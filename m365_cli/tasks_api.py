"""Explicit Graph, To Do-native and Planner-native task adapters. No write fallback."""

import base64
import json
from urllib.parse import quote, unquote, urlsplit

from .contracts import array_value, changed, object_value, typed_field
from .errors import OutlookError
from .task_validation import identifier
from .tasks_session import BASES, relative_path


def q(value):
    return quote(identifier(value), safe="")


def version(item):
    value = (
        item.get("@odata.etag") or item.get("@etag") or item.get("version") or item.get("ChangeKey")
    )
    if value is None:
        raise OutlookError(
            "missing_version", "No version was returned; refusing an unchecked update."
        )
    if type(value) not in (str, int) or not str(value):
        changed("tasks.version", "version")
    return str(value)


def review(item, expected):
    current = version(item)
    if expected is not None and current != expected:
        raise OutlookError("conflict", "The task resource changed since it was reviewed.")
    return {"If-Match": current}


def normalized(item, source, kind):
    native_todo = source == "todo" and "Id" in item
    required = {
        ("todo", "list"): "Name" if native_todo else "displayName",
        ("todo", "task"): "Subject" if native_todo else "title",
        ("todo", "step"): "displayName",
        ("todo", "group"): "Name",
        ("planner", "task"): "displayName" if "displayName" in item else "title",
        ("planner", "plan"): "title",
        ("planner", "bucket"): "name",
    }.get((source, kind))
    if required:
        typed_field(item, required, str, f"{source}.{kind}")
    for field, field_type in (
        ("status", str),
        ("percentComplete", int),
        ("isChecked", bool),
        ("isOnMyDay", bool),
    ):
        typed_field(item, field, field_type, f"{source}.{kind}", optional=True)
    result = {**item, "id": item.get("id") or item.get("Id"), "source": source, "resource": kind}
    try:
        result["version"] = version(item)
    except OutlookError:
        pass
    return result


def continuation(link, family, initial):
    """Only the exact collection may provide subsequent pages (no redirected credentials)."""
    if not isinstance(link, str) or len(link) > 32768 or not link.startswith(BASES[family]):
        changed("tasks.continuation", "origin")
    relative = link[len(BASES[family]) :]
    relative_path(family, relative)
    if unquote(urlsplit(relative).path) != unquote(urlsplit(initial).path):
        changed("tasks.continuation", "collection")
    return relative


def cursor_encode(session, path, initial):
    return base64.urlsafe_b64encode(
        json.dumps(
            {
                "v": 1,
                "account": session.account.casefold(),
                "tenant": session.tenant,
                "initial": initial,
                "path": path,
            },
            separators=(",", ":"),
        ).encode()
    ).decode()


def cursor_decode(session, cursor, initial):
    try:
        if len(cursor) > 65536:
            raise ValueError
        data = json.loads(base64.b64decode(cursor, altchars=b"-_", validate=True))
        if (
            data.get("v") != 1
            or data.get("account") != session.account.casefold()
            or data.get("tenant") != session.tenant
            or data.get("initial") != initial
        ):
            raise ValueError
        return continuation(BASES["graph"] + data["path"], "graph", initial)
    except (ValueError, TypeError, KeyError, AttributeError, OutlookError):
        raise OutlookError(
            "invalid_cursor", "Cursor does not match this account and collection."
        ) from None


class TasksAPI:
    def __init__(self, session):
        self.s = session

    def entity(self, family, path, *, key="id"):
        return object_value(self.s.request(family, path), f"tasks.{family}.entity", fields=(key,))

    def scan(self, family, path, *, key="value", id_key="id", max_pages=20):
        items, seen, paths = [], set(), set()
        initial = path
        duplicates = 0
        for _ in range(max_pages):
            if path in paths:
                changed("tasks.pagination", "cycle")
            paths.add(path)
            data = object_value(self.s.request(family, path), f"tasks.{family}.collection")
            batch = array_value(data.get(key), f"tasks.{family}.{key}", fields=(id_key,))
            if len(batch) > 10000 or len(items) + len(batch) > 10000:
                raise OutlookError(
                    "scan_budget_exceeded",
                    "Task collection exceeds the 10,000-item scan budget; use narrower containers.",
                )
            for item in batch:
                if item[id_key] in seen:
                    duplicates += 1
                else:
                    seen.add(item[id_key])
                    items.append(item)
            link = (
                data.get("@odata.nextLink")
                if family != "todo"
                else data.get("NextLink", data.get("@odata.nextLink"))
            )
            if link is None:
                return items, False, duplicates
            path = continuation(link, family, initial)
        return items, True, duplicates

    def listing(
        self,
        family,
        path,
        source,
        kind,
        *,
        limit=25,
        offset=0,
        query=None,
        max_pages=20,
        key="value",
        id_key="id",
    ):
        items, truncated, duplicates = self.scan(
            family, path, max_pages=max_pages, key=key, id_key=id_key
        )
        items = [normalized(i, source, kind) for i in items]
        scanned = len(items)
        if query is not None:
            items = [
                i
                for i in items
                if query.casefold()
                in str(
                    i.get("title", i.get("displayName", i.get("Subject", i.get("Name", ""))))
                ).casefold()
            ]
        return {
            "items": [normalized(i, source, kind) for i in items[offset : offset + limit]],
            "next_offset": offset + limit if offset + limit < len(items) else None,
            "has_more": offset + limit < len(items) or truncated,
            "truncated": truncated,
            "scanned": scanned,
            "matched": len(items),
            "duplicates_removed": duplicates,
            "coverage": {
                "source": source,
                "container": path.split("?")[0],
                "match_fields": ["title"] if query is not None else [],
                "notes_searched": False,
                "max_pages": max_pages,
            },
        }

    def get(self, family, path, source, kind, *, key="id"):
        return normalized(self.entity(family, path, key=key), source, kind)

    def create(self, family, path, data, source, kind, *, key="id"):
        result = object_value(
            self.s.request(family, path, "POST", data),
            f"tasks.{source}.{kind}.create",
            fields=(key,),
            write=True,
        )
        return self.after_write(lambda: normalized(result, source, kind))

    def patch(self, family, path, data, source, kind, expected=None, *, key="id", guarded=True):
        current = self.entity(family, path, key=key)
        headers = review(current, expected) if guarded else None
        if not guarded and expected is not None:
            raise OutlookError(
                "unsupported_precondition",
                "This endpoint has no validated atomic version precondition.",
            )
        result = self.s.request(family, path, "PATCH", data, headers)
        if result is not None:
            object_value(result, f"tasks.{source}.{kind}.update", fields=(key,), write=True)
        return self.after_write(lambda: self.get(family, path, source, kind, key=key))

    def after_write(self, callback):
        try:
            return callback()
        except OutlookError as exc:
            raise OutlookError(
                "write_outcome_unknown",
                "The write was acknowledged, but verification failed. Inspect before retrying.",
                cause_code=exc.code,
            ) from None

    def delete(self, family, path, source, kind, expected=None, *, key="id", guarded=True):
        current = self.entity(family, path, key=key)
        headers = review(current, expected) if guarded else None
        if not guarded and expected is not None:
            raise OutlookError(
                "unsupported_precondition",
                "This endpoint does not provide validated atomic conditional deletion.",
            )
        result = self.s.request(family, path, "DELETE", headers=headers)
        operation = self.s.last_meta.get("operation_id")
        if result is not None:
            changed(f"tasks.{source}.{kind}.delete", "expected_empty_ack", write=True)
        self.verify_deleted(family, path)
        return {
            "deleted": True,
            "id": current[key],
            "source": source,
            "resource": kind,
            **({"planId": current["planId"]} if current.get("planId") else {}),
            **({"operation_id": operation} if operation else {}),
        }

    def verify_deleted(self, family, path):
        try:
            self.s.request(family, path)
        except OutlookError as exc:
            if exc.code == "item_not_found":
                return
            raise OutlookError(
                "write_outcome_unknown",
                "Deletion was acknowledged but absence could not be verified.",
                cause_code=exc.code,
            ) from None
        raise OutlookError(
            "write_outcome_unknown",
            "Deletion was acknowledged but the resource remains visible. Inspect before retrying.",
        )

    @staticmethod
    def todo_path(list_id=None, task_id=None):
        p = "v1.0/me/todo/lists"
        if list_id is not None:
            p += "/" + q(list_id)
        if task_id is not None:
            p += "/tasks/" + q(task_id)
        return p

    def todo_list_guard(self, list_id):
        item = self.entity("graph", self.todo_path(list_id))
        if item.get("wellknownListName") not in ("none",):
            raise OutlookError(
                "protected_resource", "Only custom To Do lists may be renamed, grouped or deleted."
            )
        if item.get("isOwner") is not True:
            raise OutlookError(
                "forbidden", "Only an owned To Do list may be renamed, grouped or deleted."
            )
        return item

    def todo_native(self, list_id, task_id):
        self.entity("graph", self.todo_path(list_id, task_id))
        path = "taskfolders/" + q(list_id) + "/tasks/" + q(task_id)
        item = self.entity("todo", path, key="Id")
        if item["Id"] != task_id or item.get("ParentFolderId") != list_id:
            changed("todo.native_mapping", "Id/ParentFolderId")
        return path, item

    def todo_assign(self, list_id, task_id, assignment_id=None):
        _, native = self.todo_native(list_id, task_id)
        path = "tasks/" + q(task_id) + "/assignments"
        if assignment_id:
            # Only the current user's assignment can be removed by this validated adapter.
            _, task = self.todo_native(list_id, task_id)
            assignments = array_value(task.get("Assignments"), "todo.assignments", fields=("Id",))
            matches = [
                a
                for a in assignments
                if a["Id"] == assignment_id and a.get("Assignee", {}).get("Id") == self.s.oid
            ]
            if len(matches) != 1:
                raise OutlookError(
                    "forbidden", "Only the current user's verified assignment can be removed."
                )
            result = self.s.request("todo", path + "/" + q(assignment_id), "DELETE")
            if result is not None:
                changed("todo.unassign", "ack", write=True)
            return self.after_write(
                lambda: normalized(self.todo_native(list_id, task_id)[1], "todo", "task")
            )
        assignments = array_value(native.get("Assignments", []), "todo.assignments", fields=("Id",))
        existing = [a for a in assignments if a.get("Assignee", {}).get("Id") == self.s.oid]
        if len(existing) == 1:
            return {**normalized(existing[0], "todo", "assignment"), "already_assigned": True}
        if existing:
            changed("todo.assignments", "duplicate_self_assignment")
        data = {"Assignee": {"Id": self.s.oid}, "Assigner": {"Id": self.s.oid}}
        return self.create("todo", path, data, "todo", "assignment", key="Id")

    def todo_move(self, list_id, task_id, target):
        if list_id == target:
            raise OutlookError("invalid_argument", "Destination list is the source list.")
        self.entity("graph", self.todo_path(target))
        path, _ = self.todo_native(list_id, task_id)
        result = object_value(
            self.s.request("todo", path, "PATCH", {"ParentFolderId": target}),
            "todo.move",
            fields=("Id",),
            write=True,
        )
        if result.get("ParentFolderId") != target:
            changed("todo.move", "ParentFolderId", write=True)

        def verify():
            moved = self.get("graph", self.todo_path(target, result["Id"]), "todo", "task")
            self.verify_deleted("graph", self.todo_path(list_id, task_id))
            return {
                **moved,
                "previous_id": task_id,
                "list_id": target,
                "assignment_may_be_removed": True,
            }

        return self.after_write(verify)

    def todo_myday(self, list_id, task_id, day):
        path, _ = self.todo_native(list_id, task_id)
        result = self.s.request(
            "todo", path, "PATCH", {"CommittedDay": day + "T00:00:00Z" if day else None}
        )
        object_value(result, "todo.myday", fields=("Id",), write=True)
        return self.after_write(
            lambda: normalized(self.todo_native(list_id, task_id)[1], "todo", "task")
        )

    def delta(self, source, list_id=None, cursor=None):
        initial = (
            self.todo_path(list_id) + "/tasks/delta"
            if source == "todo"
            else "beta/me/planner/all/delta"
        )
        path = cursor_decode(self.s, cursor, initial) if cursor else initial
        data = object_value(self.s.request("graph", path), "tasks.delta")
        items = array_value(data.get("value"), "tasks.delta.value", fields=("id",))
        next_link, delta_link = data.get("@odata.nextLink"), data.get("@odata.deltaLink")
        if bool(next_link) == bool(delta_link):
            changed("tasks.delta", "nextLink/deltaLink")
        relative = continuation(next_link or delta_link, "graph", initial)
        return {
            "items": items,
            "source": source,
            "ordered_changes": True,
            "next_cursor": cursor_encode(self.s, relative, initial) if next_link else None,
            "delta_cursor": cursor_encode(self.s, relative, initial) if delta_link else None,
        }

    def planner_delete_task(self, task_id, expected=None, recurrence_intent=None, native=False):
        path = "beta/planner/tasks/" + q(task_id)
        task = self.entity("graph", path)
        review(task, expected)
        rec = task.get("recurrence")
        if rec is not None and not isinstance(rec, dict):
            changed("planner.task", "recurrence")
        active = bool(rec and rec.get("schedule"))
        stopped = False
        if active:
            if recurrence_intent not in ("occurrence", "stop-series"):
                raise OutlookError(
                    "recurrence_intent_required",
                    "This task recurs. Use --recurrence occurrence (may create a successor) or stop-series (stop the active schedule, then delete).",
                )
            if recurrence_intent == "stop-series":
                self.patch(
                    "graph", path, {"recurrence": {"schedule": None}}, "planner", "task", expected
                )
                stopped = True
                expected = None
        try:
            if native:
                result = self.delete("planner", "tasks('" + q(task_id) + "')", "planner", "task")
            else:
                result = self.delete("graph", path, "planner", "task", expected)
        except OutlookError as exc:
            if stopped:
                raise OutlookError(
                    "partial_completion",
                    "The recurrence schedule was stopped, but deletion did not finish reliably. Inspect the task before proceeding.",
                    completed_steps=["stop_recurrence"],
                    cause_code=exc.code,
                ) from None
            raise
        return {
            **result,
            "successor_may_be_created": active and not stopped,
            "recurrence_stopped": stopped,
        }

    def planner_move(self, task_id, target_plan, target_bucket=None):
        # Explicit cross-container operation; never create or change access to either plan.
        self.entity("graph", "beta/planner/plans/" + q(target_plan))
        if target_bucket:
            bucket = self.entity("graph", "v1.0/planner/buckets/" + q(target_bucket))
            if bucket.get("planId") != target_plan:
                raise OutlookError(
                    "invalid_argument", "Destination bucket belongs to a different plan."
                )
        path = "tasks('" + q(task_id) + "')"
        task = self.entity("planner", path)
        result = self.s.request(
            "planner",
            path + "/moveToPlan",
            "POST",
            {"targetPlanId": target_plan, "targetBucketId": target_bucket},
            review(task, None),
        )
        if result is not None:
            object_value(result, "planner.move", fields=("id",), write=True)

        def verify():
            item = self.get("graph", "beta/planner/tasks/" + q(task_id), "planner", "task")
            if item.get("planId") != target_plan:
                changed("planner.move", "planId", write=True)
            return item

        return self.after_write(verify)

    def planner_archive(self, plan_id, archive, justification, expected):
        path = "beta/planner/plans/" + q(plan_id)
        plan = self.entity("graph", path)
        headers = review(plan, expected)
        typed_field(plan, "isArchived", bool, "planner.plan.archive")
        if plan["isArchived"] is archive:
            return {**normalized(plan, "planner", "plan"), "unchanged": True}
        if (plan.get("archivalInfo") or {}).get("justification") == justification:
            raise OutlookError(
                "invalid_argument",
                "Microsoft rejects archive-state reversal with the previous justification. Supply a distinct --reason.",
            )
        result = self.s.request(
            "graph",
            path + ("/archive" if archive else "/unarchive"),
            "POST",
            {"justification": justification},
            headers,
        )
        if result is not None:
            changed("planner.archive", "ack", write=True)

        def verify():
            item = self.get("graph", path, "planner", "plan")
            if item.get("isArchived") is not archive:
                changed("planner.archive", "isArchived", write=True)
            return item

        return self.after_write(verify)

    def comment(self, task_id, action, message_id=None, text=None):
        path = "tasks('" + q(task_id) + "')/messages"
        self.entity("planner", "tasks('" + q(task_id) + "')")
        if action == "list":
            return self.listing("planner", path, "planner", "comment", limit=500)
        body = {"messageType": "plainText", "content": text, "threadType": "task"}
        if action == "create":
            return self.create("planner", path, body, "planner", "comment")
        messages, truncated, _ = self.scan("planner", path)
        if truncated:
            raise OutlookError(
                "incomplete_inventory", "Cannot safely locate the comment within the scan budget."
            )
        matches = [m for m in messages if m["id"] == message_id]
        if len(matches) != 1 or matches[0].get("isDeleted"):
            raise OutlookError("item_not_found", "Comment was not found or is already deleted.")
        # The service enforces authorship; mutations stay bound to this authenticated account.
        target = path + "('" + q(message_id) + "')"
        result = self.s.request(
            "planner",
            target,
            "PUT" if action == "update" else "DELETE",
            body if action == "update" else None,
        )
        if action == "update":
            object_value(result, "planner.comment.update", fields=("id",), write=True)
            if result.get("content") != text:
                changed("planner.comment.update", "content", write=True)
            return normalized(result, "planner", "comment")
        if result is not None:
            changed("planner.comment.delete", "ack", write=True)

        def verify():
            items, incomplete, _ = self.scan("planner", path)
            if incomplete or any(m["id"] == message_id and not m.get("isDeleted") for m in items):
                changed("planner.comment.delete", "isDeleted", write=True)
            return {"deleted": True, "id": message_id, "source": "planner", "soft_deleted": True}

        return self.after_write(verify)

    def undo(self, operation_id, task_id):
        # Task ID is required to independently verify the restoration outcome.
        result = self.s.request("planner", "undo", "POST", {"operationId": operation_id})
        if result is not None:
            changed("planner.undo", "ack", write=True)
        return self.after_write(
            lambda: self.get("planner", "tasks('" + q(task_id) + "')", "planner", "task")
        )

    def aggregate(self, source, limit=25, offset=0, query=None, max_pages=20, max_lists=50):
        items, coverage, errors = [], [], []
        if source in ("todo", "all"):
            lists, truncated, _ = self.scan("graph", self.todo_path(), max_pages=max_pages)
            coverage.append(
                {
                    "source": "todo",
                    "lists_available": len(lists),
                    "lists_scanned": min(len(lists), max_lists),
                    "truncated": truncated or len(lists) > max_lists,
                }
            )
            for container in lists[:max_lists]:
                try:
                    result = self.listing(
                        "graph",
                        self.todo_path(container["id"]) + "/tasks",
                        "todo",
                        "task",
                        limit=500,
                        query=query,
                        max_pages=max_pages,
                    )
                    items.extend({**i, "list_id": container["id"]} for i in result["items"])
                    coverage.append({**result["coverage"], "truncated": result["has_more"]})
                except OutlookError as exc:
                    errors.append(
                        {"source": "todo", "container": container["id"], "code": exc.code}
                    )
        if source in ("planner", "all"):
            result = self.listing(
                "graph",
                "beta/me/planner/tasks",
                "planner",
                "task",
                limit=500,
                query=query,
                max_pages=max_pages,
            )
            items.extend(result["items"])
            coverage.append(
                {
                    **result["coverage"],
                    "scope": "assigned_to_me_only",
                    "truncated": result["has_more"],
                }
            )
        partial = bool(errors) or any(c["truncated"] for c in coverage)
        return {
            "items": items[offset : offset + limit],
            "coverage": coverage,
            "errors": errors,
            "partial": partial,
            "tenant_wide": False,
            "notes_searched": False,
            "has_more": offset + limit < len(items) or partial,
            "next_offset": offset + limit if offset + limit < len(items) else None,
        }
