"""Task CLI: common flags for routine changes, typed JSON for structured fields."""

import base64
import json
import mimetypes
import sys
from datetime import date
from pathlib import Path

from .errors import OutlookError
from .operations import execute
from .task_validation import identifier, user_id, validate_payload
from .tasks_api import TasksAPI, q
from .tasks_session import TasksSession
from .validation import email, nonempty, paging

MUTATIONS = {
    "create",
    "update",
    "delete",
    "complete",
    "reopen",
    "assign",
    "unassign",
    "move",
    "add",
    "remove",
    "check",
    "uncheck",
    "archive",
    "unarchive",
    "group",
    "undo",
}


def register(groups, body_options, write_options, page_options):
    def pages(p, query=False):
        page_options(p)
        p.add_argument(
            "--max-pages",
            type=int,
            default=20,
            help="Bound the server scan (1–100); truncation is explicit",
        )
        if query:
            p.add_argument(
                "--query",
                help="Case-insensitive title substring in this bounded scan; no server $search",
            )

    def data_options(p, kind, action):
        data = p.add_mutually_exclusive_group()
        data.add_argument("--data", help="JSON object of editable fields; see docs/TASKS.md")
        data.add_argument("--data-file", help="UTF-8 JSON object file, or - for stdin")
        if kind == "task" or kind == "plan":
            p.add_argument("--title")
        elif kind in ("list", "step", "bucket", "group"):
            p.add_argument("--name")
        if kind == "task":
            if p.get_default("task_source") == "todo":
                p.add_argument(
                    "--status",
                    choices=(
                        "notStarted",
                        "inProgress",
                        "completed",
                        "waitingOnOthers",
                        "deferred",
                    ),
                )
            else:
                p.add_argument("--progress", type=int, help="Completion percentage, 0–100")
        if kind in ("task", "details") and (
            p.get_default("task_source") == "todo" or kind == "details"
        ):
            body_options(p)

    tasks = groups.add_parser(
        "tasks", help="Task account status, capability inventory and aggregate reads"
    ).add_subparsers(dest="command", required=True)
    p = tasks.add_parser("login", help="Reuse the running saved browser; never restart it")
    p.add_argument("--timeout", type=int, default=300)
    tasks.add_parser("status")
    tasks.add_parser(
        "capabilities", help="Implemented task commands and known validation gates (local)"
    )
    p = tasks.add_parser("list", help="To Do lists plus assigned Planner tasks; explicit coverage")
    p.add_argument("--source", choices=("todo", "planner", "all"), default="all")
    p.add_argument("--max-lists", type=int, default=50)
    pages(p, True)
    specs = {
        "todo": {
            "list": ("list", "get", "create", "update", "delete", "group"),
            "task": (
                "list",
                "get",
                "create",
                "update",
                "complete",
                "reopen",
                "delete",
                "assign",
                "unassign",
                "move",
            ),
            "step": ("list", "get", "create", "update", "check", "uncheck", "delete"),
            "link": ("list", "get", "create", "update", "delete"),
            "attachment": ("list", "get", "add", "delete"),
            "my-day": ("list", "add", "remove"),
            "group": ("list", "get", "create", "update", "delete"),
            "delta": ("list",),
        },
        "planner": {
            "plan": ("list", "get", "create", "update", "delete", "archive", "unarchive"),
            "bucket": ("list", "get", "create", "update", "delete"),
            "task": (
                "list",
                "get",
                "create",
                "update",
                "complete",
                "reopen",
                "delete",
                "assign",
                "unassign",
                "move",
            ),
            "details": ("get", "update"),
            "labels": ("get", "update"),
            "my-day": ("list", "add", "remove"),
            "history": ("list",),
            "comment": ("list", "create", "update", "delete"),
            "delta": ("list",),
            "deletion": ("undo",),
        },
    }
    for source, resources in specs.items():
        root = groups.add_parser(source, help=f"{source.title()} task operations").add_subparsers(
            dest="command", required=True
        )
        from .labels_cli import register_category, register_planner

        if source == "todo":
            register_category(root.add_parser("category"), write_options, action_dest="task_action")
        else:
            register_planner(root.add_parser("label"), write_options)
        for kind, actions in resources.items():
            sub = root.add_parser(kind).add_subparsers(dest="task_action", required=True)
            for action in actions:
                p = sub.add_parser(action)
                p.set_defaults(task_source=source)
                if action in MUTATIONS:
                    write_options(p)
                if action in (
                    "update",
                    "delete",
                    "complete",
                    "reopen",
                    "assign",
                    "unassign",
                    "archive",
                    "unarchive",
                ):
                    p.add_argument(
                        "--if-version",
                        help="Reviewed resource version; rejected when no atomic precondition is validated",
                    )
                if (
                    kind == "task"
                    and action not in ("list", "create")
                    or kind in ("list", "plan", "bucket", "group")
                    and action not in ("list", "create")
                    or kind in ("step", "link", "attachment")
                    and action in ("get", "update", "delete", "check", "uncheck")
                    or kind == "comment"
                    and action in ("update", "delete")
                ):
                    p.add_argument("id")
                if source == "todo" and (
                    kind == "task"
                    or kind in ("step", "link", "attachment", "delta")
                    or kind == "my-day"
                    and action != "list"
                ):
                    p.add_argument("--list", required=True, dest="list_id")
                if (
                    kind in ("step", "link", "attachment", "details", "comment", "deletion")
                    or kind == "my-day"
                    and action != "list"
                ):
                    p.add_argument("--task", required=True, dest="task_id")
                if source == "planner" and (
                    kind in ("bucket", "task")
                    and action in ("list", "create")
                    or kind in ("history", "labels")
                ):
                    p.add_argument(
                        "--plan",
                        required=kind != "task" or action == "create",
                        dest="plan_id",
                        help="Omit for task list to read assigned-to-me only",
                    )
                if action == "list" and kind != "delta" and kind != "comment":
                    pages(p, kind == "task")
                if kind == "delta":
                    p.add_argument(
                        "--cursor",
                        help="Opaque next_cursor or delta_cursor from this exact collection",
                    )
                if action in ("create", "update") and kind not in ("comment",):
                    data_options(p, kind, action)
                if kind == "comment" and action in ("create", "update"):
                    body_options(p)
                if kind == "plan" and action == "create":
                    p.add_argument(
                        "--container",
                        choices=("self", "group"),
                        default="self",
                        help="Private self-owned plan by default; group requires an existing group",
                    )
                    p.add_argument("--group-id")
                if kind == "plan" and action == "list":
                    p.add_argument(
                        "--group-id",
                        help="Existing Microsoft 365 group; otherwise own user-container plans",
                    )
                if kind == "plan" and action in ("archive", "unarchive"):
                    p.add_argument(
                        "--reason",
                        default="Archived using m365 CLI"
                        if action == "archive"
                        else "Unarchived using m365 CLI",
                    )
                if kind == "task" and action in ("assign", "unassign"):
                    if source == "planner":
                        p.add_argument(
                            "--user",
                            required=True,
                            help="self or exact Entra user UUID; does not grant access",
                        )
                    elif action == "unassign":
                        p.add_argument("--assignment", required=True)
                if kind == "task" and action == "move":
                    p.add_argument("--to-list" if source == "todo" else "--to-plan", required=True)
                    if source == "planner":
                        p.add_argument("--to-bucket")
                if kind == "task" and action == "delete" and source == "planner":
                    p.add_argument(
                        "--recurrence",
                        choices=("occurrence", "stop-series"),
                        help="Required when deleting an active recurring task",
                    )
                    p.add_argument(
                        "--allow-undo",
                        action="store_true",
                        help="Use the native deletion adapter and return its operation ID when available",
                    )
                if kind == "deletion":
                    p.add_argument(
                        "--operation-id",
                        required=True,
                        help="Acknowledged native deletion operation ID",
                    )
                if kind == "my-day" and action == "add" and source == "todo":
                    p.add_argument("--date", required=True, help="Explicit local day, YYYY-MM-DD")
                if kind == "list" and action == "group":
                    p.add_argument("--group-id", required=True)
                if kind == "attachment" and action == "add":
                    p.add_argument("--file", required=True, help="Single upload under 3 MB")
                if kind == "attachment" and action == "get":
                    p.add_argument(
                        "--output", required=True, help="New local path; never overwrite"
                    )


def prepare(args):
    if args.account:
        email(args.account)
    if hasattr(args, "limit"):
        paging(args.limit, args.offset)
    for key in ("max_pages", "max_lists"):
        if hasattr(args, key) and not 1 <= getattr(args, key) <= 100:
            raise OutlookError("invalid_argument", f"{key} must be 1–100.")
    if hasattr(args, "timeout") and not 1 <= args.timeout <= 3600:
        raise OutlookError("invalid_argument", "Timeout must be 1–3600 seconds.")
    for key in (
        "id",
        "list_id",
        "task_id",
        "plan_id",
        "group_id",
        "assignment",
        "to_list",
        "to_plan",
        "to_bucket",
        "operation_id",
        "if_version",
        "query",
        "request_id",
    ):
        if getattr(args, key, None) is not None:
            identifier(getattr(args, key), key)
    if getattr(args, "request_id", None) and len(args.request_id) > 200:
        raise OutlookError("invalid_request_id", "Request IDs are limited to 200 characters.")
    if args.group == "tasks":
        return args
    action, kind, source = args.task_action, args.command, args.group
    if getattr(args, "user", "self") != "self":
        user_id(args.user)
    if (
        source == "todo"
        and getattr(args, "if_version", None) is not None
        and (action in ("delete", "assign", "unassign") or kind == "group")
    ):
        raise OutlookError(
            "unsupported_precondition",
            "This native operation or To Do deletion has no validated atomic If-Match protection.",
        )
    if (
        source == "planner"
        and kind == "task"
        and action == "delete"
        and args.allow_undo
        and args.if_version is not None
    ):
        raise OutlookError(
            "unsupported_precondition",
            "--allow-undo selects native Planner versions; a Graph --if-version cannot be applied atomically to that adapter.",
        )
    if source == "planner" and kind == "comment" and getattr(args, "if_version", None) is not None:
        raise OutlookError(
            "unsupported_precondition",
            "Comment edits/deletion have no validated atomic version guard.",
        )
    if (
        kind == "plan"
        and action == "create"
        and (bool(args.group_id) != (args.container == "group"))
    ):
        raise OutlookError(
            "invalid_argument", "--group-id is required only with --container group."
        )
    if getattr(args, "group_id", None) and source == "planner":
        user_id(args.group_id)
    if getattr(args, "date", None):
        try:
            if date.fromisoformat(args.date).isoformat() != args.date:
                raise ValueError
        except ValueError:
            raise OutlookError("invalid_argument", "My Day requires YYYY-MM-DD.") from None
    if getattr(args, "body_file", None) == "-" and getattr(args, "data_file", None) == "-":
        raise OutlookError("invalid_argument", "Only one input may read stdin.")
    args.text = getattr(args, "body", None)
    try:
        if getattr(args, "body_file", None):
            args.text = (
                sys.stdin.read()
                if args.body_file == "-"
                else Path(args.body_file).read_text(encoding="utf-8")
            )
        if getattr(args, "data_file", None):
            args.data = (
                sys.stdin.read()
                if args.data_file == "-"
                else Path(args.data_file).read_text(encoding="utf-8")
            )
    except (OSError, UnicodeError):
        raise OutlookError("input_file_error", "Cannot read the UTF-8 input file.") from None
    args.payload = {}
    if hasattr(args, "data"):
        try:
            args.payload = json.loads(args.data) if args.data is not None else {}
        except ValueError:
            raise OutlookError("invalid_argument", "--data must be valid JSON.") from None
        if not isinstance(args.payload, dict):
            raise OutlookError("invalid_argument", "--data must be a JSON object.")
        values = {"title": getattr(args, "title", None)}
        if hasattr(args, "name"):
            values[
                {"todo": {"group": "Name"}.get(kind, "displayName"), "planner": "name"}[source]
            ] = args.name
        if getattr(args, "status", None) is not None:
            values["status"] = args.status
        if getattr(args, "progress", None) is not None:
            values["percentComplete"] = args.progress
        if args.text is not None:
            values["body" if source == "todo" else "description"] = (
                {"contentType": "text", "content": args.text} if source == "todo" else args.text
            )
        for key, value in values.items():
            if value is not None:
                if key in args.payload:
                    raise OutlookError(
                        "invalid_argument", f"{key} is provided both in JSON and a flag."
                    )
                args.payload[key] = value
        validate_payload(source + "." + kind, args.payload, create=action == "create")
        # Resolved input is fingerprinted; file location and formatting are not.
        args.data = None
    if kind == "comment" and action in ("create", "update"):
        nonempty(args.text, "body")
    if kind == "attachment" and action == "add":
        try:
            with Path(args.file).open("rb") as stream:
                args.file_content = stream.read(3_000_000)
        except OSError:
            raise OutlookError("input_file_error", "Cannot read attachment file.") from None
        if len(args.file_content) >= 3_000_000:
            raise OutlookError(
                "attachment_too_large",
                "Single To Do uploads must be under 3 MB; upload sessions are not implemented.",
            )
    if kind == "attachment" and action == "get" and Path(args.output).exists():
        raise OutlookError("output_exists", "Output already exists; choose a new path.")
    return args


def listing_options(args):
    return {
        "limit": args.limit,
        "offset": args.offset,
        "max_pages": args.max_pages,
        "query": getattr(args, "query", None),
    }


def todo_dispatch(api, a):
    action, kind = a.task_action, a.command
    expected = getattr(a, "if_version", None)
    if kind == "delta":
        return api.delta("todo", a.list_id, a.cursor)
    if kind == "my-day":
        if action == "list":
            return api.listing(
                "todo",
                "myDayFeed/sections/tasks",
                "todo",
                "task",
                key="Value",
                id_key="Id",
                **listing_options(a),
            )
        return api.todo_myday(a.list_id, a.task_id, a.date if action == "add" else None)
    if kind == "group":
        path = "foldergroups" + ("/" + q(a.id) if hasattr(a, "id") else "")
        if action == "list":
            return api.listing(
                "todo", path, "todo", kind, key="Value", id_key="Id", **listing_options(a)
            )
        if action == "get":
            return api.get("todo", path, "todo", kind, key="Id")
        if action == "create":
            return api.create("todo", path, a.payload, "todo", kind, key="Id")
        if action == "update":
            return api.patch(
                "todo", path, a.payload, "todo", kind, expected, key="Id", guarded=False
            )
        # Groups with lists are deliberately refused, to avoid implicit reorganization.
        lists, truncated, _ = api.scan("todo", "taskfolders", key="Value", id_key="Id")
        if truncated or any(i.get("ParentFolderGroupId") == a.id for i in lists):
            raise OutlookError("resource_not_empty", "Only an empty To Do group may be deleted.")
        return api.delete("todo", path, "todo", kind, expected, key="Id", guarded=False)
    if kind == "list":
        path = api.todo_path(getattr(a, "id", None))
        if action in ("update", "delete", "group"):
            api.todo_list_guard(a.id)
        if action == "group":
            api.entity("todo", "foldergroups/" + q(a.group_id), key="Id")
            return api.patch(
                "todo",
                "taskfolders/" + q(a.id),
                {"ParentFolderGroupId": a.group_id},
                "todo",
                "list",
                key="Id",
                guarded=False,
            )
    else:
        path = api.todo_path(a.list_id)
        if kind == "task":
            path += "/tasks" + ("/" + q(a.id) if hasattr(a, "id") else "")
            if action in ("assign", "unassign"):
                return api.todo_assign(a.list_id, a.id, getattr(a, "assignment", None))
            if action == "move":
                return api.todo_move(a.list_id, a.id, a.to_list)
        else:
            path += (
                "/tasks/"
                + q(a.task_id)
                + "/"
                + {
                    "step": "checklistItems",
                    "link": "linkedResources",
                    "attachment": "attachments",
                }[kind]
            )
            if hasattr(a, "id"):
                path += "/" + q(a.id)
    if action == "list":
        return api.listing("graph", path, "todo", kind, **listing_options(a))
    if action == "get":
        result = api.get("graph", path, "todo", kind)
        if kind != "attachment":
            return result
        try:
            content = base64.b64decode(result["contentBytes"], validate=True)
        except (KeyError, ValueError, TypeError):
            raise OutlookError(
                "api_changed", "Attachment response lacks valid contentBytes."
            ) from None
        try:
            with Path(a.output).open("xb") as stream:
                stream.write(content)
        except FileExistsError:
            raise OutlookError("output_exists", "Output already exists.") from None
        except OSError:
            raise OutlookError("output_file_error", "Could not save attachment.") from None
        return {
            "path": str(Path(a.output).resolve()),
            "bytes": len(content),
            "name": result.get("name"),
        }
    if action == "add":
        data = {
            "@odata.type": "#microsoft.graph.taskFileAttachment",
            "name": Path(a.file).name,
            "contentType": mimetypes.guess_type(a.file)[0] or "application/octet-stream",
            "contentBytes": base64.b64encode(a.file_content).decode(),
        }
        result = api.create("graph", path, data, "todo", kind)
        result.pop("contentBytes", None)
        return result
    if action == "create":
        return api.create("graph", path, a.payload, "todo", kind)
    if action == "delete":
        return api.delete("graph", path, "todo", kind, expected, guarded=False)
    data = (
        {"status": "completed" if action == "complete" else "notStarted"}
        if action in ("complete", "reopen")
        else {"isChecked": action == "check"}
        if action in ("check", "uncheck")
        else a.payload
    )
    result = api.patch(
        "graph", path, data, "todo", kind, expected, guarded=kind in ("task", "list")
    )
    if action == "complete" and result.get("recurrence"):
        result["successor_may_be_created"] = True
        result["recurrence_effect"] = (
            "Occurrence completed; the current ID may now identify the next active occurrence. Re-list before acting on the series."
        )
    return result


def planner_dispatch(api, a):
    action, kind = a.task_action, a.command
    expected = getattr(a, "if_version", None)
    if kind == "delta":
        return api.delta("planner", cursor=a.cursor)
    if kind == "comment":
        if expected is not None:
            raise OutlookError(
                "unsupported_precondition",
                "Comment versions have no validated atomic precondition.",
            )
        return api.comment(a.task_id, action, getattr(a, "id", None), a.text)
    if kind == "deletion":
        return api.undo(a.operation_id, a.task_id)
    if kind == "my-day":
        if action == "list":
            return api.listing(
                "graph", "beta/me/planner/myDayTasks", "planner", "task", **listing_options(a)
            )
        return api.patch(
            "graph",
            "beta/planner/tasks/" + q(a.task_id),
            {"isOnMyDay": action == "add"},
            "planner",
            "task",
        )
    if kind == "history":
        return api.listing(
            "graph",
            "beta/planner/plans/" + q(a.plan_id) + "/historyItems",
            "planner",
            "history",
            **listing_options(a),
        )
    if kind == "plan":
        path = "beta/planner/plans" + ("/" + q(a.id) if hasattr(a, "id") else "")
        if action == "list":
            path = (
                "beta/groups/" + q(a.group_id) + "/planner/plans"
                if a.group_id
                else "beta/users/" + q(api.s.oid) + "/planner/plans"
            )
        if action == "create":
            container = (
                "groups/" + q(a.group_id) if a.container == "group" else "users/" + q(api.s.oid)
            )
            a.payload = {
                **a.payload,
                "container": {"url": "https://graph.microsoft.com/beta/" + container},
            }
        if action in ("archive", "unarchive"):
            return api.planner_archive(a.id, action == "archive", a.reason, expected)
    elif kind == "bucket":
        path = "v1.0/planner/buckets" + ("/" + q(a.id) if hasattr(a, "id") else "")
        if action == "list":
            path = "v1.0/planner/plans/" + q(a.plan_id) + "/buckets"
        if action == "create":
            a.payload = {"orderHint": " !", **a.payload, "planId": a.plan_id}
    elif kind in ("details", "labels"):
        path = (
            "v1.0/planner/"
            + ("tasks/" + q(a.task_id) if kind == "details" else "plans/" + q(a.plan_id))
            + "/details"
        )
    else:
        path = "beta/planner/tasks" + ("/" + q(a.id) if hasattr(a, "id") else "")
        if action == "list":
            path = (
                "beta/planner/plans/" + q(a.plan_id) + "/tasks"
                if a.plan_id
                else "beta/me/planner/tasks"
            )
        if action == "create":
            a.payload = {**a.payload, "planId": a.plan_id}
        if action in ("assign", "unassign"):
            uid = api.s.oid if a.user == "self" else a.user
            a.payload = {
                "assignments": {
                    uid: {"@odata.type": "#microsoft.graph.plannerAssignment", "orderHint": " !"}
                    if action == "assign"
                    else None
                }
            }
        if action == "move":
            return api.planner_move(a.id, a.to_plan, a.to_bucket)
        if action == "delete":
            return api.planner_delete_task(a.id, expected, a.recurrence, a.allow_undo)
    if action == "list":
        return api.listing("graph", path, "planner", kind, **listing_options(a))
    if action == "get":
        return api.get("graph", path, "planner", kind)
    if action == "create":
        return api.create("graph", path, a.payload, "planner", kind)
    if action == "delete":
        return api.delete("graph", path, "planner", kind, expected)
    data = (
        {"percentComplete": 100 if action == "complete" else 0}
        if action in ("complete", "reopen")
        else a.payload
    )
    return api.patch("graph", path, data, "planner", kind, expected)


def run(args, plan):
    if args.group == "tasks" and args.command == "capabilities":
        from .maintainer import inventory

        return {
            "capabilities": [c for c in inventory()["capabilities"] if c["module"] == "tasks"],
            "not_implemented": [
                "premium scheduling",
                "sharing/membership administration",
                "large/resumable uploads",
                "copy",
                "To Do ungroup/restore",
                "other-person To Do assignment",
            ],
        }
    required = {"graph"}
    if args.group == "tasks" and args.command in ("status", "login"):
        required |= {"todo", "planner"}
    if args.group == "todo" and (
        args.command in ("group", "my-day")
        or args.task_action in ("assign", "unassign", "move", "group")
    ):
        required.add("todo")
    if args.group == "planner" and (
        args.command in ("comment", "deletion")
        or args.task_action == "move"
        or getattr(args, "allow_undo", False)
    ):
        required.add("planner")
    with TasksSession(args.account, required, getattr(args, "timeout", 30)) as session:
        if args.group == "tasks" and args.command in ("login", "status"):
            return {
                "authenticated": True,
                "account": session.account,
                "tenant": session.tenant,
                "user_id": session.oid,
                "services": sorted(required),
            }
        api = TasksAPI(session)
        from .references import References

        refs = References(session.account, session.tenant)
        if args.group == "tasks":
            return refs.tasks(
                api.aggregate(
                    args.source, args.limit, args.offset, args.query, args.max_pages, args.max_lists
                ),
                args,
            )

        def action():
            parent_task = None
            if args.group == "planner" and args.command in ("details", "comment"):
                parent_task = api.entity("graph", "beta/planner/tasks/" + q(args.task_id))
            result = (
                todo_dispatch(api, args) if args.group == "todo" else planner_dispatch(api, args)
            )
            return refs.tasks(result, args, parent_task)

        if args.mutation:
            return execute(
                args.request_id,
                session.account,
                {"service": args.group, "tenant": session.tenant, "command": plan(args)},
                action,
            )
        return action()
