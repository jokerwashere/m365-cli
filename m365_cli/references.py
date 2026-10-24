"""Account-scoped durable references and explicit native link coverage.

No network calls, inferred identity matches, or message bodies in the registry.
Provider IDs remain the identifiers accepted by service commands.
"""

import hashlib
import json
import re
import uuid
from datetime import UTC, datetime
from functools import wraps
from urllib.parse import quote, urlencode

from .errors import OutlookError
from .operations import atomic_json
from .session import read_config, state_dir

REVISION = 1
PREFIX = "m365:"


def segment(value):
    return quote(str(value), safe="")


def outlook_url(kind, identifier):
    if kind == "mail":
        return "https://outlook.office.com/mail/deeplink/read/" + segment(identifier)
    if kind == "event":
        return "https://outlook.office.com/owa/?" + urlencode(
            {"ItemID": identifier, "exvsurl": "1", "path": "/calendar/item"}
        )
    if kind == "folder":
        return "https://outlook.cloud.microsoft/mail/" + segment(identifier)
    return None


def todo_url(kind, identifier):
    # These routes use the same remote IDs as the Graph/native To Do adapters.
    return "https://to-do.office.com/tasks/" + (
        "id/" + quote(str(identifier), safe="=") + "/details"
        if kind == "task"
        else quote(str(identifier), safe="=")
    )


def planner_url(kind, identifier, tenant, plan_id=None):
    base = "https://planner.cloud.microsoft/webui/v1/"
    if kind == "plan":
        route = "plan/" + segment(identifier)
    elif plan_id:
        route = "plan/" + segment(plan_id) + "/task/" + segment(identifier)
    else:
        return None
    return base + route + ("?" + urlencode({"tid": tenant}) if tenant else "")


class Registry:
    """One atomic, private registry per account. Caller owns the CLI command lock."""

    def __init__(self, account, tenant=None):
        if not isinstance(account, str) or not account:
            raise OutlookError("account_required", "A verified or pinned account is required.")
        self.account, self.tenant = account.casefold(), tenant
        key = hashlib.sha256(self.account.encode()).hexdigest()
        self.path = state_dir() / "references" / (key + ".json")
        try:
            self.data = (
                json.loads(self.path.read_text())
                if self.path.exists()
                else {"revision": REVISION, "account": self.account, "aliases": {}, "records": {}}
            )
            if (
                self.data.get("revision") != REVISION
                or self.data.get("account") != self.account
                or not isinstance(self.data.get("aliases"), dict)
                or not isinstance(self.data.get("records"), dict)
                or any(
                    not isinstance(v, dict)
                    or not isinstance(v.get("locator"), dict)
                    or not isinstance(v.get("id"), str)
                    or (v.get("web_url") is not None and not isinstance(v.get("web_url"), str))
                    for v in self.data["records"].values()
                )
                or any(v not in self.data["records"] for v in self.data["aliases"].values())
            ):
                raise ValueError
        except (OSError, ValueError, AttributeError, TypeError):
            raise OutlookError(
                "reference_state_invalid",
                "Cannot read the reference registry. Restore its backup before running commands; do not discard move mappings.",
            ) from None
        self.changed = False

    def key(self, source, kind, identifier, parent=None):
        return json.dumps(
            [
                self.account,
                self.tenant if source != "outlook" else None,
                source,
                kind,
                parent,
                identifier,
            ],
            separators=(",", ":"),
        )

    def parent_key(self, source, kind, identifier):
        return self.prior(source, kind, identifier).get("stable_id") or PREFIX + str(
            uuid.uuid5(uuid.NAMESPACE_URL, self.key(source, kind, identifier))
        )

    def prior(self, source, kind, identifier, parent=None):
        sid = self.data["aliases"].get(self.key(source, kind, identifier, parent))
        return self.data["records"].get(sid, {})

    def add(
        self,
        item,
        source,
        kind,
        *,
        locator=None,
        parent=None,
        previous=None,
        url=None,
        url_kind="item",
        reason=None,
        status="observed",
    ):
        identifier = item.get("id") or item.get("Id")
        if not isinstance(identifier, str) or not identifier:
            item["reference_unavailable"] = "provider_id_not_returned"
            return item
        key = self.key(source, kind, identifier, parent)
        old_key = self.key(source, kind, previous, parent) if previous else key
        aliases, records = self.data["aliases"], self.data["records"]
        sid = (
            aliases.get(old_key)
            or aliases.get(key)
            or PREFIX + str(uuid.uuid5(uuid.NAMESPACE_URL, old_key))
        )
        # A newly observed destination alias must not silently discard a published reference.
        other = aliases.get(key)
        if other and other != sid:
            records[other] = {**records[other], "redirect_to": sid}
        aliases[key] = sid
        if previous:
            aliases[old_key] = sid
        prior = records.get(sid, {})
        locator = {**prior.get("locator", {}), **(locator or {}), "id": identifier}
        if not url and prior.get("web_url") and not reason:
            url, url_kind = prior["web_url"], prior["url_kind"]
        record = {
            "stable_id": sid,
            "id": identifier,
            "source": source,
            "resource": kind,
            "account": self.account,
            "tenant": self.tenant if source != "outlook" else None,
            "locator": locator,
            "parent_stable_id": parent,
            "web_url": url,
            "url_kind": url_kind if url else "unavailable",
            "url_reason": reason,
            "status": status,
            "observed_at": datetime.now(UTC).isoformat(),
            "stable_id_scope": "local_registry",
            "provider_id_stability": "may_change_on_move"
            if (source, kind) in {("outlook", "mail"), ("outlook", "event"), ("todo", "task")}
            else "provider_scoped",
        }
        if item.get("successor_may_be_created"):
            record["occurrence_identity"] = "relist_after_completion"
        records[sid] = record
        # Only repair relationships we actually observed. Never guess a replacement child ID.
        for child in records.values():
            if child.get("parent_stable_id") != sid:
                continue
            if previous or (prior and prior.get("web_url") != url):
                child["locator"] = {
                    **child.get("locator", {}),
                    **(
                        {"parent_id": identifier}
                        if source == "outlook"
                        else (
                            {"file_id": identifier}
                            if kind == "file"
                            else {"item_id": identifier}
                            if kind == "list_item"
                            else {}
                        )
                        if source == "sharepoint"
                        else {"task_id": identifier}
                    ),
                    **{k: locator[k] for k in ("list_id", "plan_id") if locator.get(k)},
                }
                if child.get("url_kind") == "parent":
                    child["web_url"] = url
                if previous and previous != identifier:
                    child["locator_needs_refresh"] = True
            if status in ("deleted", "moved_to_deleted_items"):
                child.update(
                    status="parent_deleted",
                    web_url=None,
                    url_kind="unavailable",
                    url_reason="parent_deleted",
                    locator_needs_refresh=True,
                )
        if status == "deleted" and source == "outlook" and kind == "attachment":
            removed = {identifier}
            while True:
                descendants = {
                    rec["id"]
                    for rec in records.values()
                    if rec.get("source") == "outlook"
                    and rec.get("resource") == "attachment"
                    and rec.get("locator", {}).get("parent_id") == locator.get("parent_id")
                    and rec.get("locator", {}).get("parent_attachment_id") in removed
                }
                if descendants <= removed:
                    break
                removed |= descendants
            for rec in records.values():
                if (
                    rec.get("source") == "outlook"
                    and rec.get("resource") == "attachment"
                    and rec.get("id") in removed - {identifier}
                    and rec.get("locator", {}).get("parent_id") == locator.get("parent_id")
                ):
                    rec.update(
                        status="parent_deleted",
                        web_url=None,
                        url_kind="unavailable",
                        url_reason="parent_attachment_deleted",
                        locator_needs_refresh=True,
                    )
        if status == "deleted" and kind in ("list", "plan"):
            field = "list_id" if kind == "list" else "plan_id"
            removed = {
                key
                for key, value in records.items()
                if value.get("source") == source
                and value.get("tenant") == record["tenant"]
                and value.get("locator", {}).get(field) == identifier
                and key != sid
            }
            while True:
                descendants = {
                    key
                    for key, value in records.items()
                    if value.get("parent_stable_id") in removed
                }
                if descendants <= removed:
                    break
                removed |= descendants
            for key in removed:
                records[key].update(
                    status="parent_deleted",
                    web_url=None,
                    url_kind="unavailable",
                    url_reason="parent_deleted",
                    locator_needs_refresh=True,
                )
        self.changed = True
        item.update(
            id=identifier,
            stable_id=sid,
            web_url=url,
            reference={k: v for k, v in record.items() if k not in ("id", "stable_id", "web_url")},
        )
        return item

    def finish(self, result):
        if self.changed:
            try:
                atomic_json(self.path, self.data)
            except OSError:
                # A remote mutation may already have succeeded. Preserve its acknowledgement.
                result["reference_warning"] = {
                    "code": "reference_save_failed",
                    "persisted": False,
                    "message": "The service operation completed, but local reference mappings could not be saved. Preserve this result; do not repeat a write to repair references.",
                }
        return result


def get(stable_id, account=None):
    if not isinstance(stable_id, str) or not re.fullmatch(r"m365:[0-9a-f-]{36}", stable_id):
        raise OutlookError("invalid_reference", "Expected a stable_id returned by m365.")
    registry = Registry(account or read_config().get("account"))
    record = registry.data["records"].get(stable_id)
    seen = set()
    while record and record.get("redirect_to"):
        if record["redirect_to"] in seen:
            raise OutlookError("reference_state_invalid", "Reference redirect cycle.")
        seen.add(record["redirect_to"])
        record = registry.data["records"].get(record["redirect_to"])
    if not record:
        raise OutlookError(
            "reference_not_found", "No reference exists for this account and M365_HOME."
        )
    return {**record, "cached": True, "existence_verified": False}


def reference_output(function):
    """Report schema drift safely, especially after an acknowledged mutation."""

    @wraps(function)
    def wrapped(self, result, args, *extra, **kwargs):
        try:
            return function(self, result, args, *extra, **kwargs)
        except (AttributeError, KeyError, TypeError, ValueError):
            raise OutlookError(
                "write_outcome_unknown" if getattr(args, "mutation", False) else "api_changed",
                "Reference metadata no longer matches the service response. Inspect the result and operation journal before another write.",
                contract="references.output",
            ) from None

    return wrapped


class References:
    def __init__(self, account, tenant=None):
        self.registry = Registry(account, tenant)

    def task(self, item, source, kind, context):
        r = self.registry
        identifier = item.get("id") or item.get("Id")
        if not identifier:
            return item
        c = {**context}
        for target, names in {
            "list_id": ("list_id", "ParentFolderId"),
            "plan_id": ("planId", "plan_id"),
            "task_id": ("taskId", "task_id"),
        }.items():
            for name in names:
                if isinstance(item.get(name), str):
                    c[target] = item[name]
                    break
        # Plan/task IDs are tenant-wide. Child IDs are scoped to their parent resource.
        child = kind in (
            "step",
            "link",
            "attachment",
            "assignment",
            "comment",
            "checklist",
            "reference",
        )
        parent_id = (
            c.get("task_id")
            if child
            else c.get("plan_id")
            if kind in ("label", "history")
            else None
        )
        parent = None
        if parent_id:
            parent_kind = "task" if child else "plan"
            parent = r.parent_key(source, parent_kind, parent_id)
        prior = r.prior(source, kind, identifier, parent)
        c = {**prior.get("locator", {}), **{k: v for k, v in c.items() if v is not None}}
        url, scope, reason = None, "item", None
        if source == "todo":
            if kind in ("list", "task"):
                url = todo_url(kind, identifier)
            elif child and c.get("task_id"):
                url, scope = todo_url("task", c["task_id"]), "parent"
            elif kind == "group":
                url, scope = "https://to-do.office.com/tasks/", "application"
                reason = "no_native_group_permalink"
        else:
            if kind in ("plan", "task"):
                url = planner_url(kind, identifier, r.tenant, c.get("plan_id"))
            elif kind in ("bucket", "labels", "label", "history"):
                plan_id = c.get("plan_id") or (identifier if kind == "labels" else None)
                url, scope = planner_url("plan", plan_id, r.tenant) if plan_id else None, "parent"
            elif kind in ("details", "comment", "checklist", "reference", "assignment"):
                task_id = c.get("task_id") or (identifier if kind == "details" else None)
                plan_id = c.get("plan_id") or r.prior("planner", "task", task_id).get(
                    "locator", {}
                ).get("plan_id")
                url, scope = (
                    planner_url("task", task_id, r.tenant, plan_id) if task_id else None,
                    "parent",
                )
            if not url:
                reason = "parent_context_not_returned; read the parent task/plan first"
        previous = item.get("previous_id") if source == "todo" and kind == "task" else None
        if kind == "task":
            c.pop("task_id", None)
        state = (
            "deleted"
            if item.get("deleted") or item.get("isDeleted") or "@removed" in item
            else "observed"
        )
        r.add(
            item,
            source,
            kind,
            locator=c,
            parent=parent,
            previous=previous,
            url=url,
            url_kind=scope,
            reason=reason,
            status=state,
        )
        parent_context = (
            {**c, "task_id": identifier}
            if kind in ("task", "details")
            else {**c, "plan_id": identifier}
        )
        for field, resource in (
            ("checklist", "checklist"),
            ("references", "reference"),
            ("assignments", "assignment"),
            ("categoryDescriptions", "label"),
        ):
            values = item.get(field)
            if isinstance(values, dict):
                # Keep the provider's payload maps pristine; these are read-only sidecar references.
                item.setdefault("child_references", {})[field] = {
                    key: self.task({"id": key}, source, resource, parent_context)
                    for key, value in values.items()
                    if value is not None
                }
        for field, resource in (
            ("Steps", "step"),
            ("Attachments", "attachment"),
            ("Assignments", "assignment"),
            ("LinkedResources", "link"),
        ):
            if isinstance(item.get(field), list):
                for value in item[field]:
                    if isinstance(value, dict):
                        self.task(value, source, resource, parent_context)
        return item

    @reference_output
    def tasks(self, result, args, parent_task=None):
        context = {k: getattr(args, k, None) for k in ("list_id", "task_id", "plan_id")}
        if parent_task:
            context["plan_id"] = parent_task.get("planId")
        if (
            args.group == "todo"
            and args.command == "task"
            and getattr(args, "task_action", None) == "assign"
        ):
            context["task_id"] = args.id
        default = {"my-day": "task", "deletion": "task", "delta": "task"}.get(
            args.command, args.command
        )

        def decorate(item):
            source = item.get("source", args.group)
            kind = item.get("resource", default)
            if args.command == "delta" and source == "planner":
                types = {
                    "plannerTask": "task",
                    "plannerPlan": "plan",
                    "plannerBucket": "bucket",
                    "plannerTaskDetails": "details",
                    "plannerPlanDetails": "labels",
                }
                kind = types.get(str(item.get("@odata.type", "")).split(".")[-1])
                if kind is None:
                    # An untyped tombstone cannot safely identify a task versus its details.
                    item.update(
                        stable_id=None,
                        web_url=None,
                        reference_unavailable="delta_resource_type_not_returned; read a typed resource",
                    )
                    return item
            return self.task(item, source, kind, context)

        if isinstance(result.get("items"), list):
            for item in result["items"]:
                decorate(item)
        else:
            if not result.get("id") and args.command == "attachment" and hasattr(args, "id"):
                result["id"] = args.id
            decorate(result)
        return self.registry.finish(result)

    def outlook_item(self, item, kind, *, parent=None, previous=None):
        if kind == "attachment":
            parent = parent or item.get("root_mail_id")
        identifier = item.get("id")
        if kind == "folder" and (
            item.get("FolderClass") == "IPF.Appointment"
            or (
                identifier
                and self.registry.prior("outlook", "folder", identifier).get("url_kind")
                == "application"
            )
        ):
            kind = "calendar"
        if not identifier:
            nested = (
                item.get("AttachmentId")
                if kind == "attachment"
                else item.get("FolderId") or item.get("CalendarFolderId")
            )
            if isinstance(nested, dict):
                identifier = item["id"] = nested.get("Id")
        if kind == "attachment" and identifier and not parent:
            parents = {
                rec.get("locator", {}).get("parent_id")
                for rec in self.registry.data["records"].values()
                if rec.get("source") == "outlook"
                and rec.get("resource") == "attachment"
                and rec.get("id") == identifier
                and rec.get("locator", {}).get("parent_id")
            }
            if len(parents) == 1:
                parent = next(iter(parents))
        if kind == "attachment" and identifier and parent:
            # A standalone inspection can precede discovery of the root mail. Preserve
            # that already-published handle when the parent becomes known later.
            r = self.registry
            unscoped = r.key("outlook", "attachment", identifier)
            scoped = r.key(
                "outlook", "attachment", identifier, r.parent_key("outlook", "mail", parent)
            )
            old = r.data["aliases"].get(unscoped)
            if old:
                other = r.data["aliases"].get(scoped)
                if other and other != old:
                    r.data["records"][other]["redirect_to"] = old
                r.data["aliases"][scoped] = old
        url = outlook_url(kind, identifier) if identifier else None
        scope, reason = "item", None
        if kind == "attachment":
            url, scope = outlook_url("mail", parent) if parent else None, "parent"
            reason = None if parent else "parent_mail_id_not_returned"
        elif kind == "category":
            url, scope, reason = (
                "https://outlook.office.com/mail/",
                "application",
                "no_validated_category_permalink",
            )
        elif kind in ("calendar", "calendar_group"):
            url, scope, reason = (
                "https://outlook.office.com/calendar/",
                "application",
                "no_validated_container_permalink",
            )
        status = "deleted" if item.get("deleted") else "observed"
        if item.get("moved_to_deleted_items"):
            status, url, reason = "moved_to_deleted_items", None, "new_provider_id_not_returned"
        item = self.registry.add(
            item,
            "outlook",
            "folder" if kind == "calendar" else kind,
            locator={
                **({"parent_id": parent} if parent else {}),
                **(
                    {"parent_attachment_id": item["parent_attachment_id"]}
                    if item.get("parent_attachment_id")
                    else {}
                ),
            },
            parent=self.registry.parent_key("outlook", "mail", parent) if parent else None,
            previous=previous,
            url=url,
            url_kind=scope,
            reason=reason,
            status=status,
        )
        for attachment in item.get("Attachments", []):
            self.outlook_item(
                attachment, "attachment", parent=parent if kind == "attachment" else identifier
            )
        if kind == "attachment" and isinstance(item.get("Item"), dict):
            for attachment in item["Item"].get("Attachments", []):
                attachment["parent_attachment_id"] = identifier
                self.outlook_item(attachment, "attachment", parent=parent)
        return item

    @reference_output
    def outlook(self, result, args):
        kind = (
            "event" if args.group == "calendar" else "folder" if args.group == "folder" else "mail"
        )
        if args.group == "folder" and (
            result.get("FolderClass") == "IPF.Appointment"
            or getattr(args, "kind", None) == "calendar"
        ):
            kind = "calendar"
        if args.group == "calendar" and args.command == "calendars":
            for folder in result.get("CalendarFolders", []):
                self.outlook_item(folder, "calendar")
            for group in result.get("CalendarGroups", []):
                group["id"] = group.get("GroupId") or (group.get("ItemId") or {}).get("Id")
                self.outlook_item(group, "calendar_group")
                for calendar in group.get("Calendars", []):
                    self.outlook_item(calendar, "calendar")
        elif args.group == "mail" and args.command == "attachment":
            if args.attachment_command == "list":
                items = result.get("items", [])
            elif args.attachment_command == "add":
                items = result.get("attachments", [])
            else:
                if args.attachment_command == "delete":
                    result["attachment_reference"] = {"id": args.attachment_id, "deleted": True}
                    items = [result["attachment_reference"]]
                    parent_item = self.outlook_item({"id": result["id"]}, "mail")
                    result.update(
                        {k: parent_item[k] for k in ("stable_id", "web_url", "reference")}
                    )
                else:
                    result["id"] = args.attachment_id
                    items = [result]
            if args.attachment_command == "add":
                self.outlook_item(result, "mail")
            for item in items:
                self.outlook_item(
                    item,
                    "attachment",
                    parent=getattr(args, "id", None) or getattr(args, "mail_id", None),
                )
        elif isinstance(result.get("items"), list):
            for item in result["items"]:
                if result.get("kind") == "conversation_hits" and not item.get("ItemId", {}).get(
                    "Id"
                ):
                    self.registry.add(
                        item,
                        "outlook",
                        "search_hit",
                        reason="search_hit_requires_mail_get_for_item_link",
                    )
                else:
                    self.outlook_item(item, kind)
        elif result.get("id"):
            previous = (
                result.get("original_id")
                if args.group == "mail" and args.command == "move"
                else None
            )
            self.outlook_item(result, kind, previous=previous)
        elif result.get("submitted"):
            result.update(
                stable_id=None,
                web_url=None,
                reference_unavailable="sent_item_id_not_returned; use a subsequent Sent Items read; do not resend",
            )
        return self.registry.finish(result)

    def teams_item(self, item, kind, context):
        r = self.registry
        identifier = item.get("id") or item.get("objectId") or item.get("mri")
        if kind == "person":
            mri = item.get("mri") or item.get("id")
            identifier = item.get("objectId") or (
                mri.removeprefix("8:orgid:")
                if isinstance(mri, str) and mri.startswith("8:orgid:")
                else identifier
            )
        if identifier:
            item.setdefault("id", identifier)
        url, scope, reason = None, "item", None
        c = {**context}
        if kind == "message":
            saved = r.prior("teams", "message", identifier, c.get("thread_id"))
            c = {**saved.get("locator", {}), **{k: v for k, v in c.items() if v is not None}}
            if item.get("reply_to"):
                c["reply_to"] = item["reply_to"]
        if kind == "chat":
            url = "https://teams.microsoft.com/l/chat/" + segment(identifier) + "/conversations"
        elif kind == "person":
            address = item.get("userPrincipalName") or item.get("email") or item.get("mail")
            if address:
                url = "https://teams.microsoft.com/l/chat/0/0?" + urlencode(
                    {"users": address, "tenantId": r.tenant or ""}
                )
                scope = "action"
                reason = "opens_chat_with_person; does_not_send"
            elif c.get("parent_chat_id"):
                url = (
                    "https://teams.microsoft.com/l/chat/"
                    + segment(c["parent_chat_id"])
                    + "/conversations"
                )
                scope, reason = "parent", "roster_member; opens_parent_chat"
            else:
                reason = "person_address_not_returned"
        elif kind in ("team", "channel"):
            group_id = item.get("group_id") or c.get("group_id")
            if group_id:
                route = (
                    (
                        "team/"
                        + segment(item.get("default_channel_id") or identifier)
                        + "/conversations"
                    )
                    if kind == "team"
                    else "channel/"
                    + segment(identifier)
                    + "/"
                    + segment(item.get("displayName") or "Channel")
                )
                query = {"groupId": group_id, "tenantId": r.tenant or ""}
                if item.get("channelType") != 1 and kind == "channel":
                    query["ngc"] = "true"
                if item.get("isShared"):
                    query["allowXTenantAccess"] = "true"
                url = "https://teams.microsoft.com/l/" + route + "?" + urlencode(query)
            else:
                reason = "team_group_id_not_returned"
        elif kind == "message":
            thread = c.get("thread_id")
            if thread:
                query = {"tenantId": r.tenant or ""}
                if c.get("channel"):
                    query["parentMessageId"] = (
                        c.get("reply_to") or item.get("reply_to") or identifier
                    )
                    if c.get("group_id"):
                        query["groupId"] = c["group_id"]
                    else:
                        reason = "team_group_id_not_returned"
                else:
                    query["context"] = json.dumps({"contextType": "chat"}, separators=(",", ":"))
                if not reason:
                    url = (
                        "https://teams.microsoft.com/l/message/"
                        + segment(thread)
                        + "/"
                        + segment(identifier)
                        + "?"
                        + urlencode(query)
                    )
            else:
                reason = "thread_id_not_returned"
        parent = c.get("thread_id") if kind == "message" else None
        provider_id = item.get("id")
        item["id"] = identifier
        r.add(
            item,
            "teams",
            kind,
            parent=parent,
            locator=c,
            url=url,
            url_kind=scope,
            reason=reason,
            status="deleted" if item.get("deleted") else "observed",
        )
        if kind == "person" and provider_id:
            item["id"] = provider_id
        for member in item.get("members", []):
            self.teams_item(member, "person", {"parent_chat_id": identifier})
        return item

    @reference_output
    def teams(self, result, args, api):
        kind = args.command
        context = {}
        if kind == "chat" and getattr(args, "teams_action", None) == "members":
            kind = "person"
            context["parent_chat_id"] = args.id
        if result.get("left") and result.get("thread_id"):
            result["id"] = result["thread_id"]
        if kind == "message":
            context = {
                "thread_id": args.chat or args.channel,
                "channel": bool(args.channel),
                "reply_to": getattr(args, "reply_to", None),
            }
            if args.channel:
                # target() already fetched inventory before every channel operation.
                context["group_id"] = api.channel(args.channel).get("group_id")
        if isinstance(result.get("items"), list):
            for item in result["items"]:
                self.teams_item(item, kind, context)
        elif isinstance(result.get("message"), dict):
            self.teams_item(result["message"], "message", context)
        elif isinstance(result.get("chat"), dict):
            self.teams_item(result["chat"], "chat", context)
        elif result.get("id") or result.get("mri") or result.get("objectId"):
            self.teams_item(result, kind, context)
        return self.registry.finish(result)
