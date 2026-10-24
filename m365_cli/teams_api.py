"""Conservative Teams messaging API with normalized, credential-free results."""

import base64
import json
import re
import time
from urllib.parse import parse_qsl, quote, unquote, urlencode, urlsplit

from .contracts import array_value, object_value, typed_field
from .errors import OutlookError
from .teams_composition import content as message_content
from .teams_composition import matches as body_matches
from .teams_composition import text_body as text_body  # Backward-compatible internal import.
from .validation import email, nonempty, paging

VIEW = "msnp24Equivalent|supportsMessageProperties"


def identifier(value, field="id"):
    nonempty(value, field)
    if len(value) > 2048 or any(ord(c) < 32 for c in value) or value in (".", ".."):
        raise OutlookError("invalid_argument", f"Invalid {field}.")
    return value


def segment(value):
    return quote(identifier(value), safe="")


def pick(data, names):
    return {key: data[key] for key in names if key in data}


def object_response(value):
    if not isinstance(value, dict):
        raise OutlookError("unexpected_response", "Unexpected Teams response shape.")
    return value


def list_response(value):
    if not isinstance(value, list) or any(not isinstance(x, dict) for x in value):
        raise OutlookError("unexpected_response", "Unexpected Teams list shape.")
    return value


def deleted(message):
    return bool((message.get("properties") or {}).get("deletetime"))


def normalize_message(value):
    m = object_response(value)
    object_value(m, "teams.message", fields=("id",))
    if "properties" in m and m["properties"] is not None:
        object_value(m["properties"], "teams.message.properties")
    for field in ("content", "messagetype", "from"):
        typed_field(m, field, str, "teams.message", optional=True)
    result = pick(
        m,
        (
            "id",
            "version",
            "clientmessageid",
            "messagetype",
            "content",
            "contenttype",
            "from",
            "imdisplayname",
            "composetime",
            "originalarrivaltime",
            "conversationid",
        ),
    )
    # Properties can contain embedded access tokens. Return only messaging metadata.
    props = m.get("properties") or {}
    result["properties"] = pick(
        props, ("subject", "importance", "deletetime", "edittime", "mentions", "qtdMsgs")
    )
    reply = m.get("replyChainId") or props.get("replyChainId")
    if isinstance(reply, (str, int)) and not isinstance(reply, bool):
        result["reply_to"] = str(reply)
    result["deleted"] = deleted(m)
    return result


class TeamsAPI:
    def __init__(self, session):
        self.s = session
        self._inventory = None

    def inventory(self):
        if self._inventory is None:
            data = object_response(self.s.request("csa", "api/v3/teams/users/me"))
            array_value(data.get("chats"), "teams.chats", fields=("id",))
            array_value(data.get("teams"), "teams.teams", fields=("id",))
            for chat in data["chats"]:
                array_value(chat.get("members"), "teams.chat.members")
            self._inventory = data
        return self._inventory

    def inventory_page(self, items, limit=25, offset=0):
        paging(limit, offset)
        metadata = self.inventory().get("metadata") or {}
        object_value(metadata, "teams.inventory.metadata")
        typed_field(metadata, "isPartialData", bool, "teams.inventory.metadata", optional=True)
        typed_field(metadata, "hasMoreChats", bool, "teams.inventory.metadata", optional=True)
        complete = metadata.get("isPartialData") is False and metadata.get("hasMoreChats") is False
        more = offset + limit < len(items)
        return {
            "items": items[offset : offset + limit],
            "offset": offset,
            "next_offset": offset + limit if more else None,
            "has_more": True if more else False if complete else None,
            "inventory_complete": complete,
        }

    def chats(self, limit=25, offset=0, kind=None):
        items = []
        for c in self.inventory()["chats"]:
            item = pick(
                c,
                (
                    "id",
                    "title",
                    "version",
                    "threadVersion",
                    "tenantId",
                    "isExternal",
                    "hidden",
                    "isDisabled",
                    "isMessagingDisabled",
                    "isConversationDeleted",
                ),
            )
            item["kind"] = (
                "meeting"
                if c.get("threadType") == "meeting"
                else "direct"
                if c.get("isOneOnOne")
                else "group"
            )
            item["members"] = [pick(m, ("mri", "objectId", "role")) for m in c.get("members", [])]
            if not kind or item["kind"] == kind:
                items.append(item)
        return self.inventory_page(items, limit, offset)

    def self_chat(self):
        matches = [
            c
            for c in self.inventory()["chats"]
            if c.get("isOneOnOne") is True
            and c.get("isExternal") is not True
            and len(c.get("members", [])) == 1
            and any(
                m.get("mri") == self.s.mri or m.get("objectId") == self.s.oid for m in c["members"]
            )
        ]
        verified = []
        for candidate in matches:
            thread = self.thread_raw(candidate["id"])
            members = list_response(thread.get("members"))
            if (
                len(members) == 1
                and members[0].get("id") == self.s.mri
                and str(thread.get("properties", {}).get("fixedRoster", "")).lower() == "true"
            ):
                verified.append(candidate)
        matches = verified
        if len(matches) != 1:
            raise OutlookError("not_found", "Could not identify a unique self chat.")
        return {"id": matches[0]["id"], "kind": "self", "mri": self.s.mri}

    def teams(self, limit=25, offset=0):
        return self.inventory_page(
            [
                {
                    **pick(t, ("id", "displayName", "tenantId", "isArchived", "memberRole")),
                    "group_id": (t.get("teamSiteInformation") or {}).get("groupId"),
                    "default_channel_id": next(
                        (c["id"] for c in t.get("channels", []) if c.get("isGeneral")), None
                    ),
                }
                for t in self.inventory()["teams"]
            ],
            limit,
            offset,
        )

    def channels(self, team, limit=25, offset=0):
        identifier(team, "team")
        matches = [t for t in self.inventory()["teams"] if t.get("id") == team]
        if len(matches) != 1:
            raise OutlookError("not_found", "Team is not in the available inventory.")
        return self.inventory_page(
            [
                {**self.channel_summary(c), "group_id": c.get("groupId")}
                for c in array_value(matches[0].get("channels"), "teams.channels", fields=("id",))
            ],
            limit,
            offset,
        )

    @staticmethod
    def channel_summary(c):
        return pick(
            c,
            (
                "id",
                "displayName",
                "parentTeamId",
                "groupId",
                "isMember",
                "isArchived",
                "channelType",
                "isShared",
                "threadType",
                "version",
                "threadVersion",
            ),
        )

    def channel(self, channel):
        identifier(channel, "channel")
        matches = [
            c
            for t in self.inventory()["teams"]
            for c in t.get("channels", [])
            if c.get("id") == channel
        ]
        if len(matches) != 1:
            raise OutlookError("not_found", "Channel is not in the available inventory.")
        return {**self.channel_summary(matches[0]), "group_id": matches[0].get("groupId")}

    def people(self, query):
        nonempty(query, "query")
        if len(query) > 320:
            raise OutlookError("invalid_argument", "People query is too long.")
        data = object_response(
            self.s.request(
                "mt",
                "beta/users/searchV2?includeDLs=false&enableGuest=true&includeBots=false&skypeTeamsInfo=true&source=allTrue",
                method="POST",
                body=query,
                raw=True,
            )
        )
        return {
            "items": [
                pick(
                    p,
                    (
                        "objectId",
                        "mri",
                        "email",
                        "userPrincipalName",
                        "displayName",
                        "givenName",
                        "surname",
                        "accountEnabled",
                        "userType",
                        "type",
                        "mail",
                    ),
                )
                for p in list_response(data.get("value"))
            ],
            "has_more": None,
        }

    def person(self, address):
        email(address)
        items = self.people(address)["items"]
        matches = [
            p
            for p in items
            if any(
                isinstance(p.get(k), str) and p[k].casefold() == address.casefold()
                for k in ("email", "userPrincipalName", "mail")
            )
        ]
        if len(matches) != 1 or not matches[0].get("mri"):
            raise OutlookError(
                "recipient_ambiguous" if matches else "recipient_not_found",
                "An exact, unique Teams person could not be resolved.",
            )
        if matches[0].get("accountEnabled") is False:
            raise OutlookError("recipient_disabled", "The selected Teams person is disabled.")
        return matches[0]

    def thread_raw(self, thread):
        data = object_response(
            self.s.request("chat", f"v1/threads/{segment(thread)}?view=msnp24Equivalent")
        )
        if data.get("id") != thread:
            raise OutlookError("unexpected_response", "Teams returned a different thread.")
        if data.get("tenantId") and data["tenantId"] != self.s.tenant:
            raise OutlookError("tenant_mismatch", "The thread belongs to a different tenant.")
        return data

    def chat(self, thread):
        data = self.thread_raw(thread)
        return {
            **pick(data, ("id", "tenantId", "type", "version", "rosterVersion", "eTag")),
            "properties": pick(
                data.get("properties") or {},
                ("threadType", "topic", "fixedRoster", "uniquerosterthread", "threadingMode"),
            ),
            "members": [
                pick(m, ("id", "role", "type")) for m in list_response(data.get("members"))
            ],
        }

    def resolve_chat(self, address):
        person = self.person(address)
        if person["mri"] == self.s.mri:
            return self.self_chat()
        path = "v1/threads/onetoone?" + urlencode(
            {
                "view": "msnp24Equivalent",
                "userIds": self.s.mri + "," + person["mri"],
                "threadType": "chat",
                "isFederated": "false",
            }
        )
        data = object_response(self.s.request("chat", path))
        if not data.get("id"):
            raise OutlookError(
                "unexpected_response", "Teams returned no usable direct-chat identity."
            )
        thread = self.chat(data["id"])
        if {m.get("id") for m in thread["members"]} != {self.s.mri, person["mri"]}:
            raise OutlookError(
                "identity_mismatch", "Direct chat members do not match the requested people."
            )
        return thread

    def target(self, thread, channel=False, write=False):
        identifier(thread, "thread")
        if ";messageid=" in thread:
            raise OutlookError(
                "invalid_argument", "Pass the base channel ID and --reply-to separately."
            )
        if channel:
            c = self.channel(thread)
            if write and (c.get("isArchived") is True or c.get("isMember") is False):
                raise OutlookError("read_only", "The channel is archived or you are not a member.")
            if write and (c.get("isShared") or c.get("channelType") != 1):
                raise OutlookError(
                    "unsupported_channel", "Writes currently support standard channels only."
                )
        else:
            t = self.thread_raw(thread)
            if (t.get("properties") or {}).get("threadType") not in ("chat", "meeting"):
                raise OutlookError("wrong_item_type", "Use --channel for channel messages.")
            if write and self.s.mri not in {m.get("id") for m in list_response(t.get("members"))}:
                raise OutlookError("not_a_member", "You are not a member of this chat.")
        return thread

    @staticmethod
    def message_path(thread, message=None):
        path = f"v1/users/ME/conversations/{segment(thread)}/messages"
        return path + ("/" + segment(message) if message is not None else "")

    def raw_message(self, thread, message):
        data = object_response(
            self.s.request(
                "chat", self.message_path(thread, message) + "?" + urlencode({"view": VIEW})
            )
        )
        if str(data.get("id")) != message:
            raise OutlookError("unexpected_response", "Teams returned a different message.")
        return data

    def message(self, thread, message, channel=False):
        self.target(thread, channel)
        return normalize_message(self.raw_message(thread, message))

    def cursor(self, link, thread):
        # Server links can point at the underlying chat service. Only their validated
        # path/query are reused via our authenticated, fixed Teams proxy origin.
        try:
            u = urlsplit(link)
        except (ValueError, TypeError):
            raise OutlookError("invalid_cursor", "Malformed history continuation.") from None
        expected = "/" + self.message_path(thread)
        path = re.sub(r"^/api/chatsvc/[a-z0-9-]+(?=/v1/)", "", u.path)
        if unquote(path) != unquote(expected) or u.fragment:
            raise OutlookError(
                "invalid_cursor", "History continuation belongs to another resource."
            )
        params = parse_qsl(u.query, keep_blank_values=True)
        allowed = {"view", "pageSize", "startTime", "endTime", "syncState"}
        if (
            not params
            or any(k not in allowed for k, _ in params)
            or len({k for k, _ in params}) != len(params)
        ):
            raise OutlookError("invalid_cursor", "Unsupported history continuation parameters.")
        data = {
            "v": 1,
            "account": self.s.account.casefold(),
            "tenant": self.s.tenant,
            "thread": thread,
            "params": dict(params),
        }
        return base64.urlsafe_b64encode(json.dumps(data).encode()).decode()

    def decode_cursor(self, cursor, thread, limit):
        try:
            if len(cursor) > 20000:
                raise ValueError
            data = json.loads(base64.b64decode(cursor, altchars=b"-_", validate=True))
            if (
                data.get("v") != 1
                or data.get("account") != self.s.account.casefold()
                or data.get("tenant") != self.s.tenant
                or data.get("thread") != thread
            ):
                raise ValueError
            params = data["params"]
            if (
                not isinstance(params, dict)
                or not params
                or any(
                    k not in {"view", "pageSize", "startTime", "endTime", "syncState"}
                    or not isinstance(v, str)
                    for k, v in params.items()
                )
            ):
                raise ValueError
            params.update(pageSize=str(limit), view=VIEW)
            return params
        except (ValueError, KeyError, TypeError, AttributeError):
            raise OutlookError(
                "invalid_cursor",
                "Cursor is invalid or belongs to another account, tenant, or conversation.",
            ) from None

    def messages(self, thread, *, channel=False, limit=25, cursor=None, reply_to=None):
        paging(limit)
        self.target(thread, channel)
        if reply_to:
            if not channel:
                raise OutlookError(
                    "unsupported_operation", "Reply-chain listing is supported for channels only."
                )
            self.raw_message(thread, reply_to)
        destination = thread + ";messageid=" + identifier(reply_to) if reply_to else thread
        params = (
            self.decode_cursor(cursor, destination, limit)
            if cursor
            else {"view": VIEW, "pageSize": str(limit), "startTime": "0"}
        )
        data = object_response(
            self.s.request("chat", self.message_path(destination) + "?" + urlencode(params))
        )
        if data.get("tenantId") and data["tenantId"] != self.s.tenant:
            raise OutlookError("tenant_mismatch", "History belongs to another tenant.")
        items = list_response(data.get("messages"))
        metadata = object_value(data.get("_metadata") or {}, "teams.messages.metadata")
        typed_field(metadata, "backwardLink", str, "teams.messages.metadata", optional=True)
        link = metadata.get("backwardLink")
        next_cursor = self.cursor(link, destination) if link else None
        if cursor == next_cursor:
            next_cursor = None
        return {
            "items": [normalize_message(m) for m in items],
            "next_cursor": next_cursor,
            "has_more": True if next_cursor else None,
            "thread_id": thread,
        }

    def owned_message(self, thread, message):
        current = self.raw_message(thread, message)
        sender = unquote(str(current.get("from", ""))).rstrip("/").rsplit("/", 1)[-1]
        if sender != self.s.mri:
            raise OutlookError("not_owner", "Only your own Teams messages can be changed.")
        if deleted(current):
            raise OutlookError("already_deleted", "This message is already soft-deleted.")
        if current.get("messagetype") not in ("RichText/Html", "Text"):
            raise OutlookError(
                "unsupported_message", "This message type cannot be changed by the CLI."
            )
        return current

    def verify_write(self, thread, message, predicate):
        try:
            result = self.raw_message(thread, message)
            if not predicate(result):
                raise ValueError
            return normalize_message(result)
        except Exception:
            raise OutlookError(
                "readback_failed",
                "Teams acknowledged the write but read-back verification failed. Inspect the item before retrying.",
                thread_id=thread,
                message_id=message,
            ) from None

    def send(self, thread, body, *, channel=False, reply_to=None, body_type="Text"):
        content = message_content(body, body_type)
        self.target(thread, channel, write=True)
        destination = thread
        payload = {
            "messagetype": "RichText/Html",
            "contenttype": "text",
            "content": content,
            "clientmessageid": str(time.time_ns() // 1000),
            "properties": {
                "importance": "normal",
                "subject": "",
                "mentions": "[]",
                "cards": "[]",
                "links": "[]",
                "files": "[]",
            },
        }
        if reply_to:
            source = self.raw_message(thread, reply_to)
            if deleted(source):
                raise OutlookError("already_deleted", "Cannot reply to a deleted message.")
            if not channel:
                raise OutlookError(
                    "unsupported_operation",
                    "Native replies currently support channel posts only; use message send for chats.",
                )
            destination += ";messageid=" + identifier(reply_to)
            payload["replyChainId"] = reply_to
        ack = self.s.request(
            "chat", self.message_path(destination), method="POST", body=payload, write=True
        )
        mid = ack.get("OriginalArrivalTime") if isinstance(ack, dict) else None
        if not isinstance(mid, (str, int)) or isinstance(mid, bool) or not str(mid):
            raise OutlookError(
                "write_outcome_unknown",
                "Teams did not return a message ID. Inspect history before retrying.",
                client_message_id=payload["clientmessageid"],
                thread_id=thread,
            )
        message = self.verify_write(
            thread,
            str(mid),
            lambda m: (
                body_matches(content, m.get("content"), body_type)
                and m.get("messagetype") == "RichText/Html"
                and str(m.get("clientmessageid")) == payload["clientmessageid"]
            ),
        )
        return {
            "submitted": True,
            "thread_id": thread,
            "message": message,
            **(
                {
                    "body_verification": {
                        "mode": "html_structure",
                        "server_normalized": message["content"] != content,
                    }
                }
                if body_type == "HTML"
                else {}
            ),
        }

    def update(self, thread, message, body, *, channel=False, expected=None, body_type="Text"):
        content = message_content(body, body_type)
        self.target(thread, channel, write=True)
        current = self.owned_message(thread, message)
        version = str(current.get("version", ""))
        if not re.fullmatch(r"[0-9]+", version):
            raise OutlookError("version_missing", "Message has no usable edit version.")
        if expected is not None and expected != version:
            raise OutlookError("conflict", "The message changed since the reviewed version.")
        props = current.get("properties") or {}
        # Editing rich messages can erase their hidden payload or notify mentions again.
        if any(
            props.get(k) not in (None, "", "[]", [], "{}", {})
            for k in ("mentions", "cards", "files", "qtdMsgs")
        ):
            raise OutlookError(
                "unsupported_message",
                "Editing messages containing files, cards, mentions, or quotes is unsupported.",
            )
        payload = {
            "id": message,
            "messagetype": "RichText/Html",
            "contenttype": "text",
            "content": content,
            "clientmessageid": current.get("clientmessageid"),
            "properties": pick(
                props, ("importance", "subject", "mentions", "cards", "links", "files")
            ),
        }
        self.s.request(
            "chat",
            self.message_path(thread, message),
            method="PUT",
            body=payload,
            headers={"If-Match": '"' + version + '"'},
            write=True,
        )
        result = self.verify_write(
            thread,
            message,
            lambda m: (
                body_matches(content, m.get("content"), body_type)
                and m.get("messagetype") == "RichText/Html"
                and re.fullmatch(r"[0-9]+", str(m.get("version", ""))) is not None
                and str(m.get("version")) != version
            ),
        )
        return {
            "updated": True,
            "thread_id": thread,
            "message": result,
            **(
                {
                    "body_verification": {
                        "mode": "html_structure",
                        "server_normalized": result["content"] != content,
                    }
                }
                if body_type == "HTML"
                else {}
            ),
        }

    def delete(self, thread, message, *, channel=False, expected=None):
        if expected is not None:
            raise OutlookError(
                "unsupported_precondition",
                "Teams ignores If-Match on deletion; atomic conditional deletion is unsupported.",
            )
        self.target(thread, channel, write=True)
        self.owned_message(thread, message)
        self.s.request(
            "chat",
            self.message_path(thread, message) + "?behavior=softDelete",
            method="DELETE",
            write=True,
        )
        result = self.verify_write(thread, message, lambda m: deleted(m) and not m.get("content"))
        return {
            "deleted": True,
            "thread_id": thread,
            "atomic_version_guard": False,
            "message": result,
        }

    def create_chat(self, addresses, topic=None, group=False):
        people = [self.person(address) for address in addresses]
        members = list(dict.fromkeys([self.s.mri] + [p["mri"] for p in people]))
        if not group:
            if len(addresses) != 1 or topic:
                raise OutlookError(
                    "invalid_argument",
                    "Direct chat creation needs one --to and no --topic; use --group for group chats.",
                )
            try:
                return {"created": False, "chat": self.resolve_chat(addresses[0])}
            except OutlookError as error:
                if error.code != "not_found":
                    raise
        properties = {"threadType": "chat", "chatFilesIndexId": 2}
        if group:
            nonempty(topic, "topic")
            properties["topic"] = topic
        else:
            properties.update(fixedRoster=True, uniquerosterthread=True)
        ack = self.s.request(
            "chat",
            "v1/threads",
            method="POST",
            headers={"prefer": "return=representation"},
            include_location=True,
            body={
                "members": [{"id": m, "role": "Admin"} for m in members],
                "properties": properties,
            },
            write=True,
        )
        data = ack.get("body") if isinstance(ack, dict) else None
        resource = data.get("threadResource") if isinstance(data, dict) else None
        thread = resource.get("id") if isinstance(resource, dict) else None
        if not thread and isinstance(ack, dict):
            try:
                location = urlsplit(ack.get("location") or "")
            except (ValueError, TypeError):
                location = urlsplit("")
            match = re.fullmatch(r"(?:/api/chatsvc/[a-z0-9-]+)?/v1/threads/([^/]+)", location.path)
            if match and not location.query and not location.fragment:
                thread = unquote(match[1])
        if not isinstance(thread, str) or not thread:
            raise OutlookError(
                "write_outcome_unknown",
                "Teams did not return a created thread ID. Inspect chats before retrying.",
            )
        result = self.chat_readback(thread)
        if {m.get("id") for m in result["members"]} != set(members) or (
            group and result["properties"].get("topic") != topic
        ):
            raise OutlookError(
                "readback_failed",
                "Created chat did not match its intended roster/topic.",
                thread_id=thread,
            )
        return {"created": True, "chat": result}

    def chat_readback(self, thread):
        try:
            return self.chat(thread)
        except Exception:
            raise OutlookError(
                "readback_failed",
                "Teams acknowledged the chat change but read-back failed. Inspect the chat before retrying.",
                thread_id=thread,
            ) from None

    def group_chat(self, thread):
        current = self.chat(thread)
        props = current["properties"]
        if (
            props.get("threadType") != "chat"
            or str(props.get("fixedRoster", "false")).lower() == "true"
        ):
            raise OutlookError("unsupported_operation", "This operation requires a group chat.")
        if self.s.mri not in {m.get("id") for m in current["members"]}:
            raise OutlookError("not_a_member", "You are not a member of this group chat.")
        return current

    def rename_chat(self, thread, topic):
        nonempty(topic, "topic")
        self.group_chat(thread)
        self.s.request(
            "chat",
            f"v1/threads/{segment(thread)}/properties?name=topic",
            method="PUT",
            body={"topic": topic},
            write=True,
        )
        result = self.chat_readback(thread)
        if result["properties"].get("topic") != topic:
            raise OutlookError(
                "readback_failed", "Chat rename was not visible on read-back.", thread_id=thread
            )
        return {"updated": True, "chat": result}

    def leave_chat(self, thread):
        self.group_chat(thread)
        self.s.request(
            "chat",
            f"v1/threads/{segment(thread)}/members/{segment(self.s.mri)}",
            method="DELETE",
            write=True,
        )
        result = self.chat_readback(thread)
        if self.s.mri in {m.get("id") for m in result["members"]}:
            raise OutlookError(
                "readback_failed", "Chat departure was not visible on read-back.", thread_id=thread
            )
        return {"left": True, "thread_id": thread, "deleted_for_everyone": False}
