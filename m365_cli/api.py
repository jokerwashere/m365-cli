"""Microsoft 365 mail and calendar operations over the Outlook Web JSON protocol.

Writes are never retried automatically. IDs and change keys are opaque.
"""

import base64
import uuid
from datetime import UTC

from .contracts import array_value, changed, object_value, typed_field
from .errors import OutlookError
from .validation import calendar_range, email, iso, nonempty, paging, validate_range


def typed(name, **values):
    return {"__type": f"{name}:#Exchange", **values}


def prop(uri):
    return typed("PropertyUri", FieldURI=uri)


def envelope(action, body, *, time_zone="UTC"):
    result = typed(
        action + "JsonRequest",
        Header=typed(
            "JsonRequestHeaders",
            RequestServerVersion="V2018_01_08",
            TimeZoneContext=typed(
                "TimeZoneContext", TimeZoneDefinition=typed("TimeZoneDefinitionType", Id=time_zone)
            ),
        ),
        Body=body,
    )
    return result


def walk(data):
    if isinstance(data, dict):
        yield data
        for value in data.values():
            yield from walk(value)
    elif isinstance(data, list):
        for value in data:
            yield from walk(value)


def check_response(response):
    if not isinstance(response, dict):
        raise OutlookError("unexpected_response", "Outlook returned an invalid response envelope.")
    status = response.get("status", 0)
    if type(status) is not int:
        raise OutlookError("unexpected_response", "Outlook returned an invalid HTTP status.")
    if status == 429:
        details = {"http_status": status}
        retry = response.get("retry_after")
        if isinstance(retry, str) and retry.isdecimal():
            details["retry_after_seconds"] = int(retry)
        raise OutlookError(
            "rate_limited", "Outlook throttled the request. No automatic retry was made.", **details
        )
    if status == 409 and response.get("body") == {"error": "account_mismatch"}:
        raise OutlookError("account_mismatch", "Outlook account changed before dispatch.")
    if status in (401, 403):
        raise OutlookError(
            "session_rejected",
            f"Outlook rejected the session (HTTP {status}). Run auth login; tenant policy may also restrict this access.",
        )
    if not 200 <= status < 300:
        details = [
            {
                k: str(v)[:500]
                for k, v in obj.items()
                if k in ("ExceptionName", "ResponseCode") and isinstance(v, str)
            }
            for obj in walk(response.get("body", {}))
        ]
        raise OutlookError(
            "api_error",
            f"Outlook API returned HTTP {status}.",
            http_status=status,
            server_errors=[d for d in details if d][:3],
        )
    data = response.get("body")
    if not isinstance(data, dict) or not data:
        raise OutlookError(
            "unexpected_response", "Outlook returned an empty or malformed JSON response."
        )
    for obj in walk(data):
        code = obj.get("ResponseCode")
        if code is not None and not isinstance(code, str):
            raise OutlookError(
                "unexpected_response", "Outlook returned an invalid error-code field."
            )
        if (
            (code and code != "NoError")
            or obj.get("ResponseClass") == "Error"
            or obj.get("ExceptionName")
        ):
            mapped = (
                "conflict"
                if code
                in (
                    "ErrorIrresolvableConflict",
                    "ErrorChangeKeyRequired",
                    "ErrorInvalidChangeKey",
                    "ErrorChangeKeyRequiredForWriteOperations",
                )
                else "item_not_found"
                if code == "ErrorItemNotFound"
                else "owa_error"
            )
            raise OutlookError(
                mapped,
                "Outlook rejected the operation.",
                response_code=code,
            )
        if obj.get("error") == "non_json_response":
            raise OutlookError("unexpected_response", "Outlook returned a non-JSON response.")
        if obj.get("error"):
            raise OutlookError("owa_error", "Outlook returned an error payload.")
    return data


def extract_items(data):
    # ResponseMessages.Items contains response envelopes, not mailbox items.
    return [
        obj for obj in walk(data) if isinstance(obj.get("ItemId"), dict) and obj["ItemId"].get("Id")
    ]


def normalize(item):
    return {
        "id": item.get("ItemId", {}).get("Id"),
        "change_key": item.get("ItemId", {}).get("ChangeKey"),
        **{k: v for k, v in item.items() if k not in ("__type", "ItemId")},
    }


def collection(data):
    collections = []
    found = []

    def visit(value, parent=None):
        if isinstance(value, dict):
            if "Items" in value and parent != "ResponseMessages":
                found.append(value)
                items = array_value(value["Items"], "outlook.items")
                for item in items:
                    object_value(item.get("ItemId"), "outlook.item", fields=("Id",))
                collections.extend(items)
            for key, child in value.items():
                visit(child, key)
        elif isinstance(value, list):
            for child in value:
                visit(child, parent)

    # ResponseMessages.Items is an envelope, never an empty mailbox.
    visit(data)
    if not found:
        raise OutlookError("unexpected_response", "Outlook did not return an item collection.")
    roots = [obj for obj in walk(data) if "IncludesLastItemInRange" in obj]
    root = roots[0] if roots else {}
    for key, kind in (
        ("IncludesLastItemInRange", bool),
        ("IndexedPagingOffset", int),
        ("TotalItemsInView", int),
    ):
        typed_field(root, key, kind, "outlook.pagination", optional=True)
    return {
        "items": [normalize(i) for i in collections],
        "total": root.get("TotalItemsInView"),
        "has_more": not root["IncludesLastItemInRange"]
        if "IncludesLastItemInRange" in root
        else None,
        "next_offset": root.get("IndexedPagingOffset"),
    }


def acknowledged(data):
    if not any(obj.get("ResponseCode") == "NoError" for obj in walk(data)):
        raise OutlookError(
            "write_outcome_unknown",
            "No explicit acknowledgement returned. Inspect Outlook before retrying.",
        )
    return data


def reference(item, expected=None):
    if expected is not None and item.get("change_key") != expected:
        raise OutlookError(
            "conflict",
            "The item changed since it was read; retrieve it again before applying the change.",
        )
    if not item.get("change_key"):
        raise OutlookError(
            "missing_change_key", "No change key returned; refusing an unchecked write."
        )
    return typed("ItemId", Id=item["id"], ChangeKey=item["change_key"])


def attendee_list(addresses):
    return [
        typed("Attendee", Mailbox=typed("EmailAddress", **r), ResponseType="Unknown")
        for r in recipients(addresses)
    ]


def folder_id(name):
    known = {
        "inbox",
        "drafts",
        "sentitems",
        "deleteditems",
        "calendar",
        "msgfolderroot",
        "archive",
        "junkemail",
        "outbox",
        "contacts",
        "tasks",
        "notes",
    }
    return typed(
        "DistinguishedFolderId" if name.lower() in known else "FolderId",
        Id=name.lower() if name.lower() in known else name,
    )


def recipients(addresses):
    result = []
    for address in addresses or []:
        email(address)
        # OWA writes take addresses directly; read responses wrap them in Mailbox.
        result.append(
            {
                "Name": address,
                "EmailAddress": address,
                "RoutingType": "SMTP",
                "MailboxType": "OneOff",
            }
        )
    return result


class API:
    def __init__(self, session):
        self.session = session

    def call(self, action, body, **kwargs):
        return self.session.request(action, body, **kwargs)

    def mail_list(self, folder="inbox", limit=25, offset=0):
        paging(limit, offset)
        body = typed(
            "FindItemRequest",
            ItemShape=typed("ItemResponseShape", BaseShape="Default", BodyType="Text"),
            ParentFolderIds=[folder_id(folder)],
            Traversal="Shallow",
            ShapeName="MailListItem",
            Paging=typed(
                "IndexedPageView", BasePoint="Beginning", Offset=offset, MaxEntriesReturned=limit
            ),
            SortOrder=[
                typed("SortResults", Order="Descending", Path=prop("item:DateTimeReceived"))
            ],
        )
        return collection(self.call("FindItem", body))

    def mail_search(self, query, limit=25, offset=0):
        nonempty(query, "query")
        paging(limit, offset)
        body = {
            "Cvid": str(uuid.uuid4()),
            "LogicalId": str(uuid.uuid4()),
            "Scenario": {"Name": "owa.react"},
            "TextDecorations": "Off",
            "TimeZone": "UTC",
            "EntityRequests": [
                {
                    "EntityType": "Conversation",
                    "ContentSources": ["Exchange"],
                    "EnableTopResults": False,
                    "From": offset,
                    "Size": limit,
                    "Query": {"QueryString": query},
                    "Sort": [{"Field": "Time", "SortDirection": "Desc"}],
                    "Filter": {"Or": [{"Term": {"DistinguishedFolderName": "msgfolderroot"}}]},
                }
            ],
        }
        data = self.session.request(None, body, path="/searchservice/api/v2/query")
        sets = [
            s
            for entity in array_value(data.get("EntitySets"), "outlook.search.EntitySets")
            for s in array_value(entity.get("ResultSets"), "outlook.search.ResultSets")
        ]
        if not sets:
            raise OutlookError(
                "unexpected_response", "Search response did not contain result sets."
            )
        results = []
        for group in sets:
            typed_field(group, "MoreResultsAvailable", bool, "outlook.search", optional=True)
            values = group.get("Results")
            if (
                "Results" not in group
                and type(group.get("Total")) is int
                and group["Total"] == 0
                and group.get("MoreResultsAvailable") is False
            ):
                values = []
            for r in array_value(values, "outlook.search.Results"):
                item = object_value(r.get("Source"), "outlook.search.Source")
                results.append(
                    {"id": item.get("ItemId", {}).get("Id") or item.get("ImmutableId"), **item}
                )
        has_more = (
            any(s.get("MoreResultsAvailable", False) for s in sets)
            if all("MoreResultsAvailable" in s for s in sets)
            else None
        )
        return {
            "items": results,
            "kind": "conversation_hits",
            "has_more": has_more,
            "next_offset": offset + len(results) if results else None,
        }

    def get(self, item_id, *, calendar=False, body_type="Text"):
        from .composition import body_type as validate_body_type

        validate_body_type(body_type)
        nonempty(item_id, "item ID")
        properties = [
            "item:Body",
            "item:Attachments",
            "item:Categories",
            "item:IsDraft",
            "item:ParentFolderId",
            "item:ItemClass",
            "item:Subject",
            "item:Importance",
            "item:Flag",
        ]
        if calendar:
            properties += [
                "calendar:" + s
                for s in (
                    "Start",
                    "End",
                    "EnhancedLocation",
                    "Organizer",
                    "RequiredAttendees",
                    "OptionalAttendees",
                    "IsOrganizer",
                    "IsMeeting",
                    "MeetingRequestWasSent",
                    "Resources",
                    "IsCancelled",
                    "CalendarItemType",
                    "Recurrence",
                    "UID",
                    "IsAllDayEvent",
                    "LegacyFreeBusyStatus",
                    "MyResponseType",
                )
            ]
        else:
            properties += [
                "message:" + s
                for s in (
                    "From",
                    "ToRecipients",
                    "CcRecipients",
                    "BccRecipients",
                    "ReplyTo",
                    "IsRead",
                    "InternetMessageId",
                )
            ]
        body = typed(
            "GetItemRequest",
            ItemIds=[typed("ItemId", Id=item_id)],
            ItemShape=typed(
                "ItemResponseShape",
                BaseShape="Default",
                BodyType=body_type,
                AdditionalProperties=[prop(p) for p in properties],
            ),
        )
        items = extract_items(self.call("GetItem", body))
        if len(items) != 1:
            raise OutlookError("unexpected_response", "Expected exactly one Outlook item.")
        item = normalize(items[0])
        if calendar:
            for name in ("IsOrganizer", "IsMeeting", "IsCancelled", "MeetingRequestWasSent"):
                if name in item and type(item[name]) is not bool:
                    changed("outlook.calendar.state", name)
            for name in ("RequiredAttendees", "OptionalAttendees", "Resources"):
                if name in item:
                    array_value(item[name], "outlook.calendar." + name)
        if "Attachments" in item:
            array_value(item["Attachments"], "outlook.attachments")
        if body_type == "HTML":
            value = object_value(item.get("Body"), "outlook.html_body", fields=("BodyType",))
            if value["BodyType"] != "HTML" or not isinstance(value.get("Value"), str):
                changed("outlook.html_body", "Body")
        return item

    def draft_create(self, *, subject, body, to=None, cc=None, bcc=None, body_type="Text"):
        return self._created(
            self._mail_create(
                subject=subject,
                body=body,
                body_type=body_type,
                to=to,
                cc=cc,
                bcc=bcc,
                disposition="SaveOnly",
            )
        )

    def mail_send(self, *, subject, body, to=None, cc=None, bcc=None, body_type="Text"):
        if not (to or cc or bcc):
            raise OutlookError(
                "recipient_required",
                "Sending requires at least one --to, --cc, or --bcc recipient.",
            )
        result = self._mail_create(
            subject=subject,
            body=body,
            body_type=body_type,
            to=to,
            cc=cc,
            bcc=bcc,
            disposition="SendAndSaveCopy",
        )
        # Successful sends can return no ItemId. Require an explicit server acknowledgement.
        acknowledged(result)
        return {
            "submitted": True,
            "saved_copy_folder": "sentitems",
            "items": [normalize(i) for i in extract_items(result)],
        }

    def _mail_create(self, *, subject, body, to, cc, bcc, disposition, body_type="Text"):
        from .composition import body_type as validate_body_type

        validate_body_type(body_type)
        message = typed(
            "Message",
            Subject=subject,
            Body={"BodyType": body_type, "Value": body},
            MessageDisposition=disposition,
            IsSendIndividually=False,
            Importance="Normal",
            Sensitivity="Normal",
        )
        for key, addresses in (("ToRecipients", to), ("CcRecipients", cc), ("BccRecipients", bcc)):
            if addresses:
                message[key] = recipients(addresses)
        return self.call(
            "CreateItem",
            typed(
                "CreateItemRequest",
                Items=[message],
                MessageDisposition=disposition,
                SavedItemFolderId={
                    "BaseFolderId": folder_id(
                        "drafts" if disposition == "SaveOnly" else "sentitems"
                    )
                },
                ClientSupportsIrm=True,
                ComposeOperation="newMail",
                ShapeName="MailCompose",
                ItemShape=typed("ItemResponseShape", BaseShape="IdOnly"),
            ),
        )

    @staticmethod
    def _created(result):
        acknowledged(result)
        items = extract_items(result)
        if len(items) != 1:
            raise OutlookError(
                "write_outcome_unknown",
                "The write returned no unique item. Inspect Outlook before retrying.",
            )
        return normalize(items[0])

    def update(
        self,
        item_id,
        changes,
        *,
        calendar=False,
        notify=False,
        save_only=False,
        expected=None,
        metadata=False,
    ):
        if save_only and (not calendar or notify):
            raise OutlookError(
                "invalid_argument", "Use --save-only only for calendar edits without --notify."
            )
        allowed = {
            "Subject",
            "Body",
            "ToRecipients",
            "CcRecipients",
            "BccRecipients",
            "Importance",
            "Categories",
        }
        if calendar:
            allowed = {
                "Subject",
                "Body",
                "Start",
                "End",
                "Location",
                "RequiredAttendees",
                "OptionalAttendees",
                "IsAllDayEvent",
                "LegacyFreeBusyStatus",
                "Categories",
            }
        elif metadata:
            allowed = {"IsRead", "Importance", "Categories", "Flag"}
        if not changes or set(changes) - allowed:
            raise OutlookError("invalid_changes", "Specify at least one supported field to update.")
        current = self.get(item_id, calendar=calendar)
        self._kind(current, calendar=calendar)
        if not calendar and not metadata and current.get("IsDraft") is not True:
            raise OutlookError(
                "not_a_draft", "This mail item is not a draft; refusing to modify it."
            )
        if calendar and current.get("CalendarItemType") == "RecurringMaster":
            raise OutlookError(
                "recurring_series",
                "Select a specific occurrence; automatic edits of entire recurring series are not supported.",
            )
        if calendar and save_only:
            self._require_unsent(current)
            self._active_calendar(current)
        if calendar and set(changes) == {"Categories"}:
            # Categories belong to this mailbox copy, including an attendee copy.
            notify = False
        if calendar and set(changes) != {"Categories"}:
            calendar_range(
                changes.get("Start", current.get("Start")),
                changes.get("End", current.get("End")),
                all_day=changes.get("IsAllDayEvent", current.get("IsAllDayEvent", False)),
                time_zone=getattr(self.session, "time_zone", "UTC"),
            )
            meeting = bool(
                self._meeting(current)
                or changes.get("RequiredAttendees")
                or changes.get("OptionalAttendees")
            )
            if meeting and not current.get("IsOrganizer"):
                raise OutlookError(
                    "not_organizer", "Only the organizer can update this meeting through the CLI."
                )
            if meeting and not (notify or save_only):
                raise OutlookError(
                    "notification_required",
                    "Use --notify to send meeting updates, or --save-only for a verified never-sent event.",
                )
        item_ref = reference(current, expected)
        fields = []
        for name, value in changes.items():
            scope = (
                "calendar"
                if name
                in (
                    "Start",
                    "End",
                    "Location",
                    "RequiredAttendees",
                    "OptionalAttendees",
                    "IsAllDayEvent",
                    "LegacyFreeBusyStatus",
                )
                else "message"
                if name.endswith("Recipients") or name == "IsRead"
                else "item"
            )
            fields.append(
                typed(
                    "SetItemField",
                    Path=prop(
                        f"{scope}:{'EnhancedLocation' if calendar and name == 'Location' else name}"
                    ),
                    Item=typed(
                        "CalendarItem" if calendar else "Message",
                        **{
                            "FreeBusyType"
                            if calendar and name == "LegacyFreeBusyStatus"
                            else name: value
                        },
                    ),
                )
            )
        if not fields:
            raise OutlookError("no_changes", "Specify at least one field to update.")
        body = typed(
            "UpdateItemRequest",
            ItemChanges=[{"ItemId": item_ref, "Updates": fields}],
            ConflictResolution="NeverOverwrite",
            MessageDisposition="SaveOnly",
            SendCalendarInvitationsOrCancellations="SendToAllAndSaveCopy"
            if notify
            else "SendToNone",
        )
        acknowledged(
            self.call(
                "UpdateItem",
                body,
                time_zone=getattr(self.session, "time_zone", "UTC") if calendar else "UTC",
            )
        )
        try:
            if changes.get("Body", {}).get("BodyType") == "HTML":
                result = self.get(item_id, calendar=calendar, body_type="HTML")
            else:
                result = self.get(item_id, calendar=calendar)
            if save_only:
                self._require_unsent(result)
                result["saved_only"] = True
            if calendar:
                result["notification_submitted"] = bool(notify)
            return result
        except OutlookError as error:
            raise OutlookError(
                "readback_failed",
                "The update was acknowledged, but readback failed. Inspect the item before another write.",
                id=item_id,
                write_acknowledged=True,
                cause=error.as_dict(),
            ) from None

    @staticmethod
    def _kind(item, *, calendar=False):
        name = item.get("ItemClass", "")
        valid = (
            name.startswith("IPM.Appointment")
            if calendar
            else name.startswith(("IPM.Note", "IPM.Schedule.Meeting", "REPORT.IPM.Note"))
        )
        if not valid:
            raise OutlookError(
                "wrong_item_type",
                "This ID belongs to a different item type; use the corresponding mail or calendar command.",
            )

    def draft_send(self, item_id, *, expected=None):
        item = self.get(item_id)
        self._kind(item)
        if item.get("IsDraft") is not True:
            raise OutlookError("not_a_draft", "Only an unsent draft can be sent.")
        if not any(item.get(k) for k in ("ToRecipients", "CcRecipients", "BccRecipients")):
            raise OutlookError("recipient_required", "The draft has no recipients.")
        result = self.call(
            "UpdateItem",
            typed(
                "UpdateItemRequest",
                ItemChanges=[
                    {
                        "ItemId": reference(item, expected),
                        "Updates": [
                            typed(
                                "SetItemField",
                                Path=prop("item:Subject"),
                                Item=typed("Message", Subject=item.get("Subject", "")),
                            )
                        ],
                    }
                ],
                ConflictResolution="NeverOverwrite",
                MessageDisposition="SendAndSaveCopy",
                SendCalendarInvitationsOrCancellations="SendToNone",
            ),
        )
        acknowledged(result)
        return {"submitted": True, "draft_id": item_id, "saved_copy_folder": "sentitems"}

    def respond_mail(
        self, item_id, *, body, mode="reply", to=None, send=False, expected=None, body_type="Text"
    ):
        from .composition import body_type as validate_body_type

        validate_body_type(body_type)
        kinds = {"reply": "ReplyToItem", "reply-all": "ReplyAllToItem", "forward": "ForwardItem"}
        if mode not in kinds:
            raise OutlookError("invalid_argument", "Invalid mail response mode.")
        if mode == "forward" and not to:
            raise OutlookError(
                "recipient_required", "Forwarding requires at least one --to recipient."
            )
        item = self.get(item_id)
        self._kind(item)
        if item.get("IsDraft"):
            raise OutlookError(
                "is_a_draft", "Send or edit this draft before replying or forwarding."
            )
        disposition = "SendAndSaveCopy" if send else "SaveOnly"
        reply = typed(
            kinds[mode],
            ReferenceItemId=reference(item, expected),
            MessageDisposition=disposition,
            NewBodyContent=typed("BodyContentType", BodyType=body_type, Value=body),
        )
        if to:
            reply["ToRecipients"] = recipients(to)
        result = self.call(
            "CreateItem",
            typed(
                "CreateItemRequest",
                Items=[reply],
                MessageDisposition=disposition,
                ClientSupportsIrm=True,
                ComposeOperation="forward" if mode == "forward" else "reply",
                SavedItemFolderId={"BaseFolderId": folder_id("sentitems" if send else "drafts")},
                ItemShape=typed("ItemResponseShape", BaseShape="IdOnly"),
            ),
        )
        if not send:
            return self._created(result)
        acknowledged(result)
        return {"submitted": True, "reference_id": item_id, "saved_copy_folder": "sentitems"}

    def move(self, item_id, folder, *, copy=False, expected=None):
        item = self.get(item_id)
        self._kind(item)
        item_ref = reference(item, expected)
        target = self.folder_get(folder)
        if not target.get("FolderClass", "IPF.Note").startswith("IPF.Note"):
            raise OutlookError(
                "wrong_folder_type", "Mail can only be moved or copied into a mail folder."
            )
        if item.get("ParentFolderId", {}).get("Id") == target["id"] and not copy:
            return {"id": item_id, "unchanged": True, "folder_id": target["id"]}
        action = "CopyItem" if copy else "MoveItem"
        result = self.call(
            action,
            typed(
                action + "Request",
                ItemIds=[item_ref],
                ToFolderId={"BaseFolderId": folder_id(target["id"])},
                ReturnNewItemIds=True,
            ),
        )
        acknowledged(result)
        created = self._created(result)
        return {**created, "original_id": item_id, "folder_id": target["id"]}

    def calendars(self):
        data = self.call("GetCalendarFolders", typed("GetCalendarFoldersRequest"))
        typed_field(data, "WasSuccessful", bool, "outlook.calendar_folders")
        if data["WasSuccessful"] is not True:
            raise OutlookError("owa_error", "Calendar discovery was rejected.")
        for folder in array_value(data.get("CalendarFolders"), "outlook.calendar_folders"):
            object_value(folder.get("FolderId"), "outlook.calendar_folder", fields=("Id",))
        return data

    def calendar_list(self, start, end, limit=100, folder="calendar", offset=0, query=None):
        validate_range(start, end, window=True)
        paging(limit, offset)
        body = typed(
            "GetCalendarViewRequest",
            CalendarId=typed("TargetFolderId", BaseFolderId=folder_id(folder)),
            RangeStart=iso(start)
            .astimezone(UTC)
            .replace(tzinfo=None)
            .isoformat(timespec="milliseconds"),
            RangeEnd=iso(end)
            .astimezone(UTC)
            .replace(tzinfo=None)
            .isoformat(timespec="milliseconds"),
            ClientSupportsIrm=True,
            OptimizeExtendedPropertyLoading=True,
        )
        result = collection(self.call("GetCalendarView", body))
        outside = [
            i
            for i in result["items"]
            if not (
                i.get("Start")
                and i.get("End")
                and iso(i["Start"]) <= iso(end)
                and iso(i["End"]) > iso(start)
            )
        ]
        if outside:
            raise OutlookError(
                "range_not_honored",
                "Outlook did not honor the requested calendar window; refusing to report an incomplete agenda.",
            )
        # GetCalendarView can include an occurrence starting exactly at RangeEnd.
        # The CLI's window is end-exclusive; only this verified boundary is trimmed.
        # More distant/non-overlapping events still trigger range_not_honored above.
        result["items"] = [i for i in result["items"] if iso(i["Start"]) < iso(end)]
        result["items"].sort(key=lambda i: (iso(i["Start"]), i["id"]))
        if query is not None:
            nonempty(query, "query")
            needle = query.casefold()
            result["items"] = [
                i
                for i in result["items"]
                if needle
                in (str(i.get("Subject", "")) + " " + str(i.get("Location", ""))).casefold()
            ]
        count = len(result["items"])
        if count > offset + limit:
            result["has_more"] = True
        result["items"] = result["items"][offset : offset + limit]
        result["next_offset"] = offset + len(result["items"]) if result["items"] else None
        result["window_start"], result["window_end"] = start, end
        result["returned_from_server"] = count
        return result

    def calendar_create(
        self,
        *,
        subject,
        start,
        end,
        body="",
        body_type="Text",
        location="",
        attendees=None,
        optional=None,
        folder="calendar",
        all_day=False,
        show_as="Busy",
        save_only=False,
    ):
        nonempty(subject, "subject")
        if body_type not in ("Text", "HTML"):
            raise OutlookError("invalid_body_format", "Body type must be Text or HTML.")
        calendar_range(
            start, end, all_day=all_day, time_zone=getattr(self.session, "time_zone", "UTC")
        )
        if show_as not in ("Free", "Tentative", "Busy", "OOF", "WorkingElsewhere"):
            raise OutlookError("invalid_argument", "Invalid show-as value.")
        item = typed(
            "CalendarItem",
            Subject=subject,
            Start=start,
            End=end,
            IsAllDayEvent=all_day,
            FreeBusyType=show_as,
            Body={"BodyType": body_type, "Value": body},
            Location={"DisplayName": location},
        )
        for key, addresses in (("RequiredAttendees", attendees), ("OptionalAttendees", optional)):
            if addresses:
                item[key] = attendee_list(addresses)
        result = self.call(
            "CreateItem",
            typed(
                "CreateItemRequest",
                Items=[item],
                SavedItemFolderId={"BaseFolderId": folder_id(folder)},
                SendMeetingInvitations="SendToAllAndSaveCopy"
                if (attendees or optional) and not save_only
                else "SendToNone",
            ),
            time_zone=getattr(self.session, "time_zone", "UTC"),
        )
        created = self._created(result)
        if save_only:
            try:
                created = self.get(created["id"], calendar=True, body_type=body_type)
                self._require_unsent(created)
            except OutlookError as error:
                raise OutlookError(
                    "readback_failed",
                    "Event creation was acknowledged, but its unsent state could not be verified. Inspect before another write.",
                    id=created["id"],
                    write_acknowledged=True,
                    cause=error.as_dict(),
                ) from None
            created["saved_only"] = True
        return {
            **created,
            "notification_submitted": bool((attendees or optional) and not save_only),
        }

    @staticmethod
    def _meeting(item):
        return bool(
            item.get("IsMeeting") is True
            or item.get("MeetingRequestWasSent") is True
            or any(item.get(k) for k in ("RequiredAttendees", "OptionalAttendees", "Resources"))
        )

    @staticmethod
    def _require_unsent(item):
        # IsDraft is false even for a never-sent Outlook calendar meeting.
        if item.get("IsOrganizer") is not True:
            raise OutlookError("not_organizer", "Only the organizer can save or send this meeting.")
        if item.get("CalendarItemType") != "Single":
            raise OutlookError(
                "recurring_series",
                "Save-only and initial sending require a single event, not a recurring series or occurrence.",
            )
        if item.get("IsCancelled") is True:
            raise OutlookError("cancelled_meeting", "A cancelled meeting cannot be saved or sent.")
        if item.get("MeetingRequestWasSent") is True:
            raise OutlookError(
                "meeting_already_sent",
                "Invitations were already sent. Use calendar update --notify for changes.",
            )
        if item.get("MeetingRequestWasSent") is not False or item.get("IsCancelled") is not False:
            raise OutlookError(
                "meeting_state_unknown",
                "The server did not verify a never-sent, active meeting.",
            )

    def _active_calendar(self, item):
        parent = item.get("ParentFolderId", {}).get("Id")
        if not parent:
            raise OutlookError(
                "meeting_state_unknown", "The event's containing folder could not be verified."
            )
        if parent == self.folder_get("deleteditems")["id"]:
            raise OutlookError(
                "already_deleted", "This event is in Deleted Items; it cannot be saved or sent."
            )

    def calendar_send(self, item_id, *, expected=None):
        current = self.get(item_id, calendar=True)
        self._kind(current, calendar=True)
        self._require_unsent(current)
        if not any(current.get(k) for k in ("RequiredAttendees", "OptionalAttendees", "Resources")):
            raise OutlookError("attendee_required", "The saved event has no attendees to invite.")
        item_ref = reference(current, expected)
        self._active_calendar(current)
        if not isinstance(current.get("Subject"), str):
            changed("outlook.calendar.send", "Subject")
        acknowledged(
            self.call(
                "UpdateItem",
                typed(
                    "UpdateItemRequest",
                    ItemChanges=[
                        {
                            "ItemId": item_ref,
                            "Updates": [
                                typed(
                                    "SetItemField",
                                    Path=prop("item:Subject"),
                                    Item=typed("CalendarItem", Subject=current["Subject"]),
                                )
                            ],
                        }
                    ],
                    ConflictResolution="NeverOverwrite",
                    MessageDisposition="SaveOnly",
                    SendCalendarInvitationsOrCancellations="SendToAllAndSaveCopy",
                ),
                time_zone=getattr(self.session, "time_zone", "UTC"),
            )
        )
        try:
            result = self.get(item_id, calendar=True)
            if result.get("MeetingRequestWasSent") is not True:
                changed("outlook.calendar.send", "MeetingRequestWasSent")
        except OutlookError as error:
            raise OutlookError(
                "readback_failed",
                "Sending was acknowledged, but the sent meeting state could not be verified. Inspect the event and Sent Items; do not resend.",
                id=item_id,
                write_acknowledged=True,
                cause=error.as_dict(),
            ) from None
        return {
            **result,
            "submitted": True,
            "notification_submitted": True,
            "saved_copy_folder": "sentitems",
        }

    def delete(self, item_id, *, calendar=False, notify=False, draft_only=False, expected=None):
        current = self.get(item_id, calendar=calendar)
        self._kind(current, calendar=calendar)
        if draft_only and current.get("IsDraft") is not True:
            raise OutlookError("not_a_draft", "This item is not a draft.")
        if calendar and current.get("CalendarItemType") == "RecurringMaster":
            raise OutlookError(
                "recurring_series", "Deleting an entire recurring series is not supported."
            )
        if calendar and self._meeting(current):
            if not current.get("IsOrganizer"):
                raise OutlookError(
                    "not_organizer",
                    "Use calendar respond --response decline for a meeting organized by someone else.",
                )
            if current.get("MeetingRequestWasSent") is False:
                # No prior invitation exists to cancel. Require verified state,
                # including the single-event guard, before suppressing notices.
                self._require_unsent(current)
                notify = False
            elif not notify:
                raise OutlookError(
                    "notification_required", "Use --notify to send meeting cancellations."
                )
        item_ref = reference(current, expected)
        deleted = self.folder_get("deleteditems")
        if current.get("ParentFolderId", {}).get("Id") == deleted["id"]:
            raise OutlookError(
                "already_deleted",
                "This item is already in Deleted Items; permanent deletion is not supported.",
            )
        acknowledged(
            self.call(
                "DeleteItem",
                typed(
                    "DeleteItemRequest",
                    ItemIds=[item_ref],
                    DeleteType="MoveToDeletedItems",
                    SendMeetingCancellations="SendToAllAndSaveCopy" if notify else "SendToNone",
                    AffectedTaskOccurrences="AllOccurrences",
                ),
            )
        )
        return {
            "id": item_id,
            "moved_to_deleted_items": True,
            **({"notification_submitted": bool(notify)} if calendar else {}),
        }

    def calendar_respond(self, item_id, response, *, body="", notify=False, expected=None):
        kinds = {
            "accept": "AcceptItem",
            "tentative": "TentativelyAcceptItem",
            "decline": "DeclineItem",
        }
        if response not in kinds:
            raise OutlookError(
                "invalid_argument", "Response must be accept, tentative, or decline."
            )
        current = self.get(item_id, calendar=True)
        self._kind(current, calendar=True)
        if current.get("IsOrganizer"):
            raise OutlookError("is_organizer", "The organizer cannot respond to their own meeting.")
        if current.get("IsCancelled"):
            raise OutlookError("cancelled_meeting", "Cannot respond to a cancelled meeting.")
        if current.get("CalendarItemType") == "RecurringMaster":
            raise OutlookError(
                "recurring_series", "Respond to a specific occurrence instead of an entire series."
            )
        if not (current.get("RequiredAttendees") or current.get("OptionalAttendees")):
            raise OutlookError("not_a_meeting", "This event has no attendees.")
        item = typed(
            kinds[response],
            ReferenceItemId=reference(current, expected),
            Body={"BodyType": "Text", "Value": body},
        )
        acknowledged(
            self.call(
                "CreateItem",
                typed(
                    "CreateItemRequest",
                    Items=[item],
                    MessageDisposition="SendAndSaveCopy" if notify else "SaveOnly",
                ),
            )
        )
        return {"id": item_id, "response": response, "notification_submitted": notify}

    def folder_get(self, name):
        nonempty(name, "folder")
        result = self.call(
            "GetFolder",
            typed(
                "GetFolderRequest",
                FolderIds=[folder_id(name)],
                FolderShape=typed(
                    "FolderResponseShape",
                    BaseShape="Default",
                    AdditionalProperties=[
                        prop("folder:FolderClass"),
                        prop("folder:TotalCount"),
                        prop("folder:ChildFolderCount"),
                        prop("folder:ParentFolderId"),
                    ],
                ),
            ),
        )
        return self._folder_created(result)

    @staticmethod
    def _folder_created(result):
        folders = [
            f
            for f in walk(result)
            if isinstance(f.get("FolderId"), dict) and f["FolderId"].get("Id")
        ]
        if len(folders) != 1:
            raise OutlookError(
                "unexpected_response", "Expected exactly one folder in the response."
            )
        f = folders[0]
        return {
            "id": f["FolderId"]["Id"],
            "change_key": f["FolderId"].get("ChangeKey"),
            **{k: v for k, v in f.items() if k not in ("FolderId", "__type")},
        }

    def folder_list(self, parent="msgfolderroot", limit=100, offset=0):
        paging(limit, offset)
        data = self.call(
            "FindFolder",
            typed(
                "FindFolderRequest",
                ParentFolderIds=[folder_id(parent)],
                Traversal="Shallow",
                FolderShape=typed("FolderResponseShape", BaseShape="Default"),
                Paging=typed(
                    "IndexedPageView",
                    BasePoint="Beginning",
                    Offset=offset,
                    MaxEntriesReturned=limit,
                ),
            ),
        )
        root = next(
            (r for r in walk(data) if "Folders" in r and isinstance(r["Folders"], list)), None
        )
        if root is None:
            raise OutlookError(
                "unexpected_response", "Folder response is missing its folder collection."
            )
        array_value(root["Folders"], "outlook.folders")
        typed_field(root, "IncludesLastItemInRange", bool, "outlook.folders", optional=True)
        return {
            "items": [self._folder_created(f) for f in root["Folders"]],
            "has_more": not root["IncludesLastItemInRange"]
            if "IncludesLastItemInRange" in root
            else None,
            "next_offset": root.get("IndexedPagingOffset"),
        }

    def folder_create(self, name, *, parent="msgfolderroot", kind="mail"):
        nonempty(name, "folder name")
        if kind not in ("mail", "calendar"):
            raise OutlookError("invalid_argument", "Folder kind must be mail or calendar.")
        result = self.call(
            "CreateFolder",
            typed(
                "CreateFolderRequest",
                ParentFolderId={"BaseFolderId": folder_id(parent)},
                Folders=[
                    typed("CalendarFolder" if kind == "calendar" else "Folder", DisplayName=name)
                ],
            ),
        )
        return self._folder_created(result)

    def folder_update(self, item_id, name, *, expected=None):
        nonempty(name, "folder name")
        current = self.folder_get(item_id)
        ref = reference(current, expected)
        result = self.call(
            "UpdateFolder",
            typed(
                "UpdateFolderRequest",
                FolderChanges=[
                    {
                        "FolderId": typed("FolderId", Id=ref["Id"], ChangeKey=ref["ChangeKey"]),
                        "Updates": [
                            typed(
                                "SetFolderField",
                                Path=prop("folder:DisplayName"),
                                Folder=typed("Folder", DisplayName=name),
                            )
                        ],
                    }
                ],
            ),
        )
        acknowledged(result)
        return self.folder_get(item_id)

    def folder_delete(self, item_id, *, expected=None):
        current = self.folder_get(item_id)
        for name in (
            "inbox",
            "drafts",
            "sentitems",
            "deleteditems",
            "calendar",
            "msgfolderroot",
            "junkemail",
            "archive",
            "outbox",
            "contacts",
            "tasks",
            "notes",
        ):
            builtin = self.folder_get(name)
            if (
                name == "deleteditems"
                and current.get("ParentFolderId", {}).get("Id") == builtin["id"]
            ):
                raise OutlookError(
                    "already_deleted",
                    "This folder is already in Deleted Items; permanent deletion is not supported.",
                )
            if builtin["id"] == current["id"]:
                raise OutlookError("protected_folder", "Built-in folders cannot be deleted.")
        if current.get("TotalCount") != 0 or current.get("ChildFolderCount") != 0:
            raise OutlookError("folder_not_empty", "Only empty custom folders can be deleted.")
        ref = reference(current, expected)
        acknowledged(
            self.call(
                "DeleteFolder",
                typed(
                    "DeleteFolderRequest",
                    FolderIds=[typed("FolderId", Id=ref["Id"], ChangeKey=ref["ChangeKey"])],
                    DeleteType="MoveToDeletedItems",
                ),
            )
        )
        return {"id": item_id, "moved_to_deleted_items": True}

    def attachment_add(
        self,
        item_id,
        name,
        content,
        content_type="application/octet-stream",
        *,
        expected=None,
        inline=False,
        content_id=None,
    ):
        from .attachments import MAX_FILE_BYTES

        nonempty(name, "attachment name")
        if len(content) > MAX_FILE_BYTES:
            raise OutlookError(
                "attachment_too_large", "The CLI supports file attachments up to 10 MiB."
            )
        attachment = typed(
            "FileAttachment",
            Name=name,
            ContentType=content_type,
            IsInline=inline,
            Content=base64.b64encode(content).decode(),
        )
        if content_id is not None:
            nonempty(content_id, "content ID")
            attachment["ContentId"] = content_id
        return self._attachment_create(item_id, attachment, expected=expected)

    def attachment_add_item(self, item_id, source_id, *, name=None, expected=None):
        nonempty(source_id, "source item ID")
        if item_id == source_id:
            raise OutlookError("invalid_argument", "A draft cannot be attached to itself.")
        source = self.get(source_id)
        attachment = typed("ItemIdAttachment", ItemId=typed("ItemId", Id=source_id))
        attachment["Name"] = name or source.get("Subject") or "Attached item"
        return self._attachment_create(item_id, attachment, expected=expected)

    def _attachment_create(self, item_id, attachment, *, expected=None):
        from .attachments import describe, response_attachments

        current = self.get(item_id)
        self._kind(current)
        if current.get("IsDraft") is not True:
            raise OutlookError("not_a_draft", "Attachments can only be added to drafts.")
        data = self.call(
            "CreateAttachment",
            typed(
                "CreateAttachmentRequest",
                ParentItemId=reference(current, expected),
                Attachments=[attachment],
            ),
        )
        acknowledged(data)
        attachments = response_attachments(data, write=True)
        try:
            attachments = [describe(a) for a in attachments]
        except OutlookError as error:
            raise OutlookError(
                "write_outcome_unknown",
                "Outlook acknowledged attachment creation but returned an incompatible attachment shape. Inspect the draft before retrying.",
                cause=error.as_dict(),
            ) from None
        return {"id": item_id, "attachments": attachments}

    def attachment_list(self, item_id, *, recursive=False, max_depth=10, max_items=200):
        from .attachments import listing

        return listing(self, item_id, recursive=recursive, max_depth=max_depth, max_items=max_items)

    def attachment_inspect(self, attachment_id, *, mail_id=None):
        from .attachments import describe, read

        return describe(read(self, attachment_id, mail_id=mail_id))

    def attachment_get(self, attachment_id, *, mail_id=None, format="auto"):
        from .attachments import download

        return download(self, attachment_id, mail_id=mail_id, format=format)

    def attachment_delete(self, item_id, attachment_id, *, expected=None):
        current = self.get(item_id)
        self._kind(current)
        if current.get("IsDraft") is not True:
            raise OutlookError("not_a_draft", "Attachments can only be removed from drafts.")
        reference(current, expected)
        if attachment_id not in [
            a.get("AttachmentId", {}).get("Id") for a in current.get("Attachments", [])
        ]:
            raise OutlookError(
                "attachment_not_found", "This attachment does not belong to the specified draft."
            )
        acknowledged(
            self.call(
                "DeleteAttachment",
                typed(
                    "DeleteAttachmentRequest",
                    AttachmentIds=[typed("AttachmentId", Id=attachment_id)],
                ),
            )
        )
        return {"id": item_id, "attachment_id": attachment_id, "deleted": True}
