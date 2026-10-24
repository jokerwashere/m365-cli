"""Typed Outlook attachments. Never search nested items for the requested file's bytes."""

import base64
import json

from .contracts import array_value, changed, object_value, typed_field
from .errors import OutlookError
from .validation import nonempty

MAX_FILE_BYTES = 10 * 1024 * 1024
MAX_DOWNLOAD_BYTES = 50 * 1024 * 1024
CONTRACT = "outlook.attachment"
KINDS = {
    "FileAttachment": "file",
    "ItemAttachment": "item",
    "ItemIdAttachment": "item",
    "ReferenceAttachment": "reference",
    "LinkAttachment": "link",
}


def response_attachments(data, *, write=False):
    """Only inspect attachment collections in response envelopes, never embedded items."""
    body = object_value(data.get("Body"), CONTRACT, write=write)
    if "Attachments" in body:
        groups = [body]
    else:
        messages = object_value(body.get("ResponseMessages"), CONTRACT, write=write)
        groups = array_value(messages.get("Items"), CONTRACT, write=write)
    result = []
    for group in groups:
        result.extend(array_value(group.get("Attachments"), CONTRACT, write=write))
    if not result:
        changed(CONTRACT, "Attachments", write=write)
    for attachment in result:
        object_value(attachment.get("AttachmentId"), CONTRACT, fields=("Id",), write=write)
    return result


def kind(attachment):
    value = attachment.get("__type")
    if not isinstance(value, str) or not value:
        changed(CONTRACT, "__type")
    return KINDS.get(value.split(":", 1)[0], "unknown")


def clean(value):
    """Keep item properties but exclude binary payloads and ephemeral download credentials."""
    if isinstance(value, dict):
        return {
            k: clean(v)
            for k, v in value.items()
            if k not in ("Content", "MimeContent", "AttachmentOriginalUrl")
            and "token" not in k.casefold()
        }
    if isinstance(value, list):
        return [clean(v) for v in value]
    return value


def describe(attachment):
    identifier = object_value(attachment.get("AttachmentId"), CONTRACT, fields=("Id",))["Id"]
    category = kind(attachment)
    for field, expected in (
        ("Name", str),
        ("ContentType", str),
        ("ContentId", str),
        ("IsInline", bool),
        ("Size", int),
    ):
        if attachment.get(field) is not None:
            typed_field(attachment, field, expected, CONTRACT)
    result = {
        **clean(attachment),
        "id": identifier,
        "kind": category,
        "name": attachment.get("Name"),
        "content_type": attachment.get("ContentType"),
        "is_inline": attachment.get("IsInline") or False,
    }
    if category == "item":
        item = attachment.get("Item")
        if item is not None:
            object_value(item, CONTRACT + ".item")
            result["Item"] = clean(item)
            result["Item"]["Attachments"] = [describe(a) for a in children(item)]
        result["item_class"] = (item or {}).get("ItemClass") or attachment.get("EmbeddedItemClass")
    if category in ("reference", "link"):
        # This is a provider-owned link. Never fetch it with Outlook credentials.
        result["source_url"] = next(
            (
                attachment[k]
                for k in ("AttachLongPathName", "WebUrl", "SourceUrl", "LongPathName")
                if isinstance(attachment.get(k), str) and attachment[k]
            ),
            None,
        )
        result["downloadable"] = False
        result["content_status"] = "external_reference"
    elif category == "unknown":
        result["downloadable"] = False
        result["content_status"] = "unsupported_attachment_type"
    return result


def children(item):
    value = item.get("Attachments")
    if value is None:
        if item.get("HasAttachments") is True:
            changed(CONTRACT, "Item.Attachments")
        return []
    return array_value(value, CONTRACT + ".children")


def traversal_limits(max_depth, max_items):
    if type(max_depth) is not int or not 0 <= max_depth <= 30:
        raise OutlookError("invalid_argument", "Attachment max-depth must be 0–30.")
    if type(max_items) is not int or not 1 <= max_items <= 1000:
        raise OutlookError("invalid_argument", "Attachment max-items must be 1–1000.")


def read(api, attachment_id, *, mail_id=None):
    from .api import typed

    nonempty(attachment_id, "attachment ID")
    context = {"root_mail_id": mail_id} if mail_id else {}
    if mail_id:
        # Some tenants cannot GetAttachment on cloud links. Parent metadata is sufficient
        # for these non-binary types. A supplied parent is also checked, not merely trusted.
        root = api.get(mail_id)
        for entry in children(root):
            if (entry.get("AttachmentId") or {}).get("Id") == attachment_id:
                if kind(entry) in ("reference", "link", "unknown"):
                    return entry
                break
        else:
            # Nested IDs are returned by GetAttachment and belong to their root mailbox item.
            # Verify their membership rather than issue a misleading parent URL.
            tree = listing(api, mail_id, recursive=True, max_depth=30, max_items=1000, root=root)
            if attachment_id not in {a["id"] for a in tree["items"]}:
                raise OutlookError(
                    "attachment_not_found",
                    "Attachment was not found under the specified mail item.",
                    complete=tree["complete"],
                )
            entry = next(a for a in tree["items"] if a["id"] == attachment_id)
            context.update(
                {k: entry[k] for k in ("parent_attachment_id", "ancestor_attachment_ids", "depth")}
            )
            if entry["kind"] in ("reference", "link", "unknown"):
                return entry
    data = api.call(
        "GetAttachment",
        typed("GetAttachmentRequest", AttachmentIds=[typed("AttachmentId", Id=attachment_id)]),
    )
    found = response_attachments(data)
    if len(found) != 1 or not same_attachment_id(attachment_id, found[0]["AttachmentId"]["Id"]):
        changed(CONTRACT, "requested_attachment_id")
    # Validate before returning so new/malformed shapes cannot be mistaken for file content.
    describe(found[0])
    if kind(found[0]) == "item":
        object_value(found[0].get("Item"), CONTRACT + ".item")
    # GetAttachment shortens nested IDs in its response; only the requested full
    # path remains usable as a stable locator for subsequent reads/downloads.
    return {
        **found[0],
        "AttachmentId": {**found[0]["AttachmentId"], "Id": attachment_id},
        **context,
    }


def same_attachment_id(requested, returned):
    if requested == returned:
        return True
    # Observed Exchange encoding: root prefix, one-byte path count, 20-byte path
    # entries. GetAttachment returns the root + count 1 + final entry, dropping
    # ancestor entries. Accept ONLY this exact shortening; never arbitrary ID drift.
    # IDs remain opaque outside this adapter. New encodings fail conservatively.
    try:
        full = base64.b64decode(requested, validate=True)
        short = base64.b64decode(returned, validate=True)
    except (ValueError, TypeError):
        return False
    extra = len(full) - len(short)
    index = len(short) - 21
    return (
        index > 0
        and extra > 0
        and extra % 20 == 0
        and 2 <= 1 + extra // 20 <= 31
        and full[:index] == short[:index]
        and short[index] == 1
        and full[index] == 1 + extra // 20
        and full[-20:] == short[-20:]
    )


def listing(api, mail_id, *, recursive=False, max_depth=10, max_items=200, root=None):
    traversal_limits(max_depth, max_items)
    root = api.get(mail_id) if root is None else root
    result = {"items": [], "recursive": recursive, "complete": True, "truncations": []}
    seen = set()

    def truncate(reason, identifier=None):
        result["complete"] = False
        result["truncations"].append({"reason": reason, "attachment_id": identifier})

    stack = [(raw, []) for raw in reversed(children(root))]
    while stack:
        if len(result["items"]) >= max_items:
            truncate("max_items")
            break
        raw, ancestors = stack.pop()
        entry = describe(raw)
        identifier = entry["id"]
        entry.update(
            root_mail_id=mail_id,
            parent_attachment_id=ancestors[-1] if ancestors else None,
            ancestor_attachment_ids=list(ancestors),
            depth=len(ancestors),
        )
        # Listing is metadata-only, including when a future response pre-expands an item.
        entry.pop("Item", None)
        result["items"].append(entry)
        if identifier in seen:
            truncate("repeated_attachment_id", identifier)
            continue
        seen.add(identifier)
        if recursive and entry["kind"] == "unknown":
            truncate("unsupported_attachment_type", identifier)
        if not recursive or entry["kind"] != "item":
            continue
        if len(ancestors) >= max_depth:
            truncate("max_depth", identifier)
            continue
        expanded = read(api, identifier)
        if kind(expanded) != "item":
            changed(CONTRACT, "attachment_type_changed")
        item = object_value(expanded.get("Item"), CONTRACT + ".item")
        entry["item_class"] = item.get("ItemClass") or entry.get("item_class")
        stack.extend((child, [*ancestors, identifier]) for child in reversed(children(item)))
    return result


def decode_content(value):
    if not isinstance(value, str):
        changed(CONTRACT, "Content")
    if len(value) > ((MAX_DOWNLOAD_BYTES + 2) // 3) * 4:
        raise OutlookError("attachment_too_large", "Attachment download exceeds the 50 MiB limit.")
    try:
        result = base64.b64decode(value, validate=True)
    except ValueError:
        changed(CONTRACT, "Content.base64")
    if len(result) > MAX_DOWNLOAD_BYTES:
        raise OutlookError("attachment_too_large", "Attachment download exceeds the 50 MiB limit.")
    return result


def download(api, attachment_id, *, mail_id=None, format="auto"):
    if format not in ("auto", "raw", "json"):
        raise OutlookError("invalid_argument", "Attachment format must be auto, raw or json.")
    raw = read(api, attachment_id, mail_id=mail_id)
    result = describe(raw)
    category = result["kind"]
    item = raw.get("Item") or {}
    message = item.get("__type", "").split(":", 1)[0] in (
        "Message",
        "MeetingMessage",
        "MeetingRequest",
        "MeetingResponse",
        "MeetingCancellation",
    )
    if format == "json" or (
        format == "auto" and category != "file" and not (category == "item" and message)
    ):
        result.update(
            content=(json.dumps(result, ensure_ascii=False, indent=2) + "\n").encode(),
            export_format="json",
            export_content_type="application/json",
        )
        return result
    if category == "file":
        result.update(
            content=decode_content(raw.get("Content")),
            export_format="file",
            export_content_type=result["content_type"] or "application/octet-stream",
        )
        return result
    if category == "item":
        data = api.session.download_attachment(attachment_id)
        export_format, content_type = native_format(data["content"])
        result.update(
            content=data["content"],
            export_format=export_format,
            export_content_type=content_type,
            server_content_type=data["content_type"],
            export_note="Exchange MIME conversion may omit item properties. Use --format json for structured properties.",
        )
        return result
    raise OutlookError(
        "unsupported_attachment_content",
        "This attachment has no supported embedded binary content. Use get without --output to inspect it, or --format json to save its metadata.",
        kind=category,
    )


def native_format(content):
    """Classify bytes, not the service's sometimes incorrect .ics/.vcf filename/MIME label."""
    from email.parser import BytesHeaderParser

    if content.startswith(bytes.fromhex("d0cf11e0a1b11ae1")):
        return "msg", "application/vnd.ms-outlook"
    stripped = content.lstrip()
    if stripped.startswith(b"BEGIN:VCALENDAR"):
        return "icalendar", "text/calendar"
    if stripped.startswith(b"BEGIN:VCARD"):
        return "vcard", "text/vcard"
    headers = BytesHeaderParser().parsebytes(content[:65536])
    if headers.get("MIME-Version") and headers.get("Content-Type"):
        return "mime", "message/rfc822"
    changed("outlook.attachment.download", "native_format")
