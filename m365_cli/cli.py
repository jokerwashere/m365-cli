"""JSON-first command-line interface for Microsoft 365 mail, calendar, Teams and tasks."""

import argparse
import hashlib
import json
import mimetypes
import os
import re
import sys
from pathlib import Path

from . import __version__
from .api import API, attendee_list, recipients
from .errors import OutlookError
from .operations import execute, lookup
from .session import Session, command_lock, endpoint, read_config, save_config
from .validation import email, iso, nonempty, paging, validate_range


class Parser(argparse.ArgumentParser):
    def _parse_optional(self, arg_string):
        # Planner's opaque, 28-character base64url IDs may start with one or two
        # hyphens, including as the value of --plan/--task. Known flags win.
        if (
            self.get_default("task_source") == "planner"
            and arg_string not in self._option_string_actions
            and re.fullmatch(r"[A-Za-z0-9_-]{28}", arg_string)
        ):
            return None
        return super()._parse_optional(arg_string)

    def error(self, message):
        raise OutlookError("usage_error", message)


def write_options(p, *, change=False):
    p.add_argument("--request-id", help="Stable unique ID for duplicate prevention across retries")
    p.add_argument(
        "--dry-run",
        action="store_true",
        help="Validate local input and print the plan without connecting or writing",
    )
    if change:
        p.add_argument(
            "--if-change-key", help="Reject if the item changed since this change key was read"
        )
    p.set_defaults(mutation=True)


def body_options(p, *, formatted=False):
    group = p.add_mutually_exclusive_group()
    group.add_argument("--body")
    group.add_argument("--body-file", help="UTF-8 text file, or - for stdin")
    if formatted:
        from .composition import options

        options(p)


def recipient_options(p, *, clear=False):
    for name in ("to", "cc", "bcc"):
        group = p.add_mutually_exclusive_group()
        group.add_argument(
            "--" + name, action="append", help="Plain email address; repeat for multiple recipients"
        )
        if clear:
            group.add_argument(
                "--clear-" + name, action="store_true", help="Remove all recipients in this field"
            )


def page_options(p, limit=25):
    p.add_argument("--limit", type=int, default=limit)
    p.add_argument("--offset", type=int, default=0)


def parser():
    p = Parser(
        prog="m365",
        description="Microsoft 365 mail, calendar, Teams, tasks, files and lists CLI. JSON output; direct HTTP with browser authentication.",
    )
    p.add_argument("--version", action="version", version=f"m365-cli {__version__}")
    p.add_argument("--account", help="Expected signed-in mailbox; mismatch fails before writing")
    p.set_defaults(mutation=False)
    sub = p.add_subparsers(dest="group", required=True)
    auth = sub.add_parser("auth").add_subparsers(dest="command", required=True)
    login = auth.add_parser(
        "login", help="Open the dedicated Chrome profile and verify the account"
    )
    login.add_argument("--account", default=argparse.SUPPRESS)
    login.add_argument("--timeout", type=int, default=300)
    auth.add_parser("status")
    operation = sub.add_parser(
        "operation", help="Inspect local duplicate-prevention records"
    ).add_subparsers(dest="command", required=True)
    operation.add_parser("get").add_argument("request_id")
    ref = sub.add_parser(
        "reference", help="Durable local references and SharePoint URL resolution"
    ).add_subparsers(dest="command", required=True)
    ref.add_parser("get", help="Return the latest locally observed ID and native URL").add_argument(
        "stable_id"
    )
    mail = sub.add_parser("mail").add_subparsers(dest="command", required=True)
    listing = mail.add_parser("list")
    page_options(listing)
    listing.add_argument("--folder", default="inbox")
    search = mail.add_parser("search")
    search.add_argument("query")
    page_options(search)
    get = mail.add_parser("get")
    get.add_argument("id")
    get.add_argument("--body-format", choices=("text", "html"), default="text")
    render = mail.add_parser("render", help="Render a body locally; never connects or sends")
    body_options(render, formatted=True)
    render.add_argument("--output", help="Save a new HTML preview file (never overwrites)")
    mail.add_parser("styles", help="Show local email style defaults and override fields")
    for name in ("send",):
        cmd = mail.add_parser(name, help="Send immediately and save a copy in Sent Items")
        cmd.add_argument("--subject", required=True)
        body_options(cmd, formatted=True)
        recipient_options(cmd)
        write_options(cmd)
    for name in ("reply", "forward"):
        cmd = mail.add_parser(name, help="Create a response draft; --send sends immediately")
        cmd.add_argument("id")
        body_options(cmd, formatted=True)
        write_options(cmd, change=True)
        cmd.add_argument(
            "--to", action="append", help="Override reply recipients, or specify forward recipients"
        )
        cmd.add_argument("--send", action="store_true")
        if name == "reply":
            cmd.add_argument("--all", action="store_true", help="Reply to all original recipients")
    update = mail.add_parser("update", help="Update read state, importance, categories, or flag")
    update.add_argument("id")
    write_options(update, change=True)
    update.add_argument("--read", choices=("true", "false"))
    update.add_argument("--importance", choices=("Low", "Normal", "High"))
    update.add_argument("--flag", choices=("NotFlagged", "Flagged", "Complete"))
    cats = update.add_mutually_exclusive_group()
    cats.add_argument("--category", action="append")
    cats.add_argument("--clear-categories", action="store_true")
    delete = mail.add_parser("delete", help="Move mail to Deleted Items; never permanently delete")
    delete.add_argument("id")
    write_options(delete, change=True)
    for name in ("move", "copy"):
        cmd = mail.add_parser(name)
        cmd.add_argument("id")
        cmd.add_argument("--folder", required=True)
        write_options(cmd, change=True)
    drafts = mail.add_parser("draft").add_subparsers(dest="draft_command", required=True)
    for name in ("create", "update"):
        cmd = drafts.add_parser(name)
        if name == "update":
            cmd.add_argument("id")
        cmd.add_argument("--subject", required=name == "create")
        body_options(cmd, formatted=True)
        recipient_options(cmd, clear=name == "update")
        write_options(cmd, change=name == "update")
    for name in ("get", "delete", "send"):
        cmd = drafts.add_parser(name)
        cmd.add_argument("id")
        if name != "get":
            write_options(cmd, change=True)
        else:
            cmd.add_argument("--body-format", choices=("text", "html"), default="text")
    attachments = mail.add_parser("attachment").add_subparsers(
        dest="attachment_command", required=True
    )
    listing = attachments.add_parser(
        "list", help="List typed attachments, optionally through nested items"
    )
    listing.add_argument("id", help="Root mail ID")
    listing.add_argument(
        "--recursive",
        action="store_true",
        help="Traverse attached messages/items; inspect complete and truncations",
    )
    listing.add_argument(
        "--max-depth",
        type=int,
        default=10,
        help="Maximum nested depth, 0–30; root attachments have depth 0",
    )
    listing.add_argument(
        "--max-items", type=int, default=200, help="Maximum attachment entries, 1–1000"
    )
    add = attachments.add_parser(
        "add", help="Attach a local file or a copy of an existing mailbox item to a draft"
    )
    add.add_argument("id")
    source = add.add_mutually_exclusive_group(required=True)
    source.add_argument("--file")
    source.add_argument(
        "--item-id", help="Existing mailbox item to attach, including its attachments"
    )
    add.add_argument("--name", help="Attachment display name (never used as a local output path)")
    add.add_argument("--content-type", help="File MIME type; inferred from the filename by default")
    add.add_argument(
        "--inline", action="store_true", help="Mark a file inline; requires --content-id"
    )
    add.add_argument("--content-id", help="Inline file CID, without angle brackets")
    write_options(add, change=True)
    get = attachments.add_parser(
        "get", help="Inspect any attachment, or export it to an explicit path"
    )
    get.add_argument("attachment_id")
    get.add_argument(
        "--mail-id", help="Root mail ID; verifies membership and supplies the parent URL"
    )
    get.add_argument(
        "--output",
        help="Exact output path; must not already exist. Omit to inspect JSON without file bytes",
    )
    get.add_argument(
        "--format",
        choices=("auto", "raw", "json"),
        default="auto",
        help="Export: auto saves files/messages as bytes and other types as JSON; raw requests native bytes; json saves properties",
    )
    delete = attachments.add_parser("delete")
    delete.add_argument("id")
    delete.add_argument("--attachment-id", required=True)
    write_options(delete, change=True)
    folders = sub.add_parser("folder", help="Mail and calendar folder management").add_subparsers(
        dest="command", required=True
    )
    listing = folders.add_parser("list")
    listing.add_argument("--parent", default="msgfolderroot")
    page_options(listing, 100)
    folders.add_parser("get").add_argument("id")
    create = folders.add_parser("create")
    create.add_argument("--name", required=True)
    create.add_argument("--parent", default="msgfolderroot")
    create.add_argument("--kind", choices=("mail", "calendar"), default="mail")
    write_options(create)
    update = folders.add_parser("update")
    update.add_argument("id")
    update.add_argument("--name", required=True)
    write_options(update, change=True)
    delete = folders.add_parser("delete", help="Delete only an empty custom folder")
    delete.add_argument("id")
    write_options(delete, change=True)
    cal = sub.add_parser("calendar").add_subparsers(dest="command", required=True)
    from .scheduling import register as register_scheduling

    register_scheduling(cal)
    cal.add_parser("calendars")
    for name in ("list", "search"):
        cmd = cal.add_parser(name)
        page_options(cmd, 100)
        if name == "search":
            cmd.add_argument(
                "query",
                help="Case-insensitive substring in subject or location within the time window",
            )
        cmd.add_argument("--start", required=True)
        cmd.add_argument("--end", required=True)
        cmd.add_argument("--folder", default="calendar")
    get = cal.add_parser("get")
    get.add_argument("id")
    get.add_argument("--body-format", choices=("text", "html"), default="text")
    render = cal.add_parser("render", help="Render an event body locally; never connects or books")
    body_options(render, formatted=True)
    render.add_argument("--output", help="Save a new HTML preview file (never overwrites)")
    cal.add_parser("styles", help="Show shared mail/event style defaults and override fields")
    for name in ("create", "update"):
        cmd = cal.add_parser(
            name,
            description=(
                "Create an event. Attendees receive invitations unless --save-only is supplied."
                if name == "create"
                else "Edit an event. Use --save-only for a never-sent meeting or --notify to send updates."
            ),
        )
        write_options(cmd, change=name == "update")
        delivery = cmd.add_mutually_exclusive_group()
        delivery.add_argument(
            "--save-only",
            action="store_true",
            help="Save attendees without invitations; updates require a verified never-sent single event",
        )
        if name == "update":
            cmd.add_argument("id")
            delivery.add_argument(
                "--notify", action="store_true", help="Send meeting updates to attendees"
            )
        else:
            cmd.add_argument("--folder", default="calendar")
        for field in ("subject", "start", "end"):
            cmd.add_argument("--" + field, required=name == "create")
        cmd.add_argument("--location")
        body_options(cmd, formatted=True)
        cmd.add_argument("--all-day", choices=("true", "false"))
        cmd.add_argument(
            "--show-as", choices=("Free", "Tentative", "Busy", "OOF", "WorkingElsewhere")
        )
        for field in ("attendee", "optional-attendee"):
            group = cmd.add_mutually_exclusive_group()
            group.add_argument("--" + field, action="append")
            if name == "update":
                group.add_argument("--clear-" + field + "s", action="store_true")
    send = cal.add_parser(
        "send", help="Send initial invitations for a saved, never-sent single event"
    )
    send.add_argument("id")
    write_options(send, change=True)
    delete = cal.add_parser("delete")
    delete.add_argument("id")
    delete.add_argument(
        "--notify",
        action="store_true",
        help="Send cancellations for sent meetings; verified never-sent meetings are deleted silently",
    )
    write_options(delete, change=True)
    respond = cal.add_parser("respond")
    respond.add_argument("id")
    respond.add_argument("--response", choices=("accept", "tentative", "decline"), required=True)
    respond.add_argument("--notify", action="store_true", help="Send the response to the organizer")
    body_options(respond)
    write_options(respond, change=True)
    from .teams_cli import register

    register(sub, body_options, write_options, page_options)
    from .labels_cli import register_category
    from .maintainer import register as register_maintainer
    from .tasks_cli import register as register_tasks

    register_category(
        sub.add_parser("category", help="Outlook/To Do category definitions and assignments"),
        write_options,
    )
    register_tasks(sub, body_options, write_options, page_options)
    from .sharepoint_cli import register as register_sharepoint

    register_sharepoint(sub, ref, write_options)
    register_maintainer(sub)
    return p


def prepare(args):
    if getattr(args, "sharepoint", False):
        from .sharepoint_cli import prepare as prepare_sharepoint

        return prepare_sharepoint(args)
    if getattr(args, "scheduling", False):
        from .scheduling import prepare as prepare_scheduling

        return prepare_scheduling(args)
    if hasattr(args, "label_domain"):
        from .labels_cli import prepare as prepare_labels

        return prepare_labels(args)
    if args.group in ("tasks", "todo", "planner"):
        from .tasks_cli import prepare as prepare_tasks

        return prepare_tasks(args)
    if args.group == "maintainer":
        return args
    if args.group == "teams":
        from .teams_cli import prepare as prepare_teams

        return prepare_teams(args)
    """Validate all local values before opening a browser or recording a write attempt."""
    if getattr(args, "limit", None) is not None:
        paging(args.limit, args.offset)
    if args.account:
        email(args.account)
    for key in (
        "id",
        "attachment_id",
        "mail_id",
        "item_id",
        "content_id",
        "content_type",
        "query",
        "folder",
        "name",
        "parent",
    ):
        if getattr(args, key, None) is not None:
            nonempty(getattr(args, key), key)
    if getattr(args, "request_id", None) is not None:
        nonempty(args.request_id, "request ID")
        if len(args.request_id) > 200:
            raise OutlookError("invalid_request_id", "Request IDs are limited to 200 characters.")
    if getattr(args, "timeout", 0) < 0 or getattr(args, "timeout", 0) > 3600:
        raise OutlookError("invalid_argument", "Login timeout must be 0–3600 seconds.")
    text = getattr(args, "body", None)
    try:
        if getattr(args, "body_file", None):
            limit = 1024 * 1024 + 1 if getattr(args, "body_composition", False) else -1
            if args.body_file == "-":
                text = sys.stdin.read(limit)
            else:
                with Path(args.body_file).open(encoding="utf-8") as stream:
                    text = stream.read(limit)
    except (OSError, UnicodeError):
        raise OutlookError("input_file_error", "Cannot read the body file as UTF-8 text.") from None
    args.text = text
    if getattr(args, "body_composition", False):
        from .composition import prepare as prepare_composition

        prepare_composition(args)
    for field in ("to", "cc", "bcc", "attendee", "optional_attendee"):
        for address in getattr(args, field, None) or []:
            email(address)
    if args.group == "calendar":
        for field in ("start", "end"):
            if getattr(args, field, None) is not None:
                iso(getattr(args, field))
        if getattr(args, "start", None) and getattr(args, "end", None):
            validate_range(
                args.start,
                args.end,
                window=args.command in ("list", "search"),
                all_day=getattr(args, "all_day", None) == "true",
            )
        if args.command == "create":
            nonempty(args.subject, "subject")
    if (
        args.group == "mail"
        and (args.command == "send" or args.command == "forward")
        and not any(getattr(args, field, None) for field in ("to", "cc", "bcc"))
    ):
        raise OutlookError("recipient_required", "At least one recipient is required.")
    if args.group == "mail" and args.command == "attachment":
        from .attachments import MAX_FILE_BYTES, traversal_limits

        if args.attachment_command == "list":
            traversal_limits(args.max_depth, args.max_items)
        if args.attachment_command == "add":
            if args.item_id and (args.inline or args.content_id or args.content_type):
                raise OutlookError(
                    "invalid_argument", "Inline/CID/MIME options apply only to --file."
                )
            if args.item_id == args.id:
                raise OutlookError("invalid_argument", "A draft cannot be attached to itself.")
            if bool(args.content_id) != args.inline:
                raise OutlookError("invalid_argument", "Use --inline and --content-id together.")
            if args.content_id and (
                len(args.content_id) > 998 or re.search(r"[\s<>]", args.content_id)
            ):
                raise OutlookError(
                    "invalid_argument",
                    "Content ID must contain no whitespace or angle brackets and be at most 998 characters.",
                )
            if args.content_type and (
                "/" not in args.content_type or re.search(r"[\r\n]", args.content_type)
            ):
                raise OutlookError("invalid_argument", "Specify a MIME type without line breaks.")
            if args.file is not None:
                nonempty(args.file, "attachment file path")
                try:
                    path = Path(args.file)
                    if not path.is_file():
                        raise OSError("not a regular file")
                    if path.stat().st_size > MAX_FILE_BYTES:
                        raise OutlookError(
                            "attachment_too_large",
                            "The CLI supports file attachments up to 10 MiB.",
                        )
                    with path.open("rb") as stream:
                        args.file_content = stream.read(MAX_FILE_BYTES + 1)
                    if len(args.file_content) > MAX_FILE_BYTES:
                        raise OutlookError(
                            "attachment_too_large",
                            "The CLI supports file attachments up to 10 MiB.",
                        )
                except OSError:
                    raise OutlookError(
                        "input_file_error", "Cannot read the attachment file."
                    ) from None
        if args.attachment_command == "get":
            if args.output is not None:
                nonempty(args.output, "output path")
            if not args.output and args.format != "auto":
                raise OutlookError(
                    "invalid_argument",
                    "--format requires --output; without --output, get returns structured metadata.",
                )
            if args.output and Path(args.output).exists():
                raise OutlookError(
                    "output_exists", "The output path already exists; choose a new filename."
                )
    if args.command == "update" or getattr(args, "draft_command", None) == "update":
        args.changes = changes(args)
        if not args.changes and args.group != "folder":
            raise OutlookError("no_changes", "Specify at least one field to update.")
    return args


def changes(args):
    result = {}
    for field in ("subject", "start", "end", "importance"):
        if getattr(args, field, None) is not None:
            result[field.capitalize()] = getattr(args, field)
    if args.text is not None:
        result["Body"] = {"BodyType": getattr(args, "body_type", "Text"), "Value": args.text}
    for name in ("to", "cc", "bcc"):
        if getattr(args, name, None) is not None or getattr(args, "clear_" + name, False):
            result[name.capitalize() + "Recipients"] = recipients(getattr(args, name, None))
    if getattr(args, "read", None) is not None:
        result["IsRead"] = args.read == "true"
    if getattr(args, "flag", None) is not None:
        result["Flag"] = {"FlagStatus": args.flag}
    if getattr(args, "category", None) or getattr(args, "clear_categories", False):
        result["Categories"] = args.category or []
    if getattr(args, "location", None) is not None:
        result["Location"] = {"DisplayName": args.location}
    if getattr(args, "all_day", None) is not None:
        result["IsAllDayEvent"] = args.all_day == "true"
    if getattr(args, "show_as", None) is not None:
        result["LegacyFreeBusyStatus"] = args.show_as
    for field, name in (
        ("attendee", "RequiredAttendees"),
        ("optional_attendee", "OptionalAttendees"),
    ):
        if getattr(args, field, None) is not None or getattr(args, "clear_" + field + "s", False):
            result[name] = attendee_list(getattr(args, field, None))
    return result


def dispatch(api, args, *, refs=None):
    expected = getattr(args, "if_change_key", None)
    if args.group == "folder":
        if args.command == "list":
            return api.folder_list(args.parent, args.limit, args.offset)
        if args.command == "get":
            return api.folder_get(args.id)
        if args.command == "create":
            return api.folder_create(args.name, parent=args.parent, kind=args.kind)
        if args.command == "update":
            return api.folder_update(args.id, args.name, expected=expected)
        return api.folder_delete(args.id, expected=expected)
    if args.group == "mail":
        if args.command == "list":
            return api.mail_list(args.folder, args.limit, args.offset)
        if args.command == "search":
            return api.mail_search(args.query, args.limit, args.offset)
        if args.command == "get":
            return api.get(args.id, body_type="HTML" if args.body_format == "html" else "Text")
        if args.command == "send":
            return api.mail_send(
                subject=args.subject,
                body=args.text or "",
                body_type=args.body_type,
                to=args.to,
                cc=args.cc,
                bcc=args.bcc,
            )
        if args.command in ("reply", "forward"):
            return api.respond_mail(
                args.id,
                body=args.text or "",
                body_type=args.body_type,
                mode="reply-all" if getattr(args, "all", False) else args.command,
                to=args.to,
                send=args.send,
                expected=expected,
            )
        if args.command == "update":
            return api.update(args.id, args.changes, metadata=True, expected=expected)
        if args.command == "delete":
            return api.delete(args.id, expected=expected)
        if args.command in ("move", "copy"):
            return api.move(args.id, args.folder, copy=args.command == "copy", expected=expected)
        if args.command == "attachment":
            if args.attachment_command == "list":
                return api.attachment_list(
                    args.id,
                    recursive=args.recursive,
                    max_depth=args.max_depth,
                    max_items=args.max_items,
                )
            if args.attachment_command == "add":
                if args.item_id:
                    return api.attachment_add_item(
                        args.id, args.item_id, name=args.name, expected=expected
                    )
                return api.attachment_add(
                    args.id,
                    args.name or Path(args.file).name,
                    args.file_content,
                    args.content_type
                    or mimetypes.guess_type(args.file)[0]
                    or "application/octet-stream",
                    expected=expected,
                    inline=args.inline,
                    content_id=args.content_id,
                )
            if args.attachment_command == "delete":
                return api.attachment_delete(args.id, args.attachment_id, expected=expected)
            if not args.output:
                return api.attachment_inspect(args.attachment_id, mail_id=args.mail_id)
            result = api.attachment_get(
                args.attachment_id, mail_id=args.mail_id, format=args.format
            )
            content = result.pop("content")
            if refs is not None:
                result = refs.outlook(result, args)
            if result["export_format"] == "json":
                content = (json.dumps(result, ensure_ascii=False, indent=2) + "\n").encode()
            try:
                with os.fdopen(
                    os.open(args.output, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600), "wb"
                ) as stream:
                    stream.write(content)
            except FileExistsError:
                raise OutlookError(
                    "output_exists", "The output path already exists; choose a new filename."
                ) from None
            except OSError:
                raise OutlookError(
                    "output_file_error",
                    "Could not save the attachment at the requested path. A partial file may remain.",
                ) from None
            return {**result, "path": str(Path(args.output).resolve()), "bytes": len(content)}
        if args.draft_command == "create":
            return api.draft_create(
                subject=args.subject,
                body=args.text or "",
                body_type=args.body_type,
                to=args.to,
                cc=args.cc,
                bcc=args.bcc,
            )
        if args.draft_command == "get":
            return api.get(args.id, body_type="HTML" if args.body_format == "html" else "Text")
        if args.draft_command == "send":
            return api.draft_send(args.id, expected=expected)
        if args.draft_command == "delete":
            return api.delete(args.id, draft_only=True, expected=expected)
        return api.update(args.id, args.changes, expected=expected)
    if args.command == "calendars":
        return api.calendars()
    if args.command in ("list", "search"):
        return api.calendar_list(
            args.start, args.end, args.limit, args.folder, args.offset, getattr(args, "query", None)
        )
    if args.command == "get":
        return api.get(
            args.id, calendar=True, body_type="HTML" if args.body_format == "html" else "Text"
        )
    if args.command == "delete":
        return api.delete(args.id, calendar=True, notify=args.notify, expected=expected)
    if args.command == "send":
        return api.calendar_send(args.id, expected=expected)
    if args.command == "respond":
        return api.calendar_respond(
            args.id, args.response, body=args.text or "", notify=args.notify, expected=expected
        )
    if args.command == "create":
        return api.calendar_create(
            subject=args.subject,
            start=args.start,
            end=args.end,
            body=args.text or "",
            body_type=getattr(args, "body_type", "Text"),
            location=args.location or "",
            attendees=args.attendee,
            optional=args.optional_attendee,
            folder=args.folder,
            all_day=args.all_day == "true",
            show_as=args.show_as or "Busy",
            save_only=args.save_only,
        )
    return api.update(
        args.id,
        args.changes,
        calendar=True,
        notify=args.notify,
        save_only=args.save_only,
        expected=expected,
    )


def plan(args):
    data = {
        k: v
        for k, v in vars(args).items()
        if k
        not in (
            "request_id",
            "dry_run",
            "mutation",
            "body_file",
            "file_content",
            "changes",
            "data_file",
            "body_composition",
            "style_file",
        )
    }
    # Preserve existing plain-text operation fingerprints. Rich bodies bind their
    # rendered content, body type and resolved styles, not just a style-file path.
    if getattr(args, "body_composition", False) and args.body_format == "text":
        for key in ("body_format", "body_type", "style", "composition"):
            data.pop(key, None)
    if hasattr(args, "file_content"):
        data["attachment_sha256"] = hashlib.sha256(args.file_content).hexdigest()
    # Adding an opt-in must not change fingerprints of previously recorded writes.
    if data.get("save_only") is False:
        data.pop("save_only")
    return data


def compatible_plans(args):
    # The short-lived raw calendar HTML implementation added these defaults to
    # plain-text fingerprints. Accept its records as well as the original shape,
    # only for literal text. Rich input must always bind the new rendered body.
    if (
        args.group == "calendar"
        and args.command in ("create", "update")
        and args.body_format == "text"
    ):
        return [{**plan(args), "body_format": "text", "body_type": "Text"}]
    return []


def run(args):
    prepare(args)
    from .maintainer import guard
    from .maintainer import run as run_maintainer

    guard(args)
    if args.group == "maintainer":
        return run_maintainer(args)
    if args.group in ("mail", "calendar") and args.command in ("render", "styles"):
        from .composition import preview, styles

        return preview(args) if args.command == "render" else styles()
    if args.group == "operation":
        return lookup(args.request_id)
    if args.group == "reference" and args.command == "get":
        from .references import get

        return get(args.stable_id, args.account)
    if getattr(args, "dry_run", False):
        return {"dry_run": True, "plan": plan(args), "server_validation_pending": True}
    if getattr(args, "sharepoint", False):
        from .sharepoint_cli import run as run_sharepoint

        return run_sharepoint(args, plan)
    if hasattr(args, "label_domain"):
        from .labels_cli import run as run_labels

        return run_labels(args, plan)
    if args.group in ("tasks", "todo", "planner"):
        from .tasks_cli import run as run_tasks

        return run_tasks(args, plan)
    if args.group == "teams":
        from .teams_cli import run as run_teams

        return run_teams(args, plan)
    if args.group == "auth":
        if args.command == "status" and not endpoint():
            return {"authenticated": False, "browser_running": False}
        with Session(launch=args.command == "login") as session:
            if args.command == "login":
                print("Complete sign-in in the dedicated Chrome window.", file=sys.stderr)
                status = session.authenticate(wait=args.timeout, expected=args.account)
                config = read_config()
                config["account"] = status["account"]
                save_config(config)
                return status
            return session.authenticate(expected=args.account)
    if not (args.account or read_config().get("account")):
        raise OutlookError(
            "account_required",
            "Run m365 auth login --account EMAIL to pin the mailbox before using commands.",
        )
    with Session() as session:
        session.authenticate(expected=args.account)
        api = API(session)
        from .references import References

        refs = References(session.account)

        def action():
            if getattr(args, "scheduling", False):
                from .scheduling import dispatch as scheduling_dispatch

                return scheduling_dispatch(api, args)
            result = refs.outlook(dispatch(api, args, refs=refs), args)
            if getattr(args, "composition", None) and args.body_type == "HTML":
                result["composition"] = args.composition
            return result

        if args.mutation:
            return execute(
                args.request_id,
                session.account,
                plan(args),
                action,
                compatible_payloads=compatible_plans(args),
            )
        return action()


def main():
    os.umask(0o077)
    try:
        args = parser().parse_args()
        if args.group == "maintainer" and args.command == "validate":
            # The suite owns its lock, or invokes child commands that acquire it.
            result = run(args)
        else:
            with command_lock():
                result = run(args)
        print(json.dumps({"ok": True, "data": result}, ensure_ascii=False, indent=2))
    except OutlookError as error:
        print(json.dumps({"ok": False, "error": error.as_dict()}, ensure_ascii=False))
        raise SystemExit(2 if error.code == "usage_error" else 1) from None
    except KeyboardInterrupt:
        print(
            json.dumps(
                {
                    "ok": False,
                    "error": {
                        "code": "interrupted",
                        "message": "Interrupted. Inspect any write before retrying.",
                    },
                }
            )
        )
        raise SystemExit(130) from None
    except BrokenPipeError:
        raise SystemExit(0) from None
    except Exception as error:
        print(
            json.dumps(
                {
                    "ok": False,
                    "error": {
                        "code": "runtime_error",
                        "message": f"{type(error).__name__}: operation failed. A write may have completed; inspect before retrying.",
                    },
                }
            )
        )
        raise SystemExit(1) from None


if __name__ == "__main__":
    main()
