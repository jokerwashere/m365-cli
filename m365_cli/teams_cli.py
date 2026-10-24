"""Teams argument registration, local validation, and dispatch."""

from .errors import OutlookError
from .operations import execute
from .session import endpoint
from .teams_api import TeamsAPI, identifier
from .teams_session import TeamsSession
from .validation import email, nonempty, paging


def register(groups, body_options, write_options, page_options):
    teams = groups.add_parser("teams", help="Teams people, chats, channels, and messaging")
    teams.add_argument("--tenant", help="Expected tenant UUID; defaults to the saved Teams tenant")
    commands = teams.add_subparsers(dest="command", required=True)
    commands.add_parser("status", help="Verify Teams account and tenant")
    login = commands.add_parser(
        "login", help="Open Teams in the persistent browser and await sign-in"
    )
    login.add_argument("--timeout", type=int, default=300)
    people = commands.add_parser("person").add_subparsers(dest="teams_action", required=True)
    p = people.add_parser("search", help="Find people by name or email (not message search)")
    p.add_argument("query")
    p = people.add_parser("get", help="Resolve one exact email address")
    p.add_argument("email")
    chats = commands.add_parser("chat").add_subparsers(dest="teams_action", required=True)
    p = chats.add_parser("list")
    page_options(p)
    p.add_argument("--kind", choices=("direct", "group", "meeting"))
    chats.add_parser("self", help="Get your self-chat ID")
    p = chats.add_parser("resolve", help="Find an existing direct chat by exact email")
    p.add_argument("email")
    p = chats.add_parser("get")
    p.add_argument("id")
    p = chats.add_parser("members", help="Read the chat roster")
    p.add_argument("id")
    p = chats.add_parser("create", help="Reuse/create a direct chat, or create a named group")
    p.add_argument(
        "--to", action="append", default=[], help="Exact email; repeat for group members"
    )
    p.add_argument("--group", action="store_true", dest="group_chat")
    p.add_argument("--topic", help="Required with --group")
    write_options(p)
    p = chats.add_parser("rename", help="Rename a group chat")
    p.add_argument("id")
    p.add_argument("--topic", required=True)
    write_options(p)
    p = chats.add_parser("leave", help="Leave a group chat; does not delete it for others")
    p.add_argument("id")
    write_options(p)
    team = commands.add_parser("team").add_subparsers(dest="teams_action", required=True)
    p = team.add_parser("list")
    page_options(p)
    channel = commands.add_parser("channel").add_subparsers(dest="teams_action", required=True)
    p = channel.add_parser("list")
    p.add_argument("--team", required=True)
    page_options(p)
    p = channel.add_parser("get")
    p.add_argument("id")
    messages = commands.add_parser("message").add_subparsers(dest="teams_action", required=True)
    p = messages.add_parser("render", help="Preview a text/HTML/Markdown Teams body locally")
    body_options(p, formatted=True)
    p.add_argument("--output", help="Save a new HTML preview file (never overwrites)")
    messages.add_parser("styles", help="Show shared style defaults and Teams message limits")
    for name in ("list", "get", "send", "reply", "update", "delete"):
        p = messages.add_parser(
            name,
            help={
                "reply": "Reply to a channel root post",
                "delete": "Soft-delete your message (no atomic version guard)",
            }.get(name),
        )
        target = p.add_mutually_exclusive_group(required=True)
        target.add_argument("--chat", help="Chat ID")
        target.add_argument("--channel", help="Standard channel ID")
        if name in ("get", "update", "delete"):
            p.add_argument("id", help="Message ID")
        if name == "list":
            p.add_argument("--limit", type=int, default=25)
            p.add_argument("--cursor", help="Opaque next_cursor from the previous page")
            p.add_argument("--reply-to", help="Channel root message ID to list replies")
        if name == "reply":
            p.add_argument("--reply-to", required=True, help="Channel root message ID")
        if name in ("send", "reply", "update"):
            body_options(p, formatted=True)
        if name in ("send", "reply", "update", "delete"):
            write_options(p)
        if name in ("update", "delete"):
            p.add_argument(
                "--if-version",
                help="Reviewed edit version; rejected for deletion because Teams ignores it",
            )


def prepare(args):
    import sys
    from pathlib import Path

    if args.account:
        email(args.account)
    if args.tenant:
        import uuid

        try:
            uuid.UUID(args.tenant)
        except (ValueError, AttributeError):
            raise OutlookError("invalid_argument", "Tenant must be a UUID.") from None
    if hasattr(args, "limit"):
        paging(args.limit, getattr(args, "offset", 0))
    if hasattr(args, "timeout") and not 1 <= args.timeout <= 3600:
        raise OutlookError("invalid_argument", "Timeout must be 1–3600 seconds.")
    for key in ("id", "chat", "channel", "team", "reply_to", "query", "topic"):
        if getattr(args, key, None) is not None:
            identifier(getattr(args, key), key)
    if getattr(args, "email", None):
        email(args.email)
    for recipient in getattr(args, "to", []):
        email(recipient)
    if getattr(args, "request_id", None) is not None:
        nonempty(args.request_id, "request_id")
        if len(args.request_id) > 200:
            raise OutlookError("invalid_argument", "Request ID is limited to 200 characters.")
    args.text = getattr(args, "body", None)
    if getattr(args, "body_file", None):
        try:
            from .composition import MAX_SOURCE

            if args.body_file == "-":
                args.text = sys.stdin.read(MAX_SOURCE + 1)
            else:
                with Path(args.body_file).open(encoding="utf-8") as stream:
                    args.text = stream.read(MAX_SOURCE + 1)
        except (OSError, UnicodeError):
            raise OutlookError("input_file_error", "Cannot read the UTF-8 body file.") from None
    action = getattr(args, "teams_action", None)
    if args.command == "message":
        if action in ("send", "reply", "update", "render"):
            from .teams_composition import prepare as prepare_body

            prepare_body(args)
        if getattr(args, "reply_to", None) and not args.channel:
            raise OutlookError(
                "unsupported_operation", "Native replies currently support channel posts only."
            )
        if getattr(args, "if_version", None) is not None:
            if action == "delete":
                raise OutlookError(
                    "unsupported_precondition",
                    "Teams ignores If-Match on deletion; atomic conditional deletion is unsupported.",
                )
            if not args.if_version.isdecimal():
                raise OutlookError("invalid_argument", "Message version must be a numeric string.")
    if args.command == "chat" and action == "create":
        if args.group_chat:
            nonempty(args.topic, "topic")
        elif len(args.to) != 1 or args.topic:
            raise OutlookError(
                "invalid_argument",
                "A direct chat needs exactly one --to; named groups need --group --topic.",
            )
    return args


def dispatch(api, args):
    action = getattr(args, "teams_action", None)
    if args.command == "person":
        return api.people(args.query) if action == "search" else api.person(args.email)
    if args.command == "team":
        return api.teams(args.limit, args.offset)
    if args.command == "channel":
        return (
            api.channels(args.team, args.limit, args.offset)
            if action == "list"
            else api.channel(args.id)
        )
    if args.command == "chat":
        if action == "list":
            return api.chats(args.limit, args.offset, args.kind)
        if action == "self":
            return api.self_chat()
        if action == "resolve":
            return api.resolve_chat(args.email)
        if action in ("get", "members"):
            chat = api.chat(args.id)
            return {"items": chat["members"], "thread_id": args.id} if action == "members" else chat
        if action == "create":
            return api.create_chat(args.to, args.topic, args.group_chat)
        if action == "rename":
            return api.rename_chat(args.id, args.topic)
        return api.leave_chat(args.id)
    thread = args.chat or args.channel
    channel = bool(args.channel)
    if action == "list":
        return api.messages(
            thread, channel=channel, limit=args.limit, cursor=args.cursor, reply_to=args.reply_to
        )
    if action == "get":
        return api.message(thread, args.id, channel)
    if action in ("send", "reply"):
        return api.send(
            thread,
            args.text,
            channel=channel,
            reply_to=getattr(args, "reply_to", None),
            body_type=args.body_type,
        )
    if action == "update":
        return api.update(
            thread,
            args.id,
            args.text,
            channel=channel,
            expected=args.if_version,
            body_type=args.body_type,
        )
    return api.delete(thread, args.id, channel=channel, expected=args.if_version)


def run(args, plan):
    if args.command == "message" and args.teams_action in ("render", "styles"):
        from .teams_composition import preview, styles

        return preview(args) if args.teams_action == "render" else styles()
    if args.command == "status" and not endpoint():
        return {"authenticated": False, "browser_running": False}
    with TeamsSession(launch=args.command == "login") as session:
        info = session.authenticate(
            expected=args.account,
            tenant=args.tenant,
            wait=args.timeout if args.command == "login" else 30,
        )
        if args.command in ("login", "status"):
            return info
        api = TeamsAPI(session)
        from .references import References

        refs = References(session.account, session.tenant)

        def action():
            result = refs.teams(dispatch(api, args), args, api)
            if getattr(args, "composition", None) and args.body_type == "HTML":
                result["composition"] = args.composition
            return result

        if args.mutation:
            payload = {"service": "teams", "tenant": session.tenant, "command": plan(args)}
            return execute(args.request_id, session.account, payload, action)
        return action()
