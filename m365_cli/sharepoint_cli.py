"""Files and lists CLI: explicit site scope, stable handles and bounded local files."""

import json
import os
from pathlib import Path

from .errors import OutlookError
from .operations import execute
from .references import get as reference_get
from .session import read_config
from .sharepoint_api import SharePointAPI
from .sharepoint_session import SharePointSession
from .sharepoint_validation import (
    MAX_UPLOAD,
    etag,
    filename,
    guid,
    infer_site,
    item_id,
    site_url,
    text,
    version,
)
from .validation import email


def register(groups, references, write_options):
    def command(sub, name, help, *, write=False, guard=False, unconditional=False, paged=False):
        p = sub.add_parser(name, help=help)
        p.set_defaults(sharepoint=True)
        p.add_argument(
            "--site",
            help="SharePoint site root; required unless a stable handle or URL supplies it",
        )
        if write:
            write_options(p)
        if guard:
            p.add_argument(
                "--if-version", help="Exact quoted ETag from the matching get/metadata command"
            )
        if unconditional:
            p.add_argument(
                "--allow-unconditional",
                action="store_true",
                help="Required: endpoint has no verified atomic concurrency guard",
            )
        if paged:
            p.add_argument(
                "--limit", type=int, default=50, help="Page size 1–500; follow next_cursor"
            )
            p.add_argument(
                "--cursor", help="Opaque cursor from this exact account/site/query/limit"
            )
        return p

    def ident(p, kind):
        p.add_argument("id", help="Native ID or m365: stable handle")
        p.set_defaults(sp_kind=kind)

    def data(p):
        group = p.add_mutually_exclusive_group(required=True)
        group.add_argument("--data", help="JSON object using schema internal field names")
        group.add_argument("--data-file", help="UTF-8 JSON object file, maximum 1 MiB")

    files = groups.add_parser(
        "files", help="SharePoint/OneDrive files; browser session, explicit site scope"
    ).add_subparsers(dest="command", required=True)
    p = command(files, "login", "Verify SharePoint sign-in in the saved browser")
    p.add_argument("--timeout", type=int, default=300)
    command(files, "status", "Verify account, tenant and site access")
    command(files, "capabilities", "Local file/list boundaries; no live availability claim")
    command(
        files, "libraries", "List visible document libraries and their root folder IDs", paged=True
    )
    for sub in (files, references):
        p = command(
            sub, "resolve", "Resolve an existing-access SharePoint file/folder/list-item URL"
        )
        p.add_argument("url")
    p = command(files, "search", "Site-scoped indexed literal phrase search; incomplete by design")
    p.add_argument("query")
    p.add_argument("--limit", type=int, default=25)
    p.add_argument("--offset", type=int, default=0)
    for name in (
        "get",
        "metadata",
        "versions",
        "download",
        "fetch",
        "rename",
        "copy",
        "move",
        "recycle",
        "restore",
    ):
        p = command(
            files,
            name,
            {
                "get": "Read file metadata and stable URL",
                "metadata": "Read the list-item ETag used by rename",
                "versions": "Read historical versions; current version is separate",
                "download": "Save bytes to a new local path (max 50 MiB)",
                "fetch": "Download bytes plus a reference/version/hash manifest",
                "rename": "Rename with an atomic metadata ETag guard",
                "copy": "Copy a verified byte snapshot to a new name (max 10 MiB)",
                "move": "Move within one library, without overwrite",
                "recycle": "Move a file to the recycle bin",
                "restore": "Restore historical content as a new current version",
            }[name],
            write=name in ("rename", "copy", "move", "recycle", "restore"),
            guard=name == "rename",
            unconditional=name in ("move", "recycle", "restore"),
            paged=name == "versions",
        )
        ident(p, "file")
        if name in ("download", "fetch"):
            p.add_argument("--output", required=True, help="New local file path; never overwritten")
        if name in ("download", "fetch", "copy", "restore"):
            p.add_argument(
                "--version",
                required=name == "restore",
                help="Historical version label; omit for current bytes",
            )
        if name in ("rename", "copy", "move"):
            p.add_argument("--name", required=name in ("rename", "copy"))
        if name in ("copy", "move"):
            p.add_argument(
                "--folder",
                required=True,
                help="Destination folder GUID or stable handle in the same site",
            )
    p = command(files, "upload", "Create a new file; name collision fails (max 10 MiB)", write=True)
    p.add_argument("--folder", required=True)
    p.add_argument("--file", required=True, help="Local file containing bytes")
    p.add_argument("--name", help="Destination name; defaults to local filename")
    folders = files.add_parser(
        "folder", help="Folder metadata and nonrecursive operations"
    ).add_subparsers(dest="sp_action", required=True)
    for name in ("get", "metadata", "list", "create", "rename", "recycle"):
        p = command(
            folders,
            name,
            name.capitalize() + " folder",
            write=name in ("create", "rename", "recycle"),
            guard=name == "rename",
            unconditional=name == "recycle",
            paged=name == "list",
        )
        if name == "create":
            p.add_argument("--folder", required=True, help="Parent folder GUID or stable handle")
        else:
            ident(p, "folder")
        if name in ("create", "rename"):
            p.add_argument("--name", required=True)
        if name == "list":
            p.add_argument(
                "--kind",
                choices=("file", "folder"),
                default="file",
                help="Enumerate files or subfolders separately",
            )
    p = command(files, "list", "List files or subfolders in one folder", paged=True)
    p.add_argument("--folder", required=True)
    p.add_argument("--kind", choices=("file", "folder"), default="file")

    lists = groups.add_parser(
        "lists", help="SharePoint lists, typed records, history and attachments"
    ).add_subparsers(dest="command", required=True)
    command(lists, "list", "List visible lists and libraries", paged=True)
    p = command(lists, "users", "Site-scoped person IDs for typed fields", paged=True)
    p.add_argument("--email", help="Exact email filter")
    for name in ("get", "schema", "create", "rename", "recycle"):
        p = command(
            lists,
            name,
            name.capitalize()
            + " list"
            + (" (visible generic lists only)" if name in ("create", "rename", "recycle") else ""),
            write=name in ("create", "rename", "recycle"),
            guard=name == "rename",
            unconditional=name == "recycle",
        )
        if name != "create":
            ident(p, "list")
        if name in ("create", "rename"):
            p.add_argument("--name", required=True)
    columns = lists.add_parser(
        "column", help="Add a new supported column; no destructive schema edits"
    ).add_subparsers(dest="sp_action", required=True)
    column = command(
        columns, "create", "Create a simple column on a visible generic list", write=True
    )
    column.add_argument("--list", required=True)
    column.add_argument(
        "--name", required=True, help="ASCII internal name, 1–32 letters/digits/underscores"
    )
    column.add_argument(
        "--type",
        choices=(
            "text",
            "choice",
            "person",
            "number",
            "boolean",
            "datetime",
            "url",
            "multi-choice",
            "multi-person",
        ),
        required=True,
    )
    column.add_argument("--choice", action="append", help="Repeat for each allowed choice")
    column.add_argument("--required", action="store_true", help="Only allowed on an empty list")
    for section, names in (
        ("item", ("list", "get", "create", "update", "recycle", "versions", "version")),
        ("attachment", ("list", "get", "download", "add", "recycle")),
        ("comments", ("list",)),
    ):
        sub = lists.add_parser(section).add_subparsers(dest="sp_action", required=True)
        for name in names:
            p = command(
                sub,
                name,
                name.capitalize() + " " + section,
                write=name in ("create", "update", "add", "recycle"),
                guard=name == "update",
                unconditional=name == "recycle",
                paged=name in ("list", "versions"),
            )
            p.add_argument(
                "--list", help="List GUID or stable handle; required unless item handle supplies it"
            )
            if section != "item" or name not in ("list", "create"):
                ident(p, "list_item")
            if section == "item" and name in ("create", "update"):
                data(p)
            if name == "version":
                p.add_argument("--version", required=True)
            if section == "attachment":
                if name not in ("list", "add"):
                    p.add_argument(
                        "--name", required=True, help="Exact attachment name on the record"
                    )
                if name == "add":
                    p.add_argument("--file", required=True)
                    p.add_argument("--name")
                if name == "download":
                    p.add_argument("--output", required=True)


def capabilities():
    return {
        "implemented": True,
        "live_availability": "not_checked",
        "transport": "direct HTTPS after browser bootstrap; scoped cookies and in-memory request digest",
        "limits": {
            "new_upload_bytes": MAX_UPLOAD,
            "download_bytes": 50 * 1024 * 1024,
            "max_page_size": 500,
        },
        "guarded": [
            "new-file name collision",
            "record update ETag",
            "file/folder rename metadata ETag",
            "list rename ETag",
        ],
        "unconditional_opt_in": ["move", "recycle", "file version restore"],
        "unsupported": [
            "file content replacement/publish",
            "cross-library move",
            "resumable/large upload",
            "sharing/permission changes",
            "list version restore",
            "comment writes",
            "complex field writes",
            "recursive deletion",
            "permanent deletion",
            "destructive schema mutation",
        ],
        "links": "Stable local handles plus canonical provider URLs; versions/comments link to their parent. No grant/redeem operation.",
    }


def prepare(args):
    if args.account:
        email(args.account)
    if args.command == "capabilities":
        return args
    sites, tenants = [], []
    if args.site:
        sites.append(site_url(args.site))
    # Resolve durable handles locally before browser access and fingerprint native locators.
    for key, kind in (
        ("list", "list"),
        ("folder", "folder"),
        ("id", getattr(args, "sp_kind", None)),
    ):
        value = getattr(args, key, None)
        if not value:
            continue
        if value.startswith("m365:"):
            ref = reference_get(value, args.account)
            if (
                ref.get("source") != "sharepoint"
                or ref.get("resource") != kind
                or ref.get("status") in ("deleted", "parent_deleted")
            ):
                raise OutlookError(
                    "invalid_reference",
                    "Use an active SharePoint handle of the required resource kind.",
                )
            locator = ref["locator"]
            sites.append(site_url(locator["site_url"]))
            tenants.append(ref["tenant"])
            value = locator["id"]
            if kind == "list_item":
                if (
                    getattr(args, "list", None)
                    and not args.list.startswith("m365:")
                    and guid(args.list) != guid(locator["list_id"])
                ):
                    raise OutlookError("invalid_reference", "Item handle and --list disagree.")
                args.list = locator["list_id"]
            setattr(args, key, value)
        (item_id if kind == "list_item" else guid)(value)
    if hasattr(args, "url"):
        if not sites:
            sites.append(infer_site(args.url))
        text(args.url, "URL", 8192)
    if not sites or len(set(sites)) != 1 or len(set(tenants)) > 1:
        raise OutlookError(
            "invalid_argument",
            "Supply one --site; every handle and destination must belong to that same site.",
        )
    args.site = sites[0]
    args.sp_tenant = tenants[0] if tenants else None
    if hasattr(args, "list") and not args.list:
        raise OutlookError(
            "invalid_argument", "--list is required without a list-item stable handle."
        )
    if getattr(args, "email", None):
        email(args.email)
    for key in ("if_version", "version"):
        if getattr(args, key, None) is not None:
            (etag if key == "if_version" else version)(getattr(args, key))
    if getattr(args, "allow_unconditional", None) is False:
        raise OutlookError(
            "unguarded_operation",
            "Explicit --allow-unconditional is required: this endpoint has no verified atomic version guard.",
        )
    if hasattr(args, "limit") and not 1 <= args.limit <= (100 if args.command == "search" else 500):
        raise OutlookError("invalid_argument", "Page size is outside the supported range.")
    if hasattr(args, "offset") and not 0 <= args.offset <= 10000:
        raise OutlookError("invalid_argument", "Search offset must be 0–10000.")
    if hasattr(args, "timeout") and not 0 <= args.timeout <= 3600:
        raise OutlookError("invalid_argument", "Login timeout must be 0–3600 seconds.")
    if getattr(args, "request_id", None):
        text(args.request_id, "request ID", 200)
    if getattr(args, "query", None):
        text(args.query, "query", 1000)
    if getattr(args, "file", None):
        try:
            with Path(args.file).expanduser().open("rb") as stream:
                args.file_content = stream.read(MAX_UPLOAD + 1)
        except OSError:
            raise OutlookError("invalid_file", "Cannot read the local upload file.") from None
        if len(args.file_content) > MAX_UPLOAD:
            raise OutlookError("file_too_large", "New uploads are limited to 10 MiB.")
        args.name = args.name or Path(args.file).name
    if getattr(args, "name", None):
        (text if args.group == "lists" and args.command in ("create", "rename") else filename)(
            args.name
        )
    if hasattr(args, "data"):
        try:
            raw = args.data
            if args.data_file:
                with Path(args.data_file).expanduser().open("rb") as stream:
                    raw = stream.read(1024 * 1024 + 1).decode("utf-8")
            if len(raw.encode()) > 1024 * 1024:
                raise ValueError
            args.data = json.loads(
                raw, parse_constant=lambda _: (_ for _ in ()).throw(ValueError())
            )
            if not isinstance(args.data, dict) or not args.data:
                raise ValueError
        except (OSError, ValueError, TypeError, UnicodeError):
            raise OutlookError(
                "invalid_argument", "Use a nonempty UTF-8 JSON object no larger than 1 MiB."
            ) from None
    if getattr(args, "output", None):
        args.output = str(Path(args.output).expanduser().absolute())
        outputs = [args.output] + ([args.output + ".m365.json"] if args.command == "fetch" else [])
        if any(os.path.lexists(p) or not Path(p).parent.is_dir() for p in outputs):
            raise OutlookError(
                "output_exists",
                "Output and manifest paths must be new files in an existing directory.",
            )
    return args


def save_download(args, result):
    snapshot, content = result
    paths = []
    try:
        for path, body in [(args.output, content)] + (
            [
                (
                    args.output + ".m365.json",
                    json.dumps(
                        {
                            "manifest_version": 1,
                            "account": args.account or read_config().get("account"),
                            "site": args.site,
                            **snapshot,
                        },
                        ensure_ascii=False,
                        indent=2,
                    ).encode(),
                )
            ]
            if args.command == "fetch"
            else []
        ):
            fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
            paths.append(path)
            with os.fdopen(fd, "wb") as stream:
                stream.write(body)
                stream.flush()
                os.fsync(stream.fileno())
    except OSError:
        # Only remove paths exclusively created by this call, never pre-existing files.
        for path in paths:
            try:
                Path(path).unlink()
            except OSError:
                pass
        raise OutlookError(
            "local_output_failed",
            "Could not create the local download and manifest; no remote write occurred.",
        ) from None
    return {
        **snapshot,
        "output": args.output,
        **({"manifest": args.output + ".m365.json"} if args.command == "fetch" else {}),
    }


def dispatch(api, a):
    paging = {k: getattr(a, k) for k in ("limit", "cursor") if hasattr(a, k)}
    expected = getattr(a, "if_version", None)
    allow = getattr(a, "allow_unconditional", False)
    c = a.command
    if c in ("login", "status"):
        return {
            "authenticated": True,
            "account": api.s.account,
            "tenant": api.s.tenant,
            "site": api.s.site,
            "site_id": api.s.site_id,
            "web_id": api.s.web_id,
            "user_id": api.s.user_id,
            "transport": api.s.transport_info(),
        }
    if c == "resolve":
        return api.resolve(a.url)
    if a.group == "files":
        if c == "libraries":
            return api.lists(libraries=True, **paging)
        if c == "search":
            return api.search(a.query, limit=a.limit, offset=a.offset)
        if c == "list":
            return api.children(a.folder, kind=a.kind, **paging)
        if c == "upload":
            return api.upload(a.folder, a.name, a.file_content)
        folder = c == "folder"
        c = a.sp_action if folder else c
        if c == "get":
            return api.file(a.id, folder=folder)
        if c == "metadata":
            return api.metadata(a.id, folder=folder)
        if c == "list":
            return api.children(a.id, kind=a.kind, **paging)
        if c == "create":
            return api.folder_create(a.folder, a.name)
        if c == "rename":
            return api.rename(a.id, a.name, expected=expected, folder=folder)
        if c == "recycle":
            return api.recycle_file(a.id, folder=folder, allow=allow)
        if c == "versions":
            return api.versions(a.id, **paging)
        if c in ("download", "fetch"):
            return save_download(a, api.download(a.id, label=a.version))
        if c == "copy":
            return api.copy(a.id, a.folder, a.name, label=a.version)
        if c == "move":
            return api.move(a.id, a.folder, a.name, allow=allow)
        if c == "restore":
            return api.restore_file(a.id, a.version, allow=allow)
    if c == "list":
        return api.lists(**paging)
    if c == "users":
        return api.users(a.email, **paging)
    if c == "get":
        return api.list_info(a.id)
    if c == "schema":
        return api.schema(a.id)
    if c == "create":
        return api.create_list(a.name)
    if c == "rename":
        return api.rename_list(a.id, a.name, expected=expected)
    if c == "recycle":
        return api.recycle_list(a.id, allow=allow)
    action = a.sp_action
    if c == "column":
        return api.create_column(a.list, a.name, a.type, choices=a.choice, required=a.required)
    if c == "item":
        if action == "list":
            return api.records(a.list, **paging)
        if action == "get":
            return api.record(a.list, a.id)
        if action == "create":
            return api.create_record(a.list, a.data)
        if action == "update":
            return api.update_record(a.list, a.id, a.data, expected=expected)
        if action == "recycle":
            return api.recycle_record(a.list, a.id, allow=allow)
        if action == "versions":
            return api.record_versions(a.list, a.id, **paging)
        if action == "version":
            return api.record_version(a.list, a.id, a.version)
    if c == "comments":
        return api.comments(a.list, a.id, **paging)
    if c == "attachment":
        if action == "list":
            return api.attachments(a.list, a.id, **paging)
        if action == "get":
            return api.attachment(a.list, a.id, a.name)
        if action == "download":
            return save_download(a, api.attachment_download(a.list, a.id, a.name))
        if action == "add":
            return api.attachment_add(a.list, a.id, a.name, a.file_content)
        if action == "recycle":
            return api.attachment_delete(a.list, a.id, a.name, allow=allow)
    raise OutlookError("unsupported_operation", "Unsupported SharePoint operation.")


def run(args, plan):
    if args.command == "capabilities":
        return capabilities()
    with SharePointSession(
        args.site,
        expected=args.account,
        login=args.command == "login",
        wait=getattr(args, "timeout", 30),
    ) as session:
        if args.sp_tenant and args.sp_tenant != session.tenant:
            raise OutlookError("tenant_mismatch", "Stable handle belongs to another tenant.")
        api = SharePointAPI(session)

        def action():
            try:
                return api.finish(dispatch(api, args))
            except (KeyError, TypeError, ValueError, AttributeError):
                raise OutlookError(
                    "write_outcome_unknown" if session.write_attempted else "api_changed",
                    "SharePoint returned unexpected resource metadata. Inspect any attempted write before retrying.",
                ) from None

        payload = {**plan(args), "verified_tenant": session.tenant, "verified_web": session.web_id}
        return (
            execute(args.request_id, session.account, payload, action)
            if args.mutation
            else action()
        )
