"""Explicit maintainer tools. Normal user commands never trigger capability sweeps."""

import argparse
import hashlib
import json
import subprocess
import sys
from datetime import UTC, datetime, timedelta
from pathlib import Path

from . import __version__
from .errors import OutlookError
from .operations import atomic_json
from .session import Session, read_config, save_config, state_dir

CONTRACT_REVISION = 12


def register(groups):
    root = groups.add_parser(
        "maintainer",
        help="Maintainer-only capability inventory, probes, live validation and circuit breakers",
    )
    commands = root.add_subparsers(dest="command", required=True)
    commands.add_parser(
        "capabilities", help="Inventory of every implemented command; not a live availability claim"
    )
    p = commands.add_parser(
        "check", help="Read-only API and detached HTTP probes; writes remain untested"
    )
    p.add_argument(
        "--module", choices=("all", "outlook", "teams", "tasks", "sharepoint"), default="all"
    )
    p.add_argument("--site", help="Explicit SharePoint site for read probes")
    p.add_argument("--output", help="Save a content-free JSON report (must be a new path)")
    p.add_argument("--baseline", help="Compare with a prior maintainer-check report")
    p = commands.add_parser(
        "validate", help="Run an explicit disposable-fixture write suite; journals stay private"
    )
    p.add_argument(
        "--suite",
        choices=(
            "outlook",
            "teams",
            "tasks",
            "references",
            "attachments",
            "labels",
            "availability",
            "composition",
            "calendar-composition",
            "calendar-drafts",
            "teams-composition",
            "sharepoint",
        ),
        required=True,
    )
    p.add_argument(
        "--site", help="SharePoint suite: explicitly authorized personal OneDrive site root"
    )
    p.add_argument("--allow-test-writes", action="store_true", required=True)
    p.add_argument(
        "--recipient",
        help="Required for Outlook/composition/calendar-drafts; authorized test recipient",
    )
    p.add_argument(
        "--channel", help="Optional explicitly authorized Teams channel; self-chat by default"
    )
    p.add_argument(
        "--attendee",
        action="append",
        default=[],
        help="Optional read-only colleague lookup for the availability suite; never sends invitations",
    )
    for command in ("disable", "enable"):
        p = commands.add_parser(command, help="Set/remove a local command-prefix circuit breaker")
        p.add_argument(
            "capability", help="Quoted command prefix, e.g. 'todo task update' or 'teams'"
        )
        if command == "disable":
            p.add_argument("--reason", required=True)
    commands.add_parser("disabled", help="Show configured circuit breakers")


def command_path(args):
    return " ".join(
        str(getattr(args, k))
        for k in (
            "group",
            "command",
            "draft_command",
            "attachment_command",
            "teams_action",
            "task_action",
            "sp_action",
        )
        if getattr(args, k, None)
    )


def inventory():
    from .cli import parser

    capabilities = []

    def visit(p, parts):
        children = [a for a in p._actions if isinstance(a, argparse._SubParsersAction)]
        if children:
            for sub in children:
                for name, child in sub.choices.items():
                    visit(child, [*parts, name])
        else:
            if (
                not parts
                or parts[0] in ("maintainer", "operation")
                or parts == ["reference", "get"]
            ):
                return
            module = (
                "sharepoint"
                if p.get_default("sharepoint")
                else "teams"
                if parts[0] == "teams"
                else "tasks"
                if parts[0] in ("todo", "planner", "tasks")
                else "outlook"
            )
            capabilities.append(
                {
                    "command": " ".join(parts),
                    "module": module,
                    "mutation": bool(p.get_default("mutation")),
                    "implemented": True,
                    "live_availability": "not_checked",
                    "additional_validation_suites": ["composition"]
                    if parts[0] == "mail"
                    and (
                        p.get_default("body_composition")
                        or parts[-1] in ("get", "send")
                        and "attachment" not in parts
                    )
                    else ["calendar-composition", "calendar-drafts"]
                    if parts[0] == "calendar"
                    and (p.get_default("body_composition") or parts[-1] == "get")
                    else ["calendar-drafts"]
                    if parts == ["calendar", "delete"]
                    else ["teams-composition"]
                    if parts[:2] == ["teams", "message"]
                    else [],
                    "validation_suite": "teams-composition"
                    if parts in (["teams", "message", "render"], ["teams", "message", "styles"])
                    else "calendar-drafts"
                    if parts == ["calendar", "send"]
                    else "calendar-composition"
                    if parts[:2] in (["calendar", "render"], ["calendar", "styles"])
                    else "composition"
                    if parts[:2] in (["mail", "render"], ["mail", "styles"])
                    else "availability"
                    if p.get_default("scheduling")
                    else "labels"
                    if p.get_default("label_domain")
                    else "attachments"
                    if parts[:2] == ["mail", "attachment"]
                    else module,
                    "contract_revision": CONTRACT_REVISION,
                }
            )

    visit(parser(), [])
    return {
        "version": __version__,
        "contract_revision": CONTRACT_REVISION,
        "capabilities": capabilities,
        "meaning": "Implementation inventory only. Read probes do not validate mutation semantics; write suites report their exact coverage and remaining gates.",
    }


def breakers():
    value = read_config().get("disabled_capabilities", {})
    if not isinstance(value, dict) or any(
        not isinstance(k, str) or not isinstance(v, str) for k, v in value.items()
    ):
        raise OutlookError(
            "config_invalid", "The capability circuit-breaker configuration is invalid."
        )
    return value


def guard(args):
    if args.group in ("maintainer", "operation") or (
        args.group == "reference" and args.command == "get"
    ):
        return
    path = command_path(args)
    paths = [path]
    if path in ("reference resolve", "files resolve"):
        paths += ["reference resolve", "files resolve"]
    if getattr(args, "label_domain", None) == "category":
        # Aliases share one native adapter and must not bypass its circuit breaker.
        paths += ["category " + args.label_action, "todo category " + args.label_action]
    for prefix, reason in breakers().items():
        if any(p == prefix or p.startswith(prefix + " ") for p in paths):
            raise OutlookError(
                "capability_disabled",
                "This capability was disabled locally by a maintainer.",
                capability=prefix,
                reason=reason,
            )


def compare(report, baseline):
    if (
        not isinstance(baseline, dict)
        or baseline.get("report_type") != "m365-capability-check"
        or not isinstance(baseline.get("checks"), list)
    ):
        raise OutlookError(
            "invalid_baseline", "Use a prior maintainer check report as the baseline."
        )
    previous = {
        c["name"]: c
        for c in baseline["checks"]
        if isinstance(c, dict) and isinstance(c.get("name"), str)
    }
    now = {c["name"]: c for c in report["checks"]}
    regressions = [
        name
        for name, c in now.items()
        if previous.get(name, {}).get("status") == "passed" and c["status"] == "failed"
    ]
    return {
        "regressions": regressions,
        "new_checks": sorted(now.keys() - previous.keys()),
        "not_rechecked": sorted(previous.keys() - now.keys()),
        "contract_revision_changed": baseline.get("contract_revision")
        != report["contract_revision"],
        "availability_comparison_only": True,
        "account_changed": baseline.get("account_fingerprint") != report.get("account_fingerprint"),
    }


def check(module="all", account=None, site=None):
    from .api import API
    from .tasks_api import TasksAPI, q
    from .tasks_session import TasksSession
    from .teams_api import TeamsAPI
    from .teams_session import TeamsSession

    report = {
        "report_type": "m365-capability-check",
        "account_fingerprint": hashlib.sha256(
            (account or read_config().get("account", "")).casefold().encode()
        ).hexdigest(),
        "version": __version__,
        "contract_revision": CONTRACT_REVISION,
        "checked_at": datetime.now(UTC).isoformat(),
        "module": module,
        "writes_tested": False,
        "transports": {},
        "checks": [],
        "limitations": [
            "Browser apps perform their normal background traffic during authentication.",
            "Passing read probes do not prove search semantics, delivery, mutation safety, or every container's permissions.",
            "Run the explicit validation suite for disposable-fixture semantic checks.",
        ],
    }

    def probe(name, callback):
        try:
            value = callback()
            result = {"name": name, "status": "passed"}
            if isinstance(value, dict) and isinstance(value.get("items"), list):
                result["returned_items"] = len(value["items"])
                result["empty_fixture"] = not value["items"]
                result["partial"] = bool(
                    value.get("has_more") or value.get("truncated") or value.get("partial")
                )
            report["checks"].append(result)
            return value
        except OutlookError as exc:
            report["checks"].append({"name": name, "status": "failed", "code": exc.code})
        except Exception:
            report["checks"].append(
                {"name": name, "status": "failed", "code": "unexpected_exception"}
            )
        return None

    def skip(name):
        report["checks"].append(
            {"name": name, "status": "skipped", "reason": "No suitable read fixture was available."}
        )

    def detached(name, session):
        info = session.transport_info()
        if (
            info.get("mode") != "direct_http"
            or info.get("browser_connected") is not False
            or session.page is not None
            or session.http is None
        ):
            raise OutlookError("api_changed", "Service transport is not detached from the browser.")
        report["transports"][name] = info
        return info

    if module in ("all", "outlook"):

        def outlook():
            with Session() as s:
                s.authenticate(expected=account)
                probe("outlook.transport.detached", lambda: detached("outlook", s))
                api = API(s)
                from .labels import Categories

                probe("outlook.category.list", Categories(s).listing)
                folders = probe(
                    "outlook.folder.list", lambda: api.folder_list("msgfolderroot", 1, 0)
                )
                probe("outlook.folder.get", lambda: api.folder_get("inbox"))
                mail = probe("outlook.mail.list", lambda: api.mail_list("inbox", 1, 0))
                if mail and mail.get("items"):
                    item = probe(
                        "outlook.mail.get+attachment.metadata",
                        lambda: api.get(mail["items"][0]["id"]),
                    )
                    probe(
                        "outlook.mail.get.html_body",
                        lambda: api.get(mail["items"][0]["id"], body_type="HTML"),
                    )
                    if item and item.get("Attachments"):
                        aid = item["Attachments"][0].get("AttachmentId", {}).get("Id")
                        if aid:
                            probe("outlook.attachment.get", lambda: api.attachment_get(aid))
                        else:
                            skip("outlook.attachment.get")
                    else:
                        skip("outlook.attachment.get")
                else:
                    skip("outlook.mail.get+attachment.metadata")
                    skip("outlook.mail.get.html_body")
                probe(
                    "outlook.mail.search.contract",
                    lambda: api.mail_search("M365 CLI capability probe unmatched phrase", 1, 0),
                )
                probe("outlook.calendar.calendars", api.calendars)
                start = datetime.now(UTC).replace(hour=0, minute=0, second=0, microsecond=0)
                from .cli import parser
                from .scheduling import snapshot

                def availability():
                    args = parser().parse_args(
                        [
                            "calendar",
                            "availability",
                            "--start",
                            start.isoformat(),
                            "--end",
                            (start + timedelta(days=1)).isoformat(),
                        ]
                    )
                    data = snapshot(api, args)
                    if data["attendees"][0]["status"] != "ok":
                        raise OutlookError(
                            "availability_unavailable", "Self free/busy is unavailable."
                        )
                    return {"items": data["attendees"], "partial": data["partial"]}

                probe("outlook.calendar.availability", availability)

                events = probe(
                    "outlook.calendar.list",
                    lambda: api.calendar_list(
                        start.isoformat(), (start + timedelta(days=7)).isoformat(), 1
                    ),
                )
                if events and events.get("items"):
                    event = probe(
                        "outlook.calendar.get",
                        lambda: api.get(events["items"][0]["id"], calendar=True),
                    )

                    def meeting_state():
                        if not isinstance(event, dict) or any(
                            type(event.get(k)) is not bool
                            for k in (
                                "MeetingRequestWasSent",
                                "IsMeeting",
                                "IsOrganizer",
                                "IsCancelled",
                            )
                        ):
                            raise OutlookError(
                                "api_changed", "Calendar invitation state is missing or changed."
                            )
                        return {"verified": True}

                    probe("outlook.calendar.invitation_state", meeting_state)
                    probe(
                        "outlook.calendar.get.html_body",
                        lambda: api.get(events["items"][0]["id"], calendar=True, body_type="HTML"),
                    )
                else:
                    skip("outlook.calendar.get")
                    skip("outlook.calendar.invitation_state")
                    skip("outlook.calendar.get.html_body")
                return {"authenticated": True, "folder_probe": folders is not None}

        probe("outlook.session", outlook)
    if module in ("all", "teams"):

        def teams():
            with TeamsSession() as s:
                s.authenticate(expected=account)
                api = TeamsAPI(s)
                probe("teams.chat.list", lambda: api.chats(1, 0))
                probe("teams.person.get", lambda: api.person(s.account))
                teams = probe("teams.team.list", lambda: api.teams(1, 0))
                if teams and teams["items"]:
                    channels = probe(
                        "teams.channel.list", lambda: api.channels(teams["items"][0]["id"], 1, 0)
                    )
                    if channels and channels["items"]:
                        probe("teams.channel.get", lambda: api.channel(channels["items"][0]["id"]))
                else:
                    skip("teams.channel.list")
                chat = probe("teams.chat.self", api.self_chat)
                if chat:
                    probe("teams.chat.get+members", lambda: api.chat(chat["id"]))
                    messages = probe(
                        "teams.message.list", lambda: api.messages(chat["id"], limit=1)
                    )
                    if messages and messages["items"]:
                        probe(
                            "teams.message.get",
                            lambda: api.message(chat["id"], messages["items"][0]["id"]),
                        )
                    else:
                        skip("teams.message.get")
                return {"authenticated": True}

        probe("teams.session", teams)
    if module in ("all", "tasks"):

        def tasks():
            with TasksSession(expected=account, required=("graph", "todo", "planner")) as s:
                api = TasksAPI(s)
                lists = probe(
                    "tasks.todo.list",
                    lambda: api.listing("graph", api.todo_path(), "todo", "list", limit=1),
                )
                if lists and lists["items"]:
                    path = api.todo_path(lists["items"][0]["id"])
                    probe("tasks.todo.list.get", lambda: api.get("graph", path, "todo", "list"))
                    todo = probe(
                        "tasks.todo.task.list",
                        lambda: api.listing("graph", path + "/tasks", "todo", "task", limit=1),
                    )
                    if todo and todo["items"]:
                        taskpath = path + "/tasks/" + q(todo["items"][0]["id"])
                        probe(
                            "tasks.todo.task.get",
                            lambda: api.get("graph", taskpath, "todo", "task"),
                        )
                        for child in ("checklistItems", "linkedResources", "attachments"):
                            probe(
                                "tasks.todo." + child,
                                lambda child=child: api.listing(
                                    "graph", taskpath + "/" + child, "todo", child, limit=1
                                ),
                            )
                    else:
                        skip("tasks.todo.task.get+children")
                probe(
                    "tasks.todo.my-day",
                    lambda: api.listing(
                        "todo",
                        "myDayFeed/sections/tasks",
                        "todo",
                        "task",
                        key="Value",
                        id_key="Id",
                        limit=1,
                    ),
                )
                probe(
                    "tasks.todo.groups",
                    lambda: api.listing(
                        "todo", "foldergroups", "todo", "group", key="Value", id_key="Id", limit=1
                    ),
                )
                assigned = probe(
                    "tasks.planner.assigned",
                    lambda: api.listing(
                        "graph", "beta/me/planner/tasks", "planner", "task", limit=1
                    ),
                )
                probe(
                    "tasks.planner.my-day",
                    lambda: api.listing(
                        "graph", "beta/me/planner/myDayTasks", "planner", "task", limit=1
                    ),
                )
                if assigned and assigned["items"]:
                    task = assigned["items"][0]
                    probe(
                        "tasks.planner.task.get",
                        lambda: api.get(
                            "graph", "beta/planner/tasks/" + q(task["id"]), "planner", "task"
                        ),
                    )
                    probe(
                        "tasks.planner.details",
                        lambda: api.get(
                            "graph",
                            "v1.0/planner/tasks/" + q(task["id"]) + "/details",
                            "planner",
                            "details",
                        ),
                    )
                    probe(
                        "tasks.planner.native.task",
                        lambda: api.get(
                            "planner", "tasks('" + q(task["id"]) + "')", "planner", "task"
                        ),
                    )
                    if task.get("planId"):
                        from .labels import PlannerLabels

                        probe(
                            "tasks.planner.label.list",
                            lambda: PlannerLabels(api, task["planId"]).listing(),
                        )
                        probe(
                            "tasks.planner.plan.get",
                            lambda: api.get(
                                "graph",
                                "beta/planner/plans/" + q(task["planId"]),
                                "planner",
                                "plan",
                            ),
                        )
                        probe(
                            "tasks.planner.bucket.list",
                            lambda: api.listing(
                                "graph",
                                "v1.0/planner/plans/" + q(task["planId"]) + "/buckets",
                                "planner",
                                "bucket",
                                limit=1,
                            ),
                        )
                else:
                    skip("tasks.planner.task.get+details+plan+buckets")
                return {"authenticated": True}

        probe("tasks.session", tasks)
    if module in ("all", "sharepoint"):
        if not site:
            skip("sharepoint.explicit_site_required")
        else:

            def sharepoint():
                from .sharepoint_api import SharePointAPI
                from .sharepoint_session import SharePointSession

                with SharePointSession(site, expected=account) as session:
                    probe("sharepoint.transport.detached", lambda: detached("sharepoint", session))

                    def drive():
                        data = session.request("/_api/v2.0/drive")
                        if (
                            not isinstance(data, dict)
                            or not isinstance(data.get("id"), str)
                            or not data["id"]
                        ):
                            raise OutlookError(
                                "api_changed", "SharePoint drive identity is missing or changed."
                            )
                        return data

                    probe("sharepoint.drive", drive)
                    api = SharePointAPI(session)
                    libraries = probe(
                        "sharepoint.libraries", lambda: api.lists(libraries=True, limit=2)
                    )
                    probe("sharepoint.lists", lambda: api.lists(limit=2))
                    probe("sharepoint.users", lambda: api.users(session.account, limit=1))
                    if libraries and libraries["items"]:
                        lib = libraries["items"][0]
                        probe(
                            "sharepoint.folder",
                            lambda: api.file(lib["root_folder_id"], folder=True),
                        )
                        probe("sharepoint.schema", lambda: api.schema(lib["id"]))
                        probe(
                            "sharepoint.files", lambda: api.children(lib["root_folder_id"], limit=2)
                        )
                    else:
                        skip("sharepoint.library_fixture")
                    return {"verified": True}

            probe("sharepoint.session", sharepoint)

    report["passed"] = not any(c["status"] == "failed" for c in report["checks"])
    report["skipped"] = sum(c["status"] == "skipped" for c in report["checks"])
    return report


def run(args):
    if args.account:
        from .validation import email

        email(args.account)
    if args.command == "capabilities":
        return inventory()
    if args.command in ("disable", "enable", "disabled"):
        disabled = breakers()
        if args.command != "disabled":
            prefix = " ".join(args.capability.split())
            if not prefix or not any(
                c["command"] == prefix or c["command"].startswith(prefix + " ")
                for c in inventory()["capabilities"]
            ):
                raise OutlookError(
                    "invalid_argument", "The prefix does not match an implemented capability."
                )
            if args.command == "disable":
                if not args.reason.strip() or len(args.reason) > 500:
                    raise OutlookError("invalid_argument", "Reason must be 1–500 characters.")
                disabled[prefix] = args.reason
            else:
                disabled.pop(prefix, None)
            save_config({**read_config(), "disabled_capabilities": disabled})
        return {"disabled_capabilities": disabled}
    if args.command == "validate":
        account = args.account or read_config().get("account")
        if not account:
            raise OutlookError(
                "account_required", "Validation requires a pinned or explicit account."
            )
        from .validation import email

        email(account)
        argv = [
            sys.executable,
            "-m",
            "m365_cli.live_" + args.suite.replace("-", "_"),
            "--account",
            account,
            "--allow-test-writes",
        ]
        if args.suite == "sharepoint":
            if not args.site:
                raise OutlookError(
                    "invalid_argument",
                    "SharePoint validation requires an explicitly authorized --site.",
                )
            from .sharepoint_validation import site_url

            argv.extend(("--site", site_url(args.site)))
        elif args.site:
            raise OutlookError("invalid_argument", "--site is only for the SharePoint suite.")
        if args.suite in ("outlook", "composition", "calendar-drafts"):
            if not args.recipient:
                raise OutlookError(
                    "invalid_argument",
                    "This validation suite requires an explicitly authorized --recipient.",
                )
            email(args.recipient)
            argv.extend(("--recipient", args.recipient))
        elif args.recipient:
            raise OutlookError(
                "invalid_argument",
                "--recipient is only for Outlook/composition/calendar-drafts suites.",
            )
        if args.channel:
            if args.suite not in ("teams", "teams-composition"):
                raise OutlookError("invalid_argument", "--channel is only for Teams suites.")
            argv.extend(("--channel", args.channel))
        if args.attendee:
            if args.suite != "availability":
                raise OutlookError(
                    "invalid_argument", "--attendee is only for the availability suite."
                )
            for address in args.attendee:
                email(address)
                argv.extend(("--attendee", address))
        completed = subprocess.run(argv, capture_output=True, text=True)
        # Harness stdout is human progress; journals contain the detailed, private evidence.
        if completed.stdout:
            print(completed.stdout, file=sys.stderr, end="")
        if completed.returncode:
            raise OutlookError(
                "validation_failed",
                "Live validation or cleanup failed. Inspect the private journal before another run.",
                suite=args.suite,
                journal_directory=str(state_dir()),
            )
        paths = list(state_dir().glob(args.suite + "-live-*.json"))
        latest = max(paths, key=lambda p: p.stat().st_mtime) if paths else None
        journal = json.loads(latest.read_text()) if latest else {}
        coverage = journal.get("command_coverage", [])
        implemented = {
            c["command"]
            for c in inventory()["capabilities"]
            if (
                c["validation_suite"] == args.suite
                or args.suite in c.get("additional_validation_suites", [])
                or args.suite == "references"
                or (args.suite == "attachments" and c["command"].startswith("mail attachment "))
            )
        }
        return {
            "passed": True,
            "suite": args.suite,
            "journal": str(latest) if latest else None,
            "command_coverage": coverage,
            "commands_not_exercised": sorted(implemented - set(coverage)),
            "cleanup_complete": journal.get("cleanup_complete", False),
            "limitations": "A command being exercised does not validate every flag or tenant/container variant. See docs/MAINTAINER.md.",
        }

    if args.output and Path(args.output).exists():
        raise OutlookError("output_exists", "Report output already exists; choose a new path.")
    baseline = None
    if args.baseline:
        try:
            baseline = json.loads(Path(args.baseline).read_text())
        except (OSError, ValueError):
            raise OutlookError("invalid_baseline", "Cannot read the baseline report.") from None
        compare({"checks": [], "contract_revision": CONTRACT_REVISION}, baseline)
    result = check(args.module, args.account, getattr(args, "site", None))
    if baseline is not None:
        result["comparison"] = compare(result, baseline)
    if args.output:
        atomic_json(args.output, result)
    if not result["passed"]:
        raise OutlookError(
            "capability_check_failed",
            "One or more capability probes failed; see the report.",
            report=result,
        )
    return result
