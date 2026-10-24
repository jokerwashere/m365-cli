"""Opt-in reference/link canary. Creates owned fixtures; sends only to Teams self-chat."""

import argparse
import json
import uuid
from urllib.parse import unquote

from .api import API
from .cli import dispatch, parser, plan, prepare
from .errors import OutlookError
from .maintainer import command_path, guard
from .operations import atomic_json, execute
from .references import References, get
from .session import Session, command_lock, state_dir
from .tasks_api import TasksAPI
from .tasks_cli import planner_dispatch, todo_dispatch
from .tasks_session import TasksSession
from .teams_api import TeamsAPI
from .teams_cli import dispatch as teams_dispatch
from .teams_session import TeamsSession


class Canary:
    suite = "references"
    label = "reference"

    def __init__(self, account):
        self.account = account
        self.marker = "M365 CLI " + self.label + " validation " + uuid.uuid4().hex[:12]
        self.path = state_dir() / (self.suite + "-live-" + self.marker.rsplit(" ", 1)[1] + ".json")
        for path in state_dir().glob(self.suite + "-live-*.json"):
            prior = json.loads(path.read_text())
            if not prior.get("cleanup_complete") or prior.get("unresolved"):
                raise OutlookError(
                    "cleanup_pending",
                    "Review the previous " + self.label + " validation journal before another run.",
                )
        self.journal = {
            "marker": self.marker,
            "state": "running",
            "operations": [],
            "checks": [],
            "cleanup": [],
            "command_coverage": [],
            "unresolved": [],
        }
        self.save()

    def save(self):
        atomic_json(self.path, self.journal)

    def require(self, name, condition):
        self.journal["checks"].append({"name": name, "passed": bool(condition)})
        self.save()
        print(name + (": passed" if condition else ": FAILED"), flush=True)
        if not condition:
            raise OutlookError(
                "validation_failed", self.label.capitalize() + " validation failed.", check=name
            )

    def invoke(self, api, *words):
        args = parser().parse_args(["--account", self.account, *words])
        prepare(args)
        guard(args)
        session = api.session if isinstance(api, API) else api.s
        refs = References(session.account, getattr(session, "tenant", None))
        command = command_path(args)
        entry = {"command": command, "state": "pending"}
        self.journal["operations"].append(entry)
        self.save()

        def action():
            if args.group in ("todo", "planner"):
                parent = None
                if args.group == "planner" and args.command in ("details", "comment"):
                    from .tasks_api import q

                    parent = api.entity("graph", "beta/planner/tasks/" + q(args.task_id))
                data = (todo_dispatch if args.group == "todo" else planner_dispatch)(api, args)
                return refs.tasks(data, args, parent)
            if args.group == "teams":
                return refs.teams(teams_dispatch(api, args), args, api)
            return refs.outlook(dispatch(api, args, refs=refs), args)

        try:
            data = (
                execute(
                    self.marker + ":" + str(len(self.journal["operations"])),
                    self.account,
                    {"command": plan(args)},
                    action,
                )
                if args.mutation
                else action()
            )
        except Exception as exc:
            entry.update(state="failed", code=getattr(exc, "code", "unexpected_error"))
            if isinstance(exc, OutlookError):
                entry["error"] = exc.as_dict()
            if args.mutation:
                self.journal["unresolved"].append(len(self.journal["operations"]) - 1)
            self.save()
            raise
        entry.update(state="completed", id=data.get("id"), stable_id=data.get("stable_id"))
        if command not in self.journal["command_coverage"]:
            self.journal["command_coverage"].append(command)
        self.save()
        if data.get("reference_warning"):
            self.journal["unresolved"].append(len(self.journal["operations"]) - 1)
            self.save()
            raise OutlookError(
                "validation_failed", "Reference persistence failed; do not repeat the operation."
            )
        return data

    def track(self, *words):
        entry = {"command": list(words), "state": "pending"}
        self.journal["cleanup"].append(entry)
        self.save()
        return entry

    def clean(self, api, entries):
        for entry in reversed(entries):
            if entry["state"] == "cleaned":
                continue
            try:
                self.invoke(api, *entry["command"])
                entry["state"] = "cleaned"
            except Exception as exc:
                entry["code"] = getattr(exc, "code", "unexpected_error")
            self.save()

    def opens(self, context, item, label, name):
        self.require(
            name + ".has_url", item.get("web_url") and item["reference"]["url_kind"] == "item"
        )
        page = context.new_page()
        try:
            page.goto(item["web_url"], wait_until="domcontentloaded", timeout=45000)
            page.wait_for_function(
                "s => document.body.innerText.includes(s) || Array.from(document.querySelectorAll('input,textarea')).some(e => e.value === s)",
                arg=label,
                timeout=45000,
            )
            # A generic app landing page displaying the fixture in its sidebar is insufficient.
            self.require(name + ".route_and_fixture", item["id"] in unquote(page.url))
        except OutlookError:
            raise
        except Exception:
            self.require(name + ".fixture_visible", False)
        finally:
            page.close()

    def tasks(self):
        with TasksSession(self.account, {"graph", "todo", "planner"}) as s:
            api, cleanup = TasksAPI(s), []
            try:
                a = self.invoke(api, "todo", "list", "create", "--name", self.marker + " A")
                cleanup.append(self.track("todo", "list", "delete", a["id"]))
                b = self.invoke(api, "todo", "list", "create", "--name", self.marker + " B")
                cleanup.append(self.track("todo", "list", "delete", b["id"]))
                task = self.invoke(
                    api,
                    "todo",
                    "task",
                    "create",
                    "--list",
                    a["id"],
                    "--title",
                    self.marker + " todo",
                )
                step = self.invoke(
                    api,
                    "todo",
                    "step",
                    "create",
                    "--list",
                    a["id"],
                    "--task",
                    task["id"],
                    "--name",
                    "Reference step",
                )
                self.require(
                    "todo.child_parent_url",
                    step["web_url"] == task["web_url"]
                    and step["reference"]["url_kind"] == "parent",
                )
                self.opens(s.ctx, a, self.marker + " A", "todo.list")
                self.opens(s.ctx, task, self.marker + " todo", "todo.task")
                moved = self.invoke(
                    api, "todo", "task", "move", task["id"], "--list", a["id"], "--to-list", b["id"]
                )
                self.require("todo.move_stable", moved["stable_id"] == task["stable_id"])
                self.require(
                    "todo.move_resolve", get(task["stable_id"], self.account)["id"] == moved["id"]
                )
                self.opens(s.ctx, moved, self.marker + " todo", "todo.moved_task")
                p = self.invoke(
                    api,
                    "planner",
                    "plan",
                    "create",
                    "--title",
                    self.marker + " plan",
                    "--container",
                    "self",
                )
                cleanup.append(self.track("planner", "plan", "delete", p["id"]))
                task = self.invoke(
                    api,
                    "planner",
                    "task",
                    "create",
                    "--plan",
                    p["id"],
                    "--title",
                    self.marker + " planner",
                )
                assigned = self.invoke(
                    api, "planner", "task", "assign", task["id"], "--user", "self"
                )
                self.require(
                    "planner.assignment_stable", assigned["stable_id"] == task["stable_id"]
                )
                self.opens(s.ctx, p, self.marker + " plan", "planner.plan")
                self.opens(s.ctx, assigned, self.marker + " planner", "planner.task")
                comment = self.invoke(
                    api,
                    "planner",
                    "comment",
                    "create",
                    "--task",
                    task["id"],
                    "--body",
                    "Reference validation",
                )
                self.require(
                    "planner.comment_parent_url",
                    comment["web_url"] == task["web_url"]
                    and comment["reference"]["url_kind"] == "parent",
                )
            finally:
                self.clean(api, cleanup)

    def outlook(self):
        with Session() as s:
            s.authenticate(expected=self.account)
            api, cleanup = API(s), []
            try:
                calendars = self.invoke(api, "calendar", "calendars")
                self.require(
                    "outlook.calendar_container_references",
                    all(
                        c.get("stable_id") and c.get("web_url")
                        for group in calendars.get("CalendarGroups", [])
                        for c in group.get("Calendars", [])
                    ),
                )
                folder = self.invoke(api, "folder", "create", "--name", self.marker + " folder")
                cleanup.append(self.track("folder", "delete", folder["id"]))
                draft = self.invoke(
                    api,
                    "mail",
                    "draft",
                    "create",
                    "--subject",
                    self.marker + " mail",
                    "--body",
                    "Reference validation draft",
                )
                item = self.track("mail", "draft", "delete", draft["id"])
                cleanup.append(item)
                # Link navigation is an explicit UI check, separate from the
                # detached API session. It cannot supply a transport fallback.
                with Session() as links:
                    self.opens(links.context, draft, self.marker + " mail", "outlook.mail")
                    self.opens(links.context, folder, self.marker + " folder", "outlook.folder")
                moved = self.invoke(api, "mail", "move", draft["id"], "--folder", folder["id"])
                item["command"] = ["mail", "delete", moved["id"]]
                self.save()
                self.require("outlook.move_stable", moved["stable_id"] == draft["stable_id"])
                self.require(
                    "outlook.move_resolve",
                    get(draft["stable_id"], self.account)["id"] == moved["id"],
                )
                with Session() as links:
                    self.opens(links.context, moved, self.marker + " mail", "outlook.moved_mail")
                event = self.invoke(
                    api,
                    "calendar",
                    "create",
                    "--subject",
                    self.marker + " event",
                    "--start",
                    "2031-03-04T14:00:00+01:00",
                    "--end",
                    "2031-03-04T14:30:00+01:00",
                )
                cleanup.append(self.track("calendar", "delete", event["id"]))
                with Session() as links:
                    self.opens(links.context, event, self.marker + " event", "outlook.event")
            finally:
                self.clean(api, cleanup)

    def teams(self):
        with TeamsSession() as s:
            s.authenticate(expected=self.account)
            api, cleanup = TeamsAPI(s), []
            try:
                chat = self.invoke(api, "teams", "chat", "self")
                self.require("teams.chat_url", chat["id"] in unquote(chat["web_url"]))
                sent = self.invoke(
                    api, "teams", "message", "send", "--chat", chat["id"], "--body", self.marker
                )["message"]
                entry = self.track("teams", "message", "delete", sent["id"], "--chat", chat["id"])
                cleanup.append(entry)
                fetched = self.invoke(
                    api, "teams", "message", "get", sent["id"], "--chat", chat["id"]
                )
                self.require(
                    "teams.message_reference",
                    sent["stable_id"] == fetched["stable_id"]
                    and chat["id"] in unquote(sent["web_url"])
                    and sent["id"] in unquote(sent["web_url"]),
                )
                teams = self.invoke(api, "teams", "team", "list", "--limit", "1")["items"]
                if teams:
                    self.require("teams.team_url", bool(teams[0].get("web_url")))
                    channels = self.invoke(
                        api, "teams", "channel", "list", "--team", teams[0]["id"], "--limit", "1"
                    )["items"]
                    if channels:
                        self.require("teams.channel_url", bool(channels[0].get("web_url")))
                deleted = self.invoke(api, *entry["command"])
                entry["state"] = "cleaned"
                self.save()
                self.require(
                    "teams.delete_stable",
                    deleted["message"]["stable_id"] == sent["stable_id"]
                    and deleted["message"]["deleted"],
                )
            finally:
                self.clean(api, cleanup)


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--account", required=True)
    p.add_argument("--allow-test-writes", action="store_true", required=True)
    args = p.parse_args()
    with command_lock():
        suite = Canary(args.account)
        try:
            suite.tasks()
            suite.outlook()
            suite.teams()
            suite.journal["state"] = "passed"
        except Exception as exc:
            suite.journal.update(state="failed", code=getattr(exc, "code", "unexpected_error"))
            raise
        finally:
            suite.journal["cleanup_complete"] = all(
                c["state"] == "cleaned" for c in suite.journal["cleanup"]
            )
            suite.journal["limitations"] = (
                "Browser canaries check fixture visibility and route identity for Outlook and tasks. Teams checks documented link structure and API readback. This does not test recipient permissions or every child/container variant."
            )
            suite.save()
        if not suite.journal["cleanup_complete"]:
            raise OutlookError("cleanup_pending", "Reference canary cleanup needs review.")
        print(
            json.dumps(
                {"passed": True, "journal": str(suite.path), "checks": len(suite.journal["checks"])}
            )
        )


if __name__ == "__main__":
    main()
