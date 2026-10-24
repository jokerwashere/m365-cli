"""Opt-in task semantic suite. Only labelled, owned self-only containers are changed."""

import argparse
import json
import uuid
from pathlib import Path
from tempfile import TemporaryDirectory

from .cli import parser, plan
from .errors import OutlookError
from .maintainer import command_path, guard
from .operations import atomic_json, execute
from .session import command_lock, state_dir
from .tasks_api import TasksAPI, q
from .tasks_cli import planner_dispatch, prepare, todo_dispatch
from .tasks_session import TasksSession


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--account", required=True)
    p.add_argument("--allow-test-writes", action="store_true", required=True)
    opt = p.parse_args()
    marker = "M365 CLI validation " + uuid.uuid4().hex[:12]
    path = state_dir() / ("tasks-live-" + marker.rsplit(" ", 1)[1] + ".json")
    journal = {
        "marker": marker,
        "account": opt.account,
        "state": "running",
        "operations": [],
        "cleanup": [],
        "gates_not_tested": [
            "other-person assignment",
            "group-container plans",
            "premium",
            "sharing/access administration",
            "service throttling and credential expiry (offline tests only)",
        ],
    }

    def save():
        atomic_json(path, journal)

    def require(condition, name):
        if not condition:
            raise OutlookError("validation_failed", "Semantic assertion failed.", check=name)

    # Refuse another sweep when a previous run has unresolved container creation/cleanup.
    for previous in state_dir().glob("tasks-live-*.json"):
        old = json.loads(previous.read_text())
        if (
            old.get("state") == "running"
            or old.get("unresolved")
            or any(c.get("state") != "cleaned" for c in old.get("cleanup", []))
        ):
            raise OutlookError(
                "cleanup_pending",
                "A previous task-validation journal needs review before another run.",
            )
    with (
        command_lock(),
        TasksSession(opt.account, ("graph", "todo", "planner")) as session,
        TemporaryDirectory(prefix="m365-validation-") as tmp,
    ):
        api = TasksAPI(session)
        cli_parser = parser()

        def invoke(*argv, data=None, expect=None, replay=None):
            if data is not None:
                argv = (*argv, "--data", json.dumps(data))
            args = cli_parser.parse_args(["--account", opt.account, *argv])
            op = {"command": command_path(args), "state": "pending"}
            journal["operations"].append(op)
            save()
            try:
                prepare(args)
                guard(args)
                args.request_id = replay or marker.replace(" ", "-") + "-" + str(
                    len(journal["operations"])
                )

                def action():
                    return (
                        todo_dispatch(api, args)
                        if args.group == "todo"
                        else planner_dispatch(api, args)
                    )

                result = (
                    execute(
                        args.request_id,
                        session.account,
                        {"service": args.group, "tenant": session.tenant, "command": plan(args)},
                        action,
                    )
                    if args.mutation
                    else action()
                )
                if expect:
                    raise OutlookError(
                        "validation_failed", "Expected refusal was not returned.", expected=expect
                    )
                op.update(state="passed", result=result, request_id=args.request_id)
                save()
                print(op["command"] + " passed", flush=True)
                return result
            except OutlookError as exc:
                op.update(
                    state="expected_error" if exc.code == expect else "failed",
                    code=exc.code,
                    error=exc.as_dict(),
                )
                if (
                    exc.code == "write_outcome_unknown"
                    and args.command in ("list", "plan", "group")
                    and args.task_action == "create"
                ):
                    journal["unresolved"] = True
                save()
                if exc.code != expect:
                    raise
                return None

        def container(source, kind):
            flag = "--title" if kind == "plan" else "--name"
            obj = invoke(source, kind, "create", flag, marker)
            record = {"source": source, "kind": kind, "id": obj["id"], "state": "created"}
            journal["cleanup"].append(record)
            save()
            return obj["id"]

        failure = None
        try:
            a, b = container("todo", "list"), container("todo", "list")
            group = container("todo", "group")
            invoke("todo", "group", "update", group, "--name", marker + " group")
            invoke("todo", "list", "group", b, "--group-id", group)
            invoke("todo", "list", "update", a, "--name", marker + " source")
            invoke("todo", "list", "list", "--limit", "1")
            task = invoke(
                "todo",
                "task",
                "create",
                "--list",
                a,
                "--title",
                marker + " ALPHA",
                "--body",
                "Unicode żółć ✓",
            )
            tid, oldversion = task["id"], task["version"]
            invoke(
                "todo",
                "task",
                "update",
                tid,
                "--list",
                a,
                "--status",
                "inProgress",
                "--if-version",
                oldversion,
                data={
                    "importance": "high",
                    "dueDateTime": {"dateTime": "2031-02-03T17:00:00", "timeZone": "Europe/Zurich"},
                    "isReminderOn": True,
                    "reminderDateTime": {
                        "dateTime": "2031-02-03T16:00:00",
                        "timeZone": "Europe/Zurich",
                    },
                    "categories": ["m365 validation"],
                },
            )
            invoke(
                "todo",
                "task",
                "update",
                tid,
                "--list",
                a,
                "--title",
                "must not write",
                "--if-version",
                oldversion,
                expect="conflict",
            )
            invoke(
                "todo",
                "task",
                "delete",
                tid,
                "--list",
                a,
                "--if-version",
                oldversion,
                expect="unsupported_precondition",
            )
            found = invoke("todo", "task", "list", "--list", a, "--query", "ALPHA")
            missing = invoke("todo", "task", "list", "--list", a, "--query", "MUST-NOT-MATCH")
            require(
                len(found["items"]) == 1 and not missing["items"],
                "To Do positive/negative title query",
            )
            for status in ("waitingOnOthers", "deferred", "completed", "notStarted"):
                obj = invoke("todo", "task", "update", tid, "--list", a, "--status", status)
                require(obj["status"] == status, "To Do status")
            invoke("todo", "task", "complete", tid, "--list", a)
            invoke("todo", "task", "reopen", tid, "--list", a)
            step = invoke(
                "todo", "step", "create", "--list", a, "--task", tid, "--name", marker + " step"
            )
            invoke(
                "todo",
                "step",
                "update",
                step["id"],
                "--list",
                a,
                "--task",
                tid,
                "--name",
                marker + " renamed step",
            )
            invoke("todo", "step", "check", step["id"], "--list", a, "--task", tid)
            invoke("todo", "step", "uncheck", step["id"], "--list", a, "--task", tid)
            invoke("todo", "step", "list", "--list", a, "--task", tid)
            invoke("todo", "step", "delete", step["id"], "--list", a, "--task", tid)
            link = invoke(
                "todo",
                "link",
                "create",
                "--list",
                a,
                "--task",
                tid,
                data={
                    "webUrl": "https://example.com/m365-validation",
                    "displayName": marker,
                    "applicationName": "m365 CLI",
                },
            )
            invoke(
                "todo",
                "link",
                "update",
                link["id"],
                "--list",
                a,
                "--task",
                tid,
                data={"displayName": marker + " link"},
            )
            invoke("todo", "link", "get", link["id"], "--list", a, "--task", tid)
            invoke("todo", "link", "list", "--list", a, "--task", tid)
            invoke("todo", "link", "delete", link["id"], "--list", a, "--task", tid)
            fixture = Path(tmp) / "fixture.txt"
            fixture.write_bytes(b"M365 CLI synthetic attachment\n")
            attachment = invoke(
                "todo", "attachment", "add", "--list", a, "--task", tid, "--file", str(fixture)
            )
            downloaded = Path(tmp) / "download.txt"
            invoke(
                "todo",
                "attachment",
                "get",
                attachment["id"],
                "--list",
                a,
                "--task",
                tid,
                "--output",
                str(downloaded),
            )
            require(downloaded.read_bytes() == fixture.read_bytes(), "attachment bytes")
            invoke("todo", "attachment", "list", "--list", a, "--task", tid)
            invoke("todo", "attachment", "delete", attachment["id"], "--list", a, "--task", tid)
            assignment = invoke("todo", "task", "assign", tid, "--list", a)
            invoke("todo", "task", "unassign", tid, "--list", a, "--assignment", assignment["id"])
            invoke("todo", "my-day", "add", "--list", a, "--task", tid, "--date", "2031-02-03")
            invoke("todo", "my-day", "list", "--limit", "1")
            invoke("todo", "my-day", "remove", "--list", a, "--task", tid)
            delta = invoke("todo", "delta", "list", "--list", a)
            for _ in range(30):
                if not delta["next_cursor"]:
                    break
                delta = invoke(
                    "todo", "delta", "list", "--list", a, "--cursor", delta["next_cursor"]
                )
            require(bool(delta["delta_cursor"]), "To Do delta token")
            moved = invoke("todo", "task", "move", tid, "--list", a, "--to-list", b)
            invoke("todo", "task", "delete", moved["id"], "--list", b)
            invoke("todo", "delta", "list", "--list", a, "--cursor", delta["delta_cursor"])
            recurring = invoke(
                "todo",
                "task",
                "create",
                "--list",
                a,
                "--title",
                marker + " recurring",
                data={
                    "dueDateTime": {"dateTime": "2031-02-03T17:00:00", "timeZone": "Europe/Zurich"},
                    "recurrence": {
                        "pattern": {"type": "daily", "interval": 1},
                        "range": {"type": "noEnd", "startDate": "2031-02-03"},
                    },
                },
            )
            invoke("todo", "task", "complete", recurring["id"], "--list", a)
            instances = invoke("todo", "task", "list", "--list", a, "--query", "recurring")
            require(
                any(i["status"] == "notStarted" for i in instances["items"])
                and any(i["status"] == "completed" for i in instances["items"]),
                "To Do recurring completion has active and completed occurrences",
            )
            pa, pb = container("planner", "plan"), container("planner", "plan")
            invoke("planner", "plan", "update", pa, "--title", marker + " plan")
            invoke("planner", "plan", "list", "--limit", "1")
            bucket = invoke(
                "planner", "bucket", "create", "--plan", pa, "--name", marker + " bucket"
            )
            invoke(
                "planner", "bucket", "update", bucket["id"], "--name", marker + " renamed bucket"
            )
            invoke("planner", "bucket", "list", "--plan", pa)
            task = invoke(
                "planner",
                "task",
                "create",
                "--plan",
                pa,
                "--title",
                marker + " BETA",
                data={"bucketId": bucket["id"], "priority": 3},
            )
            pid, pv = task["id"], task["version"]
            invoke(
                "planner",
                "task",
                "update",
                pid,
                "--title",
                marker + " GAMMA",
                "--progress",
                "50",
                "--if-version",
                pv,
                data={
                    "startDateTime": "2031-02-03T13:00:00Z",
                    "dueDateTime": "2031-02-03T14:00:00Z",
                    "appliedCategories": {"category1": True},
                },
            )
            invoke(
                "planner",
                "task",
                "update",
                pid,
                "--title",
                "must not write",
                "--if-version",
                pv,
                expect="conflict",
            )
            invoke("planner", "task", "delete", pid, "--if-version", pv, expect="conflict")
            invoke("planner", "task", "assign", pid, "--user", "self")
            invoke("planner", "task", "unassign", pid, "--user", "self")
            invoke("planner", "task", "assign", pid, "--user", "self")
            invoke("planner", "task", "complete", pid)
            invoke("planner", "task", "reopen", pid)
            checklist = str(uuid.uuid4())
            ref = "https%3A//example%2Ecom/m365-validation"
            detail = invoke(
                "planner",
                "details",
                "update",
                "--task",
                pid,
                "--body",
                marker + " notes",
                data={
                    "checklist": {
                        checklist: {
                            "@odata.type": "#microsoft.graph.plannerChecklistItem",
                            "title": marker + " step",
                            "isChecked": False,
                            "orderHint": " !",
                        }
                    },
                    "references": {
                        ref: {
                            "@odata.type": "#microsoft.graph.plannerExternalReference",
                            "alias": marker,
                            "type": "Other",
                            "previewPriority": " !",
                        }
                    },
                },
            )
            require(
                checklist in detail["checklist"] and ref in detail["references"], "Planner details"
            )
            invoke(
                "planner",
                "details",
                "update",
                "--task",
                pid,
                data={
                    "checklist": {
                        checklist: {
                            "@odata.type": "#microsoft.graph.plannerChecklistItem",
                            "isChecked": True,
                        }
                    }
                },
            )
            invoke(
                "planner",
                "details",
                "update",
                "--task",
                pid,
                data={"checklist": {checklist: None}, "references": {ref: None}},
            )
            invoke(
                "planner",
                "labels",
                "update",
                "--plan",
                pa,
                data={"categoryDescriptions": {"category1": marker}},
            )
            positive = invoke(
                "planner", "task", "list", "--plan", pa, "--query", "GAMMA", "--limit", "1"
            )
            negative = invoke("planner", "task", "list", "--plan", pa, "--query", "MUST-NOT-MATCH")
            require(
                len(positive["items"]) == 1 and not negative["items"],
                "Planner bounded positive/negative query",
            )
            invoke("planner", "my-day", "add", "--task", pid)
            invoke("planner", "my-day", "list", "--limit", "1")
            invoke("planner", "my-day", "remove", "--task", pid)
            comment = invoke(
                "planner", "comment", "create", "--task", pid, "--body", marker + " test"
            )
            invoke("planner", "comment", "list", "--task", pid)
            invoke(
                "planner",
                "comment",
                "update",
                comment["id"],
                "--task",
                pid,
                "--body",
                marker + " edited",
            )
            invoke("planner", "comment", "delete", comment["id"], "--task", pid)
            invoke("planner", "task", "move", pid, "--to-plan", pb)
            deleted = invoke("planner", "task", "delete", pid, "--allow-undo")
            require(bool(deleted.get("operation_id")), "native deletion operation ID")
            restored = invoke(
                "planner",
                "deletion",
                "undo",
                "--task",
                pid,
                "--operation-id",
                deleted["operation_id"],
            )
            require(restored["id"] == pid, "restored task")
            invoke("planner", "task", "delete", pid)
            recurring = invoke(
                "planner",
                "task",
                "create",
                "--plan",
                pb,
                "--title",
                marker + " recurring",
                data={
                    "recurrence": {
                        "schedule": {
                            "pattern": {"type": "daily", "interval": 1},
                            "patternStartDateTime": "2031-02-03T13:00:00Z",
                        }
                    }
                },
            )
            invoke(
                "planner", "task", "delete", recurring["id"], expect="recurrence_intent_required"
            )
            done = invoke("planner", "task", "complete", recurring["id"])
            successor = done.get("recurrence", {}).get("nextInSeriesTaskId")
            require(bool(successor), "Planner recurrence successor")
            invoke("planner", "task", "delete", successor, "--recurrence", "stop-series")
            invoke("planner", "bucket", "delete", bucket["id"])
            invoke("planner", "history", "list", "--plan", pa, "--limit", "1")
            invoke("planner", "plan", "archive", pa)
            invoke("planner", "plan", "unarchive", pa)
            invoke("planner", "delta", "list")
            journal["state"] = "passed"
        except Exception as exc:
            failure = exc
            journal["state"] = "failed"
            journal["failure_code"] = (
                exc.code if isinstance(exc, OutlookError) else type(exc).__name__
            )
            save()
        finally:
            # Lists precede groups during cleanup; deleting a known owned container also
            # removes any task whose response was lost, without guessing task IDs.
            resources = sorted(journal["cleanup"], key=lambda r: r["kind"] == "group")
            for resource in resources:
                try:
                    source, kind, rid = resource["source"], resource["kind"], resource["id"]
                    if kind == "plan":
                        obj = api.entity("graph", "beta/planner/plans/" + q(rid))
                        require(
                            obj.get("title", "").startswith(marker)
                            and obj.get("container", {}).get("type") == "user"
                            and obj["container"].get("containerId") == session.oid,
                            "owned plan cleanup",
                        )
                    elif kind == "list":
                        obj = api.todo_list_guard(rid)
                        require(obj.get("displayName", "").startswith(marker), "owned list cleanup")
                    else:
                        obj = api.entity("todo", "foldergroups/" + q(rid), key="Id")
                        require(obj.get("Name", "").startswith(marker), "group cleanup")
                    invoke(source, kind, "delete", rid)
                    resource["state"] = "cleaned"
                except Exception as exc:
                    resource["state"] = "cleanup_pending"
                    resource["error_code"] = (
                        exc.code if isinstance(exc, OutlookError) else type(exc).__name__
                    )
                save()
            journal["cleanup_complete"] = all(
                r["state"] == "cleaned" for r in resources
            ) and not journal.get("unresolved", False)
            journal["command_coverage"] = sorted(
                {o["command"] for o in journal["operations"] if o["state"] == "passed"}
            )
            save()
        if failure or not journal["cleanup_complete"]:
            raise OutlookError(
                "validation_failed",
                "Task validation or cleanup needs review in its private journal.",
                journal=str(path),
            ) from None
    print(
        "Task validation passed; all owned test containers were independently verified absent.",
        flush=True,
    )


if __name__ == "__main__":
    main()
