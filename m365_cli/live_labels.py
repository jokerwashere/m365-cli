"""Opt-in labels canary: private categories, drafts, appointment, To Do list and plan."""

import argparse
import json

from .cli import parser, run
from .errors import OutlookError
from .live_references import Canary
from .maintainer import command_path
from .references import get
from .session import command_lock


class LabelCanary(Canary):
    suite, label = "labels", "labels"

    def invoke(self, api, *words):
        args = parser().parse_args(["--account", self.account, *words])
        entry = {"command": command_path(args), "state": "pending"}
        self.journal["operations"].append(entry)
        self.save()
        if args.mutation:
            args.request_id = self.marker + ":" + str(len(self.journal["operations"]))
        try:
            result = run(args)
        except Exception as exc:
            entry.update(state="failed", code=getattr(exc, "code", "unexpected_error"))
            if isinstance(exc, OutlookError):
                entry["error"] = exc.as_dict()
            if args.mutation:
                self.journal["unresolved"].append(len(self.journal["operations"]) - 1)
            self.save()
            raise
        entry.update(state="completed", id=result.get("id"), stable_id=result.get("stable_id"))
        if entry["command"] not in self.journal["command_coverage"]:
            self.journal["command_coverage"].append(entry["command"])
        self.save()
        if result.get("reference_warning"):
            raise OutlookError("validation_failed", "Reference persistence failed.")
        return result

    def reject(self, code, *words):
        try:
            run(parser().parse_args(["--account", self.account, *words]))
        except OutlookError as exc:
            self.require("refused." + code, exc.code == code)
        else:
            self.require("refused." + code, False)

    def labels(self):
        def call(*words):
            return self.invoke(None, *words)

        cleanup = []
        baseline = call("category", "list")["items"]
        try:
            self.require("outlook.palette", len(call("category", "colors")["items"]) == 26)
            self.require("planner.palette", len(call("planner", "label", "colors")["items"]) == 25)
            cat = call("category", "create", "--name", self.marker, "--color", "sky-blue")
            cat_cleanup = self.track("category", "delete", cat["id"])
            cleanup.append(cat_cleanup)
            self.require(
                "category.reference",
                bool(cat["stable_id"]) and cat["reference"]["url_kind"] == "application",
            )
            self.require("category.get", call("category", "get", cat["id"])["color_id"] == 7)
            self.require(
                "category.filter",
                len(call("category", "list", "--query", self.marker)["items"]) == 1,
            )
            self.reject("category_exists", "category", "create", "--name", self.marker)
            self.reject(
                "conflict",
                "category",
                "update",
                cat["id"],
                "--color",
                "dark-red",
                "--if-version",
                "stale",
            )
            cat = call(
                "category",
                "update",
                cat["id"],
                "--color",
                "dark-red",
                "--if-version",
                cat["version"],
            )
            self.require("category.color", cat["color_id"] == 15)
            # A second owned label proves that add/remove preserves unrelated assignments.
            other = call(
                "todo", "category", "create", "--name", self.marker + " other", "--color", "none"
            )
            cleanup.append(self.track("todo", "category", "delete", other["id"]))
            self.require(
                "todo.category.palette", len(call("todo", "category", "colors")["items"]) == 26
            )
            self.require(
                "todo.category.list",
                len(call("todo", "category", "list", "--query", self.marker)["items"]) == 2,
            )
            self.require(
                "todo.category.get", call("todo", "category", "get", other["id"])["color_id"] == -1
            )
            other = call("todo", "category", "update", other["id"], "--color", "plum")
            self.require("todo.category.update", other["color_id"] == 24)

            draft = call(
                "mail",
                "draft",
                "create",
                "--subject",
                self.marker,
                "--body",
                "Disposable labels test",
            )
            cleanup.append(self.track("mail", "draft", "delete", draft["id"]))
            event = call(
                "calendar",
                "create",
                "--subject",
                self.marker,
                "--start",
                "2031-02-03T14:00:00+01:00",
                "--end",
                "2031-02-03T14:30:00+01:00",
            )
            cleanup.append(self.track("calendar", "delete", event["id"]))
            todo = call("todo", "list", "create", "--name", self.marker)
            cleanup.append(self.track("todo", "list", "delete", todo["id"]))
            task = call("todo", "task", "create", "--list", todo["id"], "--title", self.marker)
            call(
                "todo", "category", "apply", other["id"], "--task", task["id"], "--list", todo["id"]
            )
            for target in (
                ("--mail", draft["id"]),
                ("--event", event["id"]),
                ("--task", task["id"], "--list", todo["id"]),
            ):
                call("category", "apply", other["id"], *target)
                result = call("category", "apply", cat["id"], *target)
                field = "categories" if target[0] == "--task" else "Categories"
                self.require(
                    "assignment.add." + target[0],
                    set(result[field]) == {cat["name"], other["name"]},
                )
                result = call("category", "remove", cat["id"], *target)
                self.require("assignment.remove." + target[0], result[field] == [other["name"]])
                call("category", "apply", cat["id"], *target)
            old = cat
            cat = call("category", "update", cat["id"], "--name", self.marker + " renamed")
            cat_cleanup["command"][-1] = cat["id"]
            self.save()
            self.require(
                "category.rename.identity",
                cat["id"] != old["id"] and cat["stable_id"] == old["stable_id"],
            )
            self.require(
                "category.rename.resolve", get(old["stable_id"], self.account)["id"] == cat["id"]
            )
            self.require(
                "category.rename.mail",
                cat["name"] in call("mail", "draft", "get", draft["id"])["Categories"],
            )
            self.require(
                "category.rename.event",
                cat["name"] in call("calendar", "get", event["id"])["Categories"],
            )
            self.require(
                "category.rename.todo",
                cat["name"]
                in call("todo", "task", "get", task["id"], "--list", todo["id"])["categories"],
            )
            call("category", "delete", cat["id"])
            cat_cleanup["state"] = "cleaned"
            self.save()
            self.require(
                "category.delete.definition_only",
                cat["name"]
                in call("todo", "task", "get", task["id"], "--list", todo["id"])["categories"],
            )
            result = call(
                "todo",
                "category",
                "remove",
                "--name",
                cat["name"],
                "--task",
                task["id"],
                "--list",
                todo["id"],
            )
            self.require("category.orphan.remove", result["categories"] == [other["name"]])
            plan = call("planner", "plan", "create", "--title", self.marker)
            cleanup.append(self.track("planner", "plan", "delete", plan["id"]))
            plan_args = ("--plan", plan["id"])
            t = call("planner", "task", "create", *plan_args, "--title", self.marker)
            self.require(
                "planner.slots", len(call("planner", "label", "list", *plan_args)["items"]) == 25
            )
            label = call(
                "planner", "label", "create", *plan_args, "--name", self.marker, "--color", "pink"
            )
            self.require(
                "planner.label.reference",
                label["id"] == "category1"
                and label["reference"]["url_kind"] == "parent"
                and bool(label["web_url"]),
            )
            call(
                "planner",
                "label",
                "create",
                *plan_args,
                "--name",
                self.marker + " other",
                "--color",
                "green",
            )
            for slot in ("category1", "category4"):
                call("planner", "label", "apply", slot, *plan_args, "--task", t["id"])
            call("planner", "label", "remove", "category1", *plan_args, "--task", t["id"])
            result = call("planner", "label", "apply", "category1", *plan_args, "--task", t["id"])
            self.require("planner.preserve.other", result["appliedCategories"]["category4"] is True)
            label = call("planner", "label", "get", "category1", *plan_args)
            self.reject(
                "conflict",
                "planner",
                "label",
                "rename",
                "category1",
                *plan_args,
                "--name",
                self.marker + " renamed",
                "--if-version",
                "stale",
            )
            label = call(
                "planner",
                "label",
                "rename",
                "category1",
                *plan_args,
                "--name",
                self.marker + " renamed",
                "--if-version",
                label["version"],
            )
            self.require("planner.rename", label["name"] == self.marker + " renamed")
            self.reject(
                "label_slot_in_use",
                "planner",
                "label",
                "recolor",
                "category1",
                *plan_args,
                "--color",
                "green",
            )
            label = call("planner", "label", "recolor", "category1", *plan_args, "--color", "blue")
            self.require(
                "planner.recolor.identity",
                label["id"] == "category5"
                and label["source_label"]["id"] == "category1"
                and label["updated_task_count"] == 1,
            )
            t = call("planner", "task", "get", t["id"])
            self.require(
                "planner.recolor.assignments",
                t["appliedCategories"].get("category1", False) is False
                and t["appliedCategories"]["category5"] is True
                and t["appliedCategories"]["category4"] is True,
            )
            self.reject("label_in_use", "planner", "label", "delete", "category5", *plan_args)
            deleted = call(
                "planner", "label", "delete", "category5", *plan_args, "--remove-from-tasks"
            )
            self.require(
                "planner.delete.definition",
                deleted["definition_cleared"] and not deleted["defined"],
            )
            t = call("planner", "task", "get", t["id"])
            self.require(
                "planner.delete.assignments",
                t["appliedCategories"].get("category5", False) is False
                and t["appliedCategories"]["category4"] is True,
            )
            call("planner", "label", "remove", "category4", *plan_args, "--task", t["id"])
            call("planner", "label", "delete", "category4", *plan_args)
        finally:
            self.clean(None, cleanup)
            after = call("category", "list")["items"]

            def fields(rows):
                return {r["id"]: (r["name"], r["color_id"]) for r in rows}

            self.require(
                "existing.category.definitions.unchanged", fields(after) == fields(baseline)
            )


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--account", required=True)
    p.add_argument("--allow-test-writes", action="store_true", required=True)
    args = p.parse_args()
    with command_lock():
        suite = LabelCanary(args.account)
        try:
            suite.labels()
            suite.journal["state"] = "passed"
        except Exception as exc:
            suite.journal.update(state="failed", code=getattr(exc, "code", "unexpected_error"))
            raise
        finally:
            suite.journal["cleanup_complete"] = all(
                c["state"] == "cleaned" for c in suite.journal["cleanup"]
            )
            suite.journal["limitations"] = (
                "Self-only basic plan, appointment, draft and To Do fixtures. No invitations or messages sent. Mailbox-wide rename completion on large accounts and concurrent edits require separate review. Calendar participant/series policy is covered offline. No custom RGB colors, premium/group-plan canary or category-specific permalink."
            )
            suite.save()
        if not suite.journal["cleanup_complete"]:
            raise OutlookError("cleanup_pending", "Label canary cleanup needs review.")
        print(
            json.dumps(
                {"passed": True, "journal": str(suite.path), "checks": len(suite.journal["checks"])}
            )
        )


if __name__ == "__main__":
    main()
