"""Outlook category definitions and Planner's fixed colored label slots."""

import hashlib
import json
import re
import uuid

from .api import typed
from .contracts import array_value, changed, object_value
from .errors import OutlookError
from .tasks_api import q, review, version

# Native Outlook Web enum names. Graph/desktop renders some presets differently.
OUTLOOK_COLORS = "cranberry bronze peach marigold light-green light-teal lime sky-blue lavender pink silver steel beige mink charcoal dark-red dark-orange dark-brown gold dark-green dark-teal forest dark-blue dark-purple plum".split()
PLANNER_COLORS = "pink red yellow green blue purple bronze lime aqua gray silver brown cranberry orange peach marigold light-green dark-green teal light-blue dark-blue lavender plum light-gray dark-gray".split()


def color(value, source):
    value = str(value).strip().lower().replace("_", "-").replace(" ", "-").replace("grey", "gray")
    if source == "category":
        if value in ("none", "-1"):
            return -1
        if re.fullmatch(r"preset(?:[0-9]|1[0-9]|2[0-4])", value):
            return int(value[6:])
        # Conventional aliases that do not conflict with a native web color name.
        value = {
            "red": "cranberry",
            "orange": "bronze",
            "yellow": "marigold",
            "green": "light-green",
            "blue": "sky-blue",
            "purple": "lavender",
        }.get(value, value)
        if value in OUTLOOK_COLORS:
            return OUTLOOK_COLORS.index(value)
    elif value in PLANNER_COLORS:
        return "category" + str(PLANNER_COLORS.index(value) + 1)
    raise OutlookError(
        "invalid_color",
        "Choose a supported color from the colors command. Custom RGB/hex colors are not supported.",
    )


def palette(source):
    if source == "category":
        return {
            "items": [
                {"id": "none", "color_id": -1, "name": "none"},
                *[
                    {"id": f"preset{i}", "color_id": i, "name": name}
                    for i, name in enumerate(OUTLOOK_COLORS)
                ],
            ],
            "scope": "outlook_category_palette",
            "note": "Names follow Outlook Web. Other clients may render presets differently.",
        }
    return {
        "items": [
            {"id": f"category{i + 1}", "color": name} for i, name in enumerate(PLANNER_COLORS)
        ],
        "scope": "planner_fixed_slots",
    }


def label_name(value, *, category=False):
    if (
        not isinstance(value, str)
        or not value.strip()
        or value != value.strip()
        or len(value.encode("utf-16-le")) // 2 > 255
        or any(ord(c) < 32 for c in value)
        or (category and any(c in value for c in ",;"))
    ):
        raise OutlookError(
            "invalid_argument",
            "Label names must be trimmed, nonempty, at most 255 UTF-16 characters, and contain no control characters"
            + (", commas or semicolons." if category else "."),
        )
    return value


def slot(value):
    if not isinstance(value, str) or not re.fullmatch(r"category(?:[1-9]|1[0-9]|2[0-5])", value):
        raise OutlookError("invalid_argument", "Planner label IDs are category1–category25.")
    return value


def applied_categories(item, *, write=False):
    values = object_value(item.get("appliedCategories"), "planner.labels.task", write=write)
    for key, value in values.items():
        if key.startswith("@"):
            continue
        if not re.fullmatch(r"category(?:[1-9]|1[0-9]|2[0-5])", key) or type(value) is not bool:
            changed("planner.labels.task", "appliedCategories", write=write)
    return values


def stamp(item):
    return hashlib.sha256(
        json.dumps(
            {k: item.get(k) for k in ("Id", "Name", "Color", "KeyboardShortcut")}, sort_keys=True
        ).encode()
    ).hexdigest()


class Categories:
    def __init__(self, session):
        self.session = session

    def call(self, action, fields):
        return self.session.request(
            action, {"request": typed(action + "Request", **fields)}, native=True
        )

    def raw(self):
        data = object_value(self.call("GetMasterCategoryList", {}), "outlook.categories")
        if data.get("WasSuccessful") is not True:
            changed("outlook.categories", "WasSuccessful")
        rows = array_value(data.get("MasterList"), "outlook.categories", fields=("Id", "Name"))
        ids, names = set(), set()
        for row in rows:
            if (
                type(row.get("Color")) is not int
                or row["Id"] in ids
                or row["Name"].casefold() in names
            ):
                changed("outlook.categories", "unique_id_name_color")
            ids.add(row["Id"])
            names.add(row["Name"].casefold())
        return rows

    def item(self, row):
        number = row["Color"]
        return {
            "id": row["Id"],
            "name": row["Name"],
            "color_id": number,
            "color": OUTLOOK_COLORS[number]
            if 0 <= number < 25
            else "none"
            if number == -1
            else None,
            "preset": f"preset{number}" if 0 <= number < 25 else "none" if number == -1 else None,
            "version": stamp(row),
            "precondition_atomic": False,
            "resource": "category",
            "source": "outlook",
        }

    def listing(self):
        return {"items": [self.item(row) for row in self.raw()], "complete": True}

    def find(self, rows, identifier):
        matches = [row for row in rows if row["Id"] == identifier]
        if len(matches) != 1:
            raise OutlookError(
                "category_not_found", "Category ID is not in this mailbox's master list."
            )
        return matches[0]

    def get(self, identifier):
        return self.item(self.find(self.raw(), identifier))

    def mutate(self, action, *, identifier=None, name=None, color_id=None, expected=None):
        if action not in ("create", "update", "delete") or (action != "create" and not identifier):
            raise OutlookError("invalid_argument", "Specify a supported action and category ID.")
        rows = self.raw()
        current = self.find(rows, identifier) if identifier else None
        if current and expected is not None and stamp(current) != expected:
            raise OutlookError("conflict", "Category changed since it was reviewed.")
        if name is not None:
            label_name(name, category=True)
            if any(r["Name"].casefold() == name.casefold() and r is not current for r in rows):
                raise OutlookError("category_exists", "A category already uses this name.")
        if color_id is not None and (type(color_id) is not int or not -1 <= color_id <= 24):
            raise OutlookError("invalid_color", "Unsupported category color.")
        add, remove, recolor = [], [], []
        if action == "create":
            if name is None:
                raise OutlookError("invalid_argument", "Category creation requires a name.")
            desired = {
                "Id": str(uuid.uuid4()),
                "Name": name,
                "Color": -1 if color_id is None else color_id,
                "KeyboardShortcut": 0,
            }
            add = [desired]
        elif action == "delete":
            desired = current
            remove = [current["Name"]]
        else:
            desired = {
                **current,
                **({"Name": name} if name is not None else {}),
                **({"Color": color_id} if color_id is not None else {}),
            }
            if desired == current:
                return {**self.item(current), "unchanged": True}
            if desired["Name"] != current["Name"]:
                add, remove = [desired], [current["Name"]]
            else:
                recolor = [desired]
        data = self.call(
            "UpdateMasterCategoryList",
            {
                "AddCategoryList": add,
                "RemoveCategoryList": remove,
                "ChangeCategoryColorList": recolor,
                "UpdateCategoryLastTimeUsedList": [],
                "ChangeCategoryKeyboardShortcutList": [],
            },
        )
        if data.get("WasSuccessful") is not True or data.get("ResponseCode") != "NoError":
            changed("outlook.categories.write", "acknowledgement", write=True)
        try:
            after = self.raw()
            if action == "delete":
                if any(r["Id"] == current["Id"] for r in after):
                    changed("outlook.categories.write", "delete_readback")
                return {**self.item(current), "deleted": True, "item_assignments": "unchanged"}
            renamed = current and current["Name"] != desired["Name"]
            if action == "create" or renamed:
                candidates = [
                    r
                    for r in after
                    if r["Name"] == desired["Name"] and r["Id"] not in {x["Id"] for x in rows}
                ]
                if len(candidates) != 1 or (
                    renamed and any(r["Id"] == current["Id"] for r in after)
                ):
                    changed("outlook.categories.write", "new_identity_readback")
                actual = candidates[0]
            else:
                actual = self.find(after, desired["Id"])
            if actual["Name"] != desired["Name"] or actual["Color"] != desired["Color"]:
                changed("outlook.categories.write", "readback")
            return {
                **self.item(actual),
                "item_assignments": "unchanged",
                **(
                    {"previous_name": current["Name"], "previous_id": current["Id"]}
                    if current and current["Name"] != actual["Name"]
                    else {}
                ),
            }
        except OutlookError as error:
            raise OutlookError(
                "write_outcome_unknown",
                "Category write was acknowledged but readback failed. Inspect the master list before retrying.",
                cause=error.as_dict(),
                intended_name=desired["Name"],
                previous_id=identifier,
            ) from None


class PlannerLabels:
    def __init__(self, api, plan_id):
        self.api, self.plan_id = api, plan_id
        self.path = "v1.0/planner/plans/" + q(plan_id) + "/details"

    def details(self, expected=None):
        data = self.api.entity("graph", self.path)
        review(data, expected)
        if data["id"] != self.plan_id:
            changed("planner.labels", "plan.id")
        descriptions = object_value(data.get("categoryDescriptions"), "planner.labels")
        for key, value in descriptions.items():
            if key.startswith("@"):
                continue
            if not re.fullmatch(r"category(?:[1-9]|1[0-9]|2[0-5])", key):
                changed("planner.labels", "categoryDescriptions.slot")
            if value is None:
                descriptions[key] = ""
            elif not isinstance(value, str):
                changed("planner.labels", "description")
        return data

    def item(self, data, identifier):
        slot(identifier)
        return {
            "id": identifier,
            "name": data["categoryDescriptions"].get(identifier, ""),
            "color": PLANNER_COLORS[int(identifier[8:]) - 1],
            "defined": bool(data["categoryDescriptions"].get(identifier)),
            "planId": self.plan_id,
            "version": version(data),
            "resource": "label",
            "source": "planner",
        }

    def listing(self):
        data = self.details()
        return {
            "items": [self.item(data, f"category{i}") for i in range(1, 26)],
            "version": version(data),
            "complete": True,
        }

    def tasks(self, max_pages):
        tasks, more, duplicates = self.api.scan(
            "graph", "beta/planner/plans/" + q(self.plan_id) + "/tasks", max_pages=max_pages
        )
        if more or duplicates:
            raise OutlookError(
                "incomplete_scan",
                "Cannot modify a plan label with an incomplete or changing task scan.",
            )
        for task in tasks:
            if task.get("planId") != self.plan_id:
                changed("planner.labels", "task.planId")
            version(task)
            applied_categories(task)
        return tasks

    def mutate(
        self,
        action,
        identifier=None,
        *,
        name=None,
        target=None,
        expected=None,
        max_pages=20,
        remove_from_tasks=False,
    ):
        if (
            action not in ("create", "rename", "delete", "recolor")
            or type(max_pages) is not int
            or not 1 <= max_pages <= 100
        ):
            raise OutlookError("invalid_argument", "Invalid label action or scan budget.")
        if action in ("create", "rename") and name is None:
            raise OutlookError("invalid_argument", "A label name is required.")
        data = self.details(expected)
        if name is not None:
            label_name(name)
        if action == "create":
            identifier = slot(target)
        else:
            slot(identifier)
        current = self.item(data, identifier)
        if name and any(
            k != identifier and v.casefold() == name.casefold()
            for k, v in data["categoryDescriptions"].items()
            if not k.startswith("@")
        ):
            raise OutlookError("label_exists", "Another plan slot already uses that name.")
        if action == "rename":
            changes = {identifier: name}
            tasks = []
        else:
            tasks = self.tasks(max_pages)
            users = [t for t in tasks if t.get("appliedCategories", {}).get(identifier) is True]
            if action == "create":
                if current["defined"] or users:
                    raise OutlookError(
                        "label_slot_in_use",
                        "This color slot is already named or applied to tasks; choose another color.",
                    )
                changes = {identifier: name}
            elif action == "delete":
                if users and not remove_from_tasks:
                    raise OutlookError(
                        "label_in_use",
                        "Use --remove-from-tasks to remove this label from the plan's tasks before clearing its name.",
                        task_count=len(users),
                    )
                changes = {identifier: ""}
            else:
                slot(target)
                if target == identifier:
                    return {**current, "unchanged": True}
                if not current["defined"]:
                    raise OutlookError("label_not_defined", "Name the label before recoloring it.")
                if self.item(data, target)["defined"] or any(
                    t.get("appliedCategories", {}).get(target) is True for t in tasks
                ):
                    raise OutlookError(
                        "label_slot_in_use",
                        "Destination color is already named or applied to tasks.",
                    )
                changes = {identifier: "", target: current["name"]}
        completed, pending_task = [], None
        try:
            if action in ("delete", "recolor"):
                for task in users:
                    pending_task = task["id"]
                    values = {identifier: False, **({target: True} if action == "recolor" else {})}
                    result = self.api.patch(
                        "graph",
                        "beta/planner/tasks/" + q(task["id"]),
                        {"appliedCategories": values},
                        "planner",
                        "task",
                        version(task),
                    )
                    completed.append(task["id"])
                    applied = applied_categories(result, write=True)
                    if any(applied.get(k, False) is not v for k, v in values.items()):
                        raise OutlookError(
                            "write_outcome_unknown",
                            "Task label readback did not match.",
                            task_id=task["id"],
                        )
            pending_task = None
            updated = self.api.patch(
                "graph",
                self.path,
                {"categoryDescriptions": changes},
                "planner",
                "labels",
                version(data),
            )
        except OutlookError as exc:
            if completed:
                raise OutlookError(
                    "partial_completion",
                    "Some task labels changed, but the plan operation did not finish. Inspect before any retry; no automatic rollback was attempted.",
                    completed_task_ids=completed,
                    plan_id=self.plan_id,
                    source_label=identifier,
                    target_label=target,
                    uncertain_task_id=pending_task,
                    cause=exc.as_dict(),
                ) from None
            if exc.code == "write_outcome_unknown":
                raise OutlookError(
                    "write_outcome_unknown",
                    "A Planner label write has an uncertain outcome. Inspect the identified task and plan before retrying.",
                    plan_id=self.plan_id,
                    uncertain_task_id=pending_task,
                    completed_task_ids=completed,
                    source_label=identifier,
                    target_label=target,
                    cause=exc.as_dict(),
                ) from None
            raise
        descriptions = object_value(
            updated.get("categoryDescriptions"), "planner.labels.write", write=True
        )
        if updated.get("id") != self.plan_id:
            changed("planner.labels.write", "plan.id", write=True)
        for key, value in changes.items():
            if descriptions.get(key, "") != value:
                raise OutlookError(
                    "write_outcome_unknown",
                    "Planner label changes were acknowledged but readback did not match.",
                    plan_id=self.plan_id,
                    completed_task_ids=completed,
                )
        result = self.item(updated, target if action == "recolor" else identifier)
        result.update(
            updated_task_count=len(completed),
            scope="plan",
            snapshot_scan=action != "rename",
            atomic=False,
        )
        if action == "delete":
            result["definition_cleared"] = True
        if action == "recolor":
            result["previous_slot"] = identifier
            result["source_label"] = self.item(updated, identifier)
        return result
