"""Local task payload validation; unknown fields never pass through to Microsoft."""

import re
import uuid
from datetime import date, datetime
from urllib.parse import urlsplit
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from .errors import OutlookError
from .validation import iso, nonempty

FIELDS = {
    "todo.list": {"displayName": str},
    "todo.task": {
        "title": str,
        "body": dict,
        "importance": str,
        "status": str,
        "categories": list,
        "isReminderOn": bool,
        "startDateTime": dict,
        "dueDateTime": dict,
        "reminderDateTime": dict,
        "recurrence": dict,
    },
    "todo.step": {"displayName": str, "isChecked": bool},
    "todo.link": {"webUrl": str, "applicationName": str, "displayName": str, "externalId": str},
    "todo.group": {"Name": str},
    "planner.plan": {"title": str},
    "planner.bucket": {"name": str, "orderHint": str},
    "planner.task": {
        "title": str,
        "bucketId": str,
        "orderHint": str,
        "percentComplete": int,
        "priority": int,
        "startDateTime": str,
        "dueDateTime": str,
        "assignments": dict,
        "appliedCategories": dict,
        "recurrence": dict,
        "isOnMyDay": bool,
    },
    "planner.details": {
        "description": str,
        "previewType": str,
        "checklist": dict,
        "references": dict,
    },
    "planner.labels": {"categoryDescriptions": dict},
}
NULLABLE = {"startDateTime", "dueDateTime", "reminderDateTime", "recurrence"}


def invalid(message):
    raise OutlookError("invalid_argument", message)


def identifier(value, field="id"):
    nonempty(value, field)
    if len(value) > 4096 or any(ord(c) < 32 for c in value) or value in (".", ".."):
        invalid(f"Invalid {field}.")
    return value


def user_id(value):
    try:
        uuid.UUID(value)
    except (ValueError, TypeError, AttributeError):
        invalid("Assignee must be 'self' or an Entra user UUID.")
    return value


def subset(value, allowed, field):
    if not isinstance(value, dict) or set(value) - set(allowed):
        invalid(f"Invalid or unknown fields in {field}.")


def web_url(value):
    u = urlsplit(value)
    if (
        u.scheme not in ("http", "https")
        or not u.hostname
        or u.username
        or u.password
        or any(ord(c) < 32 for c in value)
    ):
        invalid("References require an absolute HTTP(S) URL without embedded credentials.")


def date_value(value, field):
    subset(value, ("dateTime", "timeZone"), field)
    try:
        datetime.fromisoformat(value["dateTime"])
        ZoneInfo(value["timeZone"])
    except (KeyError, ValueError, TypeError, ZoneInfoNotFoundError):
        invalid(f"{field} requires dateTime and an IANA timeZone (for example Europe/Zurich).")


def recurrence(value, planner):
    allowed = ("schedule",) if planner else ("pattern", "range")
    subset(value, allowed, "recurrence")
    if planner:
        if "schedule" not in value:
            invalid("Planner recurrence needs schedule; null schedule stops the active series.")
        if value["schedule"] is None:
            return
        subset(value["schedule"], ("pattern", "patternStartDateTime"), "schedule")
        iso(value["schedule"].get("patternStartDateTime"))
        pattern = value["schedule"].get("pattern")
    else:
        pattern = value.get("pattern")
        r = value.get("range")
        subset(
            r,
            ("type", "startDate", "endDate", "numberOfOccurrences", "recurrenceTimeZone"),
            "range",
        )
        if r.get("type") not in ("noEnd", "endDate", "numbered"):
            invalid("Invalid recurrence range type.")
        try:
            start = date.fromisoformat(r["startDate"])
            if r["type"] == "endDate" and date.fromisoformat(r["endDate"]) < start:
                invalid("Recurrence end date precedes its start.")
            if r["type"] == "numbered" and (
                type(r.get("numberOfOccurrences")) is not int or r["numberOfOccurrences"] < 1
            ):
                invalid("Numbered recurrence needs a positive occurrence count.")
        except (KeyError, ValueError, TypeError):
            invalid("Recurrence needs ISO dates and its range-specific fields.")
    subset(
        pattern,
        ("type", "interval", "month", "dayOfMonth", "daysOfWeek", "firstDayOfWeek", "index"),
        "pattern",
    )
    if (
        pattern.get("type")
        not in (
            "daily",
            "weekly",
            "absoluteMonthly",
            "relativeMonthly",
            "absoluteYearly",
            "relativeYearly",
        )
        or type(pattern.get("interval")) is not int
        or pattern["interval"] < 1
    ):
        invalid("Recurrence needs a supported pattern type and positive interval.")
    days = {"sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"}
    if "daysOfWeek" in pattern and (
        not isinstance(pattern["daysOfWeek"], list)
        or not pattern["daysOfWeek"]
        or any(x not in days for x in pattern["daysOfWeek"])
    ):
        invalid("Invalid recurrence daysOfWeek.")
    if "firstDayOfWeek" in pattern and pattern["firstDayOfWeek"] not in days:
        invalid("Invalid firstDayOfWeek.")
    if "index" in pattern and pattern["index"] not in (
        "first",
        "second",
        "third",
        "fourth",
        "last",
    ):
        invalid("Invalid recurrence index.")
    for key, upper in (("month", 12), ("dayOfMonth", 31)):
        if key in pattern and (type(pattern[key]) is not int or not 1 <= pattern[key] <= upper):
            invalid(f"Invalid recurrence {key}.")
    required = {
        "weekly": ("daysOfWeek",),
        "absoluteMonthly": ("dayOfMonth",),
        "relativeMonthly": ("daysOfWeek", "index"),
        "absoluteYearly": ("month", "dayOfMonth"),
        "relativeYearly": ("month", "daysOfWeek", "index"),
    }.get(pattern["type"], ())
    if any(k not in pattern for k in required):
        invalid("Recurrence pattern is missing required fields.")


def validate_payload(kind, data, *, create=False):
    subset(data, FIELDS[kind], kind)
    for key, value in data.items():
        if value is None and key in NULLABLE:
            if kind.startswith("planner") and key == "recurrence":
                invalid('Stop Planner recurrence with {"schedule":null}.')
            continue
        if type(value) is not FIELDS[kind][key]:
            invalid(f"Invalid type for {key}.")
        if isinstance(value, str) and "\x00" in value:
            invalid(f"NUL is not allowed in {key}.")
        if key in ("title", "displayName", "Name", "name"):
            nonempty(value, key)
        if key in ("startDateTime", "dueDateTime", "reminderDateTime"):
            if kind.startswith("todo"):
                date_value(value, key)
            else:
                iso(value)
        if key == "body":
            subset(value, ("content", "contentType"), "body")
            if value.get("contentType") not in ("text", "html") or not isinstance(
                value.get("content"), str
            ):
                invalid("Body needs content and text/html contentType.")
        if key == "status" and value not in (
            "notStarted",
            "inProgress",
            "completed",
            "waitingOnOthers",
            "deferred",
        ):
            invalid("Invalid To Do status.")
        if key == "importance" and value not in ("low", "normal", "high"):
            invalid("Invalid To Do importance.")
        if key == "percentComplete" and not 0 <= value <= 100:
            invalid("percentComplete must be 0–100.")
        if key == "priority" and not 0 <= value <= 10:
            invalid("Planner priority must be 0–10.")
        if key == "categories" and any(not isinstance(x, str) for x in value):
            invalid("Categories must contain strings.")
        if key == "recurrence":
            recurrence(value, kind.startswith("planner"))
        if key == "assignments":
            for uid, assignment in value.items():
                user_id(uid)
                if assignment is None:
                    continue
                subset(assignment, ("@odata.type", "orderHint"), "assignment")
                if assignment.get(
                    "@odata.type"
                ) != "#microsoft.graph.plannerAssignment" or not isinstance(
                    assignment.get("orderHint"), str
                ):
                    invalid("Assignments require plannerAssignment @odata.type and orderHint.")
        if key in ("appliedCategories", "categoryDescriptions"):
            for label, entry in value.items():
                if not re.fullmatch(r"category(?:[1-9]|1[0-9]|2[0-5])", label) or type(
                    entry
                ) is not (bool if key == "appliedCategories" else str):
                    invalid(
                        "Labels use category1–category25 and boolean values (task) or strings (plan)."
                    )
        if key == "previewType" and value not in (
            "automatic",
            "noPreview",
            "checklist",
            "description",
            "reference",
        ):
            invalid("Invalid Planner previewType.")
        if key in ("checklist", "references"):
            for entry_id, entry in value.items():
                if key == "checklist":
                    user_id(entry_id)
                elif not isinstance(entry_id, str) or not entry_id:
                    invalid("Reference keys must be encoded URLs.")
                if entry is None:
                    continue
                allowed = (
                    ("@odata.type", "title", "isChecked", "orderHint")
                    if key == "checklist"
                    else ("@odata.type", "alias", "type", "previewPriority")
                )
                subset(entry, allowed, key)
                wanted = (
                    "#microsoft.graph.plannerChecklistItem"
                    if key == "checklist"
                    else "#microsoft.graph.plannerExternalReference"
                )
                if entry.get("@odata.type") != wanted:
                    invalid(f"{key} entry needs its Graph @odata.type.")
                for field, fieldvalue in entry.items():
                    if type(fieldvalue) is not (bool if field == "isChecked" else str):
                        invalid(f"Invalid {key} entry field type.")
        if key == "webUrl":
            web_url(value)
    if not data:
        invalid("At least one editable field is required.")
    if create:
        required = {
            "todo.list": "displayName",
            "todo.task": "title",
            "todo.step": "displayName",
            "todo.link": "webUrl",
            "todo.group": "Name",
            "planner.plan": "title",
            "planner.bucket": "name",
            "planner.task": "title",
        }.get(kind)
        if required and required not in data:
            invalid(f"Creation requires {required}.")
    return data
