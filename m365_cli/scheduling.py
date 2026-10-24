"""Read-only multi-person free/busy and conservative local meeting-slot selection."""

import math
import re
from datetime import UTC, datetime, time, timedelta
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from babel.core import get_global

from .api import check_response, envelope, typed
from .contracts import array_value, changed, object_value
from .errors import OutlookError
from .validation import email, iso

ACTION = "GetUserAvailabilityInternal"
STATES = {"0": "free", "1": "tentative", "2": "busy", "3": "out_of_office", "4": "unknown"}
DAYS = ("mon", "tue", "wed", "thu", "fri", "sat", "sun")


def zone(value):
    try:
        return ZoneInfo(get_global("windows_zone_mapping").get(value, value))
    except (ZoneInfoNotFoundError, ValueError, TypeError):
        raise OutlookError(
            "invalid_timezone", "Use an IANA or recognized Microsoft timezone."
        ) from None


def clock(value):
    if not isinstance(value, str) or not re.fullmatch(r"(?:[01][0-9]|2[0-3]):[0-5][0-9]", value):
        raise OutlookError("invalid_argument", "Daily hours must use HH:MM (00:00–23:59).")
    return time.fromisoformat(value)


def register(cal):
    for command in ("availability", "find-slots"):
        p = cal.add_parser(
            command,
            help="Read multiple people's free/busy"
            if command == "availability"
            else "Find common meeting times; never books or sends",
        )
        p.set_defaults(scheduling=True)
        p.add_argument(
            "--attendee",
            action="append",
            default=[],
            help="Required person's or resource's SMTP address; repeat as needed",
        )
        p.add_argument("--optional-attendee", action="append", default=[])
        p.add_argument(
            "--exclude-self",
            action="store_true",
            help="Do not automatically include the signed-in organizer",
        )
        p.add_argument("--start", required=True, help="Window start with explicit UTC offset")
        p.add_argument("--end", required=True, help="Exclusive window end; maximum 31 days")
        p.add_argument(
            "--interval",
            type=int,
            default=15,
            help="Conservative free/busy granularity in minutes (5–60)",
        )
        p.add_argument(
            "--time-zone",
            default="UTC",
            help="Output and daily-hours timezone; IANA or Microsoft name (default UTC)",
        )
        if command == "find-slots":
            p.add_argument(
                "--duration", required=True, type=int, help="Meeting duration in minutes (5–1440)"
            )
            p.add_argument(
                "--work-start",
                default="09:00",
                help="Earliest daily start in --time-zone (default 09:00)",
            )
            p.add_argument(
                "--work-end",
                default="17:00",
                help="Latest daily finish in --time-zone (default 17:00)",
            )
            p.add_argument(
                "--weekdays",
                default="mon,tue,wed,thu,fri",
                help="Comma-separated days; use all for every day",
            )
            p.add_argument(
                "--max-results", type=int, default=10, help="Maximum ranked candidates (1–100)"
            )
            p.add_argument(
                "--allow-tentative",
                action="store_true",
                help="Accept tentative slots explicitly; busy/OOF/unknown still block",
            )
            p.add_argument(
                "--respect-working-hours",
                action="store_true",
                help="Also require known published working hours for required attendees; unknown/custom zones block candidates",
            )


def prepare(args):
    if args.account:
        email(args.account)
    for address in [*args.attendee, *args.optional_attendee]:
        email(address)
    if args.exclude_self and not args.attendee:
        raise OutlookError(
            "invalid_argument", "With --exclude-self, supply at least one required --attendee."
        )
    if len({v.casefold() for v in [*args.attendee, *args.optional_attendee]}) > 100:
        raise OutlookError(
            "invalid_argument",
            "At most 100 distinct attendees are supported, including the organizer.",
        )
    start, end = iso(args.start).astimezone(UTC), iso(args.end).astimezone(UTC)
    if not start < end or end - start > timedelta(days=31):
        raise OutlookError(
            "invalid_range", "Availability windows must be positive and no longer than 31 days."
        )
    if type(args.interval) is not int or not 5 <= args.interval <= 60:
        raise OutlookError("invalid_argument", "--interval must be 5–60 minutes.")
    zone(args.time_zone)
    if args.command == "find-slots":
        if type(args.duration) is not int or not 5 <= args.duration <= 1440:
            raise OutlookError("invalid_argument", "--duration must be 5–1440 minutes.")
        if type(args.max_results) is not int or not 1 <= args.max_results <= 100:
            raise OutlookError("invalid_argument", "--max-results must be 1–100.")
        if clock(args.work_start) >= clock(args.work_end):
            raise OutlookError(
                "invalid_argument",
                "Daily work end must follow start; overnight work windows are not supported.",
            )
        days = DAYS if args.weekdays == "all" else args.weekdays.split(",")
        if not days or any(d not in DAYS for d in days) or len(set(days)) != len(days):
            raise OutlookError(
                "invalid_argument", "Use unique lowercase weekday names mon,...,sun, or all."
            )
    return args


def response_check(response):
    """Only this read's exact per-recipient error array may be returned as partial data."""
    data = response.get("body") if isinstance(response, dict) else None
    body = data.get("Body") if isinstance(data, dict) else None
    if isinstance(body, dict) and isinstance(body.get("Responses"), list):
        # Global HTTP/auth/errors and error fields outside the recipient array retain
        # the normal transport checks. The adapter validates each member separately.
        filtered = {**data, "Body": {k: v for k, v in body.items() if k != "Responses"}}
        check_response({**response, "body": filtered})
        return data
    return check_response(response)


def working_hours(raw):
    if raw is None:
        return {"usable": False, "reason": "not_returned"}
    value = object_value(raw, "availability.working_hours")
    start, end, days, tz = (
        value.get(k)
        for k in (
            "WorkHoursStartTimeInMinutes",
            "WorkHoursEndTimeInMinutes",
            "WorkDays",
            "WorkingHoursTimeZoneId",
        )
    )
    if (
        type(start) is not int
        or type(end) is not int
        or not 0 <= start < end <= 1440
        or type(days) is not int
        or not 0 <= days <= 127
        or not isinstance(tz, str)
    ):
        return {"usable": False, "reason": "unrecognized_working_hours"}
    result = {
        "start_minute": start,
        "end_minute": end,
        "weekdays": [day for i, day in enumerate(DAYS) if days & (1 << ((i + 1) % 7))],
        "time_zone": tz,
    }
    try:
        result.update(usable=True, time_zone=zone(tz).key)
    except OutlookError:
        result.update(usable=False, reason="unrecognized_working_hours_timezone")
    return result


def attendees(args, account):
    required = [*([account] if not args.exclude_self else []), *args.attendee]
    result, seen = [], set()
    for role, values in (("required", required), ("optional", args.optional_attendee)):
        for address in values:
            email(address)
            key = address.casefold()
            if key not in seen:
                result.append({"email": address, "role": role})
                seen.add(key)
    if len(result) > 100:
        raise OutlookError(
            "invalid_argument",
            "At most 100 distinct attendees are supported, including the organizer.",
        )
    return result


def utc_string(value):
    return value.astimezone(UTC).isoformat().replace("+00:00", "Z")


def snapshot(api, args):
    prepare(args)
    people = attendees(args, api.session.account)
    start, end = iso(args.start).astimezone(UTC), iso(args.end).astimezone(UTC)
    seconds = args.interval * 60
    # Align outward in UTC; a partially busy merged interval is never assumed free.
    query_start = datetime.fromtimestamp(math.floor(start.timestamp() / seconds) * seconds, UTC)
    query_end = datetime.fromtimestamp(math.ceil(end.timestamp() / seconds) * seconds, UTC)
    count = int((query_end - query_start).total_seconds() // seconds)
    if count * len(people) > 1_000_000:
        raise OutlookError(
            "scan_budget_exceeded", "Use fewer attendees or a shorter availability window."
        )
    body = typed(
        "GetUserAvailabilityRequest",
        MailboxDataArray=[
            typed("MailboxData", Email=typed("EmailAddress", Address=p["email"])) for p in people
        ],
        FreeBusyViewOptions=typed(
            "FreeBusyViewOptions",
            RequestedView="FreeBusyMerged",
            TimeWindow=typed(
                "Duration", StartTime=utc_string(query_start), EndTime=utc_string(query_end)
            ),
            MergedFreeBusyIntervalInMinutes=args.interval,
        ),
    )
    data = api.session.request(
        ACTION, {"request": envelope(ACTION, body, time_zone="UTC")}, native=True
    )
    result = object_value(data.get("Body"), "availability.body")
    if result.get("ResponseCode") != "NoError" or result.get("ResponseClass") != "Success":
        changed("availability.body", "acknowledgement")
    rows = array_value(result.get("Responses"), "availability.responses")
    if len(rows) != len(people):
        changed("availability.responses", "attendee_count")
    display = zone(args.time_zone)
    for person, row in zip(people, rows, strict=True):
        message = object_value(
            row.get("ResponseMessage"),
            "availability.recipient",
            fields=("ResponseCode", "ResponseClass"),
        )
        code = message["ResponseCode"]
        if code != "NoError":
            if message["ResponseClass"] not in ("Error", "Warning") or not re.fullmatch(
                r"[A-Za-z][A-Za-z0-9_]{0,199}", code
            ):
                changed("availability.recipient", "error_code")
            person.update(
                status="unavailable",
                error={"code": code},
                availability_view="4" * count,
                working_hours={"usable": False, "reason": "availability_unavailable"},
            )
        else:
            if message["ResponseClass"] != "Success":
                changed("availability.recipient", "ResponseClass")
            view = object_value(row.get("CalendarView"), "availability.view")
            stream = view.get("MergedFreeBusy")
            if (
                view.get("FreeBusyViewType")
                not in ("MergedOnly", "FreeBusyMerged", "DetailedMerged")
                or not isinstance(stream, str)
                or len(stream) != count
                or any(c not in STATES for c in stream)
            ):
                changed("availability.view", "merged_stream")
            person.update(
                status="ok",
                availability_view=stream,
                working_hours=working_hours(view.get("WorkingHours")),
            )
        stream = person["availability_view"]
        person["complete"] = "4" not in stream
        blocks = []
        for i, char in enumerate(stream):
            a = max(start, query_start + timedelta(seconds=i * seconds))
            b = min(end, query_start + timedelta(seconds=(i + 1) * seconds))
            state = STATES[char]
            if blocks and blocks[-1]["status"] == state:
                blocks[-1]["end"] = b.astimezone(display).isoformat()
            else:
                blocks.append(
                    {
                        "start": a.astimezone(display).isoformat(),
                        "end": b.astimezone(display).isoformat(),
                        "status": state,
                    }
                )
        person["blocks"] = blocks
    return {
        "source": "outlook",
        "snapshot": True,
        "observed_at": utc_string(datetime.now(UTC)),
        "window_start": start.astimezone(display).isoformat(),
        "window_end": end.astimezone(display).isoformat(),
        "query_start": utc_string(query_start),
        "query_end": utc_string(query_end),
        "time_zone": display.key,
        "interval_minutes": args.interval,
        "availability_legend": STATES,
        "attendees": people,
        "complete": all(p["complete"] for p in people),
        "partial": any(not p["complete"] for p in people),
        "working_hours_enforced": False,
        "reference_unavailable": "free_busy_is_a_snapshot_without_a_stable_item_id_or_permalink",
    }


def inside_hours(start, end, tz, days, first, last):
    a, b = start.astimezone(tz), end.astimezone(tz)
    if a.date() != (b - timedelta(microseconds=1)).date() or DAYS[a.weekday()] not in days:
        return False
    lower = datetime.combine(a.date(), time(), tzinfo=tz) + timedelta(minutes=first)
    upper = datetime.combine(a.date(), time(), tzinfo=tz) + timedelta(minutes=last)
    # Compare instants, including DST overlaps/gaps, rather than same-zone wall time.
    return start >= lower.astimezone(UTC) and end <= upper.astimezone(UTC)


def find_slots(data, args):
    start, end = iso(data["window_start"]).astimezone(UTC), iso(data["window_end"]).astimezone(UTC)
    anchor = iso(data["query_start"])
    step, duration = timedelta(minutes=args.interval), timedelta(minutes=args.duration)
    display = zone(args.time_zone)
    days = DAYS if args.weekdays == "all" else args.weekdays.split(",")
    first, last = (clock(v).hour * 60 + clock(v).minute for v in (args.work_start, args.work_end))
    accepted = {"0", "1"} if args.allow_tentative else {"0"}
    candidates = []
    cursor = anchor
    while cursor + duration <= end:
        finish = cursor + duration
        if cursor >= start and inside_hours(cursor, finish, display, days, first, last):
            offset = int((cursor - anchor) / step)
            length = math.ceil(duration / step)
            statuses = []
            for person in data["attendees"]:
                values = set(person["availability_view"][offset : offset + length])
                state = (
                    "unknown"
                    if not values or "4" in values
                    else "out_of_office"
                    if "3" in values
                    else "busy"
                    if "2" in values
                    else "tentative"
                    if "1" in values
                    else "free"
                )
                available = bool(values) and values <= accepted
                work = "not_checked"
                if args.respect_working_hours:
                    hours = person["working_hours"]
                    work = (
                        "unknown"
                        if not hours["usable"]
                        else "inside"
                        if inside_hours(
                            cursor,
                            finish,
                            zone(hours["time_zone"]),
                            hours["weekdays"],
                            hours["start_minute"],
                            hours["end_minute"],
                        )
                        else "outside"
                    )
                    available = available and work == "inside"
                statuses.append(
                    {
                        "email": person["email"],
                        "role": person["role"],
                        "status": state,
                        "working_hours": work,
                        "acceptable": available,
                    }
                )
            if all(p["acceptable"] for p in statuses if p["role"] == "required"):
                optional = [p for p in statuses if p["role"] == "optional"]
                candidates.append(
                    {
                        "start": cursor.astimezone(display).isoformat(),
                        "end": finish.astimezone(display).isoformat(),
                        "duration_minutes": args.duration,
                        "required_available": True,
                        "optional_available_count": sum(p["acceptable"] for p in optional),
                        "optional_total": len(optional),
                        "attendees": statuses,
                    }
                )
        cursor += step
    # Stable sorting keeps chronological order when optional attendance ties, even
    # across DST's repeated local hour. UTC generation already establishes that order.
    candidates.sort(key=lambda c: -c["optional_available_count"])
    return {
        **data,
        "slots": candidates[: args.max_results],
        "candidate_count": len(candidates),
        "has_more": len(candidates) > args.max_results,
        "constraints": {
            "duration_minutes": args.duration,
            "work_start": args.work_start,
            "work_end": args.work_end,
            "weekdays": list(days),
            "allow_tentative": args.allow_tentative,
            "respect_published_working_hours": args.respect_working_hours,
        },
        "working_hours_enforced": args.respect_working_hours,
        "ranking": "optional_attendance_then_earliest",
        "booking_submitted": False,
        "unknown_required_attendees": [
            p["email"]
            for p in data["attendees"]
            if p["role"] == "required"
            and (
                not p["complete"] or args.respect_working_hours and not p["working_hours"]["usable"]
            )
        ],
        "note": "Candidates are free/busy snapshots, not reservations. Recheck before calendar create; creating with attendees sends invitations.",
    }


def dispatch(api, args):
    data = snapshot(api, args)
    return find_slots(data, args) if args.command == "find-slots" else data
