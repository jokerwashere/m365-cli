"""Input checks shared by the CLI and Python API."""

import re
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from babel.core import get_global

from .errors import OutlookError


def nonempty(value, field):
    if not isinstance(value, str) or not value.strip() or "\x00" in value:
        raise OutlookError(
            "invalid_argument", f"{field} must be a nonempty string without NUL characters."
        )
    return value


def email(value):
    if (
        not isinstance(value, str)
        or len(value) > 320
        or not re.fullmatch(r'[^@\s<>,;:"\x00-\x1f\x7f]+@[^@\s<>,;:"\x00-\x1f\x7f]+', value)
    ):
        raise OutlookError(
            "invalid_recipient",
            "Use one plain email address per recipient flag, without a display name or separators.",
        )
    local, domain = value.rsplit("@", 1)
    try:
        domain = domain.encode("idna").decode("ascii")
    except UnicodeError:
        raise OutlookError("invalid_recipient", "The recipient domain is malformed.") from None
    if (
        len(local) > 64
        or local.startswith(".")
        or local.endswith(".")
        or ".." in local
        or len(domain) > 253
        or not all(
            1 <= len(label) <= 63
            and re.fullmatch(r"[A-Za-z0-9](?:[A-Za-z0-9-]*[A-Za-z0-9])?", label)
            for label in domain.split(".")
        )
    ):
        raise OutlookError("invalid_recipient", "The recipient email address is malformed.")
    return value


def paging(limit, offset=0):
    if type(limit) is not int or not 1 <= limit <= 500 or type(offset) is not int or offset < 0:
        raise OutlookError("invalid_paging", "Limit must be 1–500 and offset must be nonnegative.")


def iso(value):
    try:
        result = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except (ValueError, TypeError, AttributeError):
        raise OutlookError(
            "invalid_datetime", "Use an ISO 8601 datetime with an explicit UTC offset."
        ) from None
    if result.tzinfo is None or "T" not in value:
        raise OutlookError(
            "timezone_required", "Use an ISO 8601 datetime with T and an explicit UTC offset."
        )
    return result


def validate_range(start, end, *, window=False, all_day=False):
    first, last = iso(start), iso(end)
    if last <= first:
        raise OutlookError("invalid_range", "End must be later than start.")
    if window and last - first > timedelta(days=366):
        raise OutlookError(
            "invalid_range", "Calendar queries are limited to 366 days; use smaller windows."
        )
    if all_day and (
        any((d.hour, d.minute, d.second, d.microsecond) != (0, 0, 0, 0) for d in (first, last))
    ):
        raise OutlookError(
            "invalid_range",
            "All-day events must start and end at local midnight; end is exclusive.",
        )


def calendar_range(start, end, *, all_day=False, time_zone="UTC"):
    validate_range(start, end)
    if not all_day:
        return
    zone_name = get_global("windows_zone_mapping").get(time_zone, time_zone)
    try:
        zone = ZoneInfo(zone_name)
    except (ZoneInfoNotFoundError, TypeError):
        raise OutlookError(
            "timezone_unknown", "Cannot validate the mailbox timezone for an all-day event."
        ) from None
    local = [iso(value).astimezone(zone) for value in (start, end)]
    if any((d.hour, d.minute, d.second, d.microsecond) != (0, 0, 0, 0) for d in local):
        raise OutlookError(
            "timezone_mismatch",
            "All-day boundaries must be midnight in the mailbox timezone shown by auth status.",
            time_zone=time_zone,
        )
