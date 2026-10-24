"""Synthetic free/busy privacy, partial-result, scheduling and timezone contracts."""

import copy
import json
import tempfile
import unittest
from unittest.mock import patch

from test_transport import ready_session, reply

from m365_cli.api import API, check_response
from m365_cli.cli import parser, prepare, run
from m365_cli.errors import OutlookError
from m365_cli.maintainer import inventory
from m365_cli.scheduling import attendees, find_slots, response_check, snapshot, working_hours

ACCOUNT = "self@example.com"
HOURS = {
    "WorkHoursStartTimeInMinutes": 540,
    "WorkHoursEndTimeInMinutes": 1020,
    "WorkDays": 62,
    "WorkingHoursTimeZoneId": "UTC",
}


def row(stream="0000", code="NoError", hours=HOURS):
    return {
        "ResponseMessage": {
            "ResponseCode": code,
            "ResponseClass": "Success" if code == "NoError" else "Error",
            "MessageText": "Do not expose private server text",
        },
        "CalendarView": {
            "FreeBusyViewType": "FreeBusyMerged" if code == "NoError" else "None",
            "MergedFreeBusy": stream if code == "NoError" else None,
            "WorkingHours": hours,
            "Items": [
                {"Subject": "Private subject", "ItemId": {"Id": "opaque", "ChangeKey": "v1"}}
            ],
        },
    }


def response(*rows):
    return {
        "Header": {},
        "Body": {"ResponseCode": "NoError", "ResponseClass": "Success", "Responses": list(rows)},
    }


class FakeSession:
    account = ACCOUNT

    def __init__(self, data):
        self.data, self.calls = data, []

    def request(self, *args, **kwargs):
        self.calls.append((args, kwargs))
        return copy.deepcopy(self.data)


def args(command="availability", *extra):
    words = [
        "--account",
        ACCOUNT,
        "calendar",
        command,
        "--start",
        "2031-02-03T09:00:00Z",
        "--end",
        "2031-02-03T10:00:00Z",
    ]
    if command == "find-slots":
        words += ["--duration", "30"]
    return prepare(parser().parse_args([*words, *extra]))


class SchedulingTests(unittest.TestCase):
    def error(self, code, action):
        with self.assertRaises(OutlookError) as caught:
            action()
        self.assertEqual(caught.exception.code, code)
        return caught.exception

    def snapshot(self, options, *rows):
        return snapshot(API(FakeSession(response(*rows))), options)

    def test_new_commands_are_readonly_and_have_specific_maintainer_suite(self):
        commands = {x["command"]: x for x in inventory()["capabilities"]}
        for command in ("availability", "find-slots"):
            self.assertFalse(commands["calendar " + command]["mutation"])
            self.assertEqual(commands["calendar " + command]["validation_suite"], "availability")

    def test_input_errors_happen_before_connecting(self):
        invalid = [
            (["--attendee", "not an email"], "invalid_recipient"),
            (["--interval", "0"], "invalid_argument"),
            (["--interval", "61"], "invalid_argument"),
            (["--time-zone", "Mars/Nowhere"], "invalid_timezone"),
            (["--start", "2031-02-03T09:00:00"], "timezone_required"),
            (["--end", "2031-04-01T00:00:00Z"], "invalid_range"),
            (["--exclude-self"], "invalid_argument"),
        ]
        with patch("m365_cli.cli.Session") as session:
            for words, code in invalid:
                self.error(code, lambda words=words: args("availability", *words))
            session.assert_not_called()
        for words in (
            ["--duration", "0"],
            ["--max-results", "101"],
            ["--weekdays", "mon,mon"],
            ["--weekdays", "Mon"],
            ["--work-start", "9am"],
            ["--work-start", "17:00", "--work-end", "09:00"],
        ):
            self.error("invalid_argument", lambda words=words: args("find-slots", *words))

    def test_identity_deduplication_self_and_required_win(self):
        options = args(
            "availability",
            "--attendee",
            "other@example.com",
            "--optional-attendee",
            ACCOUNT.upper(),
            "--optional-attendee",
            "OTHER@example.com",
        )
        self.assertEqual(
            attendees(options, ACCOUNT),
            [
                {"email": ACCOUNT, "role": "required"},
                {"email": "other@example.com", "role": "required"},
            ],
        )
        options.exclude_self = True
        self.assertEqual(attendees(options, ACCOUNT)[0]["email"], "other@example.com")

    def test_total_limit_includes_implicit_organizer(self):
        options = args()
        options.attendee = [f"p{i}@example.com" for i in range(100)]
        s = FakeSession(response())
        self.error("invalid_argument", lambda: snapshot(API(s), options))
        self.assertEqual(s.calls, [])

    def test_exact_native_wrapper_and_privacy(self):
        s = FakeSession(response(row("0123")))
        result = snapshot(API(s), args())
        call, kw = s.calls[0]
        self.assertEqual(call[0], "GetUserAvailabilityInternal")
        self.assertTrue(kw["native"])
        request = call[1]["request"]
        self.assertEqual(request["__type"], "GetUserAvailabilityInternalJsonRequest:#Exchange")
        self.assertEqual(request["Header"]["TimeZoneContext"]["TimeZoneDefinition"]["Id"], "UTC")
        self.assertEqual(request["Body"]["FreeBusyViewOptions"]["RequestedView"], "FreeBusyMerged")
        self.assertEqual(
            [b["status"] for b in result["attendees"][0]["blocks"]],
            ["free", "tentative", "busy", "out_of_office"],
        )
        text = json.dumps(result)
        for secret in ("Private subject", "opaque", "Do not expose private"):
            self.assertNotIn(secret, text)
        self.assertIn("reference_unavailable", result)

    def test_query_rounding_and_blocks_are_clipped_to_requested_window(self):
        options = args(
            "availability", "--start", "2031-02-03T09:03:00Z", "--end", "2031-02-03T09:47:00Z"
        )
        result = self.snapshot(options, row())
        self.assertEqual(result["query_start"], "2031-02-03T09:00:00Z")
        self.assertEqual(result["query_end"], "2031-02-03T10:00:00Z")
        self.assertEqual(
            result["attendees"][0]["blocks"],
            [
                {
                    "start": "2031-02-03T09:03:00+00:00",
                    "end": "2031-02-03T09:47:00+00:00",
                    "status": "free",
                }
            ],
        )

    def test_exact_response_count_and_stream_required(self):
        options = args()
        for data in (
            response(),
            response(row(), row()),
            response(row("")),
            response(row("000")),
            response(row("00000")),
            response(row("0005")),
        ):
            self.error("api_changed", lambda data=data: snapshot(API(FakeSession(data)), options))
        data = response(row())
        del data["Body"]["Responses"][0]["CalendarView"]["MergedFreeBusy"]
        self.error("api_changed", lambda: snapshot(API(FakeSession(data)), options))

    def test_success_without_view_or_with_conflicting_ack_fails_closed(self):
        for field, value in (
            ("CalendarView", None),
            ("ResponseMessage", {"ResponseCode": "NoError", "ResponseClass": "Error"}),
        ):
            value = response({**row(), field: value})
            self.error("api_changed", lambda value=value: snapshot(API(FakeSession(value)), args()))

    def test_expected_per_recipient_failures_preserve_other_results(self):
        options = args("availability", "--attendee", "denied@example.com")
        data = self.snapshot(options, row(), row(code="ErrorNoFreeBusyAccess"))
        self.assertTrue(data["partial"])
        self.assertFalse(data["complete"])
        self.assertEqual(data["attendees"][0]["status"], "ok")
        denied = data["attendees"][1]
        self.assertEqual(denied["availability_view"], "4444")
        self.assertEqual(denied["error"], {"code": "ErrorNoFreeBusyAccess"})

    def test_transport_relaxes_only_recipient_array_for_this_read(self):
        raw = {"status": 200, "body": response(row(), row(code="ErrorMailRecipientNotFound"))}
        self.error("owa_error", lambda: check_response(raw))
        self.assertIs(response_check(raw), raw["body"])
        for field, value in (
            ("ResponseCode", "ErrorAccessDenied"),
            ("ExceptionName", "SyntheticError"),
        ):
            bad = copy.deepcopy(raw)
            bad["body"]["Body"][field] = value
            self.error("owa_error", lambda bad=bad: response_check(bad))
        self.error("session_rejected", lambda: response_check({**raw, "status": 401}))
        self.error("rate_limited", lambda: response_check({**raw, "status": 429}))
        self.error("unexpected_response", lambda: response_check({"status": 200, "body": {}}))

    def test_real_session_routes_availability_and_preserves_other_read_errors(self):
        s = ready_session(ACCOUNT)
        s.http.request.return_value = reply(response(row(code="ErrorMailRecipientNotFound")))
        self.assertIn(
            "Body", s.request("GetUserAvailabilityInternal", {"request": {}}, native=True)
        )
        self.error("owa_error", lambda: s.request("GetItem", {}))
        self.assertEqual(s.http.request.call_count, 2)

    def test_all_required_busy_unknown_or_oof_prevent_candidates(self):
        options = args("find-slots", "--attendee", "other@example.com")
        for stream in ("2222", "3333", "4444"):
            data = self.snapshot(options, row(), row(stream))
            self.assertEqual(find_slots(data, options)["slots"], [])

    def test_unknown_optional_is_not_free_and_does_not_hide_known_candidates(self):
        options = args("find-slots", "--optional-attendee", "other@example.com")
        data = self.snapshot(options, row(), row(code="ErrorNoFreeBusyAccess"))
        result = find_slots(data, options)
        self.assertEqual(result["candidate_count"], 3)
        self.assertFalse(result["complete"])
        self.assertFalse(result["booking_submitted"])
        self.assertEqual(result["slots"][0]["optional_available_count"], 0)
        self.assertEqual(result["slots"][0]["attendees"][1]["status"], "unknown")

    def test_tentative_requires_explicit_opt_in(self):
        options = args("find-slots")
        data = self.snapshot(options, row("1111"))
        self.assertFalse(find_slots(data, options)["slots"])
        options.allow_tentative = True
        result = find_slots(data, options)
        self.assertEqual(len(result["slots"]), 3)
        self.assertEqual(result["slots"][0]["attendees"][0]["status"], "tentative")

    def test_partial_nodata_only_blocks_affected_intervals(self):
        options = args("find-slots")
        data = self.snapshot(options, row("4000"))
        result = find_slots(data, options)
        self.assertEqual(len(result["slots"]), 2)
        self.assertEqual(result["slots"][0]["start"], "2031-02-03T09:15:00+00:00")
        self.assertEqual(result["unknown_required_attendees"], [ACCOUNT])

    def test_optional_attendance_ranks_before_chronology_and_limit_is_explicit(self):
        options = args(
            "find-slots", "--optional-attendee", "other@example.com", "--max-results", "1"
        )
        data = self.snapshot(options, row(), row("2200"))
        result = find_slots(data, options)
        self.assertEqual(result["candidate_count"], 3)
        self.assertTrue(result["has_more"])
        self.assertEqual(result["slots"][0]["start"], "2031-02-03T09:30:00+00:00")

    def test_nonmultiple_duration_checks_entire_overlapping_merged_intervals(self):
        options = args("find-slots", "--duration", "20")
        data = self.snapshot(options, row("0020"))
        result = find_slots(data, options)
        self.assertEqual(result["candidate_count"], 1)
        self.assertEqual(result["slots"][0]["end"], "2031-02-03T09:20:00+00:00")

    def test_busy_just_after_meeting_end_in_same_merged_interval_is_conservative(self):
        options = args("find-slots", "--duration", "20")
        data = self.snapshot(options, row("0200"))
        self.assertEqual(
            find_slots(data, options)["slots"][0]["start"], "2031-02-03T09:30:00+00:00"
        )

    def test_local_workday_and_hours_are_applied_in_requested_timezone(self):
        options = args("find-slots", "--time-zone", "America/New_York")
        data = self.snapshot(options, row())
        self.assertEqual(find_slots(data, options)["slots"], [])
        options.work_start, options.work_end = "04:00", "05:00"
        result = find_slots(data, options)
        self.assertEqual(len(result["slots"]), 3)
        self.assertTrue(result["slots"][0]["start"].endswith("-05:00"))
        options.weekdays = "tue"
        self.assertFalse(find_slots(data, options)["slots"])

    def test_published_hours_day_mask_and_unknown_custom_timezone(self):
        hours = working_hours(
            {**HOURS, "WorkDays": 3, "WorkingHoursTimeZoneId": "W. Europe Standard Time"}
        )
        self.assertTrue(hours["usable"])
        self.assertEqual(hours["weekdays"], ["mon", "sun"])
        self.assertEqual(hours["time_zone"], "Europe/Berlin")
        for value in (
            None,
            {**HOURS, "WorkingHoursTimeZoneId": "tzone://Microsoft/Custom"},
            {**HOURS, "WorkHoursStartTimeInMinutes": "540"},
        ):
            self.assertFalse(working_hours(value)["usable"])

    def test_published_hours_opt_in_blocks_unknown_and_outside_hours(self):
        options = args("find-slots", "--respect-working-hours")
        for hours in (None, {**HOURS, "WorkHoursStartTimeInMinutes": 600}):
            data = self.snapshot(options, row(hours=hours))
            self.assertEqual(find_slots(data, options)["slots"], [])
        data = self.snapshot(options, row())
        self.assertEqual(len(find_slots(data, options)["slots"]), 3)

    def test_dst_fall_back_has_distinct_offsets_and_utc_chronology(self):
        options = args(
            "find-slots",
            "--start",
            "2026-10-25T00:00:00Z",
            "--end",
            "2026-10-25T02:00:00Z",
            "--time-zone",
            "Europe/Zurich",
            "--work-start",
            "00:00",
            "--work-end",
            "05:00",
            "--weekdays",
            "all",
        )
        data = self.snapshot(options, row("0" * 8))
        result = find_slots(data, options)
        self.assertEqual(result["candidate_count"], 7)
        self.assertEqual(result["slots"][0]["start"], "2026-10-25T02:00:00+02:00")
        self.assertEqual(result["slots"][4]["start"], "2026-10-25T02:00:00+01:00")

    def test_dst_spring_gap_does_not_create_nonexistent_local_times(self):
        options = args(
            "find-slots",
            "--start",
            "2027-03-28T00:00:00Z",
            "--end",
            "2027-03-28T03:00:00Z",
            "--time-zone",
            "Europe/Zurich",
            "--work-start",
            "00:00",
            "--work-end",
            "06:00",
            "--weekdays",
            "all",
            "--max-results",
            "100",
        )
        data = self.snapshot(options, row("0" * 12))
        result = find_slots(data, options)
        self.assertEqual(result["candidate_count"], 11)
        self.assertFalse(any("T02:" in r["start"] or "T02:" in r["end"] for r in result["slots"]))

    def test_circuit_breaker_blocks_scheduling_before_browser_access(self):
        options = args("find-slots")
        with (
            tempfile.TemporaryDirectory() as state,
            patch.dict("os.environ", {"M365_HOME": state}),
            patch("m365_cli.maintainer.breakers", return_value={"calendar find-slots": "test"}),
            patch("m365_cli.cli.Session") as session,
        ):
            self.error("capability_disabled", lambda: run(options))
            session.assert_not_called()


if __name__ == "__main__":
    unittest.main()
