"""Synthetic never-sent meeting lifecycle, delivery guards and request compatibility."""

import copy
import tempfile
import unittest
from unittest.mock import Mock, patch

from test_api import FakeSession, response

from m365_cli.api import API, attendee_list
from m365_cli.cli import compatible_plans, dispatch, parser, plan, prepare, run
from m365_cli.errors import OutlookError
from m365_cli.live_calendar_drafts import CalendarDraftCanary
from m365_cli.maintainer import inventory
from m365_cli.operations import execute

EVENT = {
    "ItemId": {"Id": "event", "ChangeKey": "v1"},
    "ParentFolderId": {"Id": "calendar"},
    "ItemClass": "IPM.Appointment",
    "Subject": "Synthetic meeting",
    "Start": "2031-02-17T13:00:00Z",
    "End": "2031-02-17T13:30:00Z",
    "IsDraft": False,
    "IsOrganizer": True,
    "IsMeeting": True,
    "IsCancelled": False,
    "MeetingRequestWasSent": False,
    "CalendarItemType": "Single",
    "RequiredAttendees": attendee_list(["reader@example.com"]),
    "OptionalAttendees": [],
}
DELETED = {"FolderId": {"Id": "deleted", "ChangeKey": "d1"}}
CREATE = [
    "calendar",
    "create",
    "--subject",
    "Synthetic meeting",
    "--start",
    EVENT["Start"],
    "--end",
    EVENT["End"],
]


class CalendarDraftTests(unittest.TestCase):
    def error(self, code, fn):
        with self.assertRaises(OutlookError) as exc:
            fn()
        self.assertEqual(exc.exception.code, code)
        return exc.exception

    def args(self, *words):
        return prepare(parser().parse_args(list(words)))

    def test_get_requests_authoritative_meeting_state_without_using_isdraft(self):
        s = FakeSession([response(EVENT)])
        item = API(s).get("event", calendar=True)
        self.assertFalse(item["IsDraft"])
        self.assertFalse(item["MeetingRequestWasSent"])
        fields = str(s.calls[0][1]["ItemShape"]["AdditionalProperties"])
        for name in ("MeetingRequestWasSent", "IsMeeting", "Resources"):
            self.assertIn("calendar:" + name, fields)
        for bad in (None, 0, "false"):
            s = FakeSession([response({**EVENT, "MeetingRequestWasSent": bad})])
            self.error("api_changed", lambda s=s: API(s).get("event", calendar=True))

    def test_creation_keeps_legacy_send_default_and_can_save_attendees(self):
        for save_only in (False, True):
            replies = [response({"ItemId": {"Id": "event", "ChangeKey": "v1"}})]
            if save_only:
                replies.append(response(EVENT))
            s = FakeSession(replies)
            result = API(s).calendar_create(
                subject=EVENT["Subject"],
                start=EVENT["Start"],
                end=EVENT["End"],
                attendees=["reader@example.com"],
                optional=["observer@example.com"],
                save_only=save_only,
            )
            body = s.calls[0][1]
            self.assertEqual(
                body["SendMeetingInvitations"],
                "SendToNone" if save_only else "SendToAllAndSaveCopy",
            )
            self.assertEqual(len(body["Items"][0]["RequiredAttendees"]), 1)
            self.assertEqual(len(body["Items"][0]["OptionalAttendees"]), 1)
            self.assertEqual(result["notification_submitted"], not save_only)
            if save_only:
                self.assertTrue(result["saved_only"])
                self.assertFalse(result["MeetingRequestWasSent"])

    def test_save_only_creation_missing_ack_or_unverified_state_never_retries(self):
        kw = dict(subject="Synthetic", start=EVENT["Start"], end=EVENT["End"], save_only=True)
        s = FakeSession([{}])
        self.error("write_outcome_unknown", lambda s=s: API(s).calendar_create(**kw))
        self.assertEqual(len(s.calls), 1)
        missing = {k: v for k, v in EVENT.items() if k != "MeetingRequestWasSent"}
        s = FakeSession([response(EVENT), response(missing)])
        error = self.error("readback_failed", lambda s=s: API(s).calendar_create(**kw))
        self.assertTrue(error.details["write_acknowledged"])
        self.assertEqual(error.details["id"], "event")
        self.assertEqual(len(s.calls), 2)

    def test_add_attendees_to_appointment_and_edit_rich_body_without_sending(self):
        appointment = {**EVENT, "IsMeeting": False, "RequiredAttendees": []}
        html = {"BodyType": "HTML", "Value": "<p>Prepared agenda</p>"}
        s = FakeSession(
            [
                response(appointment),
                response(DELETED),
                response({}),
                response({**EVENT, "Body": html}),
            ]
        )
        result = API(s).update(
            "event",
            {"RequiredAttendees": EVENT["RequiredAttendees"], "Body": html},
            calendar=True,
            save_only=True,
            expected="v1",
        )
        update = s.calls[2][1]
        self.assertEqual(update["SendCalendarInvitationsOrCancellations"], "SendToNone")
        self.assertEqual(update["ConflictResolution"], "NeverOverwrite")
        self.assertEqual(update["ItemChanges"][0]["ItemId"]["ChangeKey"], "v1")
        self.assertTrue(result["saved_only"])
        self.assertFalse(result["notification_submitted"])
        self.assertEqual(result["Body"], html)

    def test_save_only_cannot_suppress_updates_to_sent_or_unverified_meetings(self):
        cases = [
            ({"MeetingRequestWasSent": True, "IsDraft": True}, "meeting_already_sent"),
            ({"IsOrganizer": False}, "not_organizer"),
            ({"IsCancelled": True}, "cancelled_meeting"),
            ({"CalendarItemType": "Occurrence"}, "recurring_series"),
            ({"CalendarItemType": "RecurringMaster"}, "recurring_series"),
        ]
        for delta, code in cases:
            for operation in ("edit", "send"):
                s = FakeSession([response({**EVENT, **delta})])
                a = API(s)
                self.error(
                    code,
                    lambda a=a, operation=operation: (
                        a.calendar_send("event")
                        if operation == "send"
                        else a.update("event", {"Subject": "Edit"}, calendar=True, save_only=True)
                    ),
                )
                self.assertEqual(len(s.calls), 1)
        missing = {k: v for k, v in EVENT.items() if k != "MeetingRequestWasSent"}
        s = FakeSession([response(missing)])
        self.error("meeting_state_unknown", lambda s=s: API(s).calendar_send("event"))
        s = FakeSession([])
        self.error(
            "invalid_argument",
            lambda s=s: API(s).update(
                "event", {"Subject": "Edit"}, calendar=True, save_only=True, notify=True
            ),
        )
        self.assertFalse(s.calls)

    def test_save_only_and_initial_send_reject_deleted_events(self):
        for operation in ("edit", "send"):
            s = FakeSession(
                [response({**EVENT, "ParentFolderId": {"Id": "deleted"}}), response(DELETED)]
            )
            a = API(s)
            self.error(
                "already_deleted",
                lambda a=a, operation=operation: (
                    a.calendar_send("event")
                    if operation == "send"
                    else a.update("event", {"Subject": "Edit"}, calendar=True, save_only=True)
                ),
            )
            self.assertEqual(len(s.calls), 2)

    def test_explicit_send_keeps_identity_and_uses_reviewed_change_key(self):
        sent = {**EVENT, "MeetingRequestWasSent": True}
        s = FakeSession([response(EVENT), response(DELETED), response({}), response(sent)])
        result = API(s).calendar_send("event", expected="v1")
        update = s.calls[2][1]
        self.assertEqual(update["SendCalendarInvitationsOrCancellations"], "SendToAllAndSaveCopy")
        self.assertEqual(update["ItemChanges"][0]["ItemId"]["ChangeKey"], "v1")
        self.assertEqual(
            update["ItemChanges"][0]["Updates"][0]["Item"]["Subject"], EVENT["Subject"]
        )
        self.assertTrue(result["submitted"])
        self.assertEqual(result["id"], "event")
        s = FakeSession([response(sent)])
        self.error("meeting_already_sent", lambda s=s: API(s).calendar_send("event"))
        self.assertEqual(len(s.calls), 1)

    def test_send_preconditions_and_no_attendees_fail_before_write(self):
        s = FakeSession([response(EVENT)])
        self.error("conflict", lambda s=s: API(s).calendar_send("event", expected="stale"))
        self.assertEqual(len(s.calls), 1)
        s = FakeSession([response({**EVENT, "RequiredAttendees": [], "IsMeeting": False})])
        self.error("attendee_required", lambda s=s: API(s).calendar_send("event"))
        self.assertEqual(len(s.calls), 1)
        s = FakeSession([response(EVENT), response(DELETED)])
        self.error(
            "conflict",
            lambda s=s: API(s).update(
                "event", {"Subject": "Edit"}, calendar=True, save_only=True, expected="stale"
            ),
        )
        self.assertEqual(len(s.calls), 2)

    def test_send_unknown_ack_and_failed_readback_do_not_replay(self):
        for failure, code in (
            ({}, "write_outcome_unknown"),
            (OutlookError("write_outcome_unknown", "Network interrupted"), "write_outcome_unknown"),
        ):
            s = FakeSession([response(EVENT), response(DELETED), failure])
            self.error(code, lambda s=s: API(s).calendar_send("event"))
            self.assertEqual(len(s.calls), 3)
        for failure in (OutlookError("session_rejected", "Expired"), response(EVENT)):
            s = FakeSession([response(EVENT), response(DELETED), response({}), failure])
            error = self.error("readback_failed", lambda s=s: API(s).calendar_send("event"))
            self.assertTrue(error.details["write_acknowledged"])
            self.assertEqual(error.details["id"], "event")
            self.assertEqual(len(s.calls), 4)

    def test_silent_delete_requires_never_sent_state_and_sent_guard_survives_empty_roster(self):
        for notify in (True, False):
            s = FakeSession([response(EVENT), response(DELETED), response({})])
            result = API(s).delete("event", calendar=True, notify=notify)
            self.assertFalse(result["notification_submitted"])
            self.assertEqual(s.calls[2][1]["SendMeetingCancellations"], "SendToNone")
        for state in (True, None):
            item = {**EVENT, "RequiredAttendees": [], "MeetingRequestWasSent": state}
            if state is None:
                item.pop("MeetingRequestWasSent")
            s = FakeSession([response(item)])
            self.error("notification_required", lambda s=s: API(s).delete("event", calendar=True))
            self.assertEqual(len(s.calls), 1)

    def test_cli_dispatch_dry_run_and_delivery_choices(self):
        api = Mock()
        dispatch(api, self.args(*CREATE, "--save-only", "--attendee", "reader@example.com"))
        self.assertTrue(api.calendar_create.call_args.kwargs["save_only"])
        dispatch(
            api, self.args("calendar", "update", "event", "--save-only", "--subject", "Edited")
        )
        self.assertTrue(api.update.call_args.kwargs["save_only"])
        dispatch(api, self.args("calendar", "send", "event", "--if-change-key", "v1"))
        api.calendar_send.assert_called_once_with("event", expected="v1")
        with patch("m365_cli.cli.Session") as session:
            result = run(parser().parse_args(["calendar", "send", "event", "--dry-run"]))
            self.assertTrue(result["dry_run"])
            session.assert_not_called()
        self.error(
            "usage_error",
            lambda: parser().parse_args(
                ["calendar", "update", "event", "--subject", "Edit", "--notify", "--save-only"]
            ),
        )

    def test_save_only_fingerprint_is_explicit_and_old_records_still_replay(self):
        for words in (CREATE, ["calendar", "update", "event", "--subject", "Edit"]):
            args = self.args(*words)
            original = copy.copy(args)
            del original.save_only
            self.assertEqual(plan(original), plan(args))
            self.assertEqual(compatible_plans(original), compatible_plans(args))
            draft = self.args(*words, "--save-only")
            self.assertTrue(plan(draft)["save_only"])
            with tempfile.TemporaryDirectory() as tmp, patch.dict("os.environ", {"M365_HOME": tmp}):
                action = Mock(return_value={"id": "event"})
                execute("same", "reader@example.com", plan(original), action)
                self.assertTrue(
                    execute("same", "reader@example.com", plan(args), action)["replayed"]
                )
                self.assertEqual(action.call_count, 1)
                self.error(
                    "request_id_conflict",
                    lambda draft=draft, action=action: execute(
                        "same", "reader@example.com", plan(draft), action
                    ),
                )
                self.error(
                    "transport_error",
                    lambda draft=draft, action=action: execute(
                        "pending",
                        "reader@example.com",
                        plan(draft),
                        lambda: (_ for _ in ()).throw(
                            OutlookError("transport_error", "Interrupted")
                        ),
                    ),
                )
                self.error(
                    "write_outcome_unknown",
                    lambda draft=draft, action=action: execute(
                        "pending", "reader@example.com", plan(draft), action
                    ),
                )
                self.assertEqual(action.call_count, 1)

    def test_maintainer_inventory_includes_explicit_initial_send(self):
        row = next(r for r in inventory()["capabilities"] if r["command"] == "calendar send")
        self.assertTrue(row["mutation"])
        self.assertEqual(row["validation_suite"], "calendar-drafts")

    def test_live_canary_never_retries_failed_cleanup_when_collecting_cancellations(self):
        canary = object.__new__(CalendarDraftCanary)
        canary.marker = "Synthetic calendar validation"
        canary.journal = {"cleanup": []}
        canary.save = Mock()
        discovered = iter(
            [[], [{"id": "cancellation", "Subject": canary.marker + " cancellation"}]]
        )

        def invoke(api, *words):
            command = words[:2]
            if command == ("calendar", "create"):
                return {"id": "event"}
            if command == ("calendar", "get"):
                raise OutlookError("api_changed", "Synthetic read failure")
            if command == ("mail", "list"):
                return {"items": next(discovered)}
            if command == ("calendar", "delete"):
                raise OutlookError("write_outcome_unknown", "Synthetic cleanup interruption")
            if command == ("mail", "delete"):
                return {"id": words[2], "moved_to_deleted_items": True}
            self.fail("Unexpected canary command")

        canary.invoke = Mock(side_effect=invoke)
        self.error("api_changed", lambda: canary.scenario("reader@example.com"))
        commands = [call.args[1:3] for call in canary.invoke.call_args_list]
        self.assertEqual(commands.count(("calendar", "delete")), 1)
        self.assertEqual(commands.count(("mail", "delete")), 1)
        self.assertEqual(canary.journal["cleanup"][0]["state"], "pending")
        self.assertEqual(canary.journal["cleanup"][0]["code"], "write_outcome_unknown")
        self.assertEqual(canary.journal["cleanup"][1]["state"], "cleaned")
