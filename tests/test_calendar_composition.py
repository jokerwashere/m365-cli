"""Synthetic shared event/mail composition and calendar transport contracts."""

import io
import json
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from test_api import FakeSession, response

from m365_cli.api import API
from m365_cli.cli import compatible_plans, dispatch, parser, plan, prepare, run
from m365_cli.composition import MAX_SOURCE
from m365_cli.errors import OutlookError
from m365_cli.live_calendar_composition import CalendarCompositionCanary
from m365_cli.maintainer import inventory
from m365_cli.operations import execute

CREATE = [
    "calendar",
    "create",
    "--subject",
    "Synthetic event",
    "--start",
    "2031-02-03T14:00:00+01:00",
    "--end",
    "2031-02-03T14:30:00+01:00",
]
EVENT = {
    "ItemId": {"Id": "event", "ChangeKey": "v1"},
    "Start": "2031-02-03T13:00:00Z",
    "End": "2031-02-03T13:30:00Z",
    "CalendarItemType": "Single",
    "IsOrganizer": True,
    "Body": {"BodyType": "HTML", "Value": "<p>Stored</p>"},
}


class CalendarCompositionTests(unittest.TestCase):
    def args(self, *words):
        return prepare(parser().parse_args(list(words)))

    def error(self, code, action):
        with self.assertRaises(OutlookError) as caught:
            action()
        self.assertEqual(caught.exception.code, code)
        return caught.exception

    def test_events_use_identical_mail_renderer_for_all_formats_and_styles(self):
        source = '<p onclick="x()">**Agenda** Zażółć 📅</p><script>removed</script>'
        for format in ("text", "html", "markdown"):
            for style in ([], ["--style", "minimal"], ["--style", "outlook"]):
                if format == "text" and style:
                    continue
                body = ["--body", source, "--body-format", format, *style]
                mail = self.args("mail", "draft", "create", "--subject", "x", *body)
                for route in (CREATE, ["calendar", "update", "event"]):
                    event = self.args(*route, *body)
                    self.assertEqual(
                        (event.text, event.body_type, event.composition),
                        (mail.text, mail.body_type, mail.composition),
                    )
                    api = Mock()
                    dispatch(api, event)
                    native = (
                        api.calendar_create.call_args.kwargs
                        if event.command == "create"
                        else api.update.call_args.args[1]["Body"]
                    )
                    self.assertEqual(native.get("body", native.get("Value")), mail.text)
                    self.assertEqual(
                        native.get("body_type", native.get("BodyType")), mail.body_type
                    )

    def test_calendar_invalid_input_fails_before_session(self):
        with patch("m365_cli.cli.Session") as session:
            for flags, code in (
                (["--body-format", "html"], "body_required"),
                (["--body-format", "markdown"], "body_required"),
                (["--body", "plain", "--style", "outlook"], "invalid_style"),
                (["--body", "\x00"], "invalid_body"),
                (["--body", "x" * (MAX_SOURCE + 1)], "invalid_body"),
                (["--body", "ą" * (MAX_SOURCE // 2 + 1)], "invalid_body"),
                (["--body", "<br>" * 20001, "--body-format", "html"], "body_too_complex"),
                (
                    ["--body", "x", "--body-format", "html", "--style-file", "missing.json"],
                    "invalid_style",
                ),
                (
                    [
                        "--body",
                        "x",
                        "--body-format",
                        "html",
                        "--style",
                        "minimal",
                        "--style-file",
                        "missing.json",
                    ],
                    "invalid_style",
                ),
            ):
                for route in (CREATE, ["calendar", "update", "event"]):
                    self.error(
                        code,
                        lambda route=route, flags=flags: run(parser().parse_args([*route, *flags])),
                    )
            session.assert_not_called()

    def test_calendar_file_and_stdin_share_bounded_reader(self):
        for route in (CREATE, ["calendar", "update", "event"], ["calendar", "render"]):
            with patch("m365_cli.cli.Path.open") as opened:
                stream = opened.return_value.__enter__.return_value
                stream.read.return_value = "x" * (MAX_SOURCE + 1)
                self.error(
                    "invalid_body",
                    lambda route=route: self.args(*route, "--body-file", "large.txt"),
                )
                stream.read.assert_called_once_with(MAX_SOURCE + 1)
            stream = io.StringIO("x" * (MAX_SOURCE + 20))
            with patch("sys.stdin", stream):
                self.error(
                    "invalid_body", lambda route=route: self.args(*route, "--body-file", "-")
                )
            self.assertEqual(stream.tell(), MAX_SOURCE + 1)
            with patch("sys.stdin", io.StringIO("**Agenda 📅**")):
                args = self.args(*route, "--body-file", "-", "--body-format", "markdown")
                self.assertIn("<strong>Agenda 📅</strong>", args.text)

    def test_empty_and_omitted_bodies_keep_update_semantics(self):
        omitted = self.args("calendar", "update", "event", "--subject", "new")
        self.assertNotIn("Body", omitted.changes)
        for format, flags, expected in (
            ("text", [], "Text"),
            ("html", ["--style", "minimal"], "HTML"),
        ):
            args = self.args(
                "calendar", "update", "event", "--body", "", "--body-format", format, *flags
            )
            self.assertEqual(args.changes["Body"], {"BodyType": expected, "Value": ""})
        self.error("no_changes", lambda: self.args("calendar", "update", "event"))

    def test_local_preview_styles_and_dry_run_need_no_browser(self):
        with (
            tempfile.TemporaryDirectory() as tmp,
            patch.dict(os.environ, {"M365_HOME": tmp}),
            patch("m365_cli.cli.Session") as session,
        ):
            self.assertEqual(
                run(parser().parse_args(["calendar", "styles"])),
                run(parser().parse_args(["mail", "styles"])),
            )
            output = Path(tmp) / "event.html"
            args = [
                "calendar",
                "render",
                "--body-format",
                "markdown",
                "--body",
                "**Agenda**",
                "--output",
                str(output),
            ]
            preview = run(parser().parse_args(args))
            self.assertTrue(preview["local_only"])
            self.assertFalse(preview["submitted"])
            self.assertIn("<strong>Agenda</strong>", output.read_text())
            self.assertIn("img-src 'none'", output.read_text())
            self.assertEqual(output.stat().st_mode & 0o777, 0o600)
            self.error("output_exists", lambda: run(parser().parse_args(args)))
            self.error("body_required", lambda: run(parser().parse_args(["calendar", "render"])))
            data = run(
                parser().parse_args(
                    [*CREATE, "--body-format", "markdown", "--body", "**Agenda**", "--dry-run"]
                )
            )
            self.assertEqual(data["plan"]["body_type"], "HTML")
            self.assertEqual(data["plan"]["text"], preview["body"])
            session.assert_not_called()

    def test_calendar_reads_are_explicit_and_native(self):
        for flags, type in (([], "Text"), (["--body-format", "html"], "HTML")):
            api = Mock()
            dispatch(api, self.args("calendar", "get", "event", *flags))
            api.get.assert_called_once_with("event", calendar=True, body_type=type)
        s = FakeSession([response(EVENT)])
        API(s).get("event", calendar=True, body_type="HTML")
        shape = s.calls[0][1]["ItemShape"]
        self.assertEqual(shape["BodyType"], "HTML")
        self.assertIn("calendar:Start", json.dumps(shape))
        for body in (
            None,
            {},
            {"BodyType": "Text", "Value": "x"},
            {"BodyType": "HTML", "Value": None},
        ):
            s = FakeSession([response({**EVENT, "Body": body})])
            self.error(
                "api_changed", lambda s=s: API(s).get("event", calendar=True, body_type="HTML")
            )

    def test_calendar_create_native_html_preserves_invitation_policy(self):
        for people in (None, ["user@example.com"]):
            s = FakeSession([response(EVENT)])
            API(s).calendar_create(
                subject="x",
                start=EVENT["Start"],
                end=EVENT["End"],
                body="<b>x</b>",
                body_type="HTML",
                attendees=people,
            )
            body = s.calls[0][1]
            self.assertEqual(body["Items"][0]["Body"], {"BodyType": "HTML", "Value": "<b>x</b>"})
            self.assertEqual(
                body["SendMeetingInvitations"], "SendToAllAndSaveCopy" if people else "SendToNone"
            )

    def test_calendar_update_html_readback_and_change_key(self):
        s = FakeSession([response(EVENT), response({}), response(EVENT)])
        result = API(s).update("event", {"Body": EVENT["Body"]}, calendar=True, expected="v1")
        self.assertEqual(result["Body"]["BodyType"], "HTML")
        update = s.calls[1][1]
        self.assertEqual(update["ConflictResolution"], "NeverOverwrite")
        self.assertEqual(update["ItemChanges"][0]["ItemId"]["ChangeKey"], "v1")
        self.assertEqual(update["SendCalendarInvitationsOrCancellations"], "SendToNone")
        self.assertEqual(s.calls[2][1]["ItemShape"]["BodyType"], "HTML")

    def test_calendar_html_updates_do_not_bypass_existing_guards(self):
        meeting = {
            **EVENT,
            "RequiredAttendees": [{"Mailbox": {"EmailAddress": "user@example.com"}}],
        }
        for item, expected, code in (
            (EVENT, "stale", "conflict"),
            (meeting, "v1", "notification_required"),
            ({**meeting, "IsOrganizer": False}, "v1", "not_organizer"),
            ({**EVENT, "CalendarItemType": "RecurringMaster"}, "v1", "recurring_series"),
        ):
            s = FakeSession([response(item)])
            self.error(
                code,
                lambda s=s, expected=expected: API(s).update(
                    "event", {"Body": EVENT["Body"]}, calendar=True, expected=expected
                ),
            )
            self.assertEqual(len(s.calls), 1)
        s = FakeSession([response(meeting), response({}), response(meeting)])
        API(s).update("event", {"Body": EVENT["Body"]}, calendar=True, notify=True, expected="v1")
        self.assertEqual(
            s.calls[1][1]["SendCalendarInvitationsOrCancellations"], "SendToAllAndSaveCopy"
        )

    def test_unknown_calendar_writes_and_failed_readbacks_never_retry(self):
        for reply in (OutlookError("transport_error", "unknown"), {"Body": {}}):
            s = FakeSession([reply])
            with self.assertRaises(OutlookError):
                API(s).calendar_create(
                    subject="x", start=EVENT["Start"], end=EVENT["End"], body="x", body_type="HTML"
                )
            self.assertEqual(len(s.calls), 1)
        s = FakeSession(
            [
                response(EVENT),
                response({}),
                response({**EVENT, "Body": {"BodyType": "Text", "Value": "x"}}),
            ]
        )
        error = self.error(
            "readback_failed",
            lambda: API(s).update("event", {"Body": EVENT["Body"]}, calendar=True),
        )
        self.assertTrue(error.as_dict()["write_acknowledged"])
        self.assertEqual([c[0] for c in s.calls], ["GetItem", "UpdateItem", "GetItem"])

    def test_plain_calendar_replays_both_prior_fingerprint_shapes(self):
        for words in (
            CREATE + ["--body", "literal"],
            ["calendar", "update", "event", "--subject", "new"],
        ):
            args = self.args(*words)
            current = plan(args)
            legacy = {
                k: v
                for k, v in vars(args).items()
                if k
                not in {
                    "request_id",
                    "dry_run",
                    "mutation",
                    "body_file",
                    "changes",
                    "body_composition",
                    "body_format",
                    "body_type",
                    "composition",
                    "style",
                    "style_file",
                    "save_only",  # This option did not exist in either historical parser.
                }
            }
            self.assertEqual(current, legacy)
            for prior in (legacy, {**legacy, "body_format": "text", "body_type": "Text"}):
                with (
                    tempfile.TemporaryDirectory() as tmp,
                    patch.dict(os.environ, {"M365_HOME": tmp}),
                ):
                    action = Mock(return_value={"id": "event"})
                    execute("old", "user@example.com", prior, action)
                    result = execute(
                        "old",
                        "user@example.com",
                        current,
                        action,
                        compatible_payloads=compatible_plans(args),
                    )
                    self.assertTrue(result["replayed"])
                    action.assert_called_once()
                    self.error(
                        "request_id_conflict",
                        lambda current=current, action=action, args=args: execute(
                            "old",
                            "other@example.com",
                            current,
                            action,
                            compatible_payloads=compatible_plans(args),
                        ),
                    )
                    changed = self.args(*words, "--location", "different")
                    self.error(
                        "request_id_conflict",
                        lambda changed=changed, action=action: execute(
                            "old",
                            "user@example.com",
                            plan(changed),
                            action,
                            compatible_payloads=compatible_plans(changed),
                        ),
                    )

    def test_legacy_pending_calendar_write_stays_unknown(self):
        with tempfile.TemporaryDirectory() as tmp, patch.dict(os.environ, {"M365_HOME": tmp}):
            args = self.args(*CREATE)
            action = Mock(side_effect=OutlookError("transport_error", "unknown"))
            self.error(
                "transport_error",
                lambda: execute("pending", "user@example.com", compatible_plans(args)[0], action),
            )
            self.error(
                "write_outcome_unknown",
                lambda: execute(
                    "pending",
                    "user@example.com",
                    plan(args),
                    action,
                    compatible_payloads=compatible_plans(args),
                ),
            )
            action.assert_called_once()

    def test_calendar_style_content_binds_operation_and_no_rich_legacy_aliases(self):
        with tempfile.TemporaryDirectory() as tmp, patch.dict(os.environ, {"M365_HOME": tmp}):
            path = Path(tmp) / "style.json"
            path.write_text('{"font_size_pt":11}')
            words = [
                *CREATE,
                "--body",
                "**x**",
                "--body-format",
                "markdown",
                "--style-file",
                str(path),
            ]
            first = self.args(*words)
            self.assertEqual(compatible_plans(first), [])
            action = Mock(return_value={"id": "event"})
            execute("styled", "user@example.com", plan(first), action)
            self.assertTrue(execute("styled", "user@example.com", plan(first), action)["replayed"])
            path.write_text('{"font_size_pt":12}')
            self.error(
                "request_id_conflict",
                lambda: execute("styled", "user@example.com", plan(self.args(*words)), action),
            )
            action.assert_called_once()

    def test_calendar_has_dedicated_composition_validation(self):
        caps = {c["command"]: c for c in inventory()["capabilities"]}
        for verb in ("get", "create", "update"):
            self.assertEqual(caps["calendar " + verb]["validation_suite"], "outlook")
            self.assertIn(
                "calendar-composition", caps["calendar " + verb]["additional_validation_suites"]
            )
        for verb in ("render", "styles"):
            self.assertFalse(caps["calendar " + verb]["mutation"])
            self.assertEqual(caps["calendar " + verb]["validation_suite"], "calendar-composition")
        self.assertNotIn(
            "calendar-composition", caps["calendar respond"]["additional_validation_suites"]
        )

    def test_canary_waits_for_settled_fixture_then_writes_once(self):
        canary = CalendarCompositionCanary.__new__(CalendarCompositionCanary)
        canary.invoke = Mock(
            side_effect=[{"change_key": key} for key in ("v1", "v2", "v2", "v2")]
            + [{"updated": True}]
        )
        with patch("m365_cli.live_calendar_composition.time.sleep"):
            result = canary.update({"id": "fixture"}, "--body", "updated")
        self.assertEqual(result, {"updated": True})
        self.assertEqual(canary.invoke.call_count, 5)
        canary.invoke.assert_called_with(
            None, "calendar", "update", "fixture", "--body", "updated", "--if-change-key", "v2"
        )

    def test_canary_never_retries_rejected_write(self):
        canary = CalendarCompositionCanary.__new__(CalendarCompositionCanary)
        canary.invoke = Mock(
            side_effect=[
                {"change_key": "v1"},
                {"change_key": "v1"},
                {"change_key": "v1"},
                OutlookError("conflict", "Changed again"),
            ]
        )
        with patch("m365_cli.live_calendar_composition.time.sleep"):
            self.error("conflict", lambda: canary.update({"id": "fixture"}, "--body", "x"))
        self.assertEqual(canary.invoke.call_count, 4)
        self.assertEqual(sum(c.args[2] == "update" for c in canary.invoke.call_args_list), 1)

    def test_canary_refuses_unstable_or_missing_fixture_versions(self):
        canary = CalendarCompositionCanary.__new__(CalendarCompositionCanary)
        canary.invoke = Mock(side_effect=[{"change_key": str(i)} for i in range(8)])
        with patch("m365_cli.live_calendar_composition.time.sleep"):
            self.error("conflict", lambda: canary.update({"id": "fixture"}, "--body", "x"))
        self.assertTrue(all(c.args[2] == "get" for c in canary.invoke.call_args_list))
        canary.invoke = Mock(return_value={})
        self.error("missing_change_key", lambda: canary.update({"id": "fixture"}, "--body", "x"))
        canary.invoke.assert_called_once()
