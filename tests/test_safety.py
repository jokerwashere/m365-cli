import contextlib
import io
import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from test_api import FakeSession, response
from test_transport import ready_session, reply

from m365_cli.api import API, acknowledged, check_response, iso, recipients, reference
from m365_cli.cli import main, parser, prepare, run
from m365_cli.errors import OutlookError
from m365_cli.operations import execute
from m365_cli.session import Session, command_lock, read_config, save_config, state_dir
from m365_cli.validation import email, validate_range


class Isolated(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.env = patch.dict(os.environ, {"M365_HOME": self.temp.name})
        self.env.start()
        self.addCleanup(self.env.stop)
        self.addCleanup(self.temp.cleanup)

    def error(self, code, call):
        with self.assertRaises(OutlookError) as exc:
            call()
        self.assertEqual(exc.exception.code, code)


class ValidationTests(Isolated):
    def test_malformed_addresses(self):
        for value in (
            "",
            "@a",
            "a@",
            "a@@b",
            "A <a@b.com>",
            "a b@c.com",
            "a@b.com,c@d.com",
            "a@-b.com",
            "a@b..com",
            "a\n@b.com",
            "a\x00@b.com",
            "a@b;bad",
            "a@b:80",
            ".a@b.com",
            "a..b@c.com",
            "a@bad_domain.com",
        ):
            with self.subTest(value=value):
                self.error("invalid_recipient", lambda value=value: email(value))
        for value in ("sender+test@example.com", "a_b@example-domain.test", "a@localhost"):
            self.assertEqual(email(value), value)

    def test_dst_and_exclusive_all_day_ranges(self):
        validate_range("2026-10-25T00:00:00+02:00", "2026-10-26T00:00:00+01:00", all_day=True)
        self.error(
            "invalid_range",
            lambda: validate_range(
                "2026-10-25T10:00:00+02:00", "2026-10-26T00:00:00+01:00", all_day=True
            ),
        )
        self.error(
            "invalid_range",
            lambda: validate_range("2026-01-01T00:00:00Z", "2028-01-01T00:00:00Z", window=True),
        )
        self.error("timezone_required", lambda: iso("2026-01-01"))

    def test_bad_json_envelopes_and_http_errors(self):
        for value in (
            None,
            [],
            {},
            {"status": 200, "body": None},
            {"status": 200, "body": {}},
            {"status": 200, "body": []},
        ):
            with self.subTest(value=value):
                with self.assertRaises(OutlookError):
                    check_response(value)
        for status in (401, 403):
            self.error("session_rejected", lambda status=status: check_response({"status": status}))
        self.error("rate_limited", lambda: check_response({"status": 429, "body": {}}))
        self.error(
            "owa_error", lambda: check_response({"status": 200, "body": {"ResponseClass": "Error"}})
        )
        self.error(
            "conflict",
            lambda: check_response(
                {"status": 200, "body": {"ResponseCode": "ErrorIrresolvableConflict"}}
            ),
        )
        self.error(
            "item_not_found",
            lambda: check_response({"status": 200, "body": {"ResponseCode": "ErrorItemNotFound"}}),
        )

    def test_acknowledgement_is_required(self):
        self.error("write_outcome_unknown", lambda: acknowledged({"Body": {}}))
        self.error("missing_change_key", lambda: reference({"id": "x"}))
        self.error("conflict", lambda: reference({"id": "x", "change_key": "new"}, "old"))

    def test_invalid_cli_input_never_connects(self):
        cases = [
            (["mail", "send", "--subject", "x"], "recipient_required"),
            (["mail", "list", "--limit", "0"], "invalid_paging"),
            (["mail", "list", "--offset", "-1"], "invalid_paging"),
            (["mail", "draft", "update", "x"], "no_changes"),
            (["calendar", "update", "x", "--start", "bad"], "invalid_datetime"),
            (
                ["mail", "draft", "create", "--subject", "x", "--body-file", "/does/not/exist"],
                "input_file_error",
            ),
            (
                ["mail", "draft", "create", "--subject", "x", "--to", "Name <a@example.com>"],
                "invalid_recipient",
            ),
        ]
        with patch("m365_cli.cli.Session") as session:
            for argv, code in cases:
                with self.subTest(argv=argv):
                    self.error(code, lambda argv=argv: run(parser().parse_args(argv)))
            session.assert_not_called()

    def test_dry_run_does_not_connect_or_create_ledger(self):
        with patch("m365_cli.cli.Session") as session:
            args = parser().parse_args(
                [
                    "mail",
                    "send",
                    "--subject",
                    "x",
                    "--to",
                    "a@example.com",
                    "--request-id",
                    "safe",
                    "--dry-run",
                ]
            )
            self.assertTrue(run(args)["dry_run"])
            session.assert_not_called()
        self.assertFalse((state_dir() / "operations").exists())

    def test_clear_fields_are_explicit_and_mutually_exclusive(self):
        args = prepare(
            parser().parse_args(
                ["mail", "draft", "update", "id", "--subject", "", "--body", "", "--clear-to"]
            )
        )
        self.assertEqual(
            args.changes,
            {"Subject": "", "Body": {"BodyType": "Text", "Value": ""}, "ToRecipients": []},
        )
        self.error(
            "usage_error",
            lambda: parser().parse_args(
                ["mail", "draft", "update", "id", "--to", "a@b.com", "--clear-to"]
            ),
        )

    def test_stdin_body_and_unicode_file(self):
        with patch("sys.stdin", io.StringIO("Zażółć 📨")):
            args = prepare(
                parser().parse_args(
                    ["mail", "send", "--subject", "x", "--to", "a@b.com", "--body-file", "-"]
                )
            )
        self.assertEqual(args.text, "Zażółć 📨")
        path = Path(self.temp.name) / "body"
        path.write_bytes(b"\xff")
        self.error(
            "input_file_error",
            lambda: prepare(
                parser().parse_args(
                    ["mail", "send", "--subject", "x", "--to", "a@b.com", "--body-file", str(path)]
                )
            ),
        )

    def test_invalid_arguments_are_json_exit_two(self):
        with (
            patch.object(sys, "argv", ["m365", "mail", "nonexistent"]),
            contextlib.redirect_stdout(io.StringIO()) as out,
        ):
            with self.assertRaises(SystemExit) as exc:
                main()
        self.assertEqual(exc.exception.code, 2)
        self.assertEqual(json.loads(out.getvalue())["error"]["code"], "usage_error")

    def test_all_leaf_help_commands(self):
        def visit(p, prefix):
            actions = [
                a for a in p._actions if isinstance(a, __import__("argparse")._SubParsersAction)
            ]
            if not actions:
                with (
                    contextlib.redirect_stdout(io.StringIO()),
                    self.assertRaises(SystemExit) as exc,
                ):
                    parser().parse_args(prefix + ["--help"])
                self.assertEqual(exc.exception.code, 0)
            else:
                for name, child in actions[0].choices.items():
                    visit(child, prefix + [name])

        visit(parser(), [])


class OperationTests(Isolated):
    def test_identical_retry_replays_without_second_send(self):
        action = Mock(return_value={"submitted": True})
        a = execute("send-123", "a@example.com", {"subject": "x"}, action)
        b = execute("send-123", "a@example.com", {"subject": "x"}, action)
        self.assertTrue(a["submitted"])
        self.assertTrue(b["replayed"])
        action.assert_called_once()

    def test_same_id_different_input_or_account_is_rejected(self):
        execute("x", "a@example.com", {"body": "one"}, lambda: {"submitted": True})
        for account, payload in [
            ("b@example.com", {"body": "one"}),
            ("a@example.com", {"body": "two"}),
        ]:
            self.error(
                "request_id_conflict",
                lambda account=account, payload=payload: execute(
                    "x", account, payload, lambda: self.fail("write called")
                ),
            )

    def test_interrupted_write_stays_pending_and_blocks_retry(self):
        action = Mock(side_effect=KeyboardInterrupt)
        with self.assertRaises(KeyboardInterrupt):
            execute("interrupted", "a@b", {}, action)
        self.error("write_outcome_unknown", lambda: execute("interrupted", "a@b", {}, action))
        self.assertEqual(action.call_count, 1)

    def test_result_persistence_failure_blocks_retry(self):
        from m365_cli.operations import atomic_json

        calls = 0

        def write(path, data):
            nonlocal calls
            calls += 1
            if calls == 2:
                raise OSError("disk full")
            atomic_json(path, data)

        action = Mock(return_value={"submitted": True})
        with patch("m365_cli.operations.atomic_json", write), self.assertRaises(OSError):
            execute("disk", "a@b", {}, action)
        self.error("write_outcome_unknown", lambda: execute("disk", "a@b", {}, action))
        action.assert_called_once()

    def test_file_permissions_and_configuration_corruption(self):
        save_config({"account": "a@example.com"})
        self.assertEqual(read_config()["account"], "a@example.com")
        self.assertEqual(state_dir().stat().st_mode & 0o777, 0o700)
        self.assertEqual((state_dir() / "config.json").stat().st_mode & 0o777, 0o600)
        (state_dir() / "config.json").write_text("[]")
        self.error("config_invalid", read_config)

    def test_command_lock_prevents_concurrent_process(self):
        with command_lock():
            result = subprocess.run(
                [sys.executable, "-m", "m365_cli.cli", "auth", "status"],
                capture_output=True,
                text=True,
            )
        self.assertEqual(json.loads(result.stdout)["error"]["code"], "busy")


class MutationSafetyTests(Isolated):
    def test_stale_update_never_writes(self):
        session = FakeSession(
            [response({"ItemId": {"Id": "x", "ChangeKey": "new"}, "IsDraft": True})]
        )
        self.error(
            "conflict", lambda: API(session).update("x", {"Subject": "change"}, expected="old")
        )
        self.assertEqual(len(session.calls), 1)

    def test_metadata_update_cannot_change_message_body(self):
        session = FakeSession([])
        self.error("invalid_changes", lambda: API(session).update("x", {"Body": {}}, metadata=True))
        self.assertFalse(session.calls)

    def test_wrong_item_type_never_writes(self):
        session = FakeSession(
            [
                response(
                    {
                        "ItemId": {"Id": "x", "ChangeKey": "ck"},
                        "ItemClass": "IPM.Appointment",
                        "IsDraft": True,
                    }
                )
            ]
        )
        self.error("wrong_item_type", lambda: API(session).draft_send("x"))
        self.assertEqual(len(session.calls), 1)

    def test_nonorganizer_cannot_update_or_cancel(self):
        item = {
            "ItemId": {"Id": "x", "ChangeKey": "ck"},
            "Start": "2026-01-01T09:00:00Z",
            "End": "2026-01-01T10:00:00Z",
            "IsOrganizer": False,
            "RequiredAttendees": [{"Mailbox": {"EmailAddress": "a@b"}}],
        }
        for action in ("update", "delete"):
            session = FakeSession([response(item)])

            def call(session=session, action=action):
                if action == "update":
                    return API(session).update(
                        "x", {"Subject": "change"}, calendar=True, notify=True
                    )
                return API(session).delete("x", calendar=True, notify=True)

            self.error("not_organizer", call)
            self.assertEqual(len(session.calls), 1)

    def test_organizer_cannot_respond_to_own_meeting(self):
        session = FakeSession([response({"ItemId": {"Id": "x"}, "IsOrganizer": True})])
        self.error(
            "is_organizer", lambda: API(session).calendar_respond("x", "accept", notify=True)
        )
        self.assertEqual(len(session.calls), 1)

    def test_meeting_response_uses_change_key_and_notification_setting(self):
        item = {
            "ItemId": {"Id": "x", "ChangeKey": "ck"},
            "RequiredAttendees": [{}],
            "IsOrganizer": False,
        }
        for notify in (True, False):
            s = FakeSession([response(item), response({})])
            API(s).calendar_respond("x", "tentative", notify=notify)
            request = s.calls[1][1]
            self.assertEqual(request["Items"][0]["ReferenceItemId"]["ChangeKey"], "ck")
            self.assertEqual(
                request["MessageDisposition"], "SendAndSaveCopy" if notify else "SaveOnly"
            )

    def test_already_deleted_never_hard_deletes(self):
        session = FakeSession(
            [
                response(
                    {"ItemId": {"Id": "x", "ChangeKey": "ck"}, "ParentFolderId": {"Id": "deleted"}}
                ),
                {"FolderId": {"Id": "deleted", "ChangeKey": "ck"}},
            ]
        )
        self.error("already_deleted", lambda: API(session).delete("x"))
        self.assertNotIn("DeleteItem", [c[0] for c in session.calls])

    def test_attachment_must_belong_to_draft(self):
        session = FakeSession(
            [
                response(
                    {"ItemId": {"Id": "x", "ChangeKey": "ck"}, "IsDraft": True, "Attachments": []}
                )
            ]
        )
        self.error(
            "attachment_not_found",
            lambda: API(session).attachment_delete("x", "someone-elses-attachment"),
        )
        self.assertEqual(len(session.calls), 1)

    def test_attachment_add_size_and_draft_guards(self):
        self.error(
            "attachment_too_large",
            lambda: API(FakeSession([])).attachment_add("x", "file", b"x" * (10 * 1024 * 1024 + 1)),
        )
        session = FakeSession(
            [response({"ItemId": {"Id": "x", "ChangeKey": "ck"}, "IsDraft": False})]
        )
        self.error("not_a_draft", lambda: API(session).attachment_add("x", "file", b"ok"))

    def test_send_draft_preserves_existing_item_and_change_key(self):
        session = FakeSession(
            [
                response(
                    {
                        "ItemId": {"Id": "draft", "ChangeKey": "ck"},
                        "IsDraft": True,
                        "ToRecipients": recipients(["a@example.com"]),
                    }
                ),
                response({}),
            ]
        )
        self.assertTrue(API(session).draft_send("draft")["submitted"])
        self.assertEqual(session.calls[1][0], "UpdateItem")
        self.assertEqual(session.calls[1][1]["ItemChanges"][0]["ItemId"]["ChangeKey"], "ck")

    def test_postwrite_read_failure_is_not_reported_as_unsent(self):
        session = FakeSession(
            [
                response({"ItemId": {"Id": "x", "ChangeKey": "ck"}, "IsDraft": True}),
                response({}),
                OutlookError("transport_error", "failed"),
            ]
        )
        self.error("readback_failed", lambda: API(session).update("x", {"Subject": "updated"}))

    def test_server_identity_is_used_instead_of_cached_claim(self):
        session = Session()
        session.context = Mock()
        session.page = Mock()
        session.context.pages = []
        session.status = Mock(
            return_value={"authenticated": True, "account": "expected@example.com"}
        )
        session._bootstrap = Mock()
        session.request = Mock(return_value={"UserEmailAddress": "different@example.com"})
        self.error(
            "account_mismatch", lambda: session.authenticate(expected="expected@example.com")
        )
        self.assertEqual(session.account, "")

    def test_transport_errors_never_leak_secrets_and_mark_unknown_writes(self):
        for action, code in [
            ("GetItem", "transport_error"),
            ("CreateItem", "write_outcome_unknown"),
        ]:
            session = ready_session("a@example.com")
            session.http.request.side_effect = (
                [reply({"UserEmailAddress": "a@example.com"}), RuntimeError("Bearer SECRET")]
                if action == "CreateItem"
                else RuntimeError("Bearer SECRET")
            )
            with self.assertRaises(OutlookError) as exc:
                session.request(action, {})
            self.assertEqual(exc.exception.code, code)
            self.assertNotIn("SECRET", str(exc.exception))


if __name__ == "__main__":
    unittest.main()
