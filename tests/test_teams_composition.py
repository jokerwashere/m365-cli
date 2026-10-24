"""Synthetic Teams rich-body contracts, routing, normalization and write safety."""

import copy
import io
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from test_teams import CHANNEL, MID, THREAD, FakeSession, message, thread

from m365_cli.cli import parser, plan, prepare, run
from m365_cli.composition import MAX_SOURCE, render
from m365_cli.errors import OutlookError
from m365_cli.maintainer import inventory
from m365_cli.operations import execute
from m365_cli.teams_api import TeamsAPI
from m365_cli.teams_cli import dispatch
from m365_cli.teams_composition import MAX_MESSAGE, content, matches, text_body


class TeamsCompositionTests(unittest.TestCase):
    def args(self, *words):
        return prepare(parser().parse_args(["teams", "message", *words]))

    def error(self, code, fn):
        with self.assertRaises(OutlookError) as exc:
            fn()
        self.assertEqual(exc.exception.code, code)
        return exc.exception

    def test_shared_renderer_styles_and_local_preview(self):
        source = "# Heading\n\n**Bold** *italic* ~~strike~~ [Link](https://example.com)\n\n- First\n- Second\n\n| A | B |\n|---|---|\n|1|2|\n\n```\nprint('test')\n```"
        expected = render(source, "markdown")
        with patch("m365_cli.teams_cli.TeamsSession") as session:
            result = run(
                parser().parse_args(
                    ["teams", "message", "render", "--body-format", "markdown", "--body", source]
                )
            )
            styles = run(parser().parse_args(["teams", "message", "styles"]))
            session.assert_not_called()
        self.assertEqual(result["body"], expected["body"])
        self.assertEqual(result["message_content"], expected["body"])
        self.assertTrue(result["local_only"])
        self.assertFalse(result["submitted"])
        self.assertIn("default-src 'none'", result["preview_html"])
        self.assertEqual(styles["default"], "outlook")
        self.assertEqual(styles["limits"]["rich_rendered_bytes"], MAX_MESSAGE)
        self.assertFalse(styles["limits"]["images"])

    def test_all_message_routes_pass_rendered_html_and_guards(self):
        for action in ("send", "reply", "update"):
            for target in ("--chat", "--channel"):
                if action == "reply" and target == "--chat":
                    continue
                extra = (
                    [MID, "--if-version", "10"]
                    if action == "update"
                    else ["--reply-to", "root"]
                    if action == "reply"
                    else []
                )
                args = self.args(
                    action,
                    *extra,
                    target,
                    CHANNEL if target == "--channel" else THREAD,
                    "--body-format",
                    "markdown",
                    "--body",
                    "**Hello**",
                )
                api = Mock()
                dispatch(api, args)
                called = api.update if action == "update" else api.send
                kw = called.call_args.kwargs
                self.assertEqual(kw["body_type"], "HTML")
                self.assertEqual(kw["channel"], target == "--channel")
                self.assertIn("<strong>Hello</strong>", called.call_args.args[-1])
                if action == "reply":
                    self.assertEqual(kw["reply_to"], "root")
                if action == "update":
                    self.assertEqual(kw["expected"], "10")

    def test_html_sanitization_and_styles_share_mail_defaults(self):
        source = '<p onclick="x()" style="color:#123456">Hello <strong>world</strong></p><script>bad_token</script><a href="javascript:x()">link</a><at id="x">plain name</at>'
        args = self.args("send", "--chat", THREAD, "--body-format", "html", "--body", source)
        self.assertEqual(args.text, render(source, "html")["body"])
        for forbidden in ("onclick", "bad_token", "javascript:", "<at"):
            self.assertNotIn(forbidden, args.text)
        self.assertIn("Helvetica Neue", args.text)
        self.assertIn("#123456", args.text)
        self.assertTrue(args.composition["html_normalized"])
        self.assertGreater(len(args.composition["warnings"]), 1)

    def test_minimal_and_custom_styles_files_and_stdin(self):
        with tempfile.TemporaryDirectory() as tmp:
            style = Path(tmp) / "style.json"
            style.write_text('{"font_family":"Georgia, serif","font_size_pt":11}')
            source = Path(tmp) / "body.md"
            source.write_text("**Zażółć 📅**", encoding="utf-8")
            args = self.args(
                "send",
                "--chat",
                THREAD,
                "--body-format",
                "markdown",
                "--body-file",
                str(source),
                "--style-file",
                str(style),
            )
            self.assertIn("Georgia", args.text)
            self.assertIn("11pt", args.text)
            self.assertEqual(args.composition["style_options"]["font_size_pt"], 11)
            with patch("sys.stdin", io.StringIO(source.read_text())):
                stdin = self.args(
                    "send",
                    "--chat",
                    THREAD,
                    "--body-format",
                    "markdown",
                    "--body-file",
                    "-",
                    "--style",
                    "minimal",
                )
            self.assertEqual(stdin.text, "<p><strong>Zażółć 📅</strong></p>\n")

    def test_preview_never_overwrites_and_is_private(self):
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "preview.html"
            words = ["teams", "message", "render", "--body", "test", "--output", str(output)]
            data = run(parser().parse_args(words))
            self.assertEqual(data["message_content"], text_body("test"))
            self.assertEqual(output.stat().st_mode & 0o777, 0o600)
            self.error("output_exists", lambda: run(parser().parse_args(words)))

    def test_invalid_empty_media_and_styling_fail_before_connect(self):
        cases = [
            (["--body-format", "html", "--body", "<p> </p>"], "invalid_body"),
            (["--body-format", "html", "--body", "<script>bad</script>"], "invalid_body"),
            (
                ["--body-format", "markdown", "--body", "![image](https://example.com/a.png)"],
                "unsupported_body",
            ),
            (["--body-format", "html", "--body", '<img src="cid:logo">'], "unsupported_body"),
            (["--body", "text", "--style", "minimal"], "invalid_style"),
            (["--body-format", "markdown"], "body_required"),
            (["--body", "bad\x00text"], "invalid_body"),
        ]
        for words, code in cases:
            with patch("m365_cli.teams_cli.TeamsSession") as session:
                self.error(
                    code,
                    lambda words=words: run(
                        parser().parse_args(["teams", "message", "send", "--chat", THREAD, *words])
                    ),
                )
                session.assert_not_called()

    def test_input_and_rendered_byte_limits_and_bounded_stdin(self):
        self.assertEqual(len(content("a" * MAX_MESSAGE, "Text")), MAX_MESSAGE + 7)
        self.error("body_too_large", lambda: content("📅" * 5001, "Text"))
        self.error("body_too_large", lambda: content("<p>" + "a" * MAX_MESSAGE + "</p>", "HTML"))
        self.error("invalid_body_format", lambda: content("x", "Markdown"))
        stream = Mock()
        stream.read.return_value = "a" * (MAX_SOURCE + 1)
        with patch("sys.stdin", stream):
            self.error(
                "invalid_body",
                lambda: self.args(
                    "send", "--chat", THREAD, "--body-file", "-", "--body-format", "markdown"
                ),
            )
        stream.read.assert_called_once_with(MAX_SOURCE + 1)
        source = "\n\n".join("para" for _ in range(200))
        self.error(
            "body_too_large",
            lambda: self.args(
                "send", "--chat", THREAD, "--body-format", "markdown", "--body", source
            ),
        )

    def test_service_serialization_normalization_is_semantic_and_rich_only(self):
        before = '<div style="font-family:Helvetica Neue, Arial;font-size:12pt"><p><strong>Bold</strong> &amp; text</p><pre><code>x = "test"\n  y</code></pre></div>'
        after = '<div style="font-family:Helvetica Neue,Arial; font-size:12pt">\n<p>\n<strong>Bold</strong> &amp; text</p>\n\n<pre><code>x = &quot;test&quot;\n  y</code></pre>\n</div>'
        self.assertTrue(matches(before, after, "HTML"))
        self.assertFalse(matches(before, after, "Text"))
        for changed in (
            after.replace("Bold", "Other"),
            after.replace("12pt", "13pt"),
            after.replace("<strong>", "<em>").replace("</strong>", "</em>"),
            after.replace("\n  y", "\n y"),
            after.replace("</div>", "<script>x()</script></div>"),
        ):
            self.assertFalse(matches(before, changed, "HTML"))
        self.assertFalse(
            matches(
                '<a href="https://example.com/a">Link</a>',
                '<a href="https://example.com/b">Link</a>',
                "HTML",
            )
        )
        self.assertFalse(matches("<p>A <em>B</em></p>", "<p>A<em>B</em></p>", "HTML"))
        self.assertFalse(
            matches(
                '<p style="white-space:pre-wrap">A  B</p>',
                '<p style="white-space:pre-wrap">A B</p>',
                "HTML",
            )
        )
        self.assertFalse(
            matches(
                "<p style=\"font-family:'A  B'\">x</p>",
                "<p style=\"font-family:'A B'\">x</p>",
                "HTML",
            )
        )
        self.assertFalse(matches("<p>x</p>", "<p>x</div>", "HTML"))
        self.assertFalse(matches("<p>x</p>", None, "HTML"))

    def test_rich_send_normalizes_readback_without_escaping(self):
        body = "<p><strong>Hello</strong> &amp; world</p>"
        stored = "<p>\n<strong>Hello</strong> &amp; world</p>\n"
        session = FakeSession(
            [thread(), {"OriginalArrivalTime": MID}, message(content=stored, clientmessageid="1")]
        )
        with patch("m365_cli.teams_api.time.time_ns", return_value=1000):
            result = TeamsAPI(session).send(THREAD, body, body_type="HTML")
        self.assertEqual(session.calls[1][2]["body"]["content"], body)
        self.assertTrue(result["body_verification"]["server_normalized"])
        self.assertTrue(result["submitted"])

    def test_rich_update_retains_version_and_ownership_guards(self):
        body = "<p><em>New</em></p>"
        session = FakeSession([thread(), message(), {}, message(content=body, version="11")])
        result = TeamsAPI(session).update(THREAD, MID, body, body_type="HTML", expected="10")
        self.assertEqual(session.calls[2][2]["headers"]["If-Match"], '"10"')
        self.assertEqual(session.calls[2][2]["body"]["content"], body)
        self.assertTrue(result["updated"])
        for prior, code in (
            (message(), "conflict"),
            (message(**{"from": "someone_else"}), "not_owner"),
            (message(properties={"mentions": "[{}]"}), "unsupported_message"),
        ):
            session = FakeSession([thread(), prior])
            self.error(
                code,
                lambda session=session, code=code: TeamsAPI(session).update(
                    THREAD,
                    MID,
                    body,
                    body_type="HTML",
                    expected="9" if code == "conflict" else None,
                ),
            )
            self.assertEqual(len(session.calls), 2)

    def test_channel_reply_preserves_root_and_rich_content(self):
        inv = {
            "chats": [],
            "teams": [
                {"id": "team", "channels": [{"id": CHANNEL, "isMember": True, "channelType": 1}]}
            ],
        }
        body = "<p><em>Reply</em></p>"
        session = FakeSession(
            [
                inv,
                message(id="root"),
                {"OriginalArrivalTime": MID},
                message(content=body, clientmessageid="1", replyChainId="root"),
            ]
        )
        with patch("m365_cli.teams_api.time.time_ns", return_value=1000):
            result = TeamsAPI(session).send(
                CHANNEL, body, channel=True, reply_to="root", body_type="HTML"
            )
        self.assertIn("%3Bmessageid%3Droot", session.calls[2][1])
        self.assertEqual(session.calls[2][2]["body"]["replyChainId"], "root")
        self.assertEqual(result["message"]["reply_to"], "root")

    def test_update_requires_usable_new_version_and_preserved_rich_content(self):
        body = "<p><strong>New</strong></p>"
        for changed in (
            message(content=body, version=None),
            message(content=body, version="unknown"),
            message(content=body, version="10"),
            message(content="<p>New</p>", version="11"),
        ):
            session = FakeSession([thread(), message(), {}, changed])
            self.error(
                "readback_failed",
                lambda session=session: TeamsAPI(session).update(
                    THREAD, MID, body, body_type="HTML", expected="10"
                ),
            )
            self.assertEqual(len(session.calls), 4)

    def test_missing_ack_lost_format_and_changed_contract_do_not_retry(self):
        body = "<p><strong>Test</strong></p>"
        for response in ({}, {"id": MID}):
            session = FakeSession([thread(), response])
            self.error(
                "write_outcome_unknown",
                lambda session=session: TeamsAPI(session).send(THREAD, body, body_type="HTML"),
            )
            self.assertEqual(len(session.calls), 2)
        for stored in (
            message(content="<p>Test</p>", clientmessageid="1"),
            message(content=body, clientmessageid="1", messagetype="unknown"),
            OutlookError("session_rejected", "Expired"),
        ):
            session = FakeSession([thread(), {"OriginalArrivalTime": MID}, stored])
            with patch("m365_cli.teams_api.time.time_ns", return_value=1000):
                error = self.error(
                    "readback_failed",
                    lambda session=session: TeamsAPI(session).send(THREAD, body, body_type="HTML"),
                )
            self.assertEqual(error.details["message_id"], MID)
            self.assertEqual(len(session.calls), 3)

    def test_legacy_plain_fingerprints_replay_and_new_formats_conflict(self):
        args = self.args("send", "--chat", THREAD, "--body", "**Test**")
        old = copy.copy(args)
        for key in (
            "body_composition",
            "body_type",
            "body_format",
            "style",
            "style_file",
            "composition",
        ):
            delattr(old, key)
        self.assertEqual(plan(args), plan(old))
        rich = self.args(
            "send", "--chat", THREAD, "--body", "**Test**", "--body-format", "markdown"
        )
        self.assertNotEqual(plan(rich), plan(old))
        with tempfile.TemporaryDirectory() as tmp, patch.dict("os.environ", {"M365_HOME": tmp}):
            action = Mock(return_value={"submitted": True})
            execute("same", "test@example.com", plan(old), action)
            self.assertTrue(execute("same", "test@example.com", plan(args), action)["replayed"])
            self.error(
                "request_id_conflict",
                lambda: execute("same", "test@example.com", plan(rich), action),
            )
            self.assertEqual(action.call_count, 1)
            self.error(
                "interrupted",
                lambda: execute(
                    "pending",
                    "test@example.com",
                    plan(rich),
                    lambda: (_ for _ in ()).throw(
                        OutlookError("interrupted", "Synthetic interruption")
                    ),
                ),
            )
            self.error(
                "write_outcome_unknown",
                lambda: execute("pending", "test@example.com", plan(rich), action),
            )
            self.assertEqual(action.call_count, 1)

    def test_rich_styles_are_bound_by_value_and_dry_run_is_local(self):
        with tempfile.TemporaryDirectory() as tmp:
            style = Path(tmp) / "style.json"
            words = [
                "teams",
                "message",
                "send",
                "--chat",
                THREAD,
                "--body",
                "**Test**",
                "--body-format",
                "markdown",
                "--style-file",
                str(style),
                "--dry-run",
            ]
            with patch("m365_cli.teams_cli.TeamsSession") as session:
                style.write_text('{"font_size_pt":11}')
                first = run(parser().parse_args(words))["plan"]
                style.write_text('{"font_size_pt":14}')
                second = run(parser().parse_args(words))["plan"]
                session.assert_not_called()
            self.assertNotEqual(first, second)
            self.assertNotIn("style_file", first)
            self.assertEqual(first["composition"]["style_options"]["font_size_pt"], 11)

    def test_maintainer_inventory_and_channel_suite_option(self):
        caps = {c["command"]: c for c in inventory()["capabilities"]}
        for command in ("send", "reply", "update", "get", "delete"):
            row = caps["teams message " + command]
            self.assertEqual(row["validation_suite"], "teams")
            self.assertIn("teams-composition", row["additional_validation_suites"])
        self.assertEqual(caps["teams message render"]["validation_suite"], "teams-composition")
        args = parser().parse_args(
            [
                "--account",
                "test@example.com",
                "maintainer",
                "validate",
                "--suite",
                "teams-composition",
                "--channel",
                CHANNEL,
                "--allow-test-writes",
            ]
        )
        with (
            tempfile.TemporaryDirectory() as tmp,
            patch.dict("os.environ", {"M365_HOME": tmp}),
            patch(
                "m365_cli.maintainer.subprocess.run", return_value=Mock(returncode=0, stdout="")
            ) as child,
        ):
            run(args)
        argv = child.call_args.args[0]
        self.assertIn("m365_cli.live_teams_composition", argv)
        self.assertEqual(argv[argv.index("--channel") + 1], CHANNEL)
