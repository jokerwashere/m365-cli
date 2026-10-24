"""Synthetic rich-mail rendering, transport, compatibility and retry contracts."""

import io
import json
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from test_api import FakeSession, response

from m365_cli.api import API
from m365_cli.cli import dispatch, parser, plan, prepare, run
from m365_cli.composition import DEFAULT_STYLE, MAX_SOURCE, render
from m365_cli.errors import OutlookError
from m365_cli.maintainer import inventory
from m365_cli.operations import execute


class CompositionTests(unittest.TestCase):
    def error(self, code, action):
        with self.assertRaises(OutlookError) as caught:
            action()
        self.assertEqual(caught.exception.code, code)

    def args(self, *words):
        return prepare(parser().parse_args(list(words)))

    def test_plain_text_remains_literal_and_default(self):
        result = render("<b>literal</b>\n**not bold** 📨")
        self.assertEqual(result["body_type"], "Text")
        self.assertEqual(result["body"], "<b>literal</b>\n**not bold** 📨")
        self.assertIsNone(render(None))
        self.assertEqual(render("")["body"], "")

    def test_markdown_semantics_styles_and_mail_line_breaks(self):
        source = '# Update\n\nHello **team**, *welcome* & ~~old~~.\nSecond line\n\n- First\n- Second\n\n> Quote\n\n```python\nprint("<literal>")\n```\n\n| A | B |\n| - | - |\n| 1 | 2 |\n\n[Link](https://example.com?a=1&b=2)'
        output = render(source, "markdown")["body"]
        for expected in (
            "<h1",
            "<strong>team</strong>",
            "<em>welcome</em>",
            "<s>old</s>",
            "<br>",
            "<ul",
            "<blockquote",
            "<pre",
            "<table",
            "&lt;literal&gt;",
            "font-size:12pt",
            "https://example.com?a=1&amp;b=2",
        ):
            self.assertIn(expected, output)
        self.assertNotIn("<style", output)

    def test_markdown_raw_html_is_literal_not_executable(self):
        output = render(
            '<script>alert(1)</script>\n\n<img src="https://example.com/a">', "markdown"
        )["body"]
        self.assertNotIn("<script", output)
        self.assertNotIn("<img", output)
        self.assertIn("&lt;script&gt;", output)

    def test_html_sanitizes_active_elements_and_unsafe_urls(self):
        source = '<script>secret_script</script><style>secret_style</style><svg><script>secret_svg</script></svg><p onclick="run()">Good<a href="java&#x73;cript:run()">bad</a><img src="file:///private/file" onerror="run()"><iframe src="https://example.com">frame</iframe><a href="/relative">relative</a></p>'
        data = render(source, "html")
        for blocked in (
            "secret_script",
            "secret_style",
            "secret_svg",
            "onclick",
            "onerror",
            "javascript:",
            "file:",
            "<iframe",
            'href="/relative"',
        ):
            self.assertNotIn(blocked, data["body"])
        self.assertIn("Good", data["body"])
        self.assertTrue(data["composition"]["warnings"])

    def test_safe_links_cid_images_and_inline_css_survive(self):
        source = '<p style="color:#123456;text-align:center">Hello <a href="mailto:user@example.com">mail</a><img src="cid:logo@example.com" alt="Logo" width="80"><a href="https://example.com">site</a></p>'
        output = render(source, "html")["body"]
        for part in (
            "color:#123456",
            "text-align:center",
            "cid:logo@example.com",
            "mailto:user@example.com",
            'width="80"',
        ):
            self.assertIn(part, output)
        md = render("![Logo](cid:logo@example.com)", "markdown")["body"]
        self.assertIn('src="cid:logo@example.com"', md)

    def test_css_fetches_expressions_and_obfuscation_are_removed(self):
        for value in (
            "background:url(https://example.com/tracker)",
            "width:expression(run())",
            r"color:red;--x:u\72l(https://example.com)",
            "color:var(--leak)",
            "color:red;/*hidden*/background:url(x)",
        ):
            result = render(f'<p style="{value}">safe</p>', "html", "minimal")["body"]
            self.assertNotIn("style=", result)
        self.assertNotIn(
            "position:", render('<p style="position:fixed;color:red">x</p>', "html")["body"]
        )

    def test_html_documents_are_reduced_to_body_content(self):
        result = render(
            "<!doctype html><html><head><title>secret title</title><style>x</style></head><body><p>Hello</p></body></html>",
            "html",
        )["body"]
        self.assertIn("Hello", result)
        self.assertNotIn("secret title", result)
        self.assertNotIn("<html", result)

    def test_minimal_does_not_add_styles_but_sanitizes(self):
        result = render("**Hello**", "markdown", "minimal")["body"]
        self.assertEqual(result, "<p><strong>Hello</strong></p>\n")
        self.assertNotIn("onclick", render('<p onclick="x()">Hello</p>', "html", "minimal")["body"])

    def test_style_file_is_resolved_and_validated(self):
        with tempfile.TemporaryDirectory() as d:
            path = Path(d) / "style.json"
            path.write_text(
                json.dumps(
                    {
                        "font_family": "Georgia, serif",
                        "font_size_pt": 11,
                        "text_color": "#17365d",
                        "paragraph_spacing_pt": 6,
                    }
                )
            )
            data = render("**Hello**", "markdown", style_file=path)
            self.assertIn("font-family:Georgia, serif", data["body"])
            self.assertIn("font-size:11pt", data["body"])
            self.assertEqual(
                data["composition"]["style_options"]["link_color"], DEFAULT_STYLE["link_color"]
            )
            for value in (
                {"unknown": 1},
                {"font_size_pt": True},
                {"line_height": float("nan")},
                {"text_color": "red;position:fixed"},
                {"font_family": "Arial; background:url(x)"},
                {"paragraph_spacing_pt": 100},
                [],
            ):
                path.write_text(json.dumps(value))
                self.error("invalid_style", lambda: render("x", "markdown", style_file=path))
            path.write_text("{broken")
            self.error("invalid_style", lambda: render("x", "markdown", style_file=path))

    def test_invalid_combinations_fail_before_browser(self):
        with patch("m365_cli.cli.Session") as session:
            for words, code in (
                (["--body-format", "markdown"], "body_required"),
                (["--body", "hello", "--style", "outlook"], "invalid_style"),
                (
                    [
                        "--body",
                        "x",
                        "--body-format",
                        "html",
                        "--style",
                        "minimal",
                        "--style-file",
                        "x",
                    ],
                    "invalid_style",
                ),
            ):
                self.error(
                    code,
                    lambda words=words: run(
                        parser().parse_args(["mail", "draft", "create", "--subject", "x", *words])
                    ),
                )
            session.assert_not_called()
        self.error("invalid_body", lambda: render("x" * (MAX_SOURCE + 1), "markdown"))
        self.error("invalid_body", lambda: render("nul\x00", "html"))
        self.error("body_too_complex", lambda: render("<br>" * 20001, "html", "minimal"))

    def test_body_file_stdin_and_explicit_empty_update(self):
        with patch("sys.stdin", io.StringIO("**Zażółć 📨**")):
            a = self.args(
                "mail",
                "send",
                "--to",
                "user@example.com",
                "--subject",
                "x",
                "--body-file",
                "-",
                "--body-format",
                "markdown",
            )
        self.assertIn("<strong>Zażółć 📨</strong>", a.text)
        a = self.args(
            "mail",
            "draft",
            "update",
            "id",
            "--body",
            "",
            "--body-format",
            "html",
            "--style",
            "minimal",
        )
        self.assertEqual(a.changes["Body"], {"BodyType": "HTML", "Value": ""})
        a = self.args("mail", "draft", "update", "id", "--subject", "new")
        self.assertNotIn("Body", a.changes)

    def test_mail_file_reader_is_bounded_before_rendering(self):
        with patch("m365_cli.cli.Path.open") as opened:
            opened.return_value.__enter__.return_value.read.return_value = "x" * (MAX_SOURCE + 1)
            self.error(
                "invalid_body", lambda: self.args("mail", "render", "--body-file", "large.txt")
            )
            opened.return_value.__enter__.return_value.read.assert_called_once_with(MAX_SOURCE + 1)

    def test_local_render_needs_no_account_or_browser_and_never_overwrites(self):
        with tempfile.TemporaryDirectory() as d, patch("m365_cli.cli.Session") as session:
            output = Path(d) / "preview.html"
            result = run(
                parser().parse_args(
                    [
                        "mail",
                        "render",
                        "--body-format",
                        "markdown",
                        "--body",
                        "**Hi**",
                        "--output",
                        str(output),
                    ]
                )
            )
            self.assertFalse(result["submitted"])
            self.assertTrue(result["local_only"])
            self.assertIn("Content-Security-Policy", output.read_text())
            self.assertIn("img-src 'none'", output.read_text())
            self.assertEqual(output.stat().st_mode & 0o777, 0o600)
            self.error(
                "output_exists",
                lambda: run(
                    parser().parse_args(
                        ["mail", "render", "--body", "different", "--output", str(output)]
                    )
                ),
            )
            self.assertIn("<strong>Hi</strong>", output.read_text())
            self.assertEqual(
                run(parser().parse_args(["mail", "styles"]))["defaults"], DEFAULT_STYLE
            )
            session.assert_not_called()

    def test_local_preview_escapes_plain_text(self):
        result = run(parser().parse_args(["mail", "render", "--body", "<img src=x onerror=run()>"]))
        self.assertIn("&lt;img", result["preview_html"])
        self.assertNotIn("<img", result["preview_html"])

    def test_legacy_plain_text_operation_fingerprint_is_unchanged(self):
        a = self.args(
            "mail", "send", "--to", "user@example.com", "--subject", "x", "--body", "hello"
        )
        value = plan(a)
        self.assertFalse(
            set(value)
            & {"body_format", "style", "style_file", "mail_composition", "body_type", "composition"}
        )
        self.assertEqual(value["text"], "hello")

    def test_changed_style_file_invalidates_request_id(self):
        with tempfile.TemporaryDirectory() as d, patch.dict(os.environ, {"M365_HOME": d}):
            path = Path(d) / "style.json"
            path.write_text('{"font_size_pt":11}')
            words = [
                "mail",
                "send",
                "--to",
                "user@example.com",
                "--subject",
                "x",
                "--body",
                "Hi",
                "--body-format",
                "markdown",
                "--style-file",
                str(path),
            ]
            first = plan(self.args(*words))
            action = Mock(return_value={"submitted": True})
            execute("test", "self@example.com", first, action)
            self.assertTrue(execute("test", "self@example.com", first, action)["replayed"])
            path.write_text('{"font_size_pt":12}')
            second = plan(self.args(*words))
            self.assertNotEqual(first, second)
            self.error(
                "request_id_conflict", lambda: execute("test", "self@example.com", second, action)
            )
            action.assert_called_once()

    def test_dry_run_has_actual_html_and_resolved_style(self):
        with patch("m365_cli.cli.Session") as session:
            result = run(
                parser().parse_args(
                    [
                        "mail",
                        "send",
                        "--to",
                        "user@example.com",
                        "--subject",
                        "x",
                        "--body",
                        "**Hi**",
                        "--body-format",
                        "markdown",
                        "--dry-run",
                    ]
                )
            )
            self.assertEqual(result["plan"]["body_type"], "HTML")
            self.assertIn("<strong>Hi</strong>", result["plan"]["text"])
            session.assert_not_called()

    def test_create_and_send_use_html_and_existing_dispositions(self):
        for send in (False, True):
            s = FakeSession([response({"ItemId": {"Id": "new", "ChangeKey": "v"}})])
            api = API(s)
            (api.mail_send if send else api.draft_create)(
                subject="x", body="<b>Hi</b>", body_type="HTML", to=["user@example.com"]
            )
            body = s.calls[0][1]
            self.assertEqual(body["Items"][0]["Body"], {"BodyType": "HTML", "Value": "<b>Hi</b>"})
            self.assertEqual(body["MessageDisposition"], "SendAndSaveCopy" if send else "SaveOnly")

    def test_reply_replyall_and_forward_use_new_html_body_without_rebuilding_quote(self):
        source = {"ItemId": {"Id": "source", "ChangeKey": "v"}, "IsDraft": False}
        for mode in ("reply", "reply-all", "forward"):
            for send in (False, True):
                s = FakeSession([response(source), response({"ItemId": {"Id": "new"}})])
                API(s).respond_mail(
                    "source",
                    body="<b>New</b>",
                    body_type="HTML",
                    to=["user@example.com"],
                    mode=mode,
                    send=send,
                    expected="v",
                )
                item = s.calls[1][1]["Items"][0]
                self.assertEqual(item["NewBodyContent"]["BodyType"], "HTML")
                self.assertEqual(item["NewBodyContent"]["Value"], "<b>New</b>")
                self.assertEqual(item["ReferenceItemId"]["ChangeKey"], "v")
                self.assertNotIn("Body", item)

    def test_read_html_is_optin_and_shape_drift_fails(self):
        item = {"ItemId": {"Id": "id"}, "Body": {"BodyType": "HTML", "Value": "<p>x</p>"}}
        s = FakeSession([response(item)])
        self.assertEqual(API(s).get("id", body_type="HTML")["Body"], item["Body"])
        self.assertEqual(s.calls[0][1]["ItemShape"]["BodyType"], "HTML")
        for value in (
            None,
            {},
            {"BodyType": "Text", "Value": "x"},
            {"BodyType": "HTML", "Value": None},
        ):
            s = FakeSession([response({**item, "Body": value})])
            self.error("api_changed", lambda s=s: API(s).get("id", body_type="HTML"))

    def test_rich_update_preserves_change_key_and_html_readback(self):
        item = {
            "ItemId": {"Id": "draft", "ChangeKey": "v"},
            "IsDraft": True,
            "Body": {"BodyType": "HTML", "Value": "<p>new</p>"},
        }
        s = FakeSession([response(item), response({}), response(item)])
        API(s).update("draft", {"Body": item["Body"]}, expected="v")
        update = s.calls[1][1]
        self.assertEqual(update["ConflictResolution"], "NeverOverwrite")
        self.assertEqual(update["ItemChanges"][0]["ItemId"]["ChangeKey"], "v")
        self.assertEqual(s.calls[2][1]["ItemShape"]["BodyType"], "HTML")

    def test_api_invalid_types_and_unknown_send_never_retry(self):
        s = FakeSession([])
        self.error(
            "invalid_body_format",
            lambda: API(s).draft_create(subject="x", body="x", body_type="Markdown"),
        )
        self.assertEqual(s.calls, [])
        s = FakeSession([OutlookError("transport_error", "unknown")])
        self.error(
            "transport_error",
            lambda: API(s).mail_send(
                subject="x", body="<p>x</p>", body_type="HTML", to=["user@example.com"]
            ),
        )
        self.assertEqual(len(s.calls), 1)

    def test_all_cli_composition_routes_pass_rendered_html(self):
        for words in (
            ["mail", "send", "--to", "user@example.com", "--subject", "x"],
            ["mail", "draft", "create", "--subject", "x"],
            ["mail", "reply", "source"],
            ["mail", "reply", "source", "--all"],
            ["mail", "forward", "source", "--to", "user@example.com"],
            ["mail", "draft", "update", "source"],
        ):
            options = self.args(*words, "--body-format", "markdown", "--body", "**Hi**")
            api = Mock()
            dispatch(api, options)
            kwargs = api.mock_calls[0].kwargs
            if options.command == "draft" and options.draft_command == "update":
                self.assertEqual(api.mock_calls[0].args[1]["Body"]["BodyType"], "HTML")
            else:
                self.assertEqual(kwargs["body_type"], "HTML")
                self.assertIn("<strong>Hi</strong>", kwargs["body"])

    def test_capability_inventory_preserves_primary_suite_and_adds_composition(self):
        caps = {c["command"]: c for c in inventory()["capabilities"]}
        self.assertEqual(caps["mail send"]["validation_suite"], "outlook")
        self.assertIn("composition", caps["mail send"]["additional_validation_suites"])
        self.assertEqual(caps["mail render"]["validation_suite"], "composition")
        self.assertFalse(caps["mail render"]["mutation"])
        self.assertFalse(caps["mail styles"]["mutation"])
