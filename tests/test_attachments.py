"""Synthetic typed attachment contracts, nesting, exports, local I/O and references."""

import base64
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from test_api import FakeSession, response

from m365_cli.api import API, typed
from m365_cli.attachments import describe, native_format, response_attachments
from m365_cli.cli import dispatch, parser, prepare
from m365_cli.errors import OutlookError
from m365_cli.references import References, get, outlook_url
from m365_cli.session import Session


def attachment(identifier="a", category="FileAttachment", **extra):
    return typed(category, AttachmentId={"Id": identifier}, Name="../../unsafe", **extra)


def result(*items):
    return {
        "Body": {
            "ResponseMessages": {"Items": [{"ResponseCode": "NoError", "Attachments": list(items)}]}
        }
    }


def root(*items):
    return response(
        {"ItemId": {"Id": "mail", "ChangeKey": "ck"}, "Attachments": list(items), "IsDraft": True}
    )


def embedded(identifier="outer", *items, category="Message", **extra):
    return attachment(
        identifier, "ItemAttachment", Item=typed(category, Attachments=list(items), **extra)
    )


class AttachmentTests(unittest.TestCase):
    def error(self, code, call):
        with self.assertRaises(OutlookError) as caught:
            call()
        self.assertEqual(caught.exception.code, code)

    def test_never_download_nested_file_as_outer_attachment(self):
        inner = attachment("child", Content=base64.b64encode(b"wrong file").decode())
        outer = embedded("outer", inner, category="Contact", DisplayName="Contact")
        api = API(FakeSession([result(outer)]))
        value = api.attachment_get("outer")
        self.assertEqual(value["kind"], "item")
        self.assertEqual(value["export_format"], "json")
        self.assertEqual(json.loads(value["content"])["Item"]["DisplayName"], "Contact")
        self.assertNotIn(b"wrong file", value["content"])
        self.assertEqual(len(response_attachments(result(outer))), 1)

    def test_nested_get_response_id_shortening_preserves_full_locator(self):
        full = base64.b64encode(
            b"synthetic-root" + bytes([3]) + b"A" * 20 + b"B" * 20 + b"C" * 20
        ).decode()
        shortened = base64.b64encode(b"synthetic-root" + bytes([1]) + b"C" * 20).decode()
        value = API(FakeSession([result(attachment(shortened, Content="eA=="))])).attachment_get(
            full
        )
        self.assertEqual(value["id"], full)
        self.assertEqual(value["AttachmentId"]["Id"], full)
        self.assertEqual(value["content"], b"x")
        wrong = base64.b64encode(b"synthetic-root" + bytes([1]) + b"D" * 20).decode()
        self.error(
            "api_changed",
            lambda: API(FakeSession([result(attachment(wrong, Content="eA=="))])).attachment_get(
                full
            ),
        )

    def test_file_bytes_empty_binary_and_inline(self):
        for content in (b"", b"\x00\xff\xfe\x01", b"some content"):
            a = attachment(
                Content=base64.b64encode(content).decode(),
                IsInline=True,
                ContentId="pic@example.test",
            )
            value = API(FakeSession([result(a)])).attachment_get("a")
            self.assertEqual(value["content"], content)
            self.assertEqual(value["ContentId"], "pic@example.test")
            self.assertTrue(value["is_inline"])
            self.assertNotIn("Content", value)

    def test_malformed_metadata_and_oversized_content_fail_explicitly(self):
        for values in ({"IsInline": "false"}, {"Size": "10"}, {"ContentType": []}):
            self.error("api_changed", lambda values=values: describe(attachment(**values)))
        with patch("m365_cli.attachments.MAX_DOWNLOAD_BYTES", 8):
            encoded = base64.b64encode(b"x" * 9).decode()
            self.error(
                "attachment_too_large",
                lambda: API(FakeSession([result(attachment(Content=encoded))])).attachment_get("a"),
            )

    def test_response_contract_and_identity_fail_closed(self):
        cases = [
            result(),
            result(attachment("wrong", Content="")),
            result(attachment(), attachment()),
            {"Body": {"nested": result(attachment(Content=""))}},
            result({"Name": "x", "Content": ""}),
        ]
        for value in cases:
            self.error(
                "api_changed", lambda value=value: API(FakeSession([value])).attachment_get("a")
            )
        self.error("write_outcome_unknown", lambda: response_attachments({"Body": {}}, write=True))
        self.error(
            "api_changed",
            lambda: API(FakeSession([result(attachment(Content="bad!!"))])).attachment_get("a"),
        )
        self.error(
            "api_changed", lambda: API(FakeSession([result(attachment())])).attachment_get("a")
        )
        self.error(
            "api_changed",
            lambda: API(
                FakeSession([result(attachment("a", "ItemAttachment", Item=None))])
            ).attachment_get("a"),
        )

    def test_item_json_keeps_all_properties_without_claiming_mailbox_id(self):
        for item_type in (
            "Message",
            "CalendarItem",
            "Contact",
            "Task",
            "DistributionList",
            "FutureItem",
        ):
            a = embedded(
                "outer", category=item_type, Subject="Example", CustomProperty={"anything": 42}
            )
            value = API(FakeSession([result(a)])).attachment_get("outer", format="json")
            self.assertEqual(
                json.loads(value["content"])["Item"]["CustomProperty"], {"anything": 42}
            )
            self.assertNotIn("id", value["Item"])

    def test_message_native_export_and_explicit_raw_other_item(self):
        for category, format in (
            ("Message", "auto"),
            ("CalendarItem", "raw"),
            ("Contact", "raw"),
            ("Task", "raw"),
        ):
            session = FakeSession([result(embedded("outer", category=category))])
            payload = b"MIME-Version: 1.0\r\nContent-Type: text/plain\r\n\r\nMessage"
            session.download_attachment = Mock(
                return_value={"content": payload, "content_type": "text/x-vcard"}
            )
            value = API(session).attachment_get("outer", format=format)
            self.assertEqual(value["content"], payload)
            self.assertEqual(value["export_content_type"], "message/rfc822")
            self.assertEqual(value["export_format"], "mime")
            session.download_attachment.assert_called_once_with("outer")

    def test_native_export_format_from_bytes_not_filename(self):
        for content, expected in (
            (b"BEGIN:VCALENDAR\r\n", "icalendar"),
            (b"BEGIN:VCARD\r\n", "vcard"),
            (bytes.fromhex("d0cf11e0a1b11ae1"), "msg"),
        ):
            self.assertEqual(native_format(content)[0], expected)
        for payload in (b"<html>login</html>", b'{"error": "oops"}', b"unexpected"):
            self.error("api_changed", lambda payload=payload: native_format(payload))

    def test_cloud_link_and_unknown_metadata_never_downloaded(self):
        for category in ("ReferenceAttachment", "LinkAttachment", "FutureAttachment"):
            a = attachment(
                category=category,
                AttachLongPathName="https://example.test/file",
                Content="this must not be decoded",
            )
            session = FakeSession([root(a)])
            value = API(session).attachment_get("a", mail_id="mail")
            self.assertEqual(value["export_format"], "json")
            self.assertEqual([c[0] for c in session.calls], ["GetItem"])
            self.error(
                "unsupported_attachment_content",
                lambda a=a: API(FakeSession([root(a)])).attachment_get(
                    "a", mail_id="mail", format="raw"
                ),
            )
            self.assertNotIn("Content", describe(a))

    def test_metadata_removes_binary_and_ephemeral_download_fields(self):
        a = embedded(
            "outer",
            attachment(Content="Yg==", AttachmentOriginalUrl="data:token", DownloadToken="secret"),
        )
        text = json.dumps(describe(a))
        self.assertNotIn("secret", text)
        self.assertNotIn("data:token", text)
        self.assertNotIn("Yg==", text)

    def test_recursive_tree_depth_and_correct_nested_bytes(self):
        leaf = attachment("leaf", Content=base64.b64encode(b"leaf").decode())
        inner = embedded("inner", leaf)
        outer = embedded("outer", inner)
        session = FakeSession(
            [root(attachment("outer", "ItemAttachment")), result(outer), result(inner)]
        )
        value = API(session).attachment_list("mail", recursive=True)
        self.assertEqual([a["id"] for a in value["items"]], ["outer", "inner", "leaf"])
        self.assertEqual([a["depth"] for a in value["items"]], [0, 1, 2])
        self.assertEqual(value["items"][2]["ancestor_attachment_ids"], ["outer", "inner"])
        self.assertTrue(value["complete"])
        self.assertNotIn("Item", value["items"][0])
        self.assertEqual(
            API(FakeSession([result(leaf)])).attachment_get("leaf")["content"], b"leaf"
        )

    def test_truncation_limits_cycles_and_future_types_are_explicit(self):
        cases = [
            ([root(embedded("outer"))], {"max_depth": 0}, "max_depth"),
            ([root(attachment("one"), attachment("two"))], {"max_items": 1}, "max_items"),
            (
                [
                    root(embedded("outer")),
                    result(embedded("outer", attachment("outer", "ItemAttachment"))),
                ],
                {},
                "repeated_attachment_id",
            ),
            ([root(attachment("a", "FutureAttachment"))], {}, "unsupported_attachment_type"),
            (
                [
                    root(embedded("outer"), attachment("two")),
                    result(embedded("outer", attachment("leaf"))),
                ],
                {"max_items": 2},
                "max_items",
            ),
        ]
        for replies, options, reason in cases:
            value = API(FakeSession(replies)).attachment_list("mail", recursive=True, **options)
            self.assertFalse(value["complete"])
            self.assertEqual(value["truncations"][0]["reason"], reason)
        for options in (
            {"max_depth": -1},
            {"max_depth": 31},
            {"max_items": 0},
            {"max_items": 1001},
        ):
            self.error(
                "invalid_argument",
                lambda options=options: API(FakeSession([])).attachment_list("mail", **options),
            )

    def test_parent_membership_is_checked_including_nested(self):
        outer = embedded("outer", attachment("leaf"))
        session = FakeSession([root(outer), result(outer), result(attachment("leaf", Content=""))])
        value = API(session).attachment_get("leaf", mail_id="mail")
        self.assertEqual(value["parent_attachment_id"], "outer")
        self.error(
            "attachment_not_found",
            lambda: API(FakeSession([root()])).attachment_inspect("other", mail_id="mail"),
        )

    def test_attach_existing_item_preserves_source_and_target_change_guard(self):
        source = response(
            {"ItemId": {"Id": "source", "ChangeKey": "source-version"}, "Subject": "Existing item"}
        )
        session = FakeSession([source, root(), result(attachment("added", "ItemAttachment"))])
        API(session).attachment_add_item("mail", "source", expected="ck")
        request = session.calls[-1][1]
        self.assertEqual(request["ParentItemId"]["ChangeKey"], "ck")
        self.assertEqual(request["Attachments"][0]["__type"], "ItemIdAttachment:#Exchange")
        self.assertEqual(request["Attachments"][0]["ItemId"]["Id"], "source")
        self.error(
            "invalid_argument", lambda: API(FakeSession([])).attachment_add_item("mail", "mail")
        )

    def test_inline_upload_and_nested_delete_guard(self):
        session = FakeSession([root(), result(attachment())])
        API(session).attachment_add(
            "mail", "image.png", b"PNG", inline=True, content_id="image@example.test"
        )
        added = session.calls[-1][1]["Attachments"][0]
        self.assertTrue(added["IsInline"])
        self.assertEqual(added["ContentId"], "image@example.test")
        self.error(
            "attachment_not_found",
            lambda: API(
                FakeSession([root(embedded("outer", attachment("leaf")))])
            ).attachment_delete("mail", "leaf"),
        )

    def test_cli_options_validate_before_session(self):
        cases = [
            (["list", "mail", "--max-depth", "31"], "invalid_argument"),
            (["add", "mail", "--item-id", "mail"], "invalid_argument"),
            (["add", "mail", "--item-id", "source", "--inline"], "invalid_argument"),
            (["get", "a", "--format", "raw"], "invalid_argument"),
        ]
        cases.extend(
            [
                (["add", "mail", "--file", ""], "invalid_argument"),
                (["get", "a", "--output", ""], "invalid_argument"),
            ]
        )
        for words, code in cases:
            self.error(
                code,
                lambda words=words: prepare(parser().parse_args(["mail", "attachment", *words])),
            )
        value = prepare(parser().parse_args(["mail", "attachment", "get", "a"]))
        self.assertIsNone(value.output)

    def test_explicit_path_private_json_with_references_no_overwrite(self):
        with tempfile.TemporaryDirectory() as temp, patch.dict("os.environ", {"M365_HOME": temp}):
            output = Path(temp) / "chosen.json"
            args = prepare(
                parser().parse_args(
                    [
                        "mail",
                        "attachment",
                        "get",
                        "a",
                        "--mail-id",
                        "mail",
                        "--format",
                        "json",
                        "--output",
                        str(output),
                    ]
                )
            )
            a = attachment(
                "a", "ReferenceAttachment", AttachLongPathName="https://example.test/file"
            )
            value = dispatch(
                API(FakeSession([root(a)])), args, refs=References("reader@example.com")
            )
            saved = json.loads(output.read_text())
            self.assertEqual(saved["stable_id"], value["stable_id"])
            self.assertEqual(saved["web_url"], outlook_url("mail", "mail"))
            self.assertEqual(output.stat().st_mode & 0o777, 0o600)
            self.error("output_exists", lambda: prepare(args))
            self.assertFalse((Path(temp).parent / "unsafe").exists())

    def test_nested_references_use_root_mail_and_invalidate_on_outer_delete(self):
        with tempfile.TemporaryDirectory() as temp, patch.dict("os.environ", {"M365_HOME": temp}):
            args = parser().parse_args(["mail", "attachment", "list", "mail", "--recursive"])
            refs = References("reader@example.com")
            value = refs.outlook({"items": [describe(embedded("outer", attachment("leaf")))]}, args)
            child = value["items"][0]["Item"]["Attachments"][0]
            self.assertEqual(child["web_url"], outlook_url("mail", "mail"))
            sid = child["stable_id"]
            refs.outlook(
                {"id": "mail", "attachment_id": "outer", "deleted": True},
                parser().parse_args(
                    ["mail", "attachment", "delete", "mail", "--attachment-id", "outer"]
                ),
            )
            cached = get(sid, "reader@example.com")
            self.assertEqual(cached["status"], "parent_deleted")
            self.assertIsNone(cached["web_url"])

    def test_standalone_reference_survives_later_parent_discovery(self):
        with tempfile.TemporaryDirectory() as temp, patch.dict("os.environ", {"M365_HOME": temp}):
            refs = References("reader@example.com")
            item = refs.outlook(
                describe(attachment()), parser().parse_args(["mail", "attachment", "get", "a"])
            )
            sid = item["stable_id"]
            self.assertIsNone(item["web_url"])
            listed = References("reader@example.com").outlook(
                {"items": [describe(attachment())]},
                parser().parse_args(["mail", "attachment", "list", "mail"]),
            )
            self.assertEqual(listed["items"][0]["stable_id"], sid)
            self.assertEqual(get(sid, "reader@example.com")["web_url"], outlook_url("mail", "mail"))

    def test_download_settings_reject_identity_host_or_missing_token(self):
        for settings, identity, code in [
            (
                {"DownloadUrlBase": "https://evil.test/owa/", "FirstDownloadToken": "secret"},
                "reader@example.com",
                "api_changed",
            ),
            (
                {"DownloadUrlBase": "https://attachments.office.net/owa/"},
                "reader@example.com",
                "api_changed",
            ),
            (
                {
                    "DownloadUrlBase": "https://attachments.office.net/owa/",
                    "FirstDownloadToken": "secret",
                },
                "other@example.com",
                "account_mismatch",
            ),
        ]:
            session = Session()
            session.account = "reader@example.com"
            session.page = Mock()
            session.request = Mock(
                return_value={"UserEmailAddress": identity, "ApplicationSettings": settings}
            )
            self.error(code, lambda session=session: session.download_attachment("a"))
            session.page.evaluate.assert_not_called()
