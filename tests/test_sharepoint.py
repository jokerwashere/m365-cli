"""Synthetic guards, transport isolation, schema validation and stable references."""

import base64
import json
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from test_transport import reply

from m365_cli.cli import parser, prepare, run
from m365_cli.errors import OutlookError
from m365_cli.http_transport import Response
from m365_cli.references import get
from m365_cli.sharepoint_api import SharePointAPI
from m365_cli.sharepoint_cli import save_download
from m365_cli.sharepoint_session import SharePointSession
from m365_cli.sharepoint_validation import etag, field_payload, filename, guid, infer_site, site_url

SITE = "https://example.sharepoint.com/sites/test"
GUID = "11111111-1111-1111-1111-111111111111"
OTHER = "22222222-2222-2222-2222-222222222222"
ACCOUNT = "reader@example.com"


class SharePointTests(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.tmp = Path(tmp.name)
        env = patch.dict("os.environ", {"M365_HOME": tmp.name})
        env.start()
        self.addCleanup(env.stop)
        self.session = SharePointSession(SITE, expected=ACCOUNT)
        self.session.site_id = self.session.web_id = self.session.tenant = GUID
        self.session.http = Mock()
        self.session.cookies = Mock()
        self.session.cookies.header.return_value = "session=synthetic"
        self.session.user_agent = "synthetic-agent"
        self.api = SharePointAPI(self.session)

    def error(self, code, callback):
        with self.assertRaises(OutlookError) as caught:
            callback()
        self.assertEqual(code, caught.exception.code)
        return caught.exception

    def test_site_and_route_bounds(self):
        for site in (
            "http://example.sharepoint.com/sites/test",
            "https://evil.test/sites/test",
            SITE + "/../other",
            SITE + "?x=1",
            "https://reader@example.sharepoint.com/sites/test",
            SITE + "/%2e%2e/other",
        ):
            self.error("invalid_argument", lambda site=site: site_url(site))
        for route in (
            "https://evil.test/_api/web",
            "https://example.sharepoint.com/sites/other/_api/web",
            "/_api/../admin",
            "/_api/%2e%2e/admin",
            "/_api/web#secret",
            "/_api/\\elsewhere",
        ):
            self.error("invalid_route", lambda route=route: self.session.route(route))
        self.assertEqual(SITE, infer_site(SITE + "/Lists/Example/DispForm.aspx?ID=7"))
        self.assertEqual(
            SITE, infer_site("https://example.sharepoint.com/:x:/r/sites/test/Docs/a.xlsx")
        )

    def test_strict_identifiers_and_names(self):
        for value in ("*", "stale", '"*"\r\nextra: true'):
            self.error("invalid_argument", lambda value=value: etag(value))
        for value in ("../a", "a/b", "a\\b", "file.", "file ", "~$file", "x\n"):
            self.error("invalid_argument", lambda value=value: filename(value))
        self.error("invalid_argument", lambda: guid("00000000-0000-0000-0000-000000000000"))
        self.assertEqual("a#50%.txt", filename("a#50%.txt"))

    def test_transport_digest_and_no_retry(self):
        self.session.verify_identity = Mock()
        self.session.http.request.side_effect = [
            reply({"FormDigestValue": "synthetic-digest", "FormDigestTimeoutSeconds": 1800}),
            reply({"error": {"code": "synthetic", "message": "SENSITIVE"}}, status=412),
        ]
        error = self.error(
            "version_conflict", lambda: self.session.request("/_api/web/lists", "POST", {})
        )
        self.assertNotIn("SENSITIVE", json.dumps(error.as_dict()))
        self.assertEqual(2, self.session.http.request.call_count)
        args = self.session.http.request.call_args.kwargs
        self.assertEqual("synthetic-digest", args["headers"]["X-RequestDigest"])
        self.assertNotIn("Authorization", args["headers"])
        self.session.verify_identity.assert_called_once()

    def test_transport_identity_change_and_missing_ack(self):
        self.session.http.request.return_value = reply({"Id": 1, "Email": "other@example.com"})
        self.error("identity_unknown", lambda: self.session.request("/_api/web", "POST"))
        self.assertFalse(self.session.write_attempted)
        self.assertEqual(self.session.http.request.call_count, 1)
        self.session.verify_identity = Mock()
        self.session.digest_until = float("inf")
        self.session.http.request.return_value = reply(None)
        self.error("write_outcome_unknown", lambda: self.session.request("/_api/web", "POST"))
        self.session.http.request.side_effect = TimeoutError("secret")
        self.error("write_outcome_unknown", lambda: self.session.request("/_api/web", "POST"))

    def test_status_errors_and_bounded_download(self):
        for status, code in (
            (401, "login_required"),
            (403, "access_denied"),
            (404, "not_found"),
            (429, "throttled"),
            (500, "service_error"),
        ):
            self.session.http.request.return_value = reply(
                {"error": {"code": "bad\nsecret"}}, status=status, **{"retry-after": "12"}
            )
            e = self.error(code, lambda: self.session.request("/_api/web"))
            self.assertNotIn("service_code", e.as_dict())
        self.session.http.request.side_effect = OutlookError("response_too_large", "bounded")
        self.error("response_too_large", lambda: self.session.request("/_api/web", binary=True))
        self.session.http.request.side_effect = None
        self.session.http.request.return_value = Response(200, b"bytes", {})
        self.assertEqual(b"bytes", self.session.request("/_api/web", binary=True))

    def test_site_identity_verified_again_before_write_and_digest_refresh(self):
        self.session.http.request.side_effect = [
            reply({"Id": 1, "Email": ACCOUNT}),
            reply({"Id": GUID}),
            reply({"Id": GUID, "Url": SITE}),
            reply({"FormDigestValue": "fresh", "FormDigestTimeoutSeconds": 60}),
            reply(None, status=204),
        ]
        self.session.digest_until = 0
        self.session.request("/_api/web/lists", "POST", {}, expected=(204,))
        calls = self.session.http.request.call_args_list
        self.assertEqual(len(calls), 5)
        self.assertTrue(all(not c.kwargs["write"] for c in calls[:4]))
        self.assertEqual(calls[-1].kwargs["headers"]["X-RequestDigest"], "fresh")
        self.assertTrue(calls[-1].kwargs["write"])

    def test_invalid_digest_lifetime_or_expired_cookies_prevents_mutation(self):
        self.session.verify_identity = Mock()
        for lifetime in (None, True, -1, 20, "1800"):
            self.session.http.request.return_value = reply(
                {"FormDigestValue": "synthetic", "FormDigestTimeoutSeconds": lifetime}
            )
            self.error("api_changed", lambda: self.session.request("/_api/web/lists", "POST", {}))
            self.assertFalse(self.session.write_attempted)
        self.session.cookies.header.return_value = ""
        self.error("login_required", lambda: self.session.request("/_api/web/lists"))

    def test_paging_and_bound_cursors(self):
        self.session.request = Mock(return_value={"value": [{"id": i} for i in range(3)]})
        route = "/_api/web/lists?$top=1"
        first = self.api.page(route, limit=1)
        self.assertEqual([{"id": 0}], first["items"])
        second = self.api.page(route, limit=1, cursor=first["next_cursor"])
        self.assertEqual([{"id": 1}], second["items"])
        last = self.api.page(route, limit=1, cursor=second["next_cursor"])
        self.assertTrue(last["complete"])
        self.error(
            "invalid_cursor",
            lambda: self.api.page("/_api/web/lists?$top=2", limit=2, cursor=first["next_cursor"]),
        )
        c = json.loads(base64.urlsafe_b64decode(first["next_cursor"]))
        c["url"] = "https://evil.test/_api/web/lists?$top=1"
        self.error(
            "invalid_cursor",
            lambda: self.api.page(
                route, limit=1, cursor=base64.urlsafe_b64encode(json.dumps(c).encode()).decode()
            ),
        )

    def test_continuation_query_cannot_change_scope(self):
        self.error(
            "api_changed",
            lambda: self.api.continuation(
                "/_api/web/lists?$filter=Hidden%20eq%20false", SITE + "/_api/web/lists?$skiptoken=2"
            ),
        )
        self.error(
            "api_changed",
            lambda: self.api.continuation("/_api/web/lists", SITE + "/_api/web/siteusers"),
        )
        self.session.request = Mock(return_value={"value": {}, "odata.nextLink": None})
        self.error("api_changed", lambda: self.api.page("/_api/web/lists"))

    def fields(self):
        return [
            {
                "internal_name": name,
                "type": kind,
                "read_only": False,
                "required": False,
                "choices": ["Open", "Closed"],
            }
            for name, kind in [
                ("Title", "Text"),
                ("State", "Choice"),
                ("People", "UserMulti"),
                ("Assignee", "User"),
                ("Value", "Number"),
                ("Done", "Boolean"),
                ("When", "DateTime"),
                ("Link", "URL"),
                ("Taxonomy", "TaxonomyFieldType"),
            ]
        ]

    def test_typed_payload_and_unknown_fields(self):
        fields = self.fields()
        out = field_payload(
            {"Title": "Example", "Assignee": 5, "People": [2, 3], "Done": True, "Value": 1.5},
            fields,
        )
        self.assertEqual(5, out["AssigneeId"])
        self.assertEqual({"results": [2, 3]}, out["PeopleId"])
        for data in (
            {"Unknown": 1},
            {"State": "Unlisted"},
            {"Value": True},
            {"Value": float("nan")},
            {"Assignee": "person@example.com"},
            {"People": None},
            {"People": [1, 1]},
            {"Done": 1},
            {"When": "2031-01-02"},
            {"Link": {"url": "javascript:alert(1)"}},
        ):
            with self.assertRaises(OutlookError):
                field_payload(data, fields)
        for value in ("x", None):
            self.error(
                "unsupported_field_type",
                lambda value=value: field_payload({"Taxonomy": value}, fields),
            )
        fields[0]["read_only"] = True
        self.error("invalid_argument", lambda: field_payload({"Title": "blocked"}, fields))

    def test_guard_checks_before_write(self):
        self.api.file = Mock(return_value={"id": GUID, "name": "old"})
        self.api.library_for = Mock()
        self.session.request = Mock(return_value={"Id": 1})
        self.session.last_etag = '"current"'
        self.error("version_conflict", lambda: self.api.rename(GUID, "new", expected='"old"'))
        self.assertEqual(1, self.session.request.call_count)
        self.error("unguarded_operation", lambda: self.api.recycle_file(GUID))

    def test_stable_file_reference_survives_rename_and_scopes_records(self):
        one = self.api.finish(
            self.api.ref({"id": GUID, "name": "old", "web_url": SITE + "/Docs/old"}, "file")
        )
        two = self.api.finish(
            self.api.ref({"id": GUID, "name": "new", "web_url": SITE + "/Docs/new"}, "file")
        )
        self.assertEqual(one["stable_id"], two["stable_id"])
        self.assertEqual(SITE + "/Docs/new", get(one["stable_id"], ACCOUNT)["web_url"])
        a = self.api.ref({"id": "1", "web_url": SITE}, "list_item", parent="list-A")
        b = self.api.ref({"id": "1", "web_url": SITE}, "list_item", parent="list-B")
        self.assertNotEqual(a["stable_id"], b["stable_id"])
        args = parser().parse_args(["--account", ACCOUNT, "files", "get", one["stable_id"]])
        prepare(args)
        self.assertEqual(GUID, args.id)
        self.assertEqual(SITE, args.site)
        self.assertEqual(GUID, args.sp_tenant)

    def test_cli_dry_run_and_no_unconditional_defaults(self):
        args = parser().parse_args(
            [
                "--account",
                ACCOUNT,
                "lists",
                "item",
                "update",
                "7",
                "--list",
                GUID,
                "--site",
                SITE,
                "--data",
                '{"Title":"Example"}',
                "--dry-run",
            ]
        )
        with patch(
            "m365_cli.sharepoint_session.SharePointSession.__enter__", side_effect=AssertionError
        ):
            result = run(args)
        self.assertTrue(result["server_validation_pending"])
        args = parser().parse_args(["files", "recycle", GUID, "--site", SITE])
        self.error("unguarded_operation", lambda: prepare(args))
        for command in ("replace", "publish", "delete"):
            self.error(
                "usage_error", lambda command=command: parser().parse_args(["files", command, GUID])
            )

    def test_download_exclusive_and_manifest_rollback(self):
        output = self.tmp / "artifact"
        args = SimpleNamespace(output=str(output), command="fetch", account=ACCOUNT, site=SITE)
        snapshot = {"sha256": "synthetic", "item": {"id": GUID}}
        result = save_download(args, (snapshot, b"content"))
        self.assertEqual(b"content", output.read_bytes())
        self.assertEqual(0o600, output.stat().st_mode & 0o777)
        self.assertEqual(GUID, json.loads(Path(result["manifest"]).read_text())["item"]["id"])
        self.error("local_output_failed", lambda: save_download(args, (snapshot, b"overwrite")))
        self.assertEqual(b"content", output.read_bytes())
        output.unlink()
        self.error("local_output_failed", lambda: save_download(args, (snapshot, b"new")))
        self.assertFalse(output.exists())
        self.assertTrue(Path(result["manifest"]).exists())

    def test_download_version_race_discards_bytes(self):
        self.api.file = Mock(side_effect=[{"id": GUID, "etag": '"1"'}, {"id": GUID, "etag": '"2"'}])
        self.session.request = Mock(return_value=b"bytes")
        self.error("version_conflict", lambda: self.api.download(GUID))

    def test_malformed_post_read_is_ambiguous(self):
        self.api.list_guard = Mock()
        self.api.validate_fields = Mock(return_value={"Title": "Example"})

        def request(*args, **kwargs):
            self.session.write_attempted = True
            return {}

        self.session.request = request
        self.error(
            "write_outcome_unknown", lambda: self.api.create_record(GUID, {"Title": "Example"})
        )

    def test_acknowledged_write_with_rejected_readback_is_not_a_rejected_write(self):
        self.session.verify_identity = Mock()
        self.session.digest_until = float("inf")
        for status, code in ((401, "login_required"), (403, "access_denied"), (429, "throttled")):
            self.session.http.request.reset_mock()
            self.session.http.request.side_effect = [
                reply({"Id": GUID}, status=201),
                reply({"error": {"message": "SENSITIVE"}}, status=status),
            ]
            error = self.error(
                "write_outcome_unknown", lambda: self.api.create_list("Synthetic fixture")
            )
            self.assertTrue(error.details["write_acknowledged"])
            self.assertEqual(error.details["cause"]["code"], code)
            self.assertNotIn("SENSITIVE", json.dumps(error.as_dict()))
            self.assertEqual(self.session.http.request.call_count, 2)
        # A new operation resets acknowledgement state. Actual write rejection
        # remains access_denied and is not replayed or disguised as readback.
        self.session.http.request.reset_mock()
        self.session.http.request.side_effect = [reply({}, status=403)]
        self.error("access_denied", lambda: self.api.create_list("Another fixture"))
        self.assertFalse(self.session.write_acknowledged)
        self.session.http.request.assert_called_once()

    def test_circuit_breaker_resolver_alias(self):
        from m365_cli.maintainer import guard, inventory

        with patch("m365_cli.maintainer.breakers", return_value={"files": "API changed"}):
            args = parser().parse_args(["reference", "resolve", SITE + "/Docs/a"])
            self.error("capability_disabled", lambda: guard(args))
        commands = {x["command"]: x for x in inventory()["capabilities"]}
        self.assertEqual("sharepoint", commands["lists item update"]["module"])
        self.assertTrue(commands["files move"]["mutation"])
        self.assertFalse(commands["files fetch"]["mutation"])

    def test_cursor_limit_bound_without_server_top(self):
        self.session.request = Mock(return_value={"value": [{"id": 1}, {"id": 2}]})
        first = self.api.page("/_api/web/lists", limit=1)
        self.error(
            "invalid_cursor",
            lambda: self.api.page("/_api/web/lists", limit=2, cursor=first["next_cursor"]),
        )

    def test_collection_routes_do_not_silently_truncate(self):
        self.session.request = Mock(return_value={"value": []})
        self.api.lists(limit=1)
        self.assertNotIn("$top", self.session.request.call_args.args[0])
        self.api.file = Mock(return_value={"id": GUID})
        self.api.children(GUID, limit=1)
        self.assertNotIn("%24top", self.session.request.call_args.args[0])
        self.api.versions(GUID, limit=1)
        self.assertNotIn("$top", self.session.request.call_args.args[0])

    def test_stable_record_cannot_override_explicit_list_handle(self):
        a = self.api.finish(self.api.ref({"id": GUID, "web_url": SITE}, "list"))
        b = self.api.finish(self.api.ref({"id": OTHER, "web_url": SITE}, "list"))
        record = self.api.finish(
            self.api.ref(
                {"id": "7", "list_id": GUID, "web_url": SITE}, "list_item", parent=a["stable_id"]
            )
        )
        args = parser().parse_args(
            [
                "--account",
                ACCOUNT,
                "lists",
                "item",
                "get",
                record["stable_id"],
                "--list",
                b["stable_id"],
            ]
        )
        self.error("invalid_reference", lambda: prepare(args))

    def test_deleted_parent_invalidates_child_handle(self):
        listing = self.api.finish(self.api.ref({"id": GUID, "web_url": SITE}, "list"))
        record = self.api.finish(
            self.api.ref(
                {"id": "7", "list_id": GUID, "web_url": SITE},
                "list_item",
                parent=listing["stable_id"],
            )
        )
        self.api.finish(self.api.ref({"id": GUID, "web_url": None}, "list", status="deleted"))
        self.assertIsNone(get(record["stable_id"], ACCOUNT)["web_url"])
        args = parser().parse_args(
            ["--account", ACCOUNT, "lists", "item", "get", record["stable_id"]]
        )
        self.error("invalid_reference", lambda: prepare(args))

    def test_file_version_reference_follows_parent_url(self):
        file = self.api.finish(self.api.ref({"id": GUID, "web_url": SITE + "/old"}, "file"))
        version = self.api.finish(
            self.api.ref(
                {"id": "512", "file_id": GUID, "web_url": SITE + "/old"},
                "file_version",
                parent=file["stable_id"],
                url_kind="parent",
            )
        )
        self.api.finish(self.api.ref({"id": GUID, "web_url": SITE + "/new"}, "file"))
        stored = get(version["stable_id"], ACCOUNT)
        self.assertEqual(SITE + "/new", stored["web_url"])
        self.assertNotIn("task_id", stored["locator"])
        self.assertEqual(GUID, stored["locator"]["file_id"])

    def test_complex_field_payload_uses_server_record_type(self):
        self.session.request = Mock(
            return_value={"ListItemEntityTypeFullName": "SP.Data.ExampleListItem"}
        )
        payload, headers = self.api.write_payload(
            GUID,
            {"Link": {"__metadata": {"type": "SP.FieldUrlValue"}, "Url": "https://example.com"}},
        )
        self.assertEqual("SP.Data.ExampleListItem", payload["__metadata"]["type"])
        self.assertEqual("application/json;odata=verbose", headers["Content-Type"])
        self.session.request.return_value = {"ListItemEntityTypeFullName": "unexpected:type"}
        self.error(
            "api_changed", lambda: self.api.write_payload(GUID, {"Link": {"__metadata": {}}})
        )

    def test_search_rejects_query_syntax_breakout(self):
        self.session.request = Mock()
        for query in ('x" OR path:other', 'x\\"'):
            self.error("invalid_argument", lambda query=query: self.api.search(query))
        self.session.request.assert_not_called()

    def test_resolver_rejects_cross_site_and_duplicate_parameters(self):
        self.session.request = Mock()
        for url, code in (
            (SITE + "/Lists/X/DispForm.aspx?ID=1&ID=2", "invalid_argument"),
            (
                "https://example.sharepoint.com/sites/other/Lists/X/DispForm.aspx?ID=1",
                "site_mismatch",
            ),
            ("https://other.sharepoint.com/sites/test/Docs/file", "invalid_argument"),
        ):
            self.error(code, lambda url=url: self.api.resolve(url))
        self.session.request.assert_not_called()

    def test_folder_create_existing_name_never_posts(self):
        self.api.file = Mock(return_value={"path": "/sites/test/Docs"})
        self.api.library_for = Mock()
        self.api.file_from_path = Mock(return_value={"id": OTHER})
        self.session.request = Mock()
        self.error("conflict", lambda: self.api.folder_create(GUID, "existing"))
        self.session.request.assert_not_called()

    def test_new_column_rejects_schema_changes_and_duplicate(self):
        self.api.list_guard = Mock()
        self.api.schema = Mock(
            return_value={"items": [{"name": "Title", "internal_name": "Title"}]}
        )
        self.session.request = Mock()
        self.error("conflict", lambda: self.api.create_column(GUID, "title", "text"))
        self.error("invalid_argument", lambda: self.api.create_column(GUID, "field name", "text"))
        self.error(
            "invalid_argument",
            lambda: self.api.create_column(GUID, "State", "choice", choices=["a", "a"]),
        )
        self.session.request.assert_not_called()

    def test_nonempty_recycle_rejected_before_write(self):
        self.api.file = Mock(return_value={"id": GUID, "item_count": 1})
        self.api.library_for = Mock()
        self.session.request = Mock()
        self.error(
            "resource_not_empty", lambda: self.api.recycle_file(GUID, folder=True, allow=True)
        )
        self.api.list_guard = Mock(return_value={"id": GUID, "item_count": 1})
        self.error("resource_not_empty", lambda: self.api.recycle_list(GUID, allow=True))
        self.session.request.assert_not_called()

    def test_cross_library_move_is_rejected(self):
        self.api.file = Mock(
            side_effect=[
                {"id": GUID, "path": "/sites/test/Docs/file.txt", "name": "file.txt"},
                {"id": OTHER, "path": "/sites/test/Other"},
            ]
        )
        self.session.request = Mock(return_value={"Id": GUID})
        self.api.library_for = Mock(side_effect=[{"id": GUID}, {"id": OTHER}])
        self.error("unsupported_operation", lambda: self.api.move(GUID, OTHER, allow=True))
        self.session.request.assert_not_called()

    def test_download_size_bound_and_unknown_revision(self):
        self.api.file = Mock(return_value={"id": GUID, "etag": '"1"'})
        self.api.scan = Mock(return_value=[{"ID": 512, "VersionLabel": "1.0"}])
        self.session.request = Mock()
        self.error("not_found", lambda: self.api.download(GUID, label="2.0"))
        self.session.request.assert_not_called()

    def test_request_id_binds_upload_content_and_site(self):
        from m365_cli.cli import plan
        from m365_cli.operations import execute

        local = self.tmp / "file.txt"
        local.write_bytes(b"one")

        def arguments():
            a = parser().parse_args(
                [
                    "--account",
                    ACCOUNT,
                    "files",
                    "upload",
                    "--folder",
                    GUID,
                    "--site",
                    SITE,
                    "--file",
                    str(local),
                    "--request-id",
                    "fixture",
                ]
            )
            prepare(a)
            return a

        first = plan(arguments())
        self.assertNotIn("file_content", first)
        call = Mock(return_value={"id": GUID})
        execute("fixture", ACCOUNT, first, call)
        local.write_bytes(b"two")
        self.error(
            "request_id_conflict", lambda: execute("fixture", ACCOUNT, plan(arguments()), call)
        )
        self.assertEqual(1, call.call_count)

    def test_library_boundary_excludes_generic_lists_and_root_mutations(self):
        self.api.scan = Mock(
            return_value=[{"Id": GUID, "RootFolder": {"ServerRelativeUrl": "/sites/test/Docs"}}]
        )
        self.assertEqual(GUID, self.api.library_for({"path": "/sites/test/Docs/folder"})["id"])
        self.error(
            "protected_resource",
            lambda: self.api.library_for({"path": "/sites/test/Docs"}, protect_root=True),
        )
        self.error(
            "unsupported_operation",
            lambda: self.api.library_for({"path": "/sites/test/Lists/Issues/Attachments/7/a.txt"}),
        )
        self.error(
            "unsupported_operation",
            lambda: self.api.library_for({"path": "/sites/test/Docs-other/a.txt"}),
        )

    def test_attachment_recycle_invalidates_underlying_file_reference(self):
        file = self.api.finish(self.api.ref({"id": GUID, "web_url": SITE + "/evidence"}, "file"))
        self.api.list_guard = Mock()
        self.api.record = Mock(return_value={"stable_id": "parent"})
        self.api.attachment = Mock(
            return_value={
                "id": GUID,
                "file_id": GUID,
                "list_id": OTHER,
                "item_id": "7",
                "web_url": SITE + "/evidence",
            }
        )
        self.session.request = Mock(return_value={"value": OTHER})
        self.api.finish(self.api.attachment_delete(OTHER, "7", "evidence.txt", allow=True))
        self.assertIsNone(get(file["stable_id"], ACCOUNT)["web_url"])
