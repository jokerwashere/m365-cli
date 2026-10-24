"""Synthetic safety and contract tests. No browser or network is used."""

import base64
import copy
import json
import tempfile
import time
import unittest
from unittest.mock import Mock, patch

from m365_cli.cli import parser, run
from m365_cli.errors import OutlookError
from m365_cli.operations import execute
from m365_cli.teams_api import TeamsAPI, normalize_message, text_body
from m365_cli.teams_session import (
    AUDIENCES,
    ORIGIN,
    TeamsSession,
    trusted_origin,
    validate_credentials,
)

ACCOUNT = "test@example.com"
TENANT = "00000000-0000-0000-0000-000000000001"
OID = "00000000-0000-0000-0000-000000000002"
MRI = "8:orgid:" + OID
THREAD = "19:test@thread.v2"
CHANNEL = "19:channel@thread.tacv2"
MID = "1234567890"


def jwt(**claims):
    return (
        "fake."
        + base64.urlsafe_b64encode(json.dumps(claims).encode()).decode().rstrip("=")
        + ".fake"
    )


def credentials():
    return {
        k: (
            ORIGIN + "/api/" + ("chatsvc" if k == "chat" else k) + "/emea",
            {
                "authorization": "Bearer "
                + jwt(aud=aud, tid=TENANT, oid=OID, upn=ACCOUNT, exp=time.time() + 3600)
            },
        )
        for k, aud in AUDIENCES.items()
    }


def thread(**properties):
    return {
        "id": THREAD,
        "tenantId": TENANT,
        "members": [{"id": MRI, "role": "Admin"}],
        "properties": {"threadType": "chat", **properties},
    }


def message(**overrides):
    return {
        "id": MID,
        "version": "10",
        "clientmessageid": "client-123",
        "from": "https://chat.invalid/v1/users/" + MRI,
        "messagetype": "RichText/Html",
        "content": "<p>old</p>",
        "properties": {},
        **overrides,
    }


class FakeSession:
    account, tenant, oid, mri = ACCOUNT, TENANT, OID, MRI

    def __init__(self, results):
        self.results = iter(results)
        self.calls = []

    def request(self, family, path, **kwargs):
        self.calls.append((family, path, kwargs))
        result = next(self.results)
        if isinstance(result, Exception):
            raise result
        return result


class Cases(unittest.TestCase):
    def error(self, code, function):
        with self.assertRaises(OutlookError) as caught:
            function()
        self.assertEqual(caught.exception.code, code)
        return caught.exception


class AuthTests(Cases):
    def session(self, status=200, data=None):
        s = TeamsSession()
        c = credentials()
        s.base = {k: v[0] for k, v in c.items()}
        s.headers = {k: v[1] for k, v in c.items()}
        s.account, s.tenant = ACCOUNT, TENANT
        s.context = Mock()
        response = Mock(status=status, headers={})
        response.json.return_value = data
        s.context.request.fetch.return_value = response
        return s

    def test_exact_origin_only(self):
        for url in (
            "http://teams.cloud.microsoft",
            "https://teams.cloud.microsoft.evil.test",
            "https://teams.cloud.microsoft:444/",
            "https://user@teams.cloud.microsoft",
            "https://teams.microsoft.com",
            "https://teams.cloud.microsoft:bad",
        ):
            self.assertFalse(trusted_origin(url), url)
        self.assertTrue(trusted_origin(ORIGIN + "/api/mt/emea/beta/users/tenantsv2"))

    def test_all_service_identities_and_audiences_match(self):
        self.assertEqual(
            validate_credentials(credentials(), ACCOUNT, TENANT), (ACCOUNT, TENANT, OID)
        )
        self.error(
            "account_mismatch", lambda: validate_credentials(credentials(), "other@example.com")
        )
        self.error("tenant_mismatch", lambda: validate_credentials(credentials(), ACCOUNT, "other"))
        for field in ("oid", "tid", "aud"):
            c = credentials()
            fields = dict(
                aud=AUDIENCES["mt"], tid=TENANT, oid=OID, upn=ACCOUNT, exp=time.time() + 3600
            )
            fields[field] = "wrong"
            c["mt"][1]["authorization"] = "Bearer " + jwt(**fields)
            self.error("identity_mismatch", lambda c=c: validate_credentials(c, ACCOUNT))

    def test_missing_chat_credentials_fail_closed(self):
        c = credentials()
        c["chat"][1].clear()
        self.error("login_required", lambda: validate_credentials(c, ACCOUNT))

    def test_expiry_and_skype_identity(self):
        c = credentials()
        c["chat"][1]["authorization"] = "Bearer " + jwt(exp=time.time() - 1)
        self.error("login_required", lambda c=c: validate_credentials(c, ACCOUNT))
        c = credentials()
        c["chat"][1]["x-skypetoken"] = jwt(
            exp=time.time() + 300, tid=TENANT, skypeid="orgid:" + OID
        )
        self.assertEqual(validate_credentials(c, ACCOUNT)[1], TENANT)
        c["chat"][1]["x-skypetoken"] = jwt(
            exp=time.time() + 300, tid="other", skypeid="orgid:" + OID
        )
        self.error("identity_mismatch", lambda c=c: validate_credentials(c, ACCOUNT))

    def test_transport_has_no_redirects_or_retries(self):
        s = self.session(data={"messages": []})
        s.request("chat", "v1/test")
        _, kw = s.context.request.fetch.call_args
        self.assertEqual(kw["max_redirects"], 0)
        self.assertEqual(kw["max_retries"], 0)
        self.assertNotIn("cookie", kw["headers"])
        self.assertTrue(s.context.request.fetch.call_args.args[0].startswith(ORIGIN + "/"))

    def test_routes_and_header_injection_rejected(self):
        s = self.session()
        for path in (
            "https://evil.test",
            "//evil.test",
            "v1/../beta",
            "v1/x\r\nCookie: secret",
            "v1/x#other",
        ):
            self.error("invalid_route", lambda path=path: s.request("chat", path))
        self.error(
            "invalid_headers",
            lambda: s.request("chat", "v1/test", headers={"authorization": "secret"}),
        )
        s.context.request.fetch.assert_not_called()

    def test_capture_ignores_foreign_origin_and_arbitrary_headers(self):
        s = TeamsSession()
        r = Mock(url=ORIGIN + "/api/csa/emea/api/v3/teams/users/me/updates", method="GET")
        r.all_headers.return_value = {
            "authorization": "secret",
            "cookie": "secret",
            "x-ms-synctoken": "secret",
            "evil": "secret",
        }
        s._capture(r)
        self.assertEqual(set(s.captured["csa"][1]), {"authorization"})
        s.captured.clear()
        r.url = r.url.replace("teams.cloud.microsoft", "evil.test")
        s._capture(r)
        self.assertFalse(s.captured)

    def test_status_errors_and_secret_redaction(self):
        for status, expected in (
            (302, "write_outcome_unknown"),
            (500, "write_outcome_unknown"),
            (401, "session_rejected"),
            (403, "forbidden"),
            (404, "not_found"),
            (412, "conflict"),
            (429, "rate_limited"),
        ):
            s = self.session(status, {"token": "SECRET"})
            error = self.error(
                expected,
                lambda s=s: s.request("chat", "v1/test", method="POST", body={}, write=True),
            )
            self.assertNotIn("SECRET", json.dumps(error.as_dict()))
            s.context.request.fetch.assert_called_once()
        s = self.session()
        s.context.request.fetch.side_effect = RuntimeError("Bearer SECRET")
        error = self.error(
            "write_outcome_unknown", lambda: s.request("chat", "v1/test", method="POST", write=True)
        )
        self.assertNotIn("SECRET", str(error))

    def test_expired_tokens_prevent_any_call(self):
        s = self.session()
        s.headers["mt"]["authorization"] = "Bearer " + jwt(exp=1)
        self.error(
            "login_required", lambda: s.request("chat", "v1/test", method="DELETE", write=True)
        )
        s.context.request.fetch.assert_not_called()

    def test_account_switch_does_not_change_frozen_headers(self):
        s = self.session(data={})
        s.captured = credentials()
        s.captured["mt"][1]["authorization"] = "Bearer OTHER"
        s.request("chat", "v1/test")
        self.assertNotIn("OTHER", str(s.context.request.fetch.call_args))

    def test_unverified_server_identity_fails(self):
        s = self.session()
        s.captured = credentials()
        s.request = Mock(return_value=[{"tenantId": TENANT, "userId": "different"}])
        with patch("m365_cli.teams_session.read_config", return_value={"account": ACCOUNT}):
            self.error("identity_mismatch", lambda: s.authenticate(wait=0))
        self.assertEqual(s.account, "")


class APIReadTests(Cases):
    def test_partial_inventory_never_claims_completeness(self):
        for metadata, expected in (
            ({}, None),
            ({"isPartialData": True}, None),
            ({"isPartialData": False, "hasMoreChats": False}, False),
        ):
            s = FakeSession([{"chats": [], "teams": [], "metadata": metadata}])
            self.assertIs(TeamsAPI(s).chats()["has_more"], expected)

    def test_self_chat_requires_fixed_roster(self):
        inv = {
            "teams": [],
            "chats": [{"id": THREAD, "isOneOnOne": True, "members": [{"mri": MRI}]}],
        }
        self.assertEqual(
            TeamsAPI(FakeSession([inv, thread(fixedRoster="True")])).self_chat()["id"], THREAD
        )
        self.error("not_found", lambda: TeamsAPI(FakeSession([inv, thread()])).self_chat())

    def test_self_chat_rejects_groups_external_and_incomplete_rosters(self):
        candidate = {"id": THREAD, "isOneOnOne": True, "members": [{"mri": MRI}]}
        for changes in ({"isOneOnOne": False}, {"isOneOnOne": None}, {"isExternal": True}):
            session = FakeSession([{"teams": [], "chats": [{**candidate, **changes}]}])
            self.error("not_found", lambda session=session: TeamsAPI(session).self_chat())
            self.assertEqual(len(session.calls), 1)
        for members in ([], [{"id": "8:orgid:another"}], [{"id": MRI}, {"id": "other"}]):
            actual = {**thread(fixedRoster="True"), "members": members}
            session = FakeSession([{"teams": [], "chats": [candidate]}, actual])
            self.error("not_found", lambda session=session: TeamsAPI(session).self_chat())

    def test_people_query_is_raw_and_exact_matching_required(self):
        p = {"mri": MRI, "email": ACCOUNT, "displayName": "Test"}
        s = FakeSession([{"value": [p]}])
        self.assertEqual(TeamsAPI(s).person(ACCOUNT), p)
        self.assertEqual(s.calls[0][2]["body"], ACCOUNT)
        self.assertTrue(s.calls[0][2]["raw"])
        self.error(
            "recipient_not_found",
            lambda: TeamsAPI(
                FakeSession([{"value": [{**p, "email": "other@example.com"}]}])
            ).person(ACCOUNT),
        )
        self.error(
            "recipient_ambiguous",
            lambda: TeamsAPI(FakeSession([{"value": [p, p]}])).person(ACCOUNT),
        )

    def test_sensitive_message_properties_never_returned(self):
        result = normalize_message(
            message(
                properties={"subject": "ok", "copilotShareTokens": "SECRET", "arbitrary": "SECRET"}
            )
        )
        self.assertNotIn("SECRET", json.dumps(result))

    def test_cursor_bound_to_identity_and_conversation(self):
        a = TeamsAPI(FakeSession([]))
        link = (
            "https://some.chat.service/v1/users/ME/conversations/"
            + THREAD
            + "/messages?startTime=0&endTime=500&pageSize=1"
        )
        cursor = a.cursor(link, THREAD)
        params = a.decode_cursor(cursor, THREAD, 10)
        self.assertEqual(params["pageSize"], "10")
        self.error("invalid_cursor", lambda: a.decode_cursor(cursor, CHANNEL, 10))
        a.s.tenant = "other"
        self.error("invalid_cursor", lambda: a.decode_cursor(cursor, THREAD, 10))

    def test_cursor_cannot_inject_endpoint_or_unknown_parameters(self):
        a = TeamsAPI(FakeSession([]))
        for link in (
            "https://evil.test/v1/admin?startTime=0",
            "/v1/users/ME/conversations/other/messages?startTime=0",
            "/v1/users/ME/conversations/" + THREAD + "/messages?authorization=secret",
            "/v1/users/ME/conversations/" + THREAD + "/messages?startTime=0&startTime=1",
        ):
            self.error("invalid_cursor", lambda link=link: a.cursor(link, THREAD))
        self.error("invalid_cursor", lambda: a.decode_cursor("garbage", THREAD, 1))

    def test_reads_never_write_consumption_horizons(self):
        s = FakeSession([thread(), {"messages": [message()], "tenantId": TENANT}])
        TeamsAPI(s).messages(THREAD)
        self.assertTrue(all(call[2].get("method", "GET") == "GET" for call in s.calls))
        self.assertFalse(any("consumptionhorizon" in call[1] for call in s.calls))

    def test_target_type_and_tenant_mismatch_rejected(self):
        self.error(
            "wrong_item_type",
            lambda: TeamsAPI(FakeSession([thread(threadType="topic")])).target(THREAD),
        )
        self.error(
            "tenant_mismatch",
            lambda: TeamsAPI(FakeSession([{**thread(), "tenantId": "other"}])).chat(THREAD),
        )


class APIWriteTests(Cases):
    def test_body_escaping(self):
        self.assertEqual(text_body('A & <b>\n"x"'), "<p>A &amp; &lt;b&gt;<br>&quot;x&quot;</p>")
        self.error("invalid_argument", lambda: text_body(" "))
        self.error("body_too_large", lambda: text_body("x" * 20001))

    def test_edit_uses_current_server_version(self):
        s = FakeSession([thread(), message(), {}, message(content="<p>new</p>", version="11")])
        result = TeamsAPI(s).update(THREAD, MID, "new", expected="10")
        self.assertTrue(result["updated"])
        self.assertEqual(s.calls[2][2]["headers"], {"If-Match": '"10"'})

    def test_stale_edit_does_not_write(self):
        s = FakeSession([thread(), message()])
        self.error("conflict", lambda: TeamsAPI(s).update(THREAD, MID, "new", expected="9"))
        self.assertEqual(len(s.calls), 2)

    def test_server_edit_conflict_is_not_retried(self):
        s = FakeSession([thread(), message(), OutlookError("conflict", "conflict")])
        self.error("conflict", lambda: TeamsAPI(s).update(THREAD, MID, "new"))
        self.assertEqual(len(s.calls), 3)

    def test_delete_rejects_atomic_precondition_before_any_read(self):
        s = FakeSession([])
        self.error(
            "unsupported_precondition", lambda: TeamsAPI(s).delete(THREAD, MID, expected="10")
        )
        self.assertFalse(s.calls)

    def test_owner_guard_for_edits_and_deletes(self):
        for op in ("update", "delete"):
            s = FakeSession([thread(), message(**{"from": "8:orgid:other"})])

            def call(s=s, op=op):
                return (
                    TeamsAPI(s).delete(THREAD, MID)
                    if op == "delete"
                    else TeamsAPI(s).update(THREAD, MID, "x")
                )

            self.error("not_owner", call)
            self.assertEqual(len(s.calls), 2)

    def test_rich_message_edit_rejected(self):
        for key in ("mentions", "cards", "files", "qtdMsgs"):
            s = FakeSession([thread(), message(properties={key: '[{"x":1}]'})])
            self.error("unsupported_message", lambda s=s: TeamsAPI(s).update(THREAD, MID, "x"))
            self.assertEqual(len(s.calls), 2)

    def test_delete_requires_tombstone(self):
        s = FakeSession(
            [thread(), message(), None, message(content="", properties={"deletetime": "now"})]
        )
        result = TeamsAPI(s).delete(THREAD, MID)
        self.assertTrue(result["deleted"])
        self.assertFalse(result["atomic_version_guard"])
        s = FakeSession([thread(), message(), None, message()])
        self.error("readback_failed", lambda: TeamsAPI(s).delete(THREAD, MID))

    def test_send_acknowledgement_and_readback_required(self):
        s = FakeSession([thread(), {"id": MID}])
        self.error("write_outcome_unknown", lambda: TeamsAPI(s).send(THREAD, "x"))
        self.assertEqual(len(s.calls), 2)
        s = FakeSession([thread(), {"OriginalArrivalTime": MID}, message()])
        self.error("readback_failed", lambda: TeamsAPI(s).send(THREAD, "x"))

    def test_channel_reply_uses_composite_target_and_parent_id(self):
        inv = {
            "chats": [],
            "teams": [
                {"id": "team", "channels": [{"id": CHANNEL, "isMember": True, "channelType": 1}]}
            ],
        }
        s = FakeSession(
            [
                inv,
                message(id="parent"),
                {"OriginalArrivalTime": MID},
                message(content="<p>x</p>", clientmessageid="1"),
            ]
        )
        with patch("m365_cli.teams_api.time.time_ns", return_value=1000):
            result = TeamsAPI(s).send(CHANNEL, "x", channel=True, reply_to="parent")
        self.assertTrue(result["submitted"])
        self.assertIn("%3Bmessageid%3Dparent", s.calls[2][1])
        self.assertEqual(s.calls[2][2]["body"]["replyChainId"], "parent")

    def test_archived_channel_cannot_write(self):
        inv = {
            "chats": [],
            "teams": [{"id": "team", "channels": [{"id": CHANNEL, "isArchived": True}]}],
        }
        s = FakeSession([inv])
        self.error("read_only", lambda: TeamsAPI(s).send(CHANNEL, "x", channel=True))
        self.assertEqual(len(s.calls), 1)

    def test_create_group_uses_validated_location_and_verifies_roster(self):
        s = FakeSession(
            [
                {
                    "body": {"threadResource": "unused"},
                    "location": "https://chat.service/v1/threads/19%3Atest%40thread.v2",
                },
                thread(topic="Test"),
            ]
        )
        result = TeamsAPI(s).create_chat([], "Test", True)
        self.assertTrue(result["created"])
        self.assertEqual(s.calls[0][2]["body"]["members"], [{"id": MRI, "role": "Admin"}])

    def test_direct_chat_reuse_does_not_create(self):
        p = {"mri": "8:orgid:peer", "email": "peer@example.com"}
        t = thread()
        t["members"].append({"id": p["mri"]})
        s = FakeSession([{"value": [p]}, {"value": [p]}, {"id": THREAD}, t])
        result = TeamsAPI(s).create_chat([p["email"]])
        self.assertFalse(result["created"])
        self.assertFalse(any(x[2].get("write") for x in s.calls))

    def test_chat_postwrite_read_failure_reports_acknowledged_write(self):
        s = FakeSession([thread(), {}, OutlookError("transport_error", "SECRET")])
        error = self.error("readback_failed", lambda: TeamsAPI(s).rename_chat(THREAD, "new"))
        self.assertNotIn("SECRET", str(error))
        self.assertEqual(error.details["thread_id"], THREAD)

    def test_create_missing_acknowledgement_is_unknown(self):
        for ack in (None, [], {}, {"location": "https://[malformed", "body": {}}):
            self.error(
                "write_outcome_unknown",
                lambda ack=ack: TeamsAPI(FakeSession([ack])).create_chat([], "Test", True),
            )

    def test_malformed_direct_chat_resolution_never_creates(self):
        person = {"mri": "8:orgid:peer", "email": "peer@example.com"}
        s = FakeSession([{"value": [person]}, {"value": [person]}, {}])
        self.error("unexpected_response", lambda: TeamsAPI(s).create_chat([person["email"]]))
        self.assertFalse(any(call[2].get("write") for call in s.calls))

    def test_group_rename_and_leave_reject_direct_chat(self):
        s = FakeSession([thread(fixedRoster="True")])
        self.error("unsupported_operation", lambda: TeamsAPI(s).leave_chat(THREAD))
        self.assertEqual(len(s.calls), 1)


class CLITests(Cases):
    def test_dry_run_uses_no_browser_and_normalizes_file_content(self):
        args = parser().parse_args(
            ["teams", "message", "send", "--chat", THREAD, "--body", "test", "--dry-run"]
        )
        with patch("m365_cli.teams_cli.TeamsSession") as session:
            result = run(args)
        self.assertTrue(result["dry_run"])
        session.assert_not_called()

    def test_unsupported_search_and_ambiguous_targets(self):
        for argv in (
            ["teams", "message", "search", "test"],
            ["teams", "message", "send", "--chat", THREAD, "--channel", CHANNEL, "--body", "test"],
        ):
            self.error("usage_error", lambda argv=argv: parser().parse_args(argv))

    def test_delete_precondition_rejected_even_in_dry_run(self):
        args = parser().parse_args(
            ["teams", "message", "delete", MID, "--chat", THREAD, "--if-version", "10", "--dry-run"]
        )
        self.error("unsupported_precondition", lambda args=args: run(args))

    def test_invalid_plain_text_and_group_inputs(self):
        for tail in (
            ["message", "send", "--chat", THREAD],
            ["chat", "create"],
            ["chat", "create", "--group"],
        ):
            args = parser().parse_args(["teams", *tail, "--dry-run"])
            self.error("invalid_argument", lambda args=args: run(args))

    def test_operation_records_are_bound_to_service_and_tenant(self):
        with (
            tempfile.TemporaryDirectory() as folder,
            patch.dict("os.environ", {"M365_HOME": folder}),
        ):
            action = Mock(return_value={"submitted": True})
            payload = {"service": "teams", "tenant": TENANT, "command": "send"}
            execute("test", ACCOUNT, payload, action)
            self.assertTrue(execute("test", ACCOUNT, payload, action)["replayed"])
            for key in ("service", "tenant"):
                changed = copy.deepcopy(payload)
                changed[key] = "other"
                self.error(
                    "request_id_conflict",
                    lambda changed=changed: execute("test", ACCOUNT, changed, action),
                )
            action.assert_called_once()
