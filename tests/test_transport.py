"""Synthetic bootstrap isolation, bounded HTTP, scoped cookies and native downloads."""

import base64
import io
import json
import subprocess
import time
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import playwright

from m365_cli.errors import OutlookError
from m365_cli.http_transport import BrowserCookies, DirectHTTP, Response
from m365_cli.session import Session

ORIGIN = "https://outlook.office.com"
ACCOUNT = "reader@example.com"


def reply(data=None, status=200, **headers):
    return Response(status, json.dumps(data).encode(), headers)


def token(**claims):
    value = {"exp": time.time() + 3600, "upn": ACCOUNT, "aud": ORIGIN, **claims}
    return (
        "Bearer x."
        + base64.urlsafe_b64encode(json.dumps(value).encode()).decode().rstrip("=")
        + ".synthetic"
    )


def ready_session(account=ACCOUNT):
    s = Session()
    s.account, s.origin, s.base = account, ORIGIN, "/owa/0/service.svc"
    s.headers = {"authorization": token(upn=account), "user-agent": "synthetic-browser"}
    s.http = Mock()
    return s


def cookie(name="session", value="synthetic", **kw):
    return {
        "name": name,
        "value": value,
        "domain": "example.sharepoint.com",
        "path": "/",
        "secure": True,
        "expires": -1,
        **kw,
    }


class Cases(unittest.TestCase):
    def error(self, code, call):
        with self.assertRaises(OutlookError) as caught:
            call()
        self.assertEqual(caught.exception.code, code)
        self.assertNotIn("SENSITIVE", json.dumps(caught.exception.as_dict()))
        return caught.exception


class HTTPTests(Cases):
    def connection(self, data=b"{}", status=200, **headers):
        response = Mock()
        response.status = status
        response.will_close = False
        response.isclosed.return_value = False
        response.getheader.side_effect = lambda key, default=None: headers.get(key, default)
        stream = io.BytesIO(data)
        response.read1.side_effect = stream.read1
        conn = Mock()
        conn.getresponse.return_value = response
        p = patch("m365_cli.http_transport.http.client.HTTPSConnection", return_value=conn)
        factory = p.start()
        self.addCleanup(p.stop)
        return conn, response, factory

    def test_fixed_origin_tls_no_ambient_credentials_and_raw_bytes(self):
        conn, response, factory = self.connection(b"\0\xffabc", **{"Content-Length": "5"})
        client = DirectHTTP(ORIGIN)
        result = client.request(ORIGIN + "/api?x=1", headers={"Authorization": "SENSITIVE"})
        self.assertEqual(result.content, b"\0\xffabc")
        factory.assert_called_once_with("outlook.office.com", timeout=45)
        args = conn.request.call_args
        self.assertEqual(args.args, ("GET", "/api?x=1"))
        self.assertEqual(
            args.kwargs["headers"], {"Accept-Encoding": "identity", "Authorization": "SENSITIVE"}
        )
        self.assertNotIn("SENSITIVE", repr(result))
        client.close()
        conn.close.assert_called_once()

    def test_routes_and_header_injection_rejected_before_connect(self):
        _, _, factory = self.connection()
        for url in (
            "http://outlook.office.com/a",
            "https://outlook.office.com.evil.test/a",
            "https://other.test/a",
            ORIGIN + "/a#fragment",
            ORIGIN + "/a\r\nb",
            "https://user@outlook.office.com/a",
            ORIGIN + "/a\\b",
        ):
            self.error("invalid_route", lambda url=url: DirectHTTP(ORIGIN).request(url))
        for headers in (
            {"Authorization": "SENSITIVE\r\nHost: evil.test"},
            {"Host": "evil.test"},
            {"Content-Length": "0"},
            {"bad\nname": "x"},
        ):
            self.error(
                "invalid_headers",
                lambda headers=headers: DirectHTTP(ORIGIN).request(ORIGIN, headers=headers),
            )
        factory.assert_not_called()

    def test_redirect_never_followed_even_to_same_origin(self):
        for write in (False, True):
            with self.subTest(write=write):
                conn, response, factory = self.connection(
                    status=302, Location=ORIGIN + "/login?SENSITIVE"
                )
                self.error(
                    "write_outcome_unknown" if write else "redirect_rejected",
                    lambda write=write: DirectHTTP(ORIGIN).request(ORIGIN + "/api", write=write),
                )
                conn.request.assert_called_once()
                response.read1.assert_not_called()
                factory.assert_called_once()

    def test_declared_and_streamed_limits_close_without_retry(self):
        for declared in (None, "100"):
            for write in (False, True):
                conn, response, _ = self.connection(
                    b"123456", **({"Content-Length": declared} if declared else {})
                )
                self.error(
                    "write_outcome_unknown" if write else "response_too_large",
                    lambda write=write: DirectHTTP(ORIGIN).request(
                        ORIGIN + "/api", maximum=5, write=write
                    ),
                )
                conn.close.assert_called_once()
                conn.request.assert_called_once()
                if declared:
                    response.read1.assert_not_called()

    def test_incomplete_encoding_timeout_and_reset_are_conservative(self):
        conn, _, _ = self.connection(b"a", **{"Content-Length": "2"})
        self.error("transport_error", lambda: DirectHTTP(ORIGIN).request(ORIGIN + "/api"))
        conn, _, _ = self.connection(**{"Content-Encoding": "gzip"})
        self.error("api_changed", lambda: DirectHTTP(ORIGIN).request(ORIGIN + "/api"))
        for failure in (TimeoutError("SENSITIVE"), ConnectionResetError("SENSITIVE")):
            conn, _, _ = self.connection()
            conn.request.side_effect = failure
            self.error(
                "write_outcome_unknown",
                lambda: DirectHTTP(ORIGIN).request(ORIGIN + "/api", write=True),
            )
            conn.request.assert_called_once()

    def test_total_read_deadline(self):
        conn, _, _ = self.connection(b"123")
        with patch("m365_cli.http_transport.time.monotonic", side_effect=[0, 46]):
            self.error("transport_error", lambda: DirectHTTP(ORIGIN).request(ORIGIN + "/api"))
        conn.close.assert_called_once()


class CookieTests(Cases):
    def test_partitioned_extra_not_copied_but_actual_applicable_cookie_kept(self):
        url = "https://example.sharepoint.com/sites/test/_api/web"
        scoped = [cookie(), cookie("extra", "secret", partitionKey="https://other.test")]
        jar = BrowserCookies("session=synthetic", scoped, url)
        self.assertEqual(jar.header(url), "session=synthetic")
        jar = BrowserCookies("session=synthetic; extra=secret", scoped, url)
        self.assertEqual(jar.header(url), "session=synthetic; extra=secret")
        jar.clear()
        self.assertEqual(jar.header(url), "")

    def test_host_path_expiry_scope_and_missing_metadata(self):
        url = "https://example.sharepoint.com/sites/a/_api/web"
        jar = BrowserCookies("session=synthetic", [cookie(path="/sites/a")], url)
        for other in (
            "https://example.sharepoint.com/sites/ab/_api",
            "https://evil.test/sites/a/_api",
            "http://example.sharepoint.com/sites/a/_api",
        ):
            self.assertEqual(jar.header(other), "")
        with patch("m365_cli.http_transport.time.time", return_value=2):
            expiring = BrowserCookies("session=synthetic", [cookie(expires=3)], url)
        with patch("m365_cli.http_transport.time.time", return_value=4):
            self.assertEqual(expiring.header(url), "")
        self.error("login_required", lambda: BrowserCookies("unknown=SENSITIVE", [cookie()], url))
        self.error(
            "login_required",
            lambda: BrowserCookies("session=synthetic\r\nSENSITIVE", [cookie()], url),
        )


class OutlookHTTPTests(Cases):
    def test_reads_have_no_browser_dependency_and_preserve_wrappers(self):
        s = ready_session()
        s.http.request.return_value = reply({"Body": {"ResponseCode": "NoError"}})
        s.request("GetItem", {"sample": True}, time_zone="Europe/Zurich")
        args = s.http.request.call_args
        self.assertEqual(args.args[0], ORIGIN + "/owa/0/service.svc?action=GetItem&app=Mail&n=0")
        body = json.loads(args.kwargs["body"])
        self.assertEqual(body["Body"], {"sample": True})
        self.assertEqual(
            body["Header"]["TimeZoneContext"]["TimeZoneDefinition"]["Id"], "Europe/Zurich"
        )
        self.assertEqual(args.kwargs["headers"]["prefer"], 'IdType="ImmutableId"')
        self.assertFalse(s.browser_connected)
        self.assertIsNone(s.page)

    def test_expiry_and_claim_switch_stop_before_any_request(self):
        for claims, code in (
            ({"exp": time.time() - 1}, "login_required"),
            ({"upn": "other@example.com"}, "account_mismatch"),
        ):
            s = ready_session()
            s.headers["authorization"] = token(**claims)
            self.error(code, lambda s=s: s.request("CreateItem", {}))
            s.http.request.assert_not_called()

    def test_server_identity_guard_and_no_automatic_mutation_retry(self):
        s = ready_session()
        s.http.request.return_value = reply({"UserEmailAddress": "different@example.com"})
        self.error("account_mismatch", lambda: s.request("CreateItem", {}))
        self.assertEqual(s.http.request.call_count, 1)
        s.http.request.reset_mock()
        s.http.request.side_effect = [
            reply({"UserEmailAddress": ACCOUNT}),
            OutlookError("write_outcome_unknown", "Uncertain write"),
        ]
        self.error("write_outcome_unknown", lambda: s.request("CreateItem", {}))
        self.assertEqual(s.http.request.call_count, 2)
        self.assertTrue(s.http.request.call_args.kwargs["write"])

    def test_malformed_mutation_ack_and_throttling_preserve_unknown_state(self):
        s = ready_session()
        for data in (None, [], {}):
            s.http.request.side_effect = [reply({"UserEmailAddress": ACCOUNT}), reply(data)]
            self.error("write_outcome_unknown", lambda: s.request("UpdateItem", {}))
        s.http.request.side_effect = None
        s.http.request.return_value = reply({}, status=429, **{"retry-after": "12"})
        self.error("rate_limited", lambda: s.request("FindItem", {}))

    def test_unsafe_search_routes_and_native_actions_rejected(self):
        s = ready_session()
        for path in (
            "https://other.test/query",
            "/searchservice/api/v2/query?leak=1",
            "//other.test/api",
        ):
            self.error("invalid_route", lambda path=path: s.request(None, {}, path=path))
        self.error("invalid_argument", lambda: s.request("CreateItem", {}, native=True))
        s.http.request.assert_not_called()

    def test_download_uses_only_dedicated_origin_and_token(self):
        s = ready_session()
        s.request = Mock(
            return_value={
                "UserEmailAddress": ACCOUNT,
                "DownloadUrlBase": "https://attachments.office.net/owa/",
                "FirstDownloadToken": "SENSITIVE",
            }
        )
        client = Mock()
        client.request.return_value = Response(
            200,
            b"\0\xff",
            {
                "content-disposition": 'attachment; filename="a.eml"',
                "content-type": "message/rfc822",
            },
        )
        with patch("m365_cli.session.DirectHTTP", return_value=client) as factory:
            result = s.download_attachment("opaque?&/+")
        factory.assert_called_once_with("https://attachments.office.net")
        args = client.request.call_args
        self.assertIn(
            "reader%40example.com/service.svc/s/GetFileAttachment?id=opaque%3F%26%2F%2B",
            args.args[0],
        )
        self.assertEqual(
            args.kwargs["headers"],
            {
                "X-Token": "SENSITIVE",
                "X-AnchorMailbox": ACCOUNT,
                "User-Agent": "synthetic-browser",
                "Accept": "*/*",
            },
        )
        self.assertEqual(result["content"], b"\0\xff")
        self.assertNotIn("SENSITIVE", repr(result))
        client.request.side_effect = OutlookError("response_too_large", "large")
        with patch("m365_cli.session.DirectHTTP", return_value=client):
            self.error("attachment_too_large", lambda: s.download_attachment("id"))
        s.request.return_value["DownloadUrlBase"] = "https://attachments.office.net.evil.test/owa/"
        with patch("m365_cli.session.DirectHTTP") as factory:
            self.error("api_changed", lambda: s.download_attachment("id"))
            factory.assert_not_called()

    def test_disconnect_and_exit_clear_secrets(self):
        s = ready_session()
        pw = s.pw = Mock()
        s.context = s.page = Mock()
        s.cookies = Mock()
        s.captured[ORIGIN] = {"authorization": "SENSITIVE"}
        s.__exit__()
        pw.stop.assert_called_once()
        self.assertIsNone(s.page)
        self.assertIsNone(s.http)
        self.assertEqual(s.headers, {})
        self.assertEqual(s.captured, {})
        s.cookies.clear.assert_called_once()


class BootstrapTests(Cases):
    def test_browser_javascript_only_status_or_read_bootstrap(self):
        node = Path(playwright.__file__).parent / "driver" / "node"
        source = Path(__file__).parents[1] / "m365_cli" / "session.js"
        script = r"""
const vm=require('vm'),fs=require('fs'),assert=require('assert');
const source=fs.readFileSync(process.argv[1],'utf8');
let calls=[];const jwt=c=>'x.'+Buffer.from(JSON.stringify(c)).toString('base64url')+'.fake';
const token=jwt({exp:Date.now()/1000+3600,upn:'a@example.com',aud:'https://outlook.office.com'});
const ctx={location:{hostname:'outlook.office.com',protocol:'https:',port:''},window:{},document:{cookie:''},
 localStorage:{token:JSON.stringify({secret:token}),getItem(k){return this[k]}},sessionStorage:{},Date,Uint8Array,TextDecoder,
 atob:s=>Buffer.from(s,'base64').toString(),crypto:{randomUUID:()=> 'uuid'},AbortSignal:{timeout:()=>null},
 fetch:async(url,args)=>{calls.push({url,args});let done=false;return {status:200,headers:{get:()=>null},body:{getReader:()=>({read:async()=>done?{done:true}:(done=true,{done:false,value:Buffer.from('{"Body":{"ResponseCode":"NoError"}}')})})}}}};
vm.createContext(ctx);const run=vm.runInContext('('+source+')',ctx);
(async()=>{
 let value=await run({op:'status',headers:{},account:'a@example.com'});assert(value.authenticated);assert(!JSON.stringify(value).includes(token));
 for(const op of ['request','attachment_download'])await assert.rejects(()=>run({op,headers:{},account:'a@example.com'}));
 assert.equal(calls.length,0);
 value=await run({op:'bootstrap',headers:{},account:'b@example.com'});assert.equal(value.status,409);assert.equal(calls.length,0);
 value=await run({op:'bootstrap',headers:{},account:'a@example.com'});assert.equal(value.status,200);assert.equal(calls.length,1);
 assert(calls[0].url.includes('action=GetOwaUserConfiguration'));assert.equal(calls[0].args.redirect,'error');
 assert.equal(calls[0].args.headers['x-anchormailbox'],'SMTP:a@example.com');assert(!JSON.stringify(value).includes(token));
 ctx.location.hostname='evil.test';assert.equal((await run({op:'bootstrap',headers:{}})).authenticated,false);assert.equal(calls.length,1);
 process.stdout.write('passed');
})().catch(e=>{process.stderr.write(e.stack);process.exit(1)});
"""
        result = subprocess.run(
            [str(node), "-e", script, str(source)], capture_output=True, text=True, timeout=15
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(result.stdout, "passed")


class LiveFixtureTests(Cases):
    def test_fixture_preparation_only_reads_until_three_equal_keys(self):
        from m365_cli.live_outlook import settled_fixture

        read = Mock(
            side_effect=[{"Subject": "test fixture", "change_key": k} for k in ("a", "b", "b", "b")]
        )
        with patch("m365_cli.live_outlook.time.sleep"):
            self.assertEqual(settled_fixture(read, "test fixture"), "b")
        self.assertEqual(read.call_count, 4)

    def test_unstable_or_unlabelled_fixture_stops_without_write(self):
        from m365_cli.live_outlook import settled_fixture

        read = Mock(
            side_effect=[{"Subject": "test fixture", "change_key": str(i)} for i in range(8)]
        )
        with patch("m365_cli.live_outlook.time.sleep"), self.assertRaises(RuntimeError):
            settled_fixture(read, "test fixture")
        self.assertEqual(read.call_count, 8)
        with self.assertRaises(RuntimeError):
            settled_fixture(lambda: {"Subject": "business data", "change_key": "a"}, "test fixture")


class BootstrapSnapshotTests(Cases):
    def test_outlook_snapshot_is_origin_scoped_and_frozen(self):
        s = Session()
        s.page = Mock(url=ORIGIN + "/mail/")
        s.context = Mock()
        url = ORIGIN + "/owa/0/service.svc?action=GetOwaUserConfiguration&app=Mail&n=0"
        initial = token()
        incoming = {
            "authorization": initial,
            "cookie": "session=synthetic",
            "user-agent": "browser",
            "x-owa-canary": "browser-canary",
            "x-owa-sessionid": "browser-session",
            "unrelated": "SENSITIVE",
        }
        request = Mock(url=url, method="POST")
        request.all_headers.side_effect = lambda: dict(incoming)
        s.context.cookies.return_value = [
            cookie(domain="outlook.office.com"),
            cookie(
                "extra", "no", domain="outlook.office.com", partitionKey="https://elsewhere.test"
            ),
        ]

        def evaluate(*_):
            s._capture(request)
            return {"status": 200, "body": {"UserEmailAddress": ACCOUNT}}

        s.page.evaluate.side_effect = evaluate
        s._bootstrap(ACCOUNT)
        self.assertEqual(s.headers["authorization"], initial)
        self.assertNotIn("unrelated", s.headers)
        self.assertNotIn("cookie", s.headers)
        self.assertNotIn("x-owa-canary", s.headers)
        self.assertNotIn("x-owa-sessionid", s.headers)
        self.assertEqual(s.cookies.header(url), "")
        s.context.cookies.assert_called_once_with([url])
        incoming["authorization"] = token(upn="other@example.com")
        s._capture(request)
        self.assertEqual(s.headers["authorization"], initial)
        request.url = "https://evil.test/owa/service.svc"
        s._capture(request)
        self.assertNotIn("https://evil.test", s.captured)

    def test_canary_only_bootstrap_keeps_applicable_cookies(self):
        s = Session()
        s.page = Mock(url=ORIGIN + "/mail/")
        s.context = Mock()
        s.context.cookies.side_effect = [
            [cookie(domain="outlook.office.com")],
            [cookie(value="rotated", domain="outlook.office.com")],
        ]

        def evaluate(*_):
            s._observed = {"x-owa-canary": "synthetic", "cookie": "session=synthetic"}
            return {"status": 200, "body": {"UserEmailAddress": ACCOUNT}}

        s.page.evaluate.side_effect = evaluate
        s._bootstrap(ACCOUNT)
        self.assertEqual(s.cookies.header(s.bootstrap_url), "session=synthetic")
        s._validate_credentials()

    def test_sharepoint_only_captures_verified_bootstrap_request(self):
        import tempfile

        from m365_cli.sharepoint_session import SharePointSession

        with tempfile.TemporaryDirectory() as tmp, patch.dict("os.environ", {"M365_HOME": tmp}):
            s = SharePointSession("https://example.sharepoint.com/sites/test", expected=ACCOUNT)
            s.tenant = s.web_id = "11111111-1111-1111-1111-111111111111"
            context = Mock()
            context.cookies.side_effect = [
                [cookie()],
                [
                    cookie(value="rotated"),
                    cookie("extra", "no", partitionKey="https://elsewhere.test"),
                ],
            ]
            s.page = Mock()

            def evaluate(_, args):
                capture = context.on.call_args.args[1]
                wrong = Mock(url="https://other.test/_api/web", method="GET")
                capture(wrong)
                wrong.all_headers.assert_not_called()
                request = Mock(url=args["url"], method="GET")
                request.all_headers.return_value = {
                    "cookie": "session=synthetic",
                    "user-agent": "browser",
                    "authorization": "SENSITIVE",
                }
                capture(request)
                return {"status": 200}

            s.page.evaluate.side_effect = evaluate
            s._bootstrap(context)
            self.assertEqual(s.cookies.header(s.site + "/_api/web"), "session=synthetic")
            self.assertEqual(s.user_agent, "browser")
            context.remove_listener.assert_called_once()
            self.assertFalse(hasattr(s, "authorization"))
            s.__exit__()
            self.assertIsNone(s.http)
            self.assertEqual(s.cookies.header(s.site + "/_api/web"), "")

    def test_maintainer_reports_detached_transport_and_detects_browser_dependency(self):
        from m365_cli.maintainer import check

        s = ready_session()
        s.authenticate = Mock()
        session = Mock()
        session.__enter__ = Mock(return_value=s)
        session.__exit__ = Mock(return_value=False)
        api = Mock()
        api.mail_list.return_value = api.folder_list.return_value = (
            api.calendar_list.return_value
        ) = {"items": []}
        with (
            patch("m365_cli.maintainer.Session", return_value=session),
            patch("m365_cli.api.API", return_value=api),
            patch("m365_cli.labels.Categories"),
            patch("m365_cli.scheduling.snapshot"),
        ):
            result = check("outlook", account=ACCOUNT)
            self.assertEqual(result["transports"]["outlook"], s.transport_info())
            s.page = Mock()
            result = check("outlook", account=ACCOUNT)
            row = next(c for c in result["checks"] if c["name"] == "outlook.transport.detached")
            self.assertEqual(row["status"], "failed")
            self.assertNotIn("outlook", result["transports"])


class MaintainerDriveTests(Cases):
    def test_drive_probe_rejects_missing_identity(self):
        from m365_cli.maintainer import check

        s = Mock()
        s.page = None
        s.transport_info.return_value = {
            "mode": "direct_http",
            "browser_connected": False,
            "credential_storage": "command_memory",
        }
        manager = Mock()
        manager.__enter__ = Mock(return_value=s)
        manager.__exit__ = Mock(return_value=False)
        api = Mock()
        api.lists.return_value = {"items": []}
        for data, status in (({}, "failed"), ({"id": "synthetic-drive"}, "passed")):
            s.request.return_value = data
            with (
                patch("m365_cli.sharepoint_session.SharePointSession", return_value=manager),
                patch("m365_cli.sharepoint_api.SharePointAPI", return_value=api),
            ):
                report = check(
                    "sharepoint", account=ACCOUNT, site="https://example.sharepoint.com/sites/test"
                )
            self.assertEqual(
                next(c["status"] for c in report["checks"] if c["name"] == "sharepoint.drive"),
                status,
            )


class ReadRetryTests(Cases):
    def test_declared_read_retries_connection_once_without_mutation(self):
        from m365_cli.http_transport import service_request

        client = Mock()
        client.request.side_effect = [
            OutlookError("transport_error", "reset", reason="connection_io"),
            reply({"ok": True}),
        ]
        with patch("m365_cli.http_transport.time.sleep"):
            result = service_request(client, ORIGIN + "/read", read=True)
        self.assertEqual(result.json(), {"ok": True})
        self.assertEqual(client.request.call_count, 2)
        self.assertLessEqual(client.request.call_args.kwargs["timeout"], 45)

    def test_writes_and_non_transport_errors_never_retry(self):
        from m365_cli.http_transport import service_request

        for read, code in (
            (False, "transport_error"),
            (False, "write_outcome_unknown"),
            (True, "login_required"),
            (True, "api_changed"),
            (True, "response_too_large"),
            (True, "redirect_rejected"),
        ):
            client = Mock()
            client.request.side_effect = OutlookError(code, "synthetic failure")
            self.error(
                code,
                lambda read=read, client=client: service_request(
                    client, ORIGIN + "/api", read=read
                ),
            )
            client.request.assert_called_once()

    def test_exhausted_read_retains_sanitized_cause_and_attempt_count(self):
        from m365_cli.http_transport import service_request

        client = Mock()
        client.request.side_effect = OutlookError(
            "transport_error",
            "connection failure",
            phase="response_headers",
            reason="connection_io",
            os_error=54,
        )
        with patch("m365_cli.http_transport.time.sleep"):
            error = self.error(
                "transport_error", lambda: service_request(client, ORIGIN + "/read", read=True)
            )
        self.assertEqual(error.details["read_attempts"], 2)
        self.assertEqual(error.details["os_error"], 54)
        self.assertEqual(client.request.call_count, 2)

    def test_unknown_outlook_action_is_not_retryable(self):
        s = ready_session()
        s.http.request.side_effect = [
            reply({"UserEmailAddress": ACCOUNT}),
            OutlookError("transport_error", "unknown action failed"),
        ]
        self.error("transport_error", lambda: s.request("NewUnknownAction", {}))
        self.assertEqual(s.http.request.call_count, 2)
        self.assertTrue(s.http.request.call_args.kwargs["write"])


class ConnectionReuseTests(Cases):
    connection = HTTPTests.connection

    def test_success_reuses_origin_connection_and_failed_write_is_not_replayed(self):
        conn, response, factory = self.connection()
        response.read1.side_effect = [b"{}", b"", b"{}", b""]
        client = DirectHTTP(ORIGIN)
        client.request(ORIGIN + "/first")
        client.request(ORIGIN + "/second")
        factory.assert_called_once()
        conn.connect.assert_called_once()
        self.assertEqual(conn.request.call_count, 2)
        conn.close.assert_not_called()
        conn.request.side_effect = ConnectionResetError("SENSITIVE")
        self.error(
            "write_outcome_unknown",
            lambda: client.request(ORIGIN + "/write", method="POST", body=b"{}", write=True),
        )
        self.assertEqual(conn.request.call_count, 3)
        conn.close.assert_called_once()
        self.assertIsNone(client._connection)

    def test_server_connection_close_is_respected_after_success(self):
        conn, response, _ = self.connection()
        response.will_close = True
        client = DirectHTTP(ORIGIN)
        client.request(ORIGIN + "/read")
        conn.close.assert_called_once()
        self.assertIsNone(client._connection)
