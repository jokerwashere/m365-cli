"""Account-bound direct SharePoint HTTP; browser bootstrap, no write retries."""

import json
import re
import time
from contextlib import ExitStack
from urllib.parse import unquote, urlsplit

from .errors import OutlookError
from .http_transport import BrowserCookies, DirectHTTP, service_request
from .session import endpoint, read_config, save_config, start_browser
from .sharepoint_validation import MAX_DOWNLOAD, MAX_JSON, guid, site_url

BOOTSTRAP = """async a => {
 const c=window._spPageContextInfo || {};
 if(location.origin!==a.origin || String(c.userEmail||'').toLowerCase()!==a.account ||
    String(c.aadTenantId||'').toLowerCase()!==a.tenant ||
    String(c.webId||'').replace(/[{}]/g,'').toLowerCase()!==a.web)
   return {identityChanged:true};
 const r=await fetch(a.url,{credentials:'include',redirect:'error',
   headers:{Accept:'application/json;odata=nometadata'},signal:AbortSignal.timeout(15000)});
 await r.body?.cancel();
 return {status:r.status};
}"""


class SharePointSession:
    def __init__(self, site, *, expected=None, login=False, wait=30):
        self.site = site_url(site)
        self.origin = "https://" + urlsplit(self.site).hostname
        config = read_config()
        self.account = expected or config.get("account")
        self.tenant = (
            config.get("sharepoint_tenant")
            or config.get("tasks_tenant")
            or config.get("teams_tenant")
        )
        self.login, self.wait = login, wait
        self.web_id = self.site_id = None
        self.digest = None
        self.digest_until = 0
        self.last_etag = None
        self.write_attempted = False
        self.write_acknowledged = False
        self.http = self.cookies = None
        self.user_agent = None
        self.browser_connected = False

    def __enter__(self):
        from playwright.sync_api import sync_playwright

        self.stack = ExitStack()
        try:
            if not self.account:
                raise OutlookError(
                    "account_required", "Pin an account with m365 auth login --account EMAIL."
                )
            address = start_browser() if self.login else endpoint()
            if not address:
                raise OutlookError(
                    "login_required",
                    "Run m365 files login --site SITE using the saved browser profile.",
                )
            pw = self.stack.enter_context(sync_playwright())
            browser = pw.chromium.connect_over_cdp(address, timeout=10000)
            context = browser.contexts[0]
            self.browser_connected = True
            target = self.site + "/_layouts/15/viewlsts.aspx"
            self.page = next((p for p in context.pages if p.url.split("?")[0] == target), None)
            if self.page is None:
                self.page = context.new_page()
                self.page.goto(target, wait_until="domcontentloaded", timeout=45000)
            deadline = time.monotonic() + self.wait
            while True:
                identity = self.page.evaluate(
                    "() => {const c=window._spPageContextInfo||{};return {email:c.userEmail,tenant:c.aadTenantId,site:c.siteId,web:c.webId};}"
                )
                if identity.get("email") and identity.get("tenant"):
                    break
                if time.monotonic() >= deadline:
                    raise OutlookError(
                        "login_required",
                        "Complete SharePoint sign-in in the dedicated browser, then run files login --site SITE.",
                    )
                self.page.wait_for_timeout(250)
            if identity["email"].casefold() != self.account.casefold():
                raise OutlookError(
                    "account_mismatch", "SharePoint is signed in as a different account."
                )
            if self.tenant and identity["tenant"].casefold() != self.tenant.casefold():
                raise OutlookError(
                    "tenant_mismatch",
                    "SharePoint is in a different tenant from the pinned account.",
                )
            self.tenant = guid(identity["tenant"])
            self.site_id, self.web_id = guid(identity["site"]), guid(identity["web"])
            self._bootstrap(context)
            self.disconnect_browser()
            self.verify_identity()
            config = read_config()
            sites = config.get("sharepoint_sites", [])
            if not isinstance(sites, list) or any(not isinstance(s, str) for s in sites):
                raise OutlookError("config_invalid", "Invalid saved SharePoint site configuration.")
            save_config(
                {
                    **config,
                    "sharepoint_tenant": self.tenant,
                    "sharepoint_sites": list(dict.fromkeys([*sites, self.site]))[-100:],
                }
            )
            return self
        except OutlookError:
            self.__exit__()
            raise
        except Exception:
            self.__exit__()
            raise OutlookError(
                "login_required",
                "Could not verify the saved SharePoint session; run files login --site SITE.",
            ) from None

    def disconnect_browser(self):
        if getattr(self, "stack", None):
            self.stack.close()
        self.page = None
        self.browser_connected = False

    def __exit__(self, *args):
        self.disconnect_browser()
        if self.http:
            self.http.close()
        self.digest = self.user_agent = self.http = None
        if self.cookies:
            self.cookies.clear()

    def transport_info(self):
        return {
            "mode": "direct_http",
            "browser_connected": self.browser_connected,
            "credential_storage": "command_memory",
        }

    def _bootstrap(self, context):
        url = self.site + "/_api/web/currentuser?$select=Id,Email,LoginName"
        # The bootstrap response can rotate a cookie. Keep both scope snapshots
        # but export only values actually observed on the verified request.
        before = context.cookies([url])
        observed = {}

        def capture(request):
            if request.url == url and request.method == "GET":
                incoming = request.all_headers()
                observed.update(
                    {k: v for k, v in incoming.items() if k in ("cookie", "user-agent")}
                )

        context.on("request", capture)
        try:
            response = self.page.evaluate(
                BOOTSTRAP,
                {
                    "url": url,
                    "origin": self.origin,
                    "account": self.account.casefold(),
                    "tenant": self.tenant,
                    "web": self.web_id,
                },
            )
            if response.get("identityChanged"):
                raise OutlookError(
                    "account_mismatch", "SharePoint account or site changed during authentication."
                )
            if (
                response.get("status") != 200
                or not observed.get("cookie")
                or not observed.get("user-agent")
            ):
                raise OutlookError(
                    "login_required",
                    "SharePoint session could not be captured; run files login --site SITE.",
                )
            self.cookies = BrowserCookies(observed["cookie"], before + context.cookies([url]), url)
            self.user_agent = observed["user-agent"]
            self.http = DirectHTTP(self.origin)
        finally:
            context.remove_listener("request", capture)
            observed.clear()

    def verify_identity(self):
        user = self.request("/_api/web/currentuser?$select=Id,Email,LoginName")
        if (
            not isinstance(user, dict)
            or not isinstance(user.get("Email"), str)
            or user["Email"].casefold() != self.account.casefold()
        ):
            raise OutlookError(
                "identity_unknown",
                "SharePoint server identity did not match the expected account; no mutation was sent.",
            )
        site = self.request("/_api/site?$select=Id")
        web = self.request("/_api/web?$select=Id,Url")
        if (
            not isinstance(site, dict)
            or not isinstance(web, dict)
            or guid(site.get("Id")) != self.site_id
            or guid(web.get("Id")) != self.web_id
            or site_url(web.get("Url")) != self.site
        ):
            raise OutlookError(
                "identity_unknown",
                "SharePoint site identity could not be verified; no mutation was sent.",
            )
        self.user_id = user["Id"]

    def route(self, value):
        """Only same-site service URLs; this also validates opaque continuation URLs."""
        if not isinstance(value, str) or len(value) > 16000:
            raise OutlookError("invalid_route", "Invalid SharePoint service route.")
        url = value if value.startswith("https://") else self.site + value
        try:
            u = urlsplit(url)
            decoded = unquote(u.path, errors="strict")
            if (
                u.scheme != "https"
                or u.netloc != urlsplit(self.site).netloc
                or u.username is not None
                or u.fragment
                or not u.path.startswith(urlsplit(self.site).path + "/_api/")
                or "\\" in value
                or "\\" in decoded
                or any(ord(c) < 32 for c in value + decoded)
                or any(c in (".", "..") for c in decoded.split("/"))
            ):
                raise ValueError
        except (ValueError, UnicodeError):
            raise OutlookError(
                "invalid_route", "SharePoint returned an unsafe or cross-site service URL."
            ) from None
        return url

    def request(
        self,
        route,
        method="GET",
        body=None,
        *,
        headers=None,
        expected=(200,),
        binary=False,
        raw=False,
        internal=False,
    ):
        method = method.upper()
        if method not in ("GET", "POST", "PATCH", "PUT", "DELETE"):
            raise OutlookError("invalid_route", "Unsupported SharePoint method.")
        url = self.route(route)
        write = method != "GET" and not internal
        if internal and (route != "/_api/contextinfo" or method != "POST"):
            raise OutlookError("invalid_route", "Invalid internal SharePoint request.")
        if write:
            self.verify_identity()
        if write and time.monotonic() >= self.digest_until:
            info = self.request("/_api/contextinfo", "POST", internal=True)
            digest = info.get("FormDigestValue") if isinstance(info, dict) else None
            if not isinstance(digest, str) or not digest:
                raise OutlookError("api_changed", "SharePoint did not return a request digest.")
            self.digest = digest
            timeout = info.get("FormDigestTimeoutSeconds", 0)
            if not isinstance(timeout, (int, float)) or isinstance(timeout, bool) or timeout <= 30:
                raise OutlookError("api_changed", "SharePoint returned an invalid digest lifetime.")
            self.digest_until = time.monotonic() + min(timeout - 30, 1500)
        outgoing = {}
        if body is not None:
            outgoing["Content-Type"] = "application/octet-stream" if raw else "application/json"
        for key, value in (headers or {}).items():
            if (
                key.lower() not in ("if-match", "x-http-method", "content-type")
                or not isinstance(value, str)
                or any(ord(c) < 32 for c in value)
            ):
                raise OutlookError("invalid_argument", "Unsupported request header.")
            outgoing[key] = value
        if raw and not isinstance(body, bytes):
            raise OutlookError("invalid_argument", "Binary request body must be bytes.")
        if not self.http or not self.cookies:
            raise OutlookError(
                "login_required", "Authenticate SharePoint before making service requests."
            )
        cookie = self.cookies.header(url)
        if not cookie:
            raise OutlookError(
                "login_required", "SharePoint cookies expired; run files login --site SITE."
            )
        outgoing.update(
            {
                "Accept": "application/json;odata=nometadata",
                "Cookie": cookie,
                "User-Agent": self.user_agent,
            }
        )
        if write:
            outgoing["X-RequestDigest"] = self.digest
        payload = (
            body
            if raw
            else json.dumps(body, allow_nan=False).encode()
            if body is not None
            else None
        )
        self.last_etag = None
        if write:
            self.write_attempted = True
        try:
            response = service_request(
                self.http,
                url,
                read=method == "GET",
                method=method,
                headers=outgoing,
                body=payload,
                maximum=MAX_DOWNLOAD if binary else MAX_JSON,
                write=write,
            )
        except OutlookError:
            raise
        except Exception:
            raise OutlookError(
                "write_outcome_unknown" if write else "transport_error",
                "SharePoint request did not complete reliably. No retry was made.",
            ) from None
        result = {
            "status": response.status,
            "data": response.json(),
            "etag": response.headers.get("etag"),
            "retryAfter": response.headers.get("retry-after"),
        }
        status = result.get("status")
        if status not in expected:
            code = {
                400: "invalid_request",
                401: "login_required",
                403: "access_denied",
                404: "not_found",
                409: "conflict",
                412: "version_conflict",
                423: "resource_locked",
                429: "throttled",
            }.get(status, "write_outcome_unknown" if write else "service_error")
            data = result.get("data") or {}
            remote = (
                data.get("error", data.get("odata.error", {})) if isinstance(data, dict) else {}
            )
            service_code = remote.get("code") if isinstance(remote, dict) else None
            details = {"status": status}
            if isinstance(service_code, str) and re.fullmatch(r"[-\w., ]{1,150}", service_code):
                details["service_code"] = service_code
            retry = result.get("retryAfter")
            if isinstance(retry, str) and retry.isdigit():
                details["retry_after_seconds"] = min(int(retry), 86400)
            raise OutlookError(
                code,
                "SharePoint rejected the operation. Check identity, permissions, input and current resource version; writes are never retried automatically.",
                **details,
            )
        self.last_etag = result.get("etag")
        if binary:
            if write:
                self.write_acknowledged = True
            return response.content
        data = result.get("data")
        if data is None and status != 204:
            raise OutlookError(
                "write_outcome_unknown" if write else "api_changed",
                "SharePoint omitted the expected JSON acknowledgement.",
            )
        if write:
            self.write_acknowledged = True
        return data
