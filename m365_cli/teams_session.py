"""Teams Web transport. Service credentials stay in memory and never enter results."""

import base64
import json
import re
import time
from urllib.parse import urlsplit

from .errors import OutlookError
from .session import endpoint, read_config, save_config, start_browser

ORIGIN = "https://teams.cloud.microsoft"
FAMILIES = {"csa": "csa", "mt": "mt", "chatsvc": "chat"}
AUDIENCES = {
    "csa": "https://chatsvcagg.teams.microsoft.com",
    "mt": "https://api.spaces.skype.com",
    "chat": "https://ic3.teams.office.com",
}
# Copy only headers used by the observed Teams client, never cookies or arbitrary headers.
CLIENT_HEADERS = {
    "authorization",
    "x-skypetoken",
    "x-ms-client-caller",
    "x-ms-client-type",
    "x-ms-client-version",
    "x-ms-client-env",
    "x-ms-client-consumer-type",
    "x-ms-db-fast-failed",
    "x-ms-migration",
    "x-ms-object-id",
    "x-ms-partition",
    "x-ms-region",
    "x-ms-request-priority",
    "x-ms-session-id",
    "x-ms-user-type",
    "x-ms-ringoverride",
    "x-ms-client-ring",
    "x-ms-client-flights",
    "behavioroverride",
}


def trusted_origin(url):
    try:
        u = urlsplit(url)
        return (
            u.scheme == "https"
            and u.hostname == "teams.cloud.microsoft"
            and u.port in (None, 443)
            and u.username is None
            and u.password is None
        )
    except ValueError:
        return False


def claims(value):
    try:
        part = value.removeprefix("Bearer ").split(".")[1]
        result = json.loads(base64.urlsafe_b64decode(part + "=" * (-len(part) % 4)))
        return result if isinstance(result, dict) else {}
    except (ValueError, IndexError, TypeError):
        return {}


def live_claims(value):
    c = claims(value)
    if not isinstance(c.get("exp"), (int, float)) or c["exp"] <= time.time() + 15:
        raise OutlookError("login_required", "Teams credentials expired. Run m365 teams login.")
    return c


def validate_credentials(captured, expected, tenant=None):
    if set(captured) != {"csa", "mt", "chat"} or not expected:
        raise OutlookError("login_required", "Teams credentials are incomplete.")
    csa = live_claims(captured["csa"][1].get("authorization", ""))
    account = next(
        (
            csa.get(k)
            for k in ("upn", "preferred_username", "unique_name")
            if isinstance(csa.get(k), str) and csa[k].casefold() == expected.casefold()
        ),
        None,
    )
    oid, tid = csa.get("oid"), csa.get("tid")
    if not account or not isinstance(oid, str) or not isinstance(tid, str):
        raise OutlookError("account_mismatch", "Teams session does not match the pinned account.")
    if tenant and tid.casefold() != tenant.casefold():
        raise OutlookError("tenant_mismatch", "Teams session does not match the pinned tenant.")
    for family, (_, headers) in captured.items():
        if headers.get("authorization"):
            c = live_claims(headers["authorization"])
            if c.get("aud") != AUDIENCES[family] or c.get("tid") != tid or c.get("oid") != oid:
                raise OutlookError("identity_mismatch", "Teams service credentials disagree.")
        elif family != "chat" or not headers.get("x-skypetoken"):
            raise OutlookError("login_required", "Teams service credentials are missing.")
        if family == "chat" and headers.get("x-skypetoken"):
            c = live_claims(headers.get("x-skypetoken", ""))
            if c.get("tid") != tid or c.get("skypeid") not in (
                oid,
                "orgid:" + oid,
                "8:orgid:" + oid,
            ):
                raise OutlookError("identity_mismatch", "Teams chat credentials disagree.")
    return account, tid, oid


class TeamsSession:
    def __init__(self, *, launch=False):
        self.launch = launch
        self.pw = None
        self.captured = {}
        self.headers = {}
        self.base = {}
        self.account = self.tenant = self.oid = self.mri = ""

    def __enter__(self):
        from playwright.sync_api import sync_playwright

        address = start_browser() if self.launch else endpoint()
        if not address:
            raise OutlookError("login_required", "Run m365 teams login to open the saved session.")
        self.pw = sync_playwright().start()
        try:
            self.browser = self.pw.chromium.connect_over_cdp(address, timeout=10000)
            self.context = self.browser.contexts[0]
            self.context.on("request", self._capture)
            self.page = next((p for p in self.context.pages if trusted_origin(p.url)), None)
            if self.page is None:
                self.page = self.context.new_page()
                self.page.goto(ORIGIN, wait_until="domcontentloaded", timeout=45000)
            else:
                self.page.reload(wait_until="domcontentloaded", timeout=45000)
            return self
        except Exception:
            self.pw.stop()
            raise OutlookError(
                "login_required", "Could not open Teams Web. Run m365 teams login."
            ) from None

    def __exit__(self, *exc):
        if self.pw:
            self.pw.stop()  # Disconnect only; keep the persistent browser open.
        self.captured.clear()
        self.headers.clear()

    def _capture(self, request):
        if not trusted_origin(request.url) or request.method != "GET":
            return
        u = urlsplit(request.url)
        match = re.fullmatch(r"/api/(csa|mt|chatsvc)/([a-z0-9-]+)/(.*)", u.path)
        if not match:
            return
        family = FAMILIES[match[1]]
        if family == "csa" and match[3] not in (
            "api/v3/teams/users/me",
            "api/v3/teams/users/me/updates",
        ):
            return
        if family == "mt" and match[3] != "beta/users/tenantsv2":
            return
        try:
            headers = {k: v for k, v in request.all_headers().items() if k in CLIENT_HEADERS}
            if (
                family == "chat"
                and not (headers.get("x-skypetoken") or headers.get("authorization"))
            ) or (family != "chat" and not headers.get("authorization")):
                return
            self.captured[family] = (f"{ORIGIN}/api/{match[1]}/{match[2]}", headers)
        except Exception:
            pass  # Canceled browser requests are not command failures.

    def authenticate(self, *, expected=None, tenant=None, wait=30):
        config = read_config()
        expected = expected or config.get("account")
        tenant = tenant or config.get("teams_tenant")
        if not expected:
            raise OutlookError(
                "account_required", "Use --account EMAIL or pin an account with m365 auth login."
            )
        deadline = time.monotonic() + wait
        while len(self.captured) < 3 and time.monotonic() < deadline:
            self.page.wait_for_timeout(250)
        if len(self.captured) != 3:
            raise OutlookError(
                "login_required",
                "Complete Teams sign-in in the dedicated Chrome window, then run m365 teams login.",
            )
        account, tid, oid = validate_credentials(self.captured, expected, tenant)
        # Freeze this command's credentials; later browser account switches cannot replace them.
        self.base = {k: v[0] for k, v in self.captured.items()}
        self.headers = {k: dict(v[1]) for k, v in self.captured.items()}
        self.account, self.tenant, self.oid = account, tid, oid
        self.mri = "8:orgid:" + oid
        data = self.request("mt", "beta/users/tenantsv2")
        if not isinstance(data, list) or not any(
            isinstance(x, dict) and x.get("tenantId") == tid and x.get("userId") == oid
            for x in data
        ):
            self.account = ""
            raise OutlookError("identity_mismatch", "Teams server identity could not be verified.")
        if not config.get("teams_tenant") and (
            not config.get("account") or config["account"].casefold() == account.casefold()
        ):
            config.update(account=account, teams_tenant=tid)
            save_config(config)
        return {
            "authenticated": True,
            "account": account,
            "tenant_id": tid,
            "user_id": oid,
            "mri": self.mri,
        }

    def request(
        self,
        family,
        path,
        *,
        method="GET",
        body=None,
        raw=False,
        headers=None,
        write=False,
        include_location=False,
    ):
        # Only API-owned relative paths reach this method. Never follow server redirects.
        if (
            family not in self.base
            or not isinstance(path, str)
            or not path.startswith(("v1/", "beta/", "api/v3/"))
            or "\\" in path
            or any(c in path for c in "\r\n#")
            or ".." in urlsplit(path).path.split("/")
        ):
            raise OutlookError("invalid_route", "Invalid Teams API route.")
        if headers and any(k.lower() not in ("if-match", "prefer") for k in headers):
            raise OutlookError("invalid_headers", "Unsupported Teams request headers.")
        # Tokens are checked again before every call, never renewed by retrying a write.
        snapshot = {k: (self.base[k], v) for k, v in self.headers.items()}
        validate_credentials(snapshot, self.account, self.tenant)
        outgoing = dict(self.headers[family])
        outgoing.update(headers or {})
        if body is not None:
            outgoing["content-type"] = "application/json;charset=UTF-8"
        try:
            response = self.context.request.fetch(
                self.base[family] + "/" + path,
                method=method,
                headers=outgoing,
                data=(body if raw else json.dumps(body)) if body is not None else None,
                timeout=25000,
                max_redirects=0,
                max_retries=0,
            )
            status = response.status
            try:
                data = response.json()
            except Exception:
                data = None
            retry_after = response.headers.get("retry-after", "")
            location = response.headers.get("location", "") if include_location else None
            response.dispose()
        except Exception:
            raise OutlookError(
                "write_outcome_unknown" if write else "transport_error",
                "Teams transport failed. Inspect any write before retrying.",
            ) from None
        if 200 <= status < 300:
            if data is None and method in ("GET", "POST"):
                raise OutlookError(
                    "write_outcome_unknown" if write else "unexpected_response",
                    "Teams returned no readable acknowledgement.",
                )
            return {"body": data, "location": location} if include_location else data
        code = {
            401: "session_rejected",
            403: "forbidden",
            404: "not_found",
            409: "conflict",
            412: "conflict",
            429: "rate_limited",
        }.get(status)
        if not code:
            code = "write_outcome_unknown" if write else "api_error"
        details = {"status": status}
        if status == 429 and retry_after.isdigit():
            details["retry_after_seconds"] = int(retry_after)
        raise OutlookError(
            code, "Teams rejected the request. No automatic retry was attempted.", **details
        )
