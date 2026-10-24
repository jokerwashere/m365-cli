"""Own Chrome profile, local CDP connection, session secrets in memory only."""

import base64
import fcntl
import json
import os
import re
import subprocess
import time
import uuid
from contextlib import contextmanager
from pathlib import Path
from urllib.parse import quote, urlsplit
from urllib.request import ProxyHandler, build_opener

from .errors import OutlookError
from .http_transport import BrowserCookies, DirectHTTP, service_request

OUTLOOK_HEADERS = {
    "authorization",
    "x-owa-canary",
    "x-owa-sessionid",
    "x-anchormailbox",
    "x-tenantid",
    "prefer",
    "user-agent",
}
READ_ACTIONS = {
    "GetOwaUserConfiguration",
    "GetFolder",
    "FindFolder",
    "FindItem",
    "GetItem",
    "GetAttachment",
    "GetCalendarFolders",
    "GetCalendarView",
    "GetMasterCategoryList",
    "GetUserAvailabilityInternal",
}

OUTLOOK_HOSTS = {
    "outlook.office.com",
    "outlook.office365.com",
    "outlook.cloud.microsoft",
    "outlook.live.com",
}
LOGIN_URL = "https://outlook.office.com/mail/"
DEFAULT_BROWSER = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"


def is_outlook_url(url):
    try:
        u = urlsplit(url)
        return (
            u.scheme == "https"
            and u.hostname in OUTLOOK_HOSTS
            and u.port in (None, 443)
            and u.username is None
        )
    except ValueError:
        return False


def state_dir():
    project = Path(__file__).resolve().parents[1]
    default = (
        project / ".state"
        if (project / "pyproject.toml").exists()
        else Path.home() / ".local" / "share" / "m365-cli"
    )
    path = (
        Path(os.environ.get("M365_HOME", os.environ.get("OUTLOOK_ACCESS_HOME", default)))
        .expanduser()
        .resolve()
    )
    path.mkdir(mode=0o700, parents=True, exist_ok=True)
    path.chmod(0o700)
    return path


@contextmanager
def command_lock():
    with (state_dir() / "command.lock").open("a+") as f:
        try:
            fcntl.flock(f, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            raise OutlookError(
                "busy", "Another m365 CLI command is running. Retry after it finishes."
            ) from None
        yield


def read_config():
    p = state_dir() / "config.json"
    try:
        config = json.loads(p.read_text()) if p.exists() else {}
        if not isinstance(config, dict) or (
            "account" in config and not isinstance(config["account"], str)
        ):
            raise ValueError("invalid configuration")
        return config
    except (OSError, ValueError):
        raise OutlookError(
            "config_invalid",
            "Cannot read the saved account configuration; repair config.json or use a new M365_HOME.",
        ) from None


def save_config(config):
    from .operations import atomic_json

    p = state_dir() / "config.json"
    atomic_json(p, config)


def endpoint():
    port_file = state_dir() / "profile" / "DevToolsActivePort"
    try:
        port = int(port_file.read_text().splitlines()[0])
        if not 1024 <= port <= 65535:
            return None
        # Never use an environment HTTP proxy for the local debugging endpoint.
        with build_opener(ProxyHandler({})).open(
            f"http://127.0.0.1:{port}/json/version", timeout=1
        ) as r:
            data = json.load(r)
        ws = data["webSocketDebuggerUrl"]
        parsed = urlsplit(ws)
        if (
            parsed.scheme == "ws"
            and parsed.hostname in ("127.0.0.1", "localhost")
            and parsed.port == port
        ):
            return ws
    except (OSError, ValueError, KeyError, IndexError):
        pass
    return None


def start_browser():
    current = endpoint()
    if current:
        return current
    binary = os.environ.get(
        "M365_BROWSER", os.environ.get("OUTLOOK_ACCESS_BROWSER", DEFAULT_BROWSER)
    )
    if not Path(binary).is_file():
        raise OutlookError(
            "browser_missing", "Install Google Chrome or set M365_BROWSER to its executable."
        )
    profile = state_dir() / "profile"
    profile.mkdir(mode=0o700, exist_ok=True)
    # Normal Chrome security defaults. CDP is loopback-only, in an isolated profile.
    subprocess.Popen(
        [
            binary,
            f"--user-data-dir={profile}",
            "--remote-debugging-port=0",
            "--remote-debugging-address=127.0.0.1",
            "--no-first-run",
            "--no-default-browser-check",
            LOGIN_URL,
        ],
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        start_new_session=True,
    )
    for _ in range(100):
        if current := endpoint():
            return current
        time.sleep(0.1)
    raise OutlookError(
        "browser_start_failed",
        "Chrome did not expose its local endpoint. Check the dedicated browser window.",
    )


class Session:
    def __init__(self, *, launch=False):
        self.launch = launch
        self.headers = {}
        self.account = ""
        self.time_zone = "UTC"
        self.pw = self.browser = self.page = None
        self.context = None
        self.captured = {}
        self._observed = {}
        self.bootstrap_url = None
        self.origin = None
        self.http = self.cookies = None
        self.browser_connected = False

    def __enter__(self):
        from playwright.sync_api import sync_playwright

        address = start_browser() if self.launch else endpoint()
        if not address:
            raise OutlookError("login_required", "Run: m365 auth login --account YOUR_EMAIL")
        self.pw = sync_playwright().start()
        try:
            self.browser = self.pw.chromium.connect_over_cdp(address, timeout=10000)
            self.context = self.browser.contexts[0]
            self.browser_connected = True
            self.context.on("request", self._capture)
            self.page = next((p for p in self.context.pages if is_outlook_url(p.url)), None)
            if self.page is None:
                if not self.launch:
                    raise OutlookError(
                        "login_required", "No Outlook Web tab is open. Run m365 auth login."
                    )
                # Keep other service tabs (including Teams) intact during Outlook login.
                self.page = next(
                    (
                        p
                        for p in self.context.pages
                        if p.url == "about:blank"
                        or urlsplit(p.url).hostname
                        in ("login.microsoftonline.com", "login.live.com")
                    ),
                    None,
                )
                if self.page is None:
                    self.page = self.context.new_page()
                if self.page.url == "about:blank" or urlsplit(self.page.url).hostname not in (
                    "login.microsoftonline.com",
                    "login.live.com",
                ):
                    self.page.goto(LOGIN_URL, wait_until="domcontentloaded")
            self.page.set_default_timeout(10000)
            return self
        except Exception:
            self.pw.stop()
            raise

    def __exit__(self, *exc):
        self.disconnect_browser()
        self.headers.clear()
        self.captured.clear()
        self._observed.clear()
        if self.cookies:
            self.cookies.clear()
        if self.http:
            self.http.close()
        self.http = None

    def disconnect_browser(self):
        """Release automation after authentication; leave the user's Chrome running."""
        if self.pw:
            self.pw.stop()
        self.pw = self.browser = self.page = self.context = None
        self.browser_connected = False

    def transport_info(self):
        return {
            "mode": "direct_http",
            "browser_connected": self.browser_connected,
            "credential_storage": "command_memory",
        }

    def _capture(self, request):
        if not is_outlook_url(request.url) or "/owa/" not in urlsplit(request.url).path:
            return
        try:
            incoming = request.all_headers()
            origin = "https://" + urlsplit(request.url).hostname
            self.captured[origin] = {k: v for k, v in incoming.items() if k in OUTLOOK_HEADERS}
            if request.url == self.bootstrap_url and request.method == "POST":
                self._observed = {
                    k: v for k, v in incoming.items() if k in OUTLOOK_HEADERS | {"cookie"}
                }
        except Exception:
            pass  # A canceled request is not a failed command.

    def status(self, expected=None):
        if not is_outlook_url(self.page.url):
            return {"authenticated": False, "account": None}
        source = Path(__file__).with_name("session.js").read_text()
        return self.page.evaluate(
            source,
            {
                "op": "status",
                "headers": self.captured.get("https://" + urlsplit(self.page.url).hostname, {}),
                "account": expected,
            },
        )

    def _bootstrap(self, expected):
        from .api import check_response

        if self.http:
            self.http.close()
        if not is_outlook_url(self.page.url):
            raise OutlookError("login_required", "Outlook sign-in is required.")
        self.origin = "https://" + urlsplit(self.page.url).hostname
        self.base = (
            "/owa/service.svc"
            if self.origin == "https://outlook.cloud.microsoft"
            else "/owa/0/service.svc"
        )
        self.bootstrap_url = (
            self.origin + self.base + "?action=GetOwaUserConfiguration&app=Mail&n=0"
        )
        self._observed.clear()
        before = self.context.cookies([self.bootstrap_url])
        check_response(
            self.page.evaluate(
                Path(__file__).with_name("session.js").read_text(),
                {
                    "op": "bootstrap",
                    "headers": self.captured.get(self.origin, {}),
                    "account": expected,
                },
            )
        )
        if not self._observed or not (
            self._observed.get("authorization") or self._observed.get("x-owa-canary")
        ):
            raise OutlookError(
                "login_required",
                "Outlook credentials could not be captured from its verified bootstrap request.",
            )
        self.headers = {k: v for k, v in self._observed.items() if k in OUTLOOK_HEADERS}
        # Bearer authentication is independent of the browser's rotating cookie
        # session, canary and session-ID headers. Do not mix the two schemes.
        # Canary-only sessions still require their scoped cookies and headers.
        if self.headers.get("authorization"):
            self.headers = {
                k: v for k, v in self.headers.items() if k in ("authorization", "user-agent")
            }
        self.cookies = (
            BrowserCookies("", [], self.bootstrap_url)
            if self.headers.get("authorization")
            else BrowserCookies(
                self._observed.get("cookie", ""),
                before + self.context.cookies([self.bootstrap_url]),
                self.bootstrap_url,
            )
        )
        self._observed.clear()
        self.http = DirectHTTP(self.origin)

    def _validate_credentials(self):
        if not self.http or not self.origin:
            raise OutlookError("login_required", "Authenticate the Outlook session first.")
        token = self.headers.get("authorization")
        if token:
            try:
                part = token.removeprefix("Bearer ").split(".")[1]
                claims = json.loads(base64.urlsafe_b64decode(part + "=" * (-len(part) % 4)))
                expiry = claims["exp"]
                if (
                    not isinstance(expiry, (float, int))
                    or isinstance(expiry, bool)
                    or not time.time() + 15 < expiry < float("inf")
                ):
                    raise ValueError
            except (ValueError, KeyError, IndexError, TypeError):
                raise OutlookError(
                    "login_required",
                    "Outlook credentials expired or changed; run m365 auth login. No write was retried.",
                ) from None
            identity = (
                claims.get("preferred_username") or claims.get("upn") or claims.get("unique_name")
            )
            if identity is not None and not isinstance(identity, str):
                raise OutlookError("login_required", "Outlook credential identity changed format.")
            if self.account and identity and identity.casefold() != self.account.casefold():
                raise OutlookError(
                    "account_mismatch", "Outlook credentials do not match the verified account."
                )
        elif not (
            self.headers.get("x-owa-canary")
            and self.cookies
            and self.cookies.header(self.origin + self.base)
        ):
            raise OutlookError(
                "login_required", "Outlook session credentials are missing or expired."
            )

    def authenticate(self, *, wait=0, expected=None):
        expected = expected or read_config().get("account")
        if self.account and not self.browser_connected and self.http:
            if expected and expected.casefold() != self.account.casefold():
                raise OutlookError("account_mismatch", "Session is pinned to a different account.")
            self._validate_credentials()
            return {
                "authenticated": True,
                "account": self.account,
                "time_zone": self.time_zone,
                "transport": self.transport_info(),
            }
        deadline = time.monotonic() + wait
        refreshed = False
        while True:
            try:
                # Microsoft may replace the initial login tab.
                page = next((p for p in self.context.pages if is_outlook_url(p.url)), None)
                if page:
                    self.page = page
                info = self.status(expected)
                if info.get("authenticated"):
                    self._bootstrap(expected)
                    # Verify the actual mailbox on the server, not just cached JWT claims.
                    data = self.request(
                        "GetOwaUserConfiguration",
                        {"__type": "GetOwaUserConfigurationRequest:#Exchange"},
                        check_account=False,
                    )
                    account = find_identity(data)
                    if not account:
                        raise OutlookError(
                            "identity_unknown", "Outlook session identity could not be verified."
                        )
                    if expected and account.casefold() != expected.casefold():
                        raise OutlookError(
                            "account_mismatch",
                            f"Expected {expected}; the browser is signed in as {account}.",
                        )
                    self.account = account
                    self.time_zone = find_timezone(data) or "UTC"
                    self.disconnect_browser()
                    self.captured.clear()
                    return {
                        "authenticated": True,
                        "account": account,
                        "time_zone": self.time_zone,
                        "transport": self.transport_info(),
                    }
                if wait == 0 and not refreshed and is_outlook_url(self.page.url):
                    # A refresh lets Chrome renew expired tokens through the normal session.
                    self.page.reload(wait_until="domcontentloaded", timeout=30000)
                    self.page.wait_for_timeout(1500)
                    refreshed = True
                    continue
            except OutlookError as error:
                if (
                    error.code in ("session_rejected", "login_required")
                    and not refreshed
                    and is_outlook_url(self.page.url)
                ):
                    self.headers.clear()
                    self.captured.clear()
                    self._observed.clear()
                    if self.cookies:
                        self.cookies.clear()
                    self.page.reload(wait_until="domcontentloaded", timeout=30000)
                    self.page.wait_for_timeout(1500)
                    refreshed = True
                    continue
                if (
                    wait
                    and error.code in ("session_rejected", "login_required")
                    and time.monotonic() < deadline
                ):
                    self.page.wait_for_timeout(1000)
                    continue
                raise
            except Exception:
                if time.monotonic() >= deadline:
                    raise OutlookError(
                        "login_required",
                        "Complete sign-in in the dedicated Chrome window, then run m365 auth login again.",
                    ) from None
            if time.monotonic() >= deadline:
                raise OutlookError(
                    "login_required",
                    "Complete sign-in in the dedicated Chrome window, then run auth login again.",
                )
            self.page.wait_for_timeout(1000)

    def request(
        self, action, body, *, check_account=True, path=None, time_zone="UTC", native=False
    ):
        if native and action not in (
            "GetMasterCategoryList",
            "UpdateMasterCategoryList",
            "ApplyCategoryOperation",
            "GetUserAvailabilityInternal",
        ):
            raise OutlookError("invalid_argument", "Unapproved native Outlook action.")
        if check_account and not self.account:
            self.authenticate()
        from .api import check_response, envelope

        if (path is not None and (path != "/searchservice/api/v2/query" or action is not None)) or (
            path is None
            and (
                not isinstance(action, str)
                or not re.fullmatch(r"[A-Za-z][A-Za-z0-9]{0,99}", action)
            )
        ):
            raise OutlookError("invalid_route", "Unapproved Outlook API route.")
        self._validate_credentials()
        # Only explicitly known reads may be retried. Unknown actions remain
        # conservative mutations, including actions with unfamiliar name prefixes.
        write = action is not None and action not in READ_ACTIONS
        if write:
            try:
                identity = self.request(
                    "GetOwaUserConfiguration",
                    {"__type": "GetOwaUserConfigurationRequest:#Exchange"},
                )
            except OutlookError as error:
                error.details.update(phase="identity_verification", mutation_request_sent=False)
                raise
            if find_identity(identity).casefold() != self.account.casefold():
                raise OutlookError(
                    "account_mismatch", "Server identity changed; the write was not sent."
                )
        url = self.origin + (path or self.base + "?action=" + action + "&app=Mail&n=0")
        headers = {
            **self.headers,
            "Accept": "application/json",
            "Content-Type": "application/json; charset=utf-8",
            "X-Req-Source": "Mail",
            "X-OWA-CorrelationId": str(uuid.uuid4()),
            "prefer": 'IdType="ImmutableId"',
        }
        if action:
            headers["Action"] = action
        if self.account:
            headers["x-anchormailbox"] = "SMTP:" + self.account
        if self.cookies and (cookie := self.cookies.header(url)):
            headers["Cookie"] = cookie
        payload = envelope(action, body, time_zone=time_zone) if not path and not native else body
        try:
            response = service_request(
                self.http,
                url,
                read=not write,
                method="POST",
                headers=headers,
                body=json.dumps(payload, allow_nan=False).encode(),
                maximum=80 * 1024 * 1024,
                write=write,
            )
            result = {
                "status": response.status,
                "body": response.json(),
                "retry_after": response.headers.get("retry-after"),
            }
        except OutlookError:
            raise
        except Exception:
            # Never leak transport exception strings containing request credentials.
            raise OutlookError(
                "write_outcome_unknown" if write else "transport_error",
                "Outlook API transport failed. Inspect any write before retrying.",
                action=action,
            ) from None
        try:
            if action == "GetUserAvailabilityInternal":
                from .scheduling import response_check

                return response_check(result)
            return check_response(result)
        except OutlookError as error:
            error.details.setdefault("action", action or "search")
            if write and error.code in ("unexpected_response", "api_changed", "api_error"):
                raise OutlookError(
                    "write_outcome_unknown",
                    "Outlook did not confirm the write outcome. Inspect the mailbox before retrying.",
                    action=action,
                    cause=error.as_dict(),
                ) from None
            raise

    def download_attachment(self, attachment_id):
        """Native item export through Microsoft's download service; secrets are memory-only."""
        from .api import check_response, typed, walk
        from .contracts import changed

        config = self.request("GetOwaUserConfiguration", typed("GetOwaUserConfigurationRequest"))
        identity = find_identity(config)
        if not identity or identity.casefold() != self.account.casefold():
            raise OutlookError(
                "account_mismatch", "Outlook account changed before attachment download."
            )
        settings = [value for value in walk(config) if "DownloadUrlBase" in value]
        if (
            len(settings) != 1
            or settings[0].get("DownloadUrlBase") != "https://attachments.office.net/owa/"
            or not isinstance(settings[0].get("FirstDownloadToken"), str)
            or not settings[0]["FirstDownloadToken"]
        ):
            changed("outlook.attachment.download", "download_settings")
        user_agent = self.headers.get("user-agent")
        if not isinstance(user_agent, str) or not user_agent:
            changed("outlook.attachment.download", "client_metadata")
        url = (
            "https://attachments.office.net/owa/"
            + quote(self.account, safe="")
            + "/service.svc/s/GetFileAttachment?id="
            + quote(attachment_id, safe="")
            + "&isDownload=true"
        )
        client = DirectHTTP("https://attachments.office.net")
        try:
            response = service_request(
                client,
                url,
                read=True,
                headers={
                    "X-Token": settings[0]["FirstDownloadToken"],
                    "X-AnchorMailbox": self.account,
                    "User-Agent": user_agent,
                    "Accept": "*/*",
                },
                maximum=50 * 1024 * 1024,
            )
        except OutlookError as exc:
            if exc.code == "response_too_large":
                raise OutlookError(
                    "attachment_too_large", "Attachment download exceeds the 50 MiB limit."
                ) from None
            raise
        finally:
            client.close()
        if response.status != 200:
            check_response(
                {
                    "status": response.status,
                    "body": {},
                    "retry_after": response.headers.get("retry-after"),
                }
            )
        disposition = response.headers.get("content-disposition")
        if not isinstance(disposition, str) or not disposition.lower().startswith("attachment;"):
            changed("outlook.attachment.download", "binary_response")
        return {
            "content": response.content,
            "content_type": response.headers.get("content-type"),
        }


def find_identity(value):
    if isinstance(value, dict):
        for key in ("UserEmailAddress", "PrimarySmtpAddress", "primarySmtpAddress"):
            if isinstance(value.get(key), str) and "@" in value[key]:
                return value[key]
        for item in value.values():
            if found := find_identity(item):
                return found
    elif isinstance(value, list):
        for item in value:
            if found := find_identity(item):
                return found
    return ""


def find_timezone(value):
    if isinstance(value, dict):
        if isinstance(value.get("TimeZone"), str):
            return value["TimeZone"]
        for child in value.values():
            if zone := find_timezone(child):
                return zone
    elif isinstance(value, list):
        for child in value:
            if zone := find_timezone(child):
                return zone
    return ""
