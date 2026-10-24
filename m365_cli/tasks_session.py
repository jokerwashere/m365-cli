"""Task service transports; browser credentials remain transient and account-bound."""

import json
import re
import time
from contextlib import ExitStack
from urllib.parse import unquote, urlsplit

from .contracts import changed, object_value
from .errors import OutlookError
from .session import endpoint, read_config, save_config
from .teams_session import claims

BASES = {
    "graph": "https://graph.microsoft.com/",
    "todo": "https://substrate.office.com/todob2/api/v1/",
    "planner": "https://api.planner.svc.cloud.microsoft/taskapi/v4.0/",
}
AUDIENCES = {
    "graph": {"https://graph.microsoft.com", "00000003-0000-0000-c000-000000000000"},
    "todo": {"https://outlook.office.com", "https://substrate.office.com"},
    "planner": {"09abbdfd-ed23-44ee-a2d9-a627aa1c90f3"},
}
APPS = {"planner.cloud.microsoft", "to-do.office.com"}


def relative_path(family, path):
    if family not in BASES or not isinstance(path, str) or not path:
        raise OutlookError("invalid_argument", "Invalid task service path.")
    u = urlsplit(path)
    decoded = unquote(u.path)
    if (
        u.scheme
        or u.netloc
        or u.fragment
        or path.startswith("/")
        or "\\" in decoded
        or any(x in (".", "..") for x in decoded.split("/"))
        or any(ord(x) < 32 for x in path)
    ):
        raise OutlookError("invalid_argument", "Unsafe task service path.")
    prefixes = {
        "graph": (
            "v1.0/me",
            "beta/me",
            "v1.0/planner/",
            "beta/planner/",
            "beta/users/",
            "v1.0/groups/",
            "beta/groups/",
        ),
        "todo": ("taskfolders", "tasks/", "foldergroups", "myDayFeed/"),
        "planner": ("users('me')", "tasks", "plans(", "undo"),
    }
    if not path.startswith(prefixes[family]):
        raise OutlookError("invalid_argument", "Task service route is not allowed.")
    return path


def token_valid(token, family, account, tenant=None, oid=None):
    c = claims(token)
    return bool(
        isinstance(c.get("exp"), (int, float))
        and c["exp"] > time.time() + 15
        and c.get("aud") in AUDIENCES[family]
        and ("Todo-Internal.ReadWrite" if family == "todo" else "Tasks.ReadWrite")
        in str(c.get("scp", "")).split()
        and any(
            isinstance(c.get(k), str) and c[k].casefold() == account.casefold()
            for k in ("upn", "unique_name", "preferred_username")
        )
        and isinstance(c.get("oid"), str)
        and isinstance(c.get("tid"), str)
        and (tenant is None or c["tid"].casefold() == tenant.casefold())
        and (oid is None or c["oid"] == oid)
    )


class TasksSession:
    def __init__(self, expected=None, required=("graph",), wait=30):
        config = read_config()
        self.account = expected or config.get("account")
        self.tenant = config.get("tasks_tenant") or config.get("teams_tenant")
        self.required = set(required) | {"graph"}
        self.wait = wait
        self.tokens = {}
        self.oid = None
        self.frozen = False
        self.last_meta = {}

    def __enter__(self):
        from playwright.sync_api import sync_playwright

        self.stack = ExitStack()
        try:
            if not self.account:
                raise OutlookError("account_required", "Pin an account with m365 auth login first.")
            address = endpoint()
            if not address:
                raise OutlookError(
                    "login_required",
                    "Open the saved browser with m365 auth login, then m365 tasks login.",
                )
            pw = self.stack.enter_context(sync_playwright())
            self.browser = pw.chromium.connect_over_cdp(address)
            if not self.browser.contexts:
                raise OutlookError("login_required", "The saved browser has no active context.")
            self.ctx = self.browser.contexts[0]
            self.ctx.on("request", self.capture)
            self.stack.callback(self.ctx.remove_listener, "request", self.capture)
            pages = [p for p in self.ctx.pages if urlsplit(p.url).hostname in APPS]
            planner = next(
                (p for p in pages if urlsplit(p.url).hostname == "planner.cloud.microsoft"), None
            )
            if planner is None:
                planner = self.ctx.new_page()
                planner.goto(
                    "https://planner.cloud.microsoft/", wait_until="domcontentloaded", timeout=45000
                )
            else:
                planner.reload(wait_until="domcontentloaded", timeout=45000)
            until = time.monotonic() + self.wait
            while not self.required <= self.tokens.keys() and time.monotonic() < until:
                planner.wait_for_timeout(250)
            if not self.required <= self.tokens.keys():
                raise OutlookError(
                    "login_required",
                    "Task credentials unavailable. Run m365 tasks login and complete any sign-in in Planner/To Do.",
                    missing_services=sorted(self.required - self.tokens.keys()),
                )
            c = claims(self.tokens["graph"])
            self.tenant, self.oid = c["tid"], c["oid"]
            self.frozen = True
            if any(
                not token_valid(self.tokens[f], f, self.account, self.tenant, self.oid)
                for f in self.required
            ):
                raise OutlookError(
                    "account_mismatch",
                    "Task service credentials do not match one account and tenant.",
                )
            me = self.request("graph", "v1.0/me?$select=id,userPrincipalName,mail")
            object_value(me, "tasks.identity", fields=("id",))
            if me["id"] != self.oid or not any(
                isinstance(me.get(k), str) and me[k].casefold() == self.account.casefold()
                for k in ("userPrincipalName", "mail")
            ):
                raise OutlookError(
                    "account_mismatch", "Microsoft returned a different task account."
                )
            if "planner" in self.required:
                native = self.request("planner", "users('me')")
                if not isinstance(native, dict) or not any(
                    native.get(k) == self.oid for k in ("externalId", "primaryUserId")
                ):
                    raise OutlookError(
                        "account_mismatch", "Native Planner identity did not match Graph."
                    )
            config = read_config()
            if not config.get("tasks_tenant"):
                save_config({**config, "tasks_tenant": self.tenant})
            return self
        except OutlookError:
            self.__exit__(None, None, None)
            raise
        except Exception:
            self.__exit__(None, None, None)
            raise OutlookError(
                "login_required",
                "Cannot acquire task credentials from the saved browser. Run m365 tasks login.",
            ) from None

    def __exit__(self, *args):
        self.tokens.clear()
        self.stack.close()

    def capture(self, request):
        if self.frozen:
            return
        try:
            u = urlsplit(request.url)
            if u.scheme != "https" or u.username or u.password or u.port not in (None, 443):
                return
            for family, base in BASES.items():
                if family not in self.required:
                    continue
                b = urlsplit(base)
                if u.hostname == b.hostname and u.path.startswith(b.path):
                    token = request.all_headers().get("authorization", "")
                    if token_valid(token, family, self.account, self.tenant):
                        self.tokens[family] = token
        except Exception:
            return

    def request(self, family, path, method="GET", body=None, headers=None):
        path = relative_path(family, path)
        method = method.upper()
        if method not in ("GET", "POST", "PATCH", "PUT", "DELETE"):
            raise OutlookError("invalid_argument", "Unsupported task request method.")
        write = method != "GET"
        token = self.tokens.get(family, "")
        if not token_valid(token, family, self.account, self.tenant, self.oid):
            raise OutlookError(
                "login_required",
                "Task credentials expired or are unavailable. Run m365 tasks login.",
                service=family,
            )
        h = {
            "authorization": token,
            "accept": "application/json",
            "content-type": "application/json",
        }
        if family == "graph":
            h["Prefer"] = "include-unknown-enum-members"
        if family == "planner":
            h.update({"OData-Version": "4.01", "Prefer": "return=representation"})
        for k, v in (headers or {}).items():
            if (
                k.lower() not in ("if-match", "prefer")
                or not isinstance(v, str)
                or any(ord(c) < 32 for c in v)
            ):
                raise OutlookError("invalid_argument", "Invalid task request header.")
            h[k] = v
        self.last_meta = {}
        try:
            response = self.ctx.request.fetch(
                BASES[family] + path,
                method=method,
                headers=h,
                data=json.dumps(body) if body is not None else None,
                max_redirects=0,
                max_retries=0,
                timeout=25000,
            )
            status = response.status
            try:
                data = response.json()
            except Exception:
                data = None
            self.last_meta = {
                "status": status,
                "operation_id": response.headers.get("ms-planner-operationid")
                or response.headers.get("x-planner-operationid"),
            }
            retry = response.headers.get("retry-after", "")
            response.dispose()
        except Exception:
            raise OutlookError(
                "write_outcome_unknown" if write else "transport_error",
                "Task request did not complete reliably. Inspect any write before retrying.",
                service=family,
            ) from None
        codes = {
            401: "session_rejected",
            403: "forbidden",
            404: "item_not_found",
            409: "conflict",
            412: "conflict",
            429: "rate_limited",
        }
        if status in codes:
            details = {"service": family, "http_status": status}
            if retry.isdecimal():
                details["retry_after_seconds"] = int(retry)
            raise OutlookError(codes[status], "Task service rejected the request.", **details)
        if not 200 <= status < 300:
            code = (
                "write_outcome_unknown"
                if write and (status >= 500 or status < 400)
                else "api_error"
            )
            error = data.get("error", {}) if isinstance(data, dict) else {}
            ec = error.get("code") if isinstance(error, dict) else None
            details = {"service": family, "http_status": status}
            if isinstance(ec, str) and re.fullmatch(r"[A-Za-z0-9_.-]{1,100}", ec):
                details["response_code"] = ec
            raise OutlookError(
                code, "Task API rejected the operation; no automatic retry was made.", **details
            )
        if status == 204 and write:
            return None
        if not isinstance(data, dict):
            changed(f"tasks.{family}.response", "$", write=write)
        return data
