"""Bounded HTTPS with explicit read-only recovery and no ambient credentials or redirects."""

import http.client
import json
import re
import time
from dataclasses import dataclass
from urllib.parse import urlsplit

from .errors import OutlookError


@dataclass(repr=False)
class Response:
    status: int
    content: bytes
    headers: dict

    def json(self):
        try:
            return json.loads(self.content)
        except (ValueError, UnicodeError):
            return None


def service_request(client, url, *, read=False, **kwargs):
    """Retry a declared read once on connection failure, within one time budget."""
    deadline = time.monotonic() + 45
    for attempt in range(2 if read else 1):
        try:
            return client.request(url, timeout=max(0.1, deadline - time.monotonic()), **kwargs)
        except OutlookError as error:
            if not read or error.code != "transport_error":
                raise
            if attempt or deadline - time.monotonic() < 1:
                raise OutlookError(
                    error.code,
                    "HTTPS read did not complete reliably; no mutation was retried.",
                    **error.details,
                    read_attempts=attempt + 1,
                ) from None
            time.sleep(0.25)


class DirectHTTP:
    """One verified origin and reusable TLS connection; never replay a request."""

    def __init__(self, origin):
        try:
            u = urlsplit(origin)
            port = u.port
        except (ValueError, TypeError):
            raise OutlookError("invalid_route", "Invalid HTTPS service origin.") from None
        if (
            u.scheme != "https"
            or not u.hostname
            or u.username is not None
            or u.password is not None
            or port not in (None, 443)
            or u.path
            or u.query
            or u.fragment
        ):
            raise OutlookError("invalid_route", "Invalid HTTPS service origin.")
        self.origin, self.host = origin, u.hostname
        self._connection = None

    def close(self):
        connection, self._connection = self._connection, None
        if connection:
            connection.close()

    def request(
        self,
        url,
        *,
        method="GET",
        headers=None,
        body=None,
        maximum=16 * 1024 * 1024,
        write=False,
        timeout=45,
    ):
        u = urlsplit(url)
        if (
            u.scheme != "https"
            or u.netloc != urlsplit(self.origin).netloc
            or u.username is not None
            or u.fragment
            or "\\" in url
            or any(ord(c) < 32 or ord(c) == 127 for c in url)
        ):
            raise OutlookError("invalid_route", "Unsafe or cross-origin service request.")
        if method not in ("GET", "POST", "PATCH", "PUT", "DELETE"):
            raise OutlookError("invalid_route", "Unsupported HTTPS method.")
        outgoing = {"Accept-Encoding": "identity", **(headers or {})}
        for key, value in outgoing.items():
            if (
                not isinstance(key, str)
                or not re.fullmatch(r"[A-Za-z0-9-]+", key)
                or key.lower() in ("host", "content-length", "transfer-encoding", "connection")
                or not isinstance(value, str)
                or any(ord(c) < 32 or ord(c) == 127 for c in value)
            ):
                raise OutlookError("invalid_headers", "Invalid service request headers.")
        if body is not None and not isinstance(body, bytes):
            raise OutlookError("invalid_argument", "HTTP request content must be bytes.")
        connection = response = None
        phase = "connect"
        completed = False
        started = time.monotonic()
        try:
            # HTTPSConnection verifies TLS by default. It does not consume proxy,
            # netrc or browser state, redirect, authenticate again, or retry.
            connection = self._connection
            if connection is None or connection.sock is None:
                self.close()
                connection = http.client.HTTPSConnection(self.host, timeout=timeout)
                self._connection = connection
                connection.connect()
            socket = connection.sock
            remaining = timeout - (time.monotonic() - started)
            if remaining <= 0:
                raise TimeoutError
            socket.settimeout(remaining)
            phase = "send"
            path = u.path or "/"
            if u.query:
                path += "?" + u.query
            connection.request(method, path, body=body, headers=outgoing)
            remaining = timeout - (time.monotonic() - started)
            if remaining <= 0:
                raise TimeoutError
            socket.settimeout(remaining)
            phase = "response_headers"
            response = connection.getresponse()
            if 300 <= response.status < 400:
                raise OutlookError(
                    "write_outcome_unknown" if write else "redirect_rejected",
                    "Service redirected the request; no redirect or retry was followed.",
                )
            encoding = response.getheader("Content-Encoding", "identity").lower()
            if encoding not in ("", "identity"):
                raise OutlookError(
                    "write_outcome_unknown" if write else "api_changed",
                    "Service returned an unexpected content encoding.",
                )
            declared = response.getheader("Content-Length")
            if declared is not None and (not declared.isdigit() or int(declared) > maximum):
                raise OutlookError(
                    "write_outcome_unknown" if write else "response_too_large",
                    "Service response exceeded its bounded read contract.",
                )
            chunks, size = [], 0
            phase = "response_body"
            while True:
                if response.isclosed():
                    break
                remaining = timeout - (time.monotonic() - started)
                if remaining <= 0:
                    raise TimeoutError
                # getresponse can detach a Connection: close socket from the
                # connection object; retain it to enforce the whole-call deadline.
                socket.settimeout(remaining)
                chunk = response.read1(min(65536, maximum - size + 1))
                if not chunk:
                    break
                size += len(chunk)
                if size > maximum:
                    raise OutlookError(
                        "write_outcome_unknown" if write else "response_too_large",
                        "Service response exceeded its bounded read contract.",
                    )
                chunks.append(chunk)
            if declared is not None and size != int(declared):
                raise http.client.IncompleteRead(b"")
            result = Response(
                response.status,
                b"".join(chunks),
                {
                    key: response.getheader(key)
                    for key in ("etag", "retry-after", "content-type", "content-disposition")
                },
            )
            completed = True
            return result
        except OutlookError:
            raise
        except Exception as error:
            # Never return exception strings: they may contain URLs or credentials.
            raise OutlookError(
                "write_outcome_unknown" if write else "transport_error",
                "HTTPS request did not complete reliably. No retry was made.",
                phase=phase,
                reason="timeout"
                if isinstance(error, TimeoutError)
                else "http_protocol"
                if isinstance(error, http.client.HTTPException)
                else "connection_io"
                if isinstance(error, OSError)
                else "transport_failure",
                **(
                    {"os_error": error.errno}
                    if isinstance(error, OSError) and isinstance(error.errno, int)
                    else {}
                ),
            ) from None
        finally:
            if response:
                response.close()
            if not completed or (response and (response.will_close or response.status >= 400)):
                self.close()


class BrowserCookies:
    """Only cookies actually applicable to a verified browser bootstrap request."""

    def __init__(self, observed, scoped, url):
        self._cookies = []
        for pair in observed.split(";"):
            if not pair.strip():
                continue
            name, separator, value = pair.strip().partition("=")
            if not separator or any(ord(c) < 32 or ord(c) == 127 for c in pair):
                raise OutlookError("login_required", "Invalid browser cookie snapshot.")
            matches = [
                dict(c)
                for c in scoped
                if c.get("name") == name and c.get("value") == value and self.applies(c, url)
            ]
            if not matches:
                raise OutlookError("login_required", "Browser cookie scope could not be verified.")
            self._cookies.extend(matches)

    @staticmethod
    def applies(cookie, url):
        u = urlsplit(url)
        domain, path = cookie.get("domain", ""), cookie.get("path", "/")
        host_matches = u.hostname == domain.lstrip(".") or (
            domain.startswith(".") and u.hostname.endswith(domain)
        )
        request_path = u.path or "/"
        path_matches = request_path == path or request_path.startswith(path.rstrip("/") + "/")
        expiry = cookie.get("expires", -1)
        return bool(
            host_matches
            and path_matches
            and (not cookie.get("secure") or u.scheme == "https")
            and (expiry == -1 or expiry > time.time())
        )

    def header(self, url):
        pairs = []
        for c in self._cookies:
            pair = c["name"] + "=" + c["value"]
            if self.applies(c, url) and pair not in pairs:
                pairs.append(pair)
        return "; ".join(pairs)

    def clear(self):
        self._cookies.clear()
