"""Progress and a process deadline for explicit, read-only authentication commands."""

import json
import os
import selectors
import signal
import subprocess
import sys
import time

from .errors import OutlookError

STAGES = {
    "starting": "Starting authentication check",
    "browser_connect": "Connecting to the saved Chrome profile",
    "browser_load": "Loading the Microsoft web app",
    "sign_in": "Waiting for Microsoft sign-in; complete any prompts in Chrome",
    "browser_identity": "Checking the browser session with Microsoft",
    "direct_identity": "Verifying the account through the service API",
    "connect": "Connecting to Microsoft over HTTPS",
    "response": "Waiting for the Microsoft response",
    "read_retry": "Retrying one read-only connection check",
    "cleanup": "Releasing the connection; keeping Chrome open",
}
PREFIX = "M365_AUTH_PROGRESS "
_worker_deadline = None


def is_auth_command(args):
    return (
        args.group in ("auth", "teams", "tasks", "files")
        and args.command in ("login", "status")
        and not getattr(args, "mutation", False)
    )


def options(parser, *, login=False):
    parser.add_argument(
        "--timeout",
        type=int,
        default=300 if login else 30,
        help="Total command deadline in seconds, including browser/network waits (default: %(default)s)",
    )
    parser.add_argument(
        "--quiet", action="store_true", help="Suppress progress on stderr; keep JSON stdout"
    )


def budget(args):
    value = args.timeout
    # Preserve the old Outlook/SharePoint no-interactive-wait spelling.
    zero_allowed = args.command == "login" and args.group in ("auth", "files")
    if not isinstance(value, int) or not (0 if zero_allowed else 1) <= value <= 3600:
        raise OutlookError(
            "invalid_argument",
            "Authentication timeout must be 1–3600 seconds (Outlook/SharePoint login also allow 0 for no sign-in wait).",
        )
    return value or 30


def enable_worker(args):
    global _worker_deadline
    if not is_auth_command(args):
        raise OutlookError(
            "usage_error", "The authentication worker accepts only login/status commands."
        )
    _worker_deadline = time.monotonic() + budget(args)


def progress(stage):
    if _worker_deadline is not None and stage in STAGES:
        print(PREFIX + json.dumps({"stage": stage}), file=sys.stderr, flush=True)


def network_timeout(default):
    if _worker_deadline is None:
        return default
    return max(0.1, min(default, 15, _worker_deadline - time.monotonic()))


def connect_timeout(default):
    return min(default, 8) if _worker_deadline is not None else default


def recovery(args):
    command = f"m365 {args.group}"
    site = " --site SITE" if args.group == "files" else ""
    return (
        "Keep the saved browser profile. If Microsoft shows a sign-in/MFA prompt, "
        f"complete it and run {command} login{site}. If Outlook/the web app is also slow "
        "or failing, check network connectivity and Microsoft 365 service health, then "
        f"retry {command} status{site} later. A longer --timeout can allow more sign-in "
        "time but does not repair an outage. No service mutation was sent."
    )


def annotate(error, args, stage, elapsed):
    if error.code == "runtime_error":
        error = OutlookError(
            "auth_process_failed", "Authentication stopped before verification completed."
        )
    error.details.setdefault("auth_stage", stage)
    error.details.setdefault("elapsed_seconds", round(elapsed, 1))
    error.details.setdefault("timeout_seconds", budget(args))
    error.details.setdefault("mutation_request_sent", False)
    status = error.details.get("http_status", error.details.get("status"))
    if isinstance(status, int) and 500 <= status <= 599:
        error.args = (
            f"Microsoft could not complete the authentication check (HTTP {status}). The saved profile was kept; try the read-only status check again after service recovery.",
        )
        error.details.setdefault(
            "reason", "service_timeout" if status in (502, 504) else "service_unavailable"
        )
    if error.code in (
        "auth_timeout",
        "auth_process_failed",
        "transport_error",
        "browser_unavailable",
        "login_required",
        "session_rejected",
        "api_error",
        "service_error",
        "runtime_error",
    ):
        error.details["recovery"] = recovery(args)
    return error


def browser_failure(error):
    return OutlookError(
        "browser_unavailable",
        "Could not complete browser authentication. Chrome may have disconnected or Microsoft may not have finished loading. The saved profile was kept.",
        reason="timeout" if type(error).__name__ == "TimeoutError" else "browser_connection",
    )


def _stop(process):
    # Also stop a driver holding the pipes after its worker has already exited.
    # Persistent Chrome belongs to a separate session, including on first login.
    try:
        os.killpg(process.pid, signal.SIGKILL)
    except ProcessLookupError:
        pass
    if process.poll() is None:
        try:
            process.wait(timeout=1)
        except subprocess.TimeoutExpired:
            pass


def _supervise(command, args, *, interval=5):
    """Never retry a worker. The parent owns the deadline, not browser JavaScript."""
    limit = budget(args)
    started = time.monotonic()
    deadline = started + limit
    stage = "starting"
    work_stage = stage
    output, pending = bytearray(), bytearray()
    process = None

    def report():
        if not getattr(args, "quiet", False):
            elapsed = int(time.monotonic() - started)
            print(
                f"[m365] {STAGES[stage]} ({elapsed}s elapsed; {limit}s limit).",
                file=sys.stderr,
                flush=True,
            )

    report()
    heartbeat = time.monotonic() + interval
    try:
        process = subprocess.Popen(
            command,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            start_new_session=True,
        )
        with selectors.DefaultSelector() as selector:
            for stream in (process.stdout, process.stderr):
                os.set_blocking(stream.fileno(), False)
                selector.register(stream, selectors.EVENT_READ)
            while selector.get_map() or process.poll() is None:
                now = time.monotonic()
                if now >= deadline:
                    raise OutlookError(
                        "auth_timeout",
                        "Authentication exceeded its command deadline; the saved browser profile was kept.",
                        phase=stage,
                        reason="deadline_exceeded",
                    )
                if now >= heartbeat:
                    report()
                    heartbeat = now + interval
                for key, _ in selector.select(min(0.2, deadline - now)):
                    chunk = os.read(key.fileobj.fileno(), 65536)
                    if not chunk:
                        selector.unregister(key.fileobj)
                        continue
                    if key.fileobj is process.stdout:
                        output.extend(chunk)
                        if len(output) > 1024 * 1024:
                            raise OutlookError(
                                "auth_process_failed",
                                "Authentication returned an oversized result.",
                            )
                    else:
                        pending.extend(chunk)
                        if len(pending) > 65536:
                            raise OutlookError(
                                "auth_process_failed", "Authentication progress exceeded its limit."
                            )
                        while b"\n" in pending:
                            line, _, pending = pending.partition(b"\n")
                            try:
                                text = line.decode("utf-8")
                                event = (
                                    json.loads(text[len(PREFIX) :])
                                    if text.startswith(PREFIX)
                                    else None
                                )
                                next_stage = event.get("stage") if isinstance(event, dict) else None
                            except (UnicodeError, ValueError):
                                next_stage = None
                            # Only constant stage descriptions reach stderr. Never
                            # forward worker warnings, exception text or headers.
                            if (
                                isinstance(next_stage, str)
                                and next_stage in STAGES
                                and next_stage != stage
                            ):
                                stage = next_stage
                                if stage != "cleanup":
                                    work_stage = stage
                                report()
                                heartbeat = time.monotonic() + interval
            try:
                result = json.loads(output)
            except (ValueError, UnicodeError):
                result = None
            if (
                isinstance(result, dict)
                and result.get("ok") is True
                and process.returncode == 0
                and isinstance(result.get("data"), dict)
            ):
                if not getattr(args, "quiet", False):
                    print("[m365] Authentication check finished.", file=sys.stderr, flush=True)
                return result["data"]
            error = result.get("error") if isinstance(result, dict) else None
            if (
                isinstance(error, dict)
                and isinstance(error.get("code"), str)
                and isinstance(error.get("message"), str)
            ):
                raise OutlookError(
                    error["code"],
                    error["message"],
                    **{k: v for k, v in error.items() if k not in ("code", "message")},
                )
            raise OutlookError(
                "auth_process_failed", "Authentication ended without a valid result."
            )
    except OutlookError as error:
        failed_stage = stage if error.code == "auth_timeout" else work_stage
        raise annotate(error, args, failed_stage, time.monotonic() - started) from None
    except OSError:
        raise annotate(
            OutlookError("auth_process_failed", "Could not run the authentication check."),
            args,
            stage,
            time.monotonic() - started,
        ) from None
    finally:
        if process:
            _stop(process)
            process.stdout.close()
            process.stderr.close()


def run(argv, args):
    if not is_auth_command(args):
        raise OutlookError(
            "usage_error", "Only login/status commands can use the authentication worker."
        )
    return _supervise([sys.executable, "-m", "m365_cli.auth_worker", *argv], args)
