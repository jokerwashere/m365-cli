"""Synthetic outages, closed drivers and real process deadlines; no Microsoft traffic."""

import contextlib
import io
import json
import os
import selectors
import signal
import subprocess
import sys
import tempfile
import time
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from m365_cli import auth_control
from m365_cli.cli import main, parser
from m365_cli.errors import OutlookError
from m365_cli.session import Session


def args(**values):
    return SimpleNamespace(
        group="auth", command="status", mutation=False, timeout=1, quiet=False, **values
    )


class AuthControlTests(unittest.TestCase):
    def test_interrupt_stops_worker_before_returning_control(self):
        spawned = []
        original = subprocess.Popen

        def spawn(*a, **kw):
            process = original(*a, **kw)
            spawned.append(process)
            return process

        with (
            patch("m365_cli.auth_control.subprocess.Popen", side_effect=spawn),
            patch.object(selectors.DefaultSelector, "select", side_effect=KeyboardInterrupt),
            contextlib.redirect_stderr(io.StringIO()),
            self.assertRaises(KeyboardInterrupt),
        ):
            auth_control._supervise([sys.executable, "-c", "import time; time.sleep(60)"], args())
        self.assertEqual(len(spawned), 1)
        self.assertIsNotNone(spawned[0].poll())

    def test_deadline_stops_driver_holding_pipes_after_worker_exits(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "pid"
            heartbeat = Path(directory) / "heartbeat"
            driver = "import sys,time\nwith open(sys.argv[1], 'ab', buffering=0) as f:\n while True:\n  f.write(b'x'); time.sleep(0.02)"
            script = "import subprocess,sys; from pathlib import Path; p=subprocess.Popen([sys.executable,'-c',sys.argv[2],sys.argv[3]]); Path(sys.argv[1]).write_text(str(p.pid))"
            try:
                with (
                    contextlib.redirect_stderr(io.StringIO()),
                    self.assertRaises(OutlookError) as caught,
                ):
                    auth_control._supervise(
                        [sys.executable, "-c", script, str(path), driver, str(heartbeat)], args()
                    )
                self.assertEqual(caught.exception.code, "auth_timeout")
                size = heartbeat.stat().st_size
                self.assertGreater(size, 0)
                time.sleep(0.15)
                self.assertEqual(heartbeat.stat().st_size, size, "Worker driver was left running")
            finally:
                if path.exists():
                    try:
                        os.kill(int(path.read_text()), signal.SIGKILL)
                    except ProcessLookupError:
                        pass

    def test_auth_interruption_keeps_exit_130_and_never_suggests_unknown_write(self):
        stdout = io.StringIO()
        with (
            patch.object(sys, "argv", ["m365", "auth", "status"]),
            patch("m365_cli.auth_control.run", side_effect=KeyboardInterrupt),
            contextlib.redirect_stdout(stdout),
            self.assertRaises(SystemExit) as caught,
        ):
            main()
        self.assertEqual(caught.exception.code, 130)
        result = json.loads(stdout.getvalue())
        self.assertEqual(result["error"]["code"], "interrupted")
        self.assertIn("no service mutation", result["error"]["message"])

    def test_all_login_status_commands_offer_deadline_and_quiet(self):
        for group in ("auth", "teams", "tasks", "files"):
            for command in ("login", "status"):
                suffix = (
                    ["--site", "https://example.sharepoint.com/sites/test"]
                    if group == "files"
                    else []
                )
                parsed = parser().parse_args([group, command, *suffix])
                self.assertTrue(auth_control.is_auth_command(parsed))
                self.assertEqual(parsed.timeout, 300 if command == "login" else 30)
                parsed = parser().parse_args([group, command, *suffix, "--timeout", "7", "--quiet"])
                self.assertEqual(auth_control.budget(parsed), 7)
                self.assertTrue(parsed.quiet)
        for timeout in (-1, 0, 3601):
            parsed = parser().parse_args(["auth", "status", "--timeout", str(timeout)])
            with self.assertRaises(OutlookError):
                auth_control.budget(parsed)
        parsed = parser().parse_args(["auth", "login", "--timeout", "0"])
        self.assertEqual(auth_control.budget(parsed), 30)

    def test_supervisor_is_never_used_for_mutations_or_regular_reads(self):
        for argv in (["mail", "list"], ["mail", "delete", "synthetic", "--dry-run"]):
            parsed = parser().parse_args(argv)
            self.assertFalse(auth_control.is_auth_command(parsed))
            with (
                patch("m365_cli.auth_control._supervise") as supervise,
                self.assertRaises(OutlookError),
            ):
                auth_control.run(argv, parsed)
            supervise.assert_not_called()

    def test_worker_rejects_mutation_before_lock_or_dispatch(self):
        stdout = io.StringIO()
        with (
            patch.object(sys, "argv", ["m365", "mail", "delete", "synthetic"]),
            patch("m365_cli.cli.run") as dispatch,
            patch("m365_cli.cli.command_lock") as lock,
            contextlib.redirect_stdout(stdout),
            self.assertRaises(SystemExit),
        ):
            main(_auth_worker=True)
        self.assertEqual(json.loads(stdout.getvalue())["error"]["code"], "usage_error")
        dispatch.assert_not_called()
        lock.assert_not_called()

    def test_deadline_terminates_stalled_worker_and_reports_progress(self):
        script = 'import sys,time; print(\'M365_AUTH_PROGRESS {"stage":"browser_identity"}\',file=sys.stderr,flush=True); time.sleep(60)'
        stderr = io.StringIO()
        started = time.monotonic()
        with contextlib.redirect_stderr(stderr), self.assertRaises(OutlookError) as caught:
            auth_control._supervise([sys.executable, "-u", "-c", script], args(), interval=0.1)
        self.assertLess(time.monotonic() - started, 4)
        error = caught.exception.as_dict()
        self.assertEqual(error["code"], "auth_timeout")
        self.assertEqual(error["phase"], "browser_identity")
        self.assertEqual(error["timeout_seconds"], 1)
        self.assertFalse(error["mutation_request_sent"])
        self.assertIn("profile", error["recovery"])
        self.assertGreater(stderr.getvalue().count("elapsed"), 2)

    def test_timeout_keeps_detached_browser_process_alive(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "pid"
            script = "import subprocess,sys,time; from pathlib import Path; p=subprocess.Popen([sys.executable,'-c','import time; time.sleep(60)'],start_new_session=True,stdin=subprocess.DEVNULL,stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL); Path(sys.argv[1]).write_text(str(p.pid)); time.sleep(60)"
            try:
                with contextlib.redirect_stderr(io.StringIO()), self.assertRaises(OutlookError):
                    auth_control._supervise([sys.executable, "-c", script, str(path)], args())
                os.kill(int(path.read_text()), 0)
            finally:
                if path.exists():
                    os.kill(int(path.read_text()), signal.SIGKILL)

    def test_504_survives_pipe_warnings_and_cleanup_without_retry(self):
        result = {
            "ok": False,
            "error": {
                "code": "api_error",
                "message": "Outlook API returned HTTP 504.",
                "http_status": 504,
            },
        }
        script = 'import json,sys; print(\'M365_AUTH_PROGRESS {"stage":"browser_identity"}\',file=sys.stderr); print(\'pipe closed by peer SENSITIVE\',file=sys.stderr); print(\'M365_AUTH_PROGRESS {"stage":"cleanup"}\',file=sys.stderr); print(sys.argv[1]); sys.exit(1)'
        stderr = io.StringIO()
        with (
            contextlib.redirect_stderr(stderr),
            patch("m365_cli.auth_control.subprocess.Popen", wraps=subprocess.Popen) as spawn,
            self.assertRaises(OutlookError) as caught,
        ):
            auth_control._supervise([sys.executable, "-c", script, json.dumps(result)], args())
        spawn.assert_called_once()
        error = caught.exception.as_dict()
        self.assertEqual(error["http_status"], 504)
        self.assertEqual(error["reason"], "service_timeout")
        self.assertEqual(error["auth_stage"], "browser_identity")
        self.assertIn("service health", error["recovery"])
        self.assertNotIn("SENSITIVE", stderr.getvalue() + json.dumps(error))
        self.assertNotIn("pipe closed", stderr.getvalue())

    def test_closed_driver_without_result_is_secret_safe(self):
        script = "import sys; print('pipe closed by peer SENSITIVE',file=sys.stderr); print('SENSITIVE'); sys.exit(1)"
        stderr = io.StringIO()
        with contextlib.redirect_stderr(stderr), self.assertRaises(OutlookError) as caught:
            auth_control._supervise([sys.executable, "-c", script], args())
        self.assertEqual(caught.exception.code, "auth_process_failed")
        self.assertNotIn("SENSITIVE", stderr.getvalue() + json.dumps(caught.exception.as_dict()))

    def test_real_cli_json_stdout_and_quiet_with_no_saved_browser(self):
        with tempfile.TemporaryDirectory() as directory:
            for flags in ([], ["--quiet"]):
                process = subprocess.run(
                    [
                        sys.executable,
                        "-m",
                        "m365_cli.cli",
                        "auth",
                        "status",
                        "--timeout",
                        "5",
                        *flags,
                    ],
                    env={**os.environ, "M365_HOME": directory},
                    capture_output=True,
                    text=True,
                    timeout=8,
                )
                self.assertEqual(process.returncode, 0, process.stdout)
                self.assertEqual(
                    json.loads(process.stdout),
                    {"ok": True, "data": {"authenticated": False, "browser_running": False}},
                )
                if flags:
                    self.assertEqual(process.stderr, "")
                else:
                    self.assertIn("[m365]", process.stderr)

    def test_network_limits_only_change_inside_authentication_worker(self):
        with patch("m365_cli.auth_control._worker_deadline", None):
            self.assertEqual(auth_control.network_timeout(45), 45)
            self.assertEqual(auth_control.connect_timeout(45), 45)
        with patch("m365_cli.auth_control._worker_deadline", time.monotonic() + 30):
            self.assertEqual(auth_control.network_timeout(45), 15)
            self.assertEqual(auth_control.connect_timeout(45), 8)

    def test_cleanup_pipe_error_preserves_primary_error_and_clears_secrets(self):
        session = Session()
        driver = session.pw = Mock()
        session.headers = {"authorization": "SENSITIVE"}
        driver.stop.side_effect = BrokenPipeError("SENSITIVE")
        with (
            patch.object(Session, "__enter__", return_value=session),
            self.assertRaises(OutlookError) as caught,
        ):
            with session:
                raise OutlookError("api_error", "Outlook API returned HTTP 504.", http_status=504)
        self.assertEqual(caught.exception.details["http_status"], 504)
        self.assertEqual(session.headers, {})
        self.assertIsNone(session.pw)

    def test_browser_fetch_timeout_is_not_reported_as_bad_credentials(self):
        session = Session()
        session.page = Mock(url="https://outlook.office.com/mail/")
        session.page.evaluate.return_value = {"transport_error": True, "reason": "timeout"}
        session.context = Mock()
        with self.assertRaises(OutlookError) as caught:
            session._bootstrap("reader@example.com")
        self.assertEqual(caught.exception.code, "transport_error")
        self.assertEqual(caught.exception.details["phase"], "browser_identity")
        session.page.evaluate.assert_called_once()
