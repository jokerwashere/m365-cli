"""Synthetic task contracts, identity isolation, unsafe-write and API-drift regression tests."""

import base64
import copy
import json
import tempfile
import time
import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

from m365_cli.cli import parser, plan, run
from m365_cli.errors import OutlookError
from m365_cli.maintainer import compare, guard, inventory
from m365_cli.task_validation import validate_payload
from m365_cli.tasks_api import TasksAPI, continuation, cursor_decode
from m365_cli.tasks_session import BASES, TasksSession, relative_path, token_valid

ACCOUNT = "test@example.com"
TENANT = "00000000-0000-0000-0000-000000000001"
OID = "00000000-0000-0000-0000-000000000002"


def token(family="graph", **extra):
    values = {
        "exp": time.time() + 3600,
        "aud": "https://graph.microsoft.com",
        "tid": TENANT,
        "oid": OID,
        "upn": ACCOUNT,
        "scp": "Tasks.ReadWrite",
    }
    values.update(extra)
    return "Bearer fake." + base64.urlsafe_b64encode(json.dumps(values).encode()).decode() + ".fake"


class FakeSession:
    account, tenant, oid = ACCOUNT, TENANT, OID

    def __init__(self, responses):
        self.responses, self.calls, self.last_meta = list(responses), [], {}

    def request(self, *args, **kwargs):
        self.calls.append((args, kwargs))
        result = self.responses.pop(0)
        if isinstance(result, Exception):
            raise result
        return copy.deepcopy(result)


class TasksTests(unittest.TestCase):
    def error(self, code, callback):
        with self.assertRaises(OutlookError) as caught:
            callback()
        self.assertEqual(caught.exception.code, code)

    def test_planner_opaque_ids_starting_with_hyphens_parse_verbatim(self):
        for prefix in ("-", "--"):
            identifier = prefix + "a" * (28 - len(prefix))
            a = parser().parse_args(
                ["planner", "task", "update", identifier, "--title", "x", "--dry-run"]
            )
            self.assertEqual(a.id, identifier)
            a = parser().parse_args(
                ["planner", "task", "create", "--plan", identifier, "--title", "x", "--dry-run"]
            )
            self.assertEqual(a.plan_id, identifier)

    def test_missing_title_cannot_be_filtered_into_empty_results(self):
        self.error(
            "api_changed",
            lambda: TasksAPI(FakeSession([{"value": [{"id": "a"}]}])).listing(
                "graph", "beta/me/planner/tasks", "planner", "task", query="missing"
            ),
        )

    def test_dry_run_no_auth_and_semantic_fingerprint(self):
        with (
            tempfile.TemporaryDirectory() as state,
            patch.dict("os.environ", {"M365_HOME": state}),
            patch("m365_cli.tasks_cli.TasksSession") as session,
        ):
            args = parser().parse_args(
                [
                    "todo",
                    "task",
                    "create",
                    "--list",
                    "list",
                    "--title",
                    "Task",
                    "--data",
                    '{"importance":"high"}',
                    "--dry-run",
                ]
            )
            result = run(args)
            self.assertTrue(result["dry_run"])
            self.assertEqual(result["plan"]["payload"]["importance"], "high")
            session.assert_not_called()
            self.assertEqual(plan(args)["payload"]["title"], "Task")

    def test_invalid_payloads_never_connect(self):
        bad = [
            ["todo", "task", "update", "id", "--list", "list", "--data", '{"assignments":{}}'],
            ["planner", "task", "create", "--plan", "p", "--title", "x", "--progress", "101"],
            ["todo", "task", "delete", "id", "--list", "list", "--if-version", "v"],
            [
                "todo",
                "task",
                "update",
                "id",
                "--list",
                "list",
                "--data",
                '{"dueDateTime":{"dateTime":"2031-02-03","timeZone":"Fake/Zone"}}',
            ],
            ["planner", "plan", "create", "--title", "x", "--container", "group"],
            ["planner", "task", "update", "id", "--data", '{"assignments":{"not-a-uuid":null}}'],
        ]
        with patch("m365_cli.tasks_cli.TasksSession") as session:
            for argv in bad:
                with self.subTest(argv=argv), self.assertRaises(OutlookError):
                    run(parser().parse_args(argv + ["--dry-run"]))
            session.assert_not_called()

    def test_payload_null_and_type_distinctions(self):
        self.assertEqual(
            validate_payload("planner.task", {"dueDateTime": None}), {"dueDateTime": None}
        )
        self.error("invalid_argument", lambda: validate_payload("planner.task", {"priority": True}))
        self.error(
            "invalid_argument",
            lambda: validate_payload(
                "todo.task",
                {
                    "recurrence": {
                        "pattern": {"type": "weekly", "interval": 1},
                        "range": {"type": "noEnd", "startDate": "2031-01-01"},
                    }
                },
            ),
        )
        self.error(
            "invalid_argument", lambda: validate_payload("planner.task", {"recurrence": None})
        )

    def test_credentials_bound_to_identity_audience_scope_and_expiry(self):
        self.assertTrue(token_valid(token(), "graph", ACCOUNT, TENANT, OID))
        for fields in (
            {"upn": "someone@example.com"},
            {"oid": "other"},
            {"tid": "other"},
            {"exp": time.time() - 1},
            {"aud": "https://evil.example"},
            {"scp": "Mail.Read"},
        ):
            self.assertFalse(token_valid(token(**fields), "graph", ACCOUNT, TENANT, OID))
        self.assertFalse(token_valid(token(), "todo", ACCOUNT, TENANT, OID))

    def test_route_and_continuation_boundaries(self):
        for path in (
            "https://evil.example/me",
            "/v1.0/me",
            "v1.0/../me",
            "v1.0/%2e%2e/me",
            "v1.0/me#fragment",
            "v1.0/me\\evil",
        ):
            self.error("invalid_argument", lambda path=path: relative_path("graph", path))
        initial = "v1.0/me/todo/lists/l/tasks"
        self.assertEqual(
            continuation(BASES["graph"] + initial + "?$skiptoken=x", "graph", initial),
            initial + "?$skiptoken=x",
        )
        for url in (
            "https://evil.example/" + initial,
            BASES["graph"] + "v1.0/me/messages",
            "https://graph.microsoft.com@evil.example/" + initial,
        ):
            self.error("api_changed", lambda url=url: continuation(url, "graph", initial))

    def test_delta_preserves_order_and_binds_cursor(self):
        initial = "v1.0/me/todo/lists/list/tasks/delta"
        link = BASES["graph"] + initial + "?$deltatoken=x"
        events = [{"id": "a", "title": "changed"}, {"id": "a", "@removed": {"reason": "deleted"}}]
        s = FakeSession([{"value": events, "@odata.deltaLink": link}])
        result = TasksAPI(s).delta("todo", "list")
        self.assertEqual(result["items"], events)
        self.assertEqual(
            cursor_decode(s, result["delta_cursor"], initial), initial + "?$deltatoken=x"
        )
        self.error(
            "invalid_cursor",
            lambda: cursor_decode(s, result["delta_cursor"], initial.replace("list/", "other/")),
        )
        s.tenant = "other"
        self.error("invalid_cursor", lambda: cursor_decode(s, result["delta_cursor"], initial))

    def test_scan_deduplicates_and_enforces_local_limit_when_server_ignores_top(self):
        s = FakeSession(
            [
                {
                    "value": [
                        {"id": "a", "title": "Alpha"},
                        {"id": "a", "title": "Alpha"},
                        {"id": "b", "title": "Beta"},
                    ]
                }
            ]
        )
        result = TasksAPI(s).listing("graph", "beta/me/planner/tasks", "planner", "task", limit=1)
        self.assertEqual(len(result["items"]), 1)
        self.assertEqual(result["duplicates_removed"], 1)
        self.assertEqual(result["next_offset"], 1)
        self.assertTrue(result["has_more"])

    def test_scan_budget_and_negative_search_are_explicit(self):
        path = "beta/me/planner/tasks"
        s = FakeSession(
            [
                {
                    "value": [{"id": "a", "title": "Alpha"}],
                    "@odata.nextLink": BASES["graph"] + path + "?$skiptoken=x",
                }
            ]
        )
        result = TasksAPI(s).listing("graph", path, "planner", "task", max_pages=1, query="missing")
        self.assertEqual(result["items"], [])
        self.assertTrue(result["truncated"])
        self.assertTrue(result["has_more"])
        self.assertFalse(result["coverage"]["notes_searched"])

    def test_changed_collection_never_becomes_empty(self):
        for response in ({}, {"value": {}}, {"value": [{"title": "lost id"}]}):
            self.error(
                "api_changed",
                lambda response=response: TasksAPI(FakeSession([response])).scan(
                    "graph", "beta/me/planner/tasks"
                ),
            )

    def test_pagination_cycle_fails(self):
        path = "beta/me/planner/tasks"
        data = {"value": [], "@odata.nextLink": BASES["graph"] + path}
        self.error("api_changed", lambda: TasksAPI(FakeSession([data])).scan("graph", path))

    def test_stale_update_no_write(self):
        s = FakeSession([{"id": "a", "@odata.etag": "new"}])
        self.error(
            "conflict",
            lambda: TasksAPI(s).patch(
                "graph", "beta/planner/tasks/a", {"title": "x"}, "planner", "task", "old"
            ),
        )
        self.assertEqual(len(s.calls), 1)

    def test_unknown_create_and_changed_readback_are_not_retried(self):
        s = FakeSession([{"unexpected": "shape"}])
        self.error(
            "write_outcome_unknown",
            lambda: TasksAPI(s).create("graph", "beta/planner/tasks", {}, "planner", "task"),
        )
        self.assertEqual(len(s.calls), 1)
        s = FakeSession([{"id": "a", "@odata.etag": "v"}, None, {}])
        self.error(
            "write_outcome_unknown",
            lambda: TasksAPI(s).patch(
                "graph", "beta/planner/tasks/a", {"title": "x"}, "planner", "task"
            ),
        )
        self.assertEqual(len(s.calls), 3)

    def test_delete_requires_verified_absence(self):
        s = FakeSession([{"id": "a", "@odata.etag": "v"}, None, {"id": "a"}])
        self.error(
            "write_outcome_unknown",
            lambda: TasksAPI(s).delete("graph", "beta/planner/tasks/a", "planner", "task"),
        )
        self.assertEqual(len(s.calls), 3)

    def test_todo_mapping_must_match_exact_ids(self):
        s = FakeSession([{"id": "a"}, {"Id": "other", "ParentFolderId": "list"}])
        self.error("api_changed", lambda: TasksAPI(s).todo_native("list", "a"))
        self.assertEqual(len(s.calls), 2)

    def test_builtin_and_unowned_lists_protected(self):
        for item in (
            {"id": "x", "wellknownListName": "defaultList", "isOwner": True},
            {"id": "x", "wellknownListName": "none", "isOwner": False},
        ):
            with self.assertRaises(OutlookError):
                TasksAPI(FakeSession([item])).todo_list_guard("x")

    def test_archive_reversal_requires_different_justification_before_write(self):
        s = FakeSession(
            [
                {
                    "id": "p",
                    "title": "plan",
                    "@odata.etag": "v",
                    "isArchived": True,
                    "archivalInfo": {"justification": "same"},
                }
            ]
        )
        self.error(
            "invalid_argument", lambda: TasksAPI(s).planner_archive("p", False, "same", None)
        )
        self.assertEqual(len(s.calls), 1)
        a = parser().parse_args(["planner", "plan", "archive", "p"])
        b = parser().parse_args(["planner", "plan", "unarchive", "p"])
        self.assertNotEqual(a.reason, b.reason)

    def test_unassigned_native_todo_task_may_omit_assignments(self):
        s = FakeSession([{"id": "t"}, {"Id": "t", "ParentFolderId": "l"}, {"Id": "assignment"}])
        self.assertEqual(TasksAPI(s).todo_assign("l", "t")["id"], "assignment")
        self.assertEqual(s.calls[-1][0][2], "POST")

    def test_recurrence_delete_needs_explicit_intent(self):
        s = FakeSession(
            [{"id": "a", "@odata.etag": "v", "recurrence": {"schedule": {"pattern": {}}}}]
        )
        self.error("recurrence_intent_required", lambda: TasksAPI(s).planner_delete_task("a"))
        self.assertEqual(len(s.calls), 1)

    def test_transport_no_retry_no_redirect_and_secret_safe_error(self):
        with patch("m365_cli.tasks_session.read_config", return_value={"account": ACCOUNT}):
            s = TasksSession()
        s.tokens, s.tenant, s.oid = {"graph": token()}, TENANT, OID
        s.ctx = SimpleNamespace(request=Mock())
        s.ctx.request.fetch.side_effect = RuntimeError("SECRET SESSION TOKEN")
        self.error(
            "write_outcome_unknown", lambda: s.request("graph", "beta/planner/tasks", "POST", {})
        )
        self.assertEqual(s.ctx.request.fetch.call_count, 1)
        self.assertEqual(s.ctx.request.fetch.call_args.kwargs["max_retries"], 0)
        self.assertEqual(s.ctx.request.fetch.call_args.kwargs["max_redirects"], 0)

    def test_transport_error_classification(self):
        for status, code in (
            (401, "session_rejected"),
            (403, "forbidden"),
            (404, "item_not_found"),
            (409, "conflict"),
            (412, "conflict"),
            (429, "rate_limited"),
            (500, "write_outcome_unknown"),
            (302, "write_outcome_unknown"),
        ):
            with patch("m365_cli.tasks_session.read_config", return_value={"account": ACCOUNT}):
                s = TasksSession()
            s.tokens, s.tenant, s.oid = {"graph": token()}, TENANT, OID
            response = Mock(status=status, headers={"retry-after": "30"})
            response.json.return_value = {"error": {"message": "SECRET"}}
            s.ctx = SimpleNamespace(request=Mock())
            s.ctx.request.fetch.return_value = response
            with self.assertRaises(OutlookError) as caught:
                s.request("graph", "beta/planner/tasks", "POST", {})
            self.assertEqual(caught.exception.code, code)
            self.assertNotIn("SECRET", json.dumps(caught.exception.as_dict()))


class MaintainerTests(unittest.TestCase):
    def test_inventory_all_modules_and_mutations(self):
        caps = {c["command"]: c for c in inventory()["capabilities"]}
        for path in (
            "mail send",
            "mail draft update",
            "mail attachment add",
            "folder delete",
            "calendar respond",
            "teams message delete",
            "todo task move",
            "planner task assign",
        ):
            self.assertTrue(caps[path]["mutation"], path)
        self.assertFalse(caps["tasks list"]["mutation"])
        self.assertTrue(all(c["live_availability"] == "not_checked" for c in caps.values()))

    def test_disable_prefix_guard_before_dry_run_and_auth(self):
        with (
            patch("m365_cli.maintainer.breakers", return_value={"todo task": "API drift"}),
            patch("m365_cli.tasks_cli.TasksSession") as session,
        ):
            with self.assertRaises(OutlookError) as caught:
                run(
                    parser().parse_args(
                        ["todo", "task", "create", "--list", "l", "--title", "x", "--dry-run"]
                    )
                )
            self.assertEqual(caught.exception.code, "capability_disabled")
            session.assert_not_called()
            guard(parser().parse_args(["maintainer", "capabilities"]))

    def test_comparison_does_not_claim_skipped_or_missing_probes_passed(self):
        baseline = {
            "report_type": "m365-capability-check",
            "contract_revision": 1,
            "checks": [{"name": "a", "status": "passed"}, {"name": "b", "status": "passed"}],
        }
        result = compare(
            {"contract_revision": 1, "checks": [{"name": "a", "status": "failed"}]}, baseline
        )
        self.assertEqual(result["regressions"], ["a"])
        self.assertEqual(result["not_rechecked"], ["b"])


if __name__ == "__main__":
    unittest.main()
