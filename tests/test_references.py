"""Synthetic reference identity, isolation, link targets and acknowledged-write regressions."""

import json
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
from urllib.parse import parse_qs, urlsplit

from m365_cli.cli import parser, run
from m365_cli.errors import OutlookError
from m365_cli.operations import execute
from m365_cli.references import References, Registry, get, outlook_url, planner_url, todo_url
from m365_cli.teams_api import normalize_message

ACCOUNT = "reader@example.com"
TENANT = "00000000-0000-0000-0000-000000000001"


class ReferenceTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.env = patch.dict("os.environ", {"M365_HOME": self.tmp.name})
        self.env.start()
        self.addCleanup(self.env.stop)

    def refs(self, tenant=TENANT):
        return References(ACCOUNT, tenant)

    def task(self, identifier="task", **extra):
        r = self.refs()
        item = r.task(
            {"id": identifier, "source": "todo", "resource": "task", **extra},
            "todo",
            "task",
            {"list_id": "list"},
        )
        return r.registry.finish(item)

    def test_stable_across_process_instances_updates_and_case(self):
        a = self.task()
        b = self.task(title="Renamed")
        self.assertEqual(a["stable_id"], b["stable_id"])
        self.assertEqual(get(a["stable_id"], ACCOUNT.upper())["id"], "task")
        self.assertTrue(get(a["stable_id"], ACCOUNT)["cached"])
        self.assertFalse(get(a["stable_id"], ACCOUNT)["existence_verified"])

    def test_move_alias_persists_and_copy_stays_distinct(self):
        a = self.task("old")
        b = self.task("new", previous_id="old", list_id="destination")
        self.assertEqual(a["stable_id"], b["stable_id"])
        saved = get(a["stable_id"], ACCOUNT)
        self.assertEqual(saved["id"], "new")
        self.assertEqual(saved["locator"]["list_id"], "destination")
        self.assertEqual(saved["web_url"], todo_url("task", "new"))
        self.assertNotEqual(b["stable_id"], self.task("copy")["stable_id"])

    def test_same_id_in_other_account_tenant_or_source_is_distinct(self):
        first = self.task()["stable_id"]
        for account, tenant, source in [
            ("other@example.com", TENANT, "todo"),
            (ACCOUNT, "other", "todo"),
            (ACCOUNT, TENANT, "planner"),
        ]:
            r = References(account, tenant)
            item = r.task({"id": "task", "planId": "p"}, source, "task", {})
            self.assertNotEqual(item["stable_id"], first)
        with self.assertRaises(OutlookError) as caught:
            get(first, "other@example.com")
        self.assertEqual(caught.exception.code, "reference_not_found")

    def test_child_identity_is_parent_scoped_and_order_independent(self):
        r = self.refs()
        one = r.task({"id": "step"}, "todo", "step", {"task_id": "one"})
        two = r.task({"id": "step"}, "todo", "step", {"task_id": "two"})
        r.task({"id": "one"}, "todo", "task", {})
        again = r.task({"id": "step"}, "todo", "step", {"task_id": "one"})
        self.assertNotEqual(one["stable_id"], two["stable_id"])
        self.assertEqual(one["stable_id"], again["stable_id"])
        self.assertEqual(one["reference"]["url_kind"], "parent")
        self.assertEqual(one["web_url"], todo_url("task", "one"))

    def test_maps_remain_provider_payload_and_sidecars_reference_parent(self):
        checklist = {"entry": {"title": "One", "isChecked": False}}
        result = self.refs().task(
            {"id": "task", "checklist": checklist},
            "planner",
            "details",
            {"task_id": "task", "plan_id": "plan"},
        )
        self.assertEqual(result["checklist"], checklist)
        child = result["child_references"]["checklist"]["entry"]
        self.assertEqual(child["reference"]["url_kind"], "parent")
        self.assertEqual(child["web_url"], result["web_url"])

    def test_no_bodies_titles_or_credentials_persist(self):
        item = self.task(
            title="private title", body={"content": "private body"}, access_token="secret"
        )
        raw = next((Path(self.tmp.name) / "references").glob("*.json")).read_text()
        for secret in ("private title", "private body", "secret", "access_token"):
            self.assertNotIn(secret, raw)
        self.assertTrue(item["stable_id"])
        self.assertEqual(
            next((Path(self.tmp.name) / "references").glob("*.json")).stat().st_mode & 0o777, 0o600
        )

    def test_corrupt_registry_refuses_before_action(self):
        r = Registry(ACCOUNT)
        r.path.parent.mkdir(parents=True)
        r.path.write_text('{"revision":999}')
        with self.assertRaises(OutlookError) as caught:
            Registry(ACCOUNT)
        self.assertEqual(caught.exception.code, "reference_state_invalid")

    def test_storage_failure_preserves_ack_and_does_not_retry(self):
        r = self.refs()
        with patch("m365_cli.references.atomic_json", side_effect=OSError):
            value = r.registry.finish(r.task({"id": "task", "submitted": True}, "todo", "task", {}))
        self.assertTrue(value["submitted"])
        self.assertFalse(value["reference_warning"]["persisted"])

    def test_journal_replay_returns_same_references_without_action(self):
        action = unittest.mock.Mock(side_effect=lambda: self.task())
        a = execute("request", ACCOUNT, {"operation": "create"}, action)
        b = execute("request", ACCOUNT, {"operation": "create"}, action)
        self.assertEqual(a["stable_id"], b["stable_id"])
        self.assertTrue(b["replayed"])
        self.assertEqual(action.call_count, 1)

    def test_url_encoding_is_segment_safe(self):
        for url in (
            todo_url("task", "a/b+?&#="),
            outlook_url("mail", "a/b+?&#="),
            planner_url("task", "a/b?", TENANT, "p/q"),
        ):
            self.assertNotIn("#", url)
            self.assertIn("%2F", url)
        self.assertEqual(
            parse_qs(urlsplit(planner_url("plan", "p", TENANT)).query)["tid"], [TENANT]
        )

    def test_outlook_moves_copy_and_send_ack(self):
        args = SimpleNamespace(group="mail", command="get")
        a = self.refs().outlook({"id": "old"}, args)
        args.command = "move"
        b = self.refs().outlook({"id": "new", "original_id": "old"}, args)
        self.assertEqual(a["stable_id"], b["stable_id"])
        args.command = "copy"
        copy = self.refs().outlook({"id": "copy", "original_id": "new"}, args)
        self.assertNotEqual(a["stable_id"], copy["stable_id"])
        args.command = "send"
        sent = self.refs().outlook({"submitted": True}, args)
        self.assertIsNone(sent["stable_id"])
        self.assertIsNone(sent["web_url"])
        self.assertIn("do not resend", sent["reference_unavailable"])

    def test_deleted_outlook_does_not_advertise_stale_link(self):
        args = SimpleNamespace(group="mail", command="delete")
        result = self.refs().outlook({"id": "m", "moved_to_deleted_items": True}, args)
        self.assertIsNone(result["web_url"])
        self.assertEqual(result["reference"]["status"], "moved_to_deleted_items")

    def test_teams_message_ids_scoped_to_thread_and_reply_root_preserved(self):
        r = self.refs()
        one = r.teams_item(
            {"id": "123"}, "message", {"thread_id": "19:first@thread.v2", "channel": False}
        )
        two = r.teams_item(
            {"id": "123"}, "message", {"thread_id": "19:second@thread.v2", "channel": False}
        )
        self.assertNotEqual(one["stable_id"], two["stable_id"])
        self.assertEqual(
            json.loads(parse_qs(urlsplit(one["web_url"]).query)["context"][0]),
            {"contextType": "chat"},
        )
        msg = normalize_message(
            {"id": "reply", "replyChainId": "root", "properties": {"token": "SECRET"}}
        )
        msg = r.teams_item(
            msg, "message", {"thread_id": "channel", "channel": True, "group_id": "group"}
        )
        self.assertEqual(parse_qs(urlsplit(msg["web_url"]).query)["parentMessageId"], ["root"])
        self.assertNotIn("SECRET", json.dumps(msg))

    def test_incomplete_context_is_explicit(self):
        result = self.refs().task({"id": "task"}, "planner", "task", {})
        self.assertIsNone(result["web_url"])
        self.assertEqual(result["reference"]["url_kind"], "unavailable")
        group = self.refs().task({"id": "group"}, "todo", "group", {})
        self.assertEqual(group["reference"]["url_kind"], "application")

    def test_reference_get_offline_cli(self):
        item = self.task()
        result = run(
            parser().parse_args(["--account", ACCOUNT, "reference", "get", item["stable_id"]])
        )
        self.assertEqual(result["id"], item["id"])

    def test_recurrence_does_not_claim_occurrence_identity(self):
        item = self.task(successor_may_be_created=True)
        self.assertEqual(item["reference"]["occurrence_identity"], "relist_after_completion")

    def test_delta_order_and_duplicates_preserved(self):
        args = SimpleNamespace(group="todo", command="delta", list_id="list")
        result = self.refs().tasks(
            {"items": [{"id": "x", "title": "first"}, {"id": "x", "@removed": {}}]}, args
        )
        self.assertEqual(len(result["items"]), 2)
        self.assertEqual(result["items"][0]["stable_id"], result["items"][1]["stable_id"])
        self.assertEqual(get(result["items"][1]["stable_id"], ACCOUNT)["status"], "deleted")

    def test_parent_move_refreshes_child_link_but_requires_child_reread(self):
        r = self.refs()
        parent = r.task({"id": "old"}, "todo", "task", {"list_id": "a"})
        child = r.task({"id": "step"}, "todo", "step", {"task_id": "old", "list_id": "a"})
        r.registry.finish(parent)
        moved = self.task("new", previous_id="old", list_id="b")
        saved = get(child["stable_id"], ACCOUNT)
        self.assertEqual(saved["web_url"], moved["web_url"])
        self.assertEqual(saved["locator"]["task_id"], "new")
        self.assertTrue(saved["locator_needs_refresh"])
        r = self.refs()
        reread = r.task({"id": "step"}, "todo", "step", {"task_id": "new", "list_id": "b"})
        self.assertEqual(child["stable_id"], reread["stable_id"])
        self.assertNotIn("locator_needs_refresh", reread["reference"])

    def test_parent_delete_marks_child_reference(self):
        r = self.refs()
        parent = r.task({"id": "task"}, "todo", "task", {})
        child = r.task({"id": "step"}, "todo", "step", {"task_id": "task"})
        r.registry.finish(parent)
        self.task("task", deleted=True)
        saved = get(child["stable_id"], ACCOUNT)
        self.assertEqual(saved["status"], "parent_deleted")
        self.assertIsNone(saved["web_url"])

    def test_attachment_delete_preserves_original_parent_id_and_parent_state(self):
        args = SimpleNamespace(
            group="mail",
            command="attachment",
            attachment_command="delete",
            id="mail",
            attachment_id="attachment",
        )
        result = self.refs().outlook(
            {"id": "mail", "attachment_id": "attachment", "deleted": True}, args
        )
        self.assertEqual(result["id"], "mail")
        self.assertEqual(result["reference"]["status"], "observed")
        child = result["attachment_reference"]
        self.assertEqual(child["id"], "attachment")
        self.assertEqual(child["reference"]["status"], "deleted")
        self.assertEqual(child["reference"]["url_kind"], "parent")

    def test_todo_padding_is_preserved_for_native_router(self):
        self.assertTrue(todo_url("task", "AAAA==").endswith("/AAAA==/details"))
        self.assertTrue(todo_url("list", "AAAA==").endswith("/AAAA=="))

    def test_teams_reply_root_survives_readback_without_metadata(self):
        r = self.refs()
        first = r.teams_item(
            {"id": "reply", "reply_to": "root"},
            "message",
            {"thread_id": "channel", "channel": True, "group_id": "group"},
        )
        r.registry.finish(first)
        again = self.refs().teams_item(
            {"id": "reply"},
            "message",
            {"thread_id": "channel", "channel": True, "group_id": "group", "reply_to": None},
        )
        self.assertEqual(again["web_url"], first["web_url"])

    def test_calendar_container_identity_matches_folder_discovery(self):
        r = self.refs()
        calendar = r.outlook_item({"FolderId": {"Id": "calendar"}}, "calendar")
        folder = r.outlook_item({"id": "calendar", "FolderClass": "IPF.Appointment"}, "folder")
        self.assertEqual(calendar["stable_id"], folder["stable_id"])
        self.assertEqual(folder["reference"]["url_kind"], "application")

    def test_teams_roster_and_person_share_identity_without_changing_mri(self):
        r = self.refs()
        roster = r.teams_item({"id": "8:orgid:person"}, "person", {"parent_chat_id": "chat"})
        person = r.teams_item(
            {"objectId": "person", "mri": "8:orgid:person", "email": ACCOUNT}, "person", {}
        )
        self.assertEqual(roster["id"], "8:orgid:person")
        self.assertEqual(roster["stable_id"], person["stable_id"])
        self.assertEqual(roster["reference"]["url_kind"], "parent")
        self.assertEqual(person["reference"]["url_kind"], "action")

    def test_reference_contract_drift_is_safe_after_write(self):
        for mutation, code in [(False, "api_changed"), (True, "write_outcome_unknown")]:
            args = SimpleNamespace(group="calendar", command="calendars", mutation=mutation)
            with self.assertRaises(OutlookError) as caught:
                self.refs().outlook({"CalendarFolders": ["unexpected shape"]}, args)
            self.assertEqual(caught.exception.code, code)

    def test_calendar_group_entries_use_calendar_folder_id(self):
        args = SimpleNamespace(group="calendar", command="calendars")
        data = {
            "CalendarFolders": [{"FolderId": {"Id": "calendar"}}],
            "CalendarGroups": [
                {
                    "GroupId": "group",
                    "Calendars": [
                        {"CalendarFolderId": {"Id": "calendar"}, "ItemId": {"Id": "entry"}}
                    ],
                }
            ],
        }
        result = self.refs().outlook(data, args)
        self.assertEqual(
            result["CalendarFolders"][0]["stable_id"],
            result["CalendarGroups"][0]["Calendars"][0]["stable_id"],
        )

    def test_deleted_container_marks_known_tasks_and_children(self):
        r = self.refs()
        task = r.task({"id": "task"}, "planner", "task", {"plan_id": "plan"})
        child = r.task(
            {"id": "comment"}, "planner", "comment", {"task_id": "task", "plan_id": "plan"}
        )
        r.registry.finish(task)
        r = self.refs()
        r.registry.finish(r.task({"id": "plan", "deleted": True}, "planner", "plan", {}))
        for item in (task, child):
            saved = get(item["stable_id"], ACCOUNT)
            self.assertEqual(saved["status"], "parent_deleted")
            self.assertIsNone(saved["web_url"])

    def test_unknown_planner_delta_type_does_not_invent_resource_identity(self):
        args = SimpleNamespace(group="planner", command="delta")
        item = self.refs().tasks({"items": [{"id": "ambiguous", "@removed": {}}]}, args)["items"][0]
        self.assertIsNone(item["stable_id"])
        self.assertIsNone(item["web_url"])
        self.assertIn("delta_resource_type", item["reference_unavailable"])
