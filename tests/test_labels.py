"""Protocol fixtures and conservative failure behavior for categories and labels."""

import copy
import json
import tempfile
import unittest
from unittest.mock import Mock, patch

from test_api import FakeSession as OutlookSession
from test_api import response
from test_tasks import FakeSession
from test_transport import ready_session, reply

from m365_cli.api import API
from m365_cli.cli import parser, prepare, run
from m365_cli.errors import OutlookError
from m365_cli.labels import (
    Categories,
    PlannerLabels,
    applied_categories,
    color,
    label_name,
    palette,
    stamp,
)
from m365_cli.labels_cli import (
    category_assignment,
    category_dispatch,
    category_result,
    planner_dispatch,
)
from m365_cli.maintainer import guard, inventory
from m365_cli.references import References, get
from m365_cli.tasks_api import TasksAPI, review


def row(identifier="old", name="Alpha", value=7):
    return {"Id": identifier, "Name": name, "Color": value, "KeyboardShortcut": 0}


def catalog(*rows):
    return {"MasterList": list(rows), "WasSuccessful": True}


ACK = {"WasSuccessful": True, "ResponseCode": "NoError"}


class PlanAPI:
    def __init__(self, *, more=False, duplicates=0, fail=None):
        self.data = {
            "id": "plan",
            "@odata.etag": "v1",
            "categoryDescriptions": {"category1": "Alpha", "category4": "Other"},
        }
        self.rows = [
            {
                "id": "t1",
                "planId": "plan",
                "@odata.etag": "v2",
                "appliedCategories": {"category1": True, "category4": True},
            },
            {
                "id": "t2",
                "planId": "plan",
                "@odata.etag": "v3",
                "appliedCategories": {"category1": True},
            },
        ]
        self.more, self.duplicates, self.fail, self.calls = more, duplicates, fail, []

    def entity(self, family, path):
        if path.endswith("/details"):
            return copy.deepcopy(self.data)
        return copy.deepcopy(next(t for t in self.rows if path.endswith("/" + t["id"])))

    def scan(self, *args, **kwargs):
        return copy.deepcopy(self.rows), self.more, self.duplicates

    def patch(self, family, path, changes, source, kind, expected):
        self.calls.append((path, copy.deepcopy(changes), expected))
        if self.fail == len(self.calls):
            raise OutlookError("conflict", "Synthetic race")
        current = (
            self.data
            if kind == "labels"
            else next(t for t in self.rows if path.endswith("/" + t["id"]))
        )
        review(current, expected)
        for key, values in changes.items():
            current[key].update(values)
        return copy.deepcopy(current)


class LabelTests(unittest.TestCase):
    def error(self, code, callback):
        with self.assertRaises(OutlookError) as caught:
            callback()
        self.assertEqual(caught.exception.code, code)
        return caught.exception

    def test_palettes_do_not_mix_native_color_numbers_and_planner_slots(self):
        self.assertEqual(len(palette("category")["items"]), 26)
        self.assertEqual(len(palette("planner")["items"]), 25)
        self.assertEqual(color("blue", "category"), 7)
        self.assertEqual(color("blue", "planner"), "category5")
        for value in ("preset25", "#ff0000", "black-white", "25", True):
            self.error("invalid_color", lambda value=value: color(value, "category"))
        self.assertEqual(color("preset24", "category"), 24)
        self.assertEqual(color("none", "category"), -1)

    def test_names_reject_ambiguous_separators_and_utf16_overflow(self):
        for name in ("", " x", "x ", "a,b", "a;b", "x\n", "😀" * 128):
            self.error("invalid_argument", lambda name=name: label_name(name, category=True))
        self.assertEqual(label_name("Żółć 😀", category=True), "Żółć 😀")

    def test_local_dryrun_validation_and_aliases(self):
        with (
            tempfile.TemporaryDirectory() as state,
            patch.dict("os.environ", {"M365_HOME": state}),
            patch("m365_cli.labels_cli.Session") as session,
        ):
            for words in (["category"], ["todo", "category"]):
                result = run(
                    parser().parse_args(
                        [*words, "create", "--name", "Alpha", "--color", "blue", "--dry-run"]
                    )
                )
                self.assertEqual(result["plan"]["color_value"], 7)
            run(
                parser().parse_args(
                    [
                        "planner",
                        "label",
                        "recolor",
                        "category1",
                        "--plan",
                        "-" + "a" * 27,
                        "--color",
                        "blue",
                        "--dry-run",
                    ]
                )
            )
            self.assertEqual(len(run(parser().parse_args(["category", "colors"]))["items"]), 26)
            session.assert_not_called()
        for words in (
            ["category", "remove", "--mail", "m"],
            ["category", "remove", "a", "--name", "Alpha", "--mail", "m"],
            ["category", "apply", "a", "--task", "t"],
            ["category", "apply", "a", "--mail", "m", "--list", "l"],
            ["planner", "label", "delete", "category26", "--plan", "p"],
            ["category", "update", "a"],
        ):
            self.error("invalid_argument", lambda words=words: prepare(parser().parse_args(words)))

    def test_inventory_routes_all_label_leaves_to_canary(self):
        rows = inventory()["capabilities"]
        labels = [
            r
            for r in rows
            if r["command"].startswith(("category ", "todo category ", "planner label "))
        ]
        self.assertEqual(len(labels), 25)
        self.assertTrue(all(r["validation_suite"] == "labels" for r in labels))

    def test_category_aliases_share_circuit_breakers(self):
        for prefix, words in (
            ("category", ["todo", "category", "list"]),
            ("todo category update", ["category", "update", "old", "--color", "blue"]),
        ):
            args = parser().parse_args(words)
            with patch("m365_cli.maintainer.breakers", return_value={prefix: "contract changed"}):
                self.error("capability_disabled", lambda args=args: guard(args))

    def test_native_category_transport_is_explicit_and_allowlisted(self):
        s = ready_session()
        s.http.request.return_value = reply(catalog())
        body = {"request": {"__type": "GetMasterCategoryListRequest:#Exchange"}}
        s.request("GetMasterCategoryList", body, native=True)
        self.assertEqual(json.loads(s.http.request.call_args.kwargs["body"]), body)
        self.error("invalid_argument", lambda: s.request("CreateItem", {}, native=True))
        s.http.request.side_effect = [
            reply({"UserEmailAddress": s.account}),
            RuntimeError("secret"),
        ]
        self.error(
            "write_outcome_unknown", lambda: s.request("ApplyCategoryOperation", {}, native=True)
        )
        self.assertEqual(s.http.request.call_count, 3)

    def test_category_create_accepts_server_selected_identity(self):
        s = OutlookSession([catalog(row()), ACK, catalog(row(), row("new", "Beta", 15))])
        result = Categories(s).mutate("create", name="Beta", color_id=15)
        self.assertEqual(result["id"], "new")
        body = s.calls[1][1]["request"]
        self.assertEqual(body["AddCategoryList"][0]["Name"], "Beta")
        self.assertEqual(body["RemoveCategoryList"], [])

    def test_category_rename_preserves_local_reference_and_updates_locator(self):
        with tempfile.TemporaryDirectory() as state, patch.dict("os.environ", {"M365_HOME": state}):
            before = category_result(References("reader@example.com"), Categories(None).item(row()))
            s = OutlookSession([catalog(row()), ACK, catalog(row("new", "Beta"))])
            changed = Categories(s).mutate(
                "update", identifier="old", name="Beta", expected=stamp(row())
            )
            after = category_result(References("reader@example.com"), changed)
            self.assertEqual(before["stable_id"], after["stable_id"])
            self.assertEqual(get(before["stable_id"], "reader@example.com")["id"], "new")
            self.assertEqual(after["reference"]["url_kind"], "application")

    def test_category_stale_and_duplicate_fail_before_mutation(self):
        for expected, name, code in (
            ("stale", None, "conflict"),
            (None, "alpha", "category_exists"),
        ):
            s = OutlookSession([catalog(row())])
            self.error(
                code,
                lambda s=s, expected=expected, name=name: Categories(s).mutate(
                    "update" if expected else "create",
                    identifier="old" if expected else None,
                    expected=expected,
                    name=name,
                ),
            )
            self.assertEqual(len(s.calls), 1)

    def test_category_ack_without_change_is_unknown_never_retried(self):
        s = OutlookSession([catalog(), ACK, catalog()])
        self.error("write_outcome_unknown", lambda: Categories(s).mutate("create", name="Beta"))
        self.assertEqual(len(s.calls), 3)

    def test_category_drift_and_duplicate_definitions_fail_closed(self):
        for payload in (
            {"WasSuccessful": True},
            catalog(row(), row("two", "alpha")),
            catalog({**row(), "Color": True}),
            catalog({**row(), "Name": None}),
        ):
            self.error(
                "api_changed",
                lambda payload=payload: Categories(OutlookSession([payload])).listing(),
            )
        s = OutlookSession([catalog(row()), {"WasSuccessful": False}])
        self.error(
            "write_outcome_unknown", lambda: Categories(s).mutate("delete", identifier="old")
        )

    def test_category_delete_and_color_readback(self):
        s = OutlookSession([catalog(row()), ACK, catalog()])
        result = Categories(s).mutate("delete", identifier="old")
        self.assertTrue(result["deleted"])
        self.assertEqual(result["item_assignments"], "unchanged")
        s = OutlookSession([catalog(row()), ACK, catalog(row(value=24))])
        self.assertEqual(
            Categories(s).mutate("update", identifier="old", color_id=24)["preset"], "preset24"
        )

    def test_global_rename_failure_reports_completed_definition_with_reference(self):
        with tempfile.TemporaryDirectory() as state, patch.dict("os.environ", {"M365_HOME": state}):
            s = OutlookSession([catalog(row()), ACK, catalog(row("new", "Beta")), {}])
            args = prepare(parser().parse_args(["category", "update", "old", "--name", "Beta"]))
            exc = self.error(
                "partial_completion",
                lambda: category_dispatch(s, args, References("reader@example.com")),
            )
            self.assertEqual(exc.details["category"]["id"], "new")
            self.assertTrue(exc.details["category"]["stable_id"])

    def test_category_assignment_preserves_other_values_and_reviewed_etag(self):
        current = {"id": "t", "@odata.etag": "v1", "title": "Task", "categories": ["Other"]}
        after = {**current, "categories": ["Other", "Alpha"]}
        s = FakeSession([current, current, None, after])
        args = prepare(
            parser().parse_args(
                ["category", "apply", "c", "--task", "t", "--list", "l", "--if-version", "v1"]
            )
        )
        self.assertEqual(
            category_assignment(TasksAPI(s), args, "Alpha")["categories"], ["Other", "Alpha"]
        )
        self.assertEqual(s.calls[2][0][3], {"categories": ["Other", "Alpha"]})
        self.assertEqual(s.calls[2][0][4], {"If-Match": "v1"})

    def test_category_assignment_refuses_race_between_read_and_patch(self):
        current = {"id": "t", "@odata.etag": "v1", "categories": []}
        s = FakeSession([current, {**current, "@odata.etag": "v2"}])
        args = prepare(
            parser().parse_args(["category", "apply", "c", "--task", "t", "--list", "l"])
        )
        self.error("conflict", lambda: category_assignment(TasksAPI(s), args, "Alpha"))
        self.assertEqual(len(s.calls), 2)

    def test_event_personal_categories_do_not_notify_or_require_organizer(self):
        item = {
            "ItemId": {"Id": "e", "ChangeKey": "ck"},
            "IsOrganizer": False,
            "RequiredAttendees": [{"EmailAddress": {"Address": "other@example.com"}}],
        }
        s = OutlookSession(
            [response(item), response({}), response({**item, "Categories": ["Alpha"]})]
        )
        API(s).update("e", {"Categories": ["Alpha"]}, calendar=True, notify=True)
        self.assertEqual(s.calls[1][1]["SendCalendarInvitationsOrCancellations"], "SendToNone")
        self.assertEqual(s.calls[1][1]["ConflictResolution"], "NeverOverwrite")

    def test_missing_or_changed_planner_assignment_maps_cannot_look_unused(self):
        for value in (None, [], {"category1": "false"}, {"category26": True}):
            for writing in (False, True):
                self.error(
                    "write_outcome_unknown" if writing else "api_changed",
                    lambda value=value, writing=writing: applied_categories(
                        {"appliedCategories": value}, write=writing
                    ),
                )
        api = PlanAPI()
        del api.rows[0]["appliedCategories"]
        self.error(
            "api_changed",
            lambda: PlannerLabels(api, "plan").mutate("create", target="category5", name="Beta"),
        )
        self.assertEqual(api.calls, [])

    def test_planner_null_unset_slots_are_empty_names(self):
        api = PlanAPI()
        api.data["categoryDescriptions"]["category25"] = None
        result = PlannerLabels(api, "plan").listing()
        self.assertEqual(result["items"][-1]["name"], "")
        self.assertFalse(result["items"][-1]["defined"])

    def test_planner_unrecognized_slot_and_wrong_plan_fail_as_drift(self):
        for bad in ({"id": "other"}, {"categoryDescriptions": {"category26": "Future"}}):
            api = PlanAPI()
            api.data.update(bad)
            self.error("api_changed", lambda api=api: PlannerLabels(api, "plan").listing())

    def test_planner_unknown_first_write_reports_task_without_retry(self):
        api = PlanAPI()
        api.patch = Mock(side_effect=OutlookError("write_outcome_unknown", "Synthetic timeout"))
        exc = self.error(
            "write_outcome_unknown",
            lambda: PlannerLabels(api, "plan").mutate("recolor", "category1", target="category5"),
        )
        self.assertEqual(exc.details["uncertain_task_id"], "t1")
        self.assertEqual(api.patch.call_count, 1)

    def test_calendar_exact_end_boundary_is_filtered_without_hiding_wrong_ranges(self):
        inside = {
            "ItemId": {"Id": "inside"},
            "Start": "2031-02-03T09:00:00Z",
            "End": "2031-02-03T10:00:00Z",
        }
        boundary = {
            "ItemId": {"Id": "boundary"},
            "Start": "2031-02-04T00:00:00Z",
            "End": "2031-02-05T00:00:00Z",
            "IsAllDayEvent": True,
        }
        payload = {
            "Body": {
                "ResponseMessages": {
                    "Items": [{"ResponseCode": "NoError", "Items": [inside, boundary]}]
                }
            }
        }
        result = API(OutlookSession([payload])).calendar_list(
            "2031-02-03T00:00:00Z", "2031-02-04T00:00:00Z"
        )
        self.assertEqual([x["id"] for x in result["items"]], ["inside"])
        boundary["Start"] = "2031-02-04T00:00:01Z"
        self.error(
            "range_not_honored",
            lambda: API(OutlookSession([payload])).calendar_list(
                "2031-02-03T00:00:00Z", "2031-02-04T00:00:00Z"
            ),
        )

    def test_planner_recolor_moves_only_matching_assignments_and_exposes_new_slot(self):
        api = PlanAPI()
        result = PlannerLabels(api, "plan").mutate(
            "recolor", "category1", target="category5", expected="v1"
        )
        self.assertEqual(result["id"], "category5")
        self.assertEqual(result["previous_slot"], "category1")
        self.assertEqual(result["updated_task_count"], 2)
        self.assertEqual(
            api.calls[0][1], {"appliedCategories": {"category1": False, "category5": True}}
        )
        self.assertEqual(api.calls[0][2], "v2")
        self.assertTrue(api.rows[0]["appliedCategories"]["category4"])

    def test_planner_recolor_rejects_used_or_unnamed_but_applied_destinations(self):
        for target in ("category4", "category5"):
            api = PlanAPI()
            api.rows[0]["appliedCategories"]["category5"] = True
            self.error(
                "label_slot_in_use",
                lambda api=api, target=target: PlannerLabels(api, "plan").mutate(
                    "recolor", "category1", target=target
                ),
            )
            self.assertEqual(api.calls, [])

    def test_planner_recolor_and_create_require_complete_task_scan(self):
        for api in (PlanAPI(more=True), PlanAPI(duplicates=1)):
            self.error(
                "incomplete_scan",
                lambda api=api: PlannerLabels(api, "plan").mutate(
                    "create", target="category5", name="Beta"
                ),
            )
            self.assertEqual(api.calls, [])

    def test_planner_partial_failure_is_explicit_without_rollback_or_retry(self):
        for fail, completed in ((2, ["t1"]), (3, ["t1", "t2"])):
            api = PlanAPI(fail=fail)
            exc = self.error(
                "partial_completion",
                lambda api=api: PlannerLabels(api, "plan").mutate(
                    "recolor", "category1", target="category5"
                ),
            )
            self.assertEqual(exc.details["completed_task_ids"], completed)
            self.assertEqual(len(api.calls), fail)

    def test_planner_delete_requires_explicit_task_removal(self):
        api = PlanAPI()
        labels = PlannerLabels(api, "plan")
        self.error("label_in_use", lambda: labels.mutate("delete", "category1"))
        self.assertEqual(api.calls, [])
        result = labels.mutate("delete", "category1", remove_from_tasks=True)
        self.assertTrue(result["definition_cleared"])
        self.assertFalse(result["defined"])
        self.assertTrue(api.rows[0]["appliedCategories"]["category4"])

    def test_planner_task_must_belong_to_plan_and_label_must_be_defined(self):
        api = PlanAPI()
        api.rows[0]["planId"] = "other"
        args = prepare(
            parser().parse_args(
                ["planner", "label", "apply", "category1", "--plan", "plan", "--task", "t1"]
            )
        )
        self.error("invalid_argument", lambda: planner_dispatch(api, args))
        args.id = "category5"
        self.error("label_not_defined", lambda: planner_dispatch(api, args))
        self.assertEqual(api.calls, [])


if __name__ == "__main__":
    unittest.main()
