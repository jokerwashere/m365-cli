"""Critical-field removal and shape-change tests for the pre-existing adapters."""

import unittest

from test_api import FakeSession
from test_teams import FakeSession as TeamsSession

from m365_cli.api import API, check_response, collection
from m365_cli.errors import OutlookError
from m365_cli.teams_api import TeamsAPI, normalize_message


class DriftTests(unittest.TestCase):
    def test_missing_item_collection_in_valid_envelope_is_not_empty(self):
        with self.assertRaises(OutlookError):
            collection({"Body": {"ResponseMessages": {"Items": [{"ResponseCode": "NoError"}]}}})

    def test_additive_fields_allowed_but_changed_paging_rejected(self):
        good = {
            "RootFolder": {
                "Items": [],
                "IncludesLastItemInRange": True,
                "newFeature": {"anything": 42},
            }
        }
        self.assertFalse(collection(good)["has_more"])
        good["RootFolder"]["IncludesLastItemInRange"] = "false"
        with self.assertRaises(OutlookError):
            collection(good)

    def test_search_missing_results_not_silently_empty(self):
        with self.assertRaises(OutlookError):
            API(
                FakeSession([{"EntitySets": [{"ResultSets": [{"MoreResultsAvailable": False}]}]}])
            ).mail_search("test")

    def test_explicit_zero_total_search_may_omit_results(self):
        data = {"EntitySets": [{"ResultSets": [{"Total": 0, "MoreResultsAvailable": False}]}]}
        self.assertEqual(API(FakeSession([data])).mail_search("test")["items"], [])

    def test_calendar_discovery_requires_collection_and_success(self):
        for data in ({"WasSuccessful": True}, {"WasSuccessful": "true", "CalendarFolders": []}):
            with self.assertRaises(OutlookError):
                API(FakeSession([data])).calendars()

    def test_non_json_success_is_not_a_known_rejection(self):
        with self.assertRaises(OutlookError) as caught:
            check_response({"status": 200, "body": {"error": "non_json_response"}})
        self.assertEqual(caught.exception.code, "unexpected_response")

    def test_outlook_throttling_and_account_change_are_distinct(self):
        for response, code in (
            ({"status": 429, "body": {}, "retry_after": "7"}, "rate_limited"),
            ({"status": 409, "body": {"error": "account_mismatch"}}, "account_mismatch"),
        ):
            with self.assertRaises(OutlookError) as caught:
                check_response(response)
            self.assertEqual(caught.exception.code, code)
        with self.assertRaises(OutlookError):
            check_response({"status": 200, "body": {"ResponseCode": {"unexpected": "shape"}}})

    def test_teams_missing_channel_collection_and_message_id_rejected(self):
        with self.assertRaises(OutlookError):
            TeamsAPI(TeamsSession([{"chats": [], "teams": [{"id": "t"}]}])).channels("t")
        with self.assertRaises(OutlookError):
            normalize_message({"content": "lost id"})

    def test_teams_inventory_malformed_members_not_treated_as_empty(self):
        with self.assertRaises(OutlookError):
            TeamsAPI(TeamsSession([{"chats": [{"id": "t", "members": None}], "teams": []}])).chats()


if __name__ == "__main__":
    unittest.main()
