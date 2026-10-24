"""Regression checks for incompatibilities observed in the live Microsoft 365 session."""

import base64
import importlib.util
import unittest
from pathlib import Path

from test_api import FakeSession, response

from m365_cli.api import API, check_response
from m365_cli.errors import OutlookError


class LiveRegressionTests(unittest.TestCase):
    def test_attachment_download_omits_rejected_shape(self):
        s = FakeSession(
            [
                {
                    "Body": {
                        "Attachments": [
                            {
                                "__type": "FileAttachment:#Exchange",
                                "AttachmentId": {"Id": "attachment"},
                                "Name": "example.txt",
                                "Content": base64.b64encode(b"file data").decode(),
                            }
                        ]
                    }
                }
            ]
        )
        self.assertEqual(API(s).attachment_get("attachment")["content"], b"file data")
        self.assertNotIn("AttachmentShape", s.calls[0][1])

    def test_calendar_location_and_availability_use_owa_names(self):
        item = {
            "ItemId": {"Id": "event", "ChangeKey": "ck"},
            "Start": "2031-01-01T09:00:00Z",
            "End": "2031-01-01T10:00:00Z",
        }
        s = FakeSession([response(item), response({}), response(item)])
        API(s).update(
            "event",
            {"Location": {"DisplayName": "Room"}, "LegacyFreeBusyStatus": "Free"},
            calendar=True,
        )
        updates = s.calls[1][1]["ItemChanges"][0]["Updates"]
        self.assertEqual(updates[0]["Path"]["FieldURI"], "calendar:EnhancedLocation")
        self.assertEqual(updates[0]["Item"]["Location"]["DisplayName"], "Room")
        self.assertEqual(updates[1]["Path"]["FieldURI"], "calendar:LegacyFreeBusyStatus")
        self.assertEqual(updates[1]["Item"]["FreeBusyType"], "Free")

    def test_calendar_invitation_and_cancellation_fields_differ_from_update_field(self):
        item = {
            "ItemId": {"Id": "event", "ChangeKey": "ck"},
            "Start": "2031-01-01T09:00:00Z",
            "End": "2031-01-01T10:00:00Z",
            "IsOrganizer": True,
            "RequiredAttendees": [{}],
        }
        s = FakeSession([response(item), {"FolderId": {"Id": "deleted"}}, response({})])
        API(s).delete("event", calendar=True, notify=True)
        self.assertEqual(s.calls[-1][1]["SendMeetingCancellations"], "SendToAllAndSaveCopy")
        s = FakeSession([response(item)])
        API(s).calendar_create(
            subject="Test",
            start=item["Start"],
            end=item["End"],
            attendees=["test@example.com"],
            show_as="Free",
        )
        self.assertEqual(s.calls[0][1]["SendMeetingInvitations"], "SendToAllAndSaveCopy")
        self.assertEqual(s.calls[0][1]["Items"][0]["FreeBusyType"], "Free")

    def test_mail_cannot_mutate_tasks_contacts_or_unidentified_items(self):
        for cls in ("IPM.Task", "IPM.Contact", ""):
            s = FakeSession(
                [
                    response(
                        {
                            "ItemId": {"Id": "x", "ChangeKey": "ck"},
                            "ItemClass": cls,
                            "IsDraft": True,
                        }
                    )
                ]
            )
            with self.assertRaises(OutlookError) as error:
                API(s).update("x", {"Subject": "changed"})
            self.assertEqual(error.exception.code, "wrong_item_type")
            self.assertEqual(len(s.calls), 1)

    def test_generated_command_reference_is_current(self):
        root = Path(__file__).parents[1]
        spec = importlib.util.spec_from_file_location(
            "reference", root / "scripts" / "generate_reference.py"
        )
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        self.assertEqual((root / "docs" / "CLI.md").read_text(), module.render())

    def test_all_day_dates_use_mailbox_zone_and_handle_dst(self):
        from m365_cli.api import envelope
        from m365_cli.validation import calendar_range

        zone = "W. Europe Standard Time"
        calendar_range(
            "2026-10-25T00:00:00+02:00", "2026-10-26T00:00:00+01:00", all_day=True, time_zone=zone
        )
        calendar_range("2026-10-24T22:00:00Z", "2026-10-25T23:00:00Z", all_day=True, time_zone=zone)
        with self.assertRaises(OutlookError) as error:
            calendar_range(
                "2026-10-25T00:00:00Z", "2026-10-26T00:00:00Z", all_day=True, time_zone=zone
            )
        self.assertEqual(error.exception.code, "timezone_mismatch")
        self.assertEqual(
            envelope("CreateItem", {}, time_zone=zone)["Header"]["TimeZoneContext"][
                "TimeZoneDefinition"
            ]["Id"],
            zone,
        )

    def test_builtin_archive_and_deleted_folders_are_protected(self):
        from unittest.mock import Mock

        s = FakeSession([])
        api = API(s)
        current = {"id": "archive", "change_key": "ck", "TotalCount": 0, "ChildFolderCount": 0}
        api.folder_get = Mock(
            side_effect=lambda name: current if name == "target" else {"id": name}
        )
        with self.assertRaises(OutlookError) as error:
            api.folder_delete("target")
        self.assertEqual(error.exception.code, "protected_folder")
        current["id"] = "custom"
        current["ParentFolderId"] = {"Id": "deleteditems"}
        with self.assertRaises(OutlookError) as error:
            api.folder_delete("target")
        self.assertEqual(error.exception.code, "already_deleted")
        self.assertEqual(s.calls, [])

    def test_unknown_collection_is_not_reported_as_empty(self):
        from m365_cli.api import collection

        with self.assertRaises(OutlookError) as error:
            collection({"Body": {"UnexpectedShape": []}})
        self.assertEqual(error.exception.code, "unexpected_response")

    def test_exception_only_response_does_not_crash_error_handler(self):
        with self.assertRaises(OutlookError) as error:
            check_response({"status": 200, "body": {"ResponseClass": "Error"}})
        self.assertEqual(error.exception.code, "owa_error")


if __name__ == "__main__":
    unittest.main()
