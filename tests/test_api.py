import unittest

from m365_cli.api import API, check_response, collection, recipients, validate_range
from m365_cli.errors import OutlookError
from m365_cli.session import is_outlook_url


def response(item):
    if item.get("ItemId"):
        item = {
            "ItemClass": "IPM.Appointment"
            if any(k in item for k in ("Start", "IsOrganizer", "CalendarItemType"))
            else "IPM.Note",
            **item,
        }
    return {"Body": {"ResponseMessages": {"Items": [{"ResponseCode": "NoError", "Items": [item]}]}}}


class FakeSession:
    def __init__(self, replies):
        self.replies = iter(replies)
        self.calls = []

    def request(self, action, body, **kwargs):
        self.calls.append((action, body))
        result = next(self.replies)
        if isinstance(result, Exception):
            raise result
        return result


class APITests(unittest.TestCase):
    def test_calendar_update_uses_owa_notification_field(self):
        item = {
            "ItemId": {"Id": "event", "ChangeKey": "ck"},
            "Start": "2026-10-07T09:00:00Z",
            "End": "2026-10-07T10:00:00Z",
        }
        s = FakeSession([response(item), response({}), response(item)])
        API(s).update("event", {"Subject": "Test"}, calendar=True)
        self.assertEqual(s.calls[1][1]["SendCalendarInvitationsOrCancellations"], "SendToNone")

    def test_nested_response_error_is_not_success(self):
        with self.assertRaises(OutlookError):
            check_response(
                {
                    "status": 200,
                    "body": {
                        "Body": {
                            "ResponseMessages": {
                                "Items": [
                                    {
                                        "ResponseClass": "Error",
                                        "ResponseCode": "ErrorInvalidIdMalformed",
                                    }
                                ]
                            }
                        }
                    },
                }
            )

    def test_paging_and_items_ignore_response_envelopes(self):
        data = {
            "Body": {
                "ResponseMessages": {
                    "Items": [
                        {
                            "RootFolder": {
                                "Items": [{"ItemId": {"Id": "mail-1"}, "Subject": "Test"}],
                                "IncludesLastItemInRange": False,
                                "IndexedPagingOffset": 25,
                                "TotalItemsInView": 100,
                            }
                        }
                    ]
                }
            }
        }
        result = collection(data)
        self.assertEqual([x["id"] for x in result["items"]], ["mail-1"])
        self.assertTrue(result["has_more"])
        self.assertEqual(result["next_offset"], 25)

    def test_dates_need_offsets_and_order(self):
        for start, end in [
            ("2026-10-07T09:00:00", "2026-10-07T10:00:00"),
            ("2026-10-07T10:00:00+02:00", "2026-10-07T09:00:00+02:00"),
        ]:
            with self.assertRaises(OutlookError):
                validate_range(start, end)
        validate_range("2026-10-07T09:00:00+02:00", "2026-10-07T08:00:00Z")

    def test_draft_never_sends(self):
        s = FakeSession([response({"ItemId": {"Id": "draft"}})])
        result = API(s).draft_create(subject="Test", body="Zażółć 📨", to=["me@example.com"])
        self.assertEqual(result["id"], "draft")
        action, body = s.calls[0]
        self.assertEqual(action, "CreateItem")
        self.assertEqual(body["MessageDisposition"], "SaveOnly")
        self.assertEqual(body["Items"][0]["Body"]["Value"], "Zażółć 📨")

    def test_recipient_header_injection_rejected(self):
        with self.assertRaises(OutlookError):
            recipients(["me@example.com\r\nBcc: other@example.com"])

    def test_send_requires_recipient_before_write(self):
        s = FakeSession([])
        with self.assertRaises(OutlookError) as exc:
            API(s).mail_send(subject="Test", body="Hello")
        self.assertEqual(exc.exception.code, "recipient_required")
        self.assertEqual(s.calls, [])

    def test_send_saves_copy_and_accepts_acknowledgement_without_id(self):
        s = FakeSession(
            [
                {
                    "Body": {
                        "ResponseMessages": {
                            "Items": [{"ResponseCode": "NoError", "ResponseClass": "Success"}]
                        }
                    }
                }
            ]
        )
        result = API(s).mail_send(subject="Test", body="Hello", to=["me@example.com"])
        self.assertTrue(result["submitted"])
        request = s.calls[0][1]
        self.assertEqual(request["MessageDisposition"], "SendAndSaveCopy")
        self.assertEqual(request["SavedItemFolderId"]["BaseFolderId"]["Id"], "sentitems")
        self.assertEqual(request["Items"][0]["ToRecipients"][0]["EmailAddress"], "me@example.com")

    def test_send_unknown_outcomes_never_retry(self):
        for reply in ({}, OutlookError("transport_error", "unknown outcome")):
            with self.subTest(reply=reply):
                s = FakeSession([reply])
                with self.assertRaises(OutlookError):
                    API(s).mail_send(subject="Test", body="Hello", to=["me@example.com"])
                self.assertEqual(len(s.calls), 1)

    def test_transport_failure_does_not_retry_write(self):
        s = FakeSession([OutlookError("transport_error", "unknown outcome")])
        with self.assertRaises(OutlookError):
            API(s).draft_create(subject="Test", body="Test")
        self.assertEqual(len(s.calls), 1)

    def test_edit_non_draft_fails_without_write(self):
        s = FakeSession([response({"ItemId": {"Id": "sent", "ChangeKey": "ck"}, "IsDraft": False})])
        with self.assertRaises(OutlookError):
            API(s).update("sent", {"Subject": "Changed"})
        self.assertEqual([call[0] for call in s.calls], ["GetItem"])

    def test_update_uses_change_key_and_allows_empty_text(self):
        item = {"ItemId": {"Id": "draft", "ChangeKey": "ck"}, "IsDraft": True, "Subject": "Before"}
        s = FakeSession([response(item), response({}), response({**item, "Subject": ""})])
        result = API(s).update("draft", {"Subject": ""})
        update = s.calls[1][1]
        self.assertEqual(update["ConflictResolution"], "NeverOverwrite")
        self.assertEqual(update["ItemChanges"][0]["ItemId"]["ChangeKey"], "ck")
        self.assertEqual(result["Subject"], "")

    def test_event_create_no_attendees_no_invites(self):
        s = FakeSession([response({"ItemId": {"Id": "event"}})])
        API(s).calendar_create(
            subject="Test", start="2026-10-07T09:00:00+02:00", end="2026-10-07T09:30:00+02:00"
        )
        self.assertEqual(s.calls[0][1]["SendMeetingInvitations"], "SendToNone")

    def test_meeting_update_requires_notify(self):
        event = {
            "ItemId": {"Id": "event", "ChangeKey": "ck"},
            "Start": "2026-10-07T09:00:00Z",
            "End": "2026-10-07T10:00:00Z",
            "IsOrganizer": True,
            "RequiredAttendees": [{"Mailbox": {"EmailAddress": "someone@example.com"}}],
        }
        s = FakeSession([response(event)])
        with self.assertRaises(OutlookError) as exc:
            API(s).update("event", {"Subject": "Update"}, calendar=True)
        self.assertEqual(exc.exception.code, "notification_required")
        self.assertEqual(len(s.calls), 1)

    def test_series_update_refused(self):
        s = FakeSession(
            [
                response(
                    {
                        "ItemId": {"Id": "event", "ChangeKey": "ck"},
                        "CalendarItemType": "RecurringMaster",
                    }
                )
            ]
        )
        with self.assertRaises(OutlookError) as exc:
            API(s).update("event", {"Subject": "Update"}, calendar=True)
        self.assertEqual(exc.exception.code, "recurring_series")

    def test_wrong_calendar_range_is_not_silently_empty(self):
        s = FakeSession(
            [
                response(
                    {
                        "ItemId": {"Id": "event"},
                        "Start": "2020-01-01T09:00:00Z",
                        "End": "2020-01-01T10:00:00Z",
                    }
                )
            ]
        )
        with self.assertRaises(OutlookError):
            API(s).calendar_list("2026-10-07T00:00:00Z", "2026-10-08T00:00:00Z")

    def test_only_real_https_outlook_hosts(self):
        for url in (
            "http://outlook.office.com/mail/",
            "https://outlook.office.com.evil.test/mail/",
            "https://evil.test/?outlook.office.com",
            "https://outlook.office.com@evil.test/",
            "https://outlook.office.com:8080/mail/",
        ):
            self.assertFalse(is_outlook_url(url), url)
        self.assertTrue(is_outlook_url("https://outlook.cloud.microsoft/mail/"))


if __name__ == "__main__":
    unittest.main()
