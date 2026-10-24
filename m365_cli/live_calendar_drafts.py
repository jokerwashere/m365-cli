"""Opt-in unsent-meeting lifecycle; two owned events, one authorized recipient."""

import argparse
import json
import time

from .cli import parser, run
from .errors import OutlookError
from .live_labels import LabelCanary
from .live_outlook import settled_fixture
from .session import command_lock
from .validation import email


class CalendarDraftCanary(LabelCanary):
    suite, label = "calendar-drafts", "unsent calendar"

    def scenario(self, recipient):
        cleanup = []
        tracked_mail = set()

        def call(*words):
            return self.invoke(None, *words)

        def read(item):
            return call("calendar", "get", item["id"], "--body-format", "html")

        def review(item):
            return settled_fixture(lambda: read(item), self.marker)

        def unsent(item, name):
            value = read(item)
            self.require(
                name + ".state",
                value.get("MeetingRequestWasSent") is False and value.get("IsOrganizer") is True,
            )
            self.require(name + ".reference", bool(value.get("stable_id") and value.get("web_url")))
            return value

        def messages():
            return [
                i
                for i in call("mail", "list", "--folder", "sentitems", "--limit", "100")["items"]
                if self.marker in i.get("Subject", "")
            ]

        def track_messages():
            for item in messages():
                if item["id"] not in tracked_mail:
                    cleanup.append(self.track("mail", "delete", item["id"]))
                    tracked_mail.add(item["id"])

        try:
            abandoned = call(
                "calendar",
                "create",
                "--subject",
                self.marker + " abandoned",
                "--start",
                "2031-02-17T14:00:00+01:00",
                "--end",
                "2031-02-17T14:30:00+01:00",
                "--attendee",
                recipient,
                "--save-only",
                "--body-format",
                "markdown",
                "--body",
                "**Prepared agenda**",
                "--show-as",
                "Free",
            )
            discarded = self.track("calendar", "delete", abandoned["id"], "--notify")
            cleanup.append(discarded)
            value = unsent(abandoned, "create")
            self.require(
                "create.attendees_and_body",
                value["RequiredAttendees"][0]["Mailbox"]["EmailAddress"].casefold()
                == recipient.casefold()
                and "<strong>Prepared agenda</strong>" in value["Body"]["Value"],
            )
            self.require(
                "create.no_submission",
                abandoned.get("notification_submitted") is False
                and abandoned.get("saved_only") is True,
            )
            self.reject(
                "conflict",
                "calendar",
                "update",
                abandoned["id"],
                "--save-only",
                "--subject",
                self.marker + " rejected",
                "--if-change-key",
                "deliberately-stale",
            )
            updated = call(
                "calendar",
                "update",
                abandoned["id"],
                "--clear-attendees",
                "--optional-attendee",
                recipient,
                "--save-only",
                "--body-format",
                "html",
                "--body",
                "<p><em>Revised agenda</em></p>",
                "--if-change-key",
                review(abandoned),
            )
            value = unsent(updated, "edit")
            self.require(
                "edit.roster_and_html",
                not value.get("RequiredAttendees")
                and value["OptionalAttendees"][0]["Mailbox"]["EmailAddress"].casefold()
                == recipient.casefold()
                and "<em>Revised agenda</em>" in value["Body"]["Value"],
            )
            self.require("edit.stable_identity", updated["stable_id"] == abandoned["stable_id"])
            self.require("save_only.no_sent_copy", not messages())
            result = call(
                "calendar", "delete", abandoned["id"], "--if-change-key", review(abandoned)
            )
            discarded["state"] = "cleaned"
            self.save()
            self.require("discard.no_cancellation", result.get("notification_submitted") is False)
            deleted = call("folder", "get", "deleteditems")["id"]
            self.require(
                "discard.deleted_items", read(abandoned)["ParentFolderId"]["Id"] == deleted
            )
            self.require("discard.no_sent_copy", not messages())
            self.reject("already_deleted", "calendar", "send", abandoned["id"])

            event = call(
                "calendar",
                "create",
                "--subject",
                self.marker + " send later",
                "--start",
                "2031-02-17T15:00:00+01:00",
                "--end",
                "2031-02-17T15:30:00+01:00",
                "--show-as",
                "Free",
            )
            sent_cleanup = self.track("calendar", "delete", event["id"], "--notify")
            cleanup.append(sent_cleanup)
            self.reject("attendee_required", "calendar", "send", event["id"])
            event = call(
                "calendar",
                "update",
                event["id"],
                "--attendee",
                recipient,
                "--save-only",
                "--body",
                "Disposable validation meeting. No action needed.",
                "--if-change-key",
                review(event),
            )
            unsent(event, "appointment_to_meeting")
            self.require("converted.no_sent_copy", not messages())
            self.reject(
                "notification_required",
                "calendar",
                "update",
                event["id"],
                "--subject",
                self.marker + " rejected",
            )
            self.reject(
                "conflict", "calendar", "send", event["id"], "--if-change-key", "deliberately-stale"
            )
            key = review(event)
            sent = call("calendar", "send", event["id"], "--if-change-key", key)
            send_request_id = self.marker + ":" + str(len(self.journal["operations"]))
            self.require(
                "send.acknowledged",
                sent.get("submitted") is True and sent.get("MeetingRequestWasSent") is True,
            )
            self.require(
                "send.stable_reference",
                sent["stable_id"] == event["stable_id"] and bool(sent["web_url"]),
            )
            # Replay only this already-completed operation, with its exact input.
            replay = run(
                parser().parse_args(
                    [
                        "--account",
                        self.account,
                        "calendar",
                        "send",
                        event["id"],
                        "--if-change-key",
                        key,
                        "--request-id",
                        send_request_id,
                    ]
                )
            )
            self.require("send.cached_replay", replay.get("replayed") is True)
            self.reject("meeting_already_sent", "calendar", "send", event["id"])
            self.reject(
                "meeting_already_sent",
                "calendar",
                "update",
                event["id"],
                "--save-only",
                "--subject",
                self.marker + " rejected",
            )
            self.reject("notification_required", "calendar", "delete", event["id"])
            found = []
            for attempt in range(8):
                found = messages()
                if found:
                    break
                if attempt < 7:
                    time.sleep(2)
            self.require(
                "send.exactly_one_invitation",
                len(found) == 1 and found[0].get("ItemClass") == "IPM.Schedule.Meeting.Request",
            )
            message = call("mail", "get", found[0]["id"])
            self.require(
                "send.recipient",
                {
                    r.get("Mailbox", r).get("EmailAddress", "").casefold()
                    for r in message.get("ToRecipients", [])
                }
                == {recipient.casefold()},
            )
            self.require(
                "send.sender",
                message.get("From", {})
                .get("Mailbox", message.get("From", {}))
                .get("EmailAddress", "")
                .casefold()
                == self.account.casefold(),
            )
            track_messages()
            call("calendar", "delete", event["id"], "--notify", "--if-change-key", review(event))
            sent_cleanup["state"] = "cleaned"
            self.save()
            self.require("sent.deleted_items", read(event)["ParentFolderId"]["Id"] == deleted)
            for attempt in range(8):
                if any(i.get("ItemClass") == "IPM.Schedule.Meeting.Canceled" for i in messages()):
                    break
                if attempt < 7:
                    time.sleep(2)
            self.require(
                "sent.cancellation",
                any(i.get("ItemClass") == "IPM.Schedule.Meeting.Canceled" for i in messages()),
            )
        finally:
            try:
                track_messages()
            finally:
                self.clean(None, cleanup)
            # Cleanup may send a cancellation if the scenario stopped early.
            # Only newly discovered messages may be cleaned in this second pass;
            # a failed first cleanup has an unknown outcome and must not replay.
            attempted = len(cleanup)
            track_messages()
            self.clean(None, cleanup[attempted:])


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--account", required=True)
    p.add_argument("--recipient", required=True)
    p.add_argument("--allow-test-writes", action="store_true", required=True)
    args = p.parse_args()
    email(args.account)
    email(args.recipient)
    with command_lock():
        suite = CalendarDraftCanary(args.account)
        try:
            suite.scenario(args.recipient)
            suite.journal["state"] = "passed"
        except Exception as error:
            suite.journal.update(state="failed", code=getattr(error, "code", "unexpected_error"))
            raise
        finally:
            suite.journal["cleanup_complete"] = (
                all(c["state"] == "cleaned" for c in suite.journal["cleanup"])
                and not suite.journal["unresolved"]
            )
            suite.journal["limitations"] = (
                "Single organizer events only. Absence of sends verified through native state and bounded Sent Items reads, not recipient mailbox access. One initial invitation and cancellation to the authorized recipient; sender-side fixtures cleaned. No recurring/delegated or sent-meeting draft edits."
            )
            suite.save()
        if not suite.journal["cleanup_complete"]:
            raise OutlookError("cleanup_pending", "Calendar draft validation cleanup needs review.")
        print(
            json.dumps(
                {"passed": True, "journal": str(suite.path), "checks": len(suite.journal["checks"])}
            )
        )


if __name__ == "__main__":
    main()
