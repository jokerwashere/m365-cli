"""Opt-in free/busy canary: one self-only appointment, optional read-only colleagues."""

import argparse
import json
import time
from datetime import UTC, datetime, timedelta

from .errors import OutlookError
from .live_labels import LabelCanary
from .scheduling import utc_string
from .session import command_lock
from .validation import email


class AvailabilityCanary(LabelCanary):
    # Reuse the JSON CLI invocation/journal harness, not the labels fixture scenario.
    suite, label = "availability", "availability"

    def availability(self, colleagues):
        def call(*words):
            return self.invoke(None, *words)

        first = datetime.now(UTC).replace(hour=9, minute=0, second=0, microsecond=0) + timedelta(
            days=14
        )
        window = (
            "--start",
            utc_string(first),
            "--end",
            utc_string(first + timedelta(days=2)),
            "--time-zone",
            "UTC",
            "--interval",
            "15",
        )
        finder = ("--duration", "30", "--weekdays", "all", "--max-results", "1")
        baseline = call("calendar", "find-slots", *window, *finder)
        self.require("self.available_fixture_slot", bool(baseline["slots"]))
        chosen = baseline["slots"][0]
        exact = (
            "--start",
            chosen["start"],
            "--end",
            chosen["end"],
            "--time-zone",
            "UTC",
            "--interval",
            "15",
        )
        cleanup = []

        def wait_view(expected):
            for attempt in range(6):
                result = call("calendar", "availability", *exact)
                if result["attendees"][0]["availability_view"] == expected:
                    return result
                if attempt < 5:
                    time.sleep(2)  # Read propagation polling only; mutations are never retried.
            self.require("freebusy.propagation." + expected, False)

        try:
            fixture = call(
                "calendar",
                "create",
                "--subject",
                self.marker,
                "--start",
                chosen["start"],
                "--end",
                chosen["end"],
                "--show-as",
                "Busy",
            )
            cleanup.append(self.track("calendar", "delete", fixture["id"]))
            full = call("calendar", "get", fixture["id"])
            self.require(
                "fixture.no_attendees",
                not full.get("RequiredAttendees") and not full.get("OptionalAttendees"),
            )
            data = wait_view("22")
            self.require(
                "busy.positive_control", data["attendees"][0]["blocks"][0]["status"] == "busy"
            )
            fine = call("calendar", "availability", *exact, "--interval", "5")
            self.require(
                "busy.five_minute_resolution",
                fine["attendees"][0]["availability_view"] == "222222",
            )
            self.require("snapshot.no_event_details", self.marker not in json.dumps(data))
            self.require(
                "busy.blocks_suggestions",
                not call("calendar", "find-slots", *exact, *finder)["slots"],
            )
            call("calendar", "update", fixture["id"], "--show-as", "Tentative")
            data = wait_view("11")
            self.require(
                "tentative.positive_control",
                data["attendees"][0]["blocks"][0]["status"] == "tentative",
            )
            self.require(
                "tentative.blocks_by_default",
                not call("calendar", "find-slots", *exact, *finder)["slots"],
            )
            self.require(
                "tentative.explicit_opt_in",
                len(call("calendar", "find-slots", *exact, *finder, "--allow-tentative")["slots"])
                == 1,
            )
            call("calendar", "update", fixture["id"], "--show-as", "OOF")
            data = wait_view("33")
            self.require(
                "oof.positive_control",
                data["attendees"][0]["blocks"][0]["status"] == "out_of_office",
            )
            self.require(
                "oof.blocks_even_with_tentative",
                not call("calendar", "find-slots", *exact, *finder, "--allow-tentative")["slots"],
            )
            call("calendar", "update", fixture["id"], "--show-as", "Free")
            data = wait_view("00")
            self.require("free.positive_control", data["complete"])
            free = call("calendar", "find-slots", *exact, *finder)
            self.require(
                "free.suggested_exact_duration",
                len(free["slots"]) == 1 and free["slots"][0]["duration_minutes"] == 30,
            )
            self.require("readonly.never_books", free["booking_submitted"] is False)
            # A random non-existent address in the test account's domain is a lookup,
            # never a message recipient. Keep the exact address private in the journal.
            missing = (
                "m365-cli-missing-"
                + self.marker.rsplit(" ", 1)[1]
                + "@"
                + self.account.split("@", 1)[1]
            )
            mixed = call("calendar", "availability", *exact, "--optional-attendee", missing)
            self.require("partial.keeps_known_self", mixed["attendees"][0]["status"] == "ok")
            self.require(
                "partial.recipient_error",
                mixed["attendees"][1]["status"] == "unavailable" and mixed["complete"] is False,
            )
            self.require(
                "unknown.required_blocks",
                not call("calendar", "find-slots", *exact, *finder, "--attendee", missing)["slots"],
            )
            optional = call(
                "calendar", "find-slots", *exact, *finder, "--optional-attendee", missing
            )
            self.require(
                "unknown.optional_not_counted_free",
                len(optional["slots"]) == 1
                and optional["slots"][0]["optional_available_count"] == 0,
            )
            strict = call("calendar", "find-slots", *exact, *finder, "--respect-working-hours")
            hours = strict["attendees"][0]["working_hours"]
            if not hours["usable"]:
                self.require(
                    "unknown.working_hours_blocks",
                    not strict["slots"] and bool(strict["unknown_required_attendees"]),
                )
            if colleagues:
                flags = [v for address in colleagues for v in ("--attendee", address)]
                group = call("calendar", "availability", *window, *flags)
                self.require(
                    "colleagues.availability_read",
                    all(p["status"] == "ok" for p in group["attendees"]),
                )
                slots = call("calendar", "find-slots", *window, *finder, *flags)
                self.require(
                    "colleagues.all_required_free",
                    all(
                        all(
                            p["status"] == "free"
                            for p in slot["attendees"]
                            if p["role"] == "required"
                        )
                        for slot in slots["slots"]
                    ),
                )
                self.journal["colleague_count"] = len(colleagues)
                self.save()
        finally:
            self.clean(None, cleanup)


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--account", required=True)
    p.add_argument("--allow-test-writes", action="store_true", required=True)
    p.add_argument("--attendee", action="append", default=[])
    args = p.parse_args()
    email(args.account)
    for address in args.attendee:
        email(address)
    with command_lock():
        suite = AvailabilityCanary(args.account)
        try:
            suite.availability(args.attendee)
            suite.journal["state"] = "passed"
        except Exception as exc:
            suite.journal.update(state="failed", code=getattr(exc, "code", "unexpected_error"))
            raise
        finally:
            suite.journal["cleanup_complete"] = all(
                c["state"] == "cleaned" for c in suite.journal["cleanup"]
            )
            suite.journal["limitations"] = (
                "Only one self-owned appointment is changed and moved to Deleted Items. No invitations or messages. Optional colleagues are read-only. Cross-tenant sharing, rooms, distribution groups and individual custom working-hour zones are not inferred from this pass. Read propagation polling is bounded; writes are never retried."
            )
            suite.save()
        if not suite.journal["cleanup_complete"]:
            raise OutlookError("cleanup_pending", "Availability canary cleanup needs review.")
        print(
            json.dumps(
                {"passed": True, "journal": str(suite.path), "checks": len(suite.journal["checks"])}
            )
        )


if __name__ == "__main__":
    main()
