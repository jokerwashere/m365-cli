"""Opt-in live CLI validation. Only creates and cleans up uniquely labelled test items.

Requires explicit --writes and --recipient for mail/invitation traffic.
The private journal preserves IDs for cleanup after an interrupted validation.
"""

import argparse
import json
import os
import subprocess
import sys
import time
import uuid
from datetime import UTC, datetime, timedelta

from m365_cli.api import walk
from m365_cli.operations import atomic_json
from m365_cli.session import state_dir
from m365_cli.validation import iso


def settled_fixture(read, marker):
    """Read-only canary preparation; a later race still rejects the single write."""
    previous, stable = None, 0
    for attempt in range(8):
        item = read()
        if marker not in item.get("Subject", "") or not item.get("change_key"):
            raise RuntimeError("The labelled fixture or its change key could not be verified.")
        key = item["change_key"]
        stable = stable + 1 if key == previous else 1
        previous = key
        if stable == 3:
            return key
        if attempt < 7:
            time.sleep(1)
    raise RuntimeError("Fixture version did not settle; no update was attempted.")


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--account", required=True)
    p.add_argument(
        "--allow-test-writes", "--writes", dest="writes", action="store_true", required=True
    )
    p.add_argument("--recipient")
    p.add_argument(
        "--calendar-only",
        action="store_true",
        help="Run only calendar writes after common read checks",
    )
    args = p.parse_args()
    run_id = uuid.uuid4().hex[:12]
    marker = "[m365-cli validation " + run_id + "]"
    state = state_dir()
    journal = state / "validation-items.json"
    report_path = state / "validation-report.json"
    remaining = json.loads(journal.read_text()) if journal.exists() else []
    if remaining:
        raise RuntimeError(
            "Prior validation cleanup is pending; inspect the private journal before running again."
        )
    checks = []
    audit_path = state / ("outlook-live-" + run_id + ".json")
    audit = {"operations": [], "command_coverage": [], "state": "running"}
    for previous in state.glob("outlook-live-*.json"):
        old = json.loads(previous.read_text())
        if any(
            o.get("mutation")
            and (o.get("state") == "pending" or o.get("code") == "write_outcome_unknown")
            for o in old.get("operations", [])
        ):
            raise RuntimeError(
                "Prior Outlook validation has an unresolved write; inspect its audit journal."
            )

    def report(name, **details):
        item = {"check": name, **details}
        checks.append(item)
        atomic_json(
            report_path,
            {
                "run_id": run_id,
                "started_at": started,
                "checks": checks,
                "pending_cleanup": remaining,
            },
        )
        print(json.dumps(item), flush=True)

    def call(*command, expect_error=None, request_id=None):
        from m365_cli.cli import parser as command_parser
        from m365_cli.maintainer import command_path

        parsed = command_parser().parse_args(["--account", args.account, *command])
        path = command_path(parsed)
        if (
            not expect_error
            and "--if-change-key" not in command
            and path in ("mail update", "mail draft update", "calendar update")
        ):
            item_id = command[3] if path == "mail draft update" else command[2]
            key = settled_fixture(lambda: call(command[0], "get", item_id), marker)
            command = (*command, "--if-change-key", key)
        full = [sys.executable, "-m", "m365_cli.cli", "--account", args.account, *command]
        if parsed.mutation:
            request_id = request_id or "operation-" + str(len(audit["operations"]))
        if request_id:
            full += ["--request-id", run_id + "-" + request_id]
        op = {
            "command": command_path(parsed),
            "mutation": parsed.mutation,
            "state": "pending",
            "request_id": run_id + "-" + request_id if request_id else None,
        }
        audit["operations"].append(op)
        atomic_json(audit_path, audit)
        result = subprocess.run(full, capture_output=True, text=True, timeout=90)
        try:
            value = json.loads(result.stdout)
        except ValueError:
            raise RuntimeError("CLI did not return JSON: " + " ".join(command[:2])) from None
        op["state"] = "completed" if value.get("ok") else "failed"
        if value.get("ok"):
            op["id"] = value.get("data", {}).get("id")
            audit["command_coverage"] = sorted(set(audit["command_coverage"]) | {op["command"]})
        else:
            op["code"] = value.get("error", {}).get("code")
        atomic_json(audit_path, audit)
        if expect_error:
            if value.get("ok") or value.get("error", {}).get("code") != expect_error:
                raise RuntimeError(f"Expected {expect_error}, received {value.get('error')}")
            return value
        if result.returncode or not value.get("ok"):
            raise RuntimeError(" ".join(command[:3]) + ": " + json.dumps(value.get("error", value)))
        return value["data"]

    def remember(kind, item, notify=False):
        remaining.append({"kind": kind, "id": item["id"], "notify": notify, "marker": marker})
        atomic_json(journal, remaining)
        return item

    def forget(item_id):
        remaining[:] = [i for i in remaining if i["id"] != item_id]
        atomic_json(journal, remaining)

    def sent(subject, count=1):
        for _ in range(8):
            items = call("mail", "list", "--folder", "sentitems", "--limit", "100")["items"]
            hits = [i for i in items if i.get("Subject") == subject]
            if len(hits) >= count:
                for hit in hits:
                    detail = call("mail", "get", hit["id"])
                    destinations = [
                        o["EmailAddress"].casefold()
                        for key in ("ToRecipients", "CcRecipients", "BccRecipients")
                        for o in walk(detail.get(key, []))
                        if o.get("EmailAddress")
                    ]
                    assert destinations and set(destinations) == {args.recipient.casefold()}, (
                        "Unexpected test recipient"
                    )
                    assert (
                        detail["From"]["Mailbox"]["EmailAddress"].casefold()
                        == args.account.casefold()
                    ), "Unexpected test sender"
                return hits
            time.sleep(1)
        raise RuntimeError("Sent Items did not contain the expected test message: " + subject)

    started = datetime.now(UTC).isoformat()
    status = call("auth", "status")
    assert status["transport"] == {
        "mode": "direct_http",
        "browser_connected": False,
        "credential_storage": "command_memory",
    }
    report("transport.detached", passed=True)
    report("auth", passed=True)
    inbox = call("mail", "list", "--limit", "2")
    if inbox["items"]:
        before = inbox["items"][0]
        after = call("mail", "get", before["id"])
        assert after.get("IsRead") == before.get("IsRead")
    report("mail_read_and_unread_preservation", passed=True)
    call("mail", "search", "received:today", "--limit", "2")
    call("folder", "list", "--limit", "2")
    call("calendar", "calendars")
    today = datetime.now(UTC).replace(hour=0, minute=0, second=0, microsecond=0)
    call(
        "calendar",
        "list",
        "--start",
        today.isoformat(),
        "--end",
        (today + timedelta(days=7)).isoformat(),
    )
    report("read_search_and_discovery", passed=True)
    if not args.writes:
        return
    if not args.recipient:
        raise RuntimeError(
            "--recipient is required with --writes; validation sends email and invitations."
        )
    try:
        folder = remember(
            "folder", call("folder", "create", "--name", marker, request_id="folder-create")
        )
        call("folder", "update", folder["id"], "--name", marker + " renamed")
        cal_folder = remember(
            "folder", call("folder", "create", "--name", marker + " calendar", "--kind", "calendar")
        )
        report("folder_create_read_update", passed=True)
        if not args.calendar_only:
            draft = remember(
                "mail",
                call(
                    "mail",
                    "draft",
                    "create",
                    "--subject",
                    marker + " draft",
                    "--body",
                    "Zażółć 📨",
                    "--to",
                    args.recipient,
                    "--cc",
                    args.recipient,
                    "--bcc",
                    args.recipient,
                ),
            )
            data = call("mail", "draft", "get", draft["id"])
            assert data.get("IsDraft") is True
            call(
                "mail",
                "draft",
                "update",
                draft["id"],
                "--subject",
                marker + " draft-sent",
                "--clear-cc",
                "--clear-bcc",
                "--body",
                "M365 CLI validation. Unicode: Zażółć 📨",
                "--if-change-key",
                data["change_key"],
            )
            call(
                "mail",
                "draft",
                "update",
                draft["id"],
                "--subject",
                "MUST NOT APPLY",
                "--if-change-key",
                data["change_key"],
                expect_error="conflict",
            )
            fixture = state / "validation-file.txt"
            fixture.write_bytes("Attachment validation: Zażółć 📨\n".encode())
            call("mail", "attachment", "add", draft["id"], "--file", str(fixture))
            attached = call("mail", "attachment", "list", draft["id"])["items"]
            assert len(attached) == 1
            attachment_id = attached[0]["AttachmentId"]["Id"]
            output = state / ("validation-download-" + run_id + ".txt")
            call("mail", "attachment", "get", attachment_id, "--output", str(output))
            assert output.read_bytes() == fixture.read_bytes()
            call(
                "mail",
                "attachment",
                "get",
                attachment_id,
                "--output",
                str(output),
                expect_error="output_exists",
            )
            call("mail", "attachment", "delete", draft["id"], "--attachment-id", attachment_id)
            assert not call("mail", "attachment", "list", draft["id"])["items"]
            report("draft_crud_recipients_conflict_attachments", passed=True)
            call("mail", "draft", "send", draft["id"], request_id="draft-send")
            again = call("mail", "draft", "send", draft["id"], request_id="draft-send")
            assert again.get("replayed") is True
            original = sent(marker + " draft-sent")[0]
            forget(draft["id"])
            remember("mail", original)
            assert len(sent(marker + " draft-sent")) == 1
            report("draft_send_and_duplicate_prevention", passed=True)
            moved = call("mail", "move", original["id"], "--folder", folder["id"])
            forget(original["id"])
            remember("mail", moved)
            copied = remember("mail", call("mail", "copy", moved["id"], "--folder", folder["id"]))
            changed = call(
                "mail",
                "update",
                copied["id"],
                "--read",
                "false",
                "--importance",
                "High",
                "--category",
                "m365-cli-validation",
                "--flag",
                "Flagged",
            )
            assert changed.get("IsRead") is False and changed.get("Importance") == "High"
            assert changed.get("Flag", {}).get("FlagStatus") == "Flagged"
            call(
                "mail",
                "update",
                copied["id"],
                "--read",
                "true",
                "--clear-categories",
                "--flag",
                "Complete",
            )
            report("mail_move_copy_update", passed=True)
            reply = remember(
                "mail",
                call(
                    "mail",
                    "reply",
                    moved["id"],
                    "--to",
                    args.recipient,
                    "--body",
                    "Validation reply draft.",
                ),
            )
            assert call("mail", "draft", "get", reply["id"])["IsDraft"]
            call("mail", "draft", "delete", reply["id"])
            forget(reply["id"])
            for mode in ("reply", "forward"):
                extra = ["--all"] if mode == "reply" else []
                call(
                    "mail",
                    mode,
                    moved["id"],
                    "--to",
                    args.recipient,
                    "--body",
                    "M365 CLI validation " + mode,
                    "--send",
                    *extra,
                    request_id=mode,
                )
            call(
                "mail",
                "send",
                "--to",
                args.recipient,
                "--subject",
                marker + " direct",
                "--body",
                "M365 CLI direct-send validation.",
                request_id="direct-send",
            )
            assert (
                call(
                    "mail",
                    "send",
                    "--to",
                    args.recipient,
                    "--subject",
                    marker + " direct",
                    "--body",
                    "M365 CLI direct-send validation.",
                    request_id="direct-send",
                ).get("replayed")
                is True
            )
            for i in call("mail", "list", "--folder", "sentitems", "--limit", "100")["items"]:
                if marker in i.get("Subject", "") and not any(
                    i["id"] == r["id"] for r in remaining
                ):
                    remember("mail", i)
            assert len(sent(marker + " direct")) == 1
            search_found = False
            for _ in range(6):
                found = call("mail", "search", run_id, "--limit", "50")["items"]
                if found:
                    search_found = True
                    break
                time.sleep(2)
            if not search_found:
                raise RuntimeError(
                    "Search did not find the unique validation message after indexing wait."
                )
            report("mail_send_reply_forward_search", passed=True)
        start = "2031-02-03T09:00:00+01:00"
        end = "2031-02-03T09:30:00+01:00"
        event = remember(
            "calendar",
            call(
                "calendar",
                "create",
                "--subject",
                marker + " appointment",
                "--start",
                start,
                "--end",
                end,
                "--folder",
                cal_folder["id"],
            ),
        )
        updated = call(
            "calendar",
            "update",
            event["id"],
            "--start",
            "2031-02-03T10:00:00+01:00",
            "--end",
            "2031-02-03T10:30:00+01:00",
            "--location",
            "Validation location",
            "--body",
            "Temporary validation appointment",
            "--show-as",
            "Free",
        )
        assert iso(updated["Start"]) == iso("2031-02-03T10:00:00+01:00")
        assert updated["Location"]["DisplayName"] == "Validation location"
        assert updated["FreeBusyType"] == "Free"
        assert updated["Body"]["Value"] == "Temporary validation appointment"
        found = call(
            "calendar",
            "search",
            run_id,
            "--start",
            "2031-02-03T00:00:00+01:00",
            "--end",
            "2031-02-04T00:00:00+01:00",
            "--folder",
            cal_folder["id"],
        )
        assert len(found["items"]) == 1
        all_day = remember(
            "calendar",
            call(
                "calendar",
                "create",
                "--subject",
                marker + " all-day",
                "--start",
                "2031-02-04T00:00:00+01:00",
                "--end",
                "2031-02-05T00:00:00+01:00",
                "--all-day",
                "true",
                "--folder",
                cal_folder["id"],
            ),
        )
        checked = call("calendar", "get", all_day["id"])
        assert checked["IsAllDayEvent"] is True
        assert iso(checked["Start"]) == iso("2031-02-04T00:00:00+01:00")
        assert iso(checked["End"]) == iso("2031-02-05T00:00:00+01:00")
        report("calendar_appointment_crud_search_all_day", passed=True)
        meeting = remember(
            "calendar",
            call(
                "calendar",
                "create",
                "--subject",
                marker + " invitation",
                "--start",
                start,
                "--end",
                end,
                "--attendee",
                args.recipient,
            ),
            notify=True,
        )
        sent(marker + " invitation")
        changed_meeting = call(
            "calendar",
            "update",
            meeting["id"],
            "--subject",
            marker + " invitation-updated",
            "--notify",
            "--location",
            "Updated validation location",
            "--clear-attendees",
            "--optional-attendee",
            args.recipient,
        )
        assert not changed_meeting.get("RequiredAttendees")
        assert (
            changed_meeting["OptionalAttendees"][0]["Mailbox"]["EmailAddress"].casefold()
            == args.recipient.casefold()
        )
        sent(marker + " invitation-updated")
        call("calendar", "delete", meeting["id"], "--notify")
        forget(meeting["id"])
        cancelled = [
            i
            for i in call("mail", "list", "--folder", "sentitems", "--limit", "100")["items"]
            if run_id in i.get("Subject", "")
            and i.get("ItemClass") == "IPM.Schedule.Meeting.Canceled"
        ]
        assert cancelled, "No cancellation found in Sent Items"
        report("calendar_invitation_update_cancellation", passed=True)
    except Exception as error:
        audit["state"] = "failed"
        atomic_json(audit_path, audit)
        report("validation_failed", error=str(error))
        raise
    finally:
        # Pick up only messages containing this unpredictable marker, including notifications.
        try:
            for i in call("mail", "list", "--folder", "sentitems", "--limit", "100")["items"]:
                if marker in i.get("Subject", "") and not any(
                    i["id"] == r["id"] for r in remaining
                ):
                    remember("mail", i)
        except Exception as e:
            report("cleanup_scan_failed", error=str(e))
        for item in sorted(list(remaining), key=lambda i: i["kind"] == "folder"):
            try:
                if item["kind"] == "folder":
                    call("folder", "delete", item["id"])
                elif item["kind"] == "calendar":
                    call(
                        "calendar", "delete", item["id"], *(["--notify"] if item["notify"] else [])
                    )
                else:
                    call("mail", "delete", item["id"])
                forget(item["id"])
            except Exception as e:
                report("cleanup_failed", kind=item["kind"], error=str(e))
        report("cleanup", passed=not remaining, remaining=len(remaining))
        audit["cleanup_complete"] = not remaining
        atomic_json(audit_path, audit)
        if remaining:
            raise RuntimeError(
                "Cleanup incomplete; IDs retained in the private validation journal."
            )
    audit["state"] = "passed"
    audit["cleanup_complete"] = not remaining
    atomic_json(audit_path, audit)
    report("complete", passed=True)


if __name__ == "__main__":
    os.umask(0o077)
    try:
        main()
    except Exception as e:
        print(json.dumps({"passed": False, "error": str(e)}), flush=True)
        raise SystemExit(1) from None
