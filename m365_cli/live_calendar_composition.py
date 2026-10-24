"""Opt-in rich-event canary: three labelled self-only appointments, then cleanup."""

import argparse
import json
import time
from pathlib import Path
from tempfile import TemporaryDirectory

from .errors import OutlookError
from .live_labels import LabelCanary
from .session import command_lock, state_dir
from .validation import email


class CalendarCompositionCanary(LabelCanary):
    suite, label = "calendar-composition", "calendar composition"

    def update(self, item, *words):
        # Outlook assistants can update a newly edited event asynchronously.
        # Poll reads only for a settled version, then attempt exactly one write.
        # This is canary fixture preparation, not a product write-retry policy.
        previous, stable = None, 0
        for attempt in range(8):
            current = self.invoke(None, "calendar", "get", item["id"], "--body-format", "html")
            key = current.get("change_key")
            if not isinstance(key, str) or not key:
                raise OutlookError("missing_change_key", "The fixture has no reviewable version.")
            stable = stable + 1 if key == previous else 1
            previous = key
            if stable == 3:
                return self.invoke(
                    None, "calendar", "update", item["id"], *words, "--if-change-key", key
                )
            if attempt < 7:
                time.sleep(1)
        raise OutlookError("conflict", "Fixture version did not settle; no update was attempted.")

    def composition(self):
        def call(*words):
            return self.invoke(None, "calendar", *words)

        def read(item, name):
            data = call("get", item["id"], "--body-format", "html")
            self.require(name + ".html_body", data["Body"]["BodyType"] == "HTML")
            self.require(name + ".reference", bool(data["stable_id"] and data["web_url"]))
            self.require(
                name + ".self_only",
                not data.get("RequiredAttendees") and not data.get("OptionalAttendees"),
            )
            return data

        cleanup = []
        try:
            with TemporaryDirectory(prefix="calendar-composition-", dir=state_dir()) as tmp:
                style = Path(tmp) / "style.json"
                style.write_text('{"font_family":"Georgia, serif","font_size_pt":11}')
                source = Path(tmp) / "agenda.md"
                source.write_text(
                    "# Agenda\n\n**Review** Zażółć 📅\nNext line\n\n- First\n- Second\n\n"
                    "| Item | State |\n| --- | --- |\n| Review | Ready |\n\n`code`",
                    encoding="utf-8",
                )
                self.require(
                    "styles.defaults",
                    call("styles")["defaults"]["font_family"].startswith("Helvetica Neue"),
                )
                preview = call(
                    "render",
                    "--body-file",
                    str(source),
                    "--body-format",
                    "markdown",
                    "--output",
                    str(Path(tmp) / "preview.html"),
                )
                self.require(
                    "preview.local",
                    preview["local_only"]
                    and not preview["submitted"]
                    and "<strong>Review</strong>" in preview["body"],
                )
                for format, body in (
                    ("text", ["--body", "<b>literal</b> **literal** Zażółć 📅"]),
                    (
                        "html",
                        [
                            "--body",
                            '<p onclick="x()"><em>HTML event</em> Zażółć 📅</p><script>removed_token</script>',
                        ],
                    ),
                    ("markdown", ["--body-file", str(source)]),
                ):
                    item = call(
                        "create",
                        "--subject",
                        self.marker + " " + format,
                        "--start",
                        "2031-02-03T14:00:00+01:00",
                        "--end",
                        "2031-02-03T14:30:00+01:00",
                        "--show-as",
                        "Free",
                        "--body-format",
                        format,
                        *body,
                    )
                    cleanup.append(self.track("calendar", "delete", item["id"]))
                    data = read(item, format)
                    value = data["Body"]["Value"]
                    plain = call("get", item["id"])
                    self.require(format + ".text_read", plain["Body"]["BodyType"] == "Text")
                    if format == "text":
                        self.require(
                            "text.literal", "<b>literal</b> **literal**" in plain["Body"]["Value"]
                        )
                    else:
                        self.require(
                            format + ".default_style", "Helvetica Neue" in value and "12pt" in value
                        )
                        self.require(format + ".unicode", "Zażółć 📅" in value)
                        self.require(
                            format + ".metadata", item["composition"]["input_format"] == format
                        )
                    if format == "html":
                        self.require(
                            "html.normalized",
                            "<em>HTML event</em>" in value
                            and "removed_token" not in value
                            and "onclick" not in value,
                        )
                    if format == "markdown":
                        self.require(
                            "markdown.structure",
                            all(
                                v in value
                                for v in (
                                    "<strong>Review</strong>",
                                    "<table",
                                    "<ul",
                                    "<br",
                                    "Menlo",
                                )
                            ),
                        )

                # Exercise replacements and omission on the last owned fixture.
                updated = self.update(
                    item,
                    "--body-format",
                    "markdown",
                    "--body",
                    "**Updated**",
                    "--style-file",
                    str(style),
                )
                value = updated["Body"]["Value"]
                self.require(
                    "update.html_readback",
                    updated["Body"]["BodyType"] == "HTML" and "<strong>Updated</strong>" in value,
                )
                self.require("update.custom_style", "Georgia" in value and "11pt" in value)
                self.update(
                    item,
                    "--subject",
                    self.marker + " renamed",
                )
                data = read(item, "omitted_body")
                self.require("update.omitted_preserves_body", data["Body"]["Value"] == value)
                updated = self.update(
                    item,
                    "--body-format",
                    "html",
                    "--style",
                    "minimal",
                    "--body",
                    "<p><strong>Minimal</strong></p>",
                )
                self.require(
                    "update.html_minimal",
                    "<strong>Minimal</strong>" in updated["Body"]["Value"]
                    and "Georgia" not in updated["Body"]["Value"],
                )
                updated = self.update(
                    item,
                    "--body",
                    "<b>literal replacement</b>",
                )
                self.require(
                    "update.literal_text", "<b>literal replacement</b>" in updated["Body"]["Value"]
                )
                updated = self.update(item, "--body", "")
                self.require("update.empty_clears", not updated["Body"]["Value"].strip())
        finally:
            self.clean(None, cleanup)


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--account", required=True)
    p.add_argument("--allow-test-writes", action="store_true", required=True)
    args = p.parse_args()
    email(args.account)
    with command_lock():
        suite = CalendarCompositionCanary(args.account)
        try:
            suite.composition()
            suite.journal["state"] = "passed"
        except Exception as exc:
            suite.journal.update(state="failed", code=getattr(exc, "code", "unexpected_error"))
            raise
        finally:
            suite.journal["cleanup_complete"] = all(
                c["state"] == "cleaned" for c in suite.journal["cleanup"]
            )
            suite.journal["limitations"] = (
                "Self-only appointments. No invitations, attendee rendering or RSVP formatting tested. "
                "HTML storage/readback semantics, not pixel-identical clients. Stale keys, unknown "
                "write outcomes, input limits and notification policy covered offline."
            )
            suite.save()
        if not suite.journal["cleanup_complete"]:
            raise OutlookError("cleanup_pending", "Calendar composition cleanup needs review.")
        print(
            json.dumps(
                {"passed": True, "journal": str(suite.path), "checks": len(suite.journal["checks"])}
            )
        )


if __name__ == "__main__":
    main()
