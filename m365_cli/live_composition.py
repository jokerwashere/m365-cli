"""Opt-in rich-mail canary. Four test messages to one explicitly authorized recipient."""

import argparse
import base64
import json
import os
import time

from .errors import OutlookError
from .live_labels import LabelCanary
from .session import command_lock, state_dir
from .validation import email


class CompositionCanary(LabelCanary):
    suite, label = "composition", "composition"

    def composition(self, recipient):
        cleanup, tracked = [], set()
        self.journal["pending_sent"] = []
        files = state_dir() / ("composition-fixtures-" + self.marker.rsplit(" ", 1)[1])
        files.mkdir(mode=0o700)
        style = files / "style.json"
        style.write_text(
            json.dumps(
                {"font_family": "Georgia, serif", "font_size_pt": 11, "text_color": "#17365d"}
            )
        )
        image = files / "dot.png"
        image.write_bytes(
            base64.b64decode(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jU1sAAAAASUVORK5CYII="
            )
        )

        def call(*words):
            return self.invoke(None, *words)

        def track(item):
            if item["id"] not in tracked:
                cleanup.append(self.track("mail", "delete", item["id"]))
                tracked.add(item["id"])
            return cleanup[-1]

        def read(item, name):
            data = call("mail", "get", item["id"], "--body-format", "html")
            self.require(name + ".html_body", data["Body"]["BodyType"] == "HTML")
            self.require(name + ".stable_reference", bool(data["stable_id"] and data["web_url"]))
            return data

        def find_sent(needle):
            for attempt in range(8):
                items = call("mail", "list", "--folder", "sentitems", "--limit", "100")["items"]
                for item in items:
                    if self.marker not in item.get("Subject", ""):
                        continue
                    data = call("mail", "get", item["id"], "--body-format", "html")
                    if needle not in data["Body"]["Value"]:
                        continue
                    track(data)
                    addresses = {
                        r.get("Mailbox", r).get("EmailAddress", "").casefold()
                        for r in data.get("ToRecipients", [])
                    }
                    self.require(needle + ".recipient", addresses == {recipient.casefold()})
                    self.require(needle + ".sent_copy", data.get("IsDraft") is False)
                    sender = data.get("From", {})
                    self.require(
                        needle + ".sender",
                        sender.get("Mailbox", sender).get("EmailAddress", "").casefold()
                        == self.account.casefold(),
                    )
                    self.journal["pending_sent"].remove(needle)
                    self.save()
                    return data
                if attempt < 7:
                    time.sleep(2)  # Read-only propagation polling. Never retry a send.
            self.require(needle + ".sent_copy_found", False)

        def send(needle, words, draft_cleanup=None):
            self.journal["pending_sent"].append(needle)
            self.save()
            result = call(*words)
            self.require(needle + ".submitted", result.get("submitted") is True)
            if draft_cleanup:
                draft_cleanup.update(
                    state="cleaned", note="Draft sent; replacement copy tracked by pending_sent."
                )
                self.save()
            return find_sent(needle)

        try:
            defaults = call("mail", "styles")
            self.require("styles.defaults", defaults["defaults"]["font_size_pt"] == 12)
            preview = call(
                "mail",
                "render",
                "--body-format",
                "markdown",
                "--body",
                "**Preview**",
                "--output",
                str(files / "preview.html"),
            )
            self.require(
                "preview.local",
                preview["local_only"]
                and not preview["submitted"]
                and "<strong>Preview</strong>" in preview["body"],
            )
            source = send(
                "DIRECT-HTML",
                (
                    "mail",
                    "send",
                    "--to",
                    recipient,
                    "--subject",
                    self.marker + " source",
                    "--body-format",
                    "html",
                    "--body",
                    "<p>DIRECT-HTML: <strong>Rich email test</strong>. Zażółć 📨</p><p>Disposable validation message; no action needed.</p>",
                ),
            )
            self.require(
                "direct.html_preserved",
                "<strong>Rich email test</strong>" in source["Body"]["Value"],
            )
            plain = call("mail", "get", source["id"])
            self.require(
                "read.text_default",
                plain["Body"]["BodyType"] == "Text" and "Rich email test" in plain["Body"]["Value"],
            )
            draft = call(
                "mail",
                "draft",
                "create",
                "--to",
                recipient,
                "--subject",
                self.marker + " edited draft",
                "--body-format",
                "html",
                "--body",
                '<p><em>Original HTML draft</em></p><img src="cid:composition-dot" alt="Test dot">',
            )
            draft_cleanup = track(draft)
            initial = call("mail", "draft", "get", draft["id"], "--body-format", "html")
            self.require(
                "draft.html_create", "<em>Original HTML draft</em>" in initial["Body"]["Value"]
            )
            call(
                "mail",
                "attachment",
                "add",
                draft["id"],
                "--file",
                str(image),
                "--inline",
                "--content-id",
                "composition-dot",
                "--content-type",
                "image/png",
            )
            rich = "# Markdown mail test\n\nDRAFT-MARKDOWN: **Updated** with *emphasis* and Zażółć 📨.\nSecond line.\n\n- First item\n- Second item\n\n| Item | Status |\n| --- | --- |\n| Renderer | Ready |\n\n[Documentation](https://example.com)\n\n![Test dot](cid:composition-dot)\n\nDisposable validation; no action needed."
            current = call("mail", "draft", "get", draft["id"])
            updated = call(
                "mail",
                "draft",
                "update",
                draft["id"],
                "--body-format",
                "markdown",
                "--body",
                rich,
                "--style-file",
                str(style),
                "--if-change-key",
                current["change_key"],
            )
            value = updated["Body"]["Value"]
            self.require(
                "draft.markdown_semantics",
                all(
                    v in value
                    for v in (
                        "<strong>Updated</strong>",
                        "<table",
                        "<ul",
                        "cid:composition-dot",
                        "<br",
                    )
                ),
            )
            self.require("draft.custom_style", "Georgia" in value and "11pt" in value)
            attachments = call("mail", "attachment", "list", draft["id"])["items"]
            self.require(
                "draft.inline_attachment",
                any(
                    a.get("is_inline") and a.get("ContentId") == "composition-dot"
                    for a in attachments
                ),
            )
            sent_draft = send(
                "DRAFT-MARKDOWN",
                ("mail", "draft", "send", draft["id"], "--if-change-key", updated["change_key"]),
                draft_cleanup,
            )
            self.require(
                "draft.sent_format",
                "<table" in sent_draft["Body"]["Value"]
                and "Georgia" in sent_draft["Body"]["Value"],
            )
            for command, extra, format in (
                ("reply", (), "markdown"),
                ("reply", ("--all",), "html"),
                ("forward", (), "markdown"),
            ):
                label = command + ("-all" if extra else "")
                content = (
                    "**Response draft**"
                    if format == "markdown"
                    else "<p><strong>Response draft</strong></p>"
                )
                item = call(
                    "mail",
                    command,
                    source["id"],
                    *extra,
                    "--to",
                    recipient,
                    "--body-format",
                    format,
                    "--body",
                    content,
                )
                track(item)
                value = read(item, label)["Body"]["Value"]
                self.require(
                    label + ".new_content_and_quote",
                    "Response draft" in value and "DIRECT-HTML" in value,
                )
            sent_reply = send(
                "REPLY-MARKDOWN",
                (
                    "mail",
                    "reply",
                    source["id"],
                    "--to",
                    recipient,
                    "--send",
                    "--body-format",
                    "markdown",
                    "--body",
                    "**REPLY-MARKDOWN**: disposable formatting validation.",
                ),
            )
            self.require(
                "reply.sent_content_and_quote",
                "<strong>REPLY-MARKDOWN</strong>" in sent_reply["Body"]["Value"]
                and "DIRECT-HTML" in sent_reply["Body"]["Value"],
            )
            sent_forward = send(
                "FORWARD-HTML",
                (
                    "mail",
                    "forward",
                    source["id"],
                    "--to",
                    recipient,
                    "--send",
                    "--body-format",
                    "html",
                    "--body",
                    "<p><em>FORWARD-HTML</em>: disposable formatting validation.</p>",
                ),
            )
            self.require(
                "forward.sent_content_and_quote",
                "<em>FORWARD-HTML</em>" in sent_forward["Body"]["Value"]
                and "DIRECT-HTML" in sent_forward["Body"]["Value"],
            )
        finally:
            self.clean(None, cleanup)


def main():
    os.umask(0o077)
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--account", required=True)
    parser.add_argument("--recipient", required=True)
    parser.add_argument("--allow-test-writes", action="store_true", required=True)
    args = parser.parse_args()
    email(args.account)
    email(args.recipient)
    with command_lock():
        suite = CompositionCanary(args.account)
        try:
            suite.composition(args.recipient)
            suite.journal["state"] = "passed"
        except Exception as exc:
            suite.journal.update(state="failed", code=getattr(exc, "code", "unexpected_error"))
            raise
        finally:
            suite.journal["cleanup_complete"] = all(
                c["state"] == "cleaned" for c in suite.journal["cleanup"]
            ) and not suite.journal.get("pending_sent")
            suite.journal["limitations"] = (
                "Four labelled messages to the explicit test recipient. Cleanup moves only owned draft/sent copies to Deleted Items; delivered recipient copies cannot be recalled. Tests server-stored HTML semantics and styles, not pixel-identical rendering in every client. No write retries."
            )
            suite.save()
        if not suite.journal["cleanup_complete"]:
            raise OutlookError("cleanup_pending", "Composition cleanup needs review.")
        print(
            json.dumps(
                {"passed": True, "journal": str(suite.path), "checks": len(suite.journal["checks"])}
            )
        )


if __name__ == "__main__":
    main()
