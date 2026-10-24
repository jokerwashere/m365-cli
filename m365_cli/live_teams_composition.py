"""Opt-in rich Teams lifecycle: self-chat, self-only group, optional authorized channel."""

import argparse
import json
from pathlib import Path
from tempfile import TemporaryDirectory

from .cli import parser, run
from .errors import OutlookError
from .live_labels import LabelCanary
from .session import command_lock, state_dir
from .teams_composition import matches, text_body
from .validation import email


class TeamsCompositionCanary(LabelCanary):
    suite, label = "teams-composition", "Teams composition"

    def scenario(self, channel=None):
        def call(*words):
            return self.invoke(None, "teams", *words)

        def message_call(action, target, *words):
            try:
                return call("message", action, *target, *words)
            except OutlookError as error:
                mid = error.details.get("message_id")
                if action in ("send", "reply") and error.code == "readback_failed" and mid:
                    item = call("message", "get", mid, *target)
                    if self.marker in item.get("content", ""):
                        cleanup.append(self.track("teams", "message", "delete", mid, *target))
                raise

        cleanup = []
        try:
            self_chat = call("chat", "self")["id"]
            targets = [("self", ["--chat", self_chat])]
            group = call("chat", "create", "--group", "--topic", self.marker)["chat"]
            cleanup.append(self.track("teams", "chat", "leave", group["id"]))
            targets.append(("group", ["--chat", group["id"]]))
            if channel:
                call("channel", "get", channel)
                targets.append(("channel", ["--channel", channel]))
            self.require(
                "styles.shared_defaults",
                call("message", "styles")["defaults"]["font_family"].startswith("Helvetica Neue"),
            )
            with TemporaryDirectory(prefix="teams-composition-", dir=state_dir()) as tmp:
                style = Path(tmp) / "style.json"
                style.write_text(
                    '{"font_family":"Georgia, serif","font_size_pt":11,"link_color":"#006699"}'
                )
                markdown = Path(tmp) / "message.md"
                markdown.write_text(
                    self.marker
                    + '\n\n# Test heading\n\n**Bold** *italic* ~~strike~~ Zażółć 📅\nNext line\n\n- First\n- Second\n\n1. One\n2. Two\n\n> Quote\n\n[Link](https://example.com) and `code`\n\n| A | B |\n| --- | --- |\n| 1 | 2 |\n\n```\nprint("test")\n```',
                    encoding="utf-8",
                )
                rich = ["--body-format", "markdown", "--body-file", str(markdown)]
                preview = call(
                    "message", "render", *rich, "--output", str(Path(tmp) / "preview.html")
                )
                self.require("preview.local", preview["local_only"] and not preview["submitted"])
                for name, target in targets:
                    sent = message_call("send", target, *rich)
                    item = sent["message"]
                    cleanup.append(self.track("teams", "message", "delete", item["id"], *target))
                    self.require(
                        name + ".markdown",
                        matches(preview["message_content"], item["content"], "HTML"),
                    )
                    self.require(name + ".reference", bool(item["stable_id"] and item["web_url"]))
                    request_id = self.marker + ":" + str(len(self.journal["operations"]))
                    replay = run(
                        parser().parse_args(
                            [
                                "--account",
                                self.account,
                                "teams",
                                "message",
                                "send",
                                *target,
                                *rich,
                                "--request-id",
                                request_id,
                            ]
                        )
                    )
                    self.require(
                        name + ".cached_replay",
                        replay.get("replayed") is True and replay["message"]["id"] == item["id"],
                    )
                    html = (
                        self.marker
                        + '<p onclick="x()"><strong>HTML edit</strong> <a href="https://example.com">Link</a> Zażółć 📅</p><script>removed_token</script>'
                    )
                    edit_args = [
                        "--body-format",
                        "html",
                        "--body",
                        html,
                        "--style-file",
                        str(style),
                    ]
                    expected = call("message", "render", *edit_args)
                    updated = message_call(
                        "update",
                        target,
                        item["id"],
                        *edit_args,
                        "--if-version",
                        str(item["version"]),
                    )
                    self.require(
                        name + ".html_edit",
                        matches(expected["message_content"], updated["message"]["content"], "HTML"),
                    )
                    self.require(
                        name + ".style_and_sanitization",
                        "Georgia" in updated["message"]["content"]
                        and "removed_token" not in updated["message"]["content"]
                        and "onclick" not in updated["message"]["content"],
                    )
                    self.require(
                        name + ".same_id", updated["message"]["stable_id"] == item["stable_id"]
                    )
                    self.reject(
                        "conflict",
                        "teams",
                        "message",
                        "update",
                        item["id"],
                        *target,
                        *rich,
                        "--if-version",
                        str(item["version"]),
                    )
                    if name == "channel":
                        reply_args = [
                            "--reply-to",
                            item["id"],
                            "--body-format",
                            "html",
                            "--style",
                            "minimal",
                            "--body",
                            "<p>" + self.marker + " <em>reply</em></p>",
                        ]
                        reply = message_call("reply", target, *reply_args)["message"]
                        cleanup.append(
                            self.track("teams", "message", "delete", reply["id"], *target)
                        )
                        self.require(
                            "reply.html",
                            "<em>reply</em>" in reply["content"]
                            and reply["reference"]["locator"].get("reply_to") == item["id"],
                        )
                        updated_reply = message_call(
                            "update",
                            target,
                            reply["id"],
                            *rich,
                            "--style",
                            "minimal",
                            "--if-version",
                            str(reply["version"]),
                        )["message"]
                        self.require(
                            "reply.markdown_edit",
                            "<strong>Bold</strong>" in updated_reply["content"]
                            and updated_reply["stable_id"] == reply["stable_id"],
                        )
                        chain = message_call("list", target, "--reply-to", item["id"])
                        self.require(
                            "reply.chain", any(m["id"] == reply["id"] for m in chain["items"])
                        )
                    if name == "self":
                        literal = self.marker + " **literal** <b>literal</b>"
                        plain = message_call(
                            "update",
                            target,
                            item["id"],
                            "--body",
                            literal,
                            "--if-version",
                            str(updated["message"]["version"]),
                        )
                        self.require(
                            "text.legacy_literal", plain["message"]["content"] == text_body(literal)
                        )
                self.journal["channel_tested"] = bool(channel)
        finally:
            # One cleanup attempt per owned fixture, in reverse creation order.
            for entry in reversed(cleanup):
                self.clean(None, [entry])
                if entry["state"] == "cleaned" and entry["command"][1:3] == ["message", "delete"]:
                    _, _, _, mid, *target = entry["command"]
                    try:
                        result = message_call("get", target, mid)
                        self.require(
                            "cleanup.tombstone", result["deleted"] and not result.get("content")
                        )
                    except Exception as error:
                        entry.update(
                            state="verification_pending",
                            code=getattr(error, "code", "unexpected_error"),
                        )
                        self.save()


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--account", required=True)
    p.add_argument("--channel", help="Explicitly authorized standard channel ID")
    p.add_argument("--allow-test-writes", action="store_true", required=True)
    args = p.parse_args()
    email(args.account)
    with command_lock():
        suite = TeamsCompositionCanary(args.account)
        try:
            suite.scenario(args.channel)
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
                "Stored body and submission checks, not visual verification on every Teams client. Self-only group is left, not deleted. Channel checks run only with an explicitly authorized --channel."
            )
            suite.save()
        if not suite.journal["cleanup_complete"]:
            raise OutlookError("cleanup_pending", "Teams composition cleanup needs review.")
        print(
            json.dumps(
                {"passed": True, "journal": str(suite.path), "checks": len(suite.journal["checks"])}
            )
        )


if __name__ == "__main__":
    main()
