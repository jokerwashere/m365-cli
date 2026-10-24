"""Opt-in draft-only attachment canary. Never sends mail or fetches cloud links."""

import argparse
import base64
import json
from email import policy
from email.parser import BytesParser

from .api import API, typed
from .cli import parser
from .errors import OutlookError
from .live_references import Canary
from .maintainer import guard
from .operations import execute
from .references import get
from .session import Session, command_lock, state_dir


class AttachmentCanary(Canary):
    suite = "attachments"
    label = "attachment"

    def seed(self, api, parent, attachment):
        # Fixture construction uses the same acknowledged, version-guarded adapter as
        # CLI add. Rich item creation is a maintainer fixture, not a public JSON-write API.
        guard(parser().parse_args(["mail", "attachment", "add", parent, "--file", "fixture"]))
        entry = {"command": "fixture.CreateAttachment", "state": "pending", "mutation": True}
        self.journal["operations"].append(entry)
        self.save()
        try:
            value = execute(
                self.marker + ":" + str(len(self.journal["operations"])),
                self.account,
                {"fixture_type": attachment["__type"], "parent": parent},
                lambda: api._attachment_create(parent, attachment),
            )
            entry.update(state="completed", id=value["attachments"][0]["AttachmentId"]["Id"])
        except Exception as exc:
            entry.update(state="failed", code=getattr(exc, "code", "unexpected_error"))
            self.journal["unresolved"].append(len(self.journal["operations"]) - 1)
            raise
        finally:
            self.save()

    def check_error(self, name, code, call):
        try:
            call()
        except OutlookError as exc:
            self.require(name, exc.code == code)
        else:
            self.require(name, False)

    def attachments(self):
        with Session() as session:
            session.authenticate(expected=self.account)
            self.require(
                "transport.detached",
                session.transport_info()
                == {
                    "mode": "direct_http",
                    "browser_connected": False,
                    "credential_storage": "command_memory",
                },
            )
            api, cleanup = API(session), []
            exports = state_dir() / ("attachment-exports-" + self.marker.rsplit(" ", 1)[1])
            exports.mkdir(mode=0o700)
            binary = bytes(range(256)) + b"\x00m365 attachment canary\xff"
            leaf_bytes = b"Nested M365 validation file\n"
            path = exports / "input.bin"
            path.write_bytes(binary)
            empty = exports / "empty.dat"
            empty.touch(mode=0o600)
            try:
                drafts = []
                for label in ("target", "source"):
                    draft = self.invoke(
                        api,
                        "mail",
                        "draft",
                        "create",
                        "--subject",
                        self.marker + " " + label,
                        "--body",
                        "Disposable attachment test draft",
                    )
                    drafts.append(draft)
                    cleanup.append(self.track("mail", "draft", "delete", draft["id"]))
                target, source = drafts
                self.invoke(api, "mail", "attachment", "add", source["id"], "--file", str(path))
                self.invoke(
                    api,
                    "mail",
                    "attachment",
                    "add",
                    target["id"],
                    "--item-id",
                    source["id"],
                    "--name",
                    "Copied source",
                )
                self.invoke(
                    api,
                    "mail",
                    "attachment",
                    "add",
                    target["id"],
                    "--file",
                    str(path),
                    "--name",
                    "binary.dat",
                )
                self.invoke(api, "mail", "attachment", "add", target["id"], "--file", str(empty))
                self.invoke(
                    api,
                    "mail",
                    "attachment",
                    "add",
                    target["id"],
                    "--file",
                    str(path),
                    "--inline",
                    "--content-id",
                    "canary@m365.example",
                    "--content-type",
                    "application/octet-stream",
                    "--name",
                    "inline.dat",
                )
                leaf = typed(
                    "FileAttachment",
                    Name="leaf.txt",
                    ContentType="text/plain",
                    Content=base64.b64encode(leaf_bytes).decode(),
                )
                inner = typed(
                    "ItemAttachment",
                    Name="Inner",
                    Item=typed(
                        "Message",
                        Subject=self.marker + " inner",
                        Body={"BodyType": "Text", "Value": "Inner message"},
                        Attachments=[leaf],
                    ),
                )
                fixtures = {
                    "Nested message": typed(
                        "Message",
                        Subject=self.marker + " outer",
                        Body={"BodyType": "Text", "Value": "Outer message"},
                        Attachments=[inner],
                    ),
                    "Calendar item": typed(
                        "CalendarItem",
                        Subject=self.marker + " event",
                        Start="2031-02-03T13:00:00Z",
                        End="2031-02-03T13:30:00Z",
                    ),
                    "Contact item": typed(
                        "Contact", GivenName="M365", Surname="Validation", DisplayName=self.marker
                    ),
                    "Task item": typed("Task", Subject=self.marker + " task", PercentComplete=50),
                }
                for name, item in fixtures.items():
                    self.seed(api, target["id"], typed("ItemAttachment", Name=name, Item=item))
                listed = self.invoke(api, "mail", "attachment", "list", target["id"])
                self.require(
                    "top_level_types",
                    len(listed["items"]) == 8
                    and {a["kind"] for a in listed["items"]} == {"file", "item"},
                )
                by_name = {a["name"]: a for a in listed["items"]}
                inline = by_name["inline.dat"]
                self.require(
                    "inline_metadata",
                    inline["is_inline"] and inline["ContentId"] == "canary@m365.example",
                )
                tree = self.invoke(api, "mail", "attachment", "list", target["id"], "--recursive")
                self.require(
                    "recursive_tree",
                    tree["complete"]
                    and len(tree["items"]) == 11
                    and max(a["depth"] for a in tree["items"]) == 2,
                )
                self.require(
                    "nested_references",
                    all(
                        a.get("stable_id")
                        and a.get("web_url")
                        and a["reference"]["url_kind"] == "parent"
                        for a in tree["items"]
                    ),
                )
                self.require(
                    "metadata_only_list",
                    all("Item" not in a and "Content" not in a for a in tree["items"]),
                )
                for name, options in (
                    ("depth_bound", ["--max-depth", "0"]),
                    ("item_bound", ["--max-items", "2"]),
                ):
                    limited = self.invoke(
                        api, "mail", "attachment", "list", target["id"], "--recursive", *options
                    )
                    self.require(
                        name, limited["complete"] is False and bool(limited["truncations"])
                    )
                leaf_item = next(a for a in tree["items"] if a["name"] == "leaf.txt")
                cases = [
                    ("binary.dat", binary),
                    ("empty.dat", b""),
                    ("inline.dat", binary),
                    ("leaf.txt", leaf_bytes),
                ]
                for index, (name, content) in enumerate(cases):
                    item = leaf_item if name == "leaf.txt" else by_name[name]
                    output = exports / f"download-{index}.bin"
                    saved = self.invoke(
                        api,
                        "mail",
                        "attachment",
                        "get",
                        item["id"],
                        "--mail-id",
                        target["id"],
                        "--output",
                        str(output),
                    )
                    self.require(
                        "byte_exact_" + name,
                        output.read_bytes() == content
                        and saved["bytes"] == len(content)
                        and saved["stable_id"] == item["stable_id"],
                    )
                by_name["Inner"] = next(a for a in tree["items"] if a["name"] == "Inner")
                for name in ("Copied source", "Nested message", "Inner"):
                    output = exports / (name.replace(" ", "-") + ".eml")
                    saved = self.invoke(
                        api,
                        "mail",
                        "attachment",
                        "get",
                        by_name[name]["id"],
                        "--output",
                        str(output),
                    )
                    message = BytesParser(policy=policy.default).parsebytes(output.read_bytes())
                    expected_bytes = binary if name == "Copied source" else leaf_bytes
                    self.require(
                        "native_mime_" + name,
                        saved["export_format"] == "mime"
                        and any(
                            p.get_payload(decode=True) == expected_bytes for p in message.walk()
                        ),
                    )
                for name, expected_field in (
                    ("Calendar item", "Start"),
                    ("Contact item", "GivenName"),
                    ("Task item", "PercentComplete"),
                ):
                    item = by_name[name]
                    inspected = self.invoke(
                        api, "mail", "attachment", "get", item["id"], "--mail-id", target["id"]
                    )
                    self.require(
                        "structured_" + name,
                        expected_field in inspected["Item"]
                        and inspected["stable_id"] == item["stable_id"],
                    )
                    output = exports / (name.replace(" ", "-") + ".json")
                    saved = self.invoke(
                        api,
                        "mail",
                        "attachment",
                        "get",
                        item["id"],
                        "--mail-id",
                        target["id"],
                        "--output",
                        str(output),
                    )
                    value = json.loads(output.read_text())
                    self.require(
                        "json_export_" + name,
                        saved["export_format"] == "json"
                        and expected_field in value["Item"]
                        and value["web_url"] == item["web_url"],
                    )
                for name in ("Calendar item", "Contact item", "Task item"):
                    output = exports / (name.replace(" ", "-") + ".raw")
                    saved = self.invoke(
                        api,
                        "mail",
                        "attachment",
                        "get",
                        by_name[name]["id"],
                        "--format",
                        "raw",
                        "--output",
                        str(output),
                    )
                    self.require(
                        "explicit_raw_" + name,
                        saved["export_format"] in ("mime", "msg", "icalendar", "vcard")
                        and saved["bytes"] == output.stat().st_size
                        and bool(saved["export_note"]),
                    )
                self.check_error(
                    "wrong_parent_rejected",
                    "attachment_not_found",
                    lambda: api.attachment_inspect(leaf_item["id"], mail_id=source["id"]),
                )
                self.check_error(
                    "nested_delete_rejected",
                    "attachment_not_found",
                    lambda: api.attachment_delete(target["id"], leaf_item["id"]),
                )
                outer = by_name["Nested message"]
                self.invoke(
                    api,
                    "mail",
                    "attachment",
                    "delete",
                    target["id"],
                    "--attachment-id",
                    outer["id"],
                )
                self.require(
                    "nested_reference_invalidated",
                    get(leaf_item["stable_id"], self.account)["status"] == "parent_deleted",
                )
                for item in listed["items"]:
                    if item["id"] != outer["id"]:
                        self.invoke(
                            api,
                            "mail",
                            "attachment",
                            "delete",
                            target["id"],
                            "--attachment-id",
                            item["id"],
                        )
                self.require(
                    "all_top_level_attachments_removed",
                    self.invoke(api, "mail", "attachment", "list", target["id"])["items"] == [],
                )
                self.require(
                    "source_item_unchanged", len(api.get(source["id"])["Attachments"]) == 1
                )
            finally:
                self.clean(api, cleanup)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--account", required=True)
    parser.add_argument("--allow-test-writes", action="store_true", required=True)
    args = parser.parse_args()
    with command_lock():
        suite = AttachmentCanary(args.account)
        try:
            suite.attachments()
            suite.journal["state"] = "passed"
        except Exception as exc:
            suite.journal.update(state="failed", code=getattr(exc, "code", "unexpected_error"))
            raise
        finally:
            suite.journal["cleanup_complete"] = all(
                c["state"] == "cleaned" for c in suite.journal["cleanup"]
            )
            suite.journal["limitations"] = (
                "Draft-only fixture suite; cloud/link/future types use offline contracts. Cloud-link creation is not supported. Native exports are Exchange conversions, not lossless mailbox backups. Cleanup moves test drafts to Deleted Items."
            )
            suite.save()
        if not suite.journal["cleanup_complete"]:
            raise OutlookError("cleanup_pending", "Attachment canary cleanup needs review.")
        print(
            json.dumps(
                {"passed": True, "journal": str(suite.path), "checks": len(suite.journal["checks"])}
            )
        )


if __name__ == "__main__":
    main()
