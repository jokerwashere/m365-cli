"""Maintainer-only personal-site canary; actual CLI dispatch and owned fixtures."""

import argparse
import json
import tempfile
from pathlib import Path
from urllib.parse import urlsplit

from .errors import OutlookError
from .live_labels import LabelCanary
from .session import command_lock
from .sharepoint_validation import site_url


class SharePointCanary(LabelCanary):
    suite, label = "sharepoint", "files and lists"

    def __init__(self, account, site):
        self.site = site_url(site)
        if not urlsplit(self.site).path.startswith("/personal/"):
            raise OutlookError(
                "invalid_argument",
                "Write validation is restricted to an authorized personal OneDrive site.",
            )
        super().__init__(account)
        self.journal["site"] = self.site
        self.save()

    def invoke(self, api, *words):
        return super().invoke(api, *words, "--site", self.site)

    def reject(self, code, *words):
        return super().reject(code, *words, "--site", self.site)

    def clean(self, api, entries):
        super().clean(api, entries)
        for entry in entries:
            if entry["state"] != "cleaned":
                continue
            words = entry["command"]
            read = list(words)
            read.remove("--allow-unconditional")
            read[read.index("recycle")] = "get"
            try:
                self.invoke(None, *read)
            except OutlookError as error:
                entry["verified_absent"] = error.code == "not_found"
            else:
                entry["verified_absent"] = False
            self.save()

    def server_guard(self, route, payload, token, label):
        from .sharepoint_session import SharePointSession

        entry = {"command": "contract." + label, "state": "pending"}
        self.journal["operations"].append(entry)
        self.save()
        with SharePointSession(self.site, expected=self.account) as session:
            try:
                session.request(
                    route,
                    "POST",
                    payload,
                    headers={"If-Match": token, "X-HTTP-Method": "MERGE"},
                    expected=(204,),
                )
            except OutlookError as error:
                entry.update(state="rejected", code=error.code)
                if error.code != "version_conflict":
                    self.journal["unresolved"].append(len(self.journal["operations"]) - 1)
                self.save()
                self.require(label, error.code == "version_conflict")
            else:
                entry["state"] = "accepted_unexpectedly"
                self.save()
                self.require(label, False)

    def scenario(self):
        def call(*words):
            return self.invoke(None, *words)

        cleanup = []
        try:
            call("files", "capabilities")
            status = call("files", "status")
            self.require("identity.self", status["account"].casefold() == self.account.casefold())
            self.require(
                "transport.detached",
                status["transport"]
                == {
                    "mode": "direct_http",
                    "browser_connected": False,
                    "credential_storage": "command_memory",
                },
            )
            # Site ownership, in addition to current-user access: do not test in someone else's personal site.
            from .sharepoint_session import SharePointSession

            with SharePointSession(self.site, expected=self.account) as session:
                own = session.request(
                    "/_api/SP.UserProfiles.PeopleManager/GetMyProperties?$select=PersonalUrl"
                )
                self.require("site.self_owned", site_url(own["PersonalUrl"]) == self.site)
            libs = call("files", "libraries", "--limit", "1")
            self.require("libraries", bool(libs["items"]))
            root = libs["items"][0]["root_folder_id"]
            self.reject(
                "protected_resource",
                "files",
                "folder",
                "rename",
                root,
                "--name",
                "must-not-rename-root",
            )
            self.reject(
                "protected_resource", "files", "folder", "recycle", root, "--allow-unconditional"
            )
            folder = call("files", "folder", "create", "--folder", root, "--name", self.marker)
            cleanup.append(
                self.track("files", "folder", "recycle", folder["id"], "--allow-unconditional")
            )
            self.require("folder.identity", bool(folder["stable_id"] and folder["web_url"]))
            self.reject(
                "conflict", "files", "folder", "create", "--folder", root, "--name", self.marker
            )
            call("files", "folder", "get", folder["stable_id"])
            sub = call(
                "files", "folder", "create", "--folder", folder["id"], "--name", "destination"
            )
            cleanup.append(
                self.track("files", "folder", "recycle", sub["id"], "--allow-unconditional")
            )
            meta = call("files", "folder", "metadata", sub["id"])
            sub = call(
                "files",
                "folder",
                "rename",
                sub["id"],
                "--name",
                "renamed destination",
                "--if-version",
                meta["etag"],
            )
            call("files", "folder", "list", folder["id"], "--kind", "folder", "--limit", "1")
            with tempfile.TemporaryDirectory() as temporary:
                local = Path(temporary) / "fixture.txt"
                local.write_bytes(b"M365 CLI synthetic file fixture\n")
                file = call(
                    "files",
                    "upload",
                    "--folder",
                    folder["id"],
                    "--file",
                    str(local),
                    "--name",
                    "sample.txt",
                )
                cleanup.append(self.track("files", "recycle", file["id"], "--allow-unconditional"))
                self.require("upload.identity", bool(file["stable_id"] and file["web_url"]))
                call("files", "get", file["stable_id"])
                self.reject(
                    "invalid_request",
                    "files",
                    "upload",
                    "--folder",
                    folder["id"],
                    "--file",
                    str(local),
                    "--name",
                    "sample.txt",
                )
                resolved = call("files", "resolve", file["web_url"])
                self.require("resolve.file", resolved["item"]["id"] == file["id"])
                call("reference", "resolve", file["web_url"])
                fetched = call(
                    "files",
                    "fetch",
                    file["stable_id"],
                    "--output",
                    str(Path(temporary) / "fetched.txt"),
                )
                self.require(
                    "fetch.bytes", Path(fetched["output"]).read_bytes() == local.read_bytes()
                )
                self.require(
                    "fetch.manifest",
                    json.loads(Path(fetched["manifest"]).read_text())["sha256"] == file["sha256"],
                )
                meta = call("files", "metadata", file["id"])
                self.reject(
                    "version_conflict",
                    "files",
                    "rename",
                    file["id"],
                    "--name",
                    "rejected.txt",
                    "--if-version",
                    '"stale"',
                )
                renamed = call(
                    "files",
                    "rename",
                    file["id"],
                    "--name",
                    "renamed.txt",
                    "--if-version",
                    meta["etag"],
                )
                self.require("rename.identity", renamed["stable_id"] == file["stable_id"])
                self.server_guard(
                    "/_api/web/GetFileById('" + file["id"] + "')/ListItemAllFields",
                    {"FileLeafRef": "stale-rejected.txt"},
                    meta["etag"],
                    "file.server_stale_guard",
                )
                self.reject(
                    "version_conflict",
                    "files",
                    "rename",
                    file["id"],
                    "--name",
                    "rejected.txt",
                    "--if-version",
                    meta["etag"],
                )
                copied = call(
                    "files", "copy", file["id"], "--folder", folder["id"], "--name", "copy.txt"
                )
                cleanup.append(
                    self.track("files", "recycle", copied["id"], "--allow-unconditional")
                )
                self.require(
                    "copy.bytes", copied["sha256"] == file["sha256"] and copied["id"] != file["id"]
                )
                special = call(
                    "files",
                    "copy",
                    file["id"],
                    "--folder",
                    folder["id"],
                    "--name",
                    "50% # special.txt",
                )
                cleanup.append(
                    self.track("files", "recycle", special["id"], "--allow-unconditional")
                )
                self.require("special_name.bytes", special["sha256"] == file["sha256"])
                moved = call(
                    "files", "move", file["id"], "--folder", sub["id"], "--allow-unconditional"
                )
                self.require("move.identity", moved["stable_id"] == file["stable_id"])
                page = call("files", "list", "--folder", folder["id"], "--limit", "1")
                self.require("list.file", bool(page["items"]))
                self.require("pagination.more", bool(page["next_cursor"]))
                next_page = call(
                    "files",
                    "list",
                    "--folder",
                    folder["id"],
                    "--limit",
                    "1",
                    "--cursor",
                    page["next_cursor"],
                )
                self.require(
                    "pagination.distinct", next_page["items"][0]["id"] != page["items"][0]["id"]
                )
                versions = call("files", "versions", file["id"])
                if versions["items"]:
                    label = versions["items"][0]["version_label"]
                    call(
                        "files",
                        "download",
                        file["id"],
                        "--version",
                        label,
                        "--output",
                        str(Path(temporary) / "historical.txt"),
                    )
                    call(
                        "files", "restore", file["id"], "--version", label, "--allow-unconditional"
                    )
                else:
                    self.journal.setdefault("skipped", []).append(
                        "file history requires an earlier content revision"
                    )
                listing = call("lists", "create", "--name", self.marker)
                cleanup.append(
                    self.track("lists", "recycle", listing["id"], "--allow-unconditional")
                )
                call("lists", "get", listing["stable_id"])
                call("lists", "list", "--limit", "1")
                call("lists", "schema", listing["id"])
                call("lists", "users", "--email", self.account)
                old_list_etag = listing["etag"]
                listing = call(
                    "lists",
                    "rename",
                    listing["id"],
                    "--name",
                    self.marker + " renamed",
                    "--if-version",
                    listing["etag"],
                )
                self.server_guard(
                    "/_api/web/lists(guid'" + listing["id"] + "')",
                    {"Title": self.marker + " stale-rejected"},
                    old_list_etag,
                    "list.server_stale_guard",
                )
                for name, kind in (
                    ("Note", "text"),
                    ("State", "choice"),
                    ("Assignee", "person"),
                    ("Amount", "number"),
                    ("Done", "boolean"),
                    ("Due", "datetime"),
                    ("Link", "url"),
                    ("States", "multi-choice"),
                    ("People", "multi-person"),
                ):
                    flags = (
                        ["--choice", "Open", "--choice", "Closed"]
                        if kind in ("choice", "multi-choice")
                        else []
                    )
                    column = call(
                        "lists",
                        "column",
                        "create",
                        "--list",
                        listing["id"],
                        "--name",
                        name,
                        "--type",
                        kind,
                        *flags,
                    )
                    self.require("column." + kind, column["internal_name"] == name)
                record = call(
                    "lists",
                    "item",
                    "create",
                    "--list",
                    listing["id"],
                    "--data",
                    json.dumps(
                        {
                            "Title": self.marker,
                            "Note": "Synthetic only",
                            "State": "Open",
                            "Assignee": status["user_id"],
                            "Amount": 12.5,
                            "Done": False,
                            "Due": "2031-02-03T14:00:00+01:00",
                            "Link": {
                                "url": "https://example.com/fixture",
                                "description": "Synthetic",
                            },
                            "States": ["Open", "Closed"],
                            "People": [status["user_id"]],
                        }
                    ),
                )
                cleanup.append(
                    self.track(
                        "lists",
                        "item",
                        "recycle",
                        record["id"],
                        "--list",
                        listing["id"],
                        "--allow-unconditional",
                    )
                )
                self.require("record.identity", bool(record["stable_id"] and record["web_url"]))
                values = record["fields"]
                self.require(
                    "record.typed_values",
                    values["State"] == "Open"
                    and values["Amount"] == 12.5
                    and values["Done"] is False
                    and values["Due"] == "2031-02-03T13:00:00Z",
                )
                self.require(
                    "record.person_values",
                    values["Assignee"] == status["user_id"]
                    and values["People"] == [status["user_id"]],
                )
                self.require(
                    "record.complex_values",
                    values["States"] == ["Open", "Closed"]
                    and values["Link"]
                    == {"Url": "https://example.com/fixture", "Description": "Synthetic"},
                )
                self.require(
                    "resolve.record",
                    call("files", "resolve", record["web_url"])["item"]["stable_id"]
                    == record["stable_id"],
                )
                call("lists", "item", "get", record["stable_id"])
                self.reject(
                    "version_conflict",
                    "lists",
                    "item",
                    "update",
                    record["stable_id"],
                    "--data",
                    '{"Title":"rejected"}',
                    "--if-version",
                    '"stale"',
                )
                updated = call(
                    "lists",
                    "item",
                    "update",
                    record["stable_id"],
                    "--data",
                    json.dumps({"Title": self.marker + " updated"}),
                    "--if-version",
                    record["etag"],
                )
                self.require("record.update", updated["fields"]["Title"].endswith(" updated"))
                self.server_guard(
                    "/_api/web/lists(guid'" + listing["id"] + "')/items(" + record["id"] + ")",
                    {"Title": self.marker + " stale-rejected"},
                    record["etag"],
                    "record.server_stale_guard",
                )
                self.reject(
                    "version_conflict",
                    "lists",
                    "item",
                    "update",
                    record["stable_id"],
                    "--data",
                    '{"Title":"rejected"}',
                    "--if-version",
                    record["etag"],
                )
                call("lists", "item", "list", "--list", listing["id"], "--limit", "1")
                history = call("lists", "item", "versions", record["stable_id"], "--limit", "1")
                self.require("record.history", bool(history["items"]))
                call(
                    "lists",
                    "item",
                    "version",
                    record["stable_id"],
                    "--version",
                    history["items"][0]["version_label"],
                )
                call("lists", "comments", "list", record["stable_id"])
                attached = call(
                    "lists",
                    "attachment",
                    "add",
                    record["stable_id"],
                    "--file",
                    str(local),
                    "--name",
                    "evidence.txt",
                )
                self.require(
                    "attachment.identity", bool(attached["stable_id"] and attached["web_url"])
                )
                file_ref = call("files", "get", attached["file_id"])
                self.reject(
                    "unsupported_operation",
                    "files",
                    "rename",
                    attached["file_id"],
                    "--name",
                    "must-not-rename-attachment.txt",
                )
                call("lists", "attachment", "list", record["stable_id"])
                call("lists", "attachment", "get", record["stable_id"], "--name", "evidence.txt")
                saved = call(
                    "lists",
                    "attachment",
                    "download",
                    record["stable_id"],
                    "--name",
                    "evidence.txt",
                    "--output",
                    str(Path(temporary) / "attachment.txt"),
                )
                self.require(
                    "attachment.bytes", Path(saved["output"]).read_bytes() == local.read_bytes()
                )
                call(
                    "lists",
                    "attachment",
                    "recycle",
                    record["stable_id"],
                    "--name",
                    "evidence.txt",
                    "--allow-unconditional",
                )
                from .references import get

                self.require(
                    "attachment.file_reference_deleted",
                    get(file_ref["stable_id"], self.account)["web_url"] is None,
                )
                self.require(
                    "attachment.removed",
                    not call("lists", "attachment", "list", record["stable_id"])["items"],
                )
            call("files", "search", self.marker, "--limit", "1")
            self.journal["state"] = "passed"
        except Exception:
            self.journal["state"] = "failed"
            self.save()
            raise
        finally:
            self.clean(None, cleanup)
            self.journal["cleanup_complete"] = (
                all(x["state"] == "cleaned" and x.get("verified_absent") for x in cleanup)
                and not self.journal["unresolved"]
            )
            self.save()
        self.require("cleanup.complete", self.journal["cleanup_complete"])


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--account", required=True)
    p.add_argument("--site", required=True)
    p.add_argument("--allow-test-writes", action="store_true", required=True)
    a = p.parse_args()
    try:
        with command_lock():
            SharePointCanary(a.account, a.site).scenario()
    except OutlookError as error:
        print(json.dumps({"ok": False, "error": error.as_dict()}))
        raise SystemExit(1) from None


if __name__ == "__main__":
    main()
