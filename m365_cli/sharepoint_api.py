"""Files and typed list records over verified SharePoint browser endpoints."""

import base64
import hashlib
import json
import re
import xml.etree.ElementTree as ET
from functools import wraps
from urllib.parse import parse_qs, parse_qsl, quote, unquote, urlencode, urlsplit

from .errors import OutlookError
from .references import Registry
from .sharepoint_validation import (
    MAX_UPLOAD,
    etag,
    field_payload,
    filename,
    guid,
    https_url,
    item_id,
    literal,
    resource_path,
    text,
    version,
)


def changed():
    raise OutlookError("api_changed", "SharePoint returned a missing or changed required field.")


def obj(data, *keys):
    if not isinstance(data, dict) or any(k not in data or data[k] is None for k in keys):
        changed()
    return data


def string(value):
    if not isinstance(value, str) or not value:
        changed()
    return value


def remote_guid(value):
    try:
        return guid(value)
    except OutlookError:
        changed()


def mutation(fn):
    @wraps(fn)
    def wrapped(self, *args, **kwargs):
        self.s.write_attempted = False
        self.s.write_acknowledged = False
        try:
            return fn(self, *args, **kwargs)
        except OutlookError as error:
            if self.s.write_acknowledged:
                raise OutlookError(
                    "write_outcome_unknown",
                    "A SharePoint write was acknowledged, but the complete result could not be verified. Inspect the operation before another write.",
                    write_acknowledged=True,
                    cause=error.as_dict(),
                ) from None
            if self.s.write_attempted and error.code in (
                "api_changed",
                "invalid_argument",
                "reference_state_invalid",
                "transport_error",
                "not_found",
            ):
                raise OutlookError(
                    "write_outcome_unknown",
                    "A SharePoint write may have completed but its result could not be verified. Inspect the operation before retrying.",
                ) from None
            raise
        except (KeyError, TypeError, ValueError, AttributeError):
            raise OutlookError(
                "write_outcome_unknown" if self.s.write_attempted else "api_changed",
                "SharePoint mutation metadata could not be verified; inspect before retrying.",
            ) from None

    return wrapped


class SharePointAPI:
    def __init__(self, session):
        self.s = session
        self.registry = Registry(session.account, session.tenant)
        self.web_scope = session.site_id + "/" + session.web_id
        self._lists = {}
        self._schemas = {}

    def link(self, path):
        resource_path(path)
        return self.s.origin + quote(path, safe="/")

    def ref(self, item, kind, *, parent=None, status="observed", url_kind="item"):
        parent = parent or self.registry.parent_key("sharepoint", "web", self.web_scope)
        locator = {
            "site_url": self.s.site,
            "site_id": self.s.site_id,
            "web_id": self.s.web_id,
            "kind": kind,
            **{
                k: item[k]
                for k in ("list_id", "item_id", "file_id", "folder_id", "version_label", "name")
                if k in item
            },
        }
        self.registry.add(
            item,
            "sharepoint",
            kind,
            locator=locator,
            parent=parent,
            url=item.get("web_url"),
            status=status,
            url_kind=url_kind,
            reason="resource_recycled" if status == "deleted" else None,
        )
        return item

    def finish(self, value):
        return self.registry.finish(value)

    def guard(self, current, expected):
        token = etag(current)
        if token is None:
            changed()
        if expected is not None and etag(expected) != token:
            raise OutlookError(
                "version_conflict", "The SharePoint resource changed since the reviewed version."
            )
        return token

    def unconditional(self, allowed):
        if not allowed:
            raise OutlookError(
                "unguarded_operation",
                "This endpoint has no verified atomic version guard. Explicit --allow-unconditional is required; prefer a guarded operation when concurrent changes matter.",
            )

    def _cursor(self, route, url, offset=0, limit=50):
        return base64.urlsafe_b64encode(
            json.dumps(
                {
                    "v": 1,
                    "scope": hashlib.sha256(
                        json.dumps(
                            [self.s.account.casefold(), self.s.tenant, self.s.site, route, limit]
                        ).encode()
                    ).hexdigest(),
                    "url": url,
                    "offset": offset,
                },
                separators=(",", ":"),
            ).encode()
        ).decode()

    def page(self, route, *, limit=50, cursor=None, normalize=lambda x: x):
        if not isinstance(limit, int) or isinstance(limit, bool) or not 1 <= limit <= 500:
            raise OutlookError("invalid_argument", "--limit must be 1–500.")
        url, offset = self.s.route(route), 0
        if cursor:
            try:
                if len(cursor) > 32000:
                    raise ValueError
                c = json.loads(base64.b64decode(cursor, altchars=b"-_", validate=True))
                expected_scope = json.loads(
                    base64.urlsafe_b64decode(self._cursor(route, url, limit=limit))
                )["scope"]
                if (
                    c.get("v") != 1
                    or c.get("scope") != expected_scope
                    or type(c.get("offset")) is not int
                    or not 0 <= c["offset"] <= 100000
                ):
                    raise ValueError
                url, offset = self.s.route(c["url"]), c["offset"]
                self.continuation(route, url)
            except (ValueError, TypeError, KeyError, UnicodeError, OutlookError):
                raise OutlookError(
                    "invalid_cursor",
                    "Use the cursor with its original account, site, collection and limit.",
                ) from None
        data = obj(self.s.request(url), "value")
        rows = data["value"]
        if not isinstance(rows, list) or any(not isinstance(x, dict) for x in rows):
            changed()
        next_url = data.get("@odata.nextLink", data.get("odata.nextLink"))
        if next_url is not None:
            self.continuation(route, next_url)
        taken = rows[offset : offset + limit]
        if offset > len(rows):
            raise OutlookError(
                "cursor_stale", "The collection changed while paging; restart this read."
            )
        next_cursor = (
            self._cursor(route, url, offset + limit, limit=limit)
            if offset + limit < len(rows)
            else self._cursor(route, next_url, limit=limit)
            if next_url
            else None
        )
        return {
            "items": [normalize(x) for x in taken],
            "next_cursor": next_cursor,
            "has_more": bool(next_cursor),
            "complete": not bool(next_cursor),
            "snapshot": False,
        }

    def continuation(self, route, url):
        base = urlsplit(self.s.route(route))
        actual = urlsplit(self.s.route(url))
        if unquote(base.path).casefold() != unquote(actual.path).casefold():
            raise OutlookError("api_changed", "SharePoint continuation changed its collection.")
        fixed = dict(parse_qsl(base.query))
        incoming = dict(parse_qsl(actual.query))
        for key, value in incoming.items():
            if key.lower() not in ("$skiptoken", "$skip", "$top") and fixed.get(key) != value:
                raise OutlookError(
                    "api_changed", "SharePoint continuation changed its query scope."
                )
        for key, value in fixed.items():
            if key.lower() not in ("$top", "$skip", "$skiptoken") and incoming.get(key) != value:
                raise OutlookError(
                    "api_changed", "SharePoint continuation omitted its query scope."
                )

    def scan(self, route, max_pages=20):
        rows, cursor = [], None
        for _ in range(max_pages):
            page = self.page(route, limit=500, cursor=cursor)
            rows.extend(page["items"])
            cursor = page["next_cursor"]
            if not cursor:
                return rows
        raise OutlookError(
            "scan_limit",
            "The bounded SharePoint scan is incomplete; narrow the requested resource.",
        )

    def list_path(self, lid, iid=None):
        path = "/_api/web/lists(guid'" + guid(lid) + "')"
        return path + ("/items(" + item_id(iid) + ")" if iid is not None else "")

    def file_path(self, fid, folder=False):
        return "/_api/web/Get" + ("Folder" if folder else "File") + "ById('" + guid(fid) + "')"

    def list_info(self, lid):
        lid = guid(lid)
        if lid not in self._lists:
            data = obj(
                self.s.request(
                    self.list_path(lid)
                    + "?$expand=RootFolder&$select=Id,Title,BaseType,BaseTemplate,Hidden,ItemCount,EnableAttachments,EnableVersioning,DefaultDisplayFormUrl,RootFolder/ServerRelativeUrl,RootFolder/UniqueId"
                ),
                "Id",
                "Title",
                "BaseType",
                "BaseTemplate",
                "RootFolder",
            )
            token = self.s.last_etag
            if remote_guid(data["Id"]) != lid:
                changed()
            root = obj(data["RootFolder"], "ServerRelativeUrl", "UniqueId")
            self._lists[lid] = {
                "id": lid,
                "name": string(data["Title"]),
                "kind": "library" if data["BaseType"] == 1 else "list",
                "base_template": data["BaseTemplate"],
                "hidden": data.get("Hidden"),
                "item_count": data.get("ItemCount"),
                "attachments_enabled": data.get("EnableAttachments"),
                "versioning_enabled": data.get("EnableVersioning"),
                "etag": token,
                "root_path": string(root["ServerRelativeUrl"]),
                "root_folder_id": remote_guid(root["UniqueId"]),
                "display_form": data.get("DefaultDisplayFormUrl"),
                "web_url": self.link(root["ServerRelativeUrl"]),
            }
        return self.ref(dict(self._lists[lid]), "list")

    def lists(self, *, libraries=False, limit=50, cursor=None):
        condition = "Hidden eq false" + (" and BaseType eq 1" if libraries else "")
        route = "/_api/web/lists?" + urlencode(
            {
                "$filter": condition,
                "$select": "Id,Title,BaseType,BaseTemplate,RootFolder/ServerRelativeUrl,RootFolder/UniqueId",
                "$expand": "RootFolder",
            }
        )

        def normal(d):
            obj(d, "Id", "Title", "BaseType", "RootFolder")
            root = obj(d["RootFolder"], "ServerRelativeUrl", "UniqueId")
            return self.ref(
                {
                    "id": remote_guid(d["Id"]),
                    "name": string(d["Title"]),
                    "kind": "library" if d["BaseType"] == 1 else "list",
                    "base_template": d.get("BaseTemplate"),
                    "root_folder_id": remote_guid(root["UniqueId"]),
                    "web_url": self.link(root["ServerRelativeUrl"]),
                },
                "list",
            )

        return self.page(route, limit=limit, cursor=cursor, normalize=normal)

    def library_for(self, resource, *, protect_root=False):
        """File writes never use generic-list attachment storage or library roots."""
        path = resource_path(resource["path"])
        rows = self.scan(
            "/_api/web/lists?$filter=Hidden%20eq%20false%20and%20BaseType%20eq%201&$select=Id,RootFolder/ServerRelativeUrl&$expand=RootFolder"
        )
        matches = []
        for row in rows:
            root = resource_path(
                obj(obj(row, "Id", "RootFolder")["RootFolder"], "ServerRelativeUrl")[
                    "ServerRelativeUrl"
                ]
            )
            if path.casefold() == root.casefold() or path.casefold().startswith(
                root.rstrip("/").casefold() + "/"
            ):
                matches.append((root, remote_guid(row["Id"])))
        if not matches:
            raise OutlookError(
                "unsupported_operation",
                "File writes require a visible document library. Manage list-record attachments with lists attachment commands.",
            )
        root, lid = max(matches, key=lambda entry: len(entry[0]))
        if protect_root and path.casefold() == root.casefold():
            raise OutlookError(
                "protected_resource",
                "Library roots cannot be renamed or recycled with folder commands.",
            )
        return {"id": lid, "root_path": root}

    def file(self, fid, *, folder=False):
        fid = guid(fid)
        data = obj(
            self.s.request(self.file_path(fid, folder)), "UniqueId", "Name", "ServerRelativeUrl"
        )
        if remote_guid(data["UniqueId"]) != fid:
            changed()
        result = {
            "id": fid,
            "kind": "folder" if folder else "file",
            "name": string(data["Name"]),
            "path": resource_path(data["ServerRelativeUrl"]),
            "web_url": self.link(data["ServerRelativeUrl"]),
            "etag": data.get("ETag") or self.s.last_etag,
            "modified": data.get("TimeLastModified"),
            "created": data.get("TimeCreated"),
        }
        if folder:
            result["item_count"] = data.get("ItemCount")
        else:
            try:
                result["size"] = int(data["Length"])
            except (ValueError, KeyError, TypeError):
                changed()
            result["version_label"] = (
                f"{data['MajorVersion']}.{data['MinorVersion']}"
                if "MajorVersion" in data and "MinorVersion" in data
                else None
            )
        return self.ref(result, result["kind"])

    def file_from_path(self, path, *, folder=False):
        route = (
            "/_api/web/Get"
            + ("Folder" if folder else "File")
            + "ByServerRelativePath(decodedUrl="
            + literal(resource_path(path))
            + ")?$select=UniqueId"
        )
        data = obj(self.s.request(route), "UniqueId")
        return self.file(remote_guid(data["UniqueId"]), folder=folder)

    def children(self, folder, *, kind="file", limit=50, cursor=None):
        self.file(folder, folder=True)
        route = (
            self.file_path(folder, True)
            + ("/Folders" if kind == "folder" else "/Files")
            + "?"
            + urlencode({"$select": "UniqueId,Name,ServerRelativeUrl"})
        )

        def normal(d):
            obj(d, "UniqueId", "Name", "ServerRelativeUrl")
            return self.ref(
                {
                    "id": remote_guid(d["UniqueId"]),
                    "kind": kind,
                    "name": string(d["Name"]),
                    "web_url": self.link(d["ServerRelativeUrl"]),
                    "path": d["ServerRelativeUrl"],
                    "folder_id": guid(folder),
                },
                kind,
            )

        return self.page(route, limit=limit, cursor=cursor, normalize=normal)

    def versions(self, fid, *, limit=50, cursor=None):
        file = self.file(fid)
        route = self.file_path(fid) + "/versions"

        def normal(d):
            obj(d, "ID", "VersionLabel")
            row = {
                "id": str(d["ID"]),
                "version_label": string(d["VersionLabel"]),
                "file_id": guid(fid),
                "created": d.get("Created"),
                "size": d.get("Size"),
                "web_url": file["web_url"],
            }
            return self.ref(row, "file_version", parent=file["stable_id"], url_kind="parent")

        result = self.page(route, limit=limit, cursor=cursor, normalize=normal)
        result["current_version"] = file.get("version_label")
        result["history_scope"] = (
            "Native historical versions; current content is read with download without --version."
        )
        return result

    def download(self, fid, *, label=None):
        before = self.file(fid)
        route = self.file_path(fid)
        if label is not None:
            version(label)
            matches = [r for r in self.scan(route + "/versions") if r.get("VersionLabel") == label]
            if len(matches) != 1:
                raise OutlookError(
                    "not_found",
                    "The requested historical version is unavailable; use current download for the current revision.",
                )
            route += "/versions(" + item_id(matches[0].get("ID")) + ")"
        content = self.s.request(route + "/$value", binary=True)
        if label is None:
            after = self.file(fid)
            if not before.get("etag") or before["etag"] != after.get("etag"):
                raise OutlookError(
                    "version_conflict",
                    "The file changed during download; no local output was written.",
                )
        return {
            "item": before,
            "version_label": label or before.get("version_label"),
            "sha256": hashlib.sha256(content).hexdigest(),
            "size": len(content),
        }, content

    @mutation
    def upload(self, folder, name, content):
        filename(name)
        if not isinstance(content, bytes) or len(content) > MAX_UPLOAD:
            raise OutlookError("file_too_large", "New uploads are limited to 10 MiB.")
        self.library_for(self.file(folder, folder=True))
        data = obj(
            self.s.request(
                self.file_path(folder, True)
                + "/Files/AddUsingPath(decodedUrl="
                + literal(name)
                + ",overwrite=false)",
                "POST",
                content,
                raw=True,
                expected=(200, 201),
            ),
            "UniqueId",
        )
        created = self.file(remote_guid(data["UniqueId"]))
        if created["name"] != name or created["size"] != len(content):
            changed()
        read, _ = self.download(created["id"])
        if read["sha256"] != hashlib.sha256(content).hexdigest():
            changed()
        return {
            **created,
            "uploaded": True,
            "sha256": read["sha256"],
            "concurrency": "new_name_only",
        }

    @mutation
    def folder_create(self, parent, name):
        filename(name)
        folder = self.file(parent, folder=True)
        self.library_for(folder)
        target = folder["path"].rstrip("/") + "/" + name
        # Some folder APIs ensure existence rather than reject collisions. A read
        # refuses already-present folders; a racing create remains nondestructive.
        try:
            self.file_from_path(target, folder=True)
        except OutlookError as error:
            if error.code != "not_found":
                raise
        else:
            raise OutlookError("conflict", "A folder with that name already exists.")
        data = obj(
            self.s.request(
                "/_api/web/folders", "POST", {"ServerRelativeUrl": target}, expected=(201,)
            ),
            "UniqueId",
            "ServerRelativeUrl",
        )
        if data["ServerRelativeUrl"] != target:
            changed()
        return {
            **self.file(remote_guid(data["UniqueId"]), folder=True),
            "creation_acknowledged": True,
            "concurrency": "existing-name precheck; concurrent same-name folder creation may coalesce",
        }

    @mutation
    def rename(self, fid, name, *, expected=None, folder=False):
        filename(name)
        current = self.file(fid, folder=folder)
        self.library_for(current, protect_root=True)
        fields = obj(self.s.request(self.file_path(fid, folder) + "/ListItemAllFields"), "Id")
        token = self.guard(self.s.last_etag, expected)
        self.s.request(
            self.file_path(fid, folder) + "/ListItemAllFields",
            "POST",
            {"FileLeafRef": name},
            headers={"If-Match": token, "X-HTTP-Method": "MERGE"},
            expected=(204,),
        )
        result = self.file(fid, folder=folder)
        if result["name"] != name or result["id"] != current["id"]:
            changed()
        return {
            **result,
            "renamed": True,
            "concurrency": "list_item_etag",
            "list_item_id": fields["Id"],
        }

    def metadata(self, fid, *, folder=False):
        file = self.file(fid, folder=folder)
        fields = obj(self.s.request(self.file_path(fid, folder) + "/ListItemAllFields"), "Id")
        token = self.s.last_etag
        return {
            "item": file,
            "list_item_id": str(fields["Id"]),
            "etag": token,
            "guard_scope": "list_item_metadata",
        }

    @mutation
    def copy(self, fid, folder, name, *, label=None):
        # Download a verified snapshot, then create a new name. No delete/recreate,
        # original-file mutation, inherited metadata or history promises.
        snapshot, content = self.download(fid, label=label)
        result = self.upload(folder, name, content)
        return {**result, "copied": True, "source": snapshot, "copied_scope": "file_bytes_only"}

    @mutation
    def move(self, fid, folder, name=None, *, allow=False):
        self.unconditional(allow)
        source = self.file(fid)
        destination = self.file(folder, folder=True)
        library = self.library_for(source)
        if self.library_for(destination)["id"] != library["id"]:
            raise OutlookError(
                "unsupported_operation", "Moves are restricted to the source document library."
            )
        target = destination["path"].rstrip("/") + "/" + filename(name or source["name"])
        if any(c in target or c in source["path"] for c in "%#"):
            raise OutlookError(
                "unsupported_operation",
                "Moves of paths containing percent/hash characters are not validated; use copy to a new name instead.",
            )
        if target == source["path"]:
            raise OutlookError("no_changes", "The file is already at the destination.")
        self.s.request(
            self.file_path(fid) + "/moveTo(newurl=" + literal(target) + ",flags=0)",
            "POST",
            expected=(200, 204),
        )
        result = self.file(fid)
        if result["path"] != target:
            changed()
        result = self.ref({**result, "folder_id": guid(folder)}, "file")
        return {**result, "moved": True, "concurrency": "unconditional", "overwrite": False}

    @mutation
    def recycle_file(self, fid, *, folder=False, allow=False):
        self.unconditional(allow)
        current = self.file(fid, folder=folder)
        self.library_for(current, protect_root=True)
        if folder and current.get("item_count") != 0:
            raise OutlookError(
                "resource_not_empty", "Only empty folders can be recycled by this command."
            )
        data = obj(self.s.request(self.file_path(fid, folder) + "/recycle()", "POST"), "value")
        recycle_id = remote_guid(data["value"])
        result = self.ref({**current, "web_url": None}, current["kind"], status="deleted")
        return {
            **result,
            "recycled": True,
            "recycle_id": recycle_id,
            "concurrency": "unconditional",
        }

    @mutation
    def restore_file(self, fid, label, *, allow=False):
        self.unconditional(allow)
        version(label)
        self.library_for(self.file(fid))
        versions = self.scan(self.file_path(fid) + "/versions")
        if not any(v.get("VersionLabel") == label for v in versions):
            raise OutlookError("not_found", "The historical version is unavailable.")
        self.s.request(
            self.file_path(fid) + "/versions/restoreByLabel(versionlabel=" + literal(label) + ")",
            "POST",
            expected=(200, 204),
        )
        return {**self.file(fid), "restored_from": label, "concurrency": "unconditional"}

    def schema(self, lid):
        lid = guid(lid)
        self.list_info(lid)
        if lid not in self._schemas:
            rows = self.scan(self.list_path(lid) + "/fields")
            fields = []
            for d in rows:
                obj(d, "Id", "InternalName", "Title", "TypeAsString", "Required", "ReadOnlyField")
                xml = d.get("SchemaXml", "")
                if (
                    not isinstance(xml, str)
                    or len(xml) > 100000
                    or re.search(r"<!\s*(DOCTYPE|ENTITY)", xml, re.I)
                ):
                    changed()
                try:
                    root = ET.fromstring(xml) if xml else ET.Element("Field")
                except ET.ParseError:
                    changed()
                choices = d.get("Choices") or [
                    x.text or "" for x in root.findall("./CHOICES/CHOICE")
                ]
                if isinstance(choices, dict):
                    choices = choices.get("results")
                if not isinstance(choices, list) or any(not isinstance(v, str) for v in choices):
                    changed()
                f = {
                    "id": remote_guid(d["Id"]),
                    "internal_name": string(d["InternalName"]),
                    "name": string(d["Title"]),
                    "type": string(d["TypeAsString"]),
                    "required": d["Required"],
                    "read_only": d["ReadOnlyField"],
                    "hidden": bool(d.get("Hidden")),
                    "indexed": bool(d.get("Indexed")),
                    "default": d.get("DefaultValue"),
                    "choices": choices,
                    "fill_in": d.get("FillInChoice", root.get("FillInChoice") == "TRUE"),
                    "append_only": root.get("AppendOnly", "FALSE").upper() == "TRUE",
                    "list_id": lid,
                    "lookup_list": root.get("List"),
                    "lookup_field": root.get("ShowField"),
                    "web_url": self.s.site
                    + "/_layouts/15/FldEdit.aspx?"
                    + urlencode({"List": lid, "Field": d["InternalName"]}),
                }
                if type(f["required"]) is not bool or type(f["read_only"]) is not bool:
                    changed()
                for attr, key, cast in (
                    ("MaxLength", "max_length", int),
                    ("Min", "minimum", float),
                    ("Max", "maximum", float),
                ):
                    if root.get(attr) is not None:
                        try:
                            f[key] = cast(root.get(attr))
                        except (ValueError, TypeError):
                            changed()
                fields.append(self.ref(f, "column", parent=self.list_info(lid)["stable_id"]))
            self._schemas[lid] = fields
        return {"list": self.list_info(lid), "items": self._schemas[lid], "complete": True}

    def list_guard(self, lid):
        info = self.list_info(lid)
        if info["hidden"] or info["base_template"] != 100:
            raise OutlookError(
                "unsupported_list",
                "Record writes are limited to visible generic lists (template 100); document-library writes use files commands.",
            )
        return info

    def record_normalize(self, lid, d, token=None):
        obj(d, "Id", "GUID")
        iid = item_id(d["Id"])
        info = self.list_info(lid)
        form = info.get("display_form")
        if not isinstance(form, str) or not form.startswith("/"):
            changed()
        fields = {
            f["internal_name"]: d.get(f["internal_name"] + "Id")
            if f["type"] in ("User", "UserMulti", "Lookup", "LookupMulti")
            else d.get(f["internal_name"])
            for f in self.schema(lid)["items"]
            if not f["hidden"] and (f["internal_name"] in d or f["internal_name"] + "Id" in d)
        }
        row = {
            "id": iid,
            "guid": remote_guid(d["GUID"]),
            "list_id": guid(lid),
            "fields": fields,
            "etag": token or d.get("odata.etag") or d.get("__metadata", {}).get("etag"),
            "version_label": d.get("OData__UIVersionString"),
            "web_url": self.link(form) + "?" + urlencode({"ID": iid}),
            "has_attachments": d.get("Attachments"),
        }
        # The provider's numeric ID is scoped by the stable list parent.
        return self.ref(row, "list_item", parent=info["stable_id"])

    def record(self, lid, iid):
        data = self.s.request(self.list_path(lid, iid))
        token = self.s.last_etag
        result = self.record_normalize(lid, data, token)
        if result["id"] != item_id(iid):
            changed()
        return result

    def records(self, lid, *, limit=50, cursor=None):
        self.schema(lid)
        return self.page(
            self.list_path(lid) + "/items?" + urlencode({"$top": limit, "$orderby": "Id asc"}),
            limit=limit,
            cursor=cursor,
            normalize=lambda d: self.record_normalize(lid, d),
        )

    def validate_fields(self, lid, data, create=False):
        fields = self.schema(lid)["items"]
        payload = field_payload(data, fields, create=create)
        for f in fields:
            if f["internal_name"] not in data or data[f["internal_name"]] is None:
                continue
            if f["type"] in ("User", "UserMulti"):
                values = (
                    data[f["internal_name"]]
                    if f["type"] == "UserMulti"
                    else [data[f["internal_name"]]]
                )
                for uid in values:
                    user = obj(
                        self.s.request("/_api/web/GetUserById(" + item_id(uid) + ")?$select=Id"),
                        "Id",
                    )
                    if user["Id"] != uid:
                        changed()
            elif f["type"] in ("Lookup", "LookupMulti"):
                try:
                    target = guid(f["lookup_list"])
                except OutlookError:
                    raise OutlookError(
                        "unsupported_field_type",
                        "Lookup target is not a supported list in this web.",
                    ) from None
                values = (
                    data[f["internal_name"]]
                    if f["type"] == "LookupMulti"
                    else [data[f["internal_name"]]]
                )
                for rid in values:
                    obj(self.s.request(self.list_path(target, rid) + "?$select=Id"), "Id")
        return payload

    def write_payload(self, lid, payload):
        # URL values carry a native complex type. Verbose JSON requires the
        # server-provided record type as well; never infer it from a list title.
        if any(isinstance(v, dict) and "__metadata" in v for v in payload.values()):
            metadata = obj(
                self.s.request(self.list_path(lid) + "?$select=ListItemEntityTypeFullName"),
                "ListItemEntityTypeFullName",
            )
            name = string(metadata["ListItemEntityTypeFullName"])
            if not re.fullmatch(r"SP\.Data\.[A-Za-z0-9_]{1,500}", name):
                changed()
            return {"__metadata": {"type": name}, **payload}, {
                "Content-Type": "application/json;odata=verbose"
            }
        return payload, {}

    @mutation
    def create_record(self, lid, data):
        self.list_guard(lid)
        payload, headers = self.write_payload(lid, self.validate_fields(lid, data, True))
        result = obj(
            self.s.request(
                self.list_path(lid) + "/items", "POST", payload, headers=headers, expected=(201,)
            ),
            "Id",
        )
        return {**self.record(lid, result["Id"]), "created": True}

    @mutation
    def update_record(self, lid, iid, data, *, expected=None):
        self.list_guard(lid)
        payload, headers = self.write_payload(lid, self.validate_fields(lid, data))
        current = self.record(lid, iid)
        token = self.guard(current["etag"], expected)
        self.s.request(
            self.list_path(lid, iid),
            "POST",
            payload,
            headers={**headers, "If-Match": token, "X-HTTP-Method": "MERGE"},
            expected=(204,),
        )
        return {**self.record(lid, iid), "updated": True, "concurrency": "etag"}

    @mutation
    def recycle_record(self, lid, iid, *, allow=False):
        self.unconditional(allow)
        self.list_guard(lid)
        current = self.record(lid, iid)
        data = obj(self.s.request(self.list_path(lid, iid) + "/recycle()", "POST"), "value")
        return {
            **self.ref(
                {**current, "web_url": None},
                "list_item",
                parent=self.list_info(lid)["stable_id"],
                status="deleted",
            ),
            "recycled": True,
            "recycle_id": remote_guid(data["value"]),
            "concurrency": "unconditional",
        }

    @mutation
    def create_list(self, name):
        text(name, "list name", 255)
        result = obj(
            self.s.request(
                "/_api/web/lists",
                "POST",
                {"Title": name, "BaseTemplate": 100, "EnableVersioning": True},
                expected=(201,),
            ),
            "Id",
        )
        return {**self.list_info(remote_guid(result["Id"])), "created": True}

    @mutation
    def rename_list(self, lid, name, *, expected=None):
        text(name, "list name", 255)
        current = self.list_guard(lid)
        token = self.guard(current["etag"], expected)
        self.s.request(
            self.list_path(lid),
            "POST",
            {"Title": name},
            headers={"X-HTTP-Method": "MERGE", "If-Match": token},
            expected=(204,),
        )
        self._lists.pop(guid(lid), None)
        return {**self.list_info(lid), "updated": True, "concurrency": "etag"}

    @mutation
    def recycle_list(self, lid, *, allow=False):
        self.unconditional(allow)
        current = self.list_guard(lid)
        if current["item_count"] != 0:
            raise OutlookError("resource_not_empty", "Only empty generic lists can be recycled.")
        data = obj(self.s.request(self.list_path(lid) + "/recycle()", "POST"), "value")
        return {
            **self.ref({**current, "web_url": None}, "list", status="deleted"),
            "recycled": True,
            "recycle_id": remote_guid(data["value"]),
            "concurrency": "unconditional",
        }

    def record_versions(self, lid, iid, *, limit=50, cursor=None):
        record = self.record(lid, iid)

        def normal(d):
            obj(d, "VersionId", "VersionLabel")
            row = {
                "id": str(d["VersionId"]),
                "version_label": d["VersionLabel"],
                "list_id": guid(lid),
                "item_id": item_id(iid),
                "created": d.get("Created"),
                "web_url": record["web_url"],
            }
            return self.ref(row, "list_item_version", parent=record["stable_id"], url_kind="parent")

        return self.page(
            self.list_path(lid, iid) + "/versions",
            limit=limit,
            cursor=cursor,
            normalize=normal,
        )

    def record_version(self, lid, iid, label):
        version(label)
        record = self.record(lid, iid)
        versions = self.scan(self.list_path(lid, iid) + "/versions")
        found = [v for v in versions if v.get("VersionLabel") == label]
        if len(found) != 1:
            raise OutlookError("not_found", "The requested record revision is unavailable.")
        d = found[0]
        known = {f["internal_name"] for f in self.schema(lid)["items"] if not f["hidden"]}
        fields = {k: v for k, v in d.items() if k in known}
        return self.ref(
            {
                "id": str(d["VersionId"]),
                "version_label": label,
                "list_id": guid(lid),
                "item_id": item_id(iid),
                "fields": fields,
                "web_url": record["web_url"],
                "field_representation": "historical_native_values; not a write payload",
            },
            "list_item_version",
            parent=record["stable_id"],
            url_kind="parent",
        )

    def attachments(self, lid, iid, *, limit=50, cursor=None):
        parent = self.record(lid, iid)

        def normal(d):
            obj(d, "FileName", "ServerRelativeUrl")
            file = self.file_from_path(d["ServerRelativeUrl"])
            return self.ref(
                {
                    "id": file["id"],
                    "name": d["FileName"],
                    "list_id": guid(lid),
                    "item_id": item_id(iid),
                    "web_url": file["web_url"],
                    "size": file["size"],
                    "file_id": file["id"],
                },
                "list_attachment",
                parent=parent["stable_id"],
            )

        return self.page(
            self.list_path(lid, iid) + "/AttachmentFiles",
            limit=limit,
            cursor=cursor,
            normalize=normal,
        )

    def attachment(self, lid, iid, name):
        filename(name)
        rows = self.scan(self.list_path(lid, iid) + "/AttachmentFiles")
        matches = [a for a in rows if a.get("FileName") == name]
        if len(matches) != 1:
            raise OutlookError("not_found", "Attachment is not present on the specified record.")
        parent = self.record(lid, iid)
        file = self.file_from_path(matches[0]["ServerRelativeUrl"])
        return self.ref(
            {
                "id": file["id"],
                "name": name,
                "list_id": guid(lid),
                "item_id": item_id(iid),
                "web_url": file["web_url"],
                "size": file["size"],
                "file_id": file["id"],
            },
            "list_attachment",
            parent=parent["stable_id"],
        )

    def attachment_download(self, lid, iid, name):
        file = self.attachment(lid, iid, name)
        snapshot, content = self.download(file["file_id"])
        return {**snapshot, "item": file}, content

    @mutation
    def attachment_add(self, lid, iid, name, content):
        self.list_guard(lid)
        filename(name)
        if not isinstance(content, bytes) or len(content) > MAX_UPLOAD:
            raise OutlookError("file_too_large", "Attachment uploads are limited to 10 MiB.")
        parent = self.record(lid, iid)
        if not self.list_info(lid)["attachments_enabled"]:
            raise OutlookError("unsupported_operation", "Attachments are disabled for this list.")
        rows = self.scan(self.list_path(lid, iid) + "/AttachmentFiles")
        if any(str(a.get("FileName", "")).casefold() == name.casefold() for a in rows):
            raise OutlookError(
                "conflict",
                "An attachment with that name already exists; replacement is unsupported.",
            )
        obj(
            self.s.request(
                self.list_path(lid, iid) + "/AttachmentFiles/add(FileName=" + literal(name) + ")",
                "POST",
                content,
                raw=True,
                expected=(200, 201),
            ),
            "FileName",
            "ServerRelativeUrl",
        )
        result = self.attachment(lid, iid, name)
        snapshot, _ = self.attachment_download(lid, iid, name)
        if snapshot["sha256"] != hashlib.sha256(content).hexdigest():
            changed()
        return {
            **result,
            "added": True,
            "sha256": snapshot["sha256"],
            "parent_stable_id": parent["stable_id"],
        }

    @mutation
    def attachment_delete(self, lid, iid, name, *, allow=False):
        self.unconditional(allow)
        self.list_guard(lid)
        current = self.attachment(lid, iid, name)
        # Recycle the attached SP.File rather than permanent AttachmentFiles DELETE.
        data = obj(
            self.s.request(self.file_path(current["file_id"]) + "/recycle()", "POST"), "value"
        )
        self.ref({"id": current["file_id"], "web_url": None}, "file", status="deleted")
        return {
            **self.ref(
                {**current, "web_url": None},
                "list_attachment",
                parent=self.record(lid, iid)["stable_id"],
                status="deleted",
            ),
            "recycled": True,
            "recycle_id": remote_guid(data["value"]),
            "concurrency": "unconditional",
        }

    def comments(self, lid, iid, *, limit=50, cursor=None):
        parent = self.record(lid, iid)

        def normal(d):
            obj(d, "id", "text")
            author = d.get("author") or {}
            return self.ref(
                {
                    "id": str(d["id"]),
                    "text": d["text"],
                    "created": d.get("createdDate"),
                    "author": {k: author[k] for k in ("name", "email") if k in author},
                    "list_id": guid(lid),
                    "item_id": item_id(iid),
                    "web_url": parent["web_url"],
                },
                "list_comment",
                parent=parent["stable_id"],
                url_kind="parent",
            )

        return self.page(
            self.list_path(lid, iid) + "/comments?" + urlencode({"$top": limit}),
            limit=limit,
            cursor=cursor,
            normalize=normal,
        )

    def users(self, email=None, *, limit=50, cursor=None):
        query = {"$select": "Id,Title,Email,PrincipalType"}
        if email:
            query["$filter"] = "Email eq '" + email.replace("'", "''") + "'"

        def normal(d):
            obj(d, "Id", "Title")
            return self.ref(
                {
                    "id": item_id(d["Id"]),
                    "name": d["Title"],
                    "email": d.get("Email"),
                    "principal_type": d.get("PrincipalType"),
                    "web_url": self.s.site + "/_layouts/15/userdisp.aspx?ID=" + item_id(d["Id"]),
                },
                "site_user",
            )

        return self.page(
            "/_api/web/siteusers?" + urlencode(query), limit=limit, cursor=cursor, normalize=normal
        )

    def resolve(self, url):
        u = https_url(url)
        if u.hostname != urlsplit(self.s.site).hostname:
            raise OutlookError(
                "invalid_argument", "Link and --site must use the same SharePoint host."
            )
        path = unquote(u.path, errors="strict")
        query = parse_qs(u.query)
        if any(len(v) != 1 for v in query.values()):
            raise OutlookError(
                "invalid_argument", "Ambiguous repeated link parameters are not supported."
            )
        prefix = unquote(urlsplit(self.s.site).path)
        if not (path == prefix or path.startswith(prefix + "/") or re.match(r"^/:[a-z]:/", path)):
            raise OutlookError("site_mismatch", "Link path belongs to a different site.")
        if path.casefold().endswith("/dispform.aspx") and "/lists/" in path.casefold():
            listpath = resource_path(path.rsplit("/", 1)[0])
            ids = query.get("ID") or query.get("id")
            if not ids or len(ids) != 1:
                raise OutlookError(
                    "invalid_argument", "List display links require one ID parameter."
                )
            data = obj(
                self.s.request("/_api/web/GetList(@p)?@p=" + literal(listpath) + "&$select=Id"),
                "Id",
            )
            return {
                "kind": "list_item",
                "item": self.record(remote_guid(data["Id"]), item_id(ids[0])),
            }
        if path.casefold().endswith("/listform.aspx") and query.get("ListId") and query.get("ID"):
            return {
                "kind": "list_item",
                "item": self.record(guid(query["ListId"][0]), item_id(query["ID"][0])),
            }
        if query.get("sourcedoc"):
            return {"kind": "file", "item": self.file(guid(query["sourcedoc"][0]))}
        # Only an existing-access metadata read: do not send redemption preferences.
        encoded = "u!" + base64.urlsafe_b64encode(url.encode()).decode().rstrip("=")
        d = obj(
            self.s.request(
                "/_api/v2.0/shares/"
                + encoded
                + "/driveItem?$select=id,file,folder,sharepointIds,webUrl"
            ),
            "sharepointIds",
        )
        ids = obj(d["sharepointIds"], "listItemUniqueId", "siteId", "webId")
        if (
            remote_guid(ids["siteId"]) != self.s.site_id
            or remote_guid(ids["webId"]) != self.s.web_id
        ):
            raise OutlookError(
                "site_mismatch",
                "The resolved artifact belongs to another site; repeat with its explicit site root.",
            )
        kind = "folder" if "folder" in d else "file" if "file" in d else None
        if not kind:
            raise OutlookError(
                "unsupported_resource", "This link is not a supported file, folder or list record."
            )
        return {
            "kind": kind,
            "item": self.file(remote_guid(ids["listItemUniqueId"]), folder=kind == "folder"),
        }

    def search(self, query, *, limit=25, offset=0):
        text(query, "query", 1000)
        if any(c in query for c in ('"', "\\")):
            raise OutlookError(
                "invalid_argument",
                "Literal search phrases cannot contain double quotes or backslashes.",
            )
        if not 1 <= limit <= 100 or not 0 <= offset <= 10000:
            raise OutlookError("invalid_argument", "Search uses limit 1–100 and offset 0–10000.")
        # Literal phrase search, no caller-supplied KQL. Validate result scope as well.
        phrase = query.replace("\\", "\\\\").replace('"', '\\"')
        kql = 'path:"' + self.s.site + '" IsDocument:1 "' + phrase + '"'
        args = {
            "querytext": "'" + kql.replace("'", "''") + "'",
            "rowlimit": limit,
            "startrow": offset,
            "selectproperties": "'Title,Path,UniqueId,FileType'",
        }
        data = obj(self.s.request("/_api/search/query?" + urlencode(args)), "PrimaryQueryResult")
        relevant = obj(
            obj(data["PrimaryQueryResult"], "RelevantResults")["RelevantResults"],
            "Table",
            "TotalRows",
        )
        rows = obj(relevant["Table"], "Rows")["Rows"]
        if not isinstance(rows, list):
            changed()
        result, skipped = [], 0
        for row in rows:
            cells = obj(row, "Cells")["Cells"]
            if not isinstance(cells, list):
                changed()
            values = {
                c["Key"]: c.get("Value")
                for c in cells
                if isinstance(c, dict) and isinstance(c.get("Key"), str)
            }
            path = values.get("Path", "")
            if not isinstance(path, str) or not path.startswith(self.s.site + "/"):
                skipped += 1
                continue
            u = https_url(path)
            item = {
                "id": remote_guid(values.get("UniqueId")),
                "name": values.get("Title"),
                "web_url": self.s.origin + quote(unquote(u.path), safe="/"),
                "kind": "file",
                "file_type": values.get("FileType"),
            }
            result.append(self.ref(item, "file"))
        total = relevant["TotalRows"]
        if not isinstance(total, int) or isinstance(total, bool):
            changed()
        more = offset + len(rows) < total
        return {
            "items": result[:limit],
            "offset": offset,
            "next_offset": offset + len(rows) if more and rows else None,
            "has_more": more,
            "indexed_total": total,
            "skipped_out_of_scope": skipped,
            "coverage": "site-scoped indexed literal phrase; index freshness and completeness are not guaranteed",
        }

    @mutation
    def create_column(self, lid, name, kind, *, choices=None, required=False):
        self.list_guard(lid)
        # Deliberately narrow schema writes. No deletion, retyping or existing-field edits.
        if not re.fullmatch(r"[A-Za-z][A-Za-z0-9_]{0,31}", name):
            raise OutlookError(
                "invalid_argument",
                "New column names use 1–32 ASCII letters, digits or underscores, starting with a letter.",
            )
        types = {
            "text": (2, "SP.FieldText"),
            "choice": (6, "SP.FieldChoice"),
            "person": (20, "SP.FieldUser"),
            "number": (9, "SP.FieldNumber"),
            "boolean": (8, "SP.Field"),
            "datetime": (4, "SP.FieldDateTime"),
            "url": (11, "SP.FieldUrl"),
            "multi-choice": (15, "SP.FieldMultiChoice"),
            "multi-person": (20, "SP.FieldUser"),
        }
        if kind not in types:
            raise OutlookError("unsupported_field_type", "Unsupported new column type.")
        fields = self.schema(lid)["items"]
        if any(
            name.casefold() in (f["name"].casefold(), f["internal_name"].casefold()) for f in fields
        ):
            raise OutlookError("conflict", "A column with that name already exists.")
        if required and self.list_info(lid)["item_count"] != 0:
            raise OutlookError(
                "resource_not_empty", "Required columns may only be added to an empty list."
            )
        field_type, metadata_type = types[kind]
        payload = {
            "__metadata": {"type": metadata_type},
            "Title": name,
            "FieldTypeKind": field_type,
            "Required": required,
        }
        if kind in ("choice", "multi-choice"):
            if (
                not isinstance(choices, list)
                or not 1 <= len(choices) <= 100
                or len(set(choices)) != len(choices)
            ):
                raise OutlookError(
                    "invalid_argument", "Choice columns need 1–100 distinct --choice values."
                )
            for choice in choices:
                text(choice, "choice", 255)
            payload["Choices"] = {"results": choices}
        elif choices:
            raise OutlookError("invalid_argument", "--choice is only for a choice column.")
        if kind == "multi-person":
            payload["AllowMultipleValues"] = True
        created = obj(
            self.s.request(
                self.list_path(lid) + "/fields",
                "POST",
                payload,
                headers={"Content-Type": "application/json;odata=verbose"},
                expected=(201,),
            ),
            "Id",
            "InternalName",
        )
        identifier = remote_guid(created["Id"])
        self._schemas.pop(guid(lid), None)
        found = [f for f in self.schema(lid)["items"] if f["id"] == identifier]
        if len(found) != 1 or found[0]["internal_name"] != name:
            changed()
        return {**found[0], "created": True}
