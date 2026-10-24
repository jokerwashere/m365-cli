"""SharePoint locators and conservative, schema-driven list field validation."""

import math
import re
import uuid
from urllib.parse import quote, unquote, urlsplit, urlunsplit

from .errors import OutlookError
from .validation import iso

MAX_UPLOAD = 10 * 1024 * 1024
MAX_DOWNLOAD = 50 * 1024 * 1024
MAX_JSON = 8 * 1024 * 1024


def invalid(message):
    raise OutlookError("invalid_argument", message)


def guid(value):
    try:
        if not isinstance(value, str):
            raise ValueError
        parsed = uuid.UUID(value.strip("{}"))
        if parsed.int == 0:
            raise ValueError
        return str(parsed)
    except (ValueError, AttributeError):
        invalid("Expected a nonzero SharePoint GUID returned by the CLI.")


def item_id(value):
    if isinstance(value, bool) or not re.fullmatch(r"[1-9][0-9]{0,9}", str(value)):
        invalid("List item IDs must be positive integers.")
    return str(value)


def text(value, name="value", maximum=1024):
    if (
        not isinstance(value, str)
        or not value.strip()
        or len(value) > maximum
        or any(ord(c) < 32 for c in value)
    ):
        invalid(f"{name} must be nonempty text without control characters (max {maximum}).")
    return value


def filename(value):
    text(value, "name", 255)
    if (
        value in (".", "..")
        or value.endswith((".", " "))
        or any(c in value for c in '/\\:*?"<>|')
        or value.startswith("~$")
    ):
        invalid("Specify a single valid SharePoint file/folder name, without a path.")
    return value


def version(value):
    if not isinstance(value, str) or not re.fullmatch(r"[0-9]+\.[0-9]+", value):
        invalid("Use a version label returned by versions, such as 2.0.")
    return value


def etag(value):
    if value is None:
        return None
    text(value, "ETag", 1024)
    if value == "*" or not re.fullmatch(r'(?:W/)?"[^"\r\n]+"', value):
        invalid("Use an exact quoted ETag returned by the CLI; wildcard guards are forbidden.")
    return value


def https_url(value):
    text(value, "URL", 8192)
    try:
        u = urlsplit(value)
        if (
            u.scheme != "https"
            or u.username is not None
            or u.password is not None
            or u.port not in (None, 443)
            or not re.fullmatch(r"[a-z0-9][a-z0-9-]*\.sharepoint\.com", u.hostname or "")
            or "\\" in value
        ):
            raise ValueError
        path = unquote(u.path, errors="strict")
        if any(c in (".", "..") for c in path.split("/")) or any(ord(c) < 32 for c in path):
            raise ValueError
        if "\\" in path:
            raise ValueError
        return u
    except (ValueError, UnicodeError):
        invalid("Use an HTTPS SharePoint Online URL without credentials or unsafe path segments.")


def site_url(value):
    u = https_url(value)
    path = unquote(u.path, errors="strict").rstrip("/")
    if u.query or u.fragment or (path and not re.fullmatch(r"/(sites|teams|personal)/[^/]+", path)):
        invalid("--site must be the site root, e.g. https://tenant.sharepoint.com/sites/project.")
    return urlunsplit(("https", u.hostname, quote(path, safe="/"), "", ""))


def infer_site(value):
    u = https_url(value)
    path = unquote(u.path, errors="strict")
    path = re.sub(r"^/:[a-z]:/r/", "/", path)
    match = re.match(r"^/(sites|teams|personal)/([^/]+)(?:/|$)", path)
    if not match:
        invalid("This URL does not identify a site root; supply --site for an opaque sharing link.")
    return site_url("https://" + u.hostname + "/" + match[1] + "/" + match[2])


def literal(value):
    """An encoded OData string literal; never interpolate an unescaped path."""
    return "'" + quote(str(value).replace("'", "''"), safe="") + "'"


def resource_path(value):
    if not isinstance(value, str) or not value.startswith("/") or value.startswith("//"):
        invalid("Expected a server-relative SharePoint path.")
    if "\\" in value or any(ord(c) < 32 for c in value):
        invalid("Unsafe SharePoint path.")
    if any(c in (".", "..") for c in value.split("/")):
        invalid("Unsafe SharePoint path.")
    return value


def field_payload(data, fields, *, create=False):
    """Only known writable types; preserve omitted values and reject ambiguous casts."""
    if not isinstance(data, dict) or not data or len(data) > 100:
        invalid("--data must be a nonempty JSON object containing at most 100 fields.")
    schema = {f["internal_name"]: f for f in fields}
    out = {}
    for key, value in data.items():
        f = schema.get(key)
        if not f:
            invalid("An unknown field was supplied; use lists schema for internal field names.")
        if f["read_only"] or f.get("hidden") or f.get("append_only"):
            invalid("A supplied field is read-only, hidden or append-only.")
        if value is None:
            if f["required"]:
                invalid("A required field cannot be cleared.")
            if f["type"] in ("UserMulti", "LookupMulti", "MultiChoice"):
                invalid("Clear multi-value fields with an empty array, not null.")
            if f["type"] not in (
                "User",
                "Lookup",
                "Text",
                "Note",
                "Choice",
                "Number",
                "Currency",
                "Boolean",
                "DateTime",
                "URL",
            ):
                raise OutlookError(
                    "unsupported_field_type", "This field type has no validated write adapter."
                )
            out[key + "Id" if f["type"] in ("User", "Lookup") else key] = None
            continue
        kind = f["type"]
        if kind in ("Text", "Note", "Choice"):
            if not isinstance(value, str) or len(value) > (f.get("max_length") or 63000):
                invalid("Text/choice fields require bounded strings.")
            if f["required"] and not value.strip():
                invalid("Required text fields cannot be empty.")
            if kind == "Choice" and value not in f.get("choices", []) and not f.get("fill_in"):
                invalid("A value is not one of the field's allowed choices.")
        elif kind in ("Number", "Currency"):
            if (
                isinstance(value, bool)
                or not isinstance(value, (int, float))
                or not math.isfinite(value)
            ):
                invalid("Number fields require finite JSON numbers.")
            for boundary, compare in (
                ("minimum", lambda a, b: a < b),
                ("maximum", lambda a, b: a > b),
            ):
                if f.get(boundary) is not None and compare(value, f[boundary]):
                    invalid("A number lies outside its field's allowed range.")
        elif kind == "Boolean":
            if not isinstance(value, bool):
                invalid("Boolean fields require JSON true or false.")
        elif kind == "DateTime":
            if not isinstance(value, str):
                invalid("Date/time fields require an ISO timestamp with a UTC offset.")
            iso(value)
        elif kind in ("User", "Lookup"):
            if not isinstance(value, int) or isinstance(value, bool) or value < 1:
                invalid("Person/lookup fields require a positive site-scoped numeric ID.")
            key += "Id"
        elif kind in ("MultiChoice", "UserMulti", "LookupMulti"):
            if not isinstance(value, list) or len(value) > 100:
                invalid("Multi-value fields require a JSON array of at most 100 entries.")
            if kind == "MultiChoice":
                if any(not isinstance(v, str) or v not in f.get("choices", []) for v in value):
                    invalid("A multi-choice value is not an allowed choice.")
            else:
                if any(not isinstance(v, int) or isinstance(v, bool) or v < 1 for v in value):
                    invalid("Multi-person/lookup fields require positive site-scoped numeric IDs.")
                key += "Id"
            if len(set(value)) != len(value):
                invalid("Multi-value fields cannot contain duplicates.")
            value = {"results": value}
        elif kind == "URL":
            if (
                not isinstance(value, dict)
                or set(value) - {"url", "description"}
                or "url" not in value
            ):
                invalid("URL fields require {url, description}.")
            u = urlsplit(text(value["url"], "field URL", 2048))
            if u.scheme not in ("https", "http") or not u.hostname or u.username is not None:
                invalid("URL fields require HTTP(S) URLs without credentials.")
            description = value.get("description", "")
            if not isinstance(description, str) or len(description) > 255:
                invalid("URL descriptions must be strings of at most 255 characters.")
            value = {
                "__metadata": {"type": "SP.FieldUrlValue"},
                "Url": value["url"],
                "Description": description,
            }
        else:
            raise OutlookError(
                "unsupported_field_type",
                "This field type has no validated write adapter.",
                field_type=kind,
            )
        out[key] = value
    if create:
        for f in fields:
            if (
                f["required"]
                and not f["read_only"]
                and not f.get("hidden")
                and f.get("default") in (None, "")
                and f["internal_name"] not in data
            ):
                invalid(
                    "A required field is missing; inspect lists schema before creating the record."
                )
    return out
