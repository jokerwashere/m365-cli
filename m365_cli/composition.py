"""Shared local mail/event/Teams rendering; no network, templates or code execution."""

import html
import json
import math
import os
import re
from html.parser import HTMLParser
from pathlib import Path
from urllib.parse import urlsplit

import nh3
from markdown_it import MarkdownIt

from .errors import OutlookError

MAX_SOURCE = 1024 * 1024
MAX_RENDERED = 4 * 1024 * 1024
DEFAULT_STYLE = {
    "font_family": "Helvetica Neue, Helvetica, Arial, sans-serif",
    "font_size_pt": 12,
    "text_color": "#000000",
    "link_color": "#0563c1",
    "line_height": 1.35,
    "paragraph_spacing_pt": 8,
}
TAGS = set(
    "a b strong i em u s del p div span br hr h1 h2 h3 h4 h5 h6 ul ol li blockquote pre code table thead tbody tfoot tr th td caption img sup sub".split()
)
VOID = {"br", "hr", "img"}
CSS = set(
    "font-family font-size font-weight font-style color background-color line-height text-align text-decoration margin margin-top margin-bottom margin-left margin-right padding padding-top padding-bottom padding-left padding-right border border-top border-bottom border-left border-right border-color border-style border-width border-collapse border-spacing vertical-align width height max-width white-space word-wrap overflow-wrap".split()
)


def body_type(value):
    if value not in ("Text", "HTML"):
        raise OutlookError("invalid_body_format", "Body type must be Text or HTML.")
    return value


def options(p):
    p.set_defaults(body_composition=True)
    p.add_argument(
        "--body-format",
        choices=("text", "html", "markdown"),
        default="text",
        help="Input format; Markdown is sent as styled HTML (default text)",
    )
    p.add_argument(
        "--style",
        choices=("outlook", "minimal"),
        help="HTML/Markdown styling: outlook (default; Helvetica Neue/Arial) or minimal (no added styles)",
    )
    p.add_argument(
        "--style-file",
        help="JSON overrides for the outlook style; see mail styles, calendar styles or teams message styles",
    )


def styles():
    return {
        "styles": ["outlook", "minimal"],
        "default": "outlook",
        "defaults": DEFAULT_STYLE.copy(),
        "limits": {
            "font_size_pt": [8, 24],
            "line_height": [1, 2],
            "paragraph_spacing_pt": [0, 24],
            "colors": "#RGB or #RRGGBB",
        },
        "note": "These are CLI defaults, not your mailbox's saved font or signature. Inline HTML styles can override them.",
    }


def style_settings(path=None):
    settings = DEFAULT_STYLE.copy()
    if path:
        try:
            with Path(path).open(encoding="utf-8") as stream:
                raw = stream.read(16385)
            if len(raw) > 16384:
                raise ValueError
            extra = json.loads(raw)
        except (OSError, UnicodeError, ValueError):
            raise OutlookError(
                "invalid_style", "Cannot read style file as a JSON object (maximum 16 KiB)."
            ) from None
        if not isinstance(extra, dict) or set(extra) - set(settings):
            raise OutlookError(
                "invalid_style",
                "Unknown style keys; use mail styles or calendar styles for supported fields.",
            )
        settings.update(extra)
    font = settings["font_family"]
    if (
        not isinstance(font, str)
        or len(font) > 200
        or not all(
            re.fullmatch(r"[A-Za-z][A-Za-z0-9 -]{0,49}", name.strip()) for name in font.split(",")
        )
    ):
        raise OutlookError(
            "invalid_style", "font_family must be a comma-separated list of plain font names."
        )
    for key, low, high in (
        ("font_size_pt", 8, 24),
        ("line_height", 1, 2),
        ("paragraph_spacing_pt", 0, 24),
    ):
        value = settings[key]
        if type(value) not in (int, float) or not math.isfinite(value) or not low <= value <= high:
            raise OutlookError(
                "invalid_style",
                "Style dimensions are outside the supported ranges; use mail styles or calendar styles.",
            )
    for key in ("text_color", "link_color"):
        if not isinstance(settings[key], str) or not re.fullmatch(
            r"#(?:[0-9a-fA-F]{3}|[0-9a-fA-F]{6})", settings[key]
        ):
            raise OutlookError("invalid_style", "Style colors must use #RGB or #RRGGBB.")
    return settings


def attribute(tag, name, value):
    if name == "style" and re.search(r"[\\@]|/\*|(?:url|expression|var)\s*\(", value, re.I):
        return None
    if name in ("href", "src"):
        try:
            url = urlsplit(value)
        except ValueError:
            return None
        if any(ord(c) < 32 or c.isspace() for c in value):
            return None
        if name == "src" and url.scheme == "cid":
            return value if re.fullmatch(r"cid:[A-Za-z0-9_.@+\-]{1,200}", value) else None
        if url.scheme in ("http", "https") and url.netloc:
            return value
        if name == "href" and url.scheme in ("mailto", "tel") and url.path:
            return value
        return None
    return value


def sanitize(value):
    return nh3.clean(
        value,
        tags=TAGS,
        clean_content_tags={
            "script",
            "style",
            "iframe",
            "object",
            "embed",
            "svg",
            "math",
            "template",
            "head",
            "title",
        },
        attributes={
            "*": {"style", "title", "dir", "lang"},
            "a": {"href"},
            "img": {"src", "alt", "width", "height"},
            "ol": {"start"},
            "li": {"value"},
            "table": {"width", "cellpadding", "cellspacing", "border"},
            "th": {"colspan", "rowspan", "scope", "align", "valign"},
            "td": {"colspan", "rowspan", "align", "valign"},
        },
        attribute_filter=attribute,
        filter_style_properties=CSS,
        url_schemes={"https", "http", "mailto", "tel", "cid"},
        url_relative="deny",
    )


def tag_styles(s):
    font = f"font-family:{s['font_family']};font-size:{s['font_size_pt']:g}pt;color:{s['text_color']};line-height:{s['line_height']:g}"
    gap = f"{s['paragraph_spacing_pt']:g}pt"
    mapping = {
        "p": f"{font};margin:0 0 {gap} 0",
        "div": font,
        "a": f"color:{s['link_color']};text-decoration:underline",
        "ul": f"{font};margin:0 0 {gap} 0;padding-left:24pt",
        "ol": f"{font};margin:0 0 {gap} 0;padding-left:24pt",
        "li": "margin:0 0 3pt 0",
        "blockquote": f"{font};border-left:3px solid #d0d0d0;margin:0 0 {gap} 0;padding-left:12pt",
        "pre": f"font-family:Menlo, Consolas, Courier New, monospace;font-size:11pt;line-height:1.35;background-color:#f5f5f5;padding:8pt;white-space:pre-wrap;margin:0 0 {gap} 0",
        "code": "font-family:Menlo, Consolas, Courier New, monospace;font-size:11pt;background-color:#f5f5f5",
        "table": f"{font};border-collapse:collapse;margin:0 0 {gap} 0",
        "th": f"{font};border:1px solid #d0d0d0;padding:6pt 8pt;text-align:left;background-color:#f3f3f3;font-weight:bold",
        "td": f"{font};border:1px solid #d0d0d0;padding:6pt 8pt;vertical-align:top",
        "hr": f"border:0;border-top:1px solid #d0d0d0;margin:{gap} 0",
        "img": "max-width:100%;height:auto",
    }
    for i, scale in enumerate((1.5, 1.25, 1.1, 1, 1, 1), 1):
        mapping[f"h{i}"] = (
            f"{font};font-size:{s['font_size_pt'] * scale:g}pt;font-weight:bold;margin:12pt 0 6pt 0"
        )
    return mapping, font


class InlineStyles(HTMLParser):
    """Add presentation styles only after the library sanitizer has parsed the HTML."""

    def __init__(self, mapping):
        super().__init__(convert_charrefs=True)
        self.mapping, self.parts, self.elements = mapping, [], 0

    def handle_starttag(self, tag, attrs):
        self.elements += 1
        if self.elements > 20000:
            raise OutlookError("body_too_complex", "HTML is limited to 20,000 elements.")
        attrs = dict(attrs)
        if tag in self.mapping:
            attrs["style"] = self.mapping[tag] + (
                ";" + attrs["style"] if attrs.get("style") else ""
            )
        attributes = "".join(f' {k}="{html.escape(v or "", quote=True)}"' for k, v in attrs.items())
        self.parts.append(f"<{tag}{attributes}>")

    def handle_endtag(self, tag):
        if tag not in VOID:
            self.parts.append(f"</{tag}>")

    def handle_startendtag(self, tag, attrs):
        self.handle_starttag(tag, attrs)
        self.handle_endtag(tag)

    def handle_data(self, data):
        self.parts.append(html.escape(data, quote=False))


def render(source, input_format="text", style=None, style_file=None):
    if input_format not in ("text", "html", "markdown"):
        raise OutlookError("invalid_body_format", "Use text, html or markdown.")
    if source is None:
        if input_format != "text" or style or style_file:
            raise OutlookError(
                "body_required", "Supply --body or --body-file when selecting formatting."
            )
        return None
    if not isinstance(source, str) or len(source.encode("utf-8")) > MAX_SOURCE or "\x00" in source:
        raise OutlookError("invalid_body", "Body must be UTF-8 text without NUL, at most 1 MiB.")
    if input_format == "text":
        if style or style_file:
            raise OutlookError("invalid_style", "Styling requires --body-format html or markdown.")
        return {
            "body_type": "Text",
            "body": source,
            "composition": {"input_format": "text", "body_type": "Text"},
        }
    style = style or "outlook"
    if style not in ("outlook", "minimal") or style == "minimal" and style_file:
        raise OutlookError("invalid_style", "Style overrides require the outlook style.")
    settings = style_settings(style_file) if style == "outlook" else None
    markup = (
        MarkdownIt("commonmark", {"html": False, "maxNesting": 40, "breaks": True})
        .enable(["table", "strikethrough"])
        .render(source)
        if input_format == "markdown"
        else source
    )
    cleaned = sanitize(markup)
    mapping, base = tag_styles(settings) if settings else ({}, "")
    formatter = InlineStyles(mapping)
    formatter.feed(cleaned)
    formatter.close()
    body = "".join(formatter.parts)
    if settings:
        body = f'<div style="{html.escape(base, quote=True)}">{body}</div>'
    if len(body.encode("utf-8")) > MAX_RENDERED:
        raise OutlookError("body_too_large", "Rendered HTML exceeds 4 MiB.")
    return {
        "body_type": "HTML",
        "body": body,
        "composition": {
            "input_format": input_format,
            "body_type": "HTML",
            "style": style,
            "style_options": settings,
            "html_normalized": cleaned != markup,
            "warnings": [
                "HTML normalized to the supported email subset; review mail render output."
            ]
            if input_format == "html" and cleaned != markup
            else [],
        },
    }


def prepare(args):
    result = render(args.text, args.body_format, args.style, args.style_file)
    args.body_type = result["body_type"] if result else "Text"
    args.composition = result["composition"] if result else None
    if result:
        args.text = result["body"]
    if args.command == "render" and result is None:
        raise OutlookError("body_required", "Rendering requires --body or --body-file.")


def preview(args):
    body = args.text
    preview_font = html.escape(tag_styles(DEFAULT_STYLE)[1], quote=True)
    fragment = (
        body
        if args.body_type == "HTML"
        else f'<pre style="white-space:pre-wrap;{preview_font}">' + html.escape(body) + "</pre>"
    )
    document = (
        '<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta http-equiv="Content-Security-Policy" content="default-src \'none\'; style-src \'unsafe-inline\'; img-src \'none\'; base-uri \'none\'; form-action \'none\'"><title>Body preview</title></head><body style="margin:20px;background:#fff">'
        + fragment
        + "</body></html>"
    )
    result = {
        "body_type": args.body_type,
        "body": body,
        "composition": args.composition,
        "preview_html": document,
        "local_only": True,
        "submitted": False,
    }
    if args.output:
        try:
            with os.fdopen(
                os.open(args.output, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600), "wb"
            ) as stream:
                stream.write(document.encode("utf-8"))
        except FileExistsError:
            raise OutlookError("output_exists", "Output exists; choose a new filename.") from None
        except OSError:
            raise OutlookError(
                "output_file_error", "Could not save the preview at the requested path."
            ) from None
        result["saved_to"] = str(Path(args.output).absolute())
    return result
