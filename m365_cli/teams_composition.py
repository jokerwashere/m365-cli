"""Shared body rendering with Teams-specific message bounds and media limits."""

import html
import re
from html.parser import HTMLParser

from . import composition
from .errors import OutlookError
from .validation import nonempty

MAX_MESSAGE = 20000
BLOCKS = set(
    "p div h1 h2 h3 h4 h5 h6 ul ol li blockquote pre table thead tbody tfoot tr th td caption hr".split()
)


def css_spacing(value):
    # Only serialization whitespace outside quoted values; declaration order,
    # duplicates, values and quoted font names remain significant.
    parts = re.split(r"""("[^"]*"|'[^']*')""", value)
    for i in range(0, len(parts), 2):
        parts[i] = re.sub(r"[ \t\r\n\f]+", " ", parts[i])
        parts[i] = re.sub(r" *([:;,]) *", r"\1", parts[i])
    return "".join(parts).strip().rstrip(";")


class HtmlContract(HTMLParser):
    """Compare rich content without ignoring lost formatting or link targets."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.events, self.stack = [], []

    def handle_starttag(self, tag, attrs):
        attrs = [(k, css_spacing(v or "") if k == "style" else v) for k, v in attrs]
        self.events.append(("start", tag, tuple(sorted(attrs))))
        if tag not in composition.VOID:
            preserve = (
                tag in ("pre", "code")
                or bool(self.stack and self.stack[-1][1])
                or bool(
                    re.search(
                        r"(?:^|;)white-space:(?:pre|pre-wrap|pre-line|break-spaces)(?:;|$)",
                        dict(attrs).get("style", ""),
                    )
                )
            )
            self.stack.append((tag, preserve))

    def handle_endtag(self, tag):
        if tag not in composition.VOID:
            self.events.append(("end", tag))
            if not self.stack or self.stack.pop()[0] != tag:
                raise ValueError("Unbalanced HTML")

    def handle_data(self, value):
        preserve = bool(self.stack and self.stack[-1][1])
        if self.events and self.events[-1][0] == "text" and self.events[-1][2] == preserve:
            value = self.events.pop()[1] + value
        self.events.append(("text", value, preserve))

    def canonical(self, value):
        self.feed(value)
        self.close()
        if self.stack or len(self.events) > 60000:
            raise ValueError("Invalid HTML structure")

        def boundary(event):
            return event is None or event[0] in ("start", "end") and event[1] in BLOCKS

        result = []
        for i, event in enumerate(self.events):
            if event[0] == "text" and not event[2]:
                value = re.sub(r"[ \t\r\n\f]+", " ", event[1])
                if boundary(self.events[i - 1] if i else None):
                    value = value.lstrip(" ")
                if boundary(self.events[i + 1] if i + 1 < len(self.events) else None):
                    value = value.rstrip(" ")
                if value:
                    result.append(("text", value, False))
            else:
                result.append(event)
        return result


def matches(expected, actual, body_type):
    if not isinstance(actual, str):
        return False
    if body_type == "Text" or expected == actual:
        return expected == actual
    # Service indentation can expand the bounded submitted body.
    if len(actual.encode("utf-8")) > composition.MAX_SOURCE:
        return False
    try:
        return HtmlContract().canonical(expected) == HtmlContract().canonical(actual)
    except (ValueError, TypeError):
        return False


def text_body(value):
    nonempty(value, "body")
    if "\x00" in value:
        raise OutlookError("invalid_body", "Body must not contain NUL.")
    if len(value.encode("utf-8")) > MAX_MESSAGE:
        raise OutlookError(
            "body_too_large", "Teams plain-text bodies are limited to 20,000 UTF-8 bytes."
        )
    # Preserve legacy wire content and request-ID semantics for literal text.
    return "<p>" + html.escape(value).replace("\n", "<br>") + "</p>"


class RichBody(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.text, self.image = [], False

    def handle_starttag(self, tag, attrs):
        if tag == "img":
            self.image = True

    def handle_data(self, data):
        self.text.append(data)


def content(value, body_type="Text"):
    """Validate prepared content; the CLI sanitizes rich input with composition.render."""
    composition.body_type(body_type)
    if body_type == "Text":
        return text_body(value)
    nonempty(value, "body")
    if "\x00" in value:
        raise OutlookError("invalid_body", "Body must not contain NUL.")
    if len(value.encode("utf-8")) > MAX_MESSAGE:
        raise OutlookError(
            "body_too_large", "Teams rendered HTML is limited to 20,000 UTF-8 bytes."
        )
    parsed = RichBody()
    parsed.feed(value)
    parsed.close()
    if parsed.image:
        raise OutlookError(
            "unsupported_body", "Teams inline images/hosted content are not supported; use a link."
        )
    if not "".join(parsed.text).strip():
        raise OutlookError("invalid_body", "Teams messages require nonempty text after rendering.")
    return value


def prepare(args):
    composition.prepare(args)
    content(args.text, args.body_type)
    if args.body_type == "HTML":
        args.composition["warnings"] = (
            ["HTML normalized to the shared body subset; review teams message render output."]
            if args.composition["warnings"]
            else []
        )
        args.composition["warnings"].append(
            "Teams clients may ignore fonts, spacing, table styling or other HTML presentation; preview is not a Teams rendering guarantee."
        )


def styles():
    result = composition.styles()
    result["limits"].update(
        text_source_bytes=MAX_MESSAGE, rich_rendered_bytes=MAX_MESSAGE, images=False
    )
    result["note"] += " Teams may ignore presentation styles; minimal adds no styles."
    return result


def preview(args):
    result = composition.preview(args)
    result["message_content"] = content(args.text, args.body_type)
    return result
