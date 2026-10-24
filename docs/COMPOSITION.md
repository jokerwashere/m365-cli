# HTML and Markdown bodies

Mail, calendar events and Teams messages accept `--body-format text|html|markdown`.
The default remains `text`: existing commands keep sending literal plain text.
Rich formats work for `mail send`, `mail draft create/update`, `mail reply`
(including `--all`), `mail forward`, `calendar create/update`, and Teams
`message send/reply/update`. All use the same renderer, sanitizer and styles;
Teams adds stricter message-size/media limits described below. Sending a saved
mail draft preserves its existing body format.

## Compose, preview and send

```sh
m365 mail render --body-format markdown --body-file message.md --output preview.html
m365 mail draft create --to colleague@example.com --subject 'Project update' \
  --body-format markdown --body-file message.md
m365 mail draft get DRAFT_ID --body-format html
m365 mail draft send DRAFT_ID --if-change-key CHANGE_KEY --request-id update-001
```

`mail render` and `calendar render` are identical local previews and require no login.
Each returns JSON with the same prepared content used for mutations: `body_type`,
the actual `body` to submit, `composition` metadata and `preview_html`.
`--output` additionally writes a standalone UTF-8 HTML preview with private file
permissions, refusing to overwrite an existing path. The preview blocks scripts
and remote images. It is a layout aid, not a pixel-identical Outlook rendering.

For an authorized immediate send:

```sh
m365 mail send --to colleague@example.com --subject 'Project update' \
  --body-format markdown --body-file message.md --request-id update-002
```

Use `--body` for a short body, `--body-file PATH` for a UTF-8 file, or `--body-file -`
for stdin. File extensions never select a format automatically. `--dry-run` renders
and validates rich content locally; inspect `plan.text`, `plan.body_type` and
`plan.composition`. It does not validate recipients' permissions or contact Outlook.

Example Markdown:

```markdown
Hello team,

Here is the **weekly update**.

- The review is complete.
- Deployment is planned for Thursday.

| Item | Status |
| --- | --- |
| Review | Complete |
| Deployment | Scheduled |

[Project notes](https://example.com/project)

Thanks,
Alex
```

Headings, emphasis, ordered/unordered and nested lists, blockquotes, links, images,
fenced/inline code, tables and strikethrough are supported. Single newlines become
line breaks for email, including sign-offs. Blank lines separate paragraphs.
Embedded HTML in Markdown is escaped as literal text; select `html` for authored
HTML. Markdown is converted locally to HTML; it is not sent as Markdown source.
There are no executable templates, includes, syntax-highlighting scripts or remote
resource fetches. Task-list extensions, footnotes and automatic bare-URL linking
are not promised; use `[label](https://...)` or `<https://...>` for links.

## Calendar events

```sh
m365 calendar styles
m365 calendar render --body-format markdown --body-file agenda.md --output agenda.html
m365 calendar create --subject 'Planning' \
  --start '2031-02-03T14:00:00+01:00' --end '2031-02-03T14:30:00+01:00' \
  --body-format markdown --body-file agenda.md --request-id planning-001
m365 calendar get EVENT_ID --body-format html
m365 calendar update EVENT_ID --body-format html --body-file agenda.html \
  --style minimal --if-change-key CHANGE_KEY --request-id planning-update-001
```

These examples create a personal appointment without attendees. Supplying attendees
on creation sends invitations unless `--save-only` is supplied. A never-sent
meeting supports `calendar update --save-only` and later `calendar send EVENT_ID`;
updating a sent meeting requires `--notify`
and organizer authority. Formatting never bypasses these guards or recurrence and
change-key checks. Rich body updates return HTML readback; default `calendar get`
returns plain text. A failed readback after an acknowledged update is reported as
`readback_failed`, never retried. Creation returns the new reference and composition
metadata; explicitly get it to inspect the stored body.

The formats, style presets, `--style-file`, `--body`, file/stdin input and dry-run
behavior are identical to mail. Body updates replace the whole event description;
omitting the body preserves it. Use `--body ''` to clear it. Existing plain-text
request IDs remain recognizable, including those recorded by the earlier raw-HTML
calendar implementation. Old rich-calendar requests deliberately conflict if their
stored fingerprint differs: inspect the item before considering another write.

`calendar render` and `calendar styles` use the same implementation as their mail
counterparts. Calendar response comments (`calendar respond`) remain literal text.
This does not add event attachment writes: CID syntax can survive rendering but
will not embed an image without a corresponding existing attachment. Ordinary
HTTP(S) links/images use the same rules as mail.

## Restrained styling

```sh
m365 mail styles
m365 mail render --body-format markdown --body-file message.md --style-file style.json
m365 mail send --to colleague@example.com --subject 'Update' \
  --body-format markdown --body-file message.md --style-file style.json
```

Default `--style outlook` uses Helvetica Neue with Helvetica/Arial/sans-serif
fallbacks, 12 pt black text, 1.35 line height, 8 pt paragraph spacing, blue underlined
links, restrained headings in the same family and light table borders. Code uses
Menlo with Consolas/Courier New/monospace fallbacks at 11 pt. Styles are inline so
messages do not depend on an external stylesheet. These are CLI defaults chosen for ordinary correspondence;
they do not read or change your mailbox's saved font, signature, stationery or
organization branding. Font selection happens on the recipient's device; Arial
and Consolas provide the intended Windows fallbacks. Font availability, line
wrapping and dark-mode rendering depend on the recipient's client.

A style file is a JSON object containing any subset of these keys:

```json
{
  "font_family": "Helvetica Neue, Helvetica, Arial, sans-serif",
  "font_size_pt": 12,
  "text_color": "#000000",
  "link_color": "#0563c1",
  "line_height": 1.35,
  "paragraph_spacing_pt": 8
}
```

Unknown keys, invalid JSON, invalid colors and out-of-range values are rejected
before connecting. Font size is 8–24 pt, line height 1–2, paragraph spacing 0–24 pt;
colors use `#RGB` or `#RRGGBB`. Font families are comma-separated plain names.
Use `mail styles` to discover the installed schema/defaults. Style files contain
values, not arbitrary CSS. They are resolved before execution and their values
are bound to the operation's duplicate-prevention fingerprint.

`--style minimal` adds no wrapper or presentation styles. Use this for authored
HTML with its own inline formatting, or unadorned Markdown conversion. Sanitization
still applies. Style files require `--style outlook` (the rich-format default).
Styling flags with plain text, or format/style flags without a supplied body, fail
explicitly (except the default `--body-format text`). Updating a subject alone leaves
the draft/event body untouched. An explicit empty body clears it.

## HTML, images and replies

```sh
m365 mail draft create --subject 'HTML update' --to colleague@example.com \
  --body-format html --body-file message.html --style minimal
m365 mail reply MESSAGE_ID --body-format markdown --body-file reply.md
m365 mail forward MESSAGE_ID --to colleague@example.com \
  --body-format html --body-file introduction.html --send --request-id forward-001
```

HTML is normalized to a supported email subset: paragraphs, headings, basic text
formatting, lists, links, tables, code and images. Safe inline CSS for typography,
spacing and table borders is retained. On the same element, authored inline styles
follow and override generated defaults. Full HTML documents are reduced to body
content; document-head styles are not inlined. Classes, external stylesheets,
conditional comments, scripts, forms, embedded objects and SVG are not supported.
Unsafe URL schemes and CSS resource/executable expressions are removed. Relative
URLs are removed; use absolute HTTP(S) links or `mailto:`/`tel:` links.
`composition.html_normalized` and warnings identify normalization; inspect the
preview when supplying authored HTML. Exact original HTML bytes are not preserved.

Image sources may be absolute HTTP(S) URLs or `cid:` references. Remote images can
be blocked by recipients' settings and are never fetched by this CLI. Local paths
and data URLs are not embedded automatically. To include an inline image, create a
draft referencing `![Logo](cid:company-logo)` in Markdown or
`<img src="cid:company-logo" alt="Logo">` in HTML, then attach the file:

```sh
m365 mail attachment add DRAFT_ID --file logo.png --inline --content-id company-logo
m365 mail draft get DRAFT_ID --body-format html
m365 mail draft send DRAFT_ID --request-id logo-update-001
```

The CLI does not resolve missing CIDs automatically. HTML returned by `mail get`
is untrusted service data in JSON, never executed or opened automatically. Default
reads still request text; `mail get` and `mail draft get --body-format html` request
HTML explicitly. Check `Body.BodyType`, `Body.Value` and truncation metadata.

Reply/forward content uses Outlook's native `NewBodyContent`, preserving the
service-generated quoted thread. The CLI does not reconstruct the original body.
Formatting applies to your new introduction; the quoted history retains its own
formatting. Replies/forwards create drafts unless `--send` is supplied. No automatic
signature is inserted, and no arbitrary HTML sanitization is applied to existing
quoted history during these native operations.

Mail/event source bodies are limited to 1 MiB of UTF-8, rendered HTML to 4 MiB and 20,000
HTML elements. Style files are limited to 16 KiB. Exchange may impose additional
message limits. Calendar RSVP comments and task body formatting are
unchanged; these are separate service capabilities. The CLI does not construct a custom MIME multipart body;
Exchange controls delivery encoding and any generated text alternative.

## Teams messages, channel posts and replies

```sh
m365 teams message render --body-format markdown --body-file update.md --output preview.html
m365 teams message styles
m365 teams message send --chat CHAT_ID --body-format markdown --body-file update.md --request-id rich-chat-001
m365 teams message send --channel CHANNEL_ID --body-format markdown --body-file update.md --request-id rich-post-001
m365 teams message reply --channel CHANNEL_ID --reply-to ROOT_MESSAGE_ID --body-format html --body '<p><strong>Ready.</strong></p>' --style minimal --request-id rich-reply-001
m365 teams message update MESSAGE_ID --channel CHANNEL_ID --body-format markdown --body-file revised.md --if-version VERSION --request-id rich-edit-001
```

These routes use exactly the shared renderer and style defaults above. Local
`teams message render` also returns `message_content`, the final Teams HTML
envelope (including escaped markup for literal text). Preview/styles need no
login. Sending is immediate; Teams has no draft stage here. Updating replaces
the entire body; an empty rendered message is rejected.

Plain text retains the 20,000-byte source bound and its existing escaping.
Rich source is bounded to 1 MiB and the final rendered HTML to **20,000 UTF-8
bytes**, with no automatic truncation or splitting. Images, including Markdown,
CID and remote HTML images, are refused. Links are supported; attachments, hosted
content, cards and mentions are not introduced by formatting. Native code widgets
and syntax highlighting are not promised; code uses ordinary HTML.

Teams can normalize entities, block indentation and CSS separator spacing. Rich
readback compares HTML structure, text, attributes and styles, preserving
preformatted/code whitespace. Missing formatting or changed targets still fail
with `readback_failed`, without resending. Successful normalized readback is
reported in `body_verification`. Normalized stored HTML is still untrusted data.
Client fonts, spacing and table presentation may differ from the preview.
Completed plain-text request IDs keep their prior fingerprints; rich IDs bind
rendered HTML and resolved styles. Existing ownership, version and hidden-payload
edit guards remain. See [TEAMS.md](TEAMS.md) for discovery and channel restrictions.

## Service contracts

Implementation references: Microsoft's [Body format model](https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/body),
[Markdown parser](https://markdown-it-py.readthedocs.io/en/latest/using.html), and
[HTML sanitizer](https://nh3.readthedocs.io/en/latest/). These libraries ship as
Python package dependencies; no additional mail CLI is required. The underlying
Outlook Web JSON transport remains undocumented.
