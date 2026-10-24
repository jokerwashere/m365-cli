# Mail attachments

Use `m365 mail attachment` for files, inline files, attached Outlook items and
cloud-link metadata. Attachment IDs are opaque: use them exactly as returned,
including IDs several levels inside attached messages.

## Discover and inspect

```sh
m365 mail attachment list 'MAIL_ID'
m365 mail attachment list 'MAIL_ID' --recursive
m365 mail attachment list 'MAIL_ID' --recursive --max-depth 20 --max-items 500
m365 mail attachment get 'ATTACHMENT_ID' --mail-id 'ROOT_MAIL_ID'
```

`list` returns metadata with `kind`, `id`, `name`, `content_type`, `is_inline`,
`depth`, `parent_attachment_id`, `ancestor_attachment_ids`, and `root_mail_id`.
A recursive list is flat and depth-first: root attachments have depth 0. Each
attached item is read to discover its immediate children. Check `complete` and
`truncations`: limits, repeated IDs and unknown attachment types are explicit.
Defaults are depth 10 and 200 entries; supported limits are depth 0–30 and
1–1000 entries. A nonrecursive result describes only the root collection.
An API/read failure fails the command rather than returning a misleading complete
result. Concurrent mailbox edits can invalidate IDs; this is not a snapshot.

`get` without `--output` inspects structured JSON. An attached item's properties
are under `Item`, including its body and immediate attachment metadata. It does
not hydrate the whole tree or output base64 file payloads. Use the recursive list
and then inspect/download the child IDs you need.

Every observed attachment receives a local `stable_id` and reference metadata.
Its `web_url` opens the **root mail**, including for deeply nested attachments;
it is not a direct file-download link. `--mail-id` verifies membership and supplies
that root when it is not already known from the local reference registry. Save
both the stable handle and the provider ID. Relist after moving a mail item or
replacing an attachment; provider attachment IDs can change. Attached items are
copies inside a message, not independently editable mailbox items.

## Download and export

```sh
m365 mail attachment get 'FILE_ID' --mail-id 'MAIL_ID' --output ./report.pdf
m365 mail attachment get 'ATTACHED_MESSAGE_ID' --mail-id 'MAIL_ID' --output ./message.eml
m365 mail attachment get 'NESTED_FILE_ID' --mail-id 'MAIL_ID' --output ./nested.bin
m365 mail attachment get 'ATTACHED_ITEM_ID' --mail-id 'MAIL_ID' --format json --output ./item.json
m365 mail attachment get 'CLOUD_ATTACHMENT_ID' --mail-id 'MAIL_ID' --output ./cloud-link.json
```

| Attachment | Default (`--format auto`) | `--format raw` | `--format json` |
|---|---|---|---|
| File, including inline | Exact decoded bytes | Exact decoded bytes | Metadata without bytes |
| Attached message | Microsoft's native export, normally MIME/EML, including nested attachments | Native export | Structured item properties and child metadata |
| Attached event, contact, task, distribution list or other item | Structured JSON | Microsoft's native export, if available | Structured item properties and child metadata |
| Cloud/reference or link | JSON descriptor, including `source_url` when returned | Explicit no-embedded-content error | JSON descriptor |
| Unknown future attachment type | JSON metadata | Explicit unsupported-type error | JSON metadata |

The result reports `export_format`, `export_content_type`, byte count and exact
saved path. Native exports are classified from the returned bytes, not the
service's filename: an `.ics` or `.vcf` label can still accompany MIME content. A raw event export may be a MIME message containing a calendar part.
Contact/task MIME conversion can omit properties. Use JSON when you need those
structured properties; neither format is a lossless mailbox backup.

Output paths must be explicit and new. Files are created with mode 0600. Names
from mail are never turned into local paths, existing files are never overwritten,
and attachments are never opened or executed. A failed local write may leave a
partial file. Binary downloads are capped at 50 MiB per attachment. Uploaded `.eml`,
`.msg`, archives, encrypted files and other ordinary file attachments remain exact
bytes; the CLI does not unpack them or bypass encryption/tenant restrictions.
Native nested-item traversal follows Outlook's attachment hierarchy.

Native item downloads use the same verified Outlook login and a temporary
Microsoft download token held only in memory. Requests go to the fixed
`attachments.office.net` service with redirects disabled. Outlook bearer/canary
headers are not forwarded there. No extra sign-in or token cache is introduced.
Cloud links are never fetched automatically, and Outlook credentials are never
sent to their destinations. Access to a linked file remains subject to its own
permissions. If `GetAttachment` cannot read a cloud reference, supply `--mail-id`
so the CLI can use parent metadata.

## Add and remove

```sh
m365 mail attachment add 'DRAFT_ID' --file ./report.pdf --request-id 'attach-report-001'
m365 mail attachment add 'DRAFT_ID' --file ./image.png --inline --content-id 'image@example.test' --request-id 'attach-inline-001'
m365 mail attachment add 'DRAFT_ID' --item-id 'SOURCE_MAILBOX_ITEM_ID' --name 'Original message' --request-id 'attach-item-001'
m365 mail attachment delete 'DRAFT_ID' --attachment-id 'TOP_LEVEL_ATTACHMENT_ID' --request-id 'remove-attachment-001'
```

Files may have any content type; uploads are capped at 10 MiB. `--content-type`
overrides filename-based MIME detection. `--name` changes the display name.
Inline uploads require both `--inline` and `--content-id`; setting a CID does not
insert an HTML body reference. Use the matching `cid:` URL in an HTML body; see
[COMPOSITION.md](COMPOSITION.md).

`--item-id` copies an existing **Outlook mailbox item**, including its attachments,
into the destination draft. It leaves the source unchanged. It is not a To Do or
Planner task ID. A draft cannot be attached to itself. Creating cloud-reference
attachments is not supported.
Attach a local file or include an authorized sharing link in the body instead.

Add/delete work only on drafts. `--if-change-key`, `--request-id` and `--dry-run`
follow the normal CLI conventions. Creation passes the current parent change key;
removal checks it before dispatch, but the deletion endpoint has no atomic parent
version condition. Nested attachments cannot be edited/deleted in place: delete
the top-level attached item and add a replacement. Deleting an outer attachment
marks known descendant references as stale/deleted. There is no automatic retry
or fallback write.

## Service contracts

Microsoft's [attachment model](https://learn.microsoft.com/en-us/exchange/client-developer/exchange-web-services/attachments-and-ews-in-exchange)
and [GetAttachment contract](https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/getattachment)
describe file and attached-item semantics. The Outlook Web JSON and download
adapters remain undocumented and can change; schema changes fail with explicit
errors rather than selecting a nested file as the outer attachment.
