# Teams messaging

Teams uses the same `m365` Python CLI, state directory, and dedicated Chrome
profile as mail/calendar. Sign-in happens through Teams Web; all commands access
service data through HTTP, without clicking UI controls or using accessibility.
No Entra application registration is required.

```sh
m365 --account you@example.com teams login
m365 teams status
m365 teams person search 'Person name'
m365 teams person get person@example.com
m365 teams chat list --limit 25
m365 teams chat self
m365 teams chat resolve person@example.com
m365 teams chat get 'CHAT_ID'
m365 teams chat members 'CHAT_ID'
m365 teams team list
m365 teams channel list --team 'TEAM_ID'
m365 teams channel get 'CHANNEL_ID'
```

`teams login` opens Teams in the saved profile and waits up to 300 seconds for
normal sign-in. SSO may complete automatically. `teams status` verifies the
account and tenant against the server. The first verified Teams tenant is saved
in `config.json`; subsequent commands reject a different tenant. An explicit
`m365 teams --tenant TENANT_UUID ...` supplies the expected tenant for that call.
`--account EMAIL` remains a global argument before `teams`.

## Messages

Choose exactly one `--chat` or `--channel` using an ID returned by discovery.
Sending is immediate; Teams messages do not have a draft stage in this CLI.

```sh
m365 teams message list --chat 'CHAT_ID' --limit 20
m365 teams message list --chat 'CHAT_ID' --limit 20 --cursor 'NEXT_CURSOR'
m365 teams message get 'MESSAGE_ID' --chat 'CHAT_ID'
m365 teams message send --chat 'CHAT_ID' --body 'Hello.' --dry-run
m365 teams message send --chat 'CHAT_ID' --body-file message.txt --request-id 'chat-send-001'
m365 teams message send --channel 'CHANNEL_ID' --body 'Project update.' --request-id 'channel-post-001'
m365 teams message reply --channel 'CHANNEL_ID' --reply-to 'ROOT_MESSAGE_ID' --body 'Follow-up.' --request-id 'channel-reply-001'
m365 teams message list --channel 'CHANNEL_ID' --reply-to 'ROOT_MESSAGE_ID'
m365 teams message update 'MESSAGE_ID' --chat 'CHAT_ID' --body 'Revised text.' --if-version 'VERSION' --request-id 'chat-edit-001'
m365 teams message delete 'MESSAGE_ID' --chat 'CHAT_ID' --request-id 'chat-delete-001'
```

Bodies default to literal UTF-8 plain text, HTML-escaped before submission, with
a 20,000-byte source limit. Send, reply and update also accept
`--body-format markdown|html`, using the shared mail/calendar renderer, sanitizer,
`--style outlook|minimal` and `--style-file` schema. `--body-file -` reads bounded
UTF-8 stdin. Returned `content` is Teams HTML; treat it as
untrusted data, and do not execute embedded content. Message results preserve
`id`, `version`, `messagetype`, sender, timestamps, and selected properties.
Unknown service fields and embedded credential-bearing metadata are not exported.

A successful send returns `submitted: true`, `thread_id`, and a read-back
`message`. This acknowledges submission, not recipient delivery or reading.
Updates/deletions also require read-back verification. Rich bodies are compared
by HTML structure, text, attributes and styles, allowing harmless entity,
indentation and CSS-spacing rewrites. Code/preformatted whitespace is preserved.
Lost formatting, changed links or unexpected body types fail verification.
`body_verification.server_normalized` reports a successful serialization change;
plain-text verification remains exact. If verification fails, inspect the
returned item identifiers; do not blindly resend or repeat the mutation.

Only your own ordinary text/HTML messages can be edited or deleted. Editing a
message containing files, cards, mentions, or quoted-message metadata is refused.
Edits use the current server version in `If-Match`, even when `--if-version` is
omitted. Providing `--if-version` also protects the version you reviewed.

**Deletes are soft deletes and have no atomic version guard.** The deletion
transport does not enforce `If-Match`. Supplying `--if-version` to deletion therefore
returns `unsupported_precondition` before any request. Deletion verifies ownership
and then checks for a tombstone with empty content. It cannot guarantee that a
concurrent edit will be preserved. Retention records and deletion placeholders may
remain. Restore/undelete and permanent deletion are not implemented.

Native reply threading is supported for channel root posts. Use ordinary send for
chat messages; quoted/native chat replies are outside this release. Channel writes
are limited to standard channels. Server policy may further restrict posting,
editing, or deleting. Read access can include meeting chats and other channel types
where the signed-in user has access, subject to the channel's permissions and policy.

## Markdown, HTML and preview

```sh
m365 teams message render --body-format markdown --body-file update.md --output preview.html
m365 teams message styles
m365 teams message send --chat 'CHAT_ID' --body-format markdown --body-file update.md --request-id 'rich-chat-001'
m365 teams message send --channel 'CHANNEL_ID' --body-format markdown --body-file update.md --request-id 'rich-post-001'
m365 teams message reply --channel 'CHANNEL_ID' --reply-to 'ROOT_MESSAGE_ID' --body-format html --body '<p><strong>Update:</strong> Ready.</p>' --style minimal --request-id 'rich-reply-001'
m365 teams message update 'MESSAGE_ID' --channel 'CHANNEL_ID' --body-format markdown --body-file revised.md --if-version 'VERSION' --request-id 'rich-edit-001'
```

The rich default is the same `outlook` style as mail/calendar: Helvetica Neue,
Helvetica and Arial fallbacks, 12 pt, conservative colors and spacing. `minimal`
adds no presentation styles. The same JSON style overrides work here. Markdown
supports headings, emphasis, lists, links, quotes, tables, inline and fenced code;
raw HTML inside Markdown is escaped. Use `--body-format html` for authored HTML.
Scripts, unsafe URLs/CSS and unsupported attributes are sanitized by the shared
renderer. Inspect `composition` metadata and warnings.

Teams may ignore fonts, spacing and table styling in its clients. The local
preview shows the submitted HTML, not a guarantee of identical Teams appearance.
The API checks stored content; it does not promise native code-block widgets or
syntax highlighting. Images (including Markdown images and CID/remote HTML
images) are rejected; use ordinary links. This does not add attachments, cards,
mentions or announcements. HTML `<at>` markup does not create a mention.

Rich source has the shared 1 MiB input bound; the **rendered HTML must fit 20,000
UTF-8 bytes**. The CLI never truncates or automatically splits messages. Empty
rendered messages are rejected; use deletion to remove a message. Edits replace
the complete body and retain the existing ownership/version/hidden-payload guards.
Preview/styles run locally without authentication. Preview files never overwrite
existing files. Rich request IDs bind source, rendered HTML and resolved styles;
old plain-text request IDs retain their meaning. See [COMPOSITION.md](COMPOSITION.md).

## Native drafts

Native Teams draft management is **not supported** by this CLI.
`teams message render` prepares a local preview. It neither creates a Teams draft
nor returns a native draft ID or shareable draft URL. Sending remains explicit
and immediate. Do not use scheduled messages as an unsent-draft substitute.

## Chats and groups

`chat self` requires an unambiguous internal direct chat, a fixed roster, and a
fresh server roster containing only the signed-in user. A group with one currently
resolved member can still have pending invitees; it is never selected as a self
chat. Discovery fails if these checks cannot establish the destination. The
native `48:notes` self-chat stream is not supported by the ordinary-thread adapter.
For an explicit self-only destination, create a named group without `--to`.

```sh
m365 teams chat create --to person@example.com --request-id 'direct-chat-001'
m365 teams chat create --group --topic 'Planning' --to first@example.com --to second@example.com --request-id 'group-create-001'
m365 teams chat rename 'GROUP_CHAT_ID' --topic 'Planning updates' --request-id 'group-rename-001'
m365 teams chat leave 'GROUP_CHAT_ID' --request-id 'group-leave-001'
```

People are resolved by an exact, unique email match; a fuzzy name is never silently
chosen as a recipient. Direct-chat creation first resolves and reuses an existing
chat. New direct chats use the service's fixed, unique-roster contract. Group
creation always creates a new named group; without `--to`, it creates a self-only
group. Creation can notify the selected people. Rename and leave apply only to
group chats. Leaving does not delete the conversation for everyone. There is no
team/channel creation or deletion, or editing other people's group membership.

## Pagination and search boundary

Chat/team/channel listings page the available bootstrap inventory with `--offset`
and `--limit` (1–500). `inventory_complete: false` and `has_more: null` mean the
server did not confirm a complete inventory; do not claim that missing items do
not exist. Offsets do not retrieve historical data omitted by the bootstrap API.

Message history uses `next_cursor`. Reuse it only with the same account, tenant,
chat/channel, and reply chain. Cursors contain continuation metadata, not tokens;
treat them as opaque and private. Pages can overlap while new messages arrive, so
deduplicate by message ID. Stop when `next_cursor` is absent; `has_more: null`
means no further usable continuation was supplied, not proof of exhaustive history.
No command explicitly changes consumption horizons/read receipts. Loading Teams
Web still causes its normal background session/presence activity.

**There is no Teams message-search command.** People search remains available.
Agents may inspect bounded history pages and filter those results locally, but
must describe the inspected range and never label it a complete mailbox-wide or
Teams-wide search. Mail and calendar search remain available in their own groups.

Files/attachments, mentions, reactions, federation-specific workflows, private or
shared channel writes, calls, presence control, recordings, transcripts, membership
administration, and tasks are outside this first Teams release.

## Reliability and recovery

Use a unique stable `--request-id` for every intended mutation. The local ledger
binds Teams requests to service, account, tenant, target, and payload. Completed
replays return their cached result with `replayed: true`; pending attempts block
re-execution. This is local duplicate prevention, not server-side exactly-once
execution. Keep the same `M365_HOME` across processes.

There are no automatic API write retries or speculative endpoint fallbacks.
Authentication failures require `m365 teams login`; `conflict` requires a fresh
read and reconciliation; `rate_limited` may include `retry_after_seconds`.
`write_outcome_unknown`, interruption, and `readback_failed` require inspection of
the affected chat/channel and `m365 operation get REQUEST_ID` before another write.
A new request ID is appropriate only after the prior outcome is reconciled.

See [AUTH.md](AUTH.md), [AGENTS.md](AGENTS.md), and [CLI.md](CLI.md).

## Native links

Chats, teams, channels and messages return stable CLI references and Microsoft Teams URLs. Message IDs are scoped to their thread; channel replies retain the root-message context. People links open a chat without sending. See [REFERENCES.md](REFERENCES.md).
