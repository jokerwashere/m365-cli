# Command reference

All commands return JSON except `--help` and `--version`. Opaque IDs come from
previous command results. Global `--account EMAIL` appears **before** the group.
Mutation options appear after the command.

```sh
m365 --account you@example.com mail send --to person@example.com --subject 'Hello' --body 'Hello from M365 CLI.' --request-id 'hello-001'
```

## Shared behavior

- Online commands use browser authentication followed by direct HTTP. Outlook and
  SharePoint disconnect after bootstrap; status results report the transport.
  Each invocation still needs browser authentication. No separate credential cache,
  page-based fallback or automatic write retry exists. See [AUTH.md](AUTH.md).

- Resource results include `stable_id`, `web_url` and link-scope metadata. Resolve
  a stored handle with `m365 reference get STABLE_ID`. Commands accept provider IDs;
  files/lists also accept compatible stored SharePoint handles.
  See [REFERENCES.md](REFERENCES.md) for move tracking and native-link limits.

- Mail, calendar and Teams message composition defaults to plain text; `--body-format html|markdown` enables
  rich bodies. `--body-file -` reads UTF-8 stdin. See [COMPOSITION.md](COMPOSITION.md)
  for local previews, style files, HTML reads and supported formatting.
- Recipient flags may be repeated. Updates replace the specified recipient list;
  `--clear-*` empties it. Omitted fields are left unchanged.
- Mail/calendar: `--if-change-key` rejects a write if the reviewed item has changed. All item
  updates also use the current server change key and reject racing overwrites.
- `--request-id` prevents repeated local execution of the same mutation. Keep the
  exact input and state directory across retries. Do not automatically retry an
  ambiguous operation using a new ID.
- `--dry-run` validates local input and prints a plan. Server-dependent checks are
  deferred until real execution.
- List limits are 1–500. Offsets are nonnegative. Server pagination can change if
  the mailbox changes; use returned IDs to avoid processing an item twice.
- Built-in mail folders include `inbox`, `drafts`, `sentitems`, `deleteditems`,
  `archive`, `junkemail`, and `msgfolderroot`. Other folders use IDs.
- Calendar timestamps need an explicit UTC offset. Event list/search queries span
  at most 366 days; availability/slot queries span at most 31 days. Event search is
  a literal substring in subject/location within its window.
- Mail/calendar deletion moves items to Deleted Items. It does not permanently erase them.
  Empty custom folders may be deleted; built-in folders are protected.
- A successful send returns `submitted: true`; verify Sent Items if needed.
  This is not a recipient delivery receipt.

Attachment types, recursion, native/JSON exports and cloud-link limits are in
[ATTACHMENTS.md](ATTACHMENTS.md). Recursive results report completeness explicitly.

Teams uses the same JSON, body-file, dry-run and request-ID conventions. See
[TEAMS.md](TEAMS.md) for setup, pagination, and supported limits. Teams edits use
`--if-version`; Teams deletion rejects that option because the service ignores
atomic delete preconditions. `teams message reply` supports channel root posts.
Teams history uses opaque `--cursor`; inventory uses offsets. Teams has people
search, but no message search.
Teams sends, channel replies and edits share mail/calendar body formats, sanitizer
and styles. `teams message render/styles` work locally. Rich Teams HTML is limited
to 20,000 UTF-8 bytes; images are refused and client presentation can vary. Sending
remains immediate; rich readback tolerates only harmless serialization changes.
Native Teams drafts are not supported. Preview files are local and have no native draft ID/URL.

Category definitions/colors, shared To Do categories, and Planner label lifecycle
are documented in [LABELS.md](LABELS.md). Category-only calendar changes are
personal metadata and send no meeting updates.

Multi-person free/busy, meeting-slot selection and unknown/working-hours handling
are documented in [SCHEDULING.md](SCHEDULING.md). These commands never book or send.

SharePoint files, OneDrive and typed Lists records use `files` / `lists`. See
[SHAREPOINT.md](SHAREPOINT.md) for site-scoped authentication, stable handles,
new-name uploads, manifests, version guards, unconditional-operation opt-ins,
pagination and explicit unsupported operations.

## Examples beyond the quick start

```sh
m365 mail reply 'MESSAGE_ID' --body-file reply.txt
m365 mail reply 'MESSAGE_ID' --all --body-file reply.txt --send --request-id 'reply-001'
m365 mail forward 'MESSAGE_ID' --to person@example.com --body 'For your review.'
m365 mail copy 'MESSAGE_ID' --folder 'FOLDER_ID'
m365 mail delete 'MESSAGE_ID' --if-change-key 'CHANGE_KEY' --request-id 'delete-001'
m365 mail move 'DELETED_MESSAGE_ID' --folder inbox
m365 mail attachment list 'MESSAGE_ID' --recursive
m365 mail attachment get 'NESTED_ATTACHMENT_ID' --mail-id 'MESSAGE_ID'
m365 mail attachment get 'ATTACHED_ITEM_ID' --mail-id 'MESSAGE_ID' --format json --output ./item.json
m365 mail attachment get 'ATTACHMENT_ID' --output ./report.pdf
m365 mail attachment add 'DRAFT_ID' --file ./report.pdf
m365 mail attachment add 'DRAFT_ID' --item-id 'SOURCE_MESSAGE_ID'
m365 mail attachment delete 'DRAFT_ID' --attachment-id 'ATTACHMENT_ID'
m365 folder create --kind calendar --name 'Project calendar'
m365 folder update 'FOLDER_ID' --name 'Renamed calendar'
m365 calendar create --subject 'Planning' --start '2026-10-07T09:00:00+02:00' --end '2026-10-07T10:00:00+02:00' --attendee person@example.com --request-id 'meeting-001'
m365 calendar update 'EVENT_ID' --location 'Room 2' --notify --if-change-key 'CHANGE_KEY' --request-id 'meeting-update-001'
m365 calendar respond 'EVENT_ID' --response accept --notify --request-id 'accept-001'
m365 calendar delete 'EVENT_ID' --notify --request-id 'cancel-001'
```

To prepare a separate meeting and send it after review:

```sh
m365 calendar create --subject 'Prepared planning' --start '2031-02-17T09:00:00+01:00' --end '2031-02-17T10:00:00+01:00' --attendee person@example.com --save-only --request-id 'meeting-prepare-001'
m365 calendar update 'UNSENT_EVENT_ID' --location 'Room 2' --save-only --if-change-key 'CHANGE_KEY' --request-id 'meeting-prepare-edit-001'
m365 calendar get 'UNSENT_EVENT_ID' --body-format html
m365 calendar send 'UNSENT_EVENT_ID' --if-change-key 'LATEST_CHANGE_KEY' --request-id 'meeting-invite-001'
```

Reply and forward default to creating drafts; `--send` sends immediately.
For replies, `--to` explicitly overrides the To field; without it Outlook derives
recipients from the source message. Review recipient fields before sending a draft.
Adding calendar attendees at creation sends invitations unless `--save-only` is given.
Use `calendar update --save-only` for a verified never-sent single event, then
`calendar send EVENT_ID` to submit initial invitations. Inspect `MeetingRequestWasSent`
on `calendar get`; `IsDraft` can be false even for an unsent meeting. Verified unsent
meetings can be deleted silently. Sent meetings still require `--notify` for updates
and cancellation; `--save-only` is refused. Initial sending rejects already-sent,
received, cancelled, recurring, attendee-free and deleted events. Responses notify
the organizer only with `--notify`. See [SCHEDULING.md](SCHEDULING.md).

The sections below are generated from the parser. Regenerate with
`python scripts/generate_reference.py` after changing arguments.

## m365

```text
usage: m365 [-h] [--version] [--account ACCOUNT]
            {auth,operation,reference,mail,folder,calendar,teams,category,tasks,todo,planner,files,lists,maintainer} ...

Microsoft 365 mail, calendar, Teams, tasks, files and lists CLI. JSON output;
direct HTTP with browser authentication.

positional arguments:
  {auth,operation,reference,mail,folder,calendar,teams,category,tasks,todo,planner,files,lists,maintainer}
    operation           Inspect local duplicate-prevention records
    reference           Durable local references and SharePoint URL resolution
    folder              Mail and calendar folder management
    teams               Teams people, chats, channels, and messaging
    category            Outlook/To Do category definitions and assignments
    tasks               Task account status, capability inventory and
                        aggregate reads
    todo                Todo task operations
    planner             Planner task operations
    files               SharePoint/OneDrive files; browser session, explicit
                        site scope
    lists               SharePoint lists, typed records, history and
                        attachments
    maintainer          Maintainer-only capability inventory, probes, live
                        validation and circuit breakers

options:
  -h, --help            show this help message and exit
  --version             show program's version number and exit
  --account ACCOUNT     Expected signed-in mailbox; mismatch fails before
                        writing
```

## m365 auth

```text
usage: m365 auth [-h] {login,status} ...

positional arguments:
  {login,status}
    login         Open the dedicated Chrome profile and verify the account

options:
  -h, --help      show this help message and exit
```

## m365 auth login

```text
usage: m365 auth login [-h] [--account ACCOUNT] [--timeout TIMEOUT]

options:
  -h, --help         show this help message and exit
  --account ACCOUNT
  --timeout TIMEOUT
```

## m365 auth status

```text
usage: m365 auth status [-h]

options:
  -h, --help  show this help message and exit
```

## m365 operation

```text
usage: m365 operation [-h] {get} ...

positional arguments:
  {get}

options:
  -h, --help  show this help message and exit
```

## m365 operation get

```text
usage: m365 operation get [-h] request_id

positional arguments:
  request_id

options:
  -h, --help  show this help message and exit
```

## m365 reference

```text
usage: m365 reference [-h] {get,resolve} ...

positional arguments:
  {get,resolve}
    get          Return the latest locally observed ID and native URL
    resolve      Resolve an existing-access SharePoint file/folder/list-item
                 URL

options:
  -h, --help     show this help message and exit
```

## m365 reference get

```text
usage: m365 reference get [-h] stable_id

positional arguments:
  stable_id

options:
  -h, --help  show this help message and exit
```

## m365 reference resolve

```text
usage: m365 reference resolve [-h] [--site SITE] url

positional arguments:
  url

options:
  -h, --help   show this help message and exit
  --site SITE  SharePoint site root; required unless a stable handle or URL
               supplies it
```

## m365 mail

```text
usage: m365 mail [-h]
                 {list,search,get,render,styles,send,reply,forward,update,delete,move,copy,draft,attachment} ...

positional arguments:
  {list,search,get,render,styles,send,reply,forward,update,delete,move,copy,draft,attachment}
    render              Render a body locally; never connects or sends
    styles              Show local email style defaults and override fields
    send                Send immediately and save a copy in Sent Items
    reply               Create a response draft; --send sends immediately
    forward             Create a response draft; --send sends immediately
    update              Update read state, importance, categories, or flag
    delete              Move mail to Deleted Items; never permanently delete

options:
  -h, --help            show this help message and exit
```

## m365 mail list

```text
usage: m365 mail list [-h] [--limit LIMIT] [--offset OFFSET] [--folder FOLDER]

options:
  -h, --help       show this help message and exit
  --limit LIMIT
  --offset OFFSET
  --folder FOLDER
```

## m365 mail search

```text
usage: m365 mail search [-h] [--limit LIMIT] [--offset OFFSET] query

positional arguments:
  query

options:
  -h, --help       show this help message and exit
  --limit LIMIT
  --offset OFFSET
```

## m365 mail get

```text
usage: m365 mail get [-h] [--body-format {text,html}] id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --body-format {text,html}
```

## m365 mail render

```text
usage: m365 mail render [-h] [--body BODY | --body-file BODY_FILE]
                        [--body-format {text,html,markdown}]
                        [--style {outlook,minimal}] [--style-file STYLE_FILE]
                        [--output OUTPUT]

options:
  -h, --help            show this help message and exit
  --body BODY
  --body-file BODY_FILE
                        UTF-8 text file, or - for stdin
  --body-format {text,html,markdown}
                        Input format; Markdown is sent as styled HTML (default
                        text)
  --style {outlook,minimal}
                        HTML/Markdown styling: outlook (default; Helvetica
                        Neue/Arial) or minimal (no added styles)
  --style-file STYLE_FILE
                        JSON overrides for the outlook style; see mail styles,
                        calendar styles or teams message styles
  --output OUTPUT       Save a new HTML preview file (never overwrites)
```

## m365 mail styles

```text
usage: m365 mail styles [-h]

options:
  -h, --help  show this help message and exit
```

## m365 mail send

```text
usage: m365 mail send [-h] --subject SUBJECT [--body BODY |
                      --body-file BODY_FILE]
                      [--body-format {text,html,markdown}]
                      [--style {outlook,minimal}] [--style-file STYLE_FILE]
                      [--to TO] [--cc CC] [--bcc BCC]
                      [--request-id REQUEST_ID] [--dry-run]

options:
  -h, --help            show this help message and exit
  --subject SUBJECT
  --body BODY
  --body-file BODY_FILE
                        UTF-8 text file, or - for stdin
  --body-format {text,html,markdown}
                        Input format; Markdown is sent as styled HTML (default
                        text)
  --style {outlook,minimal}
                        HTML/Markdown styling: outlook (default; Helvetica
                        Neue/Arial) or minimal (no added styles)
  --style-file STYLE_FILE
                        JSON overrides for the outlook style; see mail styles,
                        calendar styles or teams message styles
  --to TO               Plain email address; repeat for multiple recipients
  --cc CC               Plain email address; repeat for multiple recipients
  --bcc BCC             Plain email address; repeat for multiple recipients
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
```

## m365 mail reply

```text
usage: m365 mail reply [-h] [--body BODY | --body-file BODY_FILE]
                       [--body-format {text,html,markdown}]
                       [--style {outlook,minimal}] [--style-file STYLE_FILE]
                       [--request-id REQUEST_ID] [--dry-run]
                       [--if-change-key IF_CHANGE_KEY] [--to TO] [--send]
                       [--all]
                       id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --body BODY
  --body-file BODY_FILE
                        UTF-8 text file, or - for stdin
  --body-format {text,html,markdown}
                        Input format; Markdown is sent as styled HTML (default
                        text)
  --style {outlook,minimal}
                        HTML/Markdown styling: outlook (default; Helvetica
                        Neue/Arial) or minimal (no added styles)
  --style-file STYLE_FILE
                        JSON overrides for the outlook style; see mail styles,
                        calendar styles or teams message styles
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-change-key IF_CHANGE_KEY
                        Reject if the item changed since this change key was
                        read
  --to TO               Override reply recipients, or specify forward
                        recipients
  --send
  --all                 Reply to all original recipients
```

## m365 mail forward

```text
usage: m365 mail forward [-h] [--body BODY | --body-file BODY_FILE]
                         [--body-format {text,html,markdown}]
                         [--style {outlook,minimal}] [--style-file STYLE_FILE]
                         [--request-id REQUEST_ID] [--dry-run]
                         [--if-change-key IF_CHANGE_KEY] [--to TO] [--send]
                         id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --body BODY
  --body-file BODY_FILE
                        UTF-8 text file, or - for stdin
  --body-format {text,html,markdown}
                        Input format; Markdown is sent as styled HTML (default
                        text)
  --style {outlook,minimal}
                        HTML/Markdown styling: outlook (default; Helvetica
                        Neue/Arial) or minimal (no added styles)
  --style-file STYLE_FILE
                        JSON overrides for the outlook style; see mail styles,
                        calendar styles or teams message styles
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-change-key IF_CHANGE_KEY
                        Reject if the item changed since this change key was
                        read
  --to TO               Override reply recipients, or specify forward
                        recipients
  --send
```

## m365 mail update

```text
usage: m365 mail update [-h] [--request-id REQUEST_ID] [--dry-run]
                        [--if-change-key IF_CHANGE_KEY] [--read {true,false}]
                        [--importance {Low,Normal,High}]
                        [--flag {NotFlagged,Flagged,Complete}]
                        [--category CATEGORY | --clear-categories]
                        id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-change-key IF_CHANGE_KEY
                        Reject if the item changed since this change key was
                        read
  --read {true,false}
  --importance {Low,Normal,High}
  --flag {NotFlagged,Flagged,Complete}
  --category CATEGORY
  --clear-categories
```

## m365 mail delete

```text
usage: m365 mail delete [-h] [--request-id REQUEST_ID] [--dry-run]
                        [--if-change-key IF_CHANGE_KEY]
                        id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-change-key IF_CHANGE_KEY
                        Reject if the item changed since this change key was
                        read
```

## m365 mail move

```text
usage: m365 mail move [-h] --folder FOLDER [--request-id REQUEST_ID]
                      [--dry-run] [--if-change-key IF_CHANGE_KEY]
                      id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --folder FOLDER
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-change-key IF_CHANGE_KEY
                        Reject if the item changed since this change key was
                        read
```

## m365 mail copy

```text
usage: m365 mail copy [-h] --folder FOLDER [--request-id REQUEST_ID]
                      [--dry-run] [--if-change-key IF_CHANGE_KEY]
                      id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --folder FOLDER
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-change-key IF_CHANGE_KEY
                        Reject if the item changed since this change key was
                        read
```

## m365 mail draft

```text
usage: m365 mail draft [-h] {create,update,get,delete,send} ...

positional arguments:
  {create,update,get,delete,send}

options:
  -h, --help            show this help message and exit
```

## m365 mail draft create

```text
usage: m365 mail draft create [-h] --subject SUBJECT [--body BODY |
                              --body-file BODY_FILE]
                              [--body-format {text,html,markdown}]
                              [--style {outlook,minimal}]
                              [--style-file STYLE_FILE] [--to TO] [--cc CC]
                              [--bcc BCC] [--request-id REQUEST_ID]
                              [--dry-run]

options:
  -h, --help            show this help message and exit
  --subject SUBJECT
  --body BODY
  --body-file BODY_FILE
                        UTF-8 text file, or - for stdin
  --body-format {text,html,markdown}
                        Input format; Markdown is sent as styled HTML (default
                        text)
  --style {outlook,minimal}
                        HTML/Markdown styling: outlook (default; Helvetica
                        Neue/Arial) or minimal (no added styles)
  --style-file STYLE_FILE
                        JSON overrides for the outlook style; see mail styles,
                        calendar styles or teams message styles
  --to TO               Plain email address; repeat for multiple recipients
  --cc CC               Plain email address; repeat for multiple recipients
  --bcc BCC             Plain email address; repeat for multiple recipients
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
```

## m365 mail draft update

```text
usage: m365 mail draft update [-h] [--subject SUBJECT] [--body BODY |
                              --body-file BODY_FILE]
                              [--body-format {text,html,markdown}]
                              [--style {outlook,minimal}]
                              [--style-file STYLE_FILE] [--to TO | --clear-to]
                              [--cc CC | --clear-cc] [--bcc BCC | --clear-bcc]
                              [--request-id REQUEST_ID] [--dry-run]
                              [--if-change-key IF_CHANGE_KEY]
                              id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --subject SUBJECT
  --body BODY
  --body-file BODY_FILE
                        UTF-8 text file, or - for stdin
  --body-format {text,html,markdown}
                        Input format; Markdown is sent as styled HTML (default
                        text)
  --style {outlook,minimal}
                        HTML/Markdown styling: outlook (default; Helvetica
                        Neue/Arial) or minimal (no added styles)
  --style-file STYLE_FILE
                        JSON overrides for the outlook style; see mail styles,
                        calendar styles or teams message styles
  --to TO               Plain email address; repeat for multiple recipients
  --clear-to            Remove all recipients in this field
  --cc CC               Plain email address; repeat for multiple recipients
  --clear-cc            Remove all recipients in this field
  --bcc BCC             Plain email address; repeat for multiple recipients
  --clear-bcc           Remove all recipients in this field
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-change-key IF_CHANGE_KEY
                        Reject if the item changed since this change key was
                        read
```

## m365 mail draft get

```text
usage: m365 mail draft get [-h] [--body-format {text,html}] id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --body-format {text,html}
```

## m365 mail draft delete

```text
usage: m365 mail draft delete [-h] [--request-id REQUEST_ID] [--dry-run]
                              [--if-change-key IF_CHANGE_KEY]
                              id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-change-key IF_CHANGE_KEY
                        Reject if the item changed since this change key was
                        read
```

## m365 mail draft send

```text
usage: m365 mail draft send [-h] [--request-id REQUEST_ID] [--dry-run]
                            [--if-change-key IF_CHANGE_KEY]
                            id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-change-key IF_CHANGE_KEY
                        Reject if the item changed since this change key was
                        read
```

## m365 mail attachment

```text
usage: m365 mail attachment [-h] {list,add,get,delete} ...

positional arguments:
  {list,add,get,delete}
    list                List typed attachments, optionally through nested
                        items
    add                 Attach a local file or a copy of an existing mailbox
                        item to a draft
    get                 Inspect any attachment, or export it to an explicit
                        path

options:
  -h, --help            show this help message and exit
```

## m365 mail attachment list

```text
usage: m365 mail attachment list [-h] [--recursive] [--max-depth MAX_DEPTH]
                                 [--max-items MAX_ITEMS]
                                 id

positional arguments:
  id                    Root mail ID

options:
  -h, --help            show this help message and exit
  --recursive           Traverse attached messages/items; inspect complete and
                        truncations
  --max-depth MAX_DEPTH
                        Maximum nested depth, 0–30; root attachments have
                        depth 0
  --max-items MAX_ITEMS
                        Maximum attachment entries, 1–1000
```

## m365 mail attachment add

```text
usage: m365 mail attachment add [-h] (--file FILE | --item-id ITEM_ID)
                                [--name NAME] [--content-type CONTENT_TYPE]
                                [--inline] [--content-id CONTENT_ID]
                                [--request-id REQUEST_ID] [--dry-run]
                                [--if-change-key IF_CHANGE_KEY]
                                id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --file FILE
  --item-id ITEM_ID     Existing mailbox item to attach, including its
                        attachments
  --name NAME           Attachment display name (never used as a local output
                        path)
  --content-type CONTENT_TYPE
                        File MIME type; inferred from the filename by default
  --inline              Mark a file inline; requires --content-id
  --content-id CONTENT_ID
                        Inline file CID, without angle brackets
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-change-key IF_CHANGE_KEY
                        Reject if the item changed since this change key was
                        read
```

## m365 mail attachment get

```text
usage: m365 mail attachment get [-h] [--mail-id MAIL_ID] [--output OUTPUT]
                                [--format {auto,raw,json}]
                                attachment_id

positional arguments:
  attachment_id

options:
  -h, --help            show this help message and exit
  --mail-id MAIL_ID     Root mail ID; verifies membership and supplies the
                        parent URL
  --output OUTPUT       Exact output path; must not already exist. Omit to
                        inspect JSON without file bytes
  --format {auto,raw,json}
                        Export: auto saves files/messages as bytes and other
                        types as JSON; raw requests native bytes; json saves
                        properties
```

## m365 mail attachment delete

```text
usage: m365 mail attachment delete [-h] --attachment-id ATTACHMENT_ID
                                   [--request-id REQUEST_ID] [--dry-run]
                                   [--if-change-key IF_CHANGE_KEY]
                                   id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --attachment-id ATTACHMENT_ID
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-change-key IF_CHANGE_KEY
                        Reject if the item changed since this change key was
                        read
```

## m365 folder

```text
usage: m365 folder [-h] {list,get,create,update,delete} ...

positional arguments:
  {list,get,create,update,delete}
    delete              Delete only an empty custom folder

options:
  -h, --help            show this help message and exit
```

## m365 folder list

```text
usage: m365 folder list [-h] [--parent PARENT] [--limit LIMIT]
                        [--offset OFFSET]

options:
  -h, --help       show this help message and exit
  --parent PARENT
  --limit LIMIT
  --offset OFFSET
```

## m365 folder get

```text
usage: m365 folder get [-h] id

positional arguments:
  id

options:
  -h, --help  show this help message and exit
```

## m365 folder create

```text
usage: m365 folder create [-h] --name NAME [--parent PARENT]
                          [--kind {mail,calendar}] [--request-id REQUEST_ID]
                          [--dry-run]

options:
  -h, --help            show this help message and exit
  --name NAME
  --parent PARENT
  --kind {mail,calendar}
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
```

## m365 folder update

```text
usage: m365 folder update [-h] --name NAME [--request-id REQUEST_ID]
                          [--dry-run] [--if-change-key IF_CHANGE_KEY]
                          id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --name NAME
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-change-key IF_CHANGE_KEY
                        Reject if the item changed since this change key was
                        read
```

## m365 folder delete

```text
usage: m365 folder delete [-h] [--request-id REQUEST_ID] [--dry-run]
                          [--if-change-key IF_CHANGE_KEY]
                          id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-change-key IF_CHANGE_KEY
                        Reject if the item changed since this change key was
                        read
```

## m365 calendar

```text
usage: m365 calendar [-h]
                     {availability,find-slots,calendars,list,search,get,render,styles,create,update,send,delete,respond} ...

positional arguments:
  {availability,find-slots,calendars,list,search,get,render,styles,create,update,send,delete,respond}
    availability        Read multiple people's free/busy
    find-slots          Find common meeting times; never books or sends
    render              Render an event body locally; never connects or books
    styles              Show shared mail/event style defaults and override
                        fields
    send                Send initial invitations for a saved, never-sent
                        single event

options:
  -h, --help            show this help message and exit
```

## m365 calendar availability

```text
usage: m365 calendar availability [-h] [--attendee ATTENDEE]
                                  [--optional-attendee OPTIONAL_ATTENDEE]
                                  [--exclude-self] --start START --end END
                                  [--interval INTERVAL]
                                  [--time-zone TIME_ZONE]

options:
  -h, --help            show this help message and exit
  --attendee ATTENDEE   Required person's or resource's SMTP address; repeat
                        as needed
  --optional-attendee OPTIONAL_ATTENDEE
  --exclude-self        Do not automatically include the signed-in organizer
  --start START         Window start with explicit UTC offset
  --end END             Exclusive window end; maximum 31 days
  --interval INTERVAL   Conservative free/busy granularity in minutes (5–60)
  --time-zone TIME_ZONE
                        Output and daily-hours timezone; IANA or Microsoft
                        name (default UTC)
```

## m365 calendar find-slots

```text
usage: m365 calendar find-slots [-h] [--attendee ATTENDEE]
                                [--optional-attendee OPTIONAL_ATTENDEE]
                                [--exclude-self] --start START --end END
                                [--interval INTERVAL] [--time-zone TIME_ZONE]
                                --duration DURATION [--work-start WORK_START]
                                [--work-end WORK_END] [--weekdays WEEKDAYS]
                                [--max-results MAX_RESULTS]
                                [--allow-tentative] [--respect-working-hours]

options:
  -h, --help            show this help message and exit
  --attendee ATTENDEE   Required person's or resource's SMTP address; repeat
                        as needed
  --optional-attendee OPTIONAL_ATTENDEE
  --exclude-self        Do not automatically include the signed-in organizer
  --start START         Window start with explicit UTC offset
  --end END             Exclusive window end; maximum 31 days
  --interval INTERVAL   Conservative free/busy granularity in minutes (5–60)
  --time-zone TIME_ZONE
                        Output and daily-hours timezone; IANA or Microsoft
                        name (default UTC)
  --duration DURATION   Meeting duration in minutes (5–1440)
  --work-start WORK_START
                        Earliest daily start in --time-zone (default 09:00)
  --work-end WORK_END   Latest daily finish in --time-zone (default 17:00)
  --weekdays WEEKDAYS   Comma-separated days; use all for every day
  --max-results MAX_RESULTS
                        Maximum ranked candidates (1–100)
  --allow-tentative     Accept tentative slots explicitly; busy/OOF/unknown
                        still block
  --respect-working-hours
                        Also require known published working hours for
                        required attendees; unknown/custom zones block
                        candidates
```

## m365 calendar calendars

```text
usage: m365 calendar calendars [-h]

options:
  -h, --help  show this help message and exit
```

## m365 calendar list

```text
usage: m365 calendar list [-h] [--limit LIMIT] [--offset OFFSET] --start START
                          --end END [--folder FOLDER]

options:
  -h, --help       show this help message and exit
  --limit LIMIT
  --offset OFFSET
  --start START
  --end END
  --folder FOLDER
```

## m365 calendar search

```text
usage: m365 calendar search [-h] [--limit LIMIT] [--offset OFFSET]
                            --start START --end END [--folder FOLDER]
                            query

positional arguments:
  query            Case-insensitive substring in subject or location within
                   the time window

options:
  -h, --help       show this help message and exit
  --limit LIMIT
  --offset OFFSET
  --start START
  --end END
  --folder FOLDER
```

## m365 calendar get

```text
usage: m365 calendar get [-h] [--body-format {text,html}] id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --body-format {text,html}
```

## m365 calendar render

```text
usage: m365 calendar render [-h] [--body BODY | --body-file BODY_FILE]
                            [--body-format {text,html,markdown}]
                            [--style {outlook,minimal}]
                            [--style-file STYLE_FILE] [--output OUTPUT]

options:
  -h, --help            show this help message and exit
  --body BODY
  --body-file BODY_FILE
                        UTF-8 text file, or - for stdin
  --body-format {text,html,markdown}
                        Input format; Markdown is sent as styled HTML (default
                        text)
  --style {outlook,minimal}
                        HTML/Markdown styling: outlook (default; Helvetica
                        Neue/Arial) or minimal (no added styles)
  --style-file STYLE_FILE
                        JSON overrides for the outlook style; see mail styles,
                        calendar styles or teams message styles
  --output OUTPUT       Save a new HTML preview file (never overwrites)
```

## m365 calendar styles

```text
usage: m365 calendar styles [-h]

options:
  -h, --help  show this help message and exit
```

## m365 calendar create

```text
usage: m365 calendar create [-h] [--request-id REQUEST_ID] [--dry-run]
                            [--save-only] [--folder FOLDER] --subject SUBJECT
                            --start START --end END [--location LOCATION]
                            [--body BODY | --body-file BODY_FILE]
                            [--body-format {text,html,markdown}]
                            [--style {outlook,minimal}]
                            [--style-file STYLE_FILE] [--all-day {true,false}]
                            [--show-as {Free,Tentative,Busy,OOF,WorkingElsewhere}]
                            [--attendee ATTENDEE]
                            [--optional-attendee OPTIONAL_ATTENDEE]

Create an event. Attendees receive invitations unless --save-only is supplied.

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --save-only           Save attendees without invitations; updates require a
                        verified never-sent single event
  --folder FOLDER
  --subject SUBJECT
  --start START
  --end END
  --location LOCATION
  --body BODY
  --body-file BODY_FILE
                        UTF-8 text file, or - for stdin
  --body-format {text,html,markdown}
                        Input format; Markdown is sent as styled HTML (default
                        text)
  --style {outlook,minimal}
                        HTML/Markdown styling: outlook (default; Helvetica
                        Neue/Arial) or minimal (no added styles)
  --style-file STYLE_FILE
                        JSON overrides for the outlook style; see mail styles,
                        calendar styles or teams message styles
  --all-day {true,false}
  --show-as {Free,Tentative,Busy,OOF,WorkingElsewhere}
  --attendee ATTENDEE
  --optional-attendee OPTIONAL_ATTENDEE
```

## m365 calendar update

```text
usage: m365 calendar update [-h] [--request-id REQUEST_ID] [--dry-run]
                            [--if-change-key IF_CHANGE_KEY] [--save-only |
                            --notify] [--subject SUBJECT] [--start START]
                            [--end END] [--location LOCATION] [--body BODY |
                            --body-file BODY_FILE]
                            [--body-format {text,html,markdown}]
                            [--style {outlook,minimal}]
                            [--style-file STYLE_FILE] [--all-day {true,false}]
                            [--show-as {Free,Tentative,Busy,OOF,WorkingElsewhere}]
                            [--attendee ATTENDEE | --clear-attendees]
                            [--optional-attendee OPTIONAL_ATTENDEE |
                            --clear-optional-attendees]
                            id

Edit an event. Use --save-only for a never-sent meeting or --notify to send
updates.

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-change-key IF_CHANGE_KEY
                        Reject if the item changed since this change key was
                        read
  --save-only           Save attendees without invitations; updates require a
                        verified never-sent single event
  --notify              Send meeting updates to attendees
  --subject SUBJECT
  --start START
  --end END
  --location LOCATION
  --body BODY
  --body-file BODY_FILE
                        UTF-8 text file, or - for stdin
  --body-format {text,html,markdown}
                        Input format; Markdown is sent as styled HTML (default
                        text)
  --style {outlook,minimal}
                        HTML/Markdown styling: outlook (default; Helvetica
                        Neue/Arial) or minimal (no added styles)
  --style-file STYLE_FILE
                        JSON overrides for the outlook style; see mail styles,
                        calendar styles or teams message styles
  --all-day {true,false}
  --show-as {Free,Tentative,Busy,OOF,WorkingElsewhere}
  --attendee ATTENDEE
  --clear-attendees
  --optional-attendee OPTIONAL_ATTENDEE
  --clear-optional-attendees
```

## m365 calendar send

```text
usage: m365 calendar send [-h] [--request-id REQUEST_ID] [--dry-run]
                          [--if-change-key IF_CHANGE_KEY]
                          id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-change-key IF_CHANGE_KEY
                        Reject if the item changed since this change key was
                        read
```

## m365 calendar delete

```text
usage: m365 calendar delete [-h] [--notify] [--request-id REQUEST_ID]
                            [--dry-run] [--if-change-key IF_CHANGE_KEY]
                            id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --notify              Send cancellations for sent meetings; verified never-
                        sent meetings are deleted silently
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-change-key IF_CHANGE_KEY
                        Reject if the item changed since this change key was
                        read
```

## m365 calendar respond

```text
usage: m365 calendar respond [-h] --response {accept,tentative,decline}
                             [--notify] [--body BODY | --body-file BODY_FILE]
                             [--request-id REQUEST_ID] [--dry-run]
                             [--if-change-key IF_CHANGE_KEY]
                             id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --response {accept,tentative,decline}
  --notify              Send the response to the organizer
  --body BODY
  --body-file BODY_FILE
                        UTF-8 text file, or - for stdin
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-change-key IF_CHANGE_KEY
                        Reject if the item changed since this change key was
                        read
```

## m365 teams

```text
usage: m365 teams [-h] [--tenant TENANT]
                  {status,login,person,chat,team,channel,message} ...

positional arguments:
  {status,login,person,chat,team,channel,message}
    status              Verify Teams account and tenant
    login               Open Teams in the persistent browser and await sign-in

options:
  -h, --help            show this help message and exit
  --tenant TENANT       Expected tenant UUID; defaults to the saved Teams
                        tenant
```

## m365 teams status

```text
usage: m365 teams status [-h]

options:
  -h, --help  show this help message and exit
```

## m365 teams login

```text
usage: m365 teams login [-h] [--timeout TIMEOUT]

options:
  -h, --help         show this help message and exit
  --timeout TIMEOUT
```

## m365 teams person

```text
usage: m365 teams person [-h] {search,get} ...

positional arguments:
  {search,get}
    search      Find people by name or email (not message search)
    get         Resolve one exact email address

options:
  -h, --help    show this help message and exit
```

## m365 teams person search

```text
usage: m365 teams person search [-h] query

positional arguments:
  query

options:
  -h, --help  show this help message and exit
```

## m365 teams person get

```text
usage: m365 teams person get [-h] email

positional arguments:
  email

options:
  -h, --help  show this help message and exit
```

## m365 teams chat

```text
usage: m365 teams chat [-h]
                       {list,self,resolve,get,members,create,rename,leave} ...

positional arguments:
  {list,self,resolve,get,members,create,rename,leave}
    self                Get your self-chat ID
    resolve             Find an existing direct chat by exact email
    members             Read the chat roster
    create              Reuse/create a direct chat, or create a named group
    rename              Rename a group chat
    leave               Leave a group chat; does not delete it for others

options:
  -h, --help            show this help message and exit
```

## m365 teams chat list

```text
usage: m365 teams chat list [-h] [--limit LIMIT] [--offset OFFSET]
                            [--kind {direct,group,meeting}]

options:
  -h, --help            show this help message and exit
  --limit LIMIT
  --offset OFFSET
  --kind {direct,group,meeting}
```

## m365 teams chat self

```text
usage: m365 teams chat self [-h]

options:
  -h, --help  show this help message and exit
```

## m365 teams chat resolve

```text
usage: m365 teams chat resolve [-h] email

positional arguments:
  email

options:
  -h, --help  show this help message and exit
```

## m365 teams chat get

```text
usage: m365 teams chat get [-h] id

positional arguments:
  id

options:
  -h, --help  show this help message and exit
```

## m365 teams chat members

```text
usage: m365 teams chat members [-h] id

positional arguments:
  id

options:
  -h, --help  show this help message and exit
```

## m365 teams chat create

```text
usage: m365 teams chat create [-h] [--to TO] [--group] [--topic TOPIC]
                              [--request-id REQUEST_ID] [--dry-run]

options:
  -h, --help            show this help message and exit
  --to TO               Exact email; repeat for group members
  --group
  --topic TOPIC         Required with --group
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
```

## m365 teams chat rename

```text
usage: m365 teams chat rename [-h] --topic TOPIC [--request-id REQUEST_ID]
                              [--dry-run]
                              id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --topic TOPIC
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
```

## m365 teams chat leave

```text
usage: m365 teams chat leave [-h] [--request-id REQUEST_ID] [--dry-run] id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
```

## m365 teams team

```text
usage: m365 teams team [-h] {list} ...

positional arguments:
  {list}

options:
  -h, --help  show this help message and exit
```

## m365 teams team list

```text
usage: m365 teams team list [-h] [--limit LIMIT] [--offset OFFSET]

options:
  -h, --help       show this help message and exit
  --limit LIMIT
  --offset OFFSET
```

## m365 teams channel

```text
usage: m365 teams channel [-h] {list,get} ...

positional arguments:
  {list,get}

options:
  -h, --help  show this help message and exit
```

## m365 teams channel list

```text
usage: m365 teams channel list [-h] --team TEAM [--limit LIMIT]
                               [--offset OFFSET]

options:
  -h, --help       show this help message and exit
  --team TEAM
  --limit LIMIT
  --offset OFFSET
```

## m365 teams channel get

```text
usage: m365 teams channel get [-h] id

positional arguments:
  id

options:
  -h, --help  show this help message and exit
```

## m365 teams message

```text
usage: m365 teams message [-h]
                          {render,styles,list,get,send,reply,update,delete} ...

positional arguments:
  {render,styles,list,get,send,reply,update,delete}
    render              Preview a text/HTML/Markdown Teams body locally
    styles              Show shared style defaults and Teams message limits
    list
    get
    send
    reply               Reply to a channel root post
    update
    delete              Soft-delete your message (no atomic version guard)

options:
  -h, --help            show this help message and exit
```

## m365 teams message render

```text
usage: m365 teams message render [-h] [--body BODY | --body-file BODY_FILE]
                                 [--body-format {text,html,markdown}]
                                 [--style {outlook,minimal}]
                                 [--style-file STYLE_FILE] [--output OUTPUT]

options:
  -h, --help            show this help message and exit
  --body BODY
  --body-file BODY_FILE
                        UTF-8 text file, or - for stdin
  --body-format {text,html,markdown}
                        Input format; Markdown is sent as styled HTML (default
                        text)
  --style {outlook,minimal}
                        HTML/Markdown styling: outlook (default; Helvetica
                        Neue/Arial) or minimal (no added styles)
  --style-file STYLE_FILE
                        JSON overrides for the outlook style; see mail styles,
                        calendar styles or teams message styles
  --output OUTPUT       Save a new HTML preview file (never overwrites)
```

## m365 teams message styles

```text
usage: m365 teams message styles [-h]

options:
  -h, --help  show this help message and exit
```

## m365 teams message list

```text
usage: m365 teams message list [-h] (--chat CHAT | --channel CHANNEL)
                               [--limit LIMIT] [--cursor CURSOR]
                               [--reply-to REPLY_TO]

options:
  -h, --help           show this help message and exit
  --chat CHAT          Chat ID
  --channel CHANNEL    Standard channel ID
  --limit LIMIT
  --cursor CURSOR      Opaque next_cursor from the previous page
  --reply-to REPLY_TO  Channel root message ID to list replies
```

## m365 teams message get

```text
usage: m365 teams message get [-h] (--chat CHAT | --channel CHANNEL) id

positional arguments:
  id                 Message ID

options:
  -h, --help         show this help message and exit
  --chat CHAT        Chat ID
  --channel CHANNEL  Standard channel ID
```

## m365 teams message send

```text
usage: m365 teams message send [-h] (--chat CHAT | --channel CHANNEL)
                               [--body BODY | --body-file BODY_FILE]
                               [--body-format {text,html,markdown}]
                               [--style {outlook,minimal}]
                               [--style-file STYLE_FILE]
                               [--request-id REQUEST_ID] [--dry-run]

options:
  -h, --help            show this help message and exit
  --chat CHAT           Chat ID
  --channel CHANNEL     Standard channel ID
  --body BODY
  --body-file BODY_FILE
                        UTF-8 text file, or - for stdin
  --body-format {text,html,markdown}
                        Input format; Markdown is sent as styled HTML (default
                        text)
  --style {outlook,minimal}
                        HTML/Markdown styling: outlook (default; Helvetica
                        Neue/Arial) or minimal (no added styles)
  --style-file STYLE_FILE
                        JSON overrides for the outlook style; see mail styles,
                        calendar styles or teams message styles
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
```

## m365 teams message reply

```text
usage: m365 teams message reply [-h] (--chat CHAT | --channel CHANNEL)
                                --reply-to REPLY_TO [--body BODY |
                                --body-file BODY_FILE]
                                [--body-format {text,html,markdown}]
                                [--style {outlook,minimal}]
                                [--style-file STYLE_FILE]
                                [--request-id REQUEST_ID] [--dry-run]

options:
  -h, --help            show this help message and exit
  --chat CHAT           Chat ID
  --channel CHANNEL     Standard channel ID
  --reply-to REPLY_TO   Channel root message ID
  --body BODY
  --body-file BODY_FILE
                        UTF-8 text file, or - for stdin
  --body-format {text,html,markdown}
                        Input format; Markdown is sent as styled HTML (default
                        text)
  --style {outlook,minimal}
                        HTML/Markdown styling: outlook (default; Helvetica
                        Neue/Arial) or minimal (no added styles)
  --style-file STYLE_FILE
                        JSON overrides for the outlook style; see mail styles,
                        calendar styles or teams message styles
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
```

## m365 teams message update

```text
usage: m365 teams message update [-h] (--chat CHAT | --channel CHANNEL)
                                 [--body BODY | --body-file BODY_FILE]
                                 [--body-format {text,html,markdown}]
                                 [--style {outlook,minimal}]
                                 [--style-file STYLE_FILE]
                                 [--request-id REQUEST_ID] [--dry-run]
                                 [--if-version IF_VERSION]
                                 id

positional arguments:
  id                    Message ID

options:
  -h, --help            show this help message and exit
  --chat CHAT           Chat ID
  --channel CHANNEL     Standard channel ID
  --body BODY
  --body-file BODY_FILE
                        UTF-8 text file, or - for stdin
  --body-format {text,html,markdown}
                        Input format; Markdown is sent as styled HTML (default
                        text)
  --style {outlook,minimal}
                        HTML/Markdown styling: outlook (default; Helvetica
                        Neue/Arial) or minimal (no added styles)
  --style-file STYLE_FILE
                        JSON overrides for the outlook style; see mail styles,
                        calendar styles or teams message styles
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed edit version; rejected for deletion because
                        Teams ignores it
```

## m365 teams message delete

```text
usage: m365 teams message delete [-h] (--chat CHAT | --channel CHANNEL)
                                 [--request-id REQUEST_ID] [--dry-run]
                                 [--if-version IF_VERSION]
                                 id

positional arguments:
  id                    Message ID

options:
  -h, --help            show this help message and exit
  --chat CHAT           Chat ID
  --channel CHANNEL     Standard channel ID
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed edit version; rejected for deletion because
                        Teams ignores it
```

## m365 category

```text
usage: m365 category [-h]
                     {colors,list,get,create,update,delete,apply,remove} ...

positional arguments:
  {colors,list,get,create,update,delete,apply,remove}

options:
  -h, --help            show this help message and exit
```

## m365 category colors

```text
usage: m365 category colors [-h]

options:
  -h, --help  show this help message and exit
```

## m365 category list

```text
usage: m365 category list [-h] [--query QUERY]

options:
  -h, --help     show this help message and exit
  --query QUERY  Case-insensitive name substring
```

## m365 category get

```text
usage: m365 category get [-h] id

positional arguments:
  id          Category provider ID from category list/get

options:
  -h, --help  show this help message and exit
```

## m365 category create

```text
usage: m365 category create [-h] [--request-id REQUEST_ID] [--dry-run]
                            --name NAME [--color COLOR]

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --name NAME
  --color COLOR         Name or preset0–preset24; see category colors
```

## m365 category update

```text
usage: m365 category update [-h] [--request-id REQUEST_ID] [--dry-run]
                            [--name NAME] [--color COLOR] [--definition-only]
                            [--if-version IF_VERSION]
                            id

positional arguments:
  id                    Category provider ID from category list/get

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --name NAME
  --color COLOR         Name or preset0–preset24; see category colors
  --definition-only     Leave existing item category names unchanged when
                        renaming
  --if-version IF_VERSION
                        Reviewed item version for apply/remove; category
                        review hash for definition changes (not atomic)
```

## m365 category delete

```text
usage: m365 category delete [-h] [--request-id REQUEST_ID] [--dry-run]
                            [--if-version IF_VERSION]
                            id

positional arguments:
  id                    Category provider ID from category list/get

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed item version for apply/remove; category
                        review hash for definition changes (not atomic)
```

## m365 category apply

```text
usage: m365 category apply [-h] [--request-id REQUEST_ID] [--dry-run]
                           [--if-version IF_VERSION] (--task TASK_ID |
                           --mail MAIL_ID | --event EVENT_ID) [--list LIST_ID]
                           id

positional arguments:
  id                    Category provider ID from category list/get

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed item version for apply/remove; category
                        review hash for definition changes (not atomic)
  --task TASK_ID        To Do task ID; also requires --list
  --mail MAIL_ID
  --event EVENT_ID      Calendar occurrence ID; personal metadata, no
                        notifications
  --list LIST_ID
```

## m365 category remove

```text
usage: m365 category remove [-h] [--request-id REQUEST_ID] [--dry-run]
                            [--name NAME] [--if-version IF_VERSION]
                            (--task TASK_ID | --mail MAIL_ID |
                            --event EVENT_ID) [--list LIST_ID]
                            [id]

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --name NAME           Remove a name even when its definition was deleted
  --if-version IF_VERSION
                        Reviewed item version for apply/remove; category
                        review hash for definition changes (not atomic)
  --task TASK_ID        To Do task ID; also requires --list
  --mail MAIL_ID
  --event EVENT_ID      Calendar occurrence ID; personal metadata, no
                        notifications
  --list LIST_ID
```

## m365 tasks

```text
usage: m365 tasks [-h] {login,status,capabilities,list} ...

positional arguments:
  {login,status,capabilities,list}
    login               Reuse the running saved browser; never restart it
    capabilities        Implemented task commands and known validation gates
                        (local)
    list                To Do lists plus assigned Planner tasks; explicit
                        coverage

options:
  -h, --help            show this help message and exit
```

## m365 tasks login

```text
usage: m365 tasks login [-h] [--timeout TIMEOUT]

options:
  -h, --help         show this help message and exit
  --timeout TIMEOUT
```

## m365 tasks status

```text
usage: m365 tasks status [-h]

options:
  -h, --help  show this help message and exit
```

## m365 tasks capabilities

```text
usage: m365 tasks capabilities [-h]

options:
  -h, --help  show this help message and exit
```

## m365 tasks list

```text
usage: m365 tasks list [-h] [--source {todo,planner,all}]
                       [--max-lists MAX_LISTS] [--limit LIMIT]
                       [--offset OFFSET] [--max-pages MAX_PAGES]
                       [--query QUERY]

options:
  -h, --help            show this help message and exit
  --source {todo,planner,all}
  --max-lists MAX_LISTS
  --limit LIMIT
  --offset OFFSET
  --max-pages MAX_PAGES
                        Bound the server scan (1–100); truncation is explicit
  --query QUERY         Case-insensitive title substring in this bounded scan;
                        no server $search
```

## m365 todo

```text
usage: m365 todo [-h]
                 {category,list,task,step,link,attachment,my-day,group,delta} ...

positional arguments:
  {category,list,task,step,link,attachment,my-day,group,delta}

options:
  -h, --help            show this help message and exit
```

## m365 todo category

```text
usage: m365 todo category [-h]
                          {colors,list,get,create,update,delete,apply,remove} ...

positional arguments:
  {colors,list,get,create,update,delete,apply,remove}

options:
  -h, --help            show this help message and exit
```

## m365 todo category colors

```text
usage: m365 todo category colors [-h]

options:
  -h, --help  show this help message and exit
```

## m365 todo category list

```text
usage: m365 todo category list [-h] [--query QUERY]

options:
  -h, --help     show this help message and exit
  --query QUERY  Case-insensitive name substring
```

## m365 todo category get

```text
usage: m365 todo category get [-h] id

positional arguments:
  id          Category provider ID from category list/get

options:
  -h, --help  show this help message and exit
```

## m365 todo category create

```text
usage: m365 todo category create [-h] [--request-id REQUEST_ID] [--dry-run]
                                 --name NAME [--color COLOR]

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --name NAME
  --color COLOR         Name or preset0–preset24; see category colors
```

## m365 todo category update

```text
usage: m365 todo category update [-h] [--request-id REQUEST_ID] [--dry-run]
                                 [--name NAME] [--color COLOR]
                                 [--definition-only] [--if-version IF_VERSION]
                                 id

positional arguments:
  id                    Category provider ID from category list/get

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --name NAME
  --color COLOR         Name or preset0–preset24; see category colors
  --definition-only     Leave existing item category names unchanged when
                        renaming
  --if-version IF_VERSION
                        Reviewed item version for apply/remove; category
                        review hash for definition changes (not atomic)
```

## m365 todo category delete

```text
usage: m365 todo category delete [-h] [--request-id REQUEST_ID] [--dry-run]
                                 [--if-version IF_VERSION]
                                 id

positional arguments:
  id                    Category provider ID from category list/get

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed item version for apply/remove; category
                        review hash for definition changes (not atomic)
```

## m365 todo category apply

```text
usage: m365 todo category apply [-h] [--request-id REQUEST_ID] [--dry-run]
                                [--if-version IF_VERSION] (--task TASK_ID |
                                --mail MAIL_ID | --event EVENT_ID)
                                [--list LIST_ID]
                                id

positional arguments:
  id                    Category provider ID from category list/get

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed item version for apply/remove; category
                        review hash for definition changes (not atomic)
  --task TASK_ID        To Do task ID; also requires --list
  --mail MAIL_ID
  --event EVENT_ID      Calendar occurrence ID; personal metadata, no
                        notifications
  --list LIST_ID
```

## m365 todo category remove

```text
usage: m365 todo category remove [-h] [--request-id REQUEST_ID] [--dry-run]
                                 [--name NAME] [--if-version IF_VERSION]
                                 (--task TASK_ID | --mail MAIL_ID |
                                 --event EVENT_ID) [--list LIST_ID]
                                 [id]

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --name NAME           Remove a name even when its definition was deleted
  --if-version IF_VERSION
                        Reviewed item version for apply/remove; category
                        review hash for definition changes (not atomic)
  --task TASK_ID        To Do task ID; also requires --list
  --mail MAIL_ID
  --event EVENT_ID      Calendar occurrence ID; personal metadata, no
                        notifications
  --list LIST_ID
```

## m365 todo list

```text
usage: m365 todo list [-h] {list,get,create,update,delete,group} ...

positional arguments:
  {list,get,create,update,delete,group}

options:
  -h, --help            show this help message and exit
```

## m365 todo list list

```text
usage: m365 todo list list [-h] [--limit LIMIT] [--offset OFFSET]
                           [--max-pages MAX_PAGES]

options:
  -h, --help            show this help message and exit
  --limit LIMIT
  --offset OFFSET
  --max-pages MAX_PAGES
                        Bound the server scan (1–100); truncation is explicit
```

## m365 todo list get

```text
usage: m365 todo list get [-h] id

positional arguments:
  id

options:
  -h, --help  show this help message and exit
```

## m365 todo list create

```text
usage: m365 todo list create [-h] [--request-id REQUEST_ID] [--dry-run]
                             [--data DATA | --data-file DATA_FILE]
                             [--name NAME]

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --data DATA           JSON object of editable fields; see docs/TASKS.md
  --data-file DATA_FILE
                        UTF-8 JSON object file, or - for stdin
  --name NAME
```

## m365 todo list update

```text
usage: m365 todo list update [-h] [--request-id REQUEST_ID] [--dry-run]
                             [--if-version IF_VERSION] [--data DATA |
                             --data-file DATA_FILE] [--name NAME]
                             id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed resource version; rejected when no atomic
                        precondition is validated
  --data DATA           JSON object of editable fields; see docs/TASKS.md
  --data-file DATA_FILE
                        UTF-8 JSON object file, or - for stdin
  --name NAME
```

## m365 todo list delete

```text
usage: m365 todo list delete [-h] [--request-id REQUEST_ID] [--dry-run]
                             [--if-version IF_VERSION]
                             id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed resource version; rejected when no atomic
                        precondition is validated
```

## m365 todo list group

```text
usage: m365 todo list group [-h] [--request-id REQUEST_ID] [--dry-run]
                            --group-id GROUP_ID
                            id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --group-id GROUP_ID
```

## m365 todo task

```text
usage: m365 todo task [-h]
                      {list,get,create,update,complete,reopen,delete,assign,unassign,move} ...

positional arguments:
  {list,get,create,update,complete,reopen,delete,assign,unassign,move}

options:
  -h, --help            show this help message and exit
```

## m365 todo task list

```text
usage: m365 todo task list [-h] --list LIST_ID [--limit LIMIT]
                           [--offset OFFSET] [--max-pages MAX_PAGES]
                           [--query QUERY]

options:
  -h, --help            show this help message and exit
  --list LIST_ID
  --limit LIMIT
  --offset OFFSET
  --max-pages MAX_PAGES
                        Bound the server scan (1–100); truncation is explicit
  --query QUERY         Case-insensitive title substring in this bounded scan;
                        no server $search
```

## m365 todo task get

```text
usage: m365 todo task get [-h] --list LIST_ID id

positional arguments:
  id

options:
  -h, --help      show this help message and exit
  --list LIST_ID
```

## m365 todo task create

```text
usage: m365 todo task create [-h] [--request-id REQUEST_ID] [--dry-run]
                             --list LIST_ID [--data DATA |
                             --data-file DATA_FILE] [--title TITLE]
                             [--status {notStarted,inProgress,completed,waitingOnOthers,deferred}]
                             [--body BODY | --body-file BODY_FILE]

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --list LIST_ID
  --data DATA           JSON object of editable fields; see docs/TASKS.md
  --data-file DATA_FILE
                        UTF-8 JSON object file, or - for stdin
  --title TITLE
  --status {notStarted,inProgress,completed,waitingOnOthers,deferred}
  --body BODY
  --body-file BODY_FILE
                        UTF-8 text file, or - for stdin
```

## m365 todo task update

```text
usage: m365 todo task update [-h] [--request-id REQUEST_ID] [--dry-run]
                             [--if-version IF_VERSION] --list LIST_ID
                             [--data DATA | --data-file DATA_FILE]
                             [--title TITLE]
                             [--status {notStarted,inProgress,completed,waitingOnOthers,deferred}]
                             [--body BODY | --body-file BODY_FILE]
                             id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed resource version; rejected when no atomic
                        precondition is validated
  --list LIST_ID
  --data DATA           JSON object of editable fields; see docs/TASKS.md
  --data-file DATA_FILE
                        UTF-8 JSON object file, or - for stdin
  --title TITLE
  --status {notStarted,inProgress,completed,waitingOnOthers,deferred}
  --body BODY
  --body-file BODY_FILE
                        UTF-8 text file, or - for stdin
```

## m365 todo task complete

```text
usage: m365 todo task complete [-h] [--request-id REQUEST_ID] [--dry-run]
                               [--if-version IF_VERSION] --list LIST_ID
                               id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed resource version; rejected when no atomic
                        precondition is validated
  --list LIST_ID
```

## m365 todo task reopen

```text
usage: m365 todo task reopen [-h] [--request-id REQUEST_ID] [--dry-run]
                             [--if-version IF_VERSION] --list LIST_ID
                             id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed resource version; rejected when no atomic
                        precondition is validated
  --list LIST_ID
```

## m365 todo task delete

```text
usage: m365 todo task delete [-h] [--request-id REQUEST_ID] [--dry-run]
                             [--if-version IF_VERSION] --list LIST_ID
                             id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed resource version; rejected when no atomic
                        precondition is validated
  --list LIST_ID
```

## m365 todo task assign

```text
usage: m365 todo task assign [-h] [--request-id REQUEST_ID] [--dry-run]
                             [--if-version IF_VERSION] --list LIST_ID
                             id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed resource version; rejected when no atomic
                        precondition is validated
  --list LIST_ID
```

## m365 todo task unassign

```text
usage: m365 todo task unassign [-h] [--request-id REQUEST_ID] [--dry-run]
                               [--if-version IF_VERSION] --list LIST_ID
                               --assignment ASSIGNMENT
                               id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed resource version; rejected when no atomic
                        precondition is validated
  --list LIST_ID
  --assignment ASSIGNMENT
```

## m365 todo task move

```text
usage: m365 todo task move [-h] [--request-id REQUEST_ID] [--dry-run]
                           --list LIST_ID --to-list TO_LIST
                           id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --list LIST_ID
  --to-list TO_LIST
```

## m365 todo step

```text
usage: m365 todo step [-h] {list,get,create,update,check,uncheck,delete} ...

positional arguments:
  {list,get,create,update,check,uncheck,delete}

options:
  -h, --help            show this help message and exit
```

## m365 todo step list

```text
usage: m365 todo step list [-h] --list LIST_ID --task TASK_ID [--limit LIMIT]
                           [--offset OFFSET] [--max-pages MAX_PAGES]

options:
  -h, --help            show this help message and exit
  --list LIST_ID
  --task TASK_ID
  --limit LIMIT
  --offset OFFSET
  --max-pages MAX_PAGES
                        Bound the server scan (1–100); truncation is explicit
```

## m365 todo step get

```text
usage: m365 todo step get [-h] --list LIST_ID --task TASK_ID id

positional arguments:
  id

options:
  -h, --help      show this help message and exit
  --list LIST_ID
  --task TASK_ID
```

## m365 todo step create

```text
usage: m365 todo step create [-h] [--request-id REQUEST_ID] [--dry-run]
                             --list LIST_ID --task TASK_ID [--data DATA |
                             --data-file DATA_FILE] [--name NAME]

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --list LIST_ID
  --task TASK_ID
  --data DATA           JSON object of editable fields; see docs/TASKS.md
  --data-file DATA_FILE
                        UTF-8 JSON object file, or - for stdin
  --name NAME
```

## m365 todo step update

```text
usage: m365 todo step update [-h] [--request-id REQUEST_ID] [--dry-run]
                             [--if-version IF_VERSION] --list LIST_ID
                             --task TASK_ID [--data DATA |
                             --data-file DATA_FILE] [--name NAME]
                             id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed resource version; rejected when no atomic
                        precondition is validated
  --list LIST_ID
  --task TASK_ID
  --data DATA           JSON object of editable fields; see docs/TASKS.md
  --data-file DATA_FILE
                        UTF-8 JSON object file, or - for stdin
  --name NAME
```

## m365 todo step check

```text
usage: m365 todo step check [-h] [--request-id REQUEST_ID] [--dry-run]
                            --list LIST_ID --task TASK_ID
                            id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --list LIST_ID
  --task TASK_ID
```

## m365 todo step uncheck

```text
usage: m365 todo step uncheck [-h] [--request-id REQUEST_ID] [--dry-run]
                              --list LIST_ID --task TASK_ID
                              id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --list LIST_ID
  --task TASK_ID
```

## m365 todo step delete

```text
usage: m365 todo step delete [-h] [--request-id REQUEST_ID] [--dry-run]
                             [--if-version IF_VERSION] --list LIST_ID
                             --task TASK_ID
                             id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed resource version; rejected when no atomic
                        precondition is validated
  --list LIST_ID
  --task TASK_ID
```

## m365 todo link

```text
usage: m365 todo link [-h] {list,get,create,update,delete} ...

positional arguments:
  {list,get,create,update,delete}

options:
  -h, --help            show this help message and exit
```

## m365 todo link list

```text
usage: m365 todo link list [-h] --list LIST_ID --task TASK_ID [--limit LIMIT]
                           [--offset OFFSET] [--max-pages MAX_PAGES]

options:
  -h, --help            show this help message and exit
  --list LIST_ID
  --task TASK_ID
  --limit LIMIT
  --offset OFFSET
  --max-pages MAX_PAGES
                        Bound the server scan (1–100); truncation is explicit
```

## m365 todo link get

```text
usage: m365 todo link get [-h] --list LIST_ID --task TASK_ID id

positional arguments:
  id

options:
  -h, --help      show this help message and exit
  --list LIST_ID
  --task TASK_ID
```

## m365 todo link create

```text
usage: m365 todo link create [-h] [--request-id REQUEST_ID] [--dry-run]
                             --list LIST_ID --task TASK_ID [--data DATA |
                             --data-file DATA_FILE]

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --list LIST_ID
  --task TASK_ID
  --data DATA           JSON object of editable fields; see docs/TASKS.md
  --data-file DATA_FILE
                        UTF-8 JSON object file, or - for stdin
```

## m365 todo link update

```text
usage: m365 todo link update [-h] [--request-id REQUEST_ID] [--dry-run]
                             [--if-version IF_VERSION] --list LIST_ID
                             --task TASK_ID [--data DATA |
                             --data-file DATA_FILE]
                             id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed resource version; rejected when no atomic
                        precondition is validated
  --list LIST_ID
  --task TASK_ID
  --data DATA           JSON object of editable fields; see docs/TASKS.md
  --data-file DATA_FILE
                        UTF-8 JSON object file, or - for stdin
```

## m365 todo link delete

```text
usage: m365 todo link delete [-h] [--request-id REQUEST_ID] [--dry-run]
                             [--if-version IF_VERSION] --list LIST_ID
                             --task TASK_ID
                             id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed resource version; rejected when no atomic
                        precondition is validated
  --list LIST_ID
  --task TASK_ID
```

## m365 todo attachment

```text
usage: m365 todo attachment [-h] {list,get,add,delete} ...

positional arguments:
  {list,get,add,delete}

options:
  -h, --help            show this help message and exit
```

## m365 todo attachment list

```text
usage: m365 todo attachment list [-h] --list LIST_ID --task TASK_ID
                                 [--limit LIMIT] [--offset OFFSET]
                                 [--max-pages MAX_PAGES]

options:
  -h, --help            show this help message and exit
  --list LIST_ID
  --task TASK_ID
  --limit LIMIT
  --offset OFFSET
  --max-pages MAX_PAGES
                        Bound the server scan (1–100); truncation is explicit
```

## m365 todo attachment get

```text
usage: m365 todo attachment get [-h] --list LIST_ID --task TASK_ID
                                --output OUTPUT
                                id

positional arguments:
  id

options:
  -h, --help       show this help message and exit
  --list LIST_ID
  --task TASK_ID
  --output OUTPUT  New local path; never overwrite
```

## m365 todo attachment add

```text
usage: m365 todo attachment add [-h] [--request-id REQUEST_ID] [--dry-run]
                                --list LIST_ID --task TASK_ID --file FILE

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --list LIST_ID
  --task TASK_ID
  --file FILE           Single upload under 3 MB
```

## m365 todo attachment delete

```text
usage: m365 todo attachment delete [-h] [--request-id REQUEST_ID] [--dry-run]
                                   [--if-version IF_VERSION] --list LIST_ID
                                   --task TASK_ID
                                   id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed resource version; rejected when no atomic
                        precondition is validated
  --list LIST_ID
  --task TASK_ID
```

## m365 todo my-day

```text
usage: m365 todo my-day [-h] {list,add,remove} ...

positional arguments:
  {list,add,remove}

options:
  -h, --help         show this help message and exit
```

## m365 todo my-day list

```text
usage: m365 todo my-day list [-h] [--limit LIMIT] [--offset OFFSET]
                             [--max-pages MAX_PAGES]

options:
  -h, --help            show this help message and exit
  --limit LIMIT
  --offset OFFSET
  --max-pages MAX_PAGES
                        Bound the server scan (1–100); truncation is explicit
```

## m365 todo my-day add

```text
usage: m365 todo my-day add [-h] [--request-id REQUEST_ID] [--dry-run]
                            --list LIST_ID --task TASK_ID --date DATE

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --list LIST_ID
  --task TASK_ID
  --date DATE           Explicit local day, YYYY-MM-DD
```

## m365 todo my-day remove

```text
usage: m365 todo my-day remove [-h] [--request-id REQUEST_ID] [--dry-run]
                               --list LIST_ID --task TASK_ID

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --list LIST_ID
  --task TASK_ID
```

## m365 todo group

```text
usage: m365 todo group [-h] {list,get,create,update,delete} ...

positional arguments:
  {list,get,create,update,delete}

options:
  -h, --help            show this help message and exit
```

## m365 todo group list

```text
usage: m365 todo group list [-h] [--limit LIMIT] [--offset OFFSET]
                            [--max-pages MAX_PAGES]

options:
  -h, --help            show this help message and exit
  --limit LIMIT
  --offset OFFSET
  --max-pages MAX_PAGES
                        Bound the server scan (1–100); truncation is explicit
```

## m365 todo group get

```text
usage: m365 todo group get [-h] id

positional arguments:
  id

options:
  -h, --help  show this help message and exit
```

## m365 todo group create

```text
usage: m365 todo group create [-h] [--request-id REQUEST_ID] [--dry-run]
                              [--data DATA | --data-file DATA_FILE]
                              [--name NAME]

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --data DATA           JSON object of editable fields; see docs/TASKS.md
  --data-file DATA_FILE
                        UTF-8 JSON object file, or - for stdin
  --name NAME
```

## m365 todo group update

```text
usage: m365 todo group update [-h] [--request-id REQUEST_ID] [--dry-run]
                              [--if-version IF_VERSION] [--data DATA |
                              --data-file DATA_FILE] [--name NAME]
                              id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed resource version; rejected when no atomic
                        precondition is validated
  --data DATA           JSON object of editable fields; see docs/TASKS.md
  --data-file DATA_FILE
                        UTF-8 JSON object file, or - for stdin
  --name NAME
```

## m365 todo group delete

```text
usage: m365 todo group delete [-h] [--request-id REQUEST_ID] [--dry-run]
                              [--if-version IF_VERSION]
                              id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed resource version; rejected when no atomic
                        precondition is validated
```

## m365 todo delta

```text
usage: m365 todo delta [-h] {list} ...

positional arguments:
  {list}

options:
  -h, --help  show this help message and exit
```

## m365 todo delta list

```text
usage: m365 todo delta list [-h] --list LIST_ID [--cursor CURSOR]

options:
  -h, --help       show this help message and exit
  --list LIST_ID
  --cursor CURSOR  Opaque next_cursor or delta_cursor from this exact
                   collection
```

## m365 planner

```text
usage: m365 planner [-h]
                    {label,plan,bucket,task,details,labels,my-day,history,comment,delta,deletion} ...

positional arguments:
  {label,plan,bucket,task,details,labels,my-day,history,comment,delta,deletion}

options:
  -h, --help            show this help message and exit
```

## m365 planner label

```text
usage: m365 planner label [-h]
                          {colors,list,get,create,rename,recolor,delete,apply,remove} ...

positional arguments:
  {colors,list,get,create,rename,recolor,delete,apply,remove}

options:
  -h, --help            show this help message and exit
```

## m365 planner label colors

```text
usage: m365 planner label colors [-h]

options:
  -h, --help  show this help message and exit
```

## m365 planner label list

```text
usage: m365 planner label list [-h] --plan PLAN_ID

options:
  -h, --help      show this help message and exit
  --plan PLAN_ID
```

## m365 planner label get

```text
usage: m365 planner label get [-h] --plan PLAN_ID id

positional arguments:
  id              Fixed color slot: category1–category25

options:
  -h, --help      show this help message and exit
  --plan PLAN_ID
```

## m365 planner label create

```text
usage: m365 planner label create [-h] --plan PLAN_ID --name NAME --color COLOR
                                 [--request-id REQUEST_ID] [--dry-run]
                                 [--max-pages MAX_PAGES]

options:
  -h, --help            show this help message and exit
  --plan PLAN_ID
  --name NAME
  --color COLOR         Color name from planner label colors
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --max-pages MAX_PAGES
                        Complete task scan budget (1–100); incomplete scans
                        refuse writes
```

## m365 planner label rename

```text
usage: m365 planner label rename [-h] --plan PLAN_ID --name NAME
                                 [--request-id REQUEST_ID] [--dry-run]
                                 [--if-version IF_VERSION]
                                 id

positional arguments:
  id                    Fixed color slot: category1–category25

options:
  -h, --help            show this help message and exit
  --plan PLAN_ID
  --name NAME
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed plan-details ETag; task ETag for apply/remove
```

## m365 planner label recolor

```text
usage: m365 planner label recolor [-h] --plan PLAN_ID --color COLOR
                                  [--request-id REQUEST_ID] [--dry-run]
                                  [--if-version IF_VERSION]
                                  [--max-pages MAX_PAGES]
                                  id

positional arguments:
  id                    Fixed color slot: category1–category25

options:
  -h, --help            show this help message and exit
  --plan PLAN_ID
  --color COLOR         Color name from planner label colors
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed plan-details ETag; task ETag for apply/remove
  --max-pages MAX_PAGES
                        Complete task scan budget (1–100); incomplete scans
                        refuse writes
```

## m365 planner label delete

```text
usage: m365 planner label delete [-h] --plan PLAN_ID [--request-id REQUEST_ID]
                                 [--dry-run] [--if-version IF_VERSION]
                                 [--max-pages MAX_PAGES] [--remove-from-tasks]
                                 id

positional arguments:
  id                    Fixed color slot: category1–category25

options:
  -h, --help            show this help message and exit
  --plan PLAN_ID
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed plan-details ETag; task ETag for apply/remove
  --max-pages MAX_PAGES
                        Complete task scan budget (1–100); incomplete scans
                        refuse writes
  --remove-from-tasks   Remove this slot from all tasks in the plan before
                        clearing its name
```

## m365 planner label apply

```text
usage: m365 planner label apply [-h] --plan PLAN_ID [--request-id REQUEST_ID]
                                [--dry-run] [--if-version IF_VERSION]
                                --task TASK_ID
                                id

positional arguments:
  id                    Fixed color slot: category1–category25

options:
  -h, --help            show this help message and exit
  --plan PLAN_ID
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed plan-details ETag; task ETag for apply/remove
  --task TASK_ID
```

## m365 planner label remove

```text
usage: m365 planner label remove [-h] --plan PLAN_ID [--request-id REQUEST_ID]
                                 [--dry-run] [--if-version IF_VERSION]
                                 --task TASK_ID
                                 id

positional arguments:
  id                    Fixed color slot: category1–category25

options:
  -h, --help            show this help message and exit
  --plan PLAN_ID
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed plan-details ETag; task ETag for apply/remove
  --task TASK_ID
```

## m365 planner plan

```text
usage: m365 planner plan [-h]
                         {list,get,create,update,delete,archive,unarchive} ...

positional arguments:
  {list,get,create,update,delete,archive,unarchive}

options:
  -h, --help            show this help message and exit
```

## m365 planner plan list

```text
usage: m365 planner plan list [-h] [--limit LIMIT] [--offset OFFSET]
                              [--max-pages MAX_PAGES] [--group-id GROUP_ID]

options:
  -h, --help            show this help message and exit
  --limit LIMIT
  --offset OFFSET
  --max-pages MAX_PAGES
                        Bound the server scan (1–100); truncation is explicit
  --group-id GROUP_ID   Existing Microsoft 365 group; otherwise own user-
                        container plans
```

## m365 planner plan get

```text
usage: m365 planner plan get [-h] id

positional arguments:
  id

options:
  -h, --help  show this help message and exit
```

## m365 planner plan create

```text
usage: m365 planner plan create [-h] [--request-id REQUEST_ID] [--dry-run]
                                [--data DATA | --data-file DATA_FILE]
                                [--title TITLE] [--container {self,group}]
                                [--group-id GROUP_ID]

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --data DATA           JSON object of editable fields; see docs/TASKS.md
  --data-file DATA_FILE
                        UTF-8 JSON object file, or - for stdin
  --title TITLE
  --container {self,group}
                        Private self-owned plan by default; group requires an
                        existing group
  --group-id GROUP_ID
```

## m365 planner plan update

```text
usage: m365 planner plan update [-h] [--request-id REQUEST_ID] [--dry-run]
                                [--if-version IF_VERSION] [--data DATA |
                                --data-file DATA_FILE] [--title TITLE]
                                id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed resource version; rejected when no atomic
                        precondition is validated
  --data DATA           JSON object of editable fields; see docs/TASKS.md
  --data-file DATA_FILE
                        UTF-8 JSON object file, or - for stdin
  --title TITLE
```

## m365 planner plan delete

```text
usage: m365 planner plan delete [-h] [--request-id REQUEST_ID] [--dry-run]
                                [--if-version IF_VERSION]
                                id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed resource version; rejected when no atomic
                        precondition is validated
```

## m365 planner plan archive

```text
usage: m365 planner plan archive [-h] [--request-id REQUEST_ID] [--dry-run]
                                 [--if-version IF_VERSION] [--reason REASON]
                                 id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed resource version; rejected when no atomic
                        precondition is validated
  --reason REASON
```

## m365 planner plan unarchive

```text
usage: m365 planner plan unarchive [-h] [--request-id REQUEST_ID] [--dry-run]
                                   [--if-version IF_VERSION] [--reason REASON]
                                   id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed resource version; rejected when no atomic
                        precondition is validated
  --reason REASON
```

## m365 planner bucket

```text
usage: m365 planner bucket [-h] {list,get,create,update,delete} ...

positional arguments:
  {list,get,create,update,delete}

options:
  -h, --help            show this help message and exit
```

## m365 planner bucket list

```text
usage: m365 planner bucket list [-h] --plan PLAN_ID [--limit LIMIT]
                                [--offset OFFSET] [--max-pages MAX_PAGES]

options:
  -h, --help            show this help message and exit
  --plan PLAN_ID        Omit for task list to read assigned-to-me only
  --limit LIMIT
  --offset OFFSET
  --max-pages MAX_PAGES
                        Bound the server scan (1–100); truncation is explicit
```

## m365 planner bucket get

```text
usage: m365 planner bucket get [-h] id

positional arguments:
  id

options:
  -h, --help  show this help message and exit
```

## m365 planner bucket create

```text
usage: m365 planner bucket create [-h] [--request-id REQUEST_ID] [--dry-run]
                                  --plan PLAN_ID [--data DATA |
                                  --data-file DATA_FILE] [--name NAME]

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --plan PLAN_ID        Omit for task list to read assigned-to-me only
  --data DATA           JSON object of editable fields; see docs/TASKS.md
  --data-file DATA_FILE
                        UTF-8 JSON object file, or - for stdin
  --name NAME
```

## m365 planner bucket update

```text
usage: m365 planner bucket update [-h] [--request-id REQUEST_ID] [--dry-run]
                                  [--if-version IF_VERSION] [--data DATA |
                                  --data-file DATA_FILE] [--name NAME]
                                  id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed resource version; rejected when no atomic
                        precondition is validated
  --data DATA           JSON object of editable fields; see docs/TASKS.md
  --data-file DATA_FILE
                        UTF-8 JSON object file, or - for stdin
  --name NAME
```

## m365 planner bucket delete

```text
usage: m365 planner bucket delete [-h] [--request-id REQUEST_ID] [--dry-run]
                                  [--if-version IF_VERSION]
                                  id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed resource version; rejected when no atomic
                        precondition is validated
```

## m365 planner task

```text
usage: m365 planner task [-h]
                         {list,get,create,update,complete,reopen,delete,assign,unassign,move} ...

positional arguments:
  {list,get,create,update,complete,reopen,delete,assign,unassign,move}

options:
  -h, --help            show this help message and exit
```

## m365 planner task list

```text
usage: m365 planner task list [-h] [--plan PLAN_ID] [--limit LIMIT]
                              [--offset OFFSET] [--max-pages MAX_PAGES]
                              [--query QUERY]

options:
  -h, --help            show this help message and exit
  --plan PLAN_ID        Omit for task list to read assigned-to-me only
  --limit LIMIT
  --offset OFFSET
  --max-pages MAX_PAGES
                        Bound the server scan (1–100); truncation is explicit
  --query QUERY         Case-insensitive title substring in this bounded scan;
                        no server $search
```

## m365 planner task get

```text
usage: m365 planner task get [-h] id

positional arguments:
  id

options:
  -h, --help  show this help message and exit
```

## m365 planner task create

```text
usage: m365 planner task create [-h] [--request-id REQUEST_ID] [--dry-run]
                                --plan PLAN_ID [--data DATA |
                                --data-file DATA_FILE] [--title TITLE]
                                [--progress PROGRESS]

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --plan PLAN_ID        Omit for task list to read assigned-to-me only
  --data DATA           JSON object of editable fields; see docs/TASKS.md
  --data-file DATA_FILE
                        UTF-8 JSON object file, or - for stdin
  --title TITLE
  --progress PROGRESS   Completion percentage, 0–100
```

## m365 planner task update

```text
usage: m365 planner task update [-h] [--request-id REQUEST_ID] [--dry-run]
                                [--if-version IF_VERSION] [--data DATA |
                                --data-file DATA_FILE] [--title TITLE]
                                [--progress PROGRESS]
                                id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed resource version; rejected when no atomic
                        precondition is validated
  --data DATA           JSON object of editable fields; see docs/TASKS.md
  --data-file DATA_FILE
                        UTF-8 JSON object file, or - for stdin
  --title TITLE
  --progress PROGRESS   Completion percentage, 0–100
```

## m365 planner task complete

```text
usage: m365 planner task complete [-h] [--request-id REQUEST_ID] [--dry-run]
                                  [--if-version IF_VERSION]
                                  id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed resource version; rejected when no atomic
                        precondition is validated
```

## m365 planner task reopen

```text
usage: m365 planner task reopen [-h] [--request-id REQUEST_ID] [--dry-run]
                                [--if-version IF_VERSION]
                                id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed resource version; rejected when no atomic
                        precondition is validated
```

## m365 planner task delete

```text
usage: m365 planner task delete [-h] [--request-id REQUEST_ID] [--dry-run]
                                [--if-version IF_VERSION]
                                [--recurrence {occurrence,stop-series}]
                                [--allow-undo]
                                id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed resource version; rejected when no atomic
                        precondition is validated
  --recurrence {occurrence,stop-series}
                        Required when deleting an active recurring task
  --allow-undo          Use the native deletion adapter and return its
                        operation ID when available
```

## m365 planner task assign

```text
usage: m365 planner task assign [-h] [--request-id REQUEST_ID] [--dry-run]
                                [--if-version IF_VERSION] --user USER
                                id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed resource version; rejected when no atomic
                        precondition is validated
  --user USER           self or exact Entra user UUID; does not grant access
```

## m365 planner task unassign

```text
usage: m365 planner task unassign [-h] [--request-id REQUEST_ID] [--dry-run]
                                  [--if-version IF_VERSION] --user USER
                                  id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed resource version; rejected when no atomic
                        precondition is validated
  --user USER           self or exact Entra user UUID; does not grant access
```

## m365 planner task move

```text
usage: m365 planner task move [-h] [--request-id REQUEST_ID] [--dry-run]
                              --to-plan TO_PLAN [--to-bucket TO_BUCKET]
                              id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --to-plan TO_PLAN
  --to-bucket TO_BUCKET
```

## m365 planner details

```text
usage: m365 planner details [-h] {get,update} ...

positional arguments:
  {get,update}

options:
  -h, --help    show this help message and exit
```

## m365 planner details get

```text
usage: m365 planner details get [-h] --task TASK_ID

options:
  -h, --help      show this help message and exit
  --task TASK_ID
```

## m365 planner details update

```text
usage: m365 planner details update [-h] [--request-id REQUEST_ID] [--dry-run]
                                   [--if-version IF_VERSION] --task TASK_ID
                                   [--data DATA | --data-file DATA_FILE]
                                   [--body BODY | --body-file BODY_FILE]

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed resource version; rejected when no atomic
                        precondition is validated
  --task TASK_ID
  --data DATA           JSON object of editable fields; see docs/TASKS.md
  --data-file DATA_FILE
                        UTF-8 JSON object file, or - for stdin
  --body BODY
  --body-file BODY_FILE
                        UTF-8 text file, or - for stdin
```

## m365 planner labels

```text
usage: m365 planner labels [-h] {get,update} ...

positional arguments:
  {get,update}

options:
  -h, --help    show this help message and exit
```

## m365 planner labels get

```text
usage: m365 planner labels get [-h] --plan PLAN_ID

options:
  -h, --help      show this help message and exit
  --plan PLAN_ID  Omit for task list to read assigned-to-me only
```

## m365 planner labels update

```text
usage: m365 planner labels update [-h] [--request-id REQUEST_ID] [--dry-run]
                                  [--if-version IF_VERSION] --plan PLAN_ID
                                  [--data DATA | --data-file DATA_FILE]

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed resource version; rejected when no atomic
                        precondition is validated
  --plan PLAN_ID        Omit for task list to read assigned-to-me only
  --data DATA           JSON object of editable fields; see docs/TASKS.md
  --data-file DATA_FILE
                        UTF-8 JSON object file, or - for stdin
```

## m365 planner my-day

```text
usage: m365 planner my-day [-h] {list,add,remove} ...

positional arguments:
  {list,add,remove}

options:
  -h, --help         show this help message and exit
```

## m365 planner my-day list

```text
usage: m365 planner my-day list [-h] [--limit LIMIT] [--offset OFFSET]
                                [--max-pages MAX_PAGES]

options:
  -h, --help            show this help message and exit
  --limit LIMIT
  --offset OFFSET
  --max-pages MAX_PAGES
                        Bound the server scan (1–100); truncation is explicit
```

## m365 planner my-day add

```text
usage: m365 planner my-day add [-h] [--request-id REQUEST_ID] [--dry-run]
                               --task TASK_ID

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --task TASK_ID
```

## m365 planner my-day remove

```text
usage: m365 planner my-day remove [-h] [--request-id REQUEST_ID] [--dry-run]
                                  --task TASK_ID

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --task TASK_ID
```

## m365 planner history

```text
usage: m365 planner history [-h] {list} ...

positional arguments:
  {list}

options:
  -h, --help  show this help message and exit
```

## m365 planner history list

```text
usage: m365 planner history list [-h] --plan PLAN_ID [--limit LIMIT]
                                 [--offset OFFSET] [--max-pages MAX_PAGES]

options:
  -h, --help            show this help message and exit
  --plan PLAN_ID        Omit for task list to read assigned-to-me only
  --limit LIMIT
  --offset OFFSET
  --max-pages MAX_PAGES
                        Bound the server scan (1–100); truncation is explicit
```

## m365 planner comment

```text
usage: m365 planner comment [-h] {list,create,update,delete} ...

positional arguments:
  {list,create,update,delete}

options:
  -h, --help            show this help message and exit
```

## m365 planner comment list

```text
usage: m365 planner comment list [-h] --task TASK_ID

options:
  -h, --help      show this help message and exit
  --task TASK_ID
```

## m365 planner comment create

```text
usage: m365 planner comment create [-h] [--request-id REQUEST_ID] [--dry-run]
                                   --task TASK_ID [--body BODY |
                                   --body-file BODY_FILE]

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --task TASK_ID
  --body BODY
  --body-file BODY_FILE
                        UTF-8 text file, or - for stdin
```

## m365 planner comment update

```text
usage: m365 planner comment update [-h] [--request-id REQUEST_ID] [--dry-run]
                                   [--if-version IF_VERSION] --task TASK_ID
                                   [--body BODY | --body-file BODY_FILE]
                                   id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed resource version; rejected when no atomic
                        precondition is validated
  --task TASK_ID
  --body BODY
  --body-file BODY_FILE
                        UTF-8 text file, or - for stdin
```

## m365 planner comment delete

```text
usage: m365 planner comment delete [-h] [--request-id REQUEST_ID] [--dry-run]
                                   [--if-version IF_VERSION] --task TASK_ID
                                   id

positional arguments:
  id

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Reviewed resource version; rejected when no atomic
                        precondition is validated
  --task TASK_ID
```

## m365 planner delta

```text
usage: m365 planner delta [-h] {list} ...

positional arguments:
  {list}

options:
  -h, --help  show this help message and exit
```

## m365 planner delta list

```text
usage: m365 planner delta list [-h] [--cursor CURSOR]

options:
  -h, --help       show this help message and exit
  --cursor CURSOR  Opaque next_cursor or delta_cursor from this exact
                   collection
```

## m365 planner deletion

```text
usage: m365 planner deletion [-h] {undo} ...

positional arguments:
  {undo}

options:
  -h, --help  show this help message and exit
```

## m365 planner deletion undo

```text
usage: m365 planner deletion undo [-h] [--request-id REQUEST_ID] [--dry-run]
                                  --task TASK_ID --operation-id OPERATION_ID

options:
  -h, --help            show this help message and exit
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --task TASK_ID
  --operation-id OPERATION_ID
                        Acknowledged native deletion operation ID
```

## m365 files

```text
usage: m365 files [-h]
                  {login,status,capabilities,libraries,resolve,search,get,metadata,versions,download,fetch,rename,copy,move,recycle,restore,upload,folder,list} ...

positional arguments:
  {login,status,capabilities,libraries,resolve,search,get,metadata,versions,download,fetch,rename,copy,move,recycle,restore,upload,folder,list}
    login               Verify SharePoint sign-in in the saved browser
    status              Verify account, tenant and site access
    capabilities        Local file/list boundaries; no live availability claim
    libraries           List visible document libraries and their root folder
                        IDs
    resolve             Resolve an existing-access SharePoint
                        file/folder/list-item URL
    search              Site-scoped indexed literal phrase search; incomplete
                        by design
    get                 Read file metadata and stable URL
    metadata            Read the list-item ETag used by rename
    versions            Read historical versions; current version is separate
    download            Save bytes to a new local path (max 50 MiB)
    fetch               Download bytes plus a reference/version/hash manifest
    rename              Rename with an atomic metadata ETag guard
    copy                Copy a verified byte snapshot to a new name (max 10
                        MiB)
    move                Move within one library, without overwrite
    recycle             Move a file to the recycle bin
    restore             Restore historical content as a new current version
    upload              Create a new file; name collision fails (max 10 MiB)
    folder              Folder metadata and nonrecursive operations
    list                List files or subfolders in one folder

options:
  -h, --help            show this help message and exit
```

## m365 files login

```text
usage: m365 files login [-h] [--site SITE] [--timeout TIMEOUT]

options:
  -h, --help         show this help message and exit
  --site SITE        SharePoint site root; required unless a stable handle or
                     URL supplies it
  --timeout TIMEOUT
```

## m365 files status

```text
usage: m365 files status [-h] [--site SITE]

options:
  -h, --help   show this help message and exit
  --site SITE  SharePoint site root; required unless a stable handle or URL
               supplies it
```

## m365 files capabilities

```text
usage: m365 files capabilities [-h] [--site SITE]

options:
  -h, --help   show this help message and exit
  --site SITE  SharePoint site root; required unless a stable handle or URL
               supplies it
```

## m365 files libraries

```text
usage: m365 files libraries [-h] [--site SITE] [--limit LIMIT]
                            [--cursor CURSOR]

options:
  -h, --help       show this help message and exit
  --site SITE      SharePoint site root; required unless a stable handle or
                   URL supplies it
  --limit LIMIT    Page size 1–500; follow next_cursor
  --cursor CURSOR  Opaque cursor from this exact account/site/query/limit
```

## m365 files resolve

```text
usage: m365 files resolve [-h] [--site SITE] url

positional arguments:
  url

options:
  -h, --help   show this help message and exit
  --site SITE  SharePoint site root; required unless a stable handle or URL
               supplies it
```

## m365 files search

```text
usage: m365 files search [-h] [--site SITE] [--limit LIMIT] [--offset OFFSET]
                         query

positional arguments:
  query

options:
  -h, --help       show this help message and exit
  --site SITE      SharePoint site root; required unless a stable handle or
                   URL supplies it
  --limit LIMIT
  --offset OFFSET
```

## m365 files get

```text
usage: m365 files get [-h] [--site SITE] id

positional arguments:
  id           Native ID or m365: stable handle

options:
  -h, --help   show this help message and exit
  --site SITE  SharePoint site root; required unless a stable handle or URL
               supplies it
```

## m365 files metadata

```text
usage: m365 files metadata [-h] [--site SITE] id

positional arguments:
  id           Native ID or m365: stable handle

options:
  -h, --help   show this help message and exit
  --site SITE  SharePoint site root; required unless a stable handle or URL
               supplies it
```

## m365 files versions

```text
usage: m365 files versions [-h] [--site SITE] [--limit LIMIT]
                           [--cursor CURSOR]
                           id

positional arguments:
  id               Native ID or m365: stable handle

options:
  -h, --help       show this help message and exit
  --site SITE      SharePoint site root; required unless a stable handle or
                   URL supplies it
  --limit LIMIT    Page size 1–500; follow next_cursor
  --cursor CURSOR  Opaque cursor from this exact account/site/query/limit
```

## m365 files download

```text
usage: m365 files download [-h] [--site SITE] --output OUTPUT
                           [--version VERSION]
                           id

positional arguments:
  id                 Native ID or m365: stable handle

options:
  -h, --help         show this help message and exit
  --site SITE        SharePoint site root; required unless a stable handle or
                     URL supplies it
  --output OUTPUT    New local file path; never overwritten
  --version VERSION  Historical version label; omit for current bytes
```

## m365 files fetch

```text
usage: m365 files fetch [-h] [--site SITE] --output OUTPUT [--version VERSION]
                        id

positional arguments:
  id                 Native ID or m365: stable handle

options:
  -h, --help         show this help message and exit
  --site SITE        SharePoint site root; required unless a stable handle or
                     URL supplies it
  --output OUTPUT    New local file path; never overwritten
  --version VERSION  Historical version label; omit for current bytes
```

## m365 files rename

```text
usage: m365 files rename [-h] [--site SITE] [--request-id REQUEST_ID]
                         [--dry-run] [--if-version IF_VERSION] --name NAME
                         id

positional arguments:
  id                    Native ID or m365: stable handle

options:
  -h, --help            show this help message and exit
  --site SITE           SharePoint site root; required unless a stable handle
                        or URL supplies it
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Exact quoted ETag from the matching get/metadata
                        command
  --name NAME
```

## m365 files copy

```text
usage: m365 files copy [-h] [--site SITE] [--request-id REQUEST_ID]
                       [--dry-run] [--version VERSION] --name NAME
                       --folder FOLDER
                       id

positional arguments:
  id                    Native ID or m365: stable handle

options:
  -h, --help            show this help message and exit
  --site SITE           SharePoint site root; required unless a stable handle
                        or URL supplies it
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --version VERSION     Historical version label; omit for current bytes
  --name NAME
  --folder FOLDER       Destination folder GUID or stable handle in the same
                        site
```

## m365 files move

```text
usage: m365 files move [-h] [--site SITE] [--request-id REQUEST_ID]
                       [--dry-run] [--allow-unconditional] [--name NAME]
                       --folder FOLDER
                       id

positional arguments:
  id                    Native ID or m365: stable handle

options:
  -h, --help            show this help message and exit
  --site SITE           SharePoint site root; required unless a stable handle
                        or URL supplies it
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --allow-unconditional
                        Required: endpoint has no verified atomic concurrency
                        guard
  --name NAME
  --folder FOLDER       Destination folder GUID or stable handle in the same
                        site
```

## m365 files recycle

```text
usage: m365 files recycle [-h] [--site SITE] [--request-id REQUEST_ID]
                          [--dry-run] [--allow-unconditional]
                          id

positional arguments:
  id                    Native ID or m365: stable handle

options:
  -h, --help            show this help message and exit
  --site SITE           SharePoint site root; required unless a stable handle
                        or URL supplies it
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --allow-unconditional
                        Required: endpoint has no verified atomic concurrency
                        guard
```

## m365 files restore

```text
usage: m365 files restore [-h] [--site SITE] [--request-id REQUEST_ID]
                          [--dry-run] [--allow-unconditional]
                          --version VERSION
                          id

positional arguments:
  id                    Native ID or m365: stable handle

options:
  -h, --help            show this help message and exit
  --site SITE           SharePoint site root; required unless a stable handle
                        or URL supplies it
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --allow-unconditional
                        Required: endpoint has no verified atomic concurrency
                        guard
  --version VERSION     Historical version label; omit for current bytes
```

## m365 files upload

```text
usage: m365 files upload [-h] [--site SITE] [--request-id REQUEST_ID]
                         [--dry-run] --folder FOLDER --file FILE [--name NAME]

options:
  -h, --help            show this help message and exit
  --site SITE           SharePoint site root; required unless a stable handle
                        or URL supplies it
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --folder FOLDER
  --file FILE           Local file containing bytes
  --name NAME           Destination name; defaults to local filename
```

## m365 files folder

```text
usage: m365 files folder [-h] {get,metadata,list,create,rename,recycle} ...

positional arguments:
  {get,metadata,list,create,rename,recycle}
    get                 Get folder
    metadata            Metadata folder
    list                List folder
    create              Create folder
    rename              Rename folder
    recycle             Recycle folder

options:
  -h, --help            show this help message and exit
```

## m365 files folder get

```text
usage: m365 files folder get [-h] [--site SITE] id

positional arguments:
  id           Native ID or m365: stable handle

options:
  -h, --help   show this help message and exit
  --site SITE  SharePoint site root; required unless a stable handle or URL
               supplies it
```

## m365 files folder metadata

```text
usage: m365 files folder metadata [-h] [--site SITE] id

positional arguments:
  id           Native ID or m365: stable handle

options:
  -h, --help   show this help message and exit
  --site SITE  SharePoint site root; required unless a stable handle or URL
               supplies it
```

## m365 files folder list

```text
usage: m365 files folder list [-h] [--site SITE] [--limit LIMIT]
                              [--cursor CURSOR] [--kind {file,folder}]
                              id

positional arguments:
  id                    Native ID or m365: stable handle

options:
  -h, --help            show this help message and exit
  --site SITE           SharePoint site root; required unless a stable handle
                        or URL supplies it
  --limit LIMIT         Page size 1–500; follow next_cursor
  --cursor CURSOR       Opaque cursor from this exact account/site/query/limit
  --kind {file,folder}  Enumerate files or subfolders separately
```

## m365 files folder create

```text
usage: m365 files folder create [-h] [--site SITE] [--request-id REQUEST_ID]
                                [--dry-run] --folder FOLDER --name NAME

options:
  -h, --help            show this help message and exit
  --site SITE           SharePoint site root; required unless a stable handle
                        or URL supplies it
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --folder FOLDER       Parent folder GUID or stable handle
  --name NAME
```

## m365 files folder rename

```text
usage: m365 files folder rename [-h] [--site SITE] [--request-id REQUEST_ID]
                                [--dry-run] [--if-version IF_VERSION]
                                --name NAME
                                id

positional arguments:
  id                    Native ID or m365: stable handle

options:
  -h, --help            show this help message and exit
  --site SITE           SharePoint site root; required unless a stable handle
                        or URL supplies it
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Exact quoted ETag from the matching get/metadata
                        command
  --name NAME
```

## m365 files folder recycle

```text
usage: m365 files folder recycle [-h] [--site SITE] [--request-id REQUEST_ID]
                                 [--dry-run] [--allow-unconditional]
                                 id

positional arguments:
  id                    Native ID or m365: stable handle

options:
  -h, --help            show this help message and exit
  --site SITE           SharePoint site root; required unless a stable handle
                        or URL supplies it
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --allow-unconditional
                        Required: endpoint has no verified atomic concurrency
                        guard
```

## m365 files list

```text
usage: m365 files list [-h] [--site SITE] [--limit LIMIT] [--cursor CURSOR]
                       --folder FOLDER [--kind {file,folder}]

options:
  -h, --help            show this help message and exit
  --site SITE           SharePoint site root; required unless a stable handle
                        or URL supplies it
  --limit LIMIT         Page size 1–500; follow next_cursor
  --cursor CURSOR       Opaque cursor from this exact account/site/query/limit
  --folder FOLDER
  --kind {file,folder}
```

## m365 lists

```text
usage: m365 lists [-h]
                  {list,users,get,schema,create,rename,recycle,column,item,attachment,comments} ...

positional arguments:
  {list,users,get,schema,create,rename,recycle,column,item,attachment,comments}
    list                List visible lists and libraries
    users               Site-scoped person IDs for typed fields
    get                 Get list
    schema              Schema list
    create              Create list (visible generic lists only)
    rename              Rename list (visible generic lists only)
    recycle             Recycle list (visible generic lists only)
    column              Add a new supported column; no destructive schema
                        edits

options:
  -h, --help            show this help message and exit
```

## m365 lists list

```text
usage: m365 lists list [-h] [--site SITE] [--limit LIMIT] [--cursor CURSOR]

options:
  -h, --help       show this help message and exit
  --site SITE      SharePoint site root; required unless a stable handle or
                   URL supplies it
  --limit LIMIT    Page size 1–500; follow next_cursor
  --cursor CURSOR  Opaque cursor from this exact account/site/query/limit
```

## m365 lists users

```text
usage: m365 lists users [-h] [--site SITE] [--limit LIMIT] [--cursor CURSOR]
                        [--email EMAIL]

options:
  -h, --help       show this help message and exit
  --site SITE      SharePoint site root; required unless a stable handle or
                   URL supplies it
  --limit LIMIT    Page size 1–500; follow next_cursor
  --cursor CURSOR  Opaque cursor from this exact account/site/query/limit
  --email EMAIL    Exact email filter
```

## m365 lists get

```text
usage: m365 lists get [-h] [--site SITE] id

positional arguments:
  id           Native ID or m365: stable handle

options:
  -h, --help   show this help message and exit
  --site SITE  SharePoint site root; required unless a stable handle or URL
               supplies it
```

## m365 lists schema

```text
usage: m365 lists schema [-h] [--site SITE] id

positional arguments:
  id           Native ID or m365: stable handle

options:
  -h, --help   show this help message and exit
  --site SITE  SharePoint site root; required unless a stable handle or URL
               supplies it
```

## m365 lists create

```text
usage: m365 lists create [-h] [--site SITE] [--request-id REQUEST_ID]
                         [--dry-run] --name NAME

options:
  -h, --help            show this help message and exit
  --site SITE           SharePoint site root; required unless a stable handle
                        or URL supplies it
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --name NAME
```

## m365 lists rename

```text
usage: m365 lists rename [-h] [--site SITE] [--request-id REQUEST_ID]
                         [--dry-run] [--if-version IF_VERSION] --name NAME
                         id

positional arguments:
  id                    Native ID or m365: stable handle

options:
  -h, --help            show this help message and exit
  --site SITE           SharePoint site root; required unless a stable handle
                        or URL supplies it
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Exact quoted ETag from the matching get/metadata
                        command
  --name NAME
```

## m365 lists recycle

```text
usage: m365 lists recycle [-h] [--site SITE] [--request-id REQUEST_ID]
                          [--dry-run] [--allow-unconditional]
                          id

positional arguments:
  id                    Native ID or m365: stable handle

options:
  -h, --help            show this help message and exit
  --site SITE           SharePoint site root; required unless a stable handle
                        or URL supplies it
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --allow-unconditional
                        Required: endpoint has no verified atomic concurrency
                        guard
```

## m365 lists column

```text
usage: m365 lists column [-h] {create} ...

positional arguments:
  {create}
    create    Create a simple column on a visible generic list

options:
  -h, --help  show this help message and exit
```

## m365 lists column create

```text
usage: m365 lists column create [-h] [--site SITE] [--request-id REQUEST_ID]
                                [--dry-run] --list LIST --name NAME
                                --type {text,choice,person,number,boolean,datetime,url,multi-choice,multi-person}
                                [--choice CHOICE] [--required]

options:
  -h, --help            show this help message and exit
  --site SITE           SharePoint site root; required unless a stable handle
                        or URL supplies it
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --list LIST
  --name NAME           ASCII internal name, 1–32 letters/digits/underscores
  --type {text,choice,person,number,boolean,datetime,url,multi-choice,multi-person}
  --choice CHOICE       Repeat for each allowed choice
  --required            Only allowed on an empty list
```

## m365 lists item

```text
usage: m365 lists item [-h]
                       {list,get,create,update,recycle,versions,version} ...

positional arguments:
  {list,get,create,update,recycle,versions,version}
    list                List item
    get                 Get item
    create              Create item
    update              Update item
    recycle             Recycle item
    versions            Versions item
    version             Version item

options:
  -h, --help            show this help message and exit
```

## m365 lists item list

```text
usage: m365 lists item list [-h] [--site SITE] [--limit LIMIT]
                            [--cursor CURSOR] [--list LIST]

options:
  -h, --help       show this help message and exit
  --site SITE      SharePoint site root; required unless a stable handle or
                   URL supplies it
  --limit LIMIT    Page size 1–500; follow next_cursor
  --cursor CURSOR  Opaque cursor from this exact account/site/query/limit
  --list LIST      List GUID or stable handle; required unless item handle
                   supplies it
```

## m365 lists item get

```text
usage: m365 lists item get [-h] [--site SITE] [--list LIST] id

positional arguments:
  id           Native ID or m365: stable handle

options:
  -h, --help   show this help message and exit
  --site SITE  SharePoint site root; required unless a stable handle or URL
               supplies it
  --list LIST  List GUID or stable handle; required unless item handle
               supplies it
```

## m365 lists item create

```text
usage: m365 lists item create [-h] [--site SITE] [--request-id REQUEST_ID]
                              [--dry-run] [--list LIST] (--data DATA |
                              --data-file DATA_FILE)

options:
  -h, --help            show this help message and exit
  --site SITE           SharePoint site root; required unless a stable handle
                        or URL supplies it
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --list LIST           List GUID or stable handle; required unless item
                        handle supplies it
  --data DATA           JSON object using schema internal field names
  --data-file DATA_FILE
                        UTF-8 JSON object file, maximum 1 MiB
```

## m365 lists item update

```text
usage: m365 lists item update [-h] [--site SITE] [--request-id REQUEST_ID]
                              [--dry-run] [--if-version IF_VERSION]
                              [--list LIST] (--data DATA |
                              --data-file DATA_FILE)
                              id

positional arguments:
  id                    Native ID or m365: stable handle

options:
  -h, --help            show this help message and exit
  --site SITE           SharePoint site root; required unless a stable handle
                        or URL supplies it
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --if-version IF_VERSION
                        Exact quoted ETag from the matching get/metadata
                        command
  --list LIST           List GUID or stable handle; required unless item
                        handle supplies it
  --data DATA           JSON object using schema internal field names
  --data-file DATA_FILE
                        UTF-8 JSON object file, maximum 1 MiB
```

## m365 lists item recycle

```text
usage: m365 lists item recycle [-h] [--site SITE] [--request-id REQUEST_ID]
                               [--dry-run] [--allow-unconditional]
                               [--list LIST]
                               id

positional arguments:
  id                    Native ID or m365: stable handle

options:
  -h, --help            show this help message and exit
  --site SITE           SharePoint site root; required unless a stable handle
                        or URL supplies it
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --allow-unconditional
                        Required: endpoint has no verified atomic concurrency
                        guard
  --list LIST           List GUID or stable handle; required unless item
                        handle supplies it
```

## m365 lists item versions

```text
usage: m365 lists item versions [-h] [--site SITE] [--limit LIMIT]
                                [--cursor CURSOR] [--list LIST]
                                id

positional arguments:
  id               Native ID or m365: stable handle

options:
  -h, --help       show this help message and exit
  --site SITE      SharePoint site root; required unless a stable handle or
                   URL supplies it
  --limit LIMIT    Page size 1–500; follow next_cursor
  --cursor CURSOR  Opaque cursor from this exact account/site/query/limit
  --list LIST      List GUID or stable handle; required unless item handle
                   supplies it
```

## m365 lists item version

```text
usage: m365 lists item version [-h] [--site SITE] [--list LIST]
                               --version VERSION
                               id

positional arguments:
  id                 Native ID or m365: stable handle

options:
  -h, --help         show this help message and exit
  --site SITE        SharePoint site root; required unless a stable handle or
                     URL supplies it
  --list LIST        List GUID or stable handle; required unless item handle
                     supplies it
  --version VERSION
```

## m365 lists attachment

```text
usage: m365 lists attachment [-h] {list,get,download,add,recycle} ...

positional arguments:
  {list,get,download,add,recycle}
    list                List attachment
    get                 Get attachment
    download            Download attachment
    add                 Add attachment
    recycle             Recycle attachment

options:
  -h, --help            show this help message and exit
```

## m365 lists attachment list

```text
usage: m365 lists attachment list [-h] [--site SITE] [--limit LIMIT]
                                  [--cursor CURSOR] [--list LIST]
                                  id

positional arguments:
  id               Native ID or m365: stable handle

options:
  -h, --help       show this help message and exit
  --site SITE      SharePoint site root; required unless a stable handle or
                   URL supplies it
  --limit LIMIT    Page size 1–500; follow next_cursor
  --cursor CURSOR  Opaque cursor from this exact account/site/query/limit
  --list LIST      List GUID or stable handle; required unless item handle
                   supplies it
```

## m365 lists attachment get

```text
usage: m365 lists attachment get [-h] [--site SITE] [--list LIST] --name NAME
                                 id

positional arguments:
  id           Native ID or m365: stable handle

options:
  -h, --help   show this help message and exit
  --site SITE  SharePoint site root; required unless a stable handle or URL
               supplies it
  --list LIST  List GUID or stable handle; required unless item handle
               supplies it
  --name NAME  Exact attachment name on the record
```

## m365 lists attachment download

```text
usage: m365 lists attachment download [-h] [--site SITE] [--list LIST]
                                      --name NAME --output OUTPUT
                                      id

positional arguments:
  id               Native ID or m365: stable handle

options:
  -h, --help       show this help message and exit
  --site SITE      SharePoint site root; required unless a stable handle or
                   URL supplies it
  --list LIST      List GUID or stable handle; required unless item handle
                   supplies it
  --name NAME      Exact attachment name on the record
  --output OUTPUT
```

## m365 lists attachment add

```text
usage: m365 lists attachment add [-h] [--site SITE] [--request-id REQUEST_ID]
                                 [--dry-run] [--list LIST] --file FILE
                                 [--name NAME]
                                 id

positional arguments:
  id                    Native ID or m365: stable handle

options:
  -h, --help            show this help message and exit
  --site SITE           SharePoint site root; required unless a stable handle
                        or URL supplies it
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --list LIST           List GUID or stable handle; required unless item
                        handle supplies it
  --file FILE
  --name NAME
```

## m365 lists attachment recycle

```text
usage: m365 lists attachment recycle [-h] [--site SITE]
                                     [--request-id REQUEST_ID] [--dry-run]
                                     [--allow-unconditional] [--list LIST]
                                     --name NAME
                                     id

positional arguments:
  id                    Native ID or m365: stable handle

options:
  -h, --help            show this help message and exit
  --site SITE           SharePoint site root; required unless a stable handle
                        or URL supplies it
  --request-id REQUEST_ID
                        Stable unique ID for duplicate prevention across
                        retries
  --dry-run             Validate local input and print the plan without
                        connecting or writing
  --allow-unconditional
                        Required: endpoint has no verified atomic concurrency
                        guard
  --list LIST           List GUID or stable handle; required unless item
                        handle supplies it
  --name NAME           Exact attachment name on the record
```

## m365 lists comments

```text
usage: m365 lists comments [-h] {list} ...

positional arguments:
  {list}
    list      List comments

options:
  -h, --help  show this help message and exit
```

## m365 lists comments list

```text
usage: m365 lists comments list [-h] [--site SITE] [--limit LIMIT]
                                [--cursor CURSOR] [--list LIST]
                                id

positional arguments:
  id               Native ID or m365: stable handle

options:
  -h, --help       show this help message and exit
  --site SITE      SharePoint site root; required unless a stable handle or
                   URL supplies it
  --limit LIMIT    Page size 1–500; follow next_cursor
  --cursor CURSOR  Opaque cursor from this exact account/site/query/limit
  --list LIST      List GUID or stable handle; required unless item handle
                   supplies it
```

## m365 maintainer

```text
usage: m365 maintainer [-h]
                       {capabilities,check,validate,disable,enable,disabled} ...

positional arguments:
  {capabilities,check,validate,disable,enable,disabled}
    capabilities        Inventory of every implemented command; not a live
                        availability claim
    check               Read-only API and detached HTTP probes; writes remain
                        untested
    validate            Run an explicit disposable-fixture write suite;
                        journals stay private
    disable             Set/remove a local command-prefix circuit breaker
    enable              Set/remove a local command-prefix circuit breaker
    disabled            Show configured circuit breakers

options:
  -h, --help            show this help message and exit
```

## m365 maintainer capabilities

```text
usage: m365 maintainer capabilities [-h]

options:
  -h, --help  show this help message and exit
```

## m365 maintainer check

```text
usage: m365 maintainer check [-h]
                             [--module {all,outlook,teams,tasks,sharepoint}]
                             [--site SITE] [--output OUTPUT]
                             [--baseline BASELINE]

options:
  -h, --help            show this help message and exit
  --module {all,outlook,teams,tasks,sharepoint}
  --site SITE           Explicit SharePoint site for read probes
  --output OUTPUT       Save a content-free JSON report (must be a new path)
  --baseline BASELINE   Compare with a prior maintainer-check report
```

## m365 maintainer validate

```text
usage: m365 maintainer validate [-h]
                                --suite {outlook,teams,tasks,references,attachments,labels,availability,composition,calendar-composition,calendar-drafts,teams-composition,sharepoint}
                                [--site SITE] --allow-test-writes
                                [--recipient RECIPIENT] [--channel CHANNEL]
                                [--attendee ATTENDEE]

options:
  -h, --help            show this help message and exit
  --suite {outlook,teams,tasks,references,attachments,labels,availability,composition,calendar-composition,calendar-drafts,teams-composition,sharepoint}
  --site SITE           SharePoint suite: explicitly authorized personal
                        OneDrive site root
  --allow-test-writes
  --recipient RECIPIENT
                        Required for Outlook/composition/calendar-drafts;
                        authorized test recipient
  --channel CHANNEL     Optional explicitly authorized Teams channel; self-
                        chat by default
  --attendee ATTENDEE   Optional read-only colleague lookup for the
                        availability suite; never sends invitations
```

## m365 maintainer disable

```text
usage: m365 maintainer disable [-h] --reason REASON capability

positional arguments:
  capability       Quoted command prefix, e.g. 'todo task update' or 'teams'

options:
  -h, --help       show this help message and exit
  --reason REASON
```

## m365 maintainer enable

```text
usage: m365 maintainer enable [-h] capability

positional arguments:
  capability  Quoted command prefix, e.g. 'todo task update' or 'teams'

options:
  -h, --help  show this help message and exit
```

## m365 maintainer disabled

```text
usage: m365 maintainer disabled [-h]

options:
  -h, --help  show this help message and exit
```
