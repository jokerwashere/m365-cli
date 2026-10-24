"""Generate the complete command reference from the installed argument parser."""

import argparse
from pathlib import Path

from m365_cli.cli import parser

INTRO = """# Command reference

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
"""


def render():
    sections = [INTRO]

    def visit(p, parts):
        sections.append(
            "## " + " ".join(parts) + "\n\n```text\n" + p.format_help().rstrip() + "\n```\n"
        )
        for action in p._actions:
            if isinstance(action, argparse._SubParsersAction):
                for name, child in action.choices.items():
                    visit(child, parts + [name])

    visit(parser(), ["m365"])
    return "\n".join(sections)


if __name__ == "__main__":
    (Path(__file__).parents[1] / "docs" / "CLI.md").write_text(render())
