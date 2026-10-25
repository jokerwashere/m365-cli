# m365-cli

A Python CLI for Microsoft 365 mail, calendars, Teams, To Do, Planner, SharePoint
files and lists, designed for people and agents. The command is `m365`. It reads
and changes service data through Outlook Web, Teams Web, task-service and SharePoint HTTP APIs. A dedicated, persistent
Chrome profile supplies authentication; data operations use direct HTTP.

No Microsoft Entra application registration is required. Sign in normally,
including any MFA required by your organization. Native Outlook can be closed.
Outlook and SharePoint disconnect from browser automation after verifying the
session. Each new command still needs Chrome for authentication; no separate
credential cache or unattended renewal service is installed.

## Why this exists

I built this tool around a simple principle: **if I can do something manually in
Microsoft 365, my agents should be able to do it too.** That should be possible
without an interface-clicking overlay or waiting for corporate IT to approve a
separate Microsoft Graph application. The CLI uses my existing sign-in and account
permissions to give agents direct access to the services I work with.

I use it to:

- Triage email and identify what needs my attention.
- Find appointment slots across calendars.
- Publish project artefacts to SharePoint.
- Check updates in Teams chats and channels.
- Prepare email drafts.
- Collect project-relevant information from stakeholders across these services.

For email, **I always review the drafts and send them personally**. Agents help
with the reading, organization and preparation; I make the final call on what
gets sent.

## Install with uv

You need **macOS**, **Google Chrome**, and a Microsoft 365 work or school account
with access to the services you want to use. Install Chrome in `/Applications`.
Python 3.11 or newer is required; the command below asks uv to use Python 3.13
and download it if needed. Windows is not supported. Other platforms require
platform-specific browser configuration and are outside the supported setup.

1. Install [uv](https://docs.astral.sh/uv/getting-started/installation/) if it is
   not already installed:

   ```sh
   curl -LsSf https://astral.sh/uv/install.sh | sh
   ```

   Reopen your terminal, then check `uv --version`. You also need Git for the
   GitHub installation below: check `git --version`. On macOS, if Git is missing,
   run `xcode-select --install` and finish installing the Command Line Tools.

2. Install this repository as an isolated command-line tool:

   ```sh
   uv tool install --python 3.13 'git+https://github.com/jokerwashere/m365-cli.git'
   uv tool update-shell
   ```

3. Reopen your terminal so the tool directory is on `PATH`, then check:

   ```sh
   m365 --version
   m365 --help
   ```

The package is named `m365-cli`; its command is **`m365`**. You can run it from
any directory without activating a virtual environment. Use the GitHub URL above
to install this project. No separate Playwright browser download is needed: the
CLI uses your installed Google Chrome. See uv's
[tool guide](https://docs.astral.sh/uv/guides/tools/) for environment and PATH details.

## First run: sign in and read your inbox

Replace `you@example.com` with the primary email address of your Microsoft 365
mailbox, then run:

```sh
m365 --account you@example.com auth login
```

A dedicated Chrome window opens. Complete Microsoft sign-in and any MFA or
organization prompts **in that window**. Leave the command running while you
sign in; it waits up to five minutes. Success returns JSON with `ok: true` and
`data.authenticated: true`. The CLI saves the verified mailbox address and checks
it on subsequent commands. It does not use the native Outlook application's login.

Confirm the session and read up to five inbox messages:

```sh
m365 --account you@example.com auth status
m365 --account you@example.com mail list --folder inbox --limit 5
```

These commands send nothing. A working session reports `authenticated: true`.
An empty inbox can return an empty collection. Normal results and errors are
JSON; help and version output are plain text. The explicit `--account` check is
recommended for agents and scripts, even though the verified account is saved.

### Enable additional services

Mail, calendar and availability use the Outlook session above. For the other
services, run the appropriate login command once and complete any prompts in the
same dedicated browser. Existing Microsoft single sign-on may satisfy them.

```sh
# Teams
m365 --account you@example.com teams login
m365 --account you@example.com teams status

# To Do and Planner; keep the dedicated browser running from auth login
m365 --account you@example.com tasks login
m365 --account you@example.com tasks status

# SharePoint files and lists; replace this with a site you can access
m365 --account you@example.com files login --site 'https://tenant.sharepoint.com/sites/example'
m365 --account you@example.com files status --site 'https://tenant.sharepoint.com/sites/example'
```

Service permissions still apply. See [Teams](docs/TEAMS.md), [tasks](docs/TASKS.md)
and [SharePoint / OneDrive](docs/SHAREPOINT.md) for service-specific setup and usage.

### Saved sessions and recovery

A tool installation stores its persistent Chrome profile, verified account and
operation records in **`~/.local/share/m365-cli`**. Commands reuse that state;
upgrading the package does not erase it. Keep the dedicated Chrome profile running
while using the CLI. If you close it or restart your Mac, run `auth login` again
to reopen it. The saved browser session usually avoids another password prompt,
but Microsoft can require sign-in or MFA again when a session expires.

Set `M365_HOME` to a persistent directory **before login** if you want a different
state location or separate profiles for different accounts. Keep using the same
value for later commands. Never commit or share this directory: it contains
browser authentication and private operation data. Editable source installations
use the checkout's `.state/` by default. See [authentication and recovery](docs/AUTH.md)
for browser configuration, session errors and state details.

If `m365` is not found, run `uv tool update-shell` and reopen your terminal. If the
browser is closed or authentication expires, rerun the relevant login command.
For an account mismatch, sign into the intended account in the dedicated profile;
do not remove the account check to bypass the error.

## Update or uninstall

```sh
uv tool upgrade m365-cli
m365 --version
```

To remove the installed command:

```sh
uv tool uninstall m365-cli
```

Uninstalling the tool leaves its Microsoft session and private data directory in
place. Remove that directory separately only if you intend to discard its saved
login, references and operation history.

## Run from source instead

For development or an environment using the repository's exact locked dependency
versions:

```sh
git clone https://github.com/jokerwashere/m365-cli.git
cd m365-cli
uv sync --locked
uv run --locked m365 --version
uv run --locked m365 --account you@example.com auth login
uv run --locked m365 --account you@example.com mail list --folder inbox --limit 5
```

`uv sync --locked` uses `uv.lock`; a standalone `uv tool install` resolves the
package's declared dependencies independently. See [CONTRIBUTING.md](CONTRIBUTING.md)
for development commands.

## Everyday commands

```sh
m365 mail list --folder inbox --limit 20
m365 mail search 'from:person@example.com subject:project' --limit 10
m365 mail get 'MESSAGE_ID'
m365 mail draft create --to person@example.com --subject 'Project update' --body-file message.txt
m365 mail draft update 'DRAFT_ID' --body-file revised-message.txt --if-change-key 'CHANGE_KEY'
m365 mail draft send 'DRAFT_ID' --request-id 'project-update-2026-10-06'
m365 mail update 'MESSAGE_ID' --read true
m365 mail move 'MESSAGE_ID' --folder archive
m365 calendar list --start '2026-10-06T00:00:00+02:00' --end '2026-10-13T00:00:00+02:00'
m365 calendar search 'planning' --start '2026-10-06T00:00:00+02:00' --end '2026-10-13T00:00:00+02:00'
m365 calendar create --subject 'Focus time' --start '2026-10-07T09:00:00+02:00' --end '2026-10-07T09:30:00+02:00'
```

Replace example IDs, addresses, and dates before running commands. `m365 --help`
and `m365 <group> <command> --help` describe the installed command surface.

## What is supported

| Area | Operations |
| --- | --- |
| Files / OneDrive | Libraries/folders, URL resolution, bounded listings/search, current/historical downloads, manifests, new-name uploads/copies, guarded rename, within-library move, restore and recycle |
| SharePoint lists | Schema/users, simple column creation, typed record lifecycle/history, comments reads and attachment add/read/recycle; stable IDs and URLs |
| Mail | List, search, text/HTML reads, create/edit/delete drafts, send new mail, send existing drafts, reply/reply-all, forward; plain text, HTML and styled Markdown composition |
| Mail management | Change read state, importance, categories and flags; move, copy, soft-delete, restore by moving from Deleted Items |
| Attachments | Files/inline files, attached items and recursive nesting; native message exports, structured item/link metadata, existing-item copies, draft add/remove; 10 MiB uploads and 50 MiB downloads |
| Calendar events | List, search, text/HTML reads, create, edit, soft-delete; save attendees without sending with `--save-only`, then `calendar send`; shared mail formatting/styles; appointments, all-day events and invitations |
| Availability / scheduling | Read multiple people's free/busy, required/optional attendees, common meeting slots, timezone/day/hour filters and explicit unknown handling |
| Meetings | Send organizer updates/cancellations; accept, tentatively accept or decline as an attendee |
| Folders | List/read/create/rename mail or calendar folders; delete empty custom folders |
| Teams | People lookup; chat/team/channel discovery; message history, text/HTML/Markdown sending, channel replies, editing and soft deletion; shared styles/previews; chat creation, group rename/leave, roster reads |
| Tasks | To Do and Planner CRUD, progress, assignment, recurrence, files, checklists, references, My Day, comments, moves and incremental sync; see task-specific limits |
| Categories / labels | Outlook/To Do catalog CRUD, color presets and item assignments; Planner slot naming, task labels and guarded color migration |
| Maintenance | Command inventory, read probes/baselines, disposable-fixture suites, per-command disable switches |
| Agent controls | JSON results/errors, account pinning, dry runs, change-key preconditions, command locking, local duplicate prevention |

Shared mail/event HTML/Markdown composition, local previews and styling are in [COMPOSITION.md](docs/COMPOSITION.md).

Native Teams drafts are unsupported. `teams message render` creates a local
preview; sending a Teams message is immediate.

Attachment types, nested-item examples and export limits are in
[ATTACHMENTS.md](docs/ATTACHMENTS.md).

Multi-person availability and meeting-slot selection are in [SCHEDULING.md](docs/SCHEDULING.md).

Teams setup, examples, supported limits, and pagination are in [TEAMS.md](docs/TEAMS.md).
Teams message search is deliberately excluded; people lookup remains available.

Task setup, payload examples, recurrence and limits are in [TASKS.md](docs/TASKS.md).
Category creation, colors, assignments and Planner label lifecycle are in [LABELS.md](docs/LABELS.md).
Maintainer capability probes, write suites and disable switches are in
[MAINTAINER.md](docs/MAINTAINER.md).
See [the command reference](docs/CLI.md), [the agent guide](docs/AGENTS.md),
[authentication and recovery](docs/AUTH.md), and [IDs and links](docs/REFERENCES.md).

## Writes and retries

`mail send` and `mail draft send` send immediately. Reply and forward commands
create drafts unless `--send` is specified. Adding attendees when creating a
calendar event sends invitations unless `--save-only` is supplied. Use `calendar update --save-only`
to prepare a never-sent event, then `calendar send EVENT_ID` when ready. Unsent
meetings remain in your calendar and can affect your own free/busy; they are not
mail drafts. `calendar get` exposes `MeetingRequestWasSent` as the authoritative
invitation flag (`IsDraft` can be false). Verified never-sent meetings can be
deleted without sending cancellations. Sent meetings still require `--notify`
for edits/deletion; `--save-only` cannot suppress their updates. Personal category
changes send no notifications. A calendar response notifies only with `--notify`.

Use `--dry-run` on any mutation to validate local input and review its plan without
connecting or writing. Server permissions, item existence, conflicts, and meeting
rules are checked only during execution.

Use a unique, stable `--request-id` for each intended mutation. Repeating the same
request ID with identical input returns its cached result without another write.
A different payload or account with the same ID is rejected. Teams and tasks also bind the
request to its service and tenant. An unfinished attempt
blocks retries under that ID. These records are local to `M365_HOME`; they are not
a server-side exactly-once guarantee. Keep the same state directory across agent
runs and do not reuse IDs for different actions.

There are no automatic write retries. After a timeout, interrupted operation, or
`write_outcome_unknown`, inspect the affected mailbox, chat, or channel before trying again. Successful
submission does not independently confirm recipient delivery.

Teams messages send immediately. Message edits use server-enforced version checks;
Teams soft deletion has **no atomic version guard** and rejects `--if-version`.
Only your own messages can be edited/deleted. See [Teams safeguards](docs/TEAMS.md).

## Practical boundaries

Web endpoints and Graph beta can change or be restricted by tenant policy.
Some task operations use documented Graph v1.0, with browser-bound authentication.
The CLI cannot override MFA, session expiration, server permissions, or retention
policies. No software can promise to be error-proof; the CLI fails conservatively
and reports uncertainty instead of silently retrying a write.

Mail search returns conversation hits, not every message in a thread. Calendar
search matches subject and location in an explicit window, with recurring
occurrences expanded by Outlook. `has_more: null` means completeness is unknown.
All-day boundaries must be midnight in the mailbox timezone shown by `auth status`.
For mail/calendar, entire recurring-series mutations, recurrence-rule creation, shared/delegated
mailbox routing, permanent deletion, cloud-link creation/automatic
cloud-file downloads, nested attachment mutation and attachment replacement in one
operation are outside the supported scope.
To replace a draft attachment, remove it and add the replacement.

## IDs and shareable links

Resource results include `stable_id` and `web_url`. Store the stable handle and resolve its latest known ID/link with `m365 reference get STABLE_ID`. Native links require recipient access; child items explicitly label parent links. See [IDs and links](docs/REFERENCES.md) for move tracking, local storage and service limitations.

## SharePoint files and lists

`m365 files` and `m365 lists` use the saved browser session to manage SharePoint
and OneDrive files, folders, generic lists, typed records and attachments. Resolve
a document or list-item URL with `m365 reference resolve URL`, then store its
`stable_id` and `web_url`. See [the files/lists guide](docs/SHAREPOINT.md) for
setup, examples, version guards and scope. New uploads reject collisions; file
content replacement is unsupported because the available transport cannot enforce
an atomic version precondition. Move, restore and recycle require explicit `--allow-unconditional`.
