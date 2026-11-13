# Authentication, storage, and recovery

`m365 auth login --account you@example.com` opens a dedicated Chrome profile.
Complete the normal Microsoft sign-in and MFA. The verified mailbox address is
saved and checked before later commands. `auth status` also reports the mailbox
timezone used for calendar writes and all-day validation. An optional global `--account EMAIL`
checks the expected mailbox for one invocation:

```sh
m365 --account you@example.com mail list --limit 5
```

Authentication is not copied from the native Outlook application. The browser
profile holds Microsoft cookies and browser-managed session storage. Tokens and
canary headers are used in memory by the transport; the CLI does not export a
separate token cache, store your password, or print authentication headers.

## Transport and browser dependence

Online service commands use the dedicated browser to bootstrap authentication.
Since 0.12.0, Outlook mail/calendar/availability and SharePoint files/lists then
disconnect from browser automation and use direct HTTPS for all service operations.
Bootstrap may make a read-only identity request in the page to capture its
applicable credentials. The direct client verifies server identity independently;
before mutations it verifies the account again (and SharePoint site/web identity).
There is no page-based fallback after a failed request.

A known read may retry once after an interrupted connection, within the same
45-second budget and with the same frozen credentials. Authentication rejection,
throttling, redirects, oversized responses and changed contracts are not retried.
Mutation requests are never retried, including after successful authentication
recovery. An acknowledged write's readback remains a read; a readback failure
includes the sanitized underlying cause.

Credentials are frozen for the command and remain in memory. SharePoint captures
only cookies actually sent by its verified bootstrap request, preserving their
host/path/expiry bounds, plus the observed User-Agent. Native mutation request
digests are obtained through direct HTTP. Cookie scope metadata is sampled before
and after bootstrap to tolerate cookies rotating during that identity read;
only values actually observed on the request are exported. Outlook bearer authentication
uses its own scoped service headers without browser cookies, canary or session-ID
headers; canary-only sessions retain their applicable cookies. Attachment exports
send their dedicated download token, mailbox routing and observed User-Agent to
the fixed Microsoft attachment origin. Redirects are rejected, and response bodies
are read with time and size limits. No service credentials cross to another service.

`auth status` and `files status` include `transport` with `mode: "direct_http"`,
`browser_connected: false`, and `credential_storage: "command_memory"` after
successful authentication. This describes the automation connection; Chrome's
window can remain open. Teams/tasks already use direct HTTP, but currently retain
their browser connection during commands.

Each new command still needs browser bootstrap. No additional credential cache
or session daemon has been introduced, and expiry/MFA can still require sign-in.
Authentication renewal only prepares a command; it never replays an attempted
mutation.

## State location

For an editable installation from the repository, state lives in `.state/` at the
repository root. For a regular package installation, the default is
`~/.local/share/m365-cli/`. Set `M365_HOME` to use a particular absolute directory.
Always use the same directory to preserve both login and duplicate prevention.

| Path under the state directory | Purpose |
| --- | --- |
| `profile/` | Dedicated Chrome profile, including Microsoft authentication state |
| `config.json` | Verified mailbox address and pinned Teams tenant; no password or token |
| `command.lock` | Prevents concurrent CLI operations using the same state |
| `operations/` | Private request fingerprints, pending attempts and cached mutation results |
| `validation-*.json`, `teams-live-*.json` | Private validation results and cleanup journal, when live validation is run |

The state directory is mode `0700`; CLI-created records are private files. It is
excluded from Git. Treat the entire directory as sensitive: browser sessions and
cached operation results can provide account access or contain mailbox data.
Do not commit, upload, or share it.

The CLI connects to Chrome through a loopback debugging endpoint. Processes
running as your local OS user may be able to control that dedicated session.
Chrome remains open when a command completes. Closing it does not delete its
profile. After Chrome closes or the machine restarts, run `m365 auth login` to
reopen the same profile. An interactive sign-in may be required again; the saved profile does not guarantee
silent recovery. Microsoft's session lifetime and organizational policies still apply.

`M365_BROWSER` can override the Chrome executable path. The macOS default is
`/Applications/Google Chrome.app/Contents/MacOS/Google Chrome`.

## Progress, deadlines and outages

Explicit `auth`, `teams`, `tasks`, and `files` **login/status** commands report
stage changes and a heartbeat every five seconds on stderr. Progress includes
elapsed time and the total time limit, never account identifiers or credentials.
Stdout remains exactly one JSON result. `--quiet` suppresses progress, not errors.

```sh
m365 --account you@example.com auth login --timeout 300
m365 --account you@example.com auth status --timeout 20
m365 --account you@example.com teams status --timeout 30 --quiet
m365 --account you@example.com tasks status --timeout 30
m365 --account you@example.com files status --site 'https://tenant.sharepoint.com/sites/project' --timeout 30
```

The total authentication-work deadline defaults to **300 seconds for login** and
**30 seconds for status**. `--timeout` accepts 1–3600 seconds. For compatibility,
Outlook/SharePoint `login --timeout 0` skips waiting for interactive sign-in but
still allows up to 30 seconds for verification. A larger value gives more time
for MFA or slow loading; it cannot repair an outage.

The overall limit is a ceiling, not a promise to wait that long. Individual steps
can time out earlier. In particular, connecting to **local Chrome** has a separate
**10-second connection limit**, before the Microsoft account check. Increasing
`--timeout` does not extend this connection limit. Progress names both limits:

```text
[m365] Connecting to the saved Chrome profile (1s total elapsed; 10s step limit; 300s overall limit).
```

A Chrome connection timeout reports `timeout_scope: "operation"`,
`operation_timeout_seconds: 10` and `command_timeout_seconds: 300` for the default
login budget. `elapsed_seconds` measures the entire command, including startup
and cleanup, so a value around 11 seconds is compatible with that 10-second step
limit. `timeout_seconds` remains an alias of `command_timeout_seconds` for existing
scripts; it does not identify the limit that expired. If the overall deadline
expires first, the error is `auth_timeout` with `timeout_scope: "command"`.
Other browser/network timeout errors identify the operation scope but may not
include an exact operation limit when it is unavailable.

For `browser_unavailable` at `auth_stage: "browser_connect"`, check whether the
dedicated Chrome window is responsive. If it is stuck, close only that window
and rerun the relevant login command to reopen the saved profile. Keep the profile
directory. For tasks, reopen Chrome with `m365 auth login` first, then run
`m365 tasks login`. This local connection failure is not evidence of a Microsoft outage
or expired credentials.

A parent process enforces this deadline even if DNS, TLS, browser evaluation or
driver cleanup stalls. It stops only the authentication worker and its driver,
leaving the dedicated Chrome process, saved profile and operation history intact.
Worker shutdown may take up to one additional second. This supervisor applies
only to explicit login/status commands; it does not wrap ordinary reads or writes.
Authentication network reads also use shorter request/connection timeouts. A known
read may retry once within its budget. **No worker or service mutation is retried.**

`auth_timeout` includes the last `auth_stage`, elapsed time, configured timeout
and recovery advice. Network failures retain `transport_error` and their phase;
a browser fetch timeout is not treated as rejected credentials. A disconnected
browser/driver returns `browser_unavailable` or `auth_process_failed`. Raw driver
pipe warnings are filtered, and a cleanup failure does not hide a preceding API
error. HTTP 5xx responses retain their HTTP status; 502/504 include
`reason: "service_timeout"`. These authentication errors report
`mutation_request_sent: false`.

If the web app also stalls or Microsoft returns 5xx, keep the existing profile,
check connectivity/service health and retry the read-only status check after
recovery. If Microsoft actually presents a sign-in/MFA prompt, complete it in the
saved browser and run login again. Authentication success does not authorize
replaying an earlier, unresolved write.

## Recovery

- **`login_required` / `session_rejected`:** run `m365 auth login`. If Microsoft
  rejects the account or tenant policy blocks access, resolve that in the normal
  sign-in flow. The CLI does not bypass it.
- **`account_mismatch`:** check the dedicated browser's account and your expected
  address. Do not remove the guard to send from an unintended mailbox.
- **`busy`:** another command holds the lock. Wait for that command to finish.
- **`conflict`:** read the item again, reconcile the intended changes, and submit
  with its new change key. Do not overwrite someone else's edit automatically.
- **`write_outcome_unknown` / `interrupted`:** check Sent Items, Drafts, or the
  calendar using the original subject, time, recipient, and body. Inspect a local
  attempt with `m365 operation get REQUEST_ID`. Do not resend
  automatically. Use a new request ID only after reconciling the prior outcome.
- **`readback_failed`:** the write was acknowledged but its subsequent read failed.
  Read the item again before deciding whether another write is needed.
- **`request_id_conflict`:** one request ID was reused for a different payload or
  account. Keep IDs stable for retries and unique for distinct operations.
- **`already_deleted`:** the item is already in Deleted Items. Permanent deletion
  is intentionally unavailable. Use `mail move` to restore mail to a chosen folder.

To stop using the session, sign out in the dedicated Chrome window and close it.
Deleting the profile also removes local browser state, but does not revoke other
Microsoft sessions. Keep operation records if you still need retry protection.

## Teams session

Run `m365 --account you@example.com teams login` once. Teams opens in the same
persistent Chrome profile and may reuse Microsoft SSO. Later commands connect to
that profile and reload Teams to let its normal client renew service credentials.
The CLI captures only allowlisted headers from the exact Teams HTTPS origin;
credentials are kept in memory, scoped to the relevant service, and frozen for the
command. Outlook headers are never substituted for Teams headers.

Every Teams command checks token expiry, account, user, resource audience, and
tenant agreement, then verifies user/tenant against a server response. The first
verified tenant is stored as `teams_tenant` in `config.json`. A subsequent switch
fails with `tenant_mismatch`; inspect the signed-in tenant. An explicit
`teams --tenant UUID ...` supplies the expected tenant for that invocation rather
than silently repinning configuration. Account/tenant ambiguity fails closed.

After closing Chrome, `m365 teams login` reopens the existing profile. Normal MFA,
revocation, and session-expiry rules still apply. Teams-specific auth failures use
`teams login` for recovery. `teams status` reports `authenticated: false` if the
browser is closed. Loading Teams has normal browser-client background activity.

Teams retries use the same `operations/` ledger as mail, with service and tenant in
the fingerprint. For uncertain writes, inspect the explicit chat/channel and item
IDs, not Outlook Sent Items. `already_deleted` refers to a Teams tombstone;
restoring deleted Teams messages is not supported. No separate token cache exists.

## Task services

`m365 tasks login` loads/reloads Planner in the existing dedicated browser and
waits for Graph/native task credentials. `tasks status` verifies Graph identity,
native Planner identity, account/tenant, audiences, scopes and expiry. Commands
capture only the service credentials they need, keep them in memory, and discard
them on exit. Browser cookies and its own login state remain in the existing
profile; the CLI does not persist another token cache. The task adapter never
restarts Chrome as recovery. A login or Conditional Access prompt must still be
completed by the user. Reusing a profile does not guarantee perpetual silent login.
