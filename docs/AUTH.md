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
