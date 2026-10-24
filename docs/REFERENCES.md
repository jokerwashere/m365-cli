# Item IDs and links

Managed resource results include `id`, `stable_id`, `web_url`, and `reference`.
Collections put these fields on each resource. Teams send/update/delete put them
on `message`; chat creation puts them on `chat`. Attachment results retain their
existing parent IDs; attachment deletion adds `attachment_reference`. Planner
checklists, external references, assignments and labels have `child_references`
sidecars so the original provider maps remain usable as reviewed API data.

| Field | Meaning |
| --- | --- |
| `id` | Current provider ID. Use this in service commands, with the required parent IDs. Treat IDs as opaque and case sensitive. |
| `stable_id` | Durable CLI handle, scoped to the account, service, tenant and resource. Store this alongside the URL. |
| `web_url` | Native Microsoft browser link, or null when a usable link cannot be established. |
| `reference.url_kind` | `item`, `parent`, `application`, `action`, or `unavailable`. Check this before presenting a link as a direct item link. |
| `reference.locator` | Latest known provider ID and applicable list, plan, task or thread context. |
| `reference.status` | Last observed state; deletion acknowledgements and parent deletions are explicit. |
| `reference.observed_at` | Observation time, not a freshness guarantee. |

## Store, resolve and share

```sh
m365 planner task create --plan 'PLAN_ID' --title 'Review drawings' --request-id 'drawings-001'
m365 planner task assign 'TASK_ID' --user 'DIRECTORY_USER_UUID' --request-id 'drawings-assign-001'
m365 reference get 'm365:STABLE_UUID'
```

Keep the returned `stable_id` in your database and use `web_url` in a message or
project document. Assignment returns the same task reference and a task URL.
Only assign people and send messages when authorized. Sharing a URL does not
send a notification or grant access: the recipient must already have access to
the task's plan/list, channel or mailbox. Personal task and mailbox links are not
public sharing links; assigning someone also does not make a private plan public.

`reference get` reads the local registry without opening a browser. It returns
`cached: true` and `existence_verified: false`. Re-read the resource through its
service command before changing it, and obtain a fresh version/change key. Service
commands accept provider IDs, not `stable_id` values. Explicit `--account EMAIL`
can select a registry; otherwise the pinned account is used.

## What remains stable

The private registry lives at `$M365_HOME/references/<account-hash>.json` alongside
the existing auth and operation state. It contains identifiers, parent locators,
links and observation metadata, not titles, bodies, tokens or browser headers.
Teams directory references canonicalize organization-user MRIs to directory
object IDs while retaining the original roster fields.

Files have mode 0600, the directory 0700, and writes use atomic replacement under
the command lock. Back up this directory with the rest of your private state.

A successful CLI mail move or To Do task move records the new provider ID as an
alias of the original `stable_id`. Copying creates a different reference. Planner
moves update the known plan context and URL. An old saved URL is not rewritten in
external documents: resolve the stable ID to retrieve the latest URL. There is no
hosted redirect service. The CLI handle is a local reference, not a Microsoft ID,
public URL, globally synchronized identifier, or cross-device identity guarantee.

Moves performed outside the CLI cannot be reconciled reliably from names or
subjects. Missing IDs, copies, imports, mailbox migration and recreated child
resources are never matched by guesswork. Re-read/reconcile explicitly. A parent
move updates known child parent links; `locator_needs_refresh: true` warns that a
child's provider ID may have changed. Re-list children before another child write.
Deletion preserves historical references and marks their state; it does not keep
Microsoft links alive. Recurring To Do completion may reuse an ID for the next
occurrence, so references identify that service resource, not an immutable
occurrence in time. Re-list after recurrence changes.

If reference persistence fails after Microsoft has acknowledged an operation,
the successful operation result includes `reference_warning` with `persisted:
false`. Preserve that output and repair local storage. **Do not repeat the remote
write to repair the registry.** Completed request-ID replays preserve the original
result; `reference get` reports the latest known mapping after later moves.

## Link coverage and limits

| Resource | Link target |
| --- | --- |
| Mail/draft, event, mail folder | Item/folder in Outlook. Mailbox access is required. |
| Outlook/To Do category definition | Outlook application; no supported category permalink. CLI rename aliases the replaced provider ID to its existing stable handle. |
| Outlook calendar container/group | Calendar application; no supported isolated container permalink. |
| Mail attachment, including nested items | Root mail. `--mail-id` verifies membership and supplies the root; recursive lists/get include child references. Download tokens never become URLs in the registry. |
| To Do list/task | List or task details in To Do. Native route retains trailing `=` in the opaque ID. |
| To Do step/link/file/assignment | Parent task. A linked resource's `webUrl` remains its external destination; `web_url` is its management link. |
| To Do folder group | To Do application; no native group permalink is available. |
| Planner plan/task | Plan or task in Planner, including tenant context. Task links require plan context. |
| Planner bucket/labels/history | Parent plan. Individual label IDs are scoped to the plan and identify fixed color slots; recolor returns a different slot/reference. |
| Planner details/checklist/reference/comment/assignment | Parent task. Details/comments fetch parent context before mutations. |
| Teams chat/team/channel/message | Native Teams deep link. Channel messages include group, tenant and reply-root context. |
| Teams person | Open-chat action when the address is known; roster entries otherwise link to the parent chat. It sends nothing. |

Some responses cannot identify an item: immediate mail sends return a submission
acknowledgement without the sent item's ID. They explicitly return null references
and a reason. Read Sent Items afterward; never resend merely to get a link. Search
hits and sparse delta/tombstone records may need a full item/parent read before a
URL can be supplied. Such records say `unavailable`, rather than inventing an ID
or pointing to a guessed item. No service promises every child a standalone URL.

## Link behavior

URL routes are isolated in `m365_cli/references.py`. Teams links follow the
[Microsoft Teams deep-link formats](https://learn.microsoft.com/en-us/microsoftteams/platform/concepts/build-and-test/deep-link-teams).
Planner and To Do links use browser-client routes that can change. Exchange item
link conventions are described
by [WebClientReadFormQueryString](https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/webclientreadformquerystring).
[Graph immutable IDs](https://learn.microsoft.com/en-us/graph/outlook-immutable-id)
are an opt-in Graph contract; this CLI's OWA adapter does not claim that guarantee.

Availability blocks and meeting-slot suggestions are snapshots rather than managed
calendar items. They explicitly report `reference_unavailable` and do not invent
stable IDs or permalinks. A subsequent, separately authorized `calendar create`
returns a normal event reference. See [SCHEDULING.md](SCHEDULING.md).

## SharePoint references

`files`, `lists` and `reference resolve` register account/tenant/site/web-scoped
identities. File/folder IDs are GUIDs; list-item IDs are numeric within a list,
with a separate native record GUID. Compatible stored handles are accepted by
these commands and supply site/list locators. Folder/list children, attachments,
columns, versions and comments have scoped child references. Versions/comments
link to the parent resource, explicitly marked `url_kind: parent`. Canonical file
path URLs can change on rename/move, while the GUID and local handle remain stable
within one library. Copy creates a different identity. Recycling removes the
active URL; a cached reference is never proof of current existence or permission.
