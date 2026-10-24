# To Do and Planner

`m365 tasks login` reuses the running dedicated browser and waits for Planner/To Do
sign-in. It does not restart Chrome or create another profile. `m365 tasks status`
verifies the account, tenant and task-service credentials. The configured Outlook
account remains the default; global `--account EMAIL` overrides the expected
account for this invocation. Access tokens stay in process memory. MFA, expired
sessions and tenant policy can still require interactive sign-in.

## Choose the correct backend

| Backend | Supported operations |
| --- | --- |
| `todo list` | Discover/read/create/rename/delete custom lists; place a list in a group |
| `todo task` | CRUD, all five statuses, completion/reopening, importance, notes, dates, reminders, categories, recurrence, self-assignment/unassignment, moving |
| `todo step`, `link`, `attachment` | Checklist CRUD/check/uncheck; linked-resource CRUD; file list/upload/download/delete under 3 MB |
| `todo my-day`, `group`, `delta` | My Day membership; group create/read/rename/delete when empty; incremental task changes |
| `planner plan`, `bucket` | CRUD; private user plans or existing group containers; archive/unarchive plans; bucket ordering |
| `planner task` | CRUD, progress, priority, dates, labels, assignments, bucket changes, cross-plan moves, recurrence, native deletion/undo |
| `planner details`, `labels` | Description, checklist and reference maps, preview type, plan label descriptions |
| `planner my-day`, `history`, `comment`, `delta` | My Day, plan history, text comments with editing/deletion, incremental changes |
| `todo category`, `planner label` | Category definitions/colors; fixed Planner slots, assignments and guarded color migration; see [LABELS.md](LABELS.md) |
| `tasks list` | Bounded aggregate reads across To Do lists and Planner tasks assigned to the current user |

Task IDs are opaque. Results include `source`, `resource`, `id`, and `version`
when the service supplies one. Other fields retain their API spelling. Native
To Do responses use `Id`, `Subject`, `Assignments`, etc.; Graph uses `id`, `title`,
`status`, etc. Never infer identity from a title or copy a task between services
when the user intends to change the original.

Graph v1.0 provides core To Do operations and Planner buckets/details. The current
Planner task/plan adapter uses Graph beta to retain recurrence, My Day and private
user-container information. Native To Do supplies move, assignment, groups and
My Day; native Planner supplies comments, cross-plan moves and deletion/undo.
There is no automatic adapter fallback after a failed write.

## Routine commands

```sh
m365 tasks status
m365 tasks list --source all --limit 50
m365 todo list list
m365 todo list create --name 'Project actions' --request-id 'project-list-001'
m365 todo task create --list 'LIST_ID' --title 'Review proposal' --body 'Read the revised scope.' --request-id 'task-001'
m365 todo task list --list 'LIST_ID' --query 'proposal'
m365 todo task get 'TASK_ID' --list 'LIST_ID'
m365 todo task update 'TASK_ID' --list 'LIST_ID' --status inProgress --if-version 'VERSION' --request-id 'task-progress-001'
m365 todo task complete 'TASK_ID' --list 'LIST_ID' --request-id 'task-complete-001'
m365 todo task reopen 'TASK_ID' --list 'LIST_ID' --request-id 'task-reopen-001'
m365 todo task assign 'TASK_ID' --list 'LIST_ID' --request-id 'task-self-001'
m365 todo task unassign 'TASK_ID' --list 'LIST_ID' --assignment 'ASSIGNMENT_ID' --request-id 'task-unassign-001'
m365 todo task move 'TASK_ID' --list 'SOURCE_LIST_ID' --to-list 'DESTINATION_LIST_ID' --request-id 'task-move-001'
m365 todo my-day add --task 'TASK_ID' --list 'LIST_ID' --date '2031-02-03' --request-id 'myday-001'
m365 todo my-day remove --task 'TASK_ID' --list 'LIST_ID' --request-id 'myday-remove-001'
m365 todo step create --list 'LIST_ID' --task 'TASK_ID' --name 'Check costs'
m365 todo step check 'STEP_ID' --list 'LIST_ID' --task 'TASK_ID'
m365 todo attachment add --list 'LIST_ID' --task 'TASK_ID' --file ./brief.pdf
m365 todo attachment get 'ATTACHMENT_ID' --list 'LIST_ID' --task 'TASK_ID' --output ./download.pdf
m365 planner plan create --title 'Private project' --request-id 'plan-001'
m365 planner bucket create --plan 'PLAN_ID' --name 'Next'
m365 planner task create --plan 'PLAN_ID' --title 'Review proposal' --request-id 'planner-task-001'
m365 planner task assign 'TASK_ID' --user self --request-id 'planner-assign-001'
m365 planner task update 'TASK_ID' --progress 50 --if-version 'VERSION'
m365 planner task complete 'TASK_ID'
m365 planner task list --plan 'PLAN_ID' --query 'proposal'
m365 planner details get --task 'TASK_ID'
m365 planner details update --task 'TASK_ID' --body-file notes.txt --if-version 'DETAILS_VERSION'
m365 planner comment create --task 'TASK_ID' --body 'Reviewed.' --request-id 'comment-001'
m365 planner comment update 'COMMENT_ID' --task 'TASK_ID' --body 'Reviewed, including costs.'
m365 planner comment delete 'COMMENT_ID' --task 'TASK_ID'
m365 planner task move 'TASK_ID' --to-plan 'DESTINATION_PLAN_ID' --to-bucket 'DESTINATION_BUCKET_ID'
```

To Do assignment is deliberately self-only in this release. Planner accepts
`self` or an exact Entra user UUID; it does not invite that user or grant access.
Assignment, comments,
shared-list changes, labels, plan archive and cross-plan moves can affect other
members or trigger service notifications: choose targets with that effect in mind.
Default plan creation is private. `--container group --group-id UUID` uses an
existing Microsoft 365 group; the CLI never creates/deletes the group itself.
User-plan discovery is not discovery of every shared plan. Use an explicit plan
ID or `planner plan list --group-id UUID` for a known group.

## Structured updates

Creation/update commands accept `--data JSON` or `--data-file FILE` (`-` for UTF-8
stdin). Payloads are typed and allowlisted locally; unknown fields are rejected.
Flags and JSON may be combined only when they do not specify the same field.
An omitted field is unchanged. JSON `null` clears a nullable date/reminder field;
Planner recurrence stops with `{"schedule":null}`. A dry run validates local data
without authenticating; it cannot check permissions or current versions.

To Do task example:

```json
{
  "importance": "high",
  "categories": ["Project"],
  "startDateTime": {"dateTime": "2031-02-03T09:00:00", "timeZone": "Europe/Zurich"},
  "dueDateTime": {"dateTime": "2031-02-03T17:00:00", "timeZone": "Europe/Zurich"},
  "isReminderOn": true,
  "reminderDateTime": {"dateTime": "2031-02-03T16:00:00", "timeZone": "Europe/Zurich"},
  "recurrence": {
    "pattern": {"type": "daily", "interval": 1},
    "range": {"type": "noEnd", "startDate": "2031-02-03"}
  }
}
```

To Do body JSON is `{"contentType":"text","content":"Notes"}` or `html`.
`--body`/`--body-file` are plain text. Importance is `low`, `normal`, `high`;
statuses are `notStarted`, `inProgress`, `completed`, `waitingOnOthers`, `deferred`.
Dates need an IANA timezone. Link fields are `webUrl`, `applicationName`,
`displayName`, `externalId`; step fields are `displayName` and `isChecked`.

Planner task example:

```json
{
  "bucketId": "BUCKET_ID",
  "percentComplete": 50,
  "priority": 3,
  "startDateTime": "2031-02-03T08:00:00Z",
  "dueDateTime": "2031-02-03T16:00:00Z",
  "appliedCategories": {"category1": true},
  "recurrence": {
    "schedule": {
      "pattern": {"type": "daily", "interval": 1},
      "patternStartDateTime": "2031-02-03T08:00:00Z"
    }
  }
}
```

Progress is 0–100; priority 0–10. Planner dates require an explicit UTC offset.
`orderHint` is accepted for buckets/tasks. Assignment maps are keyed by user UUID;
values use `{"@odata.type":"#microsoft.graph.plannerAssignment","orderHint":" !"}`,
or `null` to unassign. `planner task assign/unassign` builds those values for you.
Labels use `category1` through `category25`.

`planner details update --task TASK --data-file details.json` accepts:

```json
{
  "description": "Project notes",
  "previewType": "checklist",
  "checklist": {
    "33333333-3333-4333-8333-333333333333": {
      "@odata.type": "#microsoft.graph.plannerChecklistItem",
      "title": "Review costs",
      "isChecked": false,
      "orderHint": " !"
    }
  },
  "references": {
    "https%3A//example%2Ecom/brief": {
      "@odata.type": "#microsoft.graph.plannerExternalReference",
      "alias": "Brief",
      "type": "Other",
      "previewPriority": " !"
    }
  }
}
```

Use a new UUID per checklist item. To check/update an existing entry, send its
key with the type and changed fields; `null` removes it. Reference keys use Graph's
encoded URL form (encode `.`, `:`, `%`, `@`, `#`); values of `null` remove a link.
The details object has its own version, independent of the task. Plan label names
use `planner labels update --plan PLAN --data '{"categoryDescriptions":{"category1":"Urgent"}}'`.

## Versions, recurrence and deletion

Every mutation supports `--dry-run` and `--request-id`. The local operation ledger
binds task writes to account, tenant, service and resolved payload. It never retries
an unfinished write. Inspect `m365 operation get REQUEST_ID` and the affected
service after `write_outcome_unknown` or `partial_completion`.

Graph task/list updates use the current ETag and optionally compare the reviewed
`--if-version`. Planner updates/deletes also use ETags, but Microsoft can accept a
stale version when changes affect different fields; this is not a general strict
compare-and-swap guarantee. Native extensions and some child resources have no
supported atomic version guard and reject unsupported `--if-version` options.
**To Do deletion ignores If-Match and therefore rejects `--if-version`.** Built-in
To Do lists cannot be renamed/grouped/deleted by this CLI. Deleting a custom list
or plan deletes its contained tasks. Task deletion is not mailbox soft deletion.

Completing a recurring To Do task can advance the original ID to the next active
occurrence and give the completed occurrence another ID. Other layouts may return
a new active ID. Re-list and use status/dates to identify occurrences; do not
assume the original ID is completed. Planner completion supplies
`recurrence.nextInSeriesTaskId` when it creates a successor.

Deleting an active recurring Planner task requires explicit intent:

```sh
m365 planner task delete 'TASK_ID' --recurrence occurrence
m365 planner task delete 'TASK_ID' --recurrence stop-series
```

`occurrence` may generate another task. `stop-series` stops that active schedule
and then deletes the current task; historical occurrences remain. The two writes
are not atomic. A failure after stopping reports `partial_completion`, preserves
the operation as pending, and does not restart the series automatically.

For a non-recurring Planner task, `delete --allow-undo` selects the native adapter
and returns `operation_id` when supplied by Microsoft. Restore with
`planner deletion undo --task TASK_ID --operation-id OPERATION_ID`. Undo is
service-dependent, may expire, and is independently verified by reading the task.
It is not a general recycle-bin interface. Do not combine native deletion with
a Graph `--if-version`. No dependable To Do restore is implemented.

Plan archive/unarchive uses a justification. Use a distinct `--reason` when
reversing the archive state; the default reasons already differ. Reusing the
previous justification is rejected before writing. See the
[Graph unarchive contract](https://learn.microsoft.com/en-us/graph/api/plannerplan-unarchive?view=graph-rest-beta).

To Do moves can change the ID and remove assignment; use the returned ID/list.
My Day aggregate views may lag an acknowledged per-task change. A successful
per-task readback is not proof that every aggregate view has caught up.

## Search, pagination and synchronization

`--query` is a case-insensitive **local title substring**, never Microsoft's
ignored `$search`. It does not search notes. Listings follow server pages with
`--max-pages` (default 20, maximum 100), deduplicate IDs and enforce `--limit`
locally even when Planner ignores `$top`. A collection scan is limited to 10,000
items. `next_offset` refers to the scanned snapshot; rerunning can see changed
ordering. Check `truncated`, `has_more` and `coverage` before claiming completeness.

`tasks list` scans at most `--max-lists` To Do lists (default 50, maximum 100) and
up to 500 matching tasks per list/source in its aggregate view. It reports partial
results and per-list errors explicitly; use a backend-specific listing for deeper
pagination. Its Planner scope is **assigned to me**, not every task in the tenant.

```sh
m365 todo delta list --list 'LIST_ID'
m365 todo delta list --list 'LIST_ID' --cursor 'NEXT_CURSOR'
m365 todo delta list --list 'LIST_ID' --cursor 'DELTA_CURSOR'
m365 planner delta list
m365 planner delta list --cursor 'NEXT_OR_DELTA_CURSOR'
```

Delta commands return one server page. Follow `next_cursor` until `delta_cursor`,
then persist that final cursor for the next synchronization. Cursors are bound to
account, tenant and collection. Treat them as private state. Delta entries are
ordered partial changes, possibly with repeated IDs and `@removed` tombstones;
apply them in order rather than deduplicating them like a normal list. Planner's
beta delta contains mixed resource types, not only tasks. Token invalidation can
require an explicit fresh synchronization; the CLI does not silently reset it.

## Deliberate limits

Premium scheduling/dependencies/custom fields, sharing links/membership changes,
roster administration, large/resumable uploads, copy preservation, To Do ungrouping
and other-person To Do assignment are unsupported. Group-container and
other-person Planner operations require the existing service permissions and do
not grant access. These limits are separate from a tenant refusing access.
Maintainers can re-check compatibility with the [maintainer workflow](MAINTAINER.md).

## Task references and sharing

Create, read, update and assignment results include `stable_id` and `web_url`. Planner task URLs include the plan and tenant, and To Do tasks open their details pane. Steps, comments and other child entries label links to the parent task. A CLI To Do move retains its stable reference despite a changed Microsoft ID. See [REFERENCES.md](REFERENCES.md); links require the recipient to already have access.
