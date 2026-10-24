# Categories and Planner labels

`m365 category` manages the mailbox's Outlook master categories, also used by To Do.
`m365 todo category` is an alias. `m365 planner label` manages labels in one plan.
These use the existing saved browser login; no application registration is needed.

| Capability | Outlook / To Do categories | Planner labels |
| --- | --- | --- |
| List/read | Mailbox catalog; optional name filter | All 25 slots, including unnamed slots |
| Create | Unique name and optional color | Name an unused color slot |
| Rename | Change definition and submit rename on existing mailbox items | Change the slot name for the entire plan |
| Color | Choose a preset, or no color | Move name and task assignments to an unused color slot |
| Delete | Delete definition; item category names remain | Clear slot name; refuse if applied unless `--remove-from-tasks` |
| Apply/remove | One mail item, calendar occurrence or To Do task | One task in the specified plan |
| IDs/links | Provider ID plus stable local handle; Outlook application link | Plan-scoped slot ID plus stable handle; parent-plan link |

## Outlook and To Do

```sh
m365 category colors
m365 category list
m365 category list --query 'Project'
m365 category get 'CATEGORY_ID'
m365 category create --name 'Project Atlas' --color sky-blue --request-id 'atlas-category-create'
m365 category update 'CATEGORY_ID' --color dark-red --if-version 'CATEGORY_VERSION'
m365 category update 'CATEGORY_ID' --name 'Atlas follow-up' --request-id 'atlas-category-rename'
m365 category delete 'CATEGORY_ID' --if-version 'CATEGORY_VERSION'
```

IDs come from the catalog results. Category names are unique ignoring case, trimmed,
nonempty, at most 255 UTF-16 code units, and cannot contain commas, semicolons or
control characters. Create defaults to `none`. Colors are the native Outlook Web
names returned by `colors`, `none`, or `preset0` through `preset24`; conventional
`red`, `orange`, `yellow`, `green`, `blue` and `purple` aliases are accepted. Web and
desktop clients can name/render a preset differently. Custom RGB/hex values are
not supported. The newer web enum value 25 was rejected by the live server and is
not exposed. `colors` is local and does not require a signed-in browser.

Rename uses the same two native operations as Outlook Web: replace the catalog
entry, then submit a mailbox-wide category rename. This affects existing items
bearing the old category name, including mail, calendar and To Do. The result says
`item_assignments: "rename_submitted"` and `propagation_complete: false`: the
service acknowledgement cannot prove every item has finished updating. Use
`--definition-only` to leave existing item names unchanged. Color-only updates
need no assignment migration because items store category names.

Microsoft changes the provider ID on rename. The CLI returns `previous_id` and
preserves the category's `stable_id` in the private reference registry. Resolve a
stored handle with `m365 reference get STABLE_ID`; subsequent commands take the
latest provider ID. Renames performed outside this CLI cannot be reliably aliased.

Catalog `version` is a review hash. `--if-version` checks it before dispatch, but
Outlook exposes no validated atomic catalog precondition. Results explicitly say
`precondition_atomic: false`; a concurrent edit between read and write is still
possible. Definition writes require an acknowledgement **and** matching readback.
If item propagation fails after the definition changed, `partial_completion`
includes the changed category and its reference. Inspect it before another write.

### Apply and remove without replacing other categories

```sh
m365 category apply 'CATEGORY_ID' --mail 'MESSAGE_ID' --if-version 'CHANGE_KEY'
m365 category remove 'CATEGORY_ID' --mail 'MESSAGE_ID'
m365 category apply 'CATEGORY_ID' --event 'EVENT_ID'
m365 todo category apply 'CATEGORY_ID' --task 'TASK_ID' --list 'LIST_ID' --if-version 'TASK_VERSION'
m365 todo category remove 'CATEGORY_ID' --task 'TASK_ID' --list 'LIST_ID'
```

These preserve the item's other categories and use its current change key/ETag to
reject races. For these commands `--if-version` refers to the **target item**,
not the category definition. Mail updates work on drafts and received/sent mail.
Calendar categories are personal metadata and send no invitations or updates;
non-organizers may categorize their own copy. Recurring masters remain blocked;
choose an occurrence. To Do assignments require both Outlook and task credentials
for the same pinned account. Native To Do/Planner unrelated features are unchanged.

Deleting a category definition does **not** strip its name from existing items.
To remove such an orphaned name, use the explicit name selector:

```sh
m365 category remove --name 'Old category name' --mail 'MESSAGE_ID'
m365 todo category remove --name 'Old category name' --task 'TASK_ID' --list 'LIST_ID'
```

Choose either the ID or `--name`, exclusively. Name-based removal cannot create a
new category or apply an unknown name. Existing `todo task update --data` and
mail/calendar category arrays still replace the entire provided array; use the
commands above for additive/removal changes. There is no unverified mailbox-wide
strip/delete operation exposed. Keyboard shortcuts, favorites and contact category
assignment are outside this labels surface.

## Planner

```sh
m365 planner label colors
m365 planner label list --plan 'PLAN_ID'
m365 planner label get category1 --plan 'PLAN_ID'
m365 planner label create --plan 'PLAN_ID' --name 'Urgent' --color pink
m365 planner label rename category1 --plan 'PLAN_ID' --name 'Needs review' --if-version 'PLAN_DETAILS_ETAG'
m365 planner label apply category1 --plan 'PLAN_ID' --task 'TASK_ID' --if-version 'TASK_ETAG'
m365 planner label remove category1 --plan 'PLAN_ID' --task 'TASK_ID'
m365 planner label recolor category1 --plan 'PLAN_ID' --color blue --request-id 'urgent-recolor'
m365 planner label delete category5 --plan 'PLAN_ID' --remove-from-tasks --request-id 'urgent-delete'
```

There are 25 fixed slots (`category1`–`category25`) per plan. Their colors cannot
be edited. `create` names the chosen color's unused slot. `rename` affects everyone
using that plan. Applying requires a named slot and a task belonging to that plan;
removing also works for an unnamed slot. Other task labels are preserved.

`recolor` implements a move: pink (`category1`) to blue (`category5`), for example.
It reads every task in the plan, rejects a destination that is either named **or
already applied**, switches the affected tasks, then changes the two slot names.
It returns the destination label, `previous_slot`, the source slot reference and
`updated_task_count`. Stable references identify fixed slots, so this move produces
a **different** stable label ID. Store the returned destination reference.

`delete` clears a name, not the built-in color slot. If used by any tasks, it
requires `--remove-from-tasks`, which removes that slot from those tasks first.
Create, recolor and delete require a complete task scan. `--max-pages` defaults to
20 (range 1–100), with the shared 10,000-item budget; truncation or duplicates
refuse writes. The operation sees a snapshot: another client may add/change tasks
later. Graph offers no atomic transaction over all tasks and plan details.

Writes use ETags for every affected task and the plan details. A failure after
some updates returns `partial_completion`, completed task IDs and the cause.
There is no automatic retry or rollback. If the first write has an unknown
outcome, inspect that task and the plan before deciding how to repair it. The
operation journal stays pending after failure. Task membership updates and
snapshot scans cannot guarantee protection against every concurrent client.

The original `planner labels get/update --plan ... --data ...` commands remain
available for direct description-map edits. They do not migrate task assignments;
prefer the singular `planner label` commands for lifecycle operations.

## References

Every managed category/label result contains `id`, `stable_id`, `web_url` and
link-scope metadata. The CLI does not expose a standalone label/category permalink. Outlook categories link to Outlook (`url_kind: "application"`);
Planner slots link to their plan (`url_kind: "parent"`). Links do not grant access.
Share a task's item link when assigning work, rather than a label's parent link.

Outlook category operations use native Outlook Web protocols. Planner labels use
Graph plan-details and task APIs through the existing task session. See the
[maintainer guide](MAINTAINER.md) for compatibility tools.

Microsoft documents the [shared Outlook category model](https://learn.microsoft.com/en-us/graph/api/resources/outlookcategory?view=graph-rest-1.0),
[Planner category descriptions](https://learn.microsoft.com/en-us/graph/api/resources/plannercategorydescriptions?view=graph-rest-1.0)
and [Planner slot/color mapping](https://learn.microsoft.com/en-us/connectors/planner/).
Those documents describe the model; they are not a stability promise for the native
Outlook Web transport. Responses are checked conservatively after writes.
