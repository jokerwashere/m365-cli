# Maintaining Microsoft API compatibility

These commands help maintainers check service compatibility and disable affected
operations when Microsoft changes an API. They are part of the installed `m365`
package. Ordinary users and agents do not need to run capability sweeps.

## Inventory and read probes

```sh
m365 maintainer capabilities
m365 maintainer check --module all --output baseline.json
m365 maintainer check --module all --baseline baseline.json --output current.json
```

The inventory comes from the CLI parser and identifies commands, modules,
mutation status and suite assignments. It describes the implementation, not the
current availability of every operation in a particular tenant.

`check` authenticates and checks representative read contracts. It does not issue
mail, meeting, messaging or task writes; loading the web apps can still cause
normal background traffic. A missing sample is reported as `skipped`; an empty
collection cannot establish an item-field contract. A failed probe returns a
nonzero exit status and `capability_check_failed`, while still saving the report.
Read probes do not establish write behavior or coverage of every command option.

Reports contain statuses, counts, error codes and an account fingerprint for
baseline comparison. They omit message bodies and authentication material.
Baseline comparisons identify changed availability and contract revisions;
they do not guarantee API compatibility. Store reports outside the repository.

## Explicit write checks

Use `m365 maintainer validate --help` to choose a suite and inspect its arguments.
Obtain authorization for the account, destinations, test traffic and cleanup
before enabling writes. The execution flag does not replace that authorization.

```sh
m365 --account you@example.com maintainer validate --suite tasks --allow-test-writes
```

Suites create labelled disposable items and may send messages, invitations or
updates. Suites that send email require an authorized `--recipient`; Teams
channel checks require an explicitly authorized `--channel`; SharePoint checks
require an authorized personal OneDrive `--site`. Read the help before running a
suite. Cleanup cannot recall copies already delivered to other people.

Each run records command coverage, skipped checks, outcome and cleanup in a
private journal under `M365_HOME`. It may refuse a new run when an earlier write
or cleanup remains unresolved. Inspect that journal and `m365 operation get
REQUEST_ID`, reconcile the actual service state, then remove only the labelled
fixtures owned by that run. Never bypass a pending operation with a new request
ID or an automatic write retry. Keep journals and generated reports private.

## Disable affected commands

```sh
m365 maintainer disable 'todo task update' --reason 'API response changed; repair pending'
m365 maintainer disabled
m365 maintainer enable 'todo task update'
```

Quote the command prefix. Disabling a prefix blocks it and its descendants for
this local state directory; aliases share the applicable checks. It does not
change Microsoft permissions or other installations. Re-enable an operation only
after its contract and safeguards are understood and checked.

## Repair workflow

1. Record the command, safe error code and contract change. Keep raw service data,
   browser state, tokens and headers out of Git and public issues.
2. Disable the affected command locally. Inspect service state before attempting
   recovery from any unknown write outcome.
3. Update the adapter conservatively. Preserve identity checks, acknowledgement
   checks, version preconditions, command locking and the no-write-retry rule.
   Add synthetic regression coverage for the changed behavior.
4. Run the offline checks and build in [CONTRIBUTING.md](../CONTRIBUTING.md).
   Regenerate [CLI.md](CLI.md) when help changes and update the usage guides.
5. Use authorized read and disposable-item checks where needed. A skipped or
   synthetic check does not establish remote behavior. Keep results private and
   document current user-facing behavior and limitations in the public guides.
