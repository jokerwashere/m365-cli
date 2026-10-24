# Availability and meeting scheduling

`calendar availability` reads free/busy for multiple people through the existing
Outlook Web session. `calendar find-slots` computes common meeting times from that
snapshot. Both commands are read-only: they never create a meeting, reserve a time,
invite anyone, subscribe to a calendar or change sharing permissions.

## Read availability

```sh
m365 calendar availability \
  --attendee colleague@example.com \
  --attendee another@example.com \
  --start '2026-10-08T08:00:00+02:00' \
  --end '2026-10-09T18:00:00+02:00' \
  --time-zone Europe/Zurich
```

The signed-in account is included as a required attendee by default. Add
`--exclude-self` to query only the people explicitly supplied; that requires at
least one `--attendee`. Use exact SMTP addresses, not display names. Repeat
`--optional-attendee` for optional participants. Addresses are deduplicated
case-insensitively; required status takes precedence over optional status.
Global `--account you@example.com` pins the expected signed-in identity as usual.

The result contains:

- `attendees`: email, required/optional role, lookup status, compact
  `availability_view`, human-readable time `blocks`, and published `working_hours`.
- `availability_legend`: `0` free, `1` tentative, `2` busy, `3` out of office,
  `4` unknown/no data. Unknown is never interpreted as free.
- `complete` and `partial`: whether availability is known for every attendee and
  interval. A valid partial response is successful JSON with explicit gaps.
- `window_start/end`, aligned `query_start/end`, `time_zone`, `interval_minutes`,
  and `observed_at` to make the snapshot's time coverage explicit.

Expected recipient errors, such as an unknown address or inaccessible calendar,
are returned under that attendee's `error.code`; other attendees retain their
valid data. Global authentication, transport and malformed-response errors fail
the command. Server diagnostic text, appointment subjects, locations, bodies and
item IDs are omitted from these results even if Microsoft returns them.

## Find common meeting times

```sh
m365 calendar find-slots \
  --attendee colleague@example.com \
  --optional-attendee observer@example.com \
  --start '2026-10-08T00:00:00+02:00' \
  --end '2026-10-15T00:00:00+02:00' \
  --time-zone Europe/Zurich \
  --duration 45 \
  --work-start 09:00 --work-end 17:00 \
  --weekdays mon,tue,wed,thu,fri \
  --max-results 10
```

Required attendees must be known free throughout a candidate. Tentative time is
blocked unless `--allow-tentative` is explicitly supplied. Busy, out-of-office and
unknown time always block required attendees. Unknown optional attendees do not
block a candidate and are never counted as available. Results rank by most
available optional attendees, then earliest time. Every candidate includes each
attendee's status, `required_available`, optional attendance counts, exact start/
end and duration. `candidate_count` reports all matching grid starts; `has_more`
indicates the result limit. Candidates may overlap; they are alternatives.

Default daily constraints are **09:00–17:00, Monday–Friday, UTC**. Always choose
`--time-zone` explicitly for local scheduling. `--weekdays all` includes weekends.
Daily hours are caller-selected constraints, not an assumption about colleagues'
working patterns. Overnight daily work windows are not supported. Start/end must
include UTC offsets; the display/daily-hours zone may be an IANA name or recognized
Microsoft name. Candidate generation uses UTC instants, retaining correct offsets
and distinct repeated hours across daylight-saving changes.

`--respect-working-hours` additionally requires every required attendee to be
inside their published working hours. A missing, malformed, overnight or unknown-
timezone working-hours definition blocks that attendee under this option. It does
not affect ordinary free/busy lookup. Custom working-hour zones such as
`tzone://Microsoft/Custom` cannot safely be mapped to an IANA zone from this response. The CLI returns `usable: false` with a reason;
it does not substitute the organizer's timezone. Without this flag, only the
explicit daily scheduling window is enforced. Optional attendees outside their
working hours are not counted as available when the flag is enabled.

`unknown_required_attendees` identifies incomplete required availability or
unusable required working hours. Unknown segments block only the affected times;
known free segments can still produce candidates. An empty `slots` array means no
candidate satisfied the constraints, not that all calendars were necessarily busy.
Inspect completeness, attendee errors and working-hour usability to distinguish
these cases. `booking_submitted` is always false.

## Granularity, bounds and identity

- Windows must be positive and at most 31 days. Maximum 100 distinct requested
  addresses, including the automatic organizer. A one-million-cell budget bounds
  the attendee/time grid. There is no silent batching or truncation.
- `--interval` is 5–60 minutes, default 15. The API query is aligned outward on a
  UTC grid; displayed blocks are clipped to the requested window. A partially busy
  interval remains blocked. Use 5-minute resolution when short gaps matter.
- Meeting duration is 5–1440 minutes. It need not divide the interval: every
  merged interval touched by the meeting is checked. This can conservatively
  exclude a gap whose precise boundaries are finer than the selected granularity.
- `--max-results` is 1–100, default 10. The end of the requested window is exclusive.
- Free/busy blocks and proposed meeting slots are transient snapshots, not managed
  calendar items. They have no stable item ID or native permalink; results explain
  this with `reference_unavailable`. Once a meeting is created, normal event IDs,
  stable handles and URLs apply.

This is scheduling availability, not unrestricted reading or editing of another
person's calendar. Microsoft enforces the signed-in user's existing permissions
and organizational/federated sharing. Resource or distribution-list SMTP addresses
may return service-provided aggregate free/busy; this release does not promise
member-level expansion, room booking policy or cross-tenant availability. No app
registration or extra consent flow is introduced.

## Book separately, after review

Recheck the selected interval immediately before creating the meeting. Availability
can change between lookup and booking; these operations are not atomic. To book,
use the existing `calendar create` command with the selected start/end and intended
attendees, after obtaining authorization for the invitations:

```sh
m365 calendar create --subject 'Project planning' \
  --start '2026-10-08T10:00:00+02:00' --end '2026-10-08T10:45:00+02:00' \
  --attendee colleague@example.com --optional-attendee observer@example.com \
  --request-id 'project-planning-001'
```

The timestamps above are illustrative, not a claim that those people are free.
Creating with attendees sends invitations unless `--save-only` is supplied.
The availability commands never do.

## Prepare a meeting without sending

Use `calendar create ... --attendee EMAIL --save-only` to store attendees and the
agenda for review. To add attendees to an existing personal appointment, use:

```sh
m365 calendar update 'EVENT_ID' --attendee colleague@example.com --save-only --if-change-key 'CHANGE_KEY' --request-id 'prepare-001'
m365 calendar get 'EVENT_ID' --body-format html
m365 calendar send 'EVENT_ID' --if-change-key 'LATEST_CHANGE_KEY' --request-id 'invite-001'
```

`--save-only` accepts the same event text/Markdown/HTML bodies and styles. The
saved item has its normal stable ID and URL in the organizer's calendar; the
default Busy status can affect the organizer's free/busy. It is not stored as a
mail draft. Check `MeetingRequestWasSent`, not `IsDraft`, to establish whether
invitations were sent. The latter can be false in both states.

Initial sending requires explicit authorization. `calendar send` only accepts a
verified never-sent single event with attendees, and refuses subsequent sends.
After sending, use `calendar update --notify`; the CLI does not support silently
staging edits to an already-sent meeting. Save-only rejects unknown state,
non-organizer copies, cancelled events, recurring series/occurrences and deleted
items. A verified unsent meeting can be deleted silently; sent meetings require
`calendar delete --notify` for cancellation. This flow sends no RSVP responses.

## Service contracts

Microsoft documents the underlying [GetUserAvailability model](https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/getuseravailability-operation),
[free/busy view](https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/freebusyview)
and [requested detail levels](https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/requestedview).
The native Outlook Web JSON transport remains undocumented and may change; these
references do not guarantee that transport's stability.
