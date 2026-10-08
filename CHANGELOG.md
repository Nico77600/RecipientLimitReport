# Changelog — Recipient Limit Report

All notable changes are listed here. Versions follow MAJOR.MINOR.PATCH (see the guide, Annex E).
Author: Nicolas Fabert.

## [1.0.1] — 2026-10-08

### Fixed
- **Time announced before the count step**: it assumed one route per message still to count, although a
  message over the limit before expansion then rebuilds its recipient list in the same run, about one route
  per recipient. On the 7-day lab run, the console announced about 1 min 55 s for 33 counts; they took
  32 min (604 routes). The console now gives a range: one route per count at least, every route the
  recipient lists may need at most (bounded by `Counting.MaxRoutesPerMessage`), with the time of each. New test.

### Documented
- Guide, chapter 11 and Annex C: the 7-day lab run from an empty database (2026-09-28 → 2026-10-05 18:00):
  3 h 03 min, 13.2 million rows, 799,446 messages, 2,727 requests without any 429, 400,139 messages over 25,
  report written in 43 s.

## [1.0.0] — 2026-10-05

First version. Rewrite of the scripts `Invoke-DeliveredEnvelopeAuditReport.ps1` and
`Invoke-RecipientLimitBlockReport.ps1` (package *Scripts-AuditRecipientLimit*, 2026-07-15) on the
standard of **Purview DLP Report 2.1.1**: same entry point, configuration file, console, SQLite
history, CSV/HTML report, guide and tests. The report is the one of Purview DLP Report; only the
source changes: the Exchange Online **message trace** (Microsoft Graph) instead of Activity Explorer.

### Added
- A separate **User guide** (`docs/RecipientLimitReport-UserGuide.md` and `.html`) for prerequisites,
  one-time setup, everyday commands, scheduled collection, results and troubleshooting; the existing
  `RecipientLimitReport-Guide` is now explicitly the administrator/developer guide.
- **One entry point** `Invoke-RecipientLimitReport.ps1` with three modes: `Report` (default: collects and
  counts what is missing for the period, then writes the report), `Collect` (scheduled task) and `Status`
  (day-by-day content of the database). Periods `Last24Hours`, `Last7Days`, `Last30Days`, `PreviousMonth`,
  `Month`, `Day`, `Custom`. Exit codes 0 / 1 / 2.
- **Message trace collection** (Graph v1.0 `admin/exchange/tracing/messageTraces`): slices of
  `SliceHours` that never cross local midnight, parallel requests within a rolling quota shared by the
  workers and saved between runs (90 requests / 5 min by default), 429 / 401 / 5xx handled, every page
  stored at once with the part of the slice it covers (the API returns the newest rows first), so an
  interrupted run restarts where it stopped. A failed slice is requested once more for its missing part.
- **Optional sender scope** `Target.SenderDomains`: one query per accepted domain (`*@domain` is applied by
  Microsoft), so the incoming Internet mail is not read.
- **Recipient count before distribution list expansion**, as Exchange Online (`RecipientLimits`) and the DLP
  condition `RecipientCountOver` count it — a distribution list is one recipient:
  - message without distribution list: the recipients of the message trace;
  - message with distribution lists (status `expanded`), over the limit after expansion: `RcptCount` of the
    **Submit** event of the route of a list (`getDetailsByRecipient`), one request per message, kept in the
    database; `Receive` event for a message received by SMTP; a route without any count is reported in the
    console and the summary, never guessed. Separate quota (`Counting` section), retried at the next run.
- **Recipient list before distribution list expansion**, as in the DLP report: for a message over the limit
  sent to lists, the route of each recipient tells whether the sender addressed it (its route holds the
  Submit event, or Receive for SMTP) or the expansion of a list added it. Routes are read until as many
  recipients as the count are found — the lists first, the addresses already seen as members of the same
  lists last — at most `Counting.MaxRoutesPerMessage` (250) per message; above, the report keeps the
  addresses after expansion for that message and says so. Kept in the database (`recipient.direct`,
  `message.list_state`), resumed at the next run.
- **SQLite history**: one row per message trace ID and one per recipient. Once a period has settled
  (`SettlingHours`), the messages at or under the limit are removed: the database keeps the messages over
  the limit (and those still to count), not the whole message trace. Retention 180 days, lock file, runs
  table, schema version.
- **Report = Purview DLP Report**: same CSV columns (time, sender, recipients, subject, recipient count,
  Message ID) plus *Distribution lists*; one row per Message ID; same self-contained HTML page (tiles,
  distribution, Top 10 senders, filters, virtual scrolling, export of the view, details with the
  distribution lists flagged and the count source); splitting by week / day / rows above `MaxRowsPerFile`.
  Long recipient lists are cut at `MaxRecipientsListed` addresses.
- **Authentication**: certificate (app-only, client assertion built without any module), client secret
  (environment variable or prompt), interactive (MSAL). The tenant and the permission
  `ExchangeMessageTrace.Read.All` are checked in the token before any request.
- Pester tests with an in-memory message trace API (`tests\FakeGraph.cs`), guide (Markdown + HTML),
  package builder.

### Validated in the lab (2026-10-05)
- DLP pilot window (141 messages, 10 sent to distribution lists): **same 141 Message IDs and same recipient
  counts** as Purview DLP Report, distribution lists included (for example 28 recipients before expansion
  for a message with 55 addresses in the trace).
- **Same recipient lists** as Purview DLP Report for the 10 messages sent to distribution lists (26 to 35
  recipients, 31 to 90 addresses in the trace): rebuilt from 581 routes in 31 min, 10 / 10 identical.
- The 48-hour comparison window of Purview DLP Report (2026-09-27 14:24 → 2026-09-29 14:24 UTC):
  **all 149,802 messages of the DLP report, with the same recipient count**, plus 352 messages missing from
  Activity Explorer (345 delivered without any DLP evaluation, 7 blocked by the rule without an Activity
  Explorer event). 4.96 million trace rows read at about 90,000 rows per minute (quota-bound, one 429
  absorbed); database 150 MB; report of 150,154 messages written in 19 s.

### Changed (compared with the scripts of 2026-07-15)
- Every message is reported whatever its delivery status (delivered, failed, quarantined...), like the DLP
  report; the blocked-message report (`550 5.5.3`) is not part of this version.
- The application settings come from the configuration file (`state\graph-app.json` is no longer used);
  Microsoft Graph **v1.0** instead of beta.
