---
title: Recipient Limit Report
subtitle: Developer guide
version: 1.0.1
author: Nicolas Fabert
updated: 2026-10-08
---

# Recipient Limit Report — Developer guide

> Reports the Exchange Online messages sent to **more than 25 recipients** — counted as Exchange Online counts them, a **distribution list being one recipient** — as CSV and HTML files, **one row per Message ID**. The source is the Exchange Online **message trace**; the report is the same as the one of Purview DLP Report. The commands used every day are in the [user guide](RecipientLimitReport-UserGuide.md).

> [!IMPORTANT]
> Files downloaded from the Internet may be blocked by Windows and fail to run. Before using this project, unblock every file in the downloaded folder:
>
> ```powershell
> Get-ChildItem "C:\Chemin\Du\Dossier" -Recurse -File -Force | Unblock-File
> ```
>
> Replace the example path with the folder where you downloaded or extracted this project.
>
> The `Install-Module` commands in this documentation use `-Force`, so they also update or reinstall a module that is already installed. If an older version still conflicts, close every PowerShell window, open a new one (as administrator for `-Scope AllUsers`), run `Uninstall-Module <ModuleName> -AllVersions -Force`, then run the `Install-Module` command again.

```cards
target | What it answers | Who sends mail to more than 25 recipients, to whom, about what, and when — what a recipient limit of 25 would block.
download | Where the data comes from | The Exchange Online **message trace** (Microsoft Graph): every recipient of every message, 90 days.
database | What it keeps | A local **SQLite** history of the messages over the limit, beyond the 90 days of the message trace.
file | What it produces | **CSV + HTML** files in a local folder — the same report as Purview DLP Report.
```

## Quick start

```steps
Check the prerequisites | PowerShell 7.4+, an application with the Microsoft Graph permission `ExchangeMessageTrace.Read.All` and a certificate (chapter 7).
Edit the configuration | Open `config\RecipientLimitReport.config.psd1`: tenant, application, sender domains.
Run a first report | `.\Invoke-RecipientLimitReport.ps1` — the last 7 days, collected, counted, then written to `reports\`.
Schedule the daily collection | `.\Invoke-RecipientLimitReport.ps1 -Mode Collect` every day, to keep the history beyond 90 days and keep the reports fast.
```

> [!IMPORTANT]
> The tool is **read-only** for Microsoft 365: it never sends e-mail and never changes a setting. Reports stay on the local disk; the administrator sends or shares them.

# Part I · Understand

<!-- icon: book -->
## 1. Project background

**Why this project**

- To reduce the volume of e-mail, the organisation wants to **block the messages sent to more than 25 recipients**, except for the users of an **exception list**.
- Two ways of blocking are possible: a Microsoft Purview **DLP policy** (condition *recipient count over 25*), or a **recipient limit per mailbox** in Exchange Online (`Set-Mailbox -RecipientLimits`).
- **Before blocking**, the business lines must see which messages are concerned, and the exception list must be prepared from facts.

**Why a second report, from the message trace**

*Purview DLP Report* reads the matches of a DLP policy in audit mode. This tool produces **the same report without any DLP policy**, from the message trace:

```cards
shield | Independent of DLP | No policy to create or keep in audit mode: the message trace records every message anyway.
target | Works with RecipientLimits | A message over `RecipientLimits` is refused **before** the DLP rules run (lab, 2026-10-04): the DLP report no longer sees it, the message trace still does.
compare | Same numbers | Same count as Exchange Online and DLP — a distribution list is one recipient. Lab: same Message IDs and counts as the DLP report.
```

**Constraints**

| Constraint | Value |
|---|---|
| Expected volume | **30,000 to 50,000 messages over the limit per day** — and the message trace lists **every** message of the tenant (chapter 11) |
| Periods requested | last 24 hours, last 7 days, a month |
| Delivery | local files; sent by e-mail or placed on a share / OneDrive by the administrator |
| Quota of the API | 100 requests per 5 minutes for the **whole tenant**, shared with every other tool |

<!-- icon: flow -->
## 2. How it works

The tool works in **three stages**. The database sits in the middle.

```flow
download | Message trace | every recipient, 90 days
arrow | Collect | every day, or before a report
database | SQLite database | local history, no duplicates
arrow | Count | messages sent to lists only
people | Count before expansion | Submit event, once per message; then the route of each recipient for the list
arrow | Report | on demand, any period
file | CSV + HTML | one row per Message ID
```

| Stage | What happens | When |
|---|---|---|
| **1 · Collect** | Reads the message trace (`messageTraces`, one row per message and recipient) and stores it in the database. | Every day (scheduled task), and automatically before a report if something is missing. |
| **2 · Count** | For each message **over the limit after expansion** and sent to at least one **distribution list**, reads the number of recipients **before expansion** from the route of the message (`getDetailsByRecipient`). For a message still over the limit, rebuilds its **recipient list as the sender addressed it** from the route of each recipient. | Right after the collection. One request per message for the count, one per recipient for the list; kept in the database. |
| **3 · Report** | Reads the database and writes one CSV and one HTML file for the requested period. | On demand. |

**How the recipients are counted**

The message trace shows the recipients **after** expansion: a distribution list appears once with the status `expanded`, then each of its members appears as a recipient. Exchange Online (`RecipientLimits`) and the DLP condition count the recipients **before** expansion: the list counts as **one**.

| Message | What the trace shows | Count used |
|---|---|---|
| No distribution list | One row per recipient | The **rows of the trace** — every recipient was addressed by the sender. No request. |
| One or more distribution lists | The lists (`expanded`) and every member | **`RcptCount` of the Submit event** of the message — read once from the route of a list. |

Example from the lab: a message sent to 26 people and 2 lists shows **55** addresses in the trace; its Submit event says **28** — the count of the DLP report, and the count `RecipientLimits` checks.

**How the recipient list is rebuilt** — the trace does not say which addresses the sender typed: a member of a list is a recipient like the others. Its **route** does: the route of a recipient the sender addressed starts with the `Receive` and `Submit` events, the route of a member added by the expansion starts after them. For a message over the limit, the tool reads the route of each address until it has found as many addressed recipients as the count: the report then lists the **28** recipients of the example — the 2 lists and the 26 people — like the DLP report, not the 55 addresses of the trace.

**Why a local database?**

```cards
calendar | Keep the history | The message trace keeps **90 days**. The database keeps months — only the messages over the limit.
refresh | Restart safely | An interrupted collection restarts where it stopped, **without duplicates**: every page is stored with the part of the slice it covers.
clock | Report in seconds | Once a period is collected and counted, its report is written in seconds — no new download.
```

**Columns of the report** — the columns of Purview DLP Report, plus one:

| Column | Content |
|---|---|
| Received time | When Exchange Online received the message, in the report time zone, to the second |
| Sender | Sender address |
| Recipients | Recipient addresses — **optional** (`IncludeRecipientDetails`). As the sender addressed them; for a message sent to lists: the lists first, then the recipients addressed directly (the members added by the expansion are left out). |
| Subject | Message subject |
| Recipient count | Number of recipients **before expansion** |
| Distribution lists | Number of distribution lists the message was sent to (0 for most messages) |
| Message ID | Internet Message ID, the unique key of a message |

<!-- icon: lightbulb -->
## 3. Things to know

> [!WARNING]
> **The message trace keeps 90 days.** A period that was not collected within 90 days is lost for this tool. Schedule the daily collection (chapter 7).

> [!NOTE]
> **The message trace is not real time.** Rows arrive a few minutes to an hour after the message, and statuses change (pending → delivered). The tool therefore collects the **last 6 hours again** at every run (`SettlingHours`). The most recent hours of a report are provisional.

> [!TIP]
> **One row = one message over the limit, whatever happened to it.** Delivered, refused by a DLP rule, quarantined: the message is reported, like in the DLP report. The report shows what a limit **would** block; once the limit is applied, the messages it refuses still appear in the message trace.

> [!IMPORTANT]
> **The whole message trace is read.** The API cannot filter on a number of recipients: to find the messages over 25 recipients, the tool reads every recipient row of the period — about 90,000 rows per minute at most (quota of the tenant). Limit the reading to the senders of the organisation with `Target.SenderDomains`, and size the first collection (chapter 11).

# Part II · Set up

<!-- icon: checklist -->
## 4. Prerequisites

### Server or workstation

| Item | Requirement | Validated in the lab |
|---|---|---|
| Operating system | Windows 10/11 or Windows Server 2016+, x64 or ARM64 | Windows 11 x64 |
| PowerShell | **PowerShell 7.4 or later** — not Windows PowerShell 5.1 | 7.6.6 |
| Console | **Windows Terminal** — the display shown in this guide | Windows Terminal |
| Modules | **None** in certificate and secret modes. Interactive mode: `Microsoft.Graph.Authentication` (for its MSAL library) | — |
| SQLite | Supplied in `lib\sqlite` — nothing to install | SQLite 3.53.3 |
| Disk | The database keeps the messages over the limit with their recipients: about **1 KB per message**, so **1.5 GB per month** at 50,000 messages/day (about **9 GB** with the default retention of 180 days), plus the reports | 150 MB for 150,000 messages |
| Browser | Recent Edge, Chrome or Firefox, to open the HTML report | Edge |
| Network | HTTPS to `graph.microsoft.com` and `login.microsoftonline.com` | — |

```powershell
# Interactive mode only: the MSAL library of the Microsoft Graph PowerShell SDK
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
```

### Permissions

```cards
key | Application (recommended) | An app registration with the Microsoft Graph **application** permission `ExchangeMessageTrace.Read.All`, admin consent, and a **certificate** — chapter 7.
people | Administrator account (interactive) | The **delegated** permission `ExchangeMessageTrace.Read.All` consented for the client application, and an Exchange role that can read the message trace (for example *Global Reader*).
shield | Microsoft message trace service | The tenant must hold the service principal of the Microsoft application **`8bd644d1-64a1-4d4b-ae52-2e0cbf64e373`**, created once (chapter 7). Without it, Graph answers 401.
```

> [!NOTE]
> Reference: [Message trace in the Microsoft Graph API](https://learn.microsoft.com/graph/api/resources/exchangemessagetrace). The permission is read-only: it cannot change anything in Exchange Online.

<!-- icon: download -->
## 5. Installation

```steps
Copy the package | Copy the package folder (`RecipientLimitReport-<version>`, made by `tools\New-RlrPackage.ps1`) to the server, for example `D:\Tools\RecipientLimitReport`.
Unblock the files | `Get-ChildItem D:\Tools\RecipientLimitReport -Recurse -File -Force | Unblock-File` (files copied from the Internet or a share).
Edit the configuration | `config\RecipientLimitReport.config.psd1` — the *Tenant* and *Authentication* values are empty in the package: fill them in, with the *sender domains* (chapter 6).
Check without connecting | `.\Invoke-RecipientLimitReport.ps1 -Mode Status` — checks the configuration and builds the engine (a few seconds, once). On a new installation it ends with *Ready for the first collection*; later it shows the content of the database.
```

**What is in the folder**

| Path | Content |
|---|---|
| `Invoke-RecipientLimitReport.ps1` | **The only script to run** |
| `config\RecipientLimitReport.config.psd1` | All the settings |
| `RecipientLimitReport.psd1` / `.psm1` | PowerShell module used by the script |
| `src\RecipientLimitReport.Engine.cs` | C# engine: requests, database, CSV/HTML writing |
| `templates\Report.template.html` | Look and behaviour of the HTML report (the page of Purview DLP Report) |
| `lib\sqlite\` | SQLite libraries (see `THIRD-PARTY-NOTICES.md`) |
| `docs\RecipientLimitReport-UserGuide.html` · `docs\RecipientLimitReport-Guide.html` | The user guide and this guide |
| `data\` · `reports\` · `logs\` · `bin\` | Created at run time — the package contains **no database**, the first run creates an empty one. **Back up `data\`**: it is the only history beyond 90 days. |

> [!NOTE]
> The package holds only what is needed to run. The git repository of the tool also contains this Markdown guide and the user guide in `package\docs\`, `tests\` (automated tests) and `tools\` (documentation builder, package builder) — see chapters 12 to 14.

<!-- icon: settings -->
## 6. Configuration

Everything is set in **`config\RecipientLimitReport.config.psd1`**. It is a PowerShell data file: text between quotes, `$true` / `$false`, numbers, `@( )` for lists, `#` for comments.

> [!TIP]
> When a value is wrong, the tool lists **all** the problems at once and stops before doing anything. Relative paths (`.\data`) are relative to the tool folder.

### Tenant and target

| Key | Meaning |
|---|---|
| `Tenant.TenantId` | Tenant GUID — a **safety check**: the tool reads the tenant of the access token and stops if it differs. |
| `Tenant.Organization` | Initial domain `xxx.onmicrosoft.com` — display only. |
| `Target.RecipientLimit` | `25`: messages with **more** recipients are reported (26 and more), counted **before expansion**. |
| `Target.SenderDomains` | `@()` = every sender of the message trace. With the accepted domains of the organisation — `@('contoso.com', 'fabrikam.com')` — only the messages **sent by the organisation** are read: one query per domain, filtered by Microsoft (`senderAddress eq '*@contoso.com'`). Recommended in production. |

> [!CAUTION]
> **Lowering `RecipientLimit` later only applies to the data collected afterwards.** Once a period has settled, the database keeps only the messages over the limit in force: messages with 21 to 25 recipients collected with a limit of 25 are no longer in the database if the limit becomes 20. `-Mode Status` warns about it. Raising the limit is always possible.

### Authentication

| Key | Values | Meaning |
|---|---|---|
| `Mode` | `Certificate` · `ClientSecret` · `Interactive` | Application with a certificate (recommended, scheduled task), application with a secret, or an administrator signing in. |
| `AppId` | GUID | Application (client) ID. Interactive: `''` uses *Microsoft Graph Command Line Tools*. |
| `CertificateThumbprint` | 40 hex characters | Certificate mode: certificate with its private key in `Cert:\CurrentUser\My` or `Cert:\LocalMachine\My`. The client assertion is signed locally: no module is needed. |
| `ClientSecretVariable` | `RLR_CLIENT_SECRET` | Secret mode: name of the environment variable holding the secret. When it is empty in an interactive session, the secret is asked (hidden). The secret is never written anywhere. |
| `UserPrincipalName` | account | Interactive: expected account (`''` accepts any account of the tenant). |

### Collection

| Key | Default | Meaning |
|---|---|---|
| `PageSize` | `5000` | Rows per request (1–5000): 5000 = fewest requests. |
| `MaxConcurrency` | `3` | Requests in flight. 2 or 3 already use the whole quota. |
| `SliceHours` | `2` | Longest query slice. Slices never cross local midnight; up to three run in parallel. The quota of the tenant remains the limit; short slices give a finer progress table and a finer restart point. |
| `SettlingHours` | `6` | The last hours are collected again at the next run (late rows, status changes). |
| `BackfillDays` | `7` | How far back `-Mode Collect` looks; only what is missing is collected. |
| `SourceRetentionDays` | `90` | Retention of the message trace. Older missing periods are reported as lost. |
| `RetentionWarningDays` | `7` | Warn when a missing period will be lost within this many days. |
| `MaxRequests` · `PeriodSeconds` | `90` · `300` | Quota used by the tool: 90 requests per 5 minutes, out of the 100 of the tenant. |
| `RequestTimeoutSeconds` · `MaxRetries` | `180` · `5` | Per request: 5xx, timeouts and network errors are retried; 429 waits are separate. |

### Counting

The count before expansion of the messages sent to distribution lists, and the recipient list of those over the limit (`getDetailsByRecipient`, separate quota of 100 requests per 5 minutes).

| Key | Default | Meaning |
|---|---|---|
| `MaxConcurrency` | `3` | Requests in flight. |
| `MaxAttempts` | `3` | Executions that try a message before it is reported as *without a count*. |
| `MaxMessagesPerRun` | `0` | `0` = every message still to count; otherwise the rest waits for the next run. |
| `MaxRoutesPerMessage` | `250` | Requests allowed per message to rebuild its recipient list as the sender addressed it. Above, the report lists the addresses after expansion for that message (with a note). `0` = never rebuild: the count only, one request per message. |
| `MaxRequests` · `PeriodSeconds` | `90` · `300` | Quota used by the tool for this API. |

### Storage

| Key | Default | Meaning |
|---|---|---|
| `DatabasePath` | `.\data\RecipientLimitReport.sqlite` | Database file. **Back it up.** |
| `RetentionDays` | `180` | Messages older than this are deleted at each collection, and the file shrinks (`0` = keep everything). |

### Report

| Key | Default | Meaning |
|---|---|---|
| `DefaultRange` | `Last7Days` | Period used when `-Range` is not given. |
| `TimeZone` | `Europe/Paris` | Time zone of dates, days, weeks and months. |
| `OutputPath` · `FilePrefix` | `.\reports` · `RecipientLimit` | Where and how files are named. |
| `Formats` | `@('Csv','Html')` | Files to write. |
| `IncludeRecipientDetails` | `$true` | `$true`: recipient addresses in CSV and HTML. `$false`: count only — much smaller files. |
| `MaxRecipientsListed` | `500` | Addresses listed per message. A list expanded to thousands of members is cut (`(+N more)`); the count is not affected. `0` = all. |
| `SplitBy` · `MaxRowsPerFile` | `Week` · `500000` | Splitting, used only above `MaxRowsPerFile` (chapter 10). |
| `CsvDelimiter` · `Title` | `;` · *Messages with more than {0} recipients* | `;` opens directly in Excel with French regional settings. `{0}` = the limit. |

### Logging and console

| Setting | Meaning |
|---|---|
| `Logging.Path` · `Logging.RetentionDays` | One log file per day in `.\logs`, kept **30 days** (deleted at each run). |
| Environment variable `NO_COLOR` | Disables colours. `RLR_FORCE_COLOR=1` keeps them when the output is redirected. |

### Command-line overrides

The command line chooses **what to do**. A few report settings can be changed for one execution:

| Parameter | Overrides |
|---|---|
| `-IncludeRecipientDetails` · `-IncludeRecipientDetails:$false` | `Report.IncludeRecipientDetails` |
| `-SplitBy Rows\|Day\|Week` · `-MaxRowsPerFile 200000` | Splitting |
| `-OutputPath D:\Out` | `Report.OutputPath` |
| `-ConfigPath D:\Other.psd1` | Another configuration file |

<!-- icon: clock -->
## 7. Unattended execution

The daily collection keeps the history beyond 90 days and spreads the reading of the message trace over the days. Nobody is there to sign in, so it uses an **application with a certificate**.

### Create the application (once)

```steps
Register the application | Microsoft Entra admin center › **App registrations** › **New registration** — name `Recipient Limit Report - collector`, single tenant. Note the **Application (client) ID**.
Add the permission | **API permissions** › **Add a permission** › **Microsoft Graph** › **Application permissions** › `ExchangeMessageTrace.Read.All`, then **Grant admin consent**.
Create the certificate | On the server, with the account that runs the task — see the command below. Upload the `.cer` file in **Certificates & secrets**.
Create the service principal of the message trace service | Once per tenant — see below. Without it, every request answers *401 … service principal … 8bd644d1-…*; after its creation, allow **up to a few hours** for provisioning.
Fill in the configuration | `Authentication.Mode = 'Certificate'`, `AppId`, `CertificateThumbprint`, `Tenant.TenantId`, `Target.SenderDomains`.
Test by hand | `.\Invoke-RecipientLimitReport.ps1 -Range Last24Hours`.
```

```powershell
# Certificate: private key not exportable, 2 years
$cert = New-SelfSignedCertificate -Subject 'CN=RecipientLimitReport-Collector' -CertStoreLocation Cert:\CurrentUser\My `
    -KeyExportPolicy NonExportable -KeySpec Signature -KeyLength 2048 -NotAfter (Get-Date).AddYears(2)
Export-Certificate -Cert $cert -FilePath .\RecipientLimitReport-Collector.cer
$cert.Thumbprint

# Service principal of the Microsoft message trace service (once per tenant, Application Administrator)
Install-Module Microsoft.Graph.Applications -Scope CurrentUser -Force
Connect-MgGraph -Scopes 'Application.ReadWrite.All'
New-MgServicePrincipal -AppId '8bd644d1-64a1-4d4b-ae52-2e0cbf64e373'
```

> [!TIP]
> The tool checks the token before any request: right tenant, `ExchangeMessageTrace.Read.All` present in the `roles` claim. A missing consent stops it at once with the path to fix it in the Entra admin center.

### Create the scheduled task

```powershell
$action = New-ScheduledTaskAction -Execute 'pwsh.exe' `
    -Argument '-NoProfile -NonInteractive -File "D:\Tools\RecipientLimitReport\Invoke-RecipientLimitReport.ps1" -Mode Collect' `
    -WorkingDirectory 'D:\Tools\RecipientLimitReport'
$trigger  = New-ScheduledTaskTrigger -Daily -At 06:00
$settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Hours 12) -StartWhenAvailable -MultipleInstances IgnoreNew
Register-ScheduledTask -TaskName 'Recipient Limit Report - daily collection' -Action $action -Trigger $trigger `
    -Settings $settings -User 'DOMAIN\svc-rlreport' -Password '<password>' -RunLevel Limited
```

```cards
key | Certificate store | The certificate must be in **CurrentUser\My** of the account that runs the task (or LocalMachine\My with read access to the private key).
clock | First run | Collects `BackfillDays` (7) days: the duration depends on the volume of the tenant (chapter 11). Next runs: the last day only.
shield | Safety | Two collections never run at the same time (lock file). Other tools using the message trace share the same quota of the tenant.
```

# Part III · Use

<!-- icon: terminal -->
## 8. Everyday use

Open **PowerShell 7** (`pwsh`) in **Windows Terminal**, go to the tool folder, run one command.

| I want… | Command |
|---|---|
| The last 7 days (default) | `.\Invoke-RecipientLimitReport.ps1` |
| The last 24 hours | `.\Invoke-RecipientLimitReport.ps1 -Range Last24Hours` |
| The previous calendar month | `.\Invoke-RecipientLimitReport.ps1 -Range PreviousMonth` |
| A given month | `.\Invoke-RecipientLimitReport.ps1 -Range Month -Month 2026-08` |
| One day | `.\Invoke-RecipientLimitReport.ps1 -Range Day -Date 2026-09-28` |
| A custom range | `.\Invoke-RecipientLimitReport.ps1 -Range Custom -Start '2026-09-28 08:00' -End '2026-09-29'` |
| Counts only, smaller files | `.\Invoke-RecipientLimitReport.ps1 -Range PreviousMonth -IncludeRecipientDetails:$false` |
| A report without connecting | `.\Invoke-RecipientLimitReport.ps1 -NoCollect` |
| The daily collection | `.\Invoke-RecipientLimitReport.ps1 -Mode Collect` |
| What the database contains | `.\Invoke-RecipientLimitReport.ps1 -Mode Status` |
| The full help | `Get-Help .\Invoke-RecipientLimitReport.ps1 -Full` |

### What you see

![A report: title card, numbered steps, one row per collected slice, counts before expansion, files written, summary card](images/console-report.png)

```cards
info | Banner | Mode, period, limit and senders, database and log file.
checklist | Numbered steps | 1 database · 2 plan · 3 connection · 4 collection · 5 count · 6 report.
download | Collection table | One row per slice: rows, new messages, pages, duration, rate — with a progress bar, the quota in use and the remaining time.
people | Count | Messages over the limit in the trace, with and without distribution lists, and the counts read before expansion.
target | Summary card | Green when complete, amber when part of the period is missing or counts are still to read, red on error.
```

Icons: ✅ done · 🔹 information · ⚠️ attention · ❌ error · ⏩ step skipped. Everything is also written to `logs\RecipientLimitReport_yyyyMMdd.log` — without colours or icons.

### What the database contains

![Status: one line per day with a coverage bar](images/console-status.png)

Each day shows its messages over the limit, the messages **still to count** (sent to lists, count not read yet), a **coverage bar** and a state: *Complete*, *Complete, refreshed at next run*, *Missing, collectable N more days*, or *no longer collectable*.

### Exit codes

| Code | Meaning |
|---|---|
| `0` | Success |
| `1` | Failure — message on screen and in the log. What was already collected is kept. |
| `2` | Report written but **incomplete**: part of the period is missing, or messages sent to lists are still to count. The summary says which; run the same command again. |

> [!TIP]
> **Interrupted?** Close the window or press <kbd>Ctrl</kbd>+<kbd>C</kbd> at any time: every page received and every count read is already in the database. Run the **same command again** — only what is missing is read, without duplicates.

<!-- icon: chart -->
## 9. Reading the results

### The HTML report

![HTML report: summary tiles, distribution, Top 10 senders, filters and messages](images/report-overview.png)

```cards
chart | Header | Period, messages, senders, average and largest audience, distribution by recipient count.
people | Top 10 senders | Two columns. Click a sender to filter the table.
search | Filters | Search everywhere, sender (exact or partial), subject, recipient, dates, minimum recipients, Message ID.
file | Export | **Export view to CSV** saves only the filtered rows.
```

Click a row to see the message. For a message sent to distribution lists, the dialog shows the count **before expansion** and where it comes from, the number of lists and of addresses in the trace, and the recipients as the sender addressed them — the lists first, flagged:

![Message details: count before expansion, distribution lists flagged](images/report-details.png)

> [!NOTE]
> The page is the page of Purview DLP Report: the same tiles, filters, sort orders and export. The table only draws the visible rows: a 200,000-message file opens in a few seconds. When a report is split, a **Report file** bar moves between the files.

### The CSV file

- UTF-8 with BOM, `;` separator — double-click opens it correctly in Excel.
- Columns: `Received time (Europe/Paris)`, `Sender`, `Recipients` (optional), `Subject`, `Recipient count`, `Distribution lists`, `Message ID`.
- A cell starting with `=`, `+`, `-` or `@` gets a leading `'`: Excel never runs it as a formula.

### What the data means

| Item | Meaning |
|---|---|
| **One row** | One unique Message ID over the limit during the period. A message seen under two trace IDs (for example back from an on-premises server) is reported once. |
| **Received time** | When Exchange Online received the message (`receivedDateTime` of the message trace). |
| **Recipients** | To, Cc and Bcc together — **Bcc is included**: protect the files. As the sender addressed them, distribution lists first: for a message sent to lists, the members added by the expansion are left out (each address is checked in its route). When the list could not be rebuilt (`MaxRoutesPerMessage` reached, or still to read), the cell ends with *(addresses after distribution list expansion)* and the HTML page says so. |
| **Recipient count** | Recipients **before expansion**: the rows of the trace without lists, the Submit event with lists. Never estimated: a message whose count could not be read is not reported, it is counted in the console summary. |
| **Recent hours** | Provisional: late rows are added at the next run. |

<!-- icon: layers -->
## 10. Output files and splitting

Each report creates a folder `reports\<date>_<range>\` with the CSV and HTML files.

> [!TIP]
> Files are written in `<folder>.pending`, renamed only when everything is complete: **a folder without `.pending` is always a complete report**.

```steps
Up to MaxRowsPerFile messages | **One CSV + one HTML**, whatever the period (500,000 by default).
Above | Files per local **week** (Monday–Sunday, default), per **day**, or every *MaxRowsPerFile* **rows**.
Still too large | A week or a day is cut into balanced numbered parts: `_part1of2`, `_part2of2`.
```

```text
RecipientLimit_2026-09-01_to_2026-09-30.csv            a whole month, one file
RecipientLimit_2026-09-07_to_2026-09-13.html           one week of a split month
RecipientLimit_2026-09-14_to_2026-09-20_part1of2.csv   a week too large for one file
RecipientLimit_2026-09-28.csv                          one day
```

> [!TIP]
> **To send by e-mail**: use `-IncludeRecipientDetails:$false`, or zip the CSV (about 10× smaller).

<!-- icon: up -->
## 11. Volume and duration

The message trace has no filter on the number of recipients: **every recipient row** of the period is read, then the tool keeps the messages over the limit.

```cards
download | About 90,000 rows per minute | 90 requests of 5,000 rows per 5 minutes — the quota of the tenant, whatever the server or the number of threads.
people | About 1,000 routes per hour | One route per message sent to a list over the limit for its count, then about one per recipient to rebuild the list of the messages still over the limit — 90 requests per 5 minutes (a separate quota).
database | Only what is reported | Once settled, the messages at or under the limit leave the database: it holds the messages over the limit and their recipients.
```

**How to estimate the first collection**: rows per day ≈ messages per day × average recipients. For example, 400,000 messages a day with 4 recipients on average make 1.6 million rows, about **18 minutes** per day of history. With `Target.SenderDomains`, only the mail sent by the organisation is read.

> [!IMPORTANT]
> **Messages sent to distribution lists are the expensive part.** Every message **over the limit after expansion** and sent to a list needs one request — including the many messages sent to one large list, which are under the limit before expansion and will not be reported. At about 1,000 per hour, a tenant with 20,000 such messages a day needs most of the day for its counts: run the collection every day so that each run only counts one day, and use `Counting.MaxMessagesPerRun` to bound a run. The counts already read are kept: the next run continues.
>
> The messages still **over the limit before expansion** then need about **one request per recipient** to rebuild their recipient list — about 30 for a message to 26 people and a list of 30, in the lab. The addresses already seen as members of the same list are read last, so a list used again costs little more than the recipients the sender typed. `Counting.MaxRoutesPerMessage` (250) bounds one message; `0` keeps the count only.
>
> **Before it starts, the console gives a range** of routes and time (`To read: … <fewest> to <most> route(s), <time> to <time> within the quota …`). The lower bound is one route per count still to read; the upper bound adds every route the recipient lists may need. Which one applies is known only once the counts are read: the messages under the limit before expansion stop after their first route, the others rebuild their list. On the 7-day lab run, 33 counts took 604 routes (32 min).

| Lab measure (2026-10-05) | Value |
|---|---|
| Rows read per minute (quota of 90 requests / 5 min) | **about 92,000** — 1,688,908 rows in 18 min 11 s |
| One day of the DLP load test (51,325 messages over 25) | **18 min 21 s** from an empty day, **8 s** to write the same report again from the database (`-NoCollect`, 61 MB CSV + 6 MB HTML) |
| Recipient lists of the 10 pilot messages sent to lists (first time the lists are seen) | **581 routes in 31 min 17 s** — every address of the trace, the people addressed being sorted after the members |
| 7 days of the load test from an empty database (2026-09-28 → 2026-10-05 18:00) | **3 h 03 min**: collection 2 h 30 (13.2 million rows, 799,446 messages, 2,727 requests, no 429), counts 32 min (33 messages, 604 routes), report 43 s — **400,139** messages over 25 |

# Part IV · Maintain

<!-- icon: gear -->
## 12. Inside the tool

### Execution flow

```flow
database | 1 · Database | configuration, engine, lock
arrow | |
search | 2 · Plan | what is missing or recent?
arrow | |
key | 3 · Connect | only if something must be read
arrow | |
download | 4 · Collect | slices in parallel → SQLite
arrow | |
people | 5 · Count | Submit event of the messages sent to lists
arrow | |
file | 6 · Report | SQLite → CSV + HTML
```

### Collection

```steps
Plan | `Get-RlrCollectionPlan`: for each query (every sender, or each sender domain), the period minus what is already collected **and settled** (older than `SettlingHours` when collected). Beyond 90 days: reported as lost. The rest is cut into slices of `SliceHours` that never cross local midnight, the newest first.
Request | `GET /admin/exchange/tracing/messageTraces?$filter=receivedDateTime ge … and receivedDateTime le …[and senderAddress eq '*@domain']&$top=5000`, then `@odata.nextLink` until the last page. Three slices run in parallel within the shared quota (`RateLimiter`), saved in the database for the next run.
Store | One writer thread stores each page in one transaction: a row per message trace ID, a row per recipient, and the part of the slice now covered (the API returns the newest rows first). Reading a page twice is harmless.
Recover | 429: every request waits (`Retry-After`). 401: new token. 5xx, timeouts: up to 5 retries. A slice that still fails is requested once more for its missing part. 403 or a missing service principal stops the collection; what was stored is kept.
```

### Count before expansion

```steps
Select | Messages of the period with more than `RecipientLimit` rows **and** at least one row `expanded`, count not read yet, fewer than `MaxAttempts` tries.
Request | `GET /admin/exchange/tracing/messageTraces/{id}/getDetailsByRecipient(recipientAddress='<list>')` — the route of a distribution list always holds the Submit event of the message.
Read | `RcptCount` of the **Submit** event = recipients as addressed by the sender. A message received by SMTP has no Submit: `RcptCount` of the **Receive** event. The other events carry other counts (members of one list, recipients after expansion): never used.
Fallback | No count in the route of the list: the route of another recipient is read. Still nothing: the message is marked *without a count* and listed in the summary.
Rebuild | Count over the limit: the routes of the other recipients are read — the lists first, the addresses already seen as members of the same lists in other messages last — until as many recipients as the count are found. A route that holds the event of the count (`Submit`, or `Receive` for SMTP) is a recipient the sender addressed (`recipient.direct = 1`); the others were added by the expansion (`0`). At most `MaxRoutesPerMessage` requests per message (`list_state` = `Done` or `Partial`).
```

### Compaction

When a period has **settled**, the messages it holds at or under the limit — or whose count before expansion is at or under the limit — are deleted with their recipients. The database then holds what a report can show, plus the messages still to count. The limit used is recorded (`compaction_limit`).

### Report

The engine selects the messages of the period over the limit, keeps **one row per Message ID**, orders them, plans the files and streams them to the CSV and HTML writers — inside **one read transaction**, so a collection running at the same time cannot change the result halfway. The HTML embeds the data compressed (JSON → gzip → base64, blocks of 20,000 rows), like Purview DLP Report.

### Code map

| File | Part | Role |
|---|---|---|
| `Invoke-RecipientLimitReport.ps1` | — | Parameters, the 6 steps, summary, exit code. **Start reading here.** |
| `RecipientLimitReport.psm1` | Region 1 · Console and log | Theme (colours, icons), banner, steps, items, table rows, summary card, log |
| | Region 2 · Configuration | `Import-RlrConfiguration` — every check of the .psd1 |
| | Region 3 · Engine | `Initialize-RlrEngine` — loads SQLite, compiles `src\*.cs` into `bin\` |
| | Region 4 · Periods and coverage | `Resolve-RlrPeriod`, `Get-RlrSignatures`, `Get-RlrCollectionPlan`, `Get-RlrCoverage` |
| | Region 5 · Connection | `Connect-RlrGraph` (certificate, secret, interactive), `Update-RlrToken` |
| | Region 6 · Collection | `Invoke-RlrCollection` — progress, table, second pass |
| | Region 7 · Recipient count | `Invoke-RlrCounting`, `Get-RlrCountingEstimate` |
| | Region 8 · Report | `New-RlrReport`, `Move-RlrDirectory` |
| | Region 9 · Status and maintenance | `Show-RlrStatus`, `Invoke-RlrCompaction`, `Invoke-RlrRetention`, `Enter-RlrLock` |
| `src\RecipientLimitReport.Engine.cs` | `Coverage` · `GraphClient` · `RateLimiter` · `TraceCollector` · `CountCollector` · `RlrStore` · `ReportPlanner` · `ReportWriter` | Intervals · HTTP · quota · collection · counts · SQLite · splitting · CSV/HTML |
| `templates\Report.template.html` | — | HTML page. Markers `%%CHUNKS%%` then `%%META%%` receive the data. |

<!-- icon: wrench -->
## 13. Modifying the tool

> [!IMPORTANT]
> Code and comments in **English**. Keep the header block (author, version) of each file. Update `CHANGELOG.md` and the version (Annex E). Run the tests (chapter 14).

### Recipes

| I want to… | Where |
|---|---|
| Change the look of the HTML report | `templates\Report.template.html` only — colours, texts, behaviour. No rebuild. Keep the two markers, once each, outside comments. Keep it in step with the template of Purview DLP Report. |
| Add a report column | `ReportRow` and `ReportSql` (engine), `CsvOut` (header and `Add()`), `HtmlOut.Add()` / `Flush()` (new array in the chunk), then read it in the template. A new stored value needs a schema migration (`RlrStore.SchemaVersion`, `Migrate`). |
| Change splitting or file names | `ReportPlanner.Plan`, `PeriodLabel`, `FileLabel` (engine). |
| Add a period ("current week"…) | `ValidateSet` of `-Range` **and** of `Resolve-RlrPeriod`, a new `switch` branch, a test in *Periods*. |
| Change the console output | Always go through `Write-RlrStep`, `Write-RlrItem`, `Write-RlrTableRow`, `Write-RlrSummary`: they also write the log. |
| Update the SQLite libraries | nuget.org packages `Microsoft.Data.Sqlite.Core` + `SQLitePCLRaw.*` (same version): `lib/net8.0` into `package\lib\sqlite`, `runtimes/win-*/native/e_sqlite3.dll` into `package\lib\sqlite\runtimes`. Delete `bin\`, run the tests, update `THIRD-PARTY-NOTICES.md`. |
| Change this guide | Edit `package\docs\RecipientLimitReport-Guide.md` or `package\docs\RecipientLimitReport-UserGuide.md` (callouts `> [!NOTE]`, blocks `cards`, `steps`, `flow`), then run `.\tools\Build-Documentation.ps1`. |

### Pitfalls of the message trace API (lab, 2026-10-02 and 2026-10-05)

> [!CAUTION]
> **`$filter` keeps one value per property.** `senderAddress eq 'a' or senderAddress eq 'b'` is accepted but only the last value is applied, without any error. Every sender domain therefore needs its own query.

> [!CAUTION]
> **The counts of the route are not all the same count.** In the route of a message sent to two lists, `Submit` says 28 (before expansion), `Expand DL` says 4 (the members of one list) and `DLP rule` says 55 (after expansion). Only `Submit` — or `Receive` for SMTP — gives the count of `RecipientLimits`.

> [!CAUTION]
> **PowerShell pitfalls**: `-f` binds tighter than `+` (put expressions in parentheses); `$C` and `$c` are the same variable; `[Math]::Max(0, $x)` selects the integer overload — write `[Math]::Max(0.0, $x)`; an empty list returned by an `if` becomes `$null` — wrap it in `@()`.

<!-- icon: beaker -->
## 14. Testing a change

```powershell
cd D:\Tools\RecipientLimitReport
Invoke-Pester -Path .\tests\RecipientLimitReport.Tests.ps1 -Output Detailed
```

**35 tests, about 50 seconds, no connection to Microsoft 365.** Microsoft Graph is replaced by an in-memory message trace API (`tests\FakeGraph.cs`) that behaves like the real one: paging newest first, one value per property, distribution lists expanded into `expanded` rows, routes with the Submit, Receive, Expand DL and DLP rule events (none of Submit and Receive for a member added by the expansion), and failures on demand (429, 401, 500, 403).

```cards
settings | Configuration & periods | Validation, sender domains, summer/winter time, month boundaries.
download | Collection | Paging, retries, token renewal, permission error, idempotence, sender domains, plan.
people | Count and list before expansion | Submit, Receive (SMTP), route without count, one request per message for the count, recipient list rebuilt (members left out, early stop, members already seen read last, MaxRoutesPerMessage).
file | Report | Columns, one row per Message ID, recipients option and cut, splitting, HTML, CSV formula protection.
```

# Annexes

<!-- icon: lifebuoy -->
## Annex A — Troubleshooting

### Configuration and connection

| Symptom | Cause | What to do |
|---|---|---|
| `Invalid configuration (...)` and a list | Wrong values in the .psd1 | Fix every line listed. Nothing was done. |
| `Microsoft Entra sign-in failed (AADSTS700027 …)` | The certificate is not registered on the application | Upload the `.cer` of `CertificateThumbprint` in *Certificates & secrets*. The message names the most frequent AADSTS codes. |
| `Certificate … not found` / `without its private key` | The task runs with another account, or the `.cer` was imported instead of the `.pfx` | Import the certificate with its private key for the account of the task. |
| `The application has no application permission ExchangeMessageTrace.Read.All` | Permission or admin consent missing | API permissions › Microsoft Graph › Application › `ExchangeMessageTrace.Read.All` › **Grant admin consent**. |
| `401 … service principal … 8bd644d1-64a1-4d4b-ae52-2e0cbf64e373` | The message trace service is not provisioned in the tenant | `New-MgServicePrincipal -AppId 8bd644d1-64a1-4d4b-ae52-2e0cbf64e373` (chapter 7), then wait — up to a few hours. |
| `Connected to tenant …, but Tenant.TenantId is …` | Wrong tenant | Check `Tenant.TenantId` and the application. Nothing was read. |
| `Another execution is already collecting` | The scheduled task is running | Wait, or use `-NoCollect`. |

### Collection and count

| Symptom | Cause | What to do |
|---|---|---|
| *Throttled by Microsoft Graph (429)* | Another tool or an administrator uses the message trace of the tenant | Nothing: every request waits and resumes. Lower `Collection.MaxRequests` if it happens often. |
| *waiting for the quota* in the progress bar | The 90 requests of the last 5 minutes are used | Normal: the quota of the tenant is the limit. |
| `Missing … collectable for N more day(s)` | The daily collection did not run | Run `-Mode Collect` or a report of that period before the deadline; check the task. |
| *N message(s) … without a count in their route* | Neither Submit nor Receive carries a count (rare: forwarded or relayed messages) | Not reported, never guessed. Check one of them in the message trace of the Exchange admin center. |
| *still to count* after a run | `Counting.MaxMessagesPerRun` reached, or requests failed | Run again: the counts already read are kept. |
| *recipient list(s) to rebuild* after a run | The run stopped, or requests failed | Run again: the routes already read are kept. |
| *left after expansion (more than N routes to read)* | A message to large lists needs more than `Counting.MaxRoutesPerMessage` requests | Its count is right; its recipient list is the one after expansion. Raise the setting if the quota allows it. |
| `RecipientLimit is …, but the database was compacted with …` | The limit was lowered | Older periods only hold the messages over the former limit (chapter 6). |

### Report

| Symptom | Cause | What to do |
|---|---|---|
| "Report written, but INCOMPLETE", exit code 2 | Part of the period is not in the database, or counts are still to read | Run again. Periods older than 90 days are lost. |
| A message of the DLP report is missing | Received time and detection time differ by a few seconds at the edge of the period, or the message is beyond the sender domains | Compare on a slightly larger period; check `Target.SenderDomains`. |
| More addresses than the recipient count, and *(addresses after distribution list expansion)* | The list of this message could not be rebuilt: still to read, or `MaxRoutesPerMessage` reached | Run again, or raise `Counting.MaxRoutesPerMessage`. The count is right (chapter 2). |
| A folder ends with `.pending` | The execution stopped while writing | Delete it, run again. |
| "This browser cannot open the compressed data" | Old browser | Use Edge/Chrome/Firefox, or the CSV. |
| Squares instead of icons in the console | The console is not Windows Terminal | Run the tool in Windows Terminal. |
| `The database schema version is …` | Older tool, newer database | Use the matching tool version. |

<!-- icon: compare -->
## Annex B — Message trace and DLP report

Both tools produce **the same report**. They differ by their source:

| | Recipient Limit Report | Purview DLP Report |
|---|---|---|
| Source | Exchange Online message trace (Microsoft Graph) | Activity Explorer (DLP rule matches) |
| Needs a DLP policy | **No** | Yes, in audit mode |
| Sees the messages refused by `RecipientLimits` | **Yes** — they are in the message trace | No — Exchange refuses them before the DLP rules (lab, 2026-10-04) |
| Volume read | Every recipient row of the tenant (or of the sender domains) | Only the rule matches |
| Count before expansion | Trace rows, or the Submit event for messages sent to lists | Value of the rule condition |
| Recipient list | As addressed by the sender, rebuilt from the route of each recipient | As addressed by the sender (event of the rule) |
| Time shown | Received time | Detection time |
| Retention of the source | 90 days | 30 days |
| Permissions | `ExchangeMessageTrace.Read.All` (Graph) | Purview role *Information Protection Reader* |

**Lab comparison** — same periods, same messages (Annex C): on the 48-hour comparison window, **all 149,802 messages** of the DLP report are in this report **with the same recipient count**, distribution lists included. This report has **352 more** (0.23 %): messages over 25 recipients that the DLP rule never evaluated, or that it blocked without any event in Activity Explorer.

| End-to-end time on the 48-hour window, empty database | Recipient Limit Report | Purview DLP Report 2.1 |
|---|---:|---:|
| Collection | about **56 min** (4.96 million trace rows, 3 runs) | 16 min 26 s (149,807 events) |
| Report from the database | 19 s | 22 s |

<!-- icon: chart -->
## Annex C — Lab measurements

Lab tenant, 2026-10-05, application in certificate mode, quota of 90 requests per 5 minutes. The messages are the DLP load test of 2026-09-28 and 2026-09-29 (direct recipients, 26 to 60 per message) and the pilot of 2026-09-28 (200 messages, some sent to distribution lists). Reference: the database of Purview DLP Report on the same periods.

| Period (Europe/Paris) | Rows read | Messages over 25 | Counts read | Duration | Compared with the DLP report |
|---|---:|---:|---:|---:|---|
| 2026-09-28 11:00 → 13:00 (pilot) | 8,051 | 141 | 33 | 32 s | **141 / 141** Message IDs, **0** count difference (10 messages with lists) |
| 2026-09-28, whole day | 1,688,908 | 51,325 | 0 (lists already counted) | 18 min 21 s | **51,139 / 51,139** DLP messages found, **0** count difference, **+186** messages the DLP rule never evaluated (below) |
| 2026-09-27 16:24 → 2026-09-29 16:24 — the 48-hour comparison window of Purview DLP Report | 3,261,286 (the day of 28/09 already in the database) | **150,154** | 0 | 37 min 03 s | **149,802 / 149,802** DLP messages found, **0** count difference, **+352** messages missing from Activity Explorer (below) |

```cards
download | About 90,000 rows per minute | 1,688,908 rows in 341 requests (18 min 11 s), then 3,261,286 rows in 665 requests (36 min 37 s, one 429 absorbed): the quota of 90 requests per 5 minutes. The collection is bounded by Microsoft, not by the server.
people | 33 counts in 22 s | The pilot: 33 messages sent to lists over 25 after expansion, one route each — 23 of them were under 25 before expansion.
checklist | 10 recipient lists in 31 min | The 10 messages over 25 before expansion: 581 routes (one per address of the trace, 31 to 90 per message), and for **10 / 10** the same recipients as the DLP report — the lists and the people addressed, none of the members.
database | 150 MB for 150,000 messages | Only the messages over the limit stay in the database, with their recipients: about 1 KB per message. 98,674 messages at or under 25 were removed after the 48-hour window.
check | Same numbers | Every message of the DLP report, with the same recipient count, distribution lists included. The report of the 48 hours (180 MB CSV, 17 MB HTML) is written in 19 s.
```

> [!NOTE]
> **352 messages the DLP report could not show** on the 48-hour window (0.23 %), absent from Activity Explorer at any date:
> - **345 delivered** — sent by the senders of the load test to 26 to 38 recipients while the DLP rule blocked all the others. Their route has a Submit event (26, 27 … recipients) and **no DLP rule event**: the rule never evaluated them.
> - **7 blocked by the DLP rule** — their route shows `DLP rule` with the action *block* (`BA`) and the refusal `550 5.7.171`, but Activity Explorer has no event for them.
>
> The message trace records every message, whatever the DLP service did: these messages are in this report.

<!-- icon: database -->
## Annex D — Database schema

SQLite, `data\RecipientLimitReport.sqlite` (WAL mode). All times are **Unix milliseconds, UTC**. Open a **copy** with any SQLite tool, never during a collection.

| Table | One row per | Main columns |
|---|---|---|
| `metadata` | setting | `created_utc`, `created_by_version`, `compaction_limit`, quota windows |
| `run` | execution | `mode`, `status`, `account`, `details` (JSON) |
| `signature` | kind of query | `text` (`''` = every sender, `sender=*@contoso.com`) |
| `collection_interval` | slice asked to the message trace | `start_ms`, `end_ms`, `covered_from_ms`, `collected_ms`, `status`, `pages`, `rows`, `error` |
| `address` | sender or recipient address | `address` (case-insensitive) |
| `status` | delivery status | `name` (`delivered`, `failed`, `expanded` …) |
| `message` | message trace ID | `message_id`, `sender_id`, `subject`, `received_ms`, `recipient_rows`, `list_rows`, `count_value`, `count_source`, `count_state`, `count_attempts`, `list_state` (`Done`, `Partial`), `route_reads` |
| `recipient` | message and recipient | `message_key`, `address_id`, `status_id`, `direct` (`1` addressed by the sender, `0` added by the expansion, empty when not read) |

```sql
-- Last slices collected
SELECT datetime(start_ms/1000,'unixepoch') AS start_utc, datetime(end_ms/1000,'unixepoch') AS end_utc,
       datetime(covered_from_ms/1000,'unixepoch') AS covered_from_utc, status, rows, error
FROM collection_interval ORDER BY interval_id DESC LIMIT 20;

-- Messages sent to lists whose count could not be read
SELECT message_id, recipient_rows, list_rows, count_state, count_attempts, count_error
FROM message WHERE list_rows > 0 AND count_value IS NULL AND recipient_rows > 25;

-- Messages over 25 whose recipient list is shown after expansion
SELECT message_id, count_value, recipient_rows, list_state, route_reads, count_error
FROM message WHERE list_rows > 0 AND count_value > 25 AND COALESCE(list_state, '') <> 'Done';
```

> [!TIP]
> To force the re-collection of a period still within 90 days, delete its rows from `collection_interval`. Messages already stored are kept and never duplicated.

<!-- icon: tag -->
## Annex E — Versioning and release checklist

Version numbers follow **MAJOR.MINOR.PATCH** — MAJOR: incompatible change (configuration or database) · MINOR: new feature · PATCH: fix. The version appears in `package\RecipientLimitReport.psd1`, `$script:ToolVersion`, the file headers, this guide and `CHANGELOG.md`. The folder is a **git** repository (`git log --oneline`, `git tag`).

```steps
Code | Update the code and the comments, in English.
Version | Update the version everywhere, and `CHANGELOG.md`.
Tests | `Invoke-Pester -Path .\tests` — all green.
Real data | For a change of the collection, the count or the engine: one live run on a known period, compared with the previous version.
Documentation | Update this guide and the user guide, then `.\tools\Build-Documentation.ps1` to regenerate the HTML.
Release | `git add -A`, `git commit`, `git tag vX.Y.Z`.
Package | `.\tools\New-RlrPackage.ps1` — copies the files needed to run into `..\package\RecipientLimitReport-X.Y.Z`, tenant values emptied, no database. Zip this folder to deliver it.
```

<!-- icon: refresh -->
## Annex F — From the scripts of July 2026

Version 1.0.0 replaces `Invoke-DeliveredEnvelopeAuditReport.ps1` and `Invoke-RecipientLimitBlockReport.ps1` (package *Scripts-AuditRecipientLimit*, 2026-07-15). The method of the audit script is kept — the count before expansion read from the Submit event of a distribution list, and no request for the messages without lists — inside the structure of Purview DLP Report.

| Topic | Scripts of July 2026 | Version 1.0.0 |
|---|---|---|
| Entry point | Two scripts, many parameters | One script, three modes, one configuration file |
| API | Graph **beta** | Graph **v1.0** |
| Collection | Five pulls per status and per UTC day, sequential | One query per slice (or per sender domain), parallel within the quota, newest first, page by page coverage |
| Population | Delivered messages only (a message without any delivered recipient was left out) | Every message over the limit, like the DLP report |
| Messages without lists | Floor value (`≥ N`) | Exact: every row of the trace is a recipient addressed by the sender |
| Report | Summary HTML (20,000 rows), CSV of counts | The report of Purview DLP Report (CSV + HTML, recipients optional, splitting) |
| Settings | `state\graph-app.json`, vault or environment | Configuration file; secret in an environment variable or the prompt; certificate recommended |
| Blocked messages (`550 5.5.3`) | Separate report | Not in this version: in audit mode no message is blocked |
