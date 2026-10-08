---
title: Recipient Limit Report
subtitle: User guide
version: 1.0.1
author: Nicolas Fabert
updated: 2026-10-08
---

# Recipient Limit Report — User guide

> What is needed before the first report, then the commands used every day: **which messages exceed the recipient limit?**, **what did the sender address before distribution-list expansion?**, and **is the collection up to date?** The configuration details, Microsoft Graph behaviour and internals are in the [developer guide](RecipientLimitReport-Guide.md).

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
checklist | Prerequisites | PowerShell 7.4+, Microsoft Graph permission, certificate or interactive sign-in, HTTPS access.
terminal | Everyday use | One command for a period, a daily collection, or the database status.
file | Results | CSV and self-contained HTML files written locally under `reports\`.
```

# Part I · Start here

<!-- icon: checklist -->
## 1. Prerequisites

| Item | Requirement |
|---|---|
| Operating system | Windows 10/11 or Windows Server 2016+, x64 or ARM64 |
| PowerShell | **PowerShell 7.4 or later** (`pwsh`), not Windows PowerShell 5.1 |
| Console | Windows Terminal, for the emoji and the colours of the display |
| Modules | **None** in certificate and secret modes; `Microsoft.Graph.Authentication` for the interactive sign-in |
| Browser | Recent Edge, Chrome or Firefox for the HTML report |
| Network | HTTPS to `graph.microsoft.com` and `login.microsoftonline.com` |
| Disk | Space for the SQLite history and the reports; large periods can produce large CSV files |
| SQLite | Bundled in `lib\sqlite` — nothing to install |
| Microsoft Graph | `ExchangeMessageTrace.Read.All`, application permission with admin consent for unattended use, or delegated consent for interactive use |

For certificate mode, the certificate with its private key must be available to the account that runs the tool. The tenant must also have the Microsoft message trace service principal (`8bd644d1-64a1-4d4b-ae52-2e0cbf64e373`). The complete registration procedure is in the [developer guide, chapter 7](RecipientLimitReport-Guide.md#7-unattended-execution).

> [!IMPORTANT]
> The tool is **read-only** for Microsoft 365: it never sends mail and never changes a tenant setting. Reports and the SQLite database stay on the local disk.

<!-- icon: download -->
## 2. One-time setup

Everything you run is in the `package` folder of the repository, or in the folder where you extracted the zip of a [release](https://github.com/Nico77600/RecipientLimitReport/releases). Work from that folder.

```steps
Get the files | `git clone https://github.com/Nico77600/RecipientLimitReport.git`, then `cd RecipientLimitReport\package` — or extract the release zip, for example to `D:\Tools\RecipientLimitReport`.
Unblock the files | From that folder: `Get-ChildItem . -Recurse -File | Unblock-File`.
Edit the configuration | `notepad .\config\RecipientLimitReport.config.psd1` and fill in the tenant, authentication, recipient limit and sender domains.
Check the installation | `.\Invoke-RecipientLimitReport.ps1 -Mode Status` checks the configuration and builds the engine, without connecting to Microsoft 365. On a new installation it ends with *Ready for the first collection*.
Run the first report | `.\Invoke-RecipientLimitReport.ps1` collects the missing data, counts the messages sent to distribution lists and writes the report.
```

The delivered configuration is intentionally empty for tenant-specific values. The main settings are:

| Setting | Meaning |
|---|---|
| `Tenant.TenantId` | Microsoft Entra tenant GUID; the token is checked against it |
| `Authentication.Mode` | `Certificate`, `ClientSecret` or `Interactive` |
| `Authentication.AppId` | Application (client) ID; empty in interactive mode uses Microsoft Graph Command Line Tools |
| `Authentication.CertificateThumbprint` | Certificate thumbprint in the certificate store of the running account |
| `Target.RecipientLimit` | `25` reports messages with **more than** 25 recipients |
| `Target.SenderDomains` | `@()` reads every sender; accepted domains limit the report to messages sent by the organisation |

The developer guide covers client secrets, interactive sign-in, service-principal creation and every configuration key.

> [!NOTE]
> `data\`, `reports\` and `logs\` are created at the first run. Back up `data\`: it is the only history beyond the 90 days of the message trace.

# Part II · Everyday use

<!-- icon: terminal -->
## 3. Build a report

Run the commands from that same folder, in PowerShell 7:

| I want… | Command |
|---|---|
| The last 7 days (default) | `.\Invoke-RecipientLimitReport.ps1` |
| The last 24 hours | `.\Invoke-RecipientLimitReport.ps1 -Range Last24Hours` |
| The last 30 days | `.\Invoke-RecipientLimitReport.ps1 -Range Last30Days` |
| The previous calendar month | `.\Invoke-RecipientLimitReport.ps1 -Range PreviousMonth` |
| A given month | `.\Invoke-RecipientLimitReport.ps1 -Range Month -Month 2026-08` |
| One day | `.\Invoke-RecipientLimitReport.ps1 -Range Day -Date 2026-09-28` |
| A custom period | `.\Invoke-RecipientLimitReport.ps1 -Range Custom -Start '2026-09-28 08:00' -End '2026-09-29'` |
| Count only, smaller files | `.\Invoke-RecipientLimitReport.ps1 -Range PreviousMonth -IncludeRecipientDetails:$false` |
| Use only already collected data | `.\Invoke-RecipientLimitReport.ps1 -NoCollect` |
| Check the database | `.\Invoke-RecipientLimitReport.ps1 -Mode Status` |
| The full help | `Get-Help .\Invoke-RecipientLimitReport.ps1 -Full` |

Unless `-NoCollect` is used, Report mode first collects the missing part of the period, reads the recipient count before distribution-list expansion where needed, and then writes the files. The default period is controlled by `Report.DefaultRange` in the configuration.

The console shows a banner (mode, period, limit and senders, database and log file), then six numbered steps — database, plan, connection, collection, count, report — and a summary card: green when complete, amber when part of the period is missing or counts are still to read, red on error.

> [!NOTE]
> The message trace keeps about **90 days**. The database keeps the messages over the limit beyond that period, so run the daily collection to preserve history.

<!-- icon: calendar -->
## 4. Schedule the collection

The daily collection is intended for a scheduled task using certificate authentication. Change `D:\Tools\RecipientLimitReport` to your own folder:

```powershell
$action = New-ScheduledTaskAction -Execute 'pwsh.exe' `
    -Argument '-NoProfile -NonInteractive -File "D:\Tools\RecipientLimitReport\Invoke-RecipientLimitReport.ps1" -Mode Collect' `
    -WorkingDirectory 'D:\Tools\RecipientLimitReport'
$trigger = New-ScheduledTaskTrigger -Daily -At 06:00
$settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Hours 12) -StartWhenAvailable -MultipleInstances IgnoreNew
Register-ScheduledTask -TaskName 'Recipient Limit Report - daily collection' -Action $action -Trigger $trigger `
    -Settings $settings -User 'DOMAIN\svc-rlreport' -Password '<password>' -RunLevel Limited
```

The first collection checks `Collection.BackfillDays` (7 by default). Later runs collect only what is missing and refresh the most recent hours, where message-trace rows and statuses can still change.

> [!TIP]
> Two collections never run at the same time: a lock file stops the second one. If a scheduled collection is running, use `-NoCollect` for a report.

<!-- icon: chart -->
## 5. Read the results

Each report creates a new folder `reports\<date>_<range>\`. It contains:

| File | Content |
|---|---|
| `.html` | Self-contained report with summary tiles, the distribution by recipient count, Top 10 senders, search, filters, message details and **Export view to CSV** |
| `.csv` | Complete export, UTF-8 with BOM and `;` separator for Excel |

The report contains one row per unique Message ID over the limit:

| Column | Meaning |
|---|---|
| Received time | Exchange Online receive time in the configured time zone |
| Sender | Sender address |
| Recipients | Addresses as the sender addressed them; optional with `IncludeRecipientDetails` |
| Subject | Message subject |
| Recipient count | Count **before** distribution-list expansion |
| Distribution lists | Number of lists used by the message |
| Message ID | Internet Message ID |

A distribution list counts as **one recipient**. For a message sent to a list, the report lists the list first and leaves out members added by the expansion. Bcc recipients are included: protect CSV and HTML files as sensitive data.

Up to `Report.MaxRowsPerFile` messages (500,000 by default) the run writes one CSV and one HTML file. Above that, files are split by week (default), by day or by number of rows, and a week or a day too large is cut into numbered parts (`_part1of2`).

> [!TIP]
> A folder ending in `.pending` is still being written. A folder without `.pending` is a complete report.

<!-- icon: info -->
## 6. Status, logs and exit codes

Use Status to see the data held day by day, missing periods and messages still waiting for a count:

```powershell
.\Invoke-RecipientLimitReport.ps1 -Mode Status
```

Each day shows its messages over the limit, the messages still to count, a coverage bar and a state: *Complete*, *Complete, refreshed at next run*, *Missing, collectable N more days*, or *no longer collectable*.

The daily log is written to `logs\RecipientLimitReport_yyyyMMdd.log`, without colours or icons, and kept 30 days.

| Exit code | Meaning |
|---|---|
| `0` | Success |
| `1` | Failure; read the red console message and the daily log |
| `2` | Report written but incomplete; part of the period is missing or messages sent to lists still need counting |

If a run is interrupted, run the same command again. Received pages and completed counts are already stored, so the next run resumes without duplicates.

# Part III · Troubleshoot

<!-- icon: lifebuoy -->
## 7. Common situations

| Symptom | Cause | What to do |
|---|---|---|
| `Invalid configuration (...)` and a list of lines | Wrong values in the configuration file | Fix every line listed, then run `-Mode Status` again. Nothing was done. |
| `Connected to tenant …, but Tenant.TenantId is …` | The token belongs to another tenant | Correct `Tenant.TenantId`, or sign in with the expected account. |
| `401 … service principal … 8bd644d1-64a1-4d4b-ae52-2e0cbf64e373` | The message trace service is not provisioned in the tenant | Create the service principal (developer guide, chapter 7), then wait — up to a few hours. |
| `The application has no application permission ExchangeMessageTrace.Read.All` | Permission or admin consent missing | Grant `ExchangeMessageTrace.Read.All` and the admin or delegated consent. |
| `Certificate … not found` / `without its private key` | Another account runs the task, or the `.cer` was imported instead of the `.pfx` | Import the certificate with its private key for the account that runs the tool. |
| *Another execution is already collecting* | The scheduled collection is running | Wait, or use `-NoCollect`. |
| Report is incomplete (exit code `2`) | Part of the period is missing, or counts are still to read | Run the same command again; the summary says which part. |
| *waiting for the quota*, collection is slow | The tenant quota is 100 requests per 5 minutes | Normal: the tool uses 90 of them (`Collection.MaxRequests`). |
| `Missing … collectable for N more day(s)` | The daily collection did not run | Run `-Mode Collect`, or a report of that period, before the deadline; check the scheduled task. |
| Large CSV or HTML file | Recipient addresses, or a very long period | Use `-IncludeRecipientDetails:$false`, a shorter period, or `-SplitBy Day`. |
| Squares instead of icons in the console | The console is not Windows Terminal | Run the tool in Windows Terminal. |

For detailed Graph behaviour, quota sizing, configuration validation, database retention and development instructions, continue with the [developer guide](RecipientLimitReport-Guide.md).
