# Recipient Limit Report

Reports the Exchange Online messages sent to **more than 25 recipients** (the limit is configurable) — counted as Exchange Online counts them, **a distribution list being one recipient** — as CSV and HTML files, **one row per Message ID**. The source is the Exchange Online **message trace** (Microsoft Graph); the report is the same as the one of [Purview DLP Report](https://github.com/Nico77600/PurviewDlpReport).

> [!IMPORTANT]
> Files downloaded from the Internet may be blocked by Windows and fail to run. Before using this project, unblock every file in the downloaded folder:
>
> ```powershell
> Get-ChildItem "C:\Chemin\Du\Dossier" -Recurse -File -Force | Unblock-File
> ```
>
> Replace the example path with the folder where you downloaded or extracted this project.

![HTML report](docs/images/report-overview.png)

## Why

Before applying a recipient limit (`Set-Mailbox -RecipientLimits`, or a DLP policy), an organisation needs to know **which messages would be blocked** — who sends to more than 25 recipients, to whom, about what, when — and the administrators need facts to maintain the **exception list**. Purview DLP Report answers it from a DLP policy in audit mode. This tool answers it **without any DLP policy**, from the message trace, and keeps working once the limit is applied, when Exchange refuses the messages before a DLP rule sees them.

## How it works

```
Message trace (Graph)  ──►  local SQLite history  ──►  count before expansion  ──►  CSV + HTML report
(every recipient,           (beyond the 90 days of      (route of the message,       (one row per Message ID,
 90 days, newest first)      the message trace)           only for lists)               local files only)
```

- **Collect**: the message trace lists one row per message and recipient, after expansion of the distribution lists. Every row is read (optionally only the senders of the accepted domains), within the tenant quota of 100 requests per 5 minutes. Collections are restartable and never store a message twice.
- **Count**: for a message without distribution list, the recipients of the trace are the recipients the sender addressed. For a message sent to a list, the count before expansion is read from its route (`getDetailsByRecipient`, Submit event).
- **Recipient list**: for a message over the limit sent to a list, the route of each recipient tells whether the sender addressed it or the expansion of the list added it: the report lists the recipients as the sender addressed them, like the DLP report.
- **Report**: one row per unique Message ID — received time, sender, recipients (optional), subject, recipient count, distribution lists, Message ID — as CSV and a self-contained HTML file with filters, Top 10 senders and message details. Large periods are split by day, week or number of rows.
- **Read-only** for Microsoft 365: the tool never sends e-mail and never changes a setting.

![Console](docs/images/console-report.png)

## Requirements

| Item | Requirement |
|---|---|
| PowerShell | 7.4 or later |
| Modules | None in certificate and secret modes; `Microsoft.Graph.Authentication` for the interactive mode (its MSAL library) |
| Console | Windows Terminal (emoji and colours) |
| Permissions | Microsoft Graph `ExchangeMessageTrace.Read.All` — for unattended runs, an application permission with admin consent and a certificate (administrator guide, chapter 7) |
| Tenant | The service principal of the Microsoft message trace service (`8bd644d1-64a1-4d4b-ae52-2e0cbf64e373`), created once per tenant |
| SQLite | Bundled in `lib\sqlite` — nothing to install |

## Quick start

```powershell
git clone https://github.com/Nico77600/RecipientLimitReport.git
cd RecipientLimitReport
notepad .\config\RecipientLimitReport.config.psd1      # tenant, application, sender domains

.\Invoke-RecipientLimitReport.ps1 -Mode Status          # checks the configuration, no connection
.\Invoke-RecipientLimitReport.ps1                       # report of the last 7 days
.\Invoke-RecipientLimitReport.ps1 -Range PreviousMonth  # previous calendar month
.\Invoke-RecipientLimitReport.ps1 -Mode Collect         # daily collection (scheduled task)
```

The message trace keeps 90 days: schedule the daily collection to keep a longer history. The zip of each [release](https://github.com/Nico77600/RecipientLimitReport/releases) contains only the files needed to run, with both guides in HTML; `.\tools\New-RlrPackage.ps1` builds the same package from the repository.

## Documentation

| Guide | Content |
|---|---|
| **[User guide](docs/RecipientLimitReport-UserGuide.md)** | Prerequisites, one-time setup, everyday report commands, scheduled collection, how to read the results, exit codes and common situations. |
| **[Administrator guide](docs/RecipientLimitReport-Guide.md)** | Installation, configuration, unattended execution with a certificate, the message trace and distribution list counting rules, volume and duration, troubleshooting, lab measurements and the internals. |

Both guides also exist as a single HTML file (`docs/RecipientLimitReport-UserGuide.html`, `docs/RecipientLimitReport-Guide.html`): download them and open them locally, or use the copies in the release zip.

## Tests

```powershell
Invoke-Pester -Path .\tests      # Pester 5, no connection to Microsoft 365
```

## License

[MIT](LICENSE). The bundled SQLite components keep their own licenses: see [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md).

## Disclaimer

Personal project, provided as is. It is not an official Microsoft product and is not supported by Microsoft. Test it in your environment before production use.
