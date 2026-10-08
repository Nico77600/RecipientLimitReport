# Recipient Limit Report

A PowerShell 7 tool that reads the Exchange Online message trace and writes CSV and HTML reports of the messages sent to more than 25 recipients.

This folder contains everything needed to run the tool: `Invoke-RecipientLimitReport.ps1`, the module and its C# engine, the configuration, the report template, the SQLite library and the guides. Tests and build tools stay outside it, in the repository.

> [!IMPORTANT]
> Files downloaded from the Internet may be blocked by Windows. Unblock them once, from this folder:
>
> ```powershell
> Get-ChildItem . -Recurse -File | Unblock-File
> ```

## Requirements

- PowerShell 7.4 or later.
- No module with a certificate; `Microsoft.Graph.Authentication` for the interactive sign-in.
- Windows Terminal for emoji and colours.
- Microsoft Graph `ExchangeMessageTrace.Read.All`. For unattended runs: an application with a certificate and this application permission, and the service principal of the message trace service in the tenant (guide, chapter 7).
- SQLite is bundled in `lib\sqlite`.

## Quick start

```powershell
notepad .\config\RecipientLimitReport.config.psd1      # TenantId, application, sender domains

.\Invoke-RecipientLimitReport.ps1 -Mode Status          # checks the configuration, no connection
.\Invoke-RecipientLimitReport.ps1                       # report of the last 7 days
.\Invoke-RecipientLimitReport.ps1 -Range PreviousMonth  # previous calendar month
.\Invoke-RecipientLimitReport.ps1 -Mode Collect         # daily collection (scheduled task)
```

## Content

| Item | Role |
|---|---|
| `config\` | Configuration file to fill in. |
| `docs\` | User and developer guides, Markdown and self-contained HTML. |
| `lib\` | Bundled SQLite libraries. |
| `src\` | C# engine source, compiled on first use. |
| `templates\` | HTML report template. |
| `Invoke-RecipientLimitReport.ps1` | Entry script. |
| `RecipientLimitReport.psd1` | Module manifest. |
| `RecipientLimitReport.psm1` | PowerShell module used by the script. |
| `README.md` | This package readme. |
| `LICENSE` | MIT license. |
| `THIRD-PARTY-NOTICES.md` | Notices for bundled components. |

## Documentation

- [User guide](docs/RecipientLimitReport-UserGuide.md) - also `docs/RecipientLimitReport-UserGuide.html`, a single file to open locally
- [Developer guide](docs/RecipientLimitReport-Guide.md) - also `docs/RecipientLimitReport-Guide.html`

Project page, releases and change log: https://github.com/Nico77600/RecipientLimitReport

License: [MIT](LICENSE). Third-party components: [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md).
