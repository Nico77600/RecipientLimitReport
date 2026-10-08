#
#  Recipient Limit Report - configuration file
#  --------------------------------------------------------------------------
#  Author  : Nicolas Fabert
#  Version : 1.0.1
#
#  This file is read by Invoke-RecipientLimitReport.ps1. It is a PowerShell data
#  file: text between quotes, $true / $false, numbers, and @( ) for lists.
#  Lines starting with # are comments.
#
#  Relative paths (.\data, .\reports ...) are relative to the tool folder.
#  Some values can be overridden for one execution on the command line
#  (see "Command-line overrides" in docs\RecipientLimitReport-Guide.md).
#
@{
    # ---------------------------------------------------------------------
    # Tenant
    # ---------------------------------------------------------------------
    Tenant = @{
        # Microsoft Entra tenant ID (GUID). Safety check in every mode: after sign-in,
        # the tool stops if the access token belongs to another tenant.
        TenantId     = ''
        # Initial domain (xxx.onmicrosoft.com). Display only.
        Organization = ''
    }

    # ---------------------------------------------------------------------
    # What the report looks for.
    #   RecipientLimit : messages with MORE recipients than this are reported (25 => 26 and more),
    #                    counted as Exchange Online counts them for RecipientLimits and as the DLP
    #                    condition RecipientCountOver does: a distribution list is ONE recipient.
    #   SenderDomains  : senders reported. @() = every sender of the message trace. With the
    #                    accepted domains of the organisation (@('contoso.com', 'fabrikam.com')),
    #                    only the messages sent by the organisation are read: the message trace
    #                    is filtered by Microsoft (one query per domain), which skips the
    #                    incoming Internet mail.
    #   Lowering RecipientLimit later only applies to the data collected afterwards: the
    #   history keeps only the messages over the limit in force when they were collected.
    # ---------------------------------------------------------------------
    Target = @{
        RecipientLimit = 25
        SenderDomains  = @()
    }

    # ---------------------------------------------------------------------
    # Authentication to Microsoft Graph. Permission: ExchangeMessageTrace.Read.All.
    #   Certificate  : app-only, recommended (scheduled task). Application permission + admin consent.
    #   ClientSecret : app-only with a secret read from the environment variable ClientSecretVariable,
    #                  or typed at the prompt (hidden). The secret is never written anywhere.
    #   Interactive  : an administrator signs in (browser, MFA). Delegated permission. Needs the
    #                  Microsoft.Graph.Authentication module (MSAL).
    # The tenant must hold the service principal of the Microsoft application
    # 8bd644d1-64a1-4d4b-ae52-2e0cbf64e373 (guide, chapter 4).
    # ---------------------------------------------------------------------
    Authentication = @{
        Mode                  = 'Certificate'          # Certificate | ClientSecret | Interactive
        AppId                 = ''                     # application (client) ID; Interactive: '' = Microsoft Graph Command Line Tools
        CertificateThumbprint = ''                     # Certificate: in Cert:\CurrentUser\My or Cert:\LocalMachine\My
        ClientSecretVariable  = 'RLR_CLIENT_SECRET'    # ClientSecret: name of the environment variable
        UserPrincipalName     = ''                     # Interactive: expected account ('' = any account of the tenant)
    }

    # ---------------------------------------------------------------------
    # Collection of the message trace (Microsoft Graph, messageTraces).
    # Quota of the API: 100 requests per 5 minutes for the WHOLE tenant (rolling window).
    # Keep a margin for the other tools and administrators.
    # ---------------------------------------------------------------------
    Collection = @{
        PageSize              = 5000   # rows per request (1-5000): 5000 = fewest requests
        MaxConcurrency        = 3      # requests in flight. 2-3 use the whole quota; more does not go faster
        SliceHours            = 2      # length of one query slice (never crosses local midnight); short slices = finer progress and restart
        SettlingHours         = 6      # the last N hours are collected again by the next run (late rows, status changes)
        BackfillDays          = 7      # Collect mode: how many days back to check (only what is missing is collected)
        SourceRetentionDays   = 90     # the message trace keeps 90 days: older periods cannot be collected any more
        RetentionWarningDays  = 7      # warn when a missing period will become uncollectable within N days
        MaxRequests           = 90     # quota used by the tool: MaxRequests per PeriodSeconds
        PeriodSeconds         = 300
        RequestTimeoutSeconds = 180
        MaxRetries            = 5      # per request, for 5xx / timeouts / network errors (429 waits are separate)
    }

    # ---------------------------------------------------------------------
    # Recipient count before distribution list expansion (getDetailsByRecipient).
    # Only for the messages over RecipientLimit AFTER expansion that were sent to at least one
    # distribution list: one request per message, read once and kept in the database.
    # For the messages over RecipientLimit BEFORE expansion, the recipient list as the sender
    # addressed it is then rebuilt: one request per recipient until every recipient the sender
    # addressed is found (the members added by the expansion of the lists are left out).
    # This API has its own quota of 100 requests per 5 minutes for the tenant.
    # ---------------------------------------------------------------------
    Counting = @{
        MaxConcurrency      = 3
        MaxAttempts         = 3        # executions that try a message before it is reported as 'without a count'
        MaxMessagesPerRun   = 0        # 0 = every message still to count; otherwise the rest waits for the next run
        MaxRoutesPerMessage = 250      # requests per message to rebuild its recipient list; above, the list after expansion is shown (0 = never rebuild)
        MaxRequests         = 90
        PeriodSeconds       = 300
    }

    # ---------------------------------------------------------------------
    # Local database (SQLite). Keeps the history beyond the 90 days of the
    # message trace. Once a period has settled, only the messages over
    # RecipientLimit are kept, with their recipients (guide, chapter 9).
    # ---------------------------------------------------------------------
    Storage = @{
        DatabasePath  = '.\data\RecipientLimitReport.sqlite'
        RetentionDays = 180   # messages older than this are deleted at each collection (0 = keep everything)
    }

    # ---------------------------------------------------------------------
    # Report files (CSV and HTML), written locally only.
    # ---------------------------------------------------------------------
    Report = @{
        DefaultRange            = 'Last7Days'      # used when -Range is not given: Last24Hours | Last7Days | Last30Days | PreviousMonth
        TimeZone                = 'Europe/Paris'   # time zone of dates, days, weeks and months in the report
        OutputPath              = '.\reports'      # one sub-folder per execution
        FilePrefix              = 'RecipientLimit'
        Formats                 = @('Csv', 'Html')
        IncludeRecipientDetails = $true            # $true: list of recipient addresses in CSV and HTML; $false: count only
        MaxRecipientsListed     = 500              # addresses listed per message (a list expanded to thousands of members is cut; 0 = all)
        SplitBy                 = 'Week'           # used only above MaxRowsPerFile: Rows | Day | Week (Monday to Sunday)
        MaxRowsPerFile          = 500000           # no split below this number of messages; maximum 1,048,575 (Excel limit)
        CsvDelimiter            = ';'              # ';' opens directly in Excel with French regional settings
        Title                   = 'Messages with more than {0} recipients'   # {0} = RecipientLimit
    }

    # ---------------------------------------------------------------------
    # Log files (one file per day, deleted after RetentionDays).
    # ---------------------------------------------------------------------
    Logging = @{
        Path          = '.\logs'
        RetentionDays = 30
    }
}
