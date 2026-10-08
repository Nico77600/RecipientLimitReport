#Requires -Version 7.4
<#
.SYNOPSIS
    Recipient Limit Report - Exchange Online messages sent to more than N recipients
    (by default 25), counted as Exchange Online counts them: a distribution list is one
    recipient. One row per Message ID. Source: the message trace (Microsoft Graph).

.DESCRIPTION
    The tool works in three stages:

      1. COLLECT  Message trace (Microsoft Graph, messageTraces) -> local SQLite database.
                  Every recipient of every message is read; the message trace keeps
                  90 days. Collect every day (scheduled task, -Mode Collect): the
                  database keeps the history of the messages over the limit.
      2. COUNT    For the messages sent to a distribution list, the number of recipients
                  before expansion is read from the route of the message
                  (getDetailsByRecipient, Submit event): one request per message. For
                  those over the limit, the recipient list as the sender addressed it
                  is rebuilt from the route of each recipient (members added by the
                  expansion are left out), as many requests as needed to find them.
                  The console gives the range of routes and time before it starts.
      3. REPORT   SQLite database -> CSV + HTML files in a local folder, the same report
                  as Purview DLP Report. Before writing, the missing part of the period
                  (if any) is collected and counted.

    Everything is set in config\RecipientLimitReport.config.psd1; the command line only
    chooses what to do. The tool is read-only for the tenant: it never sends e-mail and
    never changes any Microsoft 365 setting.

.PARAMETER Mode
    Report  (default) Collects and counts what is missing for the period, then writes the report.
    Collect           Collects and counts the last days into the database (scheduled task). No report.
    Status            Shows what the database contains, day by day. No connection.

.PARAMETER Range
    Period of the report (Report mode). Default: Report.DefaultRange in the configuration.
      Last24Hours, Last7Days, Last30Days : rolling windows ending now
      PreviousMonth                      : the previous calendar month
      Month  -Month 2026-08              : a calendar month
      Day    -Date 2026-09-28            : a calendar day
      Custom -Start '2026-09-01 08:00' -End '2026-09-03'   (report time zone unless an offset is given)

.PARAMETER ConfigPath
    Configuration file. Default: config\RecipientLimitReport.config.psd1 next to this script.

.PARAMETER IncludeRecipientDetails
    Overrides Report.IncludeRecipientDetails for this execution.
    -IncludeRecipientDetails        : recipient addresses in the files
    -IncludeRecipientDetails:$false : recipient count only (smaller files)

.PARAMETER SplitBy
    Overrides Report.SplitBy (Rows, Day or Week) for this execution.

.PARAMETER MaxRowsPerFile
    Overrides Report.MaxRowsPerFile for this execution.

.PARAMETER OutputPath
    Overrides Report.OutputPath for this execution.

.PARAMETER NoCollect
    Report mode: do not connect to Microsoft 365; use only the data already in the database.

.EXAMPLE
    .\Invoke-RecipientLimitReport.ps1
    Report of the default period (last 7 days), collecting and counting what is missing first.

.EXAMPLE
    .\Invoke-RecipientLimitReport.ps1 -Range PreviousMonth -IncludeRecipientDetails:$false
    Report of the previous month with the recipient count only.

.EXAMPLE
    .\Invoke-RecipientLimitReport.ps1 -Mode Collect
    Daily collection (scheduled task).

.EXAMPLE
    .\Invoke-RecipientLimitReport.ps1 -Mode Status
    What is in the database, day by day, and what is missing.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.1
    Exit codes : 0 = success, 1 = failure, 2 = finished but incomplete (see the summary).
    Documentation : docs\RecipientLimitReport-Guide.md (or .html)
#>
[CmdletBinding()]
param(
    [ValidateSet('Report', 'Collect', 'Status')]
    [string]$Mode = 'Report',

    [ValidateSet('Last24Hours', 'Last7Days', 'Last30Days', 'PreviousMonth', 'Month', 'Day', 'Custom')]
    [string]$Range,
    [string]$Month,
    [string]$Date,
    [string]$Start,
    [string]$End,

    [string]$ConfigPath = (Join-Path $PSScriptRoot 'config\RecipientLimitReport.config.psd1'),
    [switch]$IncludeRecipientDetails,
    [ValidateSet('Rows', 'Day', 'Week')]
    [string]$SplitBy,
    [ValidateRange(1000, 1048575)]
    [int]$MaxRowsPerFile,
    [string]$OutputPath,
    [switch]$NoCollect
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
# Numbers and dates are displayed the same way on every server (1,234.5), whatever the regional settings.
$previousCulture = [Threading.Thread]::CurrentThread.CurrentCulture
[Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo('en-US')
$clock = [Diagnostics.Stopwatch]::StartNew()
$exitCode = 1
$runStatus = 'Failed'
$store = $null; $lock = $null; $runId = $null; $details = [ordered]@{}

try {
    # do { } while ($false): 'break' ends the execution early; the finally block always runs.
    do {
        Import-Module (Join-Path $PSScriptRoot 'RecipientLimitReport.psd1') -Force

        # ---------------------------------------------------------------------------------
        # Configuration file, then the command-line overrides.
        # ---------------------------------------------------------------------------------
        $settings = Import-RlrConfiguration -Path $ConfigPath -Root $PSScriptRoot
        if ($PSBoundParameters.ContainsKey('IncludeRecipientDetails')) { $settings.Report.IncludeRecipientDetails = [bool]$IncludeRecipientDetails }
        if ($SplitBy) { $settings.Report.SplitBy = $SplitBy }
        if ($MaxRowsPerFile) { $settings.Report.MaxRowsPerFile = $MaxRowsPerFile }
        if ($OutputPath) { $settings.Report.OutputPath = [IO.Path]::GetFullPath($OutputPath, (Get-Location).Path) }
        if ($Mode -eq 'Report' -and -not $Range) { $Range = $settings.Report.DefaultRange }
        $zone = $settings.Zone
        $limit = $settings.Target.RecipientLimit
        $domains = [string[]]@($settings.Target.SenderDomains)
        $maxAttempts = $settings.Counting.MaxAttempts

        $logPath = Start-RlrLog -Directory $settings.Logging.Path -RetentionDays $settings.Logging.RetentionDays
        Initialize-RlrEngine -Root $PSScriptRoot

        $collecting = $Mode -eq 'Collect' -or ($Mode -eq 'Report' -and -not $NoCollect)
        $totalSteps = switch ($Mode) { 'Status' { 2 } 'Collect' { 5 } default { 6 } }
        $period = $null
        if ($Mode -eq 'Report') {
            $period = Resolve-RlrPeriod -Range $Range -Month $Month -Date $Date -Start $Start -End $End -Zone $zone
        }
        $dot = [char]0x00B7
        $modeText = $Mode + $(if ($period) { " $dot $Range" } elseif ($Mode -eq 'Collect') { " $dot last $($settings.Collection.BackfillDays) days" } else { '' })
        $banner = [ordered]@{ Mode = @('Info', $modeText) }
        if ($period) { $banner['Period'] = @('Calendar', "$(Format-RlrRange $period.StartMs $period.EndMs $zone)  ($($settings.Report.TimeZone), end excluded)") }
        $banner['Limit'] = @('Target', "more than $limit recipients $dot $(Get-RlrScopeText $settings)")
        $banner['Database'] = @('Database', $settings.Storage.DatabasePath)
        $banner['Log'] = @('Log', $logPath)
        Write-RlrBanner -Title 'Recipient Limit Report' -Subtitle "Exchange Online message trace $dot one row per Message ID" -Details $banner

        # ---------------------------------------------------------------------------------
        # Step 1 - Database (and lock when this execution collects).
        # ---------------------------------------------------------------------------------
        Write-RlrStep 1 $totalSteps 'Opening the database' -Icon Database
        if ($Mode -eq 'Status' -and -not (Test-Path -LiteralPath $settings.Storage.DatabasePath)) {
            # New installation: nothing collected yet. The configuration is valid and the engine is built.
            Write-RlrItem Info 'No database yet: the first collection (-Mode Collect) or the first report creates it.'
            Write-RlrSummary -Title 'Ready for the first collection' -Values ([ordered]@{
                Config   = @('Ok', 'valid')
                Engine   = @('Ok', 'ready')
                Database = @('Database', 'none yet')
                Next     = @('Info', '.\Invoke-RecipientLimitReport.ps1 -Mode Collect')
            }) -Status Ok
            $exitCode = 0
            $runStatus = 'Completed'
            break
        }
        if ($Mode -eq 'Status') {
            $store = Open-RlrStore -Settings $settings -ReadOnly
        } else {
            if ($collecting) { $lock = Enter-RlrLock -Path ($settings.Storage.DatabasePath + '.lock') }
            $store = Open-RlrStore -Settings $settings
            if ($collecting) {
                $closed = $store.CloseAbandonedWork()
                if ($closed) { Write-RlrItem Warn "$closed slice(s) left unfinished by an interrupted execution: their missing part will be collected again." }
            }
            $runId = $store.StartRun($Mode, $null, [Environment]::MachineName, (Get-Module RecipientLimitReport).Version.ToString(), $null)
        }
        $stats = $store.GetStatistics($limit, $maxAttempts)
        Write-RlrItem Ok ("{0} messages over {4} recipients {3} {1} messages stored {3} {2}" -f (Format-RlrNumber $stats.Matches), (Format-RlrNumber $stats.Messages), (Format-RlrBytes $stats.FileBytes), $dot, $limit)

        # ---------------------------------------------------------------------------------
        # Status mode: day-by-day view, then stop.
        # ---------------------------------------------------------------------------------
        if ($Mode -eq 'Status') {
            Write-RlrStep 2 $totalSteps 'Collected data, day by day' -Icon Calendar
            Show-RlrStatus -Store $store -Settings $settings
            $expiring = @(Get-RlrExpiringGaps -Store $store -Settings $settings)
            foreach ($gap in $expiring) { Write-RlrItem Warn ("Missing {0}: the message trace keeps it for about {1} more day(s). Run -Mode Collect or a report of that period." -f (Format-RlrRange $gap.Start $gap.End $zone), $gap.DaysLeft) }
            $exitCode = 0
            $runStatus = 'Completed'
            break
        }

        # ---------------------------------------------------------------------------------
        # Step 2 - What must be collected?
        # ---------------------------------------------------------------------------------
        $nowMs = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()
        if ($Mode -eq 'Collect') {
            Write-RlrStep 2 $totalSteps 'Planning the collection' -Icon Plan
            $windowStart = [Math]::Max($nowMs - [long]$settings.Collection.BackfillDays * 86400000, $nowMs - [long]$settings.Collection.SourceRetentionDays * 86400000 + 3600000)
            $window = [pscustomobject]@{ StartMs = [long]$windowStart; EndMs = [long]$nowMs }
        } else {
            Write-RlrStep 2 $totalSteps 'Checking the data available for the period' -Icon Plan
            $window = $period
        }
        $plan = Get-RlrCollectionPlan -Store $store -StartMs $window.StartMs -EndMs $window.EndMs -Settings $settings -NowMs $nowMs
        $items = $plan.Items
        if ($items.Count) {
            $refresh = if ($plan.RefreshMs -gt 0) { " (of which {0} already collected, refreshed to catch late rows)" -f (Format-RlrDuration ($plan.RefreshMs / 1000)) } else { '' }
            $scope = if (@($plan.Signatures).Count -gt 1) { " for $(@($plan.Signatures).Count) sender domains" } else { '' }
            Write-RlrItem Info ("To collect: {0} slice(s), {1} in total{2}{3}" -f $items.Count, (Format-RlrDuration ($plan.CollectableMs / 1000)), $scope, $refresh)
        } else {
            Write-RlrItem Ok 'Everything is already in the database.'
        }
        foreach ($gap in $plan.Unrecoverable) {
            Write-RlrItem Warn ("Not in the database and older than the {0} days kept by the message trace: {1}" -f $settings.Collection.SourceRetentionDays, (Format-RlrRange $gap.Start $gap.End $zone))
        }
        if ($Mode -eq 'Report' -and $NoCollect -and $plan.NewMs -gt 0) {
            Write-RlrItem Warn '-NoCollect: the missing part will not be collected; the report will be incomplete.'
        } elseif ($Mode -eq 'Report' -and $NoCollect -and $plan.RefreshMs -gt 0) {
            Write-RlrItem Info '-NoCollect: recent data is not refreshed; rows that reached the message trace late may be missing.'
        }
        $countItems = $null
        if ($collecting) { $countItems = $store.GetCountItems($window.StartMs, $window.EndMs, $limit, $maxAttempts, $settings.Counting.MaxMessagesPerRun, $domains) }

        # ---------------------------------------------------------------------------------
        # Step 3 - Connection (only when something must be read from Microsoft 365).
        # ---------------------------------------------------------------------------------
        $connection = $null
        Write-RlrStep 3 $totalSteps 'Connecting to Microsoft Graph' -Icon Key
        if (-not $collecting) {
            Write-RlrItem Skip '-NoCollect: no connection.'
        } else {
            if ($settings.Authentication.Mode -eq 'Interactive') { Write-RlrItem Info 'A sign-in window may open: sign in with the administrator account.' -Icon People }
            # Needed when slices must be collected or counts are still to read (the collection can add
            # more messages sent to distribution lists; they are counted in step 5).
            if (-not $items.Count -and -not $countItems.Count) {
                Write-RlrItem Skip 'Not needed: nothing to collect and nothing to count.'
            } else {
                $connectClock = [Diagnostics.Stopwatch]::StartNew()
                $connection = Connect-RlrGraph -Settings $settings
                $store.UpdateRunAccount($runId, $connection.Account)
                Write-RlrItem Ok ("{0}  {3} {1} mode {3} tenant verified {3} {2}" -f $connection.Account, $connection.Mode.ToLowerInvariant(), (Format-RlrDuration $connectClock.Elapsed.TotalSeconds), $dot)
            }
        }

        # ---------------------------------------------------------------------------------
        # Step 4 - Message trace.
        # ---------------------------------------------------------------------------------
        Write-RlrStep 4 $totalSteps 'Collecting the message trace' -Icon Download
        $collection = $null
        if (-not $collecting) {
            Write-RlrItem Skip '-NoCollect: the report uses the data already collected.'
        } elseif (-not $items.Count) {
            Write-RlrItem Skip 'Nothing to collect.'
        } else {
            $collection = Invoke-RlrCollection -Store $store -Settings $settings -Connection $connection -Items $items -RunId $runId
            $details['Collection'] = $collection
            if ($collection.Throttled) { Write-RlrItem Info ("Microsoft Graph asked to slow down {0} time(s) (429); the requests waited and were sent again." -f $collection.Throttled) }
            if ($collection.Completed) {
                Write-RlrItem Ok ("Collection complete: {0} rows received, {1} new messages, {2} requests, in {3}" -f (Format-RlrNumber $collection.Rows), (Format-RlrNumber $collection.NewMessages), (Format-RlrNumber $collection.Requests), (Format-RlrDuration $collection.Seconds))
            } else {
                Write-RlrItem Fail ("Collection stopped: {0}" -f $collection.Error)
                Write-RlrItem Info 'What was received is kept. Run the same command again to continue from the missing slices.'
            }
        }
        $collectionFailed = $collection -and -not $collection.Completed

        # ---------------------------------------------------------------------------------
        # Step 5 - Recipient count before distribution list expansion.
        # ---------------------------------------------------------------------------------
        Write-RlrStep 5 $totalSteps 'Counting recipients before distribution list expansion' -Icon People
        $counting = $null
        $before = $store.GetPeriodSummary($window.StartMs, $window.EndMs, $limit, $maxAttempts, $domains)
        Write-RlrItem Info ("{0} message(s) over {1} recipients in the trace: {2} without distribution list (counted from the trace), {3} with distribution lists ({4} already counted)" -f (Format-RlrNumber $before.OverLimit), $limit, (Format-RlrNumber ($before.OverLimit - $before.WithLists)), (Format-RlrNumber $before.WithLists), (Format-RlrNumber $before.Measured)) -Icon List
        if (-not $collecting) {
            Write-RlrItem Skip '-NoCollect: no count is read.'
        } elseif ($connection -and -not ($collection -and $collection.Fatal)) {
            $countItems = $store.GetCountItems($window.StartMs, $window.EndMs, $limit, $maxAttempts, $settings.Counting.MaxMessagesPerRun, $domains)
            if (-not $countItems.Count) {
                Write-RlrItem Skip 'Nothing to count.'
            } else {
                $toCount = @($countItems | Where-Object { $null -eq $_.Count }).Count
                $toRebuild = $countItems.Count - $toCount
                $routes = Get-RlrCountingRoutes $countItems
                $routeText = if ($routes.Min -eq $routes.Max) { Format-RlrNumber $routes.Max } else { '{0} to {1}' -f (Format-RlrNumber $routes.Min), (Format-RlrNumber $routes.Max) }
                $timeText = if ($routes.Min -eq $routes.Max) { Format-RlrDuration (Get-RlrCountingEstimate $routes.Max $settings) } else { '{0} to {1}' -f (Format-RlrDuration (Get-RlrCountingEstimate $routes.Min $settings)), (Format-RlrDuration (Get-RlrCountingEstimate $routes.Max $settings)) }
                Write-RlrItem Info ("To read: {0} count(s) before expansion, {1} recipient list(s) to rebuild {5} {2} route(s), {3} within the quota of {4} requests / {6} min" -f (Format-RlrNumber $toCount), (Format-RlrNumber $toRebuild), $routeText, $timeText, $settings.Counting.MaxRequests, $dot, [int]($settings.Counting.PeriodSeconds / 60)) -Icon Clock
                if ($toCount -and $routes.Min -lt $routes.Max) { Write-RlrItem Info 'The upper bound applies to the messages over the limit before expansion: about one route per recipient to rebuild their list.' -Icon Clock }
                $counting = Invoke-RlrCounting -Store $store -Settings $settings -Connection $connection -Items $countItems
                $details['Counting'] = $counting
                $status = if ($counting.Completed -and -not $counting.Failed) { 'Ok' } else { 'Warn' }
                Write-RlrItem $status ("Read {0} of {1} message(s) in {2}: {3} with a count, {4} without a count in their route, {5} failed (tried again at the next run) {7} {6} request(s)" -f (Format-RlrNumber ($counting.Done + $counting.NoCount + $counting.Failed)), (Format-RlrNumber $counting.Items), (Format-RlrDuration $counting.Seconds), (Format-RlrNumber $counting.Done), (Format-RlrNumber $counting.NoCount), (Format-RlrNumber $counting.Failed), (Format-RlrNumber $counting.Requests), $dot)
                if ($counting.ListsDone -or $counting.ListsPartial) {
                    Write-RlrItem $(if ($counting.ListsPartial) { 'Warn' } else { 'Ok' }) ("Recipient lists before expansion: {0} rebuilt{1}" -f (Format-RlrNumber $counting.ListsDone), $(if ($counting.ListsPartial) { ", $(Format-RlrNumber $counting.ListsPartial) left after expansion (more than $($settings.Counting.MaxRoutesPerMessage) routes to read, Counting.MaxRoutesPerMessage)" } else { '' })) -Icon People
                }
                if ($counting.Fatal) { Write-RlrItem Fail $counting.Fatal }
            }
        } else {
            Write-RlrItem Skip 'Not possible: no connection.'
        }
        if ($collecting) {
            $compacted = Invoke-RlrCompaction -Store $store -Settings $settings
            if ($compacted) { Write-RlrItem Info ("History: {0} message(s) with {1} recipients or fewer removed from the settled period (only the messages over the limit are kept)." -f (Format-RlrNumber $compacted), $limit) }
            $purge = Invoke-RlrRetention -Store $store -Settings $settings
            if ($purge -and $purge.Messages) { Write-RlrItem Info ("Retention ({0} days): {1} old messages deleted." -f $settings.Storage.RetentionDays, (Format-RlrNumber $purge.Messages)) }
        }
        $after = $store.GetPeriodSummary($window.StartMs, $window.EndMs, $limit, $maxAttempts, $domains)
        $countingFailed = $counting -and $counting.Fatal

        # ---------------------------------------------------------------------------------
        # Collect mode ends here.
        # ---------------------------------------------------------------------------------
        if ($Mode -eq 'Collect') {
            $expiring = @(Get-RlrExpiringGaps -Store $store -Settings $settings)
            foreach ($gap in $expiring) { Write-RlrItem Warn ("Still missing {0}: collectable for about {1} more day(s)." -f (Format-RlrRange $gap.Start $gap.End $zone), $gap.DaysLeft) }
            $stats = $store.GetStatistics($limit, $maxAttempts)
            $failed = $collectionFailed -or $countingFailed
            $summary = [ordered]@{
                'Result'   = @($(if ($failed) { 'Fail' } else { 'Ok' }), $(if ($failed) { 'Incomplete - run again to continue' } else { 'Collection complete' }))
                'Trace'    = @('Download', $(if ($collection) { '{0} rows received, {1} new messages' -f (Format-RlrNumber $collection.Rows), (Format-RlrNumber $collection.NewMessages) } else { 'nothing to collect' }))
                'Counts'   = @('People', $(if ($counting) { '{0} read, {1} still to read {3} {2} recipient list(s) to rebuild' -f (Format-RlrNumber $counting.Done), (Format-RlrNumber $after.Pending), (Format-RlrNumber $after.ListsPending), $dot } else { '{0} still to read {2} {1} recipient list(s) to rebuild' -f (Format-RlrNumber $after.Pending), (Format-RlrNumber $after.ListsPending), $dot }))
                'Database' = @('Database', ('{0} messages over {3} recipients {2} {1}' -f (Format-RlrNumber $stats.Matches), (Format-RlrBytes $stats.FileBytes), $dot, $limit))
                'Duration' = @('Clock', (Format-RlrDuration $clock.Elapsed.TotalSeconds))
                'Log'      = @('Log', $logPath)
            }
            $status = if ($failed) { 'Fail' } elseif ($expiring.Count -or $after.Pending -or $after.ListsPending) { 'Warn' } else { 'Ok' }
            Write-RlrSummary -Title $(if ($failed) { 'Collection incomplete' } else { 'Collection finished' }) -Values $summary -Status $status
            $exitCode = if ($failed) { 1 } else { 0 }
            $runStatus = if ($failed) { 'Incomplete' } else { 'Completed' }
            break
        }

        # ---------------------------------------------------------------------------------
        # Step 6 - Report files.
        # ---------------------------------------------------------------------------------
        Write-RlrStep 6 $totalSteps 'Writing the report' -Icon Report
        $coverage = Get-RlrCoverage -Store $store -Settings $settings -StartMs $period.StartMs -EndMs $period.EndMs
        $reportClock = [Diagnostics.Stopwatch]::StartNew()
        $unmeasured = $after.Pending + $after.Unmeasurable
        $report = New-RlrReport -Store $store -Settings $settings -Period $period -Coverage $coverage -Unmeasured $unmeasured
        foreach ($file in $report.Result.Files) {
            Write-RlrItem Ok ("{0,-4} {1}   {2} messages {4} {3}" -f $file.Kind, (Split-Path $file.Path -Leaf), (Format-RlrNumber $file.Rows), (Format-RlrBytes $file.Bytes), $dot) -Icon File
        }
        Write-RlrItem Info ("Written in {0}" -f (Format-RlrDuration $reportClock.Elapsed.TotalSeconds)) -Icon Clock
        $r = $report.Result
        $complete = $coverage.Percent -ge 100 -and -not $collectionFailed -and -not $countingFailed -and $after.Pending -eq 0 -and $after.ListsPending -eq 0
        # Messages with lists are counted before the compaction of step 5 (those found under the limit are then removed).
        $withLists = [Math]::Max($before.WithLists, $after.WithLists)
        $countText = if (-not $withLists) { 'no message with distribution lists over the limit' }
            elseif (-not $after.Pending -and -not $after.Unmeasurable) { 'all {0} message(s) with distribution lists counted' -f (Format-RlrNumber $withLists) }
            else { '{0} of {1} message(s) with distribution lists counted - {2} still to count{3}' -f (Format-RlrNumber ($withLists - $after.Pending - $after.Unmeasurable)), (Format-RlrNumber $withLists), (Format-RlrNumber $after.Pending), $(if ($after.Unmeasurable) { ", $(Format-RlrNumber $after.Unmeasurable) without a count" } else { '' }) }
        if ($settings.Report.IncludeRecipientDetails -and ($after.ListsRebuilt + $after.ListsPending + $after.ListsNotRebuilt)) {
            $countText += " $dot recipient lists before expansion: {0} rebuilt" -f (Format-RlrNumber $after.ListsRebuilt)
            if ($after.ListsPending) { $countText += ', {0} still to rebuild' -f (Format-RlrNumber $after.ListsPending) }
            if ($after.ListsNotRebuilt) { $countText += ', {0} listed after expansion' -f (Format-RlrNumber $after.ListsNotRebuilt) }
        }
        $summary = [ordered]@{
            'Period'   = @('Calendar', "$(Format-RlrRange $period.StartMs $period.EndMs $zone)  ($($settings.Report.TimeZone))")
            'Messages' = @('Mail', ('{0} unique messages over {3} recipients {2} {1} senders' -f (Format-RlrNumber $r.Messages), (Format-RlrNumber $r.Senders), $dot, $limit))
            'Coverage' = @($(if ($coverage.Percent -ge 100) { 'Chart' } else { 'Warn' }), $(if ($coverage.Percent -ge 100) { '100% of the period' } else { '{0}% of the period - missing: {1}' -f $coverage.Percent, ((@($coverage.Gaps) | Select-Object -First 3 | ForEach-Object { Format-RlrRange $_.Start $_.End $zone }) -join ', ') }))
            'Lists'    = @($(if ($after.Pending -or $after.Unmeasurable -or $after.ListsPending) { 'Warn' } else { 'People' }), $countText)
            'Files'    = @('File', ('{0} file(s) {2} recipients {1}' -f $r.Files.Count, $(if ($settings.Report.IncludeRecipientDetails) { 'listed' } else { 'counted only' }), $dot))
            'Folder'   = @('Folder', $report.Directory)
            'Duration' = @('Clock', (Format-RlrDuration $clock.Elapsed.TotalSeconds))
            'Log'      = @('Log', $logPath)
        }
        Write-RlrSummary -Title $(if ($complete) { 'Report ready' } else { 'Report written, but INCOMPLETE' }) -Values $summary -Status $(if ($complete) { 'Ok' } else { 'Warn' })
        $details['Report'] = [ordered]@{ Directory = $report.Directory; Messages = $r.Messages; Files = @($r.Files | ForEach-Object { $_.Path }); CoveragePercent = $coverage.Percent; PendingCounts = $after.Pending; Unmeasurable = $after.Unmeasurable; PendingLists = $after.ListsPending; ListedAfterExpansion = $r.ListedAfterExpansion }
        $exitCode = if ($complete) { 0 } else { 2 }
        $runStatus = if ($complete) { 'Completed' } else { 'Incomplete' }
    } while ($false)
}
catch {
    $message = $_.Exception.Message
    Write-Progress -Id 1 -Activity 'Collecting the message trace' -Completed
    Write-Host ''
    if (Get-Command Write-RlrSummary -ErrorAction SilentlyContinue) {
        Write-RlrSummary -Title 'Execution stopped' -Values ([ordered]@{ Error = @('Fail', $message); Duration = @('Clock', (Format-RlrDuration $clock.Elapsed.TotalSeconds)) }) -Status Fail
    } else {
        Write-Host "  [ERROR] $message"
    }
    if (Get-Command Write-RlrLog -ErrorAction SilentlyContinue) { Write-RlrLog -Level 'ERROR' -Message ($message + "`n" + $_.ScriptStackTrace) }
    $details['Error'] = $message
    $exitCode = 1
}
finally {
    if ($store -and $runId) {
        try { $store.FinishRun($runId, $runStatus, ($details | ConvertTo-Json -Depth 6 -Compress)) } catch { }
    }
    if ($store) { $store.Dispose() }
    if (Get-Command Exit-RlrLock -ErrorAction SilentlyContinue) { Exit-RlrLock -Lock $lock }
    if (Get-Command Write-RlrLog -ErrorAction SilentlyContinue) {
        Write-RlrLog -Level 'INFO' -Message ("Exit code {0} after {1:0.0} s" -f $exitCode, $clock.Elapsed.TotalSeconds)
        Stop-RlrLog
    }
    [Threading.Thread]::CurrentThread.CurrentCulture = $previousCulture
}
exit $exitCode
