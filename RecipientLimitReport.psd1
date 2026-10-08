#
#  Recipient Limit Report - module manifest
#  --------------------------------------------------------------------------
#  Author  : Nicolas Fabert
#  Version : see ModuleVersion
#
#  Loaded by Invoke-RecipientLimitReport.ps1 (Import-Module by path).
#
@{
    RootModule        = 'RecipientLimitReport.psm1'
    ModuleVersion     = '1.0.1'
    GUID              = '5b0f3c9e-8d7a-4e21-9c4f-2a6e1d8b7f30'
    Author            = 'Nicolas Fabert'
    Description       = 'Recipient Limit Report: collects the Exchange Online message trace (Microsoft Graph) into a local SQLite database, counts the recipients of each message before distribution list expansion and produces CSV/HTML reports of the messages over a recipient limit (one row per Message ID).'
    PowerShellVersion = '7.4'

    # Functions called by Invoke-RecipientLimitReport.ps1 and by the tests. The other functions stay internal
    # to the module: add a function here only when the script or a test calls it.
    FunctionsToExport = @(
        'Import-RlrConfiguration', 'Initialize-RlrEngine', 'Open-RlrStore', 'Enter-RlrLock', 'Exit-RlrLock'
        'Start-RlrLog', 'Stop-RlrLog', 'Write-RlrLog'
        'Write-RlrBanner', 'Write-RlrStep', 'Write-RlrItem', 'Write-RlrSummary'
        'Format-RlrNumber', 'Format-RlrDuration', 'Format-RlrBytes', 'Format-RlrRange', 'Get-RlrScopeText'
        'Get-RlrTimeZone', 'Resolve-RlrPeriod', 'New-RlrRangeList', 'Get-RlrSignatures', 'Get-RlrCollectionPlan', 'Get-RlrCoverage', 'Get-RlrExpiringGaps'
        'Connect-RlrGraph', 'New-RlrCollectorOptions', 'Invoke-RlrCollection', 'Get-RlrCountingEstimate', 'Get-RlrCountingRoutes', 'Invoke-RlrCounting'
        'New-RlrReport', 'Show-RlrStatus', 'Invoke-RlrCompaction', 'Invoke-RlrRetention'
        'Get-RlrIconSet', 'Get-RlrFrameSet'
    )
}
