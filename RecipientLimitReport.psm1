#Requires -Version 7.4
<#
.SYNOPSIS
    Recipient Limit Report - PowerShell module.

.DESCRIPTION
    Helper functions used by Invoke-RecipientLimitReport.ps1. The module is organised
    in regions, in the order of an execution:

        1. Console and log        Write-Rlr* functions (what the administrator sees)
        2. Configuration          Import-RlrConfiguration (reads and checks the .psd1 file)
        3. Engine                 Initialize-RlrEngine (SQLite + compiled C# engine)
        4. Periods and coverage   Resolve-RlrPeriod, Get-RlrCollectionPlan, Get-RlrCoverage
        5. Connection             Connect-RlrGraph / Update-RlrToken (Microsoft Graph token)
        6. Collection             Invoke-RlrCollection (message trace -> SQLite)
        7. Recipient count        Invoke-RlrCounting (count and recipient list before distribution list expansion)
        8. Report                 New-RlrReport (SQLite -> CSV / HTML)
        9. Status and maintenance Show-RlrStatus, Invoke-RlrCompaction, Invoke-RlrRetention, Enter-RlrLock

    Performance-critical work (HTTP requests, JSON parsing, database, file writing) is done
    by src\RecipientLimitReport.Engine.cs, compiled on first use.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.1
    History : see CHANGELOG.md
#>
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:ToolVersion = '1.0.1'
$script:ToolRoot = $PSScriptRoot
$script:GraphRoot = 'https://graph.microsoft.com/v1.0'
# Microsoft Graph permission of the message trace API (application or delegated).
$script:Permission = 'ExchangeMessageTrace.Read.All'
$script:LogWriter = $null
$script:LogPath = $null
# ---------------------------------------------------------------------------------------------
# Console theme: colours and icons.
#   - Colours (ANSI) are disabled when the output is redirected (scheduled task, log capture)
#     or when the NO_COLOR environment variable is set; RLR_FORCE_COLOR=1 forces them.
#   - Icons: emoji in modern terminals (Windows Terminal, VS Code), simple symbols elsewhere.
#     Emoji are chosen among those always two columns wide (no variation selector), so frames
#     and columns stay aligned. The classic console (conhost) has no font fallback: a character
#     missing from its font is shown as an empty box. The symbols used outside the modern
#     terminals are therefore all present in Consolas and Lucida Console (code page 437
#     repertoire or Latin-1) - a test checks it. Force a style with the environment
#     variable RLR_ICONS = Emoji | Symbols | Ascii.
#   - Frames: rounded corners, present in Consolas (default font of the console). Lucida Console
#     and the raster font have no rounded corners: square corners are used with them (the
#     console font is read when the frame is drawn, see Get-RlrFrame).
# ---------------------------------------------------------------------------------------------
$script:C = @{ Reset = ''; Bold = ''; Dim = ''; Accent = ''; AccentBg = ''; Cyan = ''; Green = ''; Yellow = ''; Red = ''; White = '' }
if ($env:RLR_FORCE_COLOR -eq '1' -or (-not [Console]::IsOutputRedirected -and -not $env:NO_COLOR)) {
    $e = [char]27
    $script:C = @{
        Reset = "$e[0m"; Bold = "$e[1m"; Dim = "$e[90m"; White = "$e[97m"
        Accent = "$e[38;2;214;62;115m"; AccentBg = "$e[48;2;177;31;75m$e[97m"
        Cyan = "$e[38;2;97;214;214m"; Green = "$e[38;2;80;200;120m"; Yellow = "$e[38;2;240;200;90m"; Red = "$e[38;2;240;90;90m"
    }
}
$script:IconStyle = if ($env:RLR_ICONS -in 'Emoji', 'Symbols', 'Ascii') { $env:RLR_ICONS }
    elseif ([Console]::IsOutputRedirected) { 'Symbols' }
    elseif ($env:WT_SESSION -or $env:TERM_PROGRAM -eq 'vscode') { 'Emoji' }
    else { 'Symbols' }

function Get-RlrIconSet {
    <# Icons of one console style. Symbols: only characters of the classic console fonts. #>
    param([Parameter(Mandatory)][ValidateSet('Emoji', 'Symbols', 'Ascii')][string]$Style)
    $u = { param([int]$Code) [char]::ConvertFromUtf32($Code) }
    switch ($Style) {
        'Emoji' { return @{
                Logo = & $u 0x1F4EC; Ok = & $u 0x2705; Warn = (& $u 0x26A0) + [char]0xFE0F
                Fail = & $u 0x274C; Info = & $u 0x1F539; Skip = & $u 0x23E9
                Database = & $u 0x1F4BE; Plan = & $u 0x1F50E; Key = & $u 0x1F510
                Download = & $u 0x1F4E5; Report = & $u 0x1F4CA; Calendar = & $u 0x1F4C5
                File = & $u 0x1F4C4; Folder = & $u 0x1F4C1; Clock = & $u 0x23F3
                Mail = & $u 0x1F4E8; Target = & $u 0x1F3AF; Log = & $u 0x1F4DD
                Done = & $u 0x1F389; People = & $u 0x1F465; Chart = & $u 0x1F4C8; List = & $u 0x1F4CB
            } }
        'Symbols' { return @{
                Logo = & $u 0x2666; Ok = & $u 0x221A; Warn = & $u 0x25B2; Fail = & $u 0x00D7; Info = & $u 0x2022
                Skip = & $u 0x00BB; Database = & $u 0x25A0; Plan = & $u 0x25BA; Key = & $u 0x2194; Download = & $u 0x2193
                Report = & $u 0x2261; Calendar = & $u 0x263C; File = & $u 0x25AC; Folder = & $u 0x2302; Clock = & $u 0x25CB
                Mail = '@'; Target = & $u 0x25D9; Log = & $u 0x00B6; Done = & $u 0x221A; People = & $u 0x2192; Chart = & $u 0x2191; List = & $u 0x2261
            } }
        default { return @{
                Logo = '*'; Ok = '+'; Warn = '!'; Fail = 'x'; Info = '-'; Skip = '>'; Database = '#'; Plan = '?'; Key = '@'; Download = 'v'; Report = '='; Calendar = ':'
                File = '-'; Folder = '>'; Clock = '~'; Mail = '@'; Target = 'o'; Log = '='; Done = '*'; People = '&'; Chart = '^'; List = '='
            } }
    }
}

function Get-RlrFrameSet {
    <#
    .SYNOPSIS
        Frame characters. Rounded corners, except with the Ascii style and with the console fonts
        that have no rounded corners (Lucida Console, raster font 'Terminal'): square corners.
    #>
    param([Parameter(Mandatory)][ValidateSet('Emoji', 'Symbols', 'Ascii')][string]$Style, [AllowNull()][string]$FontName)
    if ($Style -eq 'Ascii') { return @{ TopLeft = [char]'+'; TopRight = [char]'+'; BottomLeft = [char]'+'; BottomRight = [char]'+'; Horizontal = [char]'-'; Vertical = [char]'|' } }
    if ($Style -eq 'Symbols' -and $FontName -in 'Lucida Console', 'Terminal') {
        return @{ TopLeft = [char]0x250C; TopRight = [char]0x2510; BottomLeft = [char]0x2514; BottomRight = [char]0x2518; Horizontal = [char]0x2500; Vertical = [char]0x2502 }
    }
    return @{ TopLeft = [char]0x256D; TopRight = [char]0x256E; BottomLeft = [char]0x2570; BottomRight = [char]0x256F; Horizontal = [char]0x2500; Vertical = [char]0x2502 }
}

function Get-RlrFrame {
    <# Frame characters for this console. In the classic console the font is read once, through the engine. #>
    if ($script:Frame) { return $script:Frame }
    $engine = [bool]('RecipientLimitReport.ConsoleFont' -as [type])
    $font = if ($script:IconStyle -eq 'Symbols' -and $engine) { [RecipientLimitReport.ConsoleFont]::FaceName() }
    $frame = Get-RlrFrameSet $script:IconStyle $font
    # Before the engine is loaded (error at start) the font is unknown: not kept, read again later.
    if ($script:IconStyle -ne 'Symbols' -or $engine) { $script:Frame = $frame }
    return $frame
}

$script:Icons = Get-RlrIconSet $script:IconStyle
$script:Frame = $null
# Emoji are two columns wide in the console; symbols are one: pad symbols so text stays aligned.
$script:IconPad = if ($script:IconStyle -eq 'Emoji') { ' ' } else { '  ' }
$script:Dot = [char]0x00B7

#region 1. Console and log ---------------------------------------------------------------

function Get-RlrIcon {
    param([Parameter(Mandatory)][string]$Name)
    return $script:Icons[$Name] + $script:IconPad
}

function Format-RlrNumber {
    param([Parameter(Mandatory)][AllowNull()]$Value)
    if ($null -eq $Value) { return '-' }
    return ([long]$Value).ToString('N0', [Globalization.CultureInfo]::GetCultureInfo('en-US'))
}

function Format-RlrDuration {
    param([Parameter(Mandatory)][double]$Seconds)
    # 0.0 (not 0): with an integer first argument PowerShell picks Math.Max(int, int) and drops the decimals.
    $t = [TimeSpan]::FromTicks([long]([Math]::Max(0.0, $Seconds) * 10000000))
    if ($t.TotalDays -ge 2) { return '{0} d {1:00} h' -f [int][Math]::Floor($t.TotalDays), $t.Hours }
    if ($t.TotalHours -ge 1) { return '{0} h {1:00} min' -f [int][Math]::Floor($t.TotalHours), $t.Minutes }
    if ($t.TotalMinutes -ge 1) { return '{0} min {1:00} s' -f $t.Minutes, $t.Seconds }
    return '{0:0.0} s' -f $t.TotalSeconds
}

function Format-RlrBytes {
    param([Parameter(Mandatory)][double]$Bytes)
    if ($Bytes -ge 1GB) { return '{0:0.00} GB' -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return '{0:0.0} MB' -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return '{0:0} KB' -f ($Bytes / 1KB) }
    return '{0:0} B' -f $Bytes
}

function Format-RlrLocalTime {
    param([Parameter(Mandatory)][long]$UnixMs, [Parameter(Mandatory)][TimeZoneInfo]$Zone, [string]$Format = 'yyyy-MM-dd HH:mm')
    return [TimeZoneInfo]::ConvertTime([DateTimeOffset]::FromUnixTimeMilliseconds($UnixMs), $Zone).ToString($Format, [Globalization.CultureInfo]::InvariantCulture)
}

function Format-RlrRange {
    param([long]$StartMs, [long]$EndMs, [TimeZoneInfo]$Zone)
    return '{0} {2} {1}' -f (Format-RlrLocalTime $StartMs $Zone), (Format-RlrLocalTime $EndMs $Zone), [char]0x2192
}

function Start-RlrLog {
    <# Opens (or continues) today's log file and deletes log files older than the retention. #>
    param([Parameter(Mandatory)][string]$Directory, [int]$RetentionDays = 30)
    [void][IO.Directory]::CreateDirectory($Directory)
    $script:LogPath = Join-Path $Directory ('RecipientLimitReport_{0:yyyyMMdd}.log' -f (Get-Date))
    $stream = [IO.FileStream]::new($script:LogPath, [IO.FileMode]::Append, [IO.FileAccess]::Write, [IO.FileShare]::ReadWrite)
    $script:LogWriter = [IO.StreamWriter]::new($stream, [Text.UTF8Encoding]::new($false))
    $script:LogWriter.AutoFlush = $true
    $limit = (Get-Date).AddDays(-$RetentionDays)
    Get-ChildItem -LiteralPath $Directory -Filter 'RecipientLimitReport_*.log' -File -ErrorAction SilentlyContinue |
        Where-Object LastWriteTime -lt $limit | Remove-Item -Force -ErrorAction SilentlyContinue
    return $script:LogPath
}

function Stop-RlrLog {
    if ($script:LogWriter) { $script:LogWriter.Dispose(); $script:LogWriter = $null }
}

function Write-RlrLog {
    <# Writes one line to the log file only (never to the console). The log never contains colours or icons. #>
    param([ValidateSet('INFO', 'OK', 'WARN', 'ERROR', 'STEP', 'DEBUG')][string]$Level = 'INFO', [Parameter(Mandatory)][AllowEmptyString()][string]$Message)
    if ($script:LogWriter) {
        $script:LogWriter.WriteLine(('{0:yyyy-MM-ddTHH:mm:ss.fffzzz} [{1,-5}] {2}' -f (Get-Date), $Level, $Message))
    }
}

function Write-RlrBanner {
    <#
    .SYNOPSIS
        Title card at the start of an execution:

          ╭──────────────────────────────────────────────────────────────────────╮
          │  📬  Recipient Limit Report                v1.0.1 · Nicolas Fabert   │
          │     Exchange Online message trace · one row per Message ID           │
          ╰──────────────────────────────────────────────────────────────────────╯
             📅  Period     2026-09-22 19:58 → 2026-09-29 19:58 ...
    .PARAMETER Details
        Ordered list of rows: key = label, value = @(IconName, Text) or plain text.
    #>
    param([Parameter(Mandatory)][string]$Title, [string]$Subtitle, [System.Collections.Specialized.OrderedDictionary]$Details)
    $C = $script:C; $F = Get-RlrFrame; $width = 74
    $right = "v$($script:ToolVersion) $([char]0x00B7) Nicolas Fabert"
    $iconWidth = if ($script:IconStyle -eq 'Emoji') { 2 } else { 1 }
    $left = "  $($script:Icons.Logo)  $Title"
    $gap = [Math]::Max(1, $width - ($left.Length - $script:Icons.Logo.Length + $iconWidth) - $right.Length - 2)
    Write-Host ''
    Write-Host ("  {0}{1}{2}{3}{4}" -f $C.Accent, $F.TopLeft, [string]::new($F.Horizontal, $width), $F.TopRight, $C.Reset)
    Write-Host ("  {0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}" -f $C.Accent, $F.Vertical, $C.Reset, $C.Bold, $left, $C.Reset, [string]::new(' ', $gap), $C.Dim, $right, '  ', ($C.Accent + $F.Vertical), $C.Reset)
    if ($Subtitle) {
        $sub = "     $Subtitle"
        Write-Host ("  {0}{1}{2}{3}{4}{5}{0}{6}{2}" -f $C.Accent, $F.Vertical, $C.Reset, $C.Dim, $sub.PadRight($width), $C.Reset, $F.Vertical)
    }
    Write-Host ("  {0}{1}{2}{3}{4}" -f $C.Accent, $F.BottomLeft, [string]::new($F.Horizontal, $width), $F.BottomRight, $C.Reset)
    if ($Details) {
        foreach ($key in $Details.Keys) {
            $value = $Details[$key]
            $icon, $text = if ($value -is [array]) { (Get-RlrIcon $value[0]), $value[1] } else { '   ', $value }
            Write-Host ("     {0}{1}{2,-10}{3} {4}" -f $icon, $C.Dim, $key, $C.Reset, $text)
        }
    }
    Write-RlrLog 'STEP' "=== $Title v$($script:ToolVersion) ==="
    if ($Details) { foreach ($key in $Details.Keys) { $v = $Details[$key]; Write-RlrLog 'INFO' ("{0}: {1}" -f $key, $(if ($v -is [array]) { $v[1] } else { $v })) } }
}

function Write-RlrStep {
    <#
    .SYNOPSIS
        Step header with a coloured number pill and an icon, e.g.

          ─── 3/6 ─ 🔐  Connecting to Microsoft Graph
    #>
    param([Parameter(Mandatory)][int]$Number, [Parameter(Mandatory)][int]$Total, [Parameter(Mandatory)][string]$Title, [string]$Icon = 'Info')
    $C = $script:C
    Write-Host ''
    Write-Host ("  {0} {1}/{2} {3} {4}{5}{6}{3}" -f $C.AccentBg, $Number, $Total, $C.Reset, (Get-RlrIcon $Icon), $C.Bold, $Title)
    Write-RlrLog 'STEP' "[$Number/$Total] $Title"
}

function Write-RlrItem {
    <# One indented result line with a status icon, also written to the log. #>
    param([ValidateSet('Ok', 'Warn', 'Fail', 'Info', 'Skip')][string]$Status = 'Info', [Parameter(Mandatory)][AllowEmptyString()][string]$Text, [string]$Icon)
    $color = @{ Ok = $script:C.Green; Warn = $script:C.Yellow; Fail = $script:C.Red; Info = ''; Skip = $script:C.Dim }[$Status]
    $level = @{ Ok = 'OK'; Warn = 'WARN'; Fail = 'ERROR'; Info = 'INFO'; Skip = 'INFO' }[$Status]
    $symbol = Get-RlrIcon $(if ($Icon) { $Icon } else { $Status })
    $textColor = if ($Status -in 'Warn', 'Fail', 'Skip') { $color } else { '' }
    Write-Host ("      {0}{1}{2}{3}{4}{2}" -f $color, $symbol, $script:C.Reset, $textColor, $Text)
    Write-RlrLog $level $Text
}

function Write-RlrTableRow {
    <#
    .SYNOPSIS
        One aligned row of the collection table (one row per message trace slice).
        Columns: status icon, slice, rows, messages, pages, duration, rate (and the sender
        domain when several domains are collected).
    #>
    param([switch]$Header, [ValidateSet('Ok', 'Warn', 'Fail')][string]$Status = 'Ok', [string]$Window, $Rows, $Messages, $Pages, [string]$Duration, [string]$Rate, [string]$Scope)
    $C = $script:C
    if ($Header) {
        Write-Host ("      {0}{1}{2,-35} {3,10} {4,10} {5,6}  {6,12}  {7,12}{8}" -f $C.Dim, ('  ' + $script:IconPad), 'Slice', 'Rows', 'Messages', 'Pages', 'Duration', 'Rate', $C.Reset)
        return
    }
    $color = @{ Ok = $C.Green; Warn = $C.Yellow; Fail = $C.Red }[$Status]
    $rowsText = Format-RlrNumber $Rows
    $rowsColor = if ([long]$Rows -gt 0) { $C.White + $C.Bold } else { $C.Dim }
    $scopeText = if ($Scope) { "  $($C.Dim)$Scope$($C.Reset)" } else { '' }
    Write-Host ("      {0}{1}{2}{3,-35} {4}{5,10}{2} {6,10} {7,6}  {8,12}  {9}{10,12}{2}{11}" -f $color, (Get-RlrIcon $Status), $C.Reset, $Window, $rowsColor, $rowsText, (Format-RlrNumber $Messages), $Pages, $Duration, $C.Dim, $Rate, $scopeText)
    Write-RlrLog $(if ($Status -eq 'Ok') { 'OK' } else { 'WARN' }) ("{0}{6}  {1} rows, {2} messages, {3} pages, {4}, {5}" -f $Window, $rowsText, (Format-RlrNumber $Messages), $Pages, $Duration, $Rate, $(if ($Scope) { " [$Scope]" } else { '' }))
}

function Write-RlrSummary {
    <#
    .SYNOPSIS
        Final summary card:

          ╭─ 🎉  Report ready ─────────────────────────────────────────────────╮
            📅  Period      ...
            ✉️  Messages    ...
          ╰────────────────────────────────────────────────────────────────────╯
    .PARAMETER Values
        Ordered list: key = label, value = @(IconName, Text) or plain text.
    #>
    param([Parameter(Mandatory)][string]$Title, [Parameter(Mandatory)][System.Collections.Specialized.OrderedDictionary]$Values, [ValidateSet('Ok', 'Warn', 'Fail')][string]$Status = 'Ok')
    $C = $script:C; $F = Get-RlrFrame; $width = 74
    $color = @{ Ok = $C.Green; Warn = $C.Yellow; Fail = $C.Red }[$Status]
    $icon = $script:Icons[@{ Ok = 'Done'; Warn = 'Warn'; Fail = 'Fail' }[$Status]]
    $iconWidth = if ($script:IconStyle -eq 'Emoji') { 2 } else { 1 }
    $head = " $icon  $Title "
    $rest = [Math]::Max(2, $width - 1 - ($head.Length - $icon.Length + $iconWidth))
    Write-Host ''
    Write-Host ("  {0}{1}{2}{3}{4}{0}{5}{6}{7}" -f $color, $F.TopLeft, $F.Horizontal, $C.Bold, $head, ($C.Reset + $color), ([string]::new($F.Horizontal, $rest) + $F.TopRight), $C.Reset)
    foreach ($key in $Values.Keys) {
        $value = $Values[$key]
        $rowIcon, $text = if ($value -is [array]) { (Get-RlrIcon $value[0]), $value[1] } else { '   ', $value }
        Write-Host ("    {0}{1}{2,-10}{3} {4}" -f $rowIcon, $C.Dim, $key, $C.Reset, $text)
        Write-RlrLog 'INFO' ("Summary - {0}: {1}" -f $key, $text)
    }
    Write-Host ("  {0}{1}{2}{3}{4}" -f $color, $F.BottomLeft, [string]::new($F.Horizontal, $width), $F.BottomRight, $C.Reset)
    Write-Host ''
}

function Write-RlrEngineEvents {
    <# Prints the events queued by the engine (429, retries, failed slices). Identical messages within 15 seconds are printed once. #>
    param([Parameter(Mandatory)]$Queue, [hashtable]$Seen = @{})
    $event = $null
    while ($Queue.TryDequeue([ref]$event)) {
        $key = $event.Text -replace '\d+', '#'
        $now = [Environment]::TickCount64
        if ($Seen.ContainsKey($key) -and $now - $Seen[$key] -lt 15000) { Write-RlrLog $(if ($event.Level -eq 'ERROR') { 'ERROR' } else { 'WARN' }) $event.Text; continue }
        $Seen[$key] = $now
        $status = switch ($event.Level) { 'ERROR' { 'Fail' } 'WARN' { 'Warn' } default { 'Info' } }
        Write-RlrItem $status $event.Text
    }
}

#endregion

#region 2. Configuration ------------------------------------------------------------------

function Resolve-RlrPath {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Root)
    $expanded = [Environment]::ExpandEnvironmentVariables($Path)
    if ([IO.Path]::IsPathRooted($expanded)) { return [IO.Path]::GetFullPath($expanded) }
    return [IO.Path]::GetFullPath((Join-Path $Root $expanded))
}

function Import-RlrConfiguration {
    <#
    .SYNOPSIS
        Reads the configuration file, checks every value and returns it with absolute paths.
        All problems are reported together so the administrator can fix them in one go.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [string]$Root = $script:ToolRoot)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Configuration file not found: $Path" }
    try { $config = Import-PowerShellDataFile -LiteralPath $Path }
    catch { throw "The configuration file is not valid PowerShell data ($Path): $($_.Exception.Message)" }

    $errors = [Collections.Generic.List[string]]::new()
    $guid = '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$'
    $domainPattern = '^[a-z0-9](?:[a-z0-9-]{0,62}[a-z0-9])?(?:\.[a-z0-9](?:[a-z0-9-]{0,62}[a-z0-9])?)+$'
    foreach ($section in 'Tenant', 'Target', 'Authentication', 'Collection', 'Counting', 'Storage', 'Report', 'Logging') {
        if (-not $config.ContainsKey($section) -or $config[$section] -isnot [hashtable]) { $errors.Add("Section '$section' is missing.") }
    }
    if ($errors.Count) { throw ("Invalid configuration ($Path):`n - " + ($errors -join "`n - ")) }

    function Get-Value([hashtable]$Section, [string]$SectionName, [string]$Key, $Default, [switch]$Required) {
        if ($Section.ContainsKey($Key) -and $null -ne $Section[$Key] -and "$($Section[$Key])" -ne '') { return $Section[$Key] }
        if ($Required) { $errors.Add("$SectionName.$Key is required.") }
        return $Default
    }
    function Test-Int($Value, [string]$Name, [long]$Min, [long]$Max) {
        $n = 0L
        if (-not [long]::TryParse("$Value", [ref]$n) -or $n -lt $Min -or $n -gt $Max) { $errors.Add("$Name must be a whole number between $Min and $Max (current value: '$Value').") }
        return [int]$n
    }
    function Test-Bool($Value, [string]$Name) {
        if ($Value -isnot [bool]) { $errors.Add("$Name must be `$true or `$false (current value: '$Value').") ; return $false }
        return $Value
    }

    $t = $config.Tenant; $g = $config.Target; $a = $config.Authentication; $c = $config.Collection; $n = $config.Counting
    $s = $config.Storage; $r = $config.Report; $l = $config.Logging
    # Sender domains: 'contoso.com', '@contoso.com' and '*@contoso.com' are the same.
    $domains = [Collections.Generic.List[string]]::new()
    foreach ($d in @(Get-Value $g 'Target' 'SenderDomains' @())) {
        $value = ([string]$d).Trim().TrimStart('*').TrimStart('@').ToLowerInvariant()
        if (-not $value) { continue }
        if ($value -notmatch $domainPattern) { $errors.Add("Target.SenderDomains: '$d' is not a domain name (example: 'contoso.com')."); continue }
        if (-not $domains.Contains($value)) { $domains.Add($value) }
    }
    $limit = Test-Int (Get-Value $g 'Target' 'RecipientLimit' 25) 'Target.RecipientLimit' 1 1000
    $title = [string](Get-Value $r 'Report' 'Title' 'Messages with more than {0} recipients')
    $settings = [ordered]@{
        ConfigPath = [IO.Path]::GetFullPath($Path)
        Tenant = [ordered]@{
            TenantId     = [string](Get-Value $t 'Tenant' 'TenantId' '' -Required)
            Organization = [string](Get-Value $t 'Tenant' 'Organization' '')
        }
        Target = [ordered]@{
            RecipientLimit = $limit
            SenderDomains  = [string[]]$domains.ToArray()
        }
        Authentication = [ordered]@{
            Mode                  = [string](Get-Value $a 'Authentication' 'Mode' 'Certificate')
            AppId                 = [string](Get-Value $a 'Authentication' 'AppId' '')
            CertificateThumbprint = [string](Get-Value $a 'Authentication' 'CertificateThumbprint' '')
            ClientSecretVariable  = [string](Get-Value $a 'Authentication' 'ClientSecretVariable' 'RLR_CLIENT_SECRET')
            UserPrincipalName     = [string](Get-Value $a 'Authentication' 'UserPrincipalName' '')
        }
        Collection = [ordered]@{
            PageSize              = Test-Int (Get-Value $c 'Collection' 'PageSize' 5000) 'Collection.PageSize' 1 5000
            MaxConcurrency        = Test-Int (Get-Value $c 'Collection' 'MaxConcurrency' 3) 'Collection.MaxConcurrency' 1 8
            SliceHours            = Test-Int (Get-Value $c 'Collection' 'SliceHours' 2) 'Collection.SliceHours' 1 24
            SettlingHours         = Test-Int (Get-Value $c 'Collection' 'SettlingHours' 6) 'Collection.SettlingHours' 0 72
            BackfillDays          = Test-Int (Get-Value $c 'Collection' 'BackfillDays' 7) 'Collection.BackfillDays' 1 90
            SourceRetentionDays   = Test-Int (Get-Value $c 'Collection' 'SourceRetentionDays' 90) 'Collection.SourceRetentionDays' 1 90
            RetentionWarningDays  = Test-Int (Get-Value $c 'Collection' 'RetentionWarningDays' 7) 'Collection.RetentionWarningDays' 0 89
            MaxRequests           = Test-Int (Get-Value $c 'Collection' 'MaxRequests' 90) 'Collection.MaxRequests' 1 100
            PeriodSeconds         = Test-Int (Get-Value $c 'Collection' 'PeriodSeconds' 300) 'Collection.PeriodSeconds' 60 3600
            RequestTimeoutSeconds = Test-Int (Get-Value $c 'Collection' 'RequestTimeoutSeconds' 180) 'Collection.RequestTimeoutSeconds' 30 600
            MaxRetries            = Test-Int (Get-Value $c 'Collection' 'MaxRetries' 5) 'Collection.MaxRetries' 0 10
        }
        Counting = [ordered]@{
            MaxConcurrency    = Test-Int (Get-Value $n 'Counting' 'MaxConcurrency' 3) 'Counting.MaxConcurrency' 1 8
            MaxAttempts       = Test-Int (Get-Value $n 'Counting' 'MaxAttempts' 3) 'Counting.MaxAttempts' 1 10
            MaxMessagesPerRun = Test-Int (Get-Value $n 'Counting' 'MaxMessagesPerRun' 0) 'Counting.MaxMessagesPerRun' 0 10000000
            MaxRoutesPerMessage = Test-Int (Get-Value $n 'Counting' 'MaxRoutesPerMessage' 250) 'Counting.MaxRoutesPerMessage' 0 100000
            MaxRequests       = Test-Int (Get-Value $n 'Counting' 'MaxRequests' 90) 'Counting.MaxRequests' 1 100
            PeriodSeconds     = Test-Int (Get-Value $n 'Counting' 'PeriodSeconds' 300) 'Counting.PeriodSeconds' 60 3600
        }
        Storage = [ordered]@{
            DatabasePath  = Resolve-RlrPath ([string](Get-Value $s 'Storage' 'DatabasePath' '.\data\RecipientLimitReport.sqlite')) $Root
            RetentionDays = Test-Int (Get-Value $s 'Storage' 'RetentionDays' 180) 'Storage.RetentionDays' 0 3650
        }
        Report = [ordered]@{
            DefaultRange            = [string](Get-Value $r 'Report' 'DefaultRange' 'Last7Days')
            TimeZone                = [string](Get-Value $r 'Report' 'TimeZone' 'Europe/Paris')
            OutputPath              = Resolve-RlrPath ([string](Get-Value $r 'Report' 'OutputPath' '.\reports')) $Root
            FilePrefix              = [string](Get-Value $r 'Report' 'FilePrefix' 'RecipientLimit')
            Formats                 = @(Get-Value $r 'Report' 'Formats' @('Csv', 'Html'))
            IncludeRecipientDetails = Test-Bool (Get-Value $r 'Report' 'IncludeRecipientDetails' $true) 'Report.IncludeRecipientDetails'
            MaxRecipientsListed     = Test-Int (Get-Value $r 'Report' 'MaxRecipientsListed' 500) 'Report.MaxRecipientsListed' 0 100000
            SplitBy                 = [string](Get-Value $r 'Report' 'SplitBy' 'Week')
            MaxRowsPerFile          = Test-Int (Get-Value $r 'Report' 'MaxRowsPerFile' 500000) 'Report.MaxRowsPerFile' 1000 1048575
            CsvDelimiter            = [string](Get-Value $r 'Report' 'CsvDelimiter' ';')
            Title                   = $title.Replace('{0}', [string]$limit)
            TemplatePath            = Join-Path $Root 'templates\Report.template.html'
        }
        Logging = [ordered]@{
            Path          = Resolve-RlrPath ([string](Get-Value $l 'Logging' 'Path' '.\logs')) $Root
            RetentionDays = Test-Int (Get-Value $l 'Logging' 'RetentionDays' 30) 'Logging.RetentionDays' 1 3650
        }
    }
    $auth = $settings.Authentication
    if ($settings.Tenant.TenantId -and $settings.Tenant.TenantId -notmatch $guid) { $errors.Add('Tenant.TenantId must be a GUID.') }
    if ($auth.Mode -notin 'Certificate', 'ClientSecret', 'Interactive') { $errors.Add("Authentication.Mode must be 'Certificate', 'ClientSecret' or 'Interactive'.") }
    if ($auth.Mode -in 'Certificate', 'ClientSecret' -and $auth.AppId -notmatch $guid) { $errors.Add("Authentication.AppId must be the application (client) ID (GUID) in $($auth.Mode) mode.") }
    if ($auth.Mode -eq 'Interactive' -and $auth.AppId -and $auth.AppId -notmatch $guid) { $errors.Add("Authentication.AppId must be empty or a GUID in Interactive mode.") }
    if ($auth.Mode -eq 'Certificate' -and $auth.CertificateThumbprint -notmatch '^[0-9a-fA-F]{40}$') { $errors.Add('Authentication.CertificateThumbprint must be a 40-character thumbprint in Certificate mode.') }
    if ($auth.Mode -eq 'ClientSecret' -and $auth.ClientSecretVariable -notmatch '^[A-Za-z_][A-Za-z0-9_]*$') { $errors.Add('Authentication.ClientSecretVariable must be the name of an environment variable in ClientSecret mode.') }
    if ($settings.Collection.RetentionWarningDays -ge $settings.Collection.SourceRetentionDays) { $errors.Add('Collection.RetentionWarningDays must be lower than Collection.SourceRetentionDays.') }
    if ($settings.Report.DefaultRange -notin 'Last24Hours', 'Last7Days', 'Last30Days', 'PreviousMonth') { $errors.Add('Report.DefaultRange must be Last24Hours, Last7Days, Last30Days or PreviousMonth.') }
    if ($settings.Report.SplitBy -notin 'Rows', 'Day', 'Week') { $errors.Add("Report.SplitBy must be 'Rows', 'Day' or 'Week'.") }
    if (-not $settings.Report.Formats.Count -or @($settings.Report.Formats | Where-Object { $_ -notin 'Csv', 'Html' }).Count) { $errors.Add("Report.Formats must contain 'Csv', 'Html' or both.") }
    if ($settings.Report.CsvDelimiter -notin ';', ',', "`t", '|') { $errors.Add("Report.CsvDelimiter must be ';', ',', '|' or a tab.") }
    if (-not $settings.Report.FilePrefix -or $settings.Report.FilePrefix.IndexOfAny([IO.Path]::GetInvalidFileNameChars()) -ge 0) { $errors.Add('Report.FilePrefix must be a valid file name part.') }
    try { $settings['Zone'] = Get-RlrTimeZone $settings.Report.TimeZone } catch { $errors.Add($_.Exception.Message) }
    if ($errors.Count) { throw ("Invalid configuration ($Path):`n - " + ($errors -join "`n - ")) }
    return $settings
}

function Get-RlrScopeText {
    <# Short description of the senders reported: 'all senders' or the domains. #>
    param([Parameter(Mandatory)]$Settings)
    $d = @($Settings.Target.SenderDomains)
    if (-not $d.Count) { return 'all senders' }
    if ($d.Count -le 3) { return 'senders ' + (($d | ForEach-Object { "*@$_" }) -join ', ') }
    return 'senders ' + (($d | Select-Object -First 3 | ForEach-Object { "*@$_" }) -join ', ') + " (+$($d.Count - 3) domains)"
}

#endregion

#region 3. Engine (SQLite + compiled C#) ------------------------------------------------------

function Initialize-RlrEngine {
    <#
    .SYNOPSIS
        Loads SQLite (lib\sqlite) and the C# engine. The engine is compiled from
        src\RecipientLimitReport.Engine.cs into bin\ the first time, and again only when
        the source file changes (the file name contains a hash of the source).
    #>
    [CmdletBinding()]
    param([string]$Root = $script:ToolRoot)
    if ('RecipientLimitReport.RlrStore' -as [type]) { return }
    $lib = Join-Path $Root 'lib\sqlite'
    $arch = if ([Runtime.InteropServices.RuntimeInformation]::ProcessArchitecture -eq 'Arm64') { 'win-arm64' } else { 'win-x64' }
    $native = Join-Path $lib "runtimes\$arch\e_sqlite3.dll"
    if (-not (Test-Path -LiteralPath $native)) { throw "SQLite native library not found: $native" }
    [void][Runtime.InteropServices.NativeLibrary]::Load($native)
    foreach ($name in 'SQLitePCLRaw.core', 'SQLitePCLRaw.provider.e_sqlite3', 'SQLitePCLRaw.batteries_v2', 'Microsoft.Data.Sqlite') {
        Add-Type -LiteralPath (Join-Path $lib "$name.dll")
    }
    [SQLitePCL.Batteries_V2]::Init()

    $source = Join-Path $Root 'src\RecipientLimitReport.Engine.cs'
    $hash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash.Substring(0, 16)
    $bin = Join-Path $Root 'bin'
    $dll = Join-Path $bin "RecipientLimitReport.Engine.$hash.dll"
    if (-not (Test-Path -LiteralPath $dll)) {
        [void][IO.Directory]::CreateDirectory($bin)
        $references = @(
            (Join-Path $lib 'Microsoft.Data.Sqlite.dll'), 'System.IO.Compression', 'System.Text.Json', 'System.Text.Encodings.Web', 'System.Data.Common', 'System.Linq',
            'System.Collections', 'System.Collections.Concurrent', 'System.Text.RegularExpressions', 'System.Runtime', 'System.Memory', 'System.ComponentModel.Primitives',
            'System.ComponentModel', 'System.Transactions.Local', 'System.Text.Encoding.Extensions', 'System.Runtime.Extensions', 'System.IO', 'System.Runtime.InteropServices',
            'System.Console', 'System.Net.Http', 'System.Net.Primitives', 'System.Threading', 'System.Threading.Tasks', 'System.Private.Uri', 'netstandard')
        # Compiled to a unique name first: two consoles starting at the same time never write the same file.
        $staging = Join-Path $bin ("compile-{0}.dll" -f [guid]::NewGuid().ToString('N'))
        Add-Type -LiteralPath $source -ReferencedAssemblies $references -OutputAssembly $staging -OutputType Library -IgnoreWarnings -WarningAction SilentlyContinue
        try { Move-Item -LiteralPath $staging -Destination $dll -Force -ErrorAction Stop } catch { if (-not (Test-Path -LiteralPath $dll)) { throw } }
        # Older compiled versions are removed when they are not in use.
        Get-ChildItem -LiteralPath $bin -Filter 'RecipientLimitReport.Engine.*.dll' | Where-Object FullName -ne $dll | Remove-Item -Force -ErrorAction SilentlyContinue
        Get-ChildItem -LiteralPath $bin -Filter 'compile-*.dll' | Remove-Item -Force -ErrorAction SilentlyContinue
    }
    if (-not ('RecipientLimitReport.RlrStore' -as [type])) { Add-Type -LiteralPath $dll }
}

function Open-RlrStore {
    param([Parameter(Mandatory)]$Settings, [switch]$ReadOnly)
    $path = $Settings.Storage.DatabasePath
    if ($ReadOnly -and -not (Test-Path -LiteralPath $path)) { throw "The database does not exist yet: $path. Run a collection first (-Mode Collect)." }
    $store = [RecipientLimitReport.RlrStore]::new($path, [bool]$ReadOnly, $script:ToolVersion)
    $store.MaxRoutesPerMessage = $Settings.Counting.MaxRoutesPerMessage
    return $store
}

#endregion

#region 4. Periods and coverage ------------------------------------------------------------

function Get-RlrTimeZone {
    <# Accepts an IANA name (Europe/Paris) or a Windows name (Romance Standard Time). #>
    param([Parameter(Mandatory)][string]$Id)
    try { return [TimeZoneInfo]::FindSystemTimeZoneById($Id) } catch { }
    $windowsId = $null
    if ([TimeZoneInfo]::TryConvertIanaIdToWindowsId($Id, [ref]$windowsId)) { try { return [TimeZoneInfo]::FindSystemTimeZoneById($windowsId) } catch { } }
    $known = @{ 'Europe/Paris' = 'Romance Standard Time'; 'Europe/Brussels' = 'Romance Standard Time'; 'Europe/London' = 'GMT Standard Time'; 'UTC' = 'UTC' }
    if ($known.ContainsKey($Id)) { return [TimeZoneInfo]::FindSystemTimeZoneById($known[$Id]) }
    throw "Unknown time zone '$Id' (Report.TimeZone). Use a name such as 'Europe/Paris' or 'Romance Standard Time'."
}

function ConvertTo-RlrUnixMs {
    <# Text date -> Unix ms. Without an explicit offset (Z, +02:00) the date is read in the report time zone. #>
    param([Parameter(Mandatory)][string]$Text, [Parameter(Mandatory)][TimeZoneInfo]$Zone)
    $culture = [Globalization.CultureInfo]::InvariantCulture
    if ($Text -match '(Z|[+-]\d{2}:?\d{2})$') {
        return [DateTimeOffset]::Parse($Text, $culture).ToUnixTimeMilliseconds()
    }
    $formats = [string[]]@('yyyy-MM-dd', 'yyyy-MM-dd HH:mm', 'yyyy-MM-ddTHH:mm', 'yyyy-MM-dd HH:mm:ss', 'yyyy-MM-ddTHH:mm:ss')
    $local = [datetime]::MinValue
    if (-not [datetime]::TryParseExact($Text, $formats, $culture, [Globalization.DateTimeStyles]::None, [ref]$local)) {
        throw "Invalid date '$Text'. Use yyyy-MM-dd, 'yyyy-MM-dd HH:mm' or an ISO 8601 value with an offset."
    }
    return [RecipientLimitReport.Coverage]::LocalToUnixMs($local, $Zone)
}

function Resolve-RlrPeriod {
    <#
    .SYNOPSIS
        Converts a range name into a [start, end) period in Unix milliseconds.
    .DESCRIPTION
        Last24Hours / Last7Days / Last30Days : rolling windows ending now.
        PreviousMonth : the previous calendar month in the report time zone.
        Month (-Month yyyy-MM), Day (-Date yyyy-MM-dd), Custom (-Start / -End).
        The end is never later than now.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet('Last24Hours', 'Last7Days', 'Last30Days', 'PreviousMonth', 'Month', 'Day', 'Custom')][string]$Range,
        [string]$Month, [string]$Date, [string]$Start, [string]$End,
        [Parameter(Mandatory)][TimeZoneInfo]$Zone,
        [DateTimeOffset]$Now = [DateTimeOffset]::UtcNow
    )
    $culture = [Globalization.CultureInfo]::InvariantCulture
    $nowMs = $Now.ToUnixTimeMilliseconds(); $nowMs -= $nowMs % 1000
    $day = 86400000L
    switch ($Range) {
        'Last24Hours' { $s = $nowMs - $day; $e = $nowMs }
        'Last7Days' { $s = $nowMs - 7 * $day; $e = $nowMs }
        'Last30Days' { $s = $nowMs - 30 * $day; $e = $nowMs }
        'PreviousMonth' {
            $local = [TimeZoneInfo]::ConvertTime($Now, $Zone).DateTime
            $first = [datetime]::new($local.Year, $local.Month, 1).AddMonths(-1)
            $s = [RecipientLimitReport.Coverage]::LocalToUnixMs($first, $Zone)
            $e = [RecipientLimitReport.Coverage]::LocalToUnixMs($first.AddMonths(1), $Zone)
        }
        'Month' {
            $first = [datetime]::MinValue
            if (-not $Month -or -not [datetime]::TryParseExact($Month, 'yyyy-MM', $culture, 'None', [ref]$first)) { throw "-Range Month requires -Month in the format yyyy-MM (for example -Month 2026-08)." }
            $s = [RecipientLimitReport.Coverage]::LocalToUnixMs($first, $Zone)
            $e = [RecipientLimitReport.Coverage]::LocalToUnixMs($first.AddMonths(1), $Zone)
        }
        'Day' {
            $d = [datetime]::MinValue
            if (-not $Date -or -not [datetime]::TryParseExact($Date, 'yyyy-MM-dd', $culture, 'None', [ref]$d)) { throw "-Range Day requires -Date in the format yyyy-MM-dd." }
            $s = [RecipientLimitReport.Coverage]::LocalToUnixMs($d, $Zone)
            $e = [RecipientLimitReport.Coverage]::LocalToUnixMs($d.AddDays(1), $Zone)
        }
        'Custom' {
            if (-not $Start -or -not $End) { throw '-Range Custom requires -Start and -End.' }
            $s = ConvertTo-RlrUnixMs $Start $Zone
            $e = ConvertTo-RlrUnixMs $End $Zone
        }
    }
    if ($e -gt $nowMs) { $e = $nowMs }
    if ($e -le $s) { throw "The requested period is empty or in the future ($Range)." }
    [pscustomobject]@{
        Range   = $Range
        StartMs = [long]$s
        EndMs   = [long]$e
        Label   = [RecipientLimitReport.ReportPlanner]::PeriodLabel($s, $e, $Zone)
    }
}

function Get-RlrRangeLength {
    <# Total length in milliseconds of a list of ranges (each with Start and End). #>
    param([AllowNull()][AllowEmptyCollection()]$Ranges)
    $sum = 0L
    foreach ($r in @($Ranges)) { if ($null -ne $r) { $sum += [long]$r.End - [long]$r.Start } }
    return $sum
}

function New-RlrRangeList {
    param([Parameter()][AllowEmptyCollection()][object[]]$Ranges)
    $list = [Collections.Generic.List[RecipientLimitReport.TimeRange]]::new()
    foreach ($r in $Ranges) { if ($null -ne $r) { $list.Add([RecipientLimitReport.TimeRange]::new([long]$r.Start, [long]$r.End)) } }
    return , $list
}

function Get-RlrSignatures {
    <#
    .SYNOPSIS
        Message trace queries needed by the configuration: one query on every sender, or one
        query per sender domain (the API keeps a single value per property, so a list of domains
        cannot be sent in one query; '*@domain' is applied by the service).
    #>
    param([Parameter(Mandatory)]$Settings)
    $domains = @($Settings.Target.SenderDomains)
    if (-not $domains.Count) { return , @([pscustomobject]@{ Signature = ''; Condition = ''; Label = 'all senders'; Domain = '' }) }
    $list = foreach ($d in $domains) {
        [pscustomobject]@{ Signature = "sender=*@$d"; Condition = "senderAddress eq '*@$d'"; Label = "*@$d"; Domain = $d }
    }
    return , @($list)
}

function Get-RlrSignatureTexts {
    param([Parameter(Mandatory)]$Settings)
    return [string[]]@((Get-RlrSignatures $Settings) | ForEach-Object { $_.Signature })
}

function Get-RlrCollectionPlan {
    <#
    .SYNOPSIS
        Decides what must be collected for a period.
    .DESCRIPTION
        For each message trace query (every sender, or each sender domain):
        - ranges already collected and "settled" (collected at least SettlingHours after
          their end) are skipped;
        - missing ranges younger than the message trace retention (90 days) are collected,
          cut into slices of at most SliceHours that never cross local midnight;
        - missing ranges older than the retention are reported as unrecoverable.
    .OUTPUTS
        Items (work items for the engine, newest first) and the lengths to display.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Store, [Parameter(Mandatory)][long]$StartMs, [Parameter(Mandatory)][long]$EndMs,
        [Parameter(Mandatory)]$Settings, [long]$NowMs = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()
    )
    $c = $Settings.Collection
    # One hour of margin: the oldest hour may already be purged by the time it is requested.
    $earliest = $NowMs - [long]$c.SourceRetentionDays * 86400000 + 3600000
    $items = [Collections.Generic.List[RecipientLimitReport.WorkItem]]::new()
    $lost = [Collections.Generic.List[RecipientLimitReport.TimeRange]]::new()
    $collectableMs = 0L; $newMs = 0L
    $signatures = Get-RlrSignatures $Settings
    foreach ($sig in $signatures) {
        $settled = $Store.GetCoveredRanges($sig.Signature, [long]$c.SettlingHours * 3600000)
        $missing = [RecipientLimitReport.Coverage]::Gaps($settled, $StartMs, $EndMs)
        $collectable = [Collections.Generic.List[RecipientLimitReport.TimeRange]]::new()
        foreach ($gap in $missing) {
            if ($gap.End -le $earliest) { $lost.Add($gap); continue }
            if ($gap.Start -lt $earliest) {
                $lost.Add([RecipientLimitReport.TimeRange]::new($gap.Start, $earliest))
                $collectable.Add([RecipientLimitReport.TimeRange]::new($earliest, $gap.End))
            } else { $collectable.Add($gap) }
        }
        $collectableMs += Get-RlrRangeLength $collectable
        # Part of the collectable time never collected before (the rest is a refresh of recent data).
        $coveredAny = $Store.GetCoveredRanges($sig.Signature, 0)
        foreach ($range in $collectable) { $newMs += Get-RlrRangeLength ([RecipientLimitReport.Coverage]::Gaps($coveredAny, $range.Start, $range.End)) }
        foreach ($slice in [RecipientLimitReport.Coverage]::SplitIntoSlices($collectable, $Settings.Zone, $c.SliceHours)) {
            $item = [RecipientLimitReport.WorkItem]::new()
            $item.Signature = $sig.Signature; $item.Condition = $sig.Condition; $item.Label = $sig.Label
            $item.StartMs = $slice.Start; $item.EndMs = $slice.End
            $items.Add($item)
        }
    }
    $ordered = [Collections.Generic.List[RecipientLimitReport.WorkItem]]::new()
    $index = 0
    foreach ($item in ($items | Sort-Object -Property @{ Expression = 'EndMs'; Descending = $true }, @{ Expression = 'Label'; Descending = $false })) { $item.Index = ++$index; $ordered.Add($item) }
    $merged = [RecipientLimitReport.Coverage]::Merge($lost)
    [pscustomobject]@{
        Items            = $ordered
        Signatures       = $signatures
        CollectableMs    = $collectableMs
        NewMs            = $newMs
        RefreshMs        = $collectableMs - $newMs
        Unrecoverable    = $merged
        UnrecoverableMs  = Get-RlrRangeLength $merged
        EarliestSourceMs = $earliest
    }
}

function Get-RlrCoverage {
    <# Share of a period present in the database for every query of the configuration (settled or not). #>
    param([Parameter(Mandatory)]$Store, [Parameter(Mandatory)]$Settings, [Parameter(Mandatory)][long]$StartMs, [Parameter(Mandatory)][long]$EndMs)
    $covered = $Store.GetCoverage((Get-RlrSignatureTexts $Settings), 0)
    $gaps = [RecipientLimitReport.Coverage]::Gaps($covered, $StartMs, $EndMs)
    $missingMs = Get-RlrRangeLength $gaps
    [pscustomobject]@{
        Percent   = if ($EndMs -gt $StartMs) { [Math]::Round(100.0 * ($EndMs - $StartMs - $missingMs) / ($EndMs - $StartMs), 2) } else { 100 }
        Gaps      = $gaps
        MissingMs = $missingMs
    }
}

function Get-RlrExpiringGaps {
    <# Missing ranges that the message trace will no longer return within RetentionWarningDays. #>
    param([Parameter(Mandatory)]$Store, [Parameter(Mandatory)]$Settings, [long]$NowMs = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds())
    $c = $Settings.Collection
    $stats = $Store.GetStatistics($Settings.Target.RecipientLimit, $Settings.Counting.MaxAttempts)
    if ($null -eq $stats.FirstCoveredMs) { return @() }
    $windowStart = [Math]::Max([long]$stats.FirstCoveredMs, $NowMs - [long]$c.SourceRetentionDays * 86400000)
    $gaps = [RecipientLimitReport.Coverage]::Gaps($Store.GetCoverage((Get-RlrSignatureTexts $Settings), 0), $windowStart, $NowMs - [long]$c.SettlingHours * 3600000)
    $limit = $NowMs - [long]($c.SourceRetentionDays - $c.RetentionWarningDays) * 86400000
    foreach ($gap in $gaps) {
        if ($gap.Start -lt $limit) {
            [pscustomobject]@{ Start = $gap.Start; End = $gap.End; DaysLeft = [Math]::Max(0, [Math]::Floor(($gap.Start - ($NowMs - [long]$c.SourceRetentionDays * 86400000)) / 86400000.0)) }
        }
    }
}

#endregion

#region 5. Connection ------------------------------------------------------------------------
#
#  Three ways to obtain an access token for https://graph.microsoft.com with ExchangeMessageTrace.Read.All:
#
#    Certificate   app-only, recommended for scheduled tasks. The client assertion (a JWT signed with the
#                  private key of the certificate) is built here: no PowerShell module is needed.
#    ClientSecret  app-only with a secret, read from an environment variable (Authentication.ClientSecretVariable)
#                  or typed at the prompt (hidden). Never stored by the tool. Microsoft recommends a certificate.
#    Interactive   an administrator signs in (browser, MFA). Uses MSAL from the Microsoft.Graph.Authentication
#                  module. Delegated permission ExchangeMessageTrace.Read.All on the client application.
#
#  The token is checked before any request: right tenant, permission present. The requests run in the
#  engine; Update-RlrToken is called by the progress loops and renews the token 5 minutes before it
#  expires (or at once after a 401).

function Get-RlrTokenClaims {
    <# Payload of a JWT (no signature check: only used to read tid, roles, scp, app name). #>
    param([Parameter(Mandatory)][string]$Token)
    $parts = $Token.Split('.')
    if ($parts.Count -lt 2) { throw 'The access token is not a JWT.' }
    $p = $parts[1].Replace('-', '+').Replace('_', '/')
    while ($p.Length % 4) { $p += '=' }
    return [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($p)) | ConvertFrom-Json
}

function Get-RlrCertificate {
    <# Certificate with its private key, from Cert:\CurrentUser\My or Cert:\LocalMachine\My. #>
    param([Parameter(Mandatory)][string]$Thumbprint)
    $thumb = $Thumbprint.Trim().ToUpperInvariant()
    foreach ($store in 'Cert:\CurrentUser\My', 'Cert:\LocalMachine\My') {
        $cert = Get-Item -LiteralPath (Join-Path $store $thumb) -ErrorAction SilentlyContinue
        if ($cert) {
            if (-not $cert.HasPrivateKey) { throw "Certificate $thumb found in $store without its private key: import the .pfx (not the .cer) for the account that runs the tool." }
            if ($cert.NotAfter -lt (Get-Date)) { throw "Certificate $thumb expired on $($cert.NotAfter.ToString('yyyy-MM-dd')). Upload a new certificate to the application and update Authentication.CertificateThumbprint." }
            return $cert
        }
    }
    throw "Certificate $thumb not found in Cert:\CurrentUser\My nor Cert:\LocalMachine\My (account $([Environment]::UserName)). See the guide, chapter 4."
}

function ConvertTo-RlrBase64Url { param([byte[]]$Bytes) [Convert]::ToBase64String($Bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_') }

function New-RlrClientAssertion {
    <# Client assertion (RFC 7523) signed with the certificate: RS256, header x5t = SHA-1 thumbprint, valid 10 minutes. #>
    param([Parameter(Mandatory)][Security.Cryptography.X509Certificates.X509Certificate2]$Certificate, [Parameter(Mandatory)][string]$TenantId, [Parameter(Mandatory)][string]$AppId)
    $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    $header = [ordered]@{ alg = 'RS256'; typ = 'JWT'; x5t = (ConvertTo-RlrBase64Url $Certificate.GetCertHash()) } | ConvertTo-Json -Compress
    $claims = [ordered]@{ aud = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"; iss = $AppId; sub = $AppId; jti = [guid]::NewGuid().ToString(); nbf = $now - 60; iat = $now; exp = $now + 600 } | ConvertTo-Json -Compress
    $unsigned = (ConvertTo-RlrBase64Url ([Text.Encoding]::UTF8.GetBytes($header))) + '.' + (ConvertTo-RlrBase64Url ([Text.Encoding]::UTF8.GetBytes($claims)))
    $rsa = [Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate)
    if (-not $rsa) { throw "Certificate $($Certificate.Thumbprint): the private key is not an RSA key, or it cannot be used by this account." }
    try { $signature = $rsa.SignData([Text.Encoding]::ASCII.GetBytes($unsigned), [Security.Cryptography.HashAlgorithmName]::SHA256, [Security.Cryptography.RSASignaturePadding]::Pkcs1) }
    finally { $rsa.Dispose() }
    return "$unsigned.$(ConvertTo-RlrBase64Url $signature)"
}

function Get-RlrEntraErrorHint {
    <# A sentence for the most frequent Microsoft Entra sign-in errors. #>
    param([string]$Message)
    $hints = [ordered]@{
        'AADSTS700016'  = 'the application ID is not found in this tenant (Authentication.AppId, Tenant.TenantId).'
        'AADSTS90002'   = 'the tenant ID is not found (Tenant.TenantId).'
        'AADSTS700027'  = 'the certificate is not registered on the application, or not the right one (thumbprint).'
        'AADSTS7000215' = 'the client secret is not valid for this application.'
        'AADSTS7000222' = 'the client secret has expired: create a new one, or move to a certificate.'
        'AADSTS700024'  = 'the clock of this computer is not on time (the assertion is outside its validity).'
        'AADSTS65001'   = 'admin consent is missing for ExchangeMessageTrace.Read.All.'
        'AADSTS50105'   = 'the account is not assigned to the application.'
        'AADSTS53003'   = 'blocked by Conditional Access.'
    }
    foreach ($code in $hints.Keys) { if ($Message -match $code) { return "$code - $($hints[$code])" } }
    return $null
}

function Get-RlrAppToken {
    <# App-only token (client credentials): certificate or secret. Returns @{ Token; ExpiresMs }. #>
    param([Parameter(Mandatory)]$Settings, $Certificate, [Security.SecureString]$Secret)
    $tenant = $Settings.Tenant.TenantId; $appId = $Settings.Authentication.AppId
    $body = @{ client_id = $appId; scope = 'https://graph.microsoft.com/.default'; grant_type = 'client_credentials' }
    if ($Certificate) {
        $body['client_assertion_type'] = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
        $body['client_assertion'] = New-RlrClientAssertion -Certificate $Certificate -TenantId $tenant -AppId $appId
    } else {
        $body['client_secret'] = [Net.NetworkCredential]::new('', $Secret).Password
    }
    try { $r = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenant/oauth2/v2.0/token" -Body $body -ErrorAction Stop }
    catch {
        $detail = $_.ErrorDetails.Message
        $text = try { ($detail | ConvertFrom-Json).error_description } catch { $_.Exception.Message }
        if (-not $text) { $text = $_.Exception.Message }
        $hint = Get-RlrEntraErrorHint $text
        throw ("Microsoft Entra sign-in failed{0}: {1}" -f $(if ($hint) { " ($hint)" } else { '' }), ($text -split "`r?`n")[0])
    }
    finally { $body.Clear() }
    return @{ Token = $r.access_token; ExpiresMs = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds() + 1000L * [int]$r.expires_in }
}

function Import-RlrMsal {
    <# MSAL (Microsoft.Identity.Client) from the Microsoft.Graph.Authentication module: Interactive mode only. #>
    if ('Microsoft.Identity.Client.PublicClientApplicationBuilder' -as [type]) { return }
    $source = Get-Module -ListAvailable Microsoft.Graph.Authentication | Sort-Object Version -Descending |
        ForEach-Object { Join-Path $_.ModuleBase 'Dependencies' } |
        Where-Object { Test-Path -LiteralPath (Join-Path $_ 'Core\Microsoft.Identity.Client.dll') } | Select-Object -First 1
    if (-not $source) { throw 'Interactive mode needs the Microsoft.Graph.Authentication module (Install-Module Microsoft.Graph.Authentication -Scope CurrentUser), or use Certificate mode.' }
    Add-Type -LiteralPath (Join-Path $source 'Microsoft.IdentityModel.Abstractions.dll')
    Add-Type -LiteralPath (Join-Path $source 'Core\Microsoft.Identity.Client.dll')
}

function Connect-RlrGraph {
    <#
    .SYNOPSIS
        Obtains the first access token and checks it. Returns the connection used by Update-RlrToken.
    .OUTPUTS
        @{ Mode; Account; AppName; TenantId; Token; ExpiresMs; Renew (scriptblock) }
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Settings)
    $auth = $Settings.Authentication
    $connection = [ordered]@{ Mode = $auth.Mode; Account = ''; AppName = ''; TenantId = $Settings.Tenant.TenantId; Token = $null; ExpiresMs = 0; Renew = $null; Certificate = $null }
    switch ($auth.Mode) {
        'Certificate' {
            $cert = Get-RlrCertificate $auth.CertificateThumbprint
            if ($cert.NotAfter -lt (Get-Date).AddDays(30)) { Write-RlrItem Warn ("Certificate {0} expires on {1}: renew it." -f $cert.Thumbprint, $cert.NotAfter.ToString('yyyy-MM-dd')) }
            $connection.Certificate = $cert
            $connection.Renew = { param($s, $c) Get-RlrAppToken -Settings $s -Certificate $c.Certificate }
        }
        'ClientSecret' {
            $value = [Environment]::GetEnvironmentVariable($auth.ClientSecretVariable)
            if ($value) { $secret = ConvertTo-SecureString $value -AsPlainText -Force; $value = $null }
            elseif ([Environment]::UserInteractive -and -not [Console]::IsInputRedirected) { $secret = Read-Host -AsSecureString "      Client secret of application $($auth.AppId)" }
            else { throw "ClientSecret mode: the environment variable $($auth.ClientSecretVariable) is empty and the session is not interactive." }
            if (-not $secret -or $secret.Length -eq 0) { throw 'No client secret given.' }
            $connection['Secret'] = $secret
            $connection.Renew = { param($s, $c) Get-RlrAppToken -Settings $s -Secret $c.Secret }
        }
        'Interactive' {
            Import-RlrMsal
            $clientId = if ($auth.AppId) { $auth.AppId } else { '14d82eec-204b-4c2f-b7e8-296a70dab67e' }   # Microsoft Graph Command Line Tools
            $app = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($clientId).WithTenantId($Settings.Tenant.TenantId).WithRedirectUri('http://localhost').Build()
            $scopes = [string[]]@("https://graph.microsoft.com/$($script:Permission)")
            $cts = [Threading.CancellationTokenSource]::new([TimeSpan]::FromMinutes(5))
            $request = $app.AcquireTokenInteractive($scopes).WithUseEmbeddedWebView($false).WithPrompt([Microsoft.Identity.Client.Prompt]::SelectAccount)
            if ($auth.UserPrincipalName) { $request = $request.WithLoginHint($auth.UserPrincipalName) }
            try { $result = $request.ExecuteAsync($cts.Token).GetAwaiter().GetResult() }
            catch { $hint = Get-RlrEntraErrorHint $_.Exception.Message; throw ("Interactive sign-in failed{0}: {1}" -f $(if ($hint) { " ($hint)" } else { '' }), $_.Exception.Message) }
            $connection['Msal'] = $app; $connection['MsalAccount'] = $result.Account; $connection['Scopes'] = $scopes
            $connection.Token = $result.AccessToken; $connection.ExpiresMs = $result.ExpiresOn.ToUnixTimeMilliseconds()
            $connection.Renew = { param($s, $c)
                $r = $c.Msal.AcquireTokenSilent($c.Scopes, $c.MsalAccount).WithForceRefresh($true).ExecuteAsync().GetAwaiter().GetResult()
                @{ Token = $r.AccessToken; ExpiresMs = $r.ExpiresOn.ToUnixTimeMilliseconds() } }
        }
    }
    if (-not $connection.Token) {
        $t = & $connection.Renew $Settings $connection
        $connection.Token = $t.Token; $connection.ExpiresMs = $t.ExpiresMs
    }

    # Checks before the first request: right tenant, permission granted.
    $claims = Get-RlrTokenClaims $connection.Token
    if ($claims.tid -ne $Settings.Tenant.TenantId) { throw "Connected to tenant $($claims.tid), but Tenant.TenantId is $($Settings.Tenant.TenantId). Nothing was read." }
    if ($auth.Mode -eq 'Interactive') {
        $connection.Account = [string]$claims.upn
        if (-not $connection.Account) { $connection.Account = [string]$claims.unique_name }
        if ($auth.UserPrincipalName -and $connection.Account -ne $auth.UserPrincipalName) { throw "Signed in as $($connection.Account), but Authentication.UserPrincipalName is $($auth.UserPrincipalName)." }
        $scopes = @("$($claims.scp)" -split ' ')
        if ($scopes -notcontains $script:Permission) { throw "The token has no delegated permission $($script:Permission) (scopes: $($claims.scp)). An administrator must consent to it for the client application." }
    } else {
        $roles = @($claims.roles)
        $connection.AppName = [string]$claims.app_displayname
        $connection.Account = if ($connection.AppName) { "$($connection.AppName) ($($Settings.Authentication.AppId))" } else { "application $($Settings.Authentication.AppId)" }
        if ($roles -notcontains $script:Permission) {
            throw "The application has no application permission $($script:Permission) with admin consent (roles in the token: $(if ($roles.Count) { $roles -join ', ' } else { 'none' })). Entra admin center > App registrations > API permissions > Microsoft Graph > Application permissions, then 'Grant admin consent'."
        }
    }
    return $connection
}

function Update-RlrToken {
    <# Renews the token held by the engine when it expires within 5 minutes or after a 401. Returns $true when renewed. #>
    param([Parameter(Mandatory)]$Connection, [Parameter(Mandatory)][RecipientLimitReport.TokenSlot]$Slot, [Parameter(Mandatory)]$Settings, [switch]$Force)
    if (-not $Force -and -not $Slot.NeedsRefresh(300000)) { return $false }
    $t = & $Connection.Renew $Settings $Connection
    $Connection.Token = $t.Token; $Connection.ExpiresMs = $t.ExpiresMs
    $Slot.Set($t.Token, $t.ExpiresMs)
    Write-RlrLog 'INFO' ("Access token renewed, valid until {0:HH:mm:ss}." -f [DateTimeOffset]::FromUnixTimeMilliseconds($t.ExpiresMs).ToLocalTime())
    return $true
}

#endregion

#region 6. Collection ---------------------------------------------------------------------------

function New-RlrCollectorOptions {
    param([Parameter(Mandatory)]$Settings, [ValidateSet('Collection', 'Counting')][string]$Kind = 'Collection')
    $o = [RecipientLimitReport.CollectorOptions]::new()
    $o.GraphRoot = $script:GraphRoot
    $o.PageSize = $Settings.Collection.PageSize
    $o.MaxConcurrency = if ($Kind -eq 'Counting') { $Settings.Counting.MaxConcurrency } else { $Settings.Collection.MaxConcurrency }
    $o.MaxRetries = $Settings.Collection.MaxRetries
    $o.TimeoutSeconds = $Settings.Collection.RequestTimeoutSeconds
    $o.UserAgent = "RecipientLimitReport/$($script:ToolVersion)"
    return $o
}

function Write-RlrSliceRow {
    <# Prints the table row of one finished slice (once). #>
    param([Parameter(Mandatory)][RecipientLimitReport.WorkItem]$Item, [Parameter(Mandatory)][TimeZoneInfo]$Zone, [switch]$ShowScope)
    if ($Item.Reported -or $Item.State -notin 'Done', 'Failed') { return }
    $Item.Reported = $true
    $seconds = if ($Item.EndedMs -gt $Item.StartedMs -and $Item.StartedMs -gt 0) { ($Item.EndedMs - $Item.StartedMs) / 1000.0 } else { 0.0 }
    $rate = if ($seconds -gt 0) { '{0} rows/s' -f (Format-RlrNumber ([Math]::Round($Item.Rows / $seconds))) } else { '-' }
    $status = if ($Item.State -eq 'Done') { 'Ok' } else { 'Fail' }
    Write-RlrTableRow -Status $status -Window (Format-RlrRange $Item.StartMs $Item.EndMs $Zone) -Rows $Item.Rows -Messages $Item.NewMessages -Pages $Item.Pages -Duration (Format-RlrDuration $seconds) -Rate $rate -Scope $(if ($ShowScope) { $Item.Label })
}

function Invoke-RlrCollection {
    <#
    .SYNOPSIS
        Collects the message trace slices of a plan into the database.
    .DESCRIPTION
        MaxConcurrency workers send the requests within the shared quota of the tenant
        (Collection.MaxRequests per PeriodSeconds); one writer stores every page at once, so an
        interrupted execution never loses what was already received: the next run collects
        only the rest. Every 250 ms the loop renews the access token when needed, prints the
        events of the engine (429, retries) and the slices finished, and refreshes the progress.
        A slice that failed (network, 5xx after the retries) is tried once more for its missing
        part. A permanent error (permission, token) stops the collection.
    .OUTPUTS
        An object with the counters and the final state (Completed / Error).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Store, [Parameter(Mandatory)]$Settings, [Parameter(Mandatory)]$Connection,
        [Parameter(Mandatory)][Collections.Generic.List[RecipientLimitReport.WorkItem]]$Items, [Parameter(Mandatory)][long]$RunId,
        [RecipientLimitReport.CollectorOptions]$Options
    )
    $zone = $Settings.Zone; $c = $Settings.Collection
    $stats = [ordered]@{ Completed = $false; Error = $null; Fatal = $false; Items = $Items.Count; Done = 0; Failed = 0; Retried = 0; Pages = 0L; Rows = 0L; NewMessages = 0L; Requests = 0L; Throttled = 0L; Retries = 0L; Seconds = 0.0; WaitedSeconds = 0.0 }
    if (-not $Items.Count) { $stats.Completed = $true; return [pscustomobject]$stats }
    $slot = [RecipientLimitReport.TokenSlot]::new()
    $slot.Set($Connection.Token, $Connection.ExpiresMs)
    $limiter = [RecipientLimitReport.RateLimiter]::new($c.MaxRequests, [long]$c.PeriodSeconds * 1000, $Store.GetRequestStamps('graph_requests'))
    if ($limiter.InWindow -gt 0) { Write-RlrItem Info ("{0} request(s) of the last {1} min already counted (previous run): the quota is shared by the tenant." -f $limiter.InWindow, [int]($c.PeriodSeconds / 60)) -Icon Clock }
    if (-not $Options) { $Options = New-RlrCollectorOptions $Settings }
    $showScope = @($Settings.Target.SenderDomains).Count -gt 1
    $clock = [Diagnostics.Stopwatch]::StartNew()
    $seen = @{}
    $queue = $Items
    $all = [Collections.Generic.List[RecipientLimitReport.WorkItem]]::new($Items)
    Write-RlrTableRow -Header
    try {
        for ($pass = 1; $pass -le 2; $pass++) {
            $collector = [RecipientLimitReport.TraceCollector]::new($Store, $Options, $slot, $limiter, $RunId)
            $cts = [Threading.CancellationTokenSource]::new()
            $task = $collector.RunAsync($queue, $cts.Token)
            $lastLog = 0; $lastDraw = 0; $tokenWarned = $false
            try {
                while (-not $task.IsCompleted) {
                    try { [void](Update-RlrToken -Connection $Connection -Slot $slot -Settings $Settings) }
                    catch { if (-not $tokenWarned) { Write-RlrItem Warn "Access token not renewed: $($_.Exception.Message)"; $tokenWarned = $true } }
                    Write-RlrEngineEvents $collector.Events $seen
                    foreach ($i in $queue) { Write-RlrSliceRow $i $zone -ShowScope:$showScope }
                    $elapsed = $clock.Elapsed.TotalSeconds
                    if ($elapsed - $lastDraw -ge 0.5) {
                        $lastDraw = $elapsed
                        $progress = $collector.Progress
                        $parts = [Collections.Generic.List[string]]::new()
                        $parts.Add(('slice {0}/{1}' -f [Math]::Min($queue.Count, $collector.ItemsDone + $collector.ItemsFailed + 1), $queue.Count))
                        $parts.Add(('{0} rows' -f (Format-RlrNumber ($stats.Rows + $collector.Rows))))
                        $parts.Add(('{0} requests' -f (Format-RlrNumber ($stats.Requests + $collector.Requests))))
                        $parts.Add(('quota {0}/{1}' -f $limiter.InWindow, $limiter.MaxRequests))
                        $delay = $limiter.DelayMs()
                        if ($limiter.PausedUntil -gt [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()) { $parts.Add(('paused {0} s (429)' -f [int][Math]::Ceiling($delay / 1000))) }
                        elseif ($delay -gt 1500 -and $collector.ItemsRunning -gt 0) { $parts.Add(('waiting for the quota {0} s' -f [int][Math]::Ceiling($delay / 1000))) }
                        $parts.Add($(if ($progress -gt 0.02 -and $elapsed -gt 5) { 'remaining ~' + (Format-RlrDuration ($elapsed / $progress * (1 - $progress))) } else { 'estimating...' }))
                        Write-Progress -Id 1 -Activity 'Collecting the message trace' -PercentComplete ([Math]::Min(100, [int](100 * $progress))) -Status ($parts -join '  |  ')
                    }
                    if ($elapsed - $lastLog -ge 60) {
                        $lastLog = $elapsed
                        Write-RlrLog 'INFO' ("Progress {0:P0}: {1}/{2} slices, {3} requests, {4} rows, {5} new messages, {6} throttled" -f $collector.Progress, $collector.ItemsDone, $queue.Count, $collector.Requests, $collector.Rows, $collector.NewMessages, $collector.Throttled)
                    }
                    [void]$task.Wait(250)
                }
            }
            finally {
                if (-not $task.IsCompleted) {
                    # Ctrl+C: stop the requests, let the writer store what was received.
                    Write-Progress -Id 1 -Activity 'Collecting the message trace' -Completed
                    Write-Host '      Stopping: the pages already received are being saved...'
                    $cts.Cancel()
                    try { [void]$task.Wait(120000) } catch { }
                    Write-RlrLog 'WARN' 'Collection interrupted by the user (Ctrl+C).'
                }
                Write-Progress -Id 1 -Activity 'Collecting the message trace' -Completed
                Write-RlrEngineEvents $collector.Events $seen
                foreach ($i in $queue) { Write-RlrSliceRow $i $zone -ShowScope:$showScope }
                try { $Store.SetRequestStamps('graph_requests', $limiter.Stamps()) } catch { }
            }
            if ($task.IsFaulted) { throw $task.Exception.GetBaseException() }
            $stats.Pages += $collector.Pages; $stats.Rows += $collector.Rows; $stats.NewMessages += $collector.NewMessages
            $stats.Requests += $collector.Requests; $stats.Throttled += $collector.Throttled; $stats.Retries += $collector.Retries
            if ($collector.Fatal) { $stats.Error = $collector.Fatal; $stats.Fatal = $true; break }
            $failed = @($queue | Where-Object { $_.State -eq 'Failed' -and -not $_.Unrecoverable })
            if (-not $failed.Count -or $pass -eq 2) { break }
            # Second pass: only the part of each failed slice that was not stored yet.
            $retry = [Collections.Generic.List[RecipientLimitReport.WorkItem]]::new()
            foreach ($f in $failed) {
                $end = if ($f.IntervalId -gt 0) { $f.CoveredFromMs } else { $f.EndMs }
                if ($end -le $f.StartMs) { continue }
                $n = [RecipientLimitReport.WorkItem]::new()
                $n.Signature = $f.Signature; $n.Condition = $f.Condition; $n.Label = $f.Label; $n.StartMs = $f.StartMs; $n.EndMs = $end; $n.Index = $f.Index
                $retry.Add($n); $all.Add($n)
            }
            if (-not $retry.Count) { break }
            $stats.Retried = $retry.Count
            Write-RlrItem Warn ("{0} slice(s) failed: their missing part is requested once more." -f $retry.Count)
            $queue = $retry
        }
    }
    finally {
        $stats.Seconds = $clock.Elapsed.TotalSeconds
        $stats.WaitedSeconds = $limiter.WaitedMs / 1000.0
    }
    # A slice counts as failed only when its last attempt failed.
    $final = @{}
    foreach ($i in $all) { $final["$($i.Signature)|$($i.StartMs)"] = $i }
    $stats.Done = @($final.Values | Where-Object State -eq 'Done').Count
    $stats.Failed = @($final.Values | Where-Object State -ne 'Done').Count
    if (-not $stats.Error -and $stats.Failed) { $stats.Error = "$($stats.Failed) slice(s) not collected completely: " + ((@($final.Values | Where-Object { $_.State -ne 'Done' -and $_.Error }) | Select-Object -First 1).Error) }
    $stats.Completed = -not $stats.Error
    return [pscustomobject]$stats
}

#endregion

#region 7. Recipient count before distribution list expansion ------------------------------------
#
#  The message trace lists the recipients AFTER expansion: a distribution list appears once with the
#  status 'expanded', and each of its members appears as a recipient. Exchange Online (RecipientLimits)
#  and the DLP condition RecipientCountOver count the recipients BEFORE expansion (a list = one
#  recipient). For a message without any list both counts are the same; for a message sent to a list,
#  the count before expansion is read from the route of the list (getDetailsByRecipient): the Submit
#  event carries RcptCount, the number of recipients the sender addressed. One request per message,
#  only for messages over the limit after expansion; the answer is kept in the database.
#  For a message over the limit BEFORE expansion, the recipient list as the sender addressed it is then
#  rebuilt: the route of a recipient the sender addressed starts with the Submit event, the route of a
#  member added by the expansion starts after it. Routes are read until as many recipients as the count
#  are found, at most Counting.MaxRoutesPerMessage per message.

function Get-RlrCountingEstimate {
    <# Time needed to send N getDetailsByRecipient requests within its quota. #>
    param([Parameter(Mandatory)][long]$Requests, [Parameter(Mandatory)]$Settings)
    $n = $Settings.Counting
    return [double]$Requests * 1.05 * $n.PeriodSeconds / $n.MaxRequests
}

function Get-RlrCountingRoutes {
    <#
    .SYNOPSIS
        Fewest and most routes the counting step can read for the items of Store.GetCountItems.
    .DESCRIPTION
        Fewest: one route per count still to read, and for a known count over the limit one route per
        recipient the sender addressed and not found yet. Most: every route listed for the item (its
        Targets, already bounded by Counting.MaxRoutesPerMessage). A count still to read gives no clue
        on its own: when it is over the limit, the list is rebuilt in the same run, about one route per
        recipient (lab: 33 counts to read, 604 requests, 10 lists rebuilt).
    #>
    param([Parameter(Mandatory)][AllowEmptyCollection()][Collections.Generic.List[RecipientLimitReport.CountItem]]$Items)
    [long]$min = 0; [long]$max = 0
    foreach ($item in $Items) {
        $targets = $item.Targets.Count
        $max += $targets
        if ($null -eq $item.Count) { $min += [Math]::Min(1, $targets) }
        elseif ($item.RebuildList -and $item.Count -gt $item.Limit) { $min += [Math]::Min($targets, [Math]::Max(0, $item.Count - $item.DirectKnown)) }
    }
    [pscustomobject]@{ Min = $min; Max = $max }
}

function Invoke-RlrCounting {
    <#
    .SYNOPSIS
        Reads the routes of the messages returned by Store.GetCountItems: recipient count before
        expansion, then the recipient list before expansion of the messages over the limit.
    .OUTPUTS
        An object with the counters: Done (count read), NoCount (route without a count), Failed,
        ListsDone (recipient list rebuilt), ListsPartial (MaxRoutesPerMessage reached).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Store, [Parameter(Mandatory)]$Settings, [Parameter(Mandatory)]$Connection,
        [Parameter(Mandatory)][Collections.Generic.List[RecipientLimitReport.CountItem]]$Items,
        [RecipientLimitReport.CollectorOptions]$Options
    )
    $n = $Settings.Counting
    $slot = [RecipientLimitReport.TokenSlot]::new()
    $slot.Set($Connection.Token, $Connection.ExpiresMs)
    $limiter = [RecipientLimitReport.RateLimiter]::new($n.MaxRequests, [long]$n.PeriodSeconds * 1000, $Store.GetRequestStamps('graph_detail_requests'))
    if (-not $Options) { $Options = New-RlrCollectorOptions $Settings -Kind Counting }
    $collector = [RecipientLimitReport.CountCollector]::new($Store, $Options, $slot, $limiter)
    $cts = [Threading.CancellationTokenSource]::new()
    $clock = [Diagnostics.Stopwatch]::StartNew()
    $seen = @{}
    $task = $collector.RunAsync($Items, $cts.Token)
    try {
        $lastDraw = 0
        while (-not $task.IsCompleted) {
            try { [void](Update-RlrToken -Connection $Connection -Slot $slot -Settings $Settings) } catch { }
            Write-RlrEngineEvents $collector.Events $seen
            $elapsed = $clock.Elapsed.TotalSeconds
            if ($elapsed - $lastDraw -ge 0.5) {
                $lastDraw = $elapsed
                $progress = $collector.Progress
                $left = if ($progress -gt 0.02 -and $elapsed -gt 5) { 'remaining ~' + (Format-RlrDuration ($elapsed / $progress * (1 - $progress))) } else { 'estimating...' }
                Write-Progress -Id 2 -Activity 'Counting recipients before distribution list expansion' -PercentComplete ([Math]::Min(100, [int](100 * $progress))) `
                    -Status ('message {0}/{1}  |  {2} requests  |  quota {3}/{4}  |  {5}' -f $collector.Finished, $Items.Count, (Format-RlrNumber $collector.Requests), $limiter.InWindow, $limiter.MaxRequests, $left)
            }
            [void]$task.Wait(250)
        }
    }
    finally {
        if (-not $task.IsCompleted) {
            Write-Host '      Stopping: the counts and routes already read are being saved...'
            $cts.Cancel()
            try { [void]$task.Wait(60000) } catch { }
            Write-RlrLog 'WARN' 'Counting interrupted by the user (Ctrl+C).'
        }
        Write-Progress -Id 2 -Activity 'Counting recipients before distribution list expansion' -Completed
        Write-RlrEngineEvents $collector.Events $seen
        try { $Store.SetRequestStamps('graph_detail_requests', $limiter.Stamps()) } catch { }
    }
    if ($task.IsFaulted) { throw $task.Exception.GetBaseException() }
    [pscustomobject]@{
        Items        = $Items.Count
        Done         = [long]$collector.Done
        NoCount      = [long]$collector.NoCount
        Failed       = [long]$collector.Failed
        ListsDone    = [long]$collector.ListsDone
        ListsPartial = [long]$collector.ListsPartial
        Requests     = $collector.Requests
        Throttled    = $collector.Throttled
        Seconds      = $clock.Elapsed.TotalSeconds
        Fatal        = $collector.Fatal
        Completed    = -not $collector.Fatal -and ($collector.Finished -eq $Items.Count)
    }
}

#endregion

#region 8. Report --------------------------------------------------------------------------------

function Move-RlrDirectory {
    <# Directory move with retries: antivirus or indexing can briefly lock files that were just written. #>
    param([Parameter(Mandatory)][string]$Source, [Parameter(Mandatory)][string]$Destination, [int]$Attempts = 10)
    for ($i = 1; $i -le $Attempts; $i++) {
        try { [IO.Directory]::Move($Source, $Destination); return }
        catch {
            $e = $_.Exception; while ($e.InnerException) { $e = $e.InnerException }
            $code = $e.HResult -band 0xFFFF
            $transient = $e -is [UnauthorizedAccessException] -or $code -in 5, 32, 33
            if (-not $transient -or $i -eq $Attempts) { throw }
            Start-Sleep -Milliseconds (250 * $i)
        }
    }
}

function New-RlrReport {
    <#
    .SYNOPSIS
        Writes the CSV and HTML files of a period from the database.
    .DESCRIPTION
        Files are written in a "<folder>.pending" directory, which is renamed only when
        every file is complete: a folder without ".pending" is always a finished report.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Store, [Parameter(Mandatory)]$Settings, [Parameter(Mandatory)]$Period,
        [Parameter(Mandatory)]$Coverage, [long]$Unmeasured = 0
    )
    $r = $Settings.Report
    $zone = $Settings.Zone
    [void][IO.Directory]::CreateDirectory($r.OutputPath)
    $name = '{0:yyyy-MM-dd_HHmmss}_{1}' -f (Get-Date), $Period.Range
    $final = Join-Path $r.OutputPath $name
    for ($n = 2; (Test-Path -LiteralPath $final) -or (Test-Path -LiteralPath "$final.pending"); $n++) { $final = Join-Path $r.OutputPath "${name}_$n" }
    $staging = "$final.pending"
    $request = [RecipientLimitReport.ReportRequest]::new()
    $request.StartMs = $Period.StartMs
    $request.EndMs = $Period.EndMs
    $request.Zone = $zone
    $request.TimeZoneLabel = $r.TimeZone
    $request.Limit = $Settings.Target.RecipientLimit
    $request.SenderDomains = [string[]]@($Settings.Target.SenderDomains)
    $request.SplitBy = $r.SplitBy
    $request.MaxRowsPerFile = $r.MaxRowsPerFile
    $request.IncludeRecipientDetails = [bool]$r.IncludeRecipientDetails
    $request.MaxRecipientsListed = $r.MaxRecipientsListed
    $request.WriteCsv = 'Csv' -in $r.Formats
    $request.WriteHtml = 'Html' -in $r.Formats
    $request.CsvDelimiter = $r.CsvDelimiter
    $request.OutputDirectory = $staging
    $request.FilePrefix = $r.FilePrefix
    $request.HtmlTemplatePath = $r.TemplatePath
    $request.Title = $r.Title
    $request.ScopeLabel = Get-RlrScopeText $Settings
    $request.ToolVersion = $script:ToolVersion
    $request.RangeName = $Period.Range
    $request.CoveragePercent = $Coverage.Percent
    $request.Unmeasured = $Unmeasured
    if ($Coverage.Percent -lt 100) {
        $request.CoverageNote = 'no data collected for ' + ((@($Coverage.Gaps) | Select-Object -First 3 | ForEach-Object { Format-RlrRange $_.Start $_.End $zone }) -join ', ') + $(if (@($Coverage.Gaps).Count -gt 3) { ', ...' } else { '' })
    }
    $result = $Store.WriteReport($request)
    Move-RlrDirectory -Source $staging -Destination $final
    foreach ($file in $result.Files) { $file.Path = Join-Path $final (Split-Path $file.Path -Leaf) }
    [pscustomobject]@{ Directory = $final; Result = $result }
}

#endregion

#region 9. Status and maintenance ------------------------------------------------------------------

function Show-RlrStatus {
    <#
    .SYNOPSIS
        Prints the database content and, day by day, what has been collected: date, messages
        over the limit, messages still to count, a coverage bar and the state.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Store, [Parameter(Mandatory)]$Settings, [int]$Days = 35)
    $zone = $Settings.Zone; $cfg = $Settings.Collection; $K = $script:C; $dot = [char]0x00B7
    $limit = $Settings.Target.RecipientLimit
    $now = [DateTimeOffset]::UtcNow
    $nowMs = $now.ToUnixTimeMilliseconds()
    $s = $Store.GetStatistics($limit, $Settings.Counting.MaxAttempts)
    Write-RlrItem Info ('{0}  ({1})' -f $Store.DatabasePath, (Format-RlrBytes $s.FileBytes)) -Icon Database
    Write-RlrItem Info ('{0} messages over {1} recipients {5} {2} still to count {5} {6} recipient list(s) to rebuild {5} {3} messages stored {5} {4} executions' -f (Format-RlrNumber $s.Matches), $limit, (Format-RlrNumber $s.Pending), (Format-RlrNumber $s.Messages), (Format-RlrNumber $s.Runs), $dot, (Format-RlrNumber $s.ListsPending)) -Icon Chart
    if ($null -ne $s.CompactionLimit -and $s.CompactionLimit -gt $limit) { Write-RlrItem Warn ("Target.RecipientLimit is {0}, but the database was compacted with {1}: messages with {2} to {1} recipients collected before are not in the database." -f $limit, $s.CompactionLimit, ($limit + 1)) }
    if ($null -eq $s.FirstCoveredMs) { Write-RlrItem Warn 'Nothing has been collected yet. Run: .\Invoke-RecipientLimitReport.ps1 -Mode Collect'; return }
    Write-RlrItem Info ('History {0}   {1} last collection {2}' -f (Format-RlrRange $s.FirstCoveredMs $s.LastCoveredMs $zone), $dot, $(if ($s.LastSuccessfulCollectionMs) { Format-RlrLocalTime $s.LastSuccessfulCollectionMs $zone 'yyyy-MM-dd HH:mm:ss' } else { '-' })) -Icon Calendar
    Write-Host ''
    $today = [TimeZoneInfo]::ConvertTime($now, $zone).DateTime.Date
    $first = $today.AddDays( - ($Days - 1))
    $firstCollected = [TimeZoneInfo]::ConvertTime([DateTimeOffset]::FromUnixTimeMilliseconds($s.FirstCoveredMs), $zone).DateTime.Date
    if ($firstCollected -gt $first) { $first = $firstCollected }
    $Days = [int]($today - $first).TotalDays + 1
    $rows = $Store.GetDailyStatistics((Get-RlrSignatureTexts $Settings), [string[]]@($Settings.Target.SenderDomains), $first, $Days, $zone, [long]$cfg.SettlingHours * 3600000, $limit, $Settings.Counting.MaxAttempts)
    $sourceStart = $nowMs - [long]$cfg.SourceRetentionDays * 86400000
    $en = [Globalization.CultureInfo]::GetCultureInfo('en-US')
    Write-Host ('      {0}{1}{2,-15} {3,10} {4,10}   {5,-16} {6}{7}' -f $K.Dim, ('  ' + $script:IconPad), 'Day', 'Messages', 'To read', 'Collected', 'State', $K.Reset)
    foreach ($d in $rows) {
        if ($d.StartMs -ge $nowMs) { continue }
        $length = [Math]::Max(1, [Math]::Min($d.EndMs, $nowMs) - $d.StartMs)
        $percent = [int][Math]::Min(100, [Math]::Floor(100.0 * $d.CoveredMs / $length))
        # State, status icon and colour of the day.
        if ($d.LocalDate -eq $today) { $state = if ($percent -ge 100) { 'Today, provisional' } else { 'Today, collection in progress' }; $status = 'Info' }
        elseif ($percent -ge 100 -and $d.FinalCoveredMs -ge ($d.EndMs - $d.StartMs)) { $state = 'Complete'; $status = 'Ok' }
        elseif ($percent -ge 100) { $state = 'Complete, refreshed at next run'; $status = 'Ok' }
        elseif ($d.EndMs -le $sourceStart) { $state = if ($percent -gt 0) { 'Partial, rest no longer collectable' } else { 'Missing, no longer collectable' }; $status = 'Fail' }
        elseif ($d.StartMs -lt $sourceStart) { $state = "Partial, earlier hours beyond the $($cfg.SourceRetentionDays)-day retention"; $status = 'Info' }
        else {
            $left = [Math]::Max(0, [Math]::Floor(($d.StartMs - $sourceStart) / 86400000.0))
            $state = '{0}, collectable {1} more day(s)' -f $(if ($percent -gt 0) { 'Partial' } else { 'Missing' }), $left; $status = 'Warn'
        }
        if ($d.Pending -gt 0 -and $status -eq 'Ok') { $state += ', routes to read'; $status = 'Info' }
        $color = @{ Ok = $K.Green; Warn = $K.Yellow; Fail = $K.Red; Info = $K.Cyan }[$status]
        # Coverage bar: 12 cells.
        $filled = [int][Math]::Round(12 * $percent / 100.0)
        $bar = $color + [string]::new([char]0x2588, $filled) + $K.Dim + [string]::new([char]0x2591, 12 - $filled) + $K.Reset
        $day = '{0} {1}' -f $d.LocalDate.ToString('ddd', $en), $d.LocalDate.ToString('yyyy-MM-dd')
        $messages = if ($d.Matches) { $K.Bold + (Format-RlrNumber $d.Matches).PadLeft(10) + $K.Reset } else { $K.Dim + '0'.PadLeft(10) + $K.Reset }
        $pending = if ($d.Pending) { $K.Yellow + (Format-RlrNumber $d.Pending).PadLeft(10) + $K.Reset } else { $K.Dim + '-'.PadLeft(10) + $K.Reset }
        Write-Host ('      {0}{1}{2}{3,-15} {4} {5}   {6} {7,4}%  {0}{8}{2}' -f $color, (Get-RlrIcon $status), $K.Reset, $day, $messages, $pending, $bar, $percent, $state)
        Write-RlrLog 'INFO' ("Status {0}: {1} messages over the limit, {2} with routes to read, {3}% collected, {4}" -f $d.LocalDate.ToString('yyyy-MM-dd'), $d.Matches, $d.Pending, $percent, $state)
    }
}

function Invoke-RlrCompaction {
    <#
    .SYNOPSIS
        Removes from the settled part of the history the messages that no report can show
        (Target.RecipientLimit recipients or fewer). Returns the number of messages removed.
    #>
    param([Parameter(Mandatory)]$Store, [Parameter(Mandatory)]$Settings)
    $settled = $Store.GetCoverage((Get-RlrSignatureTexts $Settings), [long]$Settings.Collection.SettlingHours * 3600000)
    return $Store.Compact($settled, $Settings.Target.RecipientLimit)
}

function Invoke-RlrRetention {
    <# Deletes database content older than Storage.RetentionDays (0 = keep everything). #>
    param([Parameter(Mandatory)]$Store, [Parameter(Mandatory)]$Settings)
    if ($Settings.Storage.RetentionDays -le 0) { return $null }
    $cutoff = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds() - [long]$Settings.Storage.RetentionDays * 86400000
    return $Store.PurgeBefore($cutoff)
}

function Enter-RlrLock {
    <#
    .SYNOPSIS
        Prevents two executions from collecting at the same time (for example the scheduled
        task and an administrator). Returns the lock, to be released with Exit-RlrLock.
    #>
    param([Parameter(Mandatory)][string]$Path, [int]$TimeoutSeconds = 30)
    [void][IO.Directory]::CreateDirectory((Split-Path $Path -Parent))
    $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
    while ($true) {
        try {
            $stream = [IO.FileStream]::new($Path, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::Read)
            $stream.SetLength(0)
            $bytes = [Text.Encoding]::UTF8.GetBytes(("{0} pid {1} since {2:o}" -f [Environment]::MachineName, $PID, (Get-Date)))
            $stream.Write($bytes, 0, $bytes.Length); $stream.Flush()
            return $stream
        } catch [IO.IOException] {
            if ([DateTime]::UtcNow -ge $deadline) {
                $owner = try { [IO.File]::ReadAllText($Path) } catch { 'unknown' }
                throw "Another execution is already collecting ($owner). Wait for it to finish, or use -NoCollect to build a report from the data already collected."
            }
            Start-Sleep -Seconds 2
        }
    }
}

function Exit-RlrLock {
    param($Lock)
    if ($Lock) { $Lock.Dispose() }
}

#endregion
