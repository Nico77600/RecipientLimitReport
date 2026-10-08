#Requires -Version 7.4

<#
.SYNOPSIS
    Renders the graphics of the GitHub README from the Recipient Limit Report guide, in a light and a dark
    version: banner, why, how it works, same report as Purview DLP Report.

.DESCRIPTION
    GitHub renders Markdown only: the custom blocks of the guide (cards, flow) and its theme are lost. This
    tool renders them as images with the CSS of the built HTML guide and the icons of
    tools\Build-Documentation.ps1, so that the README and the guide always look the same. The README shows
    them with <picture>, which picks the light or dark image from the theme of the reader.

    Sources:
      package\docs\RecipientLimitReport-Guide.md     the cards and flow blocks (chapters 1 and 2), the version
      package\docs\RecipientLimitReport-Guide.html   the CSS (run tools\Build-Documentation.ps1 first)
      tools\Build-Documentation.ps1          the icons

    Screenshots: Microsoft Edge in headless mode, with a temporary profile, 2x resolution. Only local files
    are opened. Output: package\docs\images\readme-<name>-light.png and readme-<name>-dark.png.

.PARAMETER OutputFolder
    Default: package\docs\images next to the tools folder.

.PARAMETER KeepWork
    Keeps the work folder (the HTML pages of the graphics) and shows its path.

.EXAMPLE
    .\tools\Build-Documentation.ps1; .\tools\New-ReadmeImages.ps1

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.1  (from Message Trace Report)
    Part of : Recipient Limit Report (repository tool, not in the package)
#>
[CmdletBinding()]
param(
    [string]$OutputFolder,
    [switch]$KeepWork
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
if (-not $OutputFolder) { $OutputFolder = Join-Path $root 'package\docs\images' }

#region Assets of the guide ------------------------------------------------------------------------
function ConvertTo-ReadmeInline([string]$Text) {
    # Inline Markdown of a guide block (code, bold, italic) -> HTML.
    $h = [System.Net.WebUtility]::HtmlEncode($Text.Trim())
    $h = [regex]::Replace($h, '`([^`]+)`', '<code>$1</code>')
    $h = [regex]::Replace($h, '\*\*([^*]+)\*\*', '<strong>$1</strong>')
    return [regex]::Replace($h, '(?<![\w*])\*([^*\s][^*]*)\*(?![\w*])', '<em>$1</em>')
}

function Get-ReadmeAssets {
    param([string]$Root)
    $builder = Join-Path $Root 'tools\Build-Documentation.ps1'
    $guideHtml = Join-Path $Root 'package\docs\RecipientLimitReport-Guide.html'
    $guideMd = Join-Path $Root 'package\docs\RecipientLimitReport-Guide.md'
    if (-not (Test-Path $guideHtml)) { throw 'package\docs\RecipientLimitReport-Guide.html not found: run tools\Build-Documentation.ps1 first (it holds the CSS of the graphics).' }
    # Icons: the $Icons table of the documentation builder, read without running the builder.
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($builder, [ref]$null, [ref]$null)
    $assign = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$Icons' }, $true)
    if (-not $assign) { throw "Icon table not found in $builder." }
    $md = [IO.File]::ReadAllText($guideMd) -replace "`r`n", "`n"
    $blocks = foreach ($m in [regex]::Matches($md, '(?s)```(flow|cards)\n(.*?)\n```')) {
        [pscustomobject]@{ Kind = $m.Groups[1].Value; Lines = @($m.Groups[2].Value -split "`n" | Where-Object { $_.Trim() }) }
    }
    [pscustomobject]@{
        Icons   = & ([scriptblock]::Create($assign.Right.Extent.Text))
        Css     = [regex]::Match([IO.File]::ReadAllText($guideHtml), '(?s)<style>(.*?)</style>').Groups[1].Value
        Version = [regex]::Match($md, '(?m)^version:\s*(\S+)').Groups[1].Value
        Flows   = @($blocks | Where-Object Kind -eq 'flow')
        Cards   = @($blocks | Where-Object Kind -eq 'cards')
    }
}

function Get-ReadmeIcon([string]$Name, [string]$Class = 'icon') {
    $path = $assets.Icons[$Name]; if (-not $path) { $path = $assets.Icons['info'] }
    "<svg class=""$Class"" viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" stroke-width=""1.7"" stroke-linecap=""round"" stroke-linejoin=""round"">$path</svg>"
}

function ConvertTo-ReadmeFlow([string[]]$Lines, [switch]$Vertical) {
    # Vertical: the nodes are stacked, icon on the left, with a downward arrow and its label.
    $items = foreach ($l in $Lines) {
        $icon, $title, $sub = $l.Split('|', 3).ForEach({ $_.Trim() })
        $title = [System.Net.WebUtility]::HtmlEncode($title); $sub = [System.Net.WebUtility]::HtmlEncode($sub)
        if ($Vertical) {
            if ($icon -eq 'arrow') {
                $note = if ($sub) { "<span class=""flow-sub"">$sub</span>" } else { '' }
                "<div class=""rb-varrow""><svg viewBox=""0 0 12 30""><path d=""M6 1v26M1 21l5 6 5-6"" fill=""none"" stroke=""currentColor"" stroke-width=""1.6""/></svg><span class=""flow-label"">$title</span>$note</div>"
            } else {
                "<div class=""rb-vnode""><div class=""flow-icon"">$(Get-ReadmeIcon $icon)</div><div><div class=""flow-title"">$title</div><div class=""flow-text"">$sub</div></div></div>"
            }
        } elseif ($icon -eq 'arrow') {
            $class = if ($title -or $sub) { 'flow-arrow' } else { 'flow-arrow rb-bare' }
            "<div class=""$class""><span class=""flow-label"">$title</span><svg viewBox=""0 0 40 12""><path d=""M0 6h36M31 1l6 5-6 5"" fill=""none"" stroke=""currentColor"" stroke-width=""1.6""/></svg><span class=""flow-sub"">$sub</span></div>"
        } else {
            "<div class=""flow-node""><div class=""flow-icon"">$(Get-ReadmeIcon $icon)</div><div class=""flow-title"">$title</div><div class=""flow-text"">$sub</div></div>"
        }
    }
    $class = if ($Vertical) { 'flow rb-vflow' } else { 'flow rb-flow' }
    "<div class=""$class"">$($items -join '')</div>"
}

function ConvertTo-ReadmeCards([string[]]$Lines, [string]$Class = '') {
    $items = foreach ($l in $Lines) {
        $icon, $title, $text = $l.Split('|', 3).ForEach({ $_.Trim() })
        "<div class=""card-item""><div class=""card-icon"">$(Get-ReadmeIcon $icon)</div><div><div class=""card-title"">$(ConvertTo-ReadmeInline $title)</div><div class=""card-text"">$(ConvertTo-ReadmeInline $text)</div></div></div>"
    }
    "<div class=""cards $Class"">$($items -join '')</div>"
}

function Get-ReadmePill([string]$Text, [string]$Tone) { "<span class=""rb-pill"" style=""--tone: var(--cp-$Tone)"">$Text</span>" }
#endregion

#region Styles of the graphics, on top of the CSS of the guide -------------------------------------
$Script:ReadmeCss = @'
html, body { background: #ffffff; }
html[data-theme="dark"], html[data-theme="dark"] body { background: #0d1117; }
:root { --cp-info: #0078d4; --cp-violet: #7c3aed; --cp-teal: #0d9488; }
html[data-theme="dark"] { --cp-info: #4da6ff; --cp-violet: #a78bfa; --cp-teal: #2dd4bf; }
body { display: block; margin: 0; padding: 0; }
.canvas { padding: 6px; }
.rb-pill { display: inline-block; padding: 1px 10px; margin: 8px 6px 0 0; border-radius: 999px; font-size: 11.5px; font-weight: 600; line-height: 1.6;
  color: var(--tone); background: color-mix(in srgb, var(--tone) 11%, transparent); border: 1px solid color-mix(in srgb, var(--tone) 38%, transparent); }
.rb-caption { font-size: 11.5px; font-weight: 700; letter-spacing: 0.1em; text-transform: uppercase; color: var(--cp-accent); margin: 0 0 8px 4px; }
.rb-caption span { color: var(--cp-text-muted); font-weight: 600; letter-spacing: 0.04em; text-transform: none; font-size: 12.5px; }
/* Before / after */
.rb-bench { display: grid; gap: 14px; padding: 20px 22px; border-radius: 16px; background: var(--cp-surface); border: 1px solid var(--cp-border); }
.rb-row { display: grid; grid-template-columns: 250px minmax(0, 1fr); gap: 18px; align-items: center; }
.rb-row .label { font-size: 13.5px; font-weight: 650; color: var(--cp-text); } .rb-row .label span { display: block; font-weight: 500; font-size: 12px; color: var(--cp-text-muted); margin-top: 2px; }
.rb-bars { display: grid; gap: 6px; }
.rb-bar { display: flex; align-items: center; gap: 10px; font-size: 12.5px; color: var(--cp-text-muted); }
.rb-bar i { display: block; height: 18px; border-radius: 6px; min-width: 6px; }
.rb-bar.old i { background: color-mix(in srgb, var(--cp-text-muted) 45%, transparent); } .rb-bar.new i { background: var(--cp-accent); }
.rb-bar b { color: var(--cp-text); font-weight: 650; } .rb-bar .bad { color: var(--cp-danger); font-weight: 600; }
.rb-legend { display: flex; gap: 18px; font-size: 12px; color: var(--cp-text-muted); margin: 0 0 4px 4px; } .rb-legend i { display: inline-block; width: 12px; height: 12px; border-radius: 4px; margin-right: 6px; vertical-align: -1px; }
/* Banner */
.rb-hero { margin: 0; padding: 32px 36px 30px; }
.rb-hero-grid { position: relative; display: grid; grid-template-columns: minmax(0, 1fr) 240px; gap: 34px; align-items: center; }
.rb-hero h1 { font-size: 35px; }
.rb-hero .lead { margin: 18px 0 0; font-size: 17px; max-width: none; }
.rb-hero .badges { margin: 20px 0 0; }
.rb-stats { position: relative; display: grid; gap: 10px; }
.rb-stat { display: flex; align-items: center; gap: 14px; padding: 12px 16px; border-radius: 14px; background: var(--cp-panel-strong); border: 1px solid var(--cp-border); box-shadow: 0 1px 2px rgba(0, 0, 0, 0.08); }
.rb-stat b { font-size: 30px; line-height: 1; color: var(--cp-accent); font-weight: 750; min-width: 40px; text-align: center; }
.rb-stat span { font-size: 13px; color: var(--cp-text-muted); line-height: 1.35; }
.rb-stat strong { display: block; color: var(--cp-text); font-size: 14px; }
/* Cards and flows */
.cards { margin: 0; }
.rb-cards2 { grid-template-columns: 1fr 1fr; }
.rb-flow { margin: 0; flex-wrap: nowrap; padding: 18px; gap: 4px; }
.rb-flow .flow-node { flex: 1 1 0; min-width: 0; padding: 14px 10px; }
.rb-flow .flow-title { font-size: 13.5px; overflow-wrap: anywhere; }
.rb-flow .flow-arrow { min-width: 0; width: 84px; flex: 0 0 84px; }
.rb-flow .flow-arrow.rb-bare { width: 46px; flex-basis: 46px; }
.rb-flow .flow-sub { max-width: 84px; }
.rb-space { height: 18px; }
/* How it works: vertical pipeline and the three modes */
.rb-hiw { display: grid; grid-template-columns: minmax(0, 1.08fr) minmax(0, 1fr); gap: 16px; align-items: stretch; }
.rb-col { display: flex; flex-direction: column; }
.rb-vflow { flex: 1; flex-direction: column; flex-wrap: nowrap; align-items: stretch; justify-content: center; gap: 0; margin: 0; padding: 16px 18px; }
.rb-vnode { display: flex; align-items: center; gap: 14px; padding: 11px 16px; border-radius: 12px; background: var(--cp-surface); border: 1px solid var(--cp-border); }
.rb-vnode .flow-icon { margin: 0; flex-shrink: 0; }
.rb-vnode .flow-text { margin-top: 1px; }
.rb-varrow { display: flex; align-items: center; gap: 10px; min-height: 36px; padding-left: 31px; }
.rb-varrow svg { width: 12px; height: 28px; color: var(--cp-accent); flex-shrink: 0; }
.rb-varrow .flow-sub { max-width: none; font-size: 12px; }
.rb-modes { flex: 1; display: flex; flex-direction: column; gap: 10px; }
.rb-modes .card-item { flex: 1; align-items: center; }
.rb-modes .card-title { display: flex; align-items: center; gap: 8px; }
.rb-chip { font-size: 11px; font-weight: 600; padding: 0 8px; border-radius: 999px; border: 1px solid var(--cp-border); color: var(--cp-text-muted); }
.rb-chip.hot { color: var(--cp-accent-fg); background: var(--cp-accent); border-color: var(--cp-accent); }
'@
#endregion

#region Rendering (Microsoft Edge, headless) -------------------------------------------------------
function Save-Screenshot([string]$Html, [string]$Png, [int]$Width, [int]$Height, [int]$Scale = 1) {
    $url = 'file:///' + ($Html -replace '\\', '/')
    $profilePath = Join-Path $work 'edge-profile'
    if (Test-Path $Png) { Remove-Item $Png -Force }
    # Start-Process, not &: an Edge helper process can keep the output pipe open after the capture.
    $edgeArgs = @('--headless=new', '--disable-gpu', '--hide-scrollbars', '--no-first-run', "--user-data-dir=`"$profilePath`"", "--window-size=$Width,$Height", "--force-device-scale-factor=$Scale", "--screenshot=`"$Png`"", "`"$url`"")
    $proc = Start-Process -FilePath $edge -ArgumentList $edgeArgs -PassThru -WindowStyle Hidden
    $deadline = (Get-Date).AddSeconds(45)
    while (-not (Test-Path $Png) -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 300 }
    if (-not $proc.WaitForExit(10000)) { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue }
    if (-not (Test-Path $Png)) { throw "Screenshot not written: $Png" }
}

function Get-PageHeight([string]$Html, [int]$Width) {
    # Height of the .canvas element: the page writes it in body[data-h], read with --dump-dom.
    $url = 'file:///' + ($Html -replace '\\', '/')
    $dom = Join-Path $work ('dom-' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.html')
    $edgeArgs = @('--headless=new', '--disable-gpu', '--hide-scrollbars', '--no-first-run', "--user-data-dir=`"$(Join-Path $work 'edge-profile')`"", "--window-size=$Width,2000", '--dump-dom', "`"$url`"")
    $proc = Start-Process -FilePath $edge -ArgumentList $edgeArgs -PassThru -WindowStyle Hidden -RedirectStandardOutput $dom
    if (-not $proc.WaitForExit(45000)) { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue }
    # Edge helper processes inherit the output handle: read in shared mode, retry until written.
    $m = $null
    for ($i = 0; $i -lt 20 -and -not ($m -and $m.Success); $i++) {
        $stream = [IO.File]::Open($dom, 'Open', 'Read', 'ReadWrite')
        try { $text = [IO.StreamReader]::new($stream).ReadToEnd() } finally { $stream.Dispose() }
        $m = [regex]::Match($text, 'data-h="(\d+)"')
        if (-not $m.Success) { Start-Sleep -Milliseconds 250 }
    }
    if (-not $m.Success) { throw "Height not measured: $Html" }
    return [int]$m.Groups[1].Value
}

function New-ReadmeGraphic {
    # One graphic, light and dark: HTML page -> height measured by Edge -> 2x screenshot.
    param([string]$Name, [string]$Body, [int]$Width)
    $pages = @{}
    foreach ($theme in 'light', 'dark') {
        $html = "<!doctype html><html lang=""en"" data-theme=""$theme""><head><meta charset=""utf-8""><style>$($assets.Css)`n$($Script:ReadmeCss)</style></head>" +
            "<body><div class=""canvas"" style=""width:$($Width)px"">$Body</div><script>document.body.setAttribute('data-h', Math.ceil(document.querySelector('.canvas').getBoundingClientRect().height));</script></body></html>"
        $pages[$theme] = Join-Path $work "readme-$Name-$theme.html"
        [IO.File]::WriteAllText($pages[$theme], $html, [Text.UTF8Encoding]::new($false))
    }
    $height = Get-PageHeight $pages['light'] $Width
    foreach ($theme in 'light', 'dark') { Save-Screenshot $pages[$theme] (Join-Path $OutputFolder "readme-$Name-$theme.png") $Width $height 2 }
}
#endregion

#region Main ---------------------------------------------------------------------------------------
$edge = @("${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $edge) { throw 'Microsoft Edge not found: it takes the screenshots (headless mode).' }
$work = Join-Path ([IO.Path]::GetTempPath()) ('rlr-readme-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $work, $OutputFolder -Force | Out-Null
$Script:assets = Get-ReadmeAssets -Root $root
if ($assets.Flows.Count -lt 1 -or $assets.Cards.Count -lt 2) { throw 'The guide must hold a flow block (how it works) and 2 cards blocks (overview, why the message trace).' }
$mid = '&middot;'

try {
    Write-Host 'Rendering the README graphics (light and dark, 2x)...'

    # Banner: the hero of the guide, with the key figures.
    $badges = @(
        "<span class=""badge badge-accent"">Version $($assets.Version)</span>"
        "<span class=""badge"">$(Get-ReadmeIcon 'terminal' 'icon-sm')PowerShell 7.4+</span>"
        "<span class=""badge"">$(Get-ReadmeIcon 'search' 'icon-sm')Message trace, Microsoft Graph</span>"
        "<span class=""badge"">$(Get-ReadmeIcon 'database' 'icon-sm')SQLite, nothing to install</span>"
        "<span class=""badge"">$(Get-ReadmeIcon 'tag' 'icon-sm')MIT license</span>"
    ) -join ''
    $banner = "<header class=""hero rb-hero""><div class=""rb-hero-grid""><div>" +
        "<div class=""hero-top""><div class=""hero-logo"">$(Get-ReadmeIcon 'mail')</div><div><div class=""eyebrow"">Exchange Online $mid Message trace $mid PowerShell</div><h1>Recipient Limit Report</h1></div></div>" +
        "<p class=""lead"">The messages sent to <strong>more than 25 recipients</strong>, counted as Exchange Online counts them &mdash; <strong>a distribution list is one recipient</strong> &mdash; from the <strong>message trace</strong>, as <strong>CSV and HTML</strong> reports, one row per Message ID. The report of Purview DLP Report, without any DLP policy.</p>" +
        "<div class=""badges"">$badges</div></div>" +
        "<div class=""rb-stats"">" +
        "<div class=""rb-stat""><b>25</b><span><strong>recipients</strong>the limit, configurable</span></div>" +
        "<div class=""rb-stat""><b>1</b><span><strong>per distribution list</strong>count before expansion</span></div>" +
        "<div class=""rb-stat""><b>0</b><span><strong>DLP policy needed</strong>the message trace sees all</span></div>" +
        "</div></div></header>"
    New-ReadmeGraphic -Name 'banner' -Body $banner -Width 1080

    # Why: what the report answers (first cards block of the guide) and why the message trace (second).
    $why = "<div class=""rb-caption"">What the report answers <span>$mid before and after a recipient limit is applied</span></div>" +
        (ConvertTo-ReadmeCards $assets.Cards[0].Lines 'rb-cards2') + "<div class=""rb-space""></div>" +
        "<div class=""rb-caption"">Why the message trace <span>$mid the same report as Purview DLP Report, another source</span></div>" +
        (ConvertTo-ReadmeCards $assets.Cards[1].Lines)
    New-ReadmeGraphic -Name 'principles' -Body $why -Width 1080

    # How it works: the flow of chapter 2, as a vertical pipeline, and the three modes.
    $modes = @(
        [pscustomobject]@{ Icon = 'file'; Name = 'Report'; Chip = '<span class="rb-chip hot">default</span>'; Text = 'A period: collects and counts <strong>only what the database lacks</strong>, then writes the CSV and HTML files. The same report again takes seconds.'; Pills = (Get-ReadmePill 'Last24Hours' 'info') + (Get-ReadmePill 'Last7Days' 'info') + (Get-ReadmePill 'PreviousMonth' 'violet') + (Get-ReadmePill 'Day · Month · Custom' 'teal') }
        [pscustomobject]@{ Icon = 'calendar'; Name = 'Collect'; Chip = '<span class="rb-chip">scheduled</span>'; Text = 'Every day: the history <strong>beyond the 90 days</strong> of the message trace, only the messages over the limit kept.'; Pills = (Get-ReadmePill 'Incremental' 'success') + (Get-ReadmePill 'Restartable' 'teal') + (Get-ReadmePill '90 requests / 5 min' 'warning') }
        [pscustomobject]@{ Icon = 'chart'; Name = 'Status'; Chip = '<span class="rb-chip">no connection</span>'; Text = 'What the database holds, day by day, what is still to count and the last runs.'; Pills = (Get-ReadmePill 'Complete' 'success') + (Get-ReadmePill 'Partial' 'warning') }
    )
    $modeHtml = ($modes | ForEach-Object { "<div class=""card-item""><div class=""card-icon"">$(Get-ReadmeIcon $_.Icon)</div><div><div class=""card-title"">$($_.Name) $($_.Chip)</div><div class=""card-text"">$($_.Text)</div><div>$($_.Pills)</div></div></div>" }) -join ''
    $howItWorks = "<div class=""rb-hiw""><div class=""rb-col""><div class=""rb-caption"">One entry point <span>$mid from the message trace to the report</span></div>$(ConvertTo-ReadmeFlow $assets.Flows[0].Lines -Vertical)</div>" +
        "<div class=""rb-col""><div class=""rb-caption"">Three modes <span>$mid read-only for Microsoft 365</span></div><div class=""rb-modes"">$modeHtml</div></div></div>"
    New-ReadmeGraphic -Name 'how-it-works' -Body $howItWorks -Width 1080

    # Same report as Purview DLP Report: the 48-hour comparison window of Annex B and C.
    $bench = @(
        @{ Label = 'Messages of the DLP report'; Sub = '48-hour window, lab'; Old = 149802; OldText = '<b>149,802</b> &mdash; Purview DLP Report (Activity Explorer)'; New = 149802; NewText = '<b>149,802 / 149,802</b> found, same recipient count' }
        @{ Label = 'Messages over 25 recipients'; Sub = 'same window'; Old = 149802; OldText = '149,802 &mdash; DLP rule matches only'; New = 150154; NewText = '<b>150,154</b> &mdash; <b>+352</b> the DLP rule never recorded' }
        @{ Label = 'Messages sent to lists'; Sub = 'pilot, 10 messages'; Old = 10; OldText = '10 &mdash; recipients as addressed'; New = 10; NewText = '<b>10 / 10</b> same recipient list, members left out' }
    )
    $rows = foreach ($b in $bench) {
        $max = [double][Math]::Max($b.Old, $b.New)
        "<div class=""rb-row""><div class=""label"">$($b.Label)<span>$($b.Sub)</span></div><div class=""rb-bars"">" +
        "<div class=""rb-bar old""><i style=""width:$([Math]::Round(52 * $b.Old / $max, 1))%""></i>$($b.OldText)</div>" +
        "<div class=""rb-bar new""><i style=""width:$([Math]::Round(52 * $b.New / $max, 1))%""></i>$($b.NewText)</div></div></div>"
    }
    $benchHtml = "<div class=""rb-caption"">Measured in a lab tenant <span>$mid same period, both tools, empty databases</span></div>" +
        "<div class=""rb-legend""><span><i style=""background:color-mix(in srgb, var(--cp-text-muted) 45%, transparent)""></i>Purview DLP Report</span><span><i style=""background:var(--cp-accent)""></i>Recipient Limit Report</span></div>" +
        "<div class=""rb-bench"">$($rows -join '')</div>"
    New-ReadmeGraphic -Name 'benchmark' -Body $benchHtml -Width 1080

    Get-ChildItem $OutputFolder -Filter 'readme-*.png' | Select-Object Name, @{ n = 'KB'; e = { [math]::Round($_.Length / 1KB) } } | Format-Table -AutoSize | Out-String | Write-Host
} finally {
    # Edge helper processes of the temporary profile, if any are left.
    Get-CimInstance Win32_Process -Filter "Name='msedge.exe'" | Where-Object { $_.CommandLine -like "*$work*" } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    if ($KeepWork) { Write-Host "Work folder: $work" } else { Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue }
}
#endregion
