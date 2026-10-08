#Requires -Version 7.4
<#
.SYNOPSIS
    Copies the files needed to run Recipient Limit Report into a separate folder, ready to be zipped.

.DESCRIPTION
    The package contains only what Invoke-RecipientLimitReport.ps1 needs at run time, plus the HTML guide
    and the licence notice of the SQLite binaries:
        Invoke-RecipientLimitReport.ps1, RecipientLimitReport.psd1, RecipientLimitReport.psm1,
        config\, src\, templates\, lib\sqlite\, docs\RecipientLimitReport-Guide.html, THIRD-PARTY-NOTICES.md
    It never copies data\, reports\, logs\ or bin\: there is no database in the package, the tool
    creates an empty one at the first run.

    The configuration file is copied with the tenant values emptied (TenantId, Organization,
    UserPrincipalName, AppId, CertificateThumbprint) and without sender domains: the administrator fills
    them in (guide, chapter 6). The script then checks that none of these values appears anywhere in the
    package.

.PARAMETER Destination
    Package folder. Default: package\RecipientLimitReport-<version>, next to the tool folder.

.PARAMETER Force
    Replace the destination folder if it already contains a package. A folder that contains a data\
    sub-folder (a package that has been run) is never replaced.

.EXAMPLE
    .\tools\New-RlrPackage.ps1
    Creates ..\package\RecipientLimitReport-2.1.1.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.1
#>
[CmdletBinding()]
param(
    [string]$Destination,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$version = (Import-PowerShellDataFile (Join-Path $root 'RecipientLimitReport.psd1')).ModuleVersion
if (-not $Destination) { $Destination = Join-Path (Split-Path $root -Parent) "package\RecipientLimitReport-$version" }
$Destination = [IO.Path]::GetFullPath($Destination, (Get-Location).Path).TrimEnd('\')

$rootPrefix = [IO.Path]::GetFullPath($root).TrimEnd('\') + '\'
if (($Destination + '\').StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase) -or $rootPrefix.StartsWith($Destination + '\', [StringComparison]::OrdinalIgnoreCase)) {
    throw "The destination must be outside the tool folder: $Destination"
}
if (Test-Path -LiteralPath $Destination) {
    if (-not $Force) { throw "The destination already exists: $Destination. Use -Force to replace it." }
    if (-not (Test-Path -LiteralPath (Join-Path $Destination 'Invoke-RecipientLimitReport.ps1'))) { throw "The destination is not a Recipient Limit Report package, it is not replaced: $Destination" }
    if (Test-Path -LiteralPath (Join-Path $Destination 'data')) { throw "The destination contains a data folder (a database may be in it), it is not replaced: $Destination" }
    Remove-Item -LiteralPath $Destination -Recurse -Force
}

# ---- Files needed at run time ---------------------------------------------------------------------------
$files = [Collections.Generic.List[string]]::new()
foreach ($f in 'Invoke-RecipientLimitReport.ps1', 'RecipientLimitReport.psd1', 'RecipientLimitReport.psm1', 'THIRD-PARTY-NOTICES.md',
    'src\RecipientLimitReport.Engine.cs', 'templates\Report.template.html', 'docs\RecipientLimitReport-Guide.html') { $files.Add($f) }
Get-ChildItem -LiteralPath (Join-Path $root 'lib\sqlite') -Recurse -File | ForEach-Object { $files.Add($_.FullName.Substring($rootPrefix.Length)) }

foreach ($f in $files) {
    $source = Join-Path $root $f
    if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { throw "Missing file in the tool folder: $f" }
    $target = Join-Path $Destination $f
    [void][IO.Directory]::CreateDirectory((Split-Path $target -Parent))
    Copy-Item -LiteralPath $source -Destination $target
}

# ---- Configuration with the tenant values emptied -------------------------------------------------------
$configRelative = 'config\RecipientLimitReport.config.psd1'
$config = [IO.File]::ReadAllText((Join-Path $root $configRelative))
$emptied = [Collections.Generic.List[string]]::new()
foreach ($key in 'TenantId', 'Organization', 'UserPrincipalName', 'AppId', 'CertificateThumbprint') {
    $pattern = "(?m)^(\s*$key\s*=\s*)'([^']*)'"
    $found = [regex]::Matches($config, $pattern)
    if ($found.Count -ne 1) { throw "The key $key must appear exactly once in $configRelative (found $($found.Count))." }
    if ($found[0].Groups[2].Value) { $emptied.Add($found[0].Groups[2].Value) }
    $config = [regex]::Replace($config, $pattern, '$1''''')
}
# Sender domains of the organisation: emptied too (@() = every sender).
$domains = [regex]::Matches($config, "(?m)^(\s*SenderDomains\s*=\s*)@\(([^)]*)\)")
if ($domains.Count -ne 1) { throw "The key SenderDomains must appear exactly once in $configRelative (found $($domains.Count))." }
foreach ($m in [regex]::Matches($domains[0].Groups[2].Value, "'([^']+)'")) { $emptied.Add($m.Groups[1].Value) }
$config = [regex]::Replace($config, "(?m)^(\s*SenderDomains\s*=\s*)@\([^)]*\)", '$1@()')
$configTarget = Join-Path $Destination $configRelative
[void][IO.Directory]::CreateDirectory((Split-Path $configTarget -Parent))
[IO.File]::WriteAllText($configTarget, $config, [Text.UTF8Encoding]::new($true))
$files.Add($configRelative)

# ---- Checks ---------------------------------------------------------------------------------------------
$problems = [Collections.Generic.List[string]]::new()
foreach ($name in 'data', 'reports', 'logs', 'bin') {
    if (Test-Path -LiteralPath (Join-Path $Destination $name)) { $problems.Add("Folder $name\ must not be in the package.") }
}
Get-ChildItem -LiteralPath $Destination -Recurse -File -Include '*.sqlite', '*.sqlite-*', '*.db' | ForEach-Object { $problems.Add("Database file in the package: $($_.Name)") }
$textFiles = Get-ChildItem -LiteralPath $Destination -Recurse -File | Where-Object Extension -in '.ps1', '.psm1', '.psd1', '.cs', '.html', '.md'
foreach ($value in $emptied) {
    foreach ($file in $textFiles) {
        if ([IO.File]::ReadAllText($file.FullName).IndexOf($value, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
            $problems.Add("A tenant value of the configuration appears in $($file.FullName.Substring($Destination.Length + 1)).")
        }
    }
}
if ($problems.Count) { throw ("Package not valid ($Destination):`n - " + ($problems -join "`n - ")) }

$all = Get-ChildItem -LiteralPath $Destination -Recurse -File
Write-Host ''
Write-Host "  Recipient Limit Report $version - package ready" -ForegroundColor Green
Write-Host "  Folder   : $Destination"
Write-Host ("  Content  : {0} files, {1:N1} MB" -f $all.Count, (($all | Measure-Object Length -Sum).Sum / 1MB))
Write-Host "  Database : none - the tool creates an empty database at the first run"
Write-Host "  Config   : tenant values emptied ($($emptied.Count)) - fill in Tenant, Target and Authentication (guide, chapter 6)"
Write-Host ''
$all | Sort-Object FullName | ForEach-Object { '    {0,12:N0}  {1}' -f $_.Length, $_.FullName.Substring($Destination.Length + 1) }
