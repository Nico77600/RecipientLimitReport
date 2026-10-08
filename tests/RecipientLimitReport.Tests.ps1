#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
    Recipient Limit Report - automated tests (Pester 5 or later).
    Author  : Nicolas Fabert
    Version : 1.0.1

    Run:  Invoke-Pester -Path .\tests\RecipientLimitReport.Tests.ps1 -Output Detailed

    No connection to Microsoft 365 is made: Microsoft Graph is replaced by an in-memory message
    trace API (tests\FakeGraph.cs) that behaves like the real one measured in the lab: one value per
    property in $filter, newest first, paging, distribution lists expanded into 'expanded' rows,
    getDetailsByRecipient with the Submit / Receive / Expand DL / DLP rule events, and failures on
    demand (429, 401, 500, 403).
#>

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    $script:Root = Join-Path $script:RepoRoot 'package'
    Import-Module (Join-Path $script:Root 'RecipientLimitReport.psd1') -Force
    Initialize-RlrEngine -Root $script:Root
    if (-not ('RlrTests.FakeGraph' -as [type])) {
        Add-Type -LiteralPath (Join-Path $PSScriptRoot 'FakeGraph.cs') -ReferencedAssemblies @(
            'System.Net.Http', 'System.Net.Primitives', 'System.Text.Json', 'System.Linq', 'System.Collections', 'System.Text.RegularExpressions',
            'System.Runtime', 'System.Threading', 'System.Threading.Tasks', 'System.Private.Uri', 'System.Memory', 'System.Text.Encodings.Web', 'netstandard')
    }
    $script:Paris = Get-RlrTimeZone 'Europe/Paris'

    # Test configuration = the delivered configuration file with fictitious tenant values
    # (the delivered file may contain real values or empty ones).
    $script:ConfigDirectory = Join-Path ([IO.Path]::GetTempPath()) ('RecipientLimitReportTests-' + [guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($script:ConfigDirectory)
    $script:ConfigPath = Join-Path $script:ConfigDirectory 'test.config.psd1'
    $text = [IO.File]::ReadAllText((Join-Path $script:Root 'config\RecipientLimitReport.config.psd1'))
    $values = [ordered]@{
        TenantId = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'; Organization = 'contoso.onmicrosoft.com'
        AppId = '11111111-2222-3333-4444-555555555555'; CertificateThumbprint = '0123456789ABCDEF0123456789ABCDEF01234567'
    }
    foreach ($key in $values.Keys) {
        $pattern = "(?m)^(\s*$key\s*=\s*)'[^']*'"
        if ([regex]::Matches($text, $pattern).Count -ne 1) { throw "The key $key must appear once in the delivered configuration." }
        $text = [regex]::Replace($text, $pattern, "`${1}'$($values[$key])'")
    }
    [IO.File]::WriteAllText($script:ConfigPath, $text, [Text.UTF8Encoding]::new($true))

    function New-TestSettings {
        <# Settings of the test configuration with a database and folders of their own. #>
        param([string]$Name = [guid]::NewGuid().ToString('N'), [string[]]$Domains)
        $s = Import-RlrConfiguration -Path $script:ConfigPath -Root $script:Root
        $dir = Join-Path $script:ConfigDirectory $Name
        $s.Storage.DatabasePath = Join-Path $dir 'data\test.sqlite'
        $s.Report.OutputPath = Join-Path $dir 'reports'
        $s.Logging.Path = Join-Path $dir 'logs'
        if ($Domains) { $s.Target.SenderDomains = $Domains }
        # The fake Graph has no quota: a wide limit keeps the tests fast (the limiter itself is still used).
        $s.Collection.MaxRequests = 100; $s.Collection.PeriodSeconds = 1
        $s.Counting.MaxRequests = 100; $s.Counting.PeriodSeconds = 1
        return $s
    }

    function New-TestConnection {
        [ordered]@{ Mode = 'Certificate'; Account = 'test'; Token = 'token-1'; ExpiresMs = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds() + 3600000
            Renew = { param($s, $c) @{ Token = 'token-2'; ExpiresMs = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds() + 3600000 } } }
    }

    function New-TestOptions {
        param([Parameter(Mandatory)]$Settings, [Parameter(Mandatory)]$Fake, [string]$Kind = 'Collection', [int]$PageSize = 5000)
        $o = New-RlrCollectorOptions $Settings -Kind $Kind
        $o.Handler = $Fake
        $o.GraphRoot = 'https://graph.test/v1.0'
        $o.PageSize = $PageSize
        $o.DefaultRetryAfterSeconds = 1
        return $o
    }

    function Get-Addresses { param([string]$Prefix, [int]$Count) 1..$Count | ForEach-Object { '{0}{1:000}@contoso.com' -f $Prefix, $_ } }

    function Invoke-TestRun {
        <# Collects [StartMs, EndMs) from the fake Graph, then counts the messages with distribution lists. #>
        param([Parameter(Mandatory)]$Store, [Parameter(Mandatory)]$Settings, [Parameter(Mandatory)]$Fake, [long]$StartMs, [long]$EndMs, [int]$PageSize = 5000, [long]$NowMs = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds(), [switch]$CollectOnly)
        $runId = $Store.StartRun('Test', 'test', 'host', '1.0.0', $null)
        $plan = Get-RlrCollectionPlan -Store $Store -StartMs $StartMs -EndMs $EndMs -Settings $Settings -NowMs $NowMs
        $collection = Invoke-RlrCollection -Store $Store -Settings $Settings -Connection (New-TestConnection) -Items $plan.Items -RunId $runId -Options (New-TestOptions $Settings $Fake -PageSize $PageSize) 6>$null
        if ($CollectOnly) { return [pscustomobject]@{ Plan = $plan; Collection = $collection; Counting = $null } }
        $items = $Store.GetCountItems($StartMs, $EndMs, $Settings.Target.RecipientLimit, $Settings.Counting.MaxAttempts, 0, [string[]]@($Settings.Target.SenderDomains))
        $counting = if ($items.Count) { Invoke-RlrCounting -Store $Store -Settings $Settings -Connection (New-TestConnection) -Items $items -Options (New-TestOptions $Settings $Fake -Kind Counting) 6>$null }
        [pscustomobject]@{ Plan = $plan; Collection = $collection; Counting = $counting }
    }
}

AfterAll {
    if ($script:ConfigDirectory -and (Test-Path -LiteralPath $script:ConfigDirectory)) { Remove-Item -LiteralPath $script:ConfigDirectory -Recurse -Force -ErrorAction SilentlyContinue }
}

Describe 'Configuration' {
    It 'loads the delivered configuration file (tenant values filled in)' {
        $s = Import-RlrConfiguration -Path $script:ConfigPath -Root $script:Root
        $s.Target.RecipientLimit | Should -Be 25
        @($s.Target.SenderDomains).Count | Should -Be 0
        $s.Report.Title | Should -Be 'Messages with more than 25 recipients'
        [IO.Path]::IsPathRooted($s.Storage.DatabasePath) | Should -BeTrue
        $s.Zone.Id | Should -Not -BeNullOrEmpty
    }
    It 'normalizes the sender domains and rejects an invalid one' {
        $path = Join-Path $TestDrive 'domains.psd1'
        (Get-Content $script:ConfigPath -Raw).Replace('SenderDomains  = @()', "SenderDomains  = @('*@Contoso.com', '@fabrikam.com', 'contoso.com')") | Set-Content -LiteralPath $path
        (Import-RlrConfiguration -Path $path -Root $script:Root).Target.SenderDomains | Should -Be @('contoso.com', 'fabrikam.com')
        (Get-Content $script:ConfigPath -Raw).Replace('SenderDomains  = @()', "SenderDomains  = @('not a domain')") | Set-Content -LiteralPath $path
        { Import-RlrConfiguration -Path $path -Root $script:Root } | Should -Throw -ExpectedMessage '*Target.SenderDomains*'
    }
    It 'reports every invalid value at once' {
        $path = Join-Path $TestDrive 'bad.psd1'
        $text = Get-Content $script:ConfigPath -Raw
        $text = $text.Replace('PageSize              = 5000', 'PageSize              = 9000').Replace("SplitBy                 = 'Week'", "SplitBy                 = 'Month'").Replace('RecipientLimit = 25', 'RecipientLimit = 0')
        Set-Content -LiteralPath $path -Value $text
        { Import-RlrConfiguration -Path $path -Root $script:Root } | Should -Throw -ExpectedMessage '*Target.RecipientLimit*Collection.PageSize*Report.SplitBy*'
    }
    It 'requires the application settings in Certificate mode' {
        $path = Join-Path $TestDrive 'cert.psd1'
        $text = (Get-Content $script:ConfigPath -Raw) -replace "(?m)^(\s*AppId\s*=\s*)'[^']*'", "`${1}''" -replace "(?m)^(\s*CertificateThumbprint\s*=\s*)'[^']*'", "`${1}''"
        Set-Content -LiteralPath $path -Value $text
        { Import-RlrConfiguration -Path $path -Root $script:Root } | Should -Throw -ExpectedMessage '*AppId*CertificateThumbprint*'
    }
}

Describe 'Console characters' {
    It 'uses only characters of the classic console fonts outside the emoji style' {
        # Repertoire of Consolas and Lucida Console (checked glyph by glyph): code page 437 and Latin-1.
        # The classic console has no font fallback: any other character is shown as an empty box.
        $safe = [Collections.Generic.HashSet[int]]::new()
        foreach ($c in (0x20..0x7E) + (0xA0..0xFF)) { [void]$safe.Add($c) }
        foreach ($c in '☺☻♥♦♣♠•◘○◙♂♀♪♫☼►◄↕‼¶§▬↨↑↓→←∟↔▲▼⌂₧ƒ⌐░▒▓│┤╡╢╖╕╣║╗╝╜╛┐└┴┬├─┼╞╟╚╔╩╦╠═╬╧╨╤╥╙╘╒╓╫╪┘┌█▄▌▐▀αΓπΣστΦΘΩδ∞φε∩≡≥≤⌠⌡≈∙√ⁿ■'.ToCharArray()) { [void]$safe.Add([int]$c) }
        $used = [Collections.Generic.List[string]]::new()
        $sets = (Get-RlrIconSet 'Symbols'), (Get-RlrIconSet 'Ascii'), (Get-RlrFrameSet 'Symbols' 'Lucida Console'), (Get-RlrFrameSet 'Symbols' 'Terminal'), (Get-RlrFrameSet 'Ascii' $null)
        foreach ($set in $sets) { foreach ($key in $set.Keys) { $used.Add("set.$key=$($set[$key])") } }
        # Rounded corners: only for the fonts that have them (Consolas: checked glyph by glyph).
        $rounded = (Get-RlrFrameSet 'Symbols' 'Consolas'), (Get-RlrFrameSet 'Symbols' $null)
        foreach ($set in $rounded) { foreach ($key in $set.Keys) { if ([int]$set[$key] -notin 0x256D, 0x256E, 0x256F, 0x2570) { $used.Add("rounded.$key=$($set[$key])") } } }
        ($rounded[0].TopLeft, $sets[2].TopLeft) | Should -Be @([char]0x256D, [char]0x250C)
        # Characters written directly by the module and the script (the two style functions excluded).
        foreach ($file in 'RecipientLimitReport.psm1', 'Invoke-RecipientLimitReport.ps1') {
            $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $script:Root $file), [ref]$null, [ref]$null)
            $text = $ast.Extent.Text
            $skip = $ast.FindAll({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -in 'Get-RlrIconSet', 'Get-RlrFrameSet' }, $true)
            foreach ($f in @($skip) | Sort-Object { $_.Extent.StartOffset } -Descending) { $text = $text.Remove($f.Extent.StartOffset, $f.Extent.EndOffset - $f.Extent.StartOffset) }
            foreach ($m in [regex]::Matches($text, '\[char\]0x([0-9A-Fa-f]{4})')) { $used.Add("${file}=$([char][Convert]::ToInt32($m.Groups[1].Value, 16))") }
        }
        $bad = @($used | Where-Object { $v = $_.Substring($_.IndexOf('=') + 1); @($v.ToCharArray() | Where-Object { -not $safe.Contains([int]$_) }).Count -gt 0 })
        $used.Count | Should -BeGreaterThan 40
        $bad | Should -BeNullOrEmpty
    }
}

Describe 'Entry script' {
    It 'Status on a new installation confirms the configuration and creates no database' {
        $dir = Join-Path $TestDrive 'fresh'
        $path = Join-Path $TestDrive 'fresh.psd1'
        $text = Get-Content $script:ConfigPath -Raw
        $text = $text.Replace("'.\data\RecipientLimitReport.sqlite'", "'$dir\data\RecipientLimitReport.sqlite'").Replace("Path          = '.\logs'", "Path          = '$dir\logs'")
        Set-Content -LiteralPath $path -Value $text
        $output = & pwsh -NoProfile -File (Join-Path $script:Root 'Invoke-RecipientLimitReport.ps1') -Mode Status -ConfigPath $path 2>&1 | Out-String
        $LASTEXITCODE | Should -Be 0
        $output | Should -Match 'Ready for the first collection'
        Test-Path -LiteralPath (Join-Path $dir 'data') | Should -BeFalse
    }
}

Describe 'Periods' {
    BeforeAll { $script:Now = [DateTimeOffset]::Parse('2026-11-15T10:30:45.678Z') }
    It 'Last24Hours ends now, rounded to the second' {
        $p = Resolve-RlrPeriod -Range Last24Hours -Zone $script:Paris -Now $script:Now
        $p.EndMs | Should -Be ([DateTimeOffset]::Parse('2026-11-15T10:30:45Z').ToUnixTimeMilliseconds())
        ($p.EndMs - $p.StartMs) | Should -Be 86400000
    }
    It 'PreviousMonth follows the report time zone and summer time' {
        $p = Resolve-RlrPeriod -Range PreviousMonth -Zone $script:Paris -Now $script:Now
        $p.StartMs | Should -Be ([DateTimeOffset]::Parse('2026-09-30T22:00:00Z').ToUnixTimeMilliseconds())
        $p.EndMs | Should -Be ([DateTimeOffset]::Parse('2026-10-31T23:00:00Z').ToUnixTimeMilliseconds())
    }
    It 'Custom dates without offset are local; the end is capped to now' {
        $p = Resolve-RlrPeriod -Range Custom -Start '2026-11-15 08:00' -End '2026-11-20' -Zone $script:Paris -Now $script:Now
        $p.StartMs | Should -Be ([DateTimeOffset]::Parse('2026-11-15T07:00:00Z').ToUnixTimeMilliseconds())
        $p.EndMs | Should -Be ([DateTimeOffset]::Parse('2026-11-15T10:30:45Z').ToUnixTimeMilliseconds())
    }
    It 'rejects a missing -Month' { { Resolve-RlrPeriod -Range Month -Zone $script:Paris -Now $script:Now } | Should -Throw '*yyyy-MM*' }
}

Describe 'Coverage arithmetic' {
    It 'finds the gaps of a period and the intersection of two coverages' {
        $covered = New-RlrRangeList @([pscustomobject]@{ Start = 10; End = 20 }, [pscustomobject]@{ Start = 30; End = 40 })
        $gaps = [RecipientLimitReport.Coverage]::Gaps($covered, 0, 50)
        ($gaps | ForEach-Object { "$($_.Start)-$($_.End)" }) -join ',' | Should -Be '0-10,20-30,40-50'
        $other = New-RlrRangeList @([pscustomobject]@{ Start = 15; End = 35 })
        $both = [RecipientLimitReport.Coverage]::Intersect($covered, $other)
        ($both | ForEach-Object { "$($_.Start)-$($_.End)" }) -join ',' | Should -Be '15-20,30-35'
    }
    It 'cuts slices at local midnight and returns the newest first' {
        $start = [DateTimeOffset]::Parse('2026-09-28T20:00:00Z').ToUnixTimeMilliseconds()
        $end = [DateTimeOffset]::Parse('2026-09-29T08:00:00Z').ToUnixTimeMilliseconds()
        $slices = [RecipientLimitReport.Coverage]::SplitIntoSlices((New-RlrRangeList @([pscustomobject]@{ Start = $start; End = $end })), $script:Paris, 6)
        $slices.Count | Should -Be 3
        $slices[0].End | Should -Be $end
        $midnight = [DateTimeOffset]::Parse('2026-09-28T22:00:00Z').ToUnixTimeMilliseconds()
        @($slices | Where-Object { $_.Start -lt $midnight -and $_.End -gt $midnight }).Count | Should -Be 0
    }
}

Describe 'Recipient count in the route' {
    It 'takes RcptCount of the Submit event and never the counts of the other events' {
        $body = '{"value":[{"event":"Receive","data":"<root/>"},{"event":"Submit","data":"<root><MEP Name=\"RcptCount\" Integer=\"28\" /></root>"},{"event":"Expand DL","data":"<root><MEP Name=\"RcptCount\" Integer=\"4\" /></root>"},{"event":"DLP rule","data":"<root><MEP Name=\"RcptCount\" Integer=\"55\" /></root>"}]}'
        $next = $null
        $e = [RecipientLimitReport.TraceParser]::ParseCount($body, [ref]$next)
        $e.Count | Should -Be 28
        $e.Source | Should -Be 'Submit'
    }
    It 'uses the Receive event of a message received by SMTP, and returns no count otherwise' {
        $next = $null
        $smtp = [RecipientLimitReport.TraceParser]::ParseCount('{"value":[{"event":"Receive","data":"<root><MEP Name=\"RcptCount\" Integer=\"30\" /></root>"},{"event":"DLP rule","data":"<root><MEP Name=\"RcptCount\" Integer=\"80\" /></root>"}]}', [ref]$next)
        ($smtp.Count, $smtp.Source) | Should -Be @(30, 'Receive')
        $none = [RecipientLimitReport.TraceParser]::ParseCount('{"value":[{"event":"Expand DL","data":"<root><MEP Name=\"RcptCount\" Integer=\"4\" /></root>"}]}', [ref]$next)
        $none.Count | Should -BeNullOrEmpty
    }
    It 'tells a recipient addressed by the sender from a member added by the expansion of a list' {
        # Lab, 2026-10-05: the route of a direct recipient starts with Receive and Submit, the route of a member after them.
        $next = $null
        $direct = [RecipientLimitReport.TraceParser]::ParseCount('{"value":[{"event":"Receive","data":"<root/>"},{"event":"Submit","data":"<root><MEP Name=\"RcptCount\" Integer=\"27\" /></root>"},{"event":"Fail","data":"<root><MEP Name=\"RcptCount\" Integer=\"1\" /></root>"}]}', [ref]$next)
        $member = [RecipientLimitReport.TraceParser]::ParseCount('{"value":[{"event":"Fail","data":"<root><MEP Name=\"RcptCount\" Integer=\"1\" /></root>"},{"event":"DLP rule","data":"<root><MEP Name=\"RcptCount\" Integer=\"54\" /></root>"}]}', [ref]$next)
        [RecipientLimitReport.CountItem]::IsDirect($direct, 'Submit') | Should -BeTrue
        [RecipientLimitReport.CountItem]::IsDirect($member, 'Submit') | Should -BeFalse
        [RecipientLimitReport.CountItem]::IsDirect($member, 'Receive') | Should -BeFalse
    }
}

Describe 'Collection and count (fake Graph)' {
    BeforeEach {
        $script:Settings = New-TestSettings
        $script:Store = Open-RlrStore -Settings $script:Settings
        $script:Fake = [RlrTests.FakeGraph]::new()
        $script:T0 = [DateTimeOffset]::UtcNow.AddDays(-2).Date
        $script:Start = [DateTimeOffset]::new($script:T0, [TimeSpan]::Zero).ToUnixTimeMilliseconds()
        $script:End = $script:Start + 86400000
        $script:Big = (Get-Addresses 'big' 30)
    }
    AfterEach { $script:Store.Dispose() }

    It 'counts a message without distribution list from its trace rows (no route read)' {
        [void]$script:Fake.AddMessage('alice@contoso.com', $script:Big, [DateTimeOffset]::new($script:T0.AddHours(10), [TimeSpan]::Zero), 'Thirty')
        [void]$script:Fake.AddMessage('alice@contoso.com', (Get-Addresses 'small' 25), [DateTimeOffset]::new($script:T0.AddHours(11), [TimeSpan]::Zero), 'Twenty-five')
        $run = Invoke-TestRun -Store $script:Store -Settings $script:Settings -Fake $script:Fake -StartMs $script:Start -EndMs $script:End -PageSize 7
        $run.Collection.Completed | Should -BeTrue
        $run.Counting | Should -BeNullOrEmpty
        $s = $script:Store.GetPeriodSummary($script:Start, $script:End, 25, 3, [string[]]@())
        ($s.Messages, $s.OverLimit, $s.WithLists, $s.Matches) | Should -Be @(2, 1, 0, 1)
        @($script:Fake.Requests | Where-Object { $_ -match 'getDetailsByRecipient' }).Count | Should -Be 0
    }
    It 'reads the count before expansion of a message sent to distribution lists, one request per message' {
        $script:Store.MaxRoutesPerMessage = 0   # count only: the recipient lists are not rebuilt
        $lists = [Collections.Generic.Dictionary[string, string[]]]::new()
        $lists['dl-sales@contoso.com'] = Get-Addresses 'sales' 40
        $lists['dl-it@contoso.com'] = Get-Addresses 'it' 10
        [void]$script:Fake.AddMessage('alice@contoso.com', (Get-Addresses 'direct' 3), [DateTimeOffset]::new($script:T0.AddHours(9), [TimeSpan]::Zero), 'Lists under', $lists)
        $lists2 = [Collections.Generic.Dictionary[string, string[]]]::new()
        $lists2['dl-all@contoso.com'] = Get-Addresses 'all' 5
        [void]$script:Fake.AddMessage('bob@contoso.com', (Get-Addresses 'direct' 26), [DateTimeOffset]::new($script:T0.AddHours(8), [TimeSpan]::Zero), 'Lists over', $lists2)
        $run = Invoke-TestRun -Store $script:Store -Settings $script:Settings -Fake $script:Fake -StartMs $script:Start -EndMs $script:End
        $run.Counting.Done | Should -Be 2
        @($script:Fake.Requests | Where-Object { $_ -match 'getDetailsByRecipient' }).Count | Should -Be 2
        # 3 direct + 2 lists = 5 (under the limit, although 55 rows); 26 direct + 1 list = 27.
        $s = $script:Store.GetPeriodSummary($script:Start, $script:End, 25, 3, [string[]]@())
        ($s.OverLimit, $s.WithLists, $s.Measured, $s.Matches, $s.ListsPending, $s.ListsNotRebuilt) | Should -Be @(2, 2, 2, 1, 0, 1)
    }
    It 'rebuilds the recipient list before expansion and stops once every addressed recipient is found' {
        # 26 direct + 1 list of 31 members, one of them also addressed directly: 57 rows, 27 recipients before expansion.
        $lists = [Collections.Generic.Dictionary[string, string[]]]::new(); $lists['dl-all@contoso.com'] = @(Get-Addresses 'm' 30) + 'd001@contoso.com'
        [void]$script:Fake.AddMessage('bob@contoso.com', (Get-Addresses 'd' 26), [DateTimeOffset]::new($script:T0.AddHours(8), [TimeSpan]::Zero), 'First', $lists)
        $run = Invoke-TestRun -Store $script:Store -Settings $script:Settings -Fake $script:Fake -StartMs $script:Start -EndMs $script:End
        ($run.Counting.Done, $run.Counting.ListsDone) | Should -Be @(1, 1)
        # The list (count), then d001 ... d026: the 30 members are never read.
        $routes = @($script:Fake.Requests | Where-Object { $_ -match 'getDetailsByRecipient' })
        $routes.Count | Should -Be 27
        @($routes | Where-Object { $_ -match "recipientAddress='m\d+" }).Count | Should -Be 0
        $s = $script:Store.GetPeriodSummary($script:Start, $script:End, 25, 3, [string[]]@())
        ($s.Matches, $s.ListsRebuilt, $s.ListsPending) | Should -Be @(1, 1, 0)

        # Next day, the same list with direct recipients sorted after the members: the members already
        # seen in the list are read last, so they are not read at all.
        $lists2 = [Collections.Generic.Dictionary[string, string[]]]::new(); $lists2['dl-all@contoso.com'] = $lists['dl-all@contoso.com']
        [void]$script:Fake.AddMessage('bob@contoso.com', (Get-Addresses 'z' 26), [DateTimeOffset]::new($script:T0.AddHours(32), [TimeSpan]::Zero), 'Second', $lists2)
        $script:Fake.Requests.Clear()
        [void](Invoke-TestRun -Store $script:Store -Settings $script:Settings -Fake $script:Fake -StartMs $script:End -EndMs ($script:End + 86400000))
        $routes = @($script:Fake.Requests | Where-Object { $_ -match 'getDetailsByRecipient' })
        @($routes | Where-Object { $_ -match "recipientAddress='m\d+" }).Count | Should -Be 0
        $routes.Count | Should -Be 28   # the list, d001 (a member this time, not seen as one before), z001 ... z026

        $period = [pscustomobject]@{ Range = 'Custom'; StartMs = $script:Start; EndMs = $script:End + 86400000 }
        $r = New-RlrReport -Store $script:Store -Settings $script:Settings -Period $period -Coverage (Get-RlrCoverage -Store $script:Store -Settings $script:Settings -StartMs $period.StartMs -EndMs $period.EndMs)
        $csv = Import-Csv (@($r.Result.Files | Where-Object Kind -eq 'CSV')[0].Path) -Delimiter ';'
        foreach ($row in $csv) {
            $people = $row.Recipients -split '; '
            $people.Count | Should -Be 27
            $people[0] | Should -Be 'dl-all@contoso.com'
            @($people | Where-Object { $_ -like 'm*' }).Count | Should -Be 0
        }
        ($csv | Where-Object Subject -eq 'First').Recipients | Should -Match 'd001@contoso.com'
        ($csv | Where-Object Subject -eq 'Second').Recipients | Should -Not -Match 'd001@contoso.com'
    }
    It 'announces the range of routes before counting: one per count at least, every listed route at most' {
        # Over the limit before expansion: 26 direct + 1 list of 30 (57 rows). Under: 3 direct + 1 list of 40 (44 rows).
        $over = [Collections.Generic.Dictionary[string, string[]]]::new(); $over['dl-all@contoso.com'] = Get-Addresses 'm' 30
        [void]$script:Fake.AddMessage('bob@contoso.com', (Get-Addresses 'd' 26), [DateTimeOffset]::new($script:T0.AddHours(8), [TimeSpan]::Zero), 'Over', $over)
        $under = [Collections.Generic.Dictionary[string, string[]]]::new(); $under['dl-big@contoso.com'] = Get-Addresses 'b' 40
        [void]$script:Fake.AddMessage('alice@contoso.com', (Get-Addresses 'u' 3), [DateTimeOffset]::new($script:T0.AddHours(9), [TimeSpan]::Zero), 'Under', $under)
        [void](Invoke-TestRun -Store $script:Store -Settings $script:Settings -Fake $script:Fake -StartMs $script:Start -EndMs $script:End -CollectOnly)
        $items = $script:Store.GetCountItems($script:Start, $script:End, 25, 3, 0, [string[]]@())
        $range = Get-RlrCountingRoutes $items
        ($range.Min, $range.Max) | Should -Be @(2, 101)
        $counting = Invoke-RlrCounting -Store $script:Store -Settings $script:Settings -Connection (New-TestConnection) -Items $items -Options (New-TestOptions $script:Settings $script:Fake -Kind Counting) 6>$null
        # 'Under': the list only (4 recipients). 'Over': the list, then d001 ... d026.
        $counting.Requests | Should -Be 28
        $counting.Requests | Should -BeGreaterOrEqual $range.Min
        $counting.Requests | Should -BeLessOrEqual $range.Max

        # A count already read: the recipients the sender addressed and not found yet are the minimum.
        $known = [Collections.Generic.List[RecipientLimitReport.CountItem]]::new()
        $item = [RecipientLimitReport.CountItem]::new(); $item.Limit = 25; $item.RebuildList = $true; $item.Count = 27; $item.DirectKnown = 5
        foreach ($a in Get-Addresses 'k' 40) { $item.Targets.Add($a) }
        $known.Add($item)
        $range = Get-RlrCountingRoutes $known
        ($range.Min, $range.Max) | Should -Be @(22, 40)
        (Get-RlrCountingRoutes ([Collections.Generic.List[RecipientLimitReport.CountItem]]::new())).Max | Should -Be 0
    }
    It 'rebuilds the list of a message received by SMTP from the Receive event' {
        $lists = [Collections.Generic.Dictionary[string, string[]]]::new(); $lists['dl-a@contoso.com'] = Get-Addresses 'a' 30
        [void]$script:Fake.AddMessage('app@contoso.com', (Get-Addresses 'r' 26), [DateTimeOffset]::new($script:T0.AddHours(7), [TimeSpan]::Zero), 'Relay', $lists, 'delivered', 'Smtp')
        $run = Invoke-TestRun -Store $script:Store -Settings $script:Settings -Fake $script:Fake -StartMs $script:Start -EndMs $script:End
        $run.Counting.ListsDone | Should -Be 1
        $r = New-RlrReport -Store $script:Store -Settings $script:Settings -Period ([pscustomobject]@{ Range = 'Custom'; StartMs = $script:Start; EndMs = $script:End }) -Coverage (Get-RlrCoverage -Store $script:Store -Settings $script:Settings -StartMs $script:Start -EndMs $script:End)
        $row = Import-Csv (@($r.Result.Files | Where-Object Kind -eq 'CSV')[0].Path) -Delimiter ';'
        $row.'Recipient count' | Should -Be '27'
        @($row.Recipients -split '; ' | Where-Object { $_ -like 'r*' }).Count | Should -Be 26
        @($row.Recipients -split '; ' | Where-Object { $_ -like 'a0*' }).Count | Should -Be 0
    }
    It 'keeps the list after expansion when MaxRoutesPerMessage is reached, and continues it no further' {
        $script:Store.MaxRoutesPerMessage = 10
        $lists = [Collections.Generic.Dictionary[string, string[]]]::new(); $lists['dl-all@contoso.com'] = Get-Addresses 'm' 30
        [void]$script:Fake.AddMessage('bob@contoso.com', (Get-Addresses 'd' 26), [DateTimeOffset]::new($script:T0.AddHours(8), [TimeSpan]::Zero), 'Big', $lists)
        $run = Invoke-TestRun -Store $script:Store -Settings $script:Settings -Fake $script:Fake -StartMs $script:Start -EndMs $script:End
        ($run.Counting.ListsDone, $run.Counting.ListsPartial) | Should -Be @(0, 1)
        @($script:Fake.Requests | Where-Object { $_ -match 'getDetailsByRecipient' }).Count | Should -Be 10
        $s = $script:Store.GetPeriodSummary($script:Start, $script:End, 25, 3, [string[]]@())
        ($s.Matches, $s.ListsRebuilt, $s.ListsPending, $s.ListsNotRebuilt) | Should -Be @(1, 0, 0, 1)
        $script:Store.GetCountItems($script:Start, $script:End, 25, 3, 0, [string[]]@()).Count | Should -Be 0
        $r = New-RlrReport -Store $script:Store -Settings $script:Settings -Period ([pscustomobject]@{ Range = 'Custom'; StartMs = $script:Start; EndMs = $script:End }) -Coverage (Get-RlrCoverage -Store $script:Store -Settings $script:Settings -StartMs $script:Start -EndMs $script:End)
        $r.Result.ListedAfterExpansion | Should -Be 1
        $row = Import-Csv (@($r.Result.Files | Where-Object Kind -eq 'CSV')[0].Path) -Delimiter ';'
        $row.Recipients | Should -BeLike 'dl-all@contoso.com; *; (addresses after distribution list expansion)'
        @($row.Recipients -split '; ').Count | Should -Be 58
    }
    It 'uses the Receive count of a message received by SMTP and reports a route without any count' {
        $script:Store.MaxRoutesPerMessage = 0   # count only
        $lists = [Collections.Generic.Dictionary[string, string[]]]::new(); $lists['dl-a@contoso.com'] = Get-Addresses 'a' 30
        [void]$script:Fake.AddMessage('app@contoso.com', (Get-Addresses 'd' 26), [DateTimeOffset]::new($script:T0.AddHours(7), [TimeSpan]::Zero), 'Relay', $lists, 'delivered', 'Smtp')
        [void]$script:Fake.AddMessage('app2@contoso.com', (Get-Addresses 'e' 2), [DateTimeOffset]::new($script:T0.AddHours(6), [TimeSpan]::Zero), 'Unknown', $lists, 'delivered', 'None')
        $run = Invoke-TestRun -Store $script:Store -Settings $script:Settings -Fake $script:Fake -StartMs $script:Start -EndMs $script:End
        ($run.Counting.Done, $run.Counting.NoCount) | Should -Be @(1, 1)
        $s = $script:Store.GetPeriodSummary($script:Start, $script:End, 25, 3, [string[]]@())
        ($s.Matches, $s.Unmeasurable, $s.Pending) | Should -Be @(1, 1, 0)
        # The route of a list without count is read twice: the list, then another recipient.
        @($script:Fake.Requests | Where-Object { $_ -match 'getDetailsByRecipient' }).Count | Should -Be 3
    }
    It 'waits after a 429, retries a 500 and renews the token after a 401' {
        [void]$script:Fake.AddMessage('alice@contoso.com', $script:Big, [DateTimeOffset]::new($script:T0.AddHours(10), [TimeSpan]::Zero), 'Thirty')
        $script:Fake.ThrottleNext = 1; $script:Fake.ServerErrorNext = 1; $script:Fake.UnauthorizedNext = 1
        $run = Invoke-TestRun -Store $script:Store -Settings $script:Settings -Fake $script:Fake -StartMs $script:Start -EndMs $script:End
        $run.Collection.Completed | Should -BeTrue
        $run.Collection.Throttled | Should -Be 1
        $script:Store.GetPeriodSummary($script:Start, $script:End, 25, 3, [string[]]@()).Matches | Should -Be 1
    }
    It 'stops on a permission error and keeps the period uncollected' {
        [void]$script:Fake.AddMessage('alice@contoso.com', $script:Big, [DateTimeOffset]::new($script:T0.AddHours(10), [TimeSpan]::Zero), 'Thirty')
        $script:Fake.Forbidden = $true
        $run = Invoke-TestRun -Store $script:Store -Settings $script:Settings -Fake $script:Fake -StartMs $script:Start -EndMs $script:End
        $run.Collection.Completed | Should -BeFalse
        $run.Collection.Fatal | Should -BeTrue
        $run.Collection.Error | Should -Match 'ExchangeMessageTrace.Read.All'
        (Get-RlrCoverage -Store $script:Store -Settings $script:Settings -StartMs $script:Start -EndMs $script:End).Percent | Should -Be 0
    }
    It 'collects the same period twice without duplicates and updates a changed status' {
        [void]$script:Fake.AddMessage('alice@contoso.com', $script:Big, [DateTimeOffset]::new($script:T0.AddHours(10), [TimeSpan]::Zero), 'Thirty', $null, 'pending')
        $now = [DateTimeOffset]::new($script:T0.AddHours(30), [TimeSpan]::Zero).ToUnixTimeMilliseconds()   # collected 6 h after the end: not settled
        [void](Invoke-TestRun -Store $script:Store -Settings $script:Settings -Fake $script:Fake -StartMs $script:Start -EndMs $script:End)
        foreach ($r in $script:Fake.Rows) { $r.Status = 'delivered' }
        $settings = $script:Settings; $settings.Collection.SettlingHours = 72   # everything is "recent": collected again
        $run = Invoke-TestRun -Store $script:Store -Settings $settings -Fake $script:Fake -StartMs $script:Start -EndMs $script:End
        $run.Plan.Items.Count | Should -BeGreaterThan 0
        $stats = $script:Store.GetStatistics(25, 3)
        ($stats.Messages, $script:Store.CountRecipientRows(), $stats.Matches) | Should -Be @(1, 30, 1)
    }
    It 'collects one query per sender domain and reports only those senders' {
        $settings = New-TestSettings -Domains @('contoso.com', 'fabrikam.com')
        $store = Open-RlrStore -Settings $settings
        try {
            [void]$script:Fake.AddMessage('alice@contoso.com', $script:Big, [DateTimeOffset]::new($script:T0.AddHours(10), [TimeSpan]::Zero), 'Internal')
            [void]$script:Fake.AddMessage('news@external.example', $script:Big, [DateTimeOffset]::new($script:T0.AddHours(11), [TimeSpan]::Zero), 'Newsletter')
            [void]$script:Fake.AddMessage('bob@fabrikam.com', $script:Big, [DateTimeOffset]::new($script:T0.AddHours(12), [TimeSpan]::Zero), 'Partner')
            $run = Invoke-TestRun -Store $store -Settings $settings -Fake $script:Fake -StartMs $script:Start -EndMs $script:End
            @($run.Plan.Items | Select-Object -ExpandProperty Signature -Unique) | Should -Be @('sender=*@contoso.com', 'sender=*@fabrikam.com')
            @($script:Fake.Requests | Where-Object { $_ -match "senderAddress eq '\*@contoso.com'" }).Count | Should -BeGreaterThan 0
            $store.GetPeriodSummary($script:Start, $script:End, 25, 3, $settings.Target.SenderDomains).Matches | Should -Be 2
            (Get-RlrCoverage -Store $store -Settings $settings -StartMs $script:Start -EndMs $script:End).Percent | Should -Be 100
        } finally { $store.Dispose() }
    }
    It 'plans only what is missing and flags what is older than the 90 days of the message trace' {
        [void](Invoke-TestRun -Store $script:Store -Settings $script:Settings -Fake $script:Fake -StartMs $script:Start -EndMs $script:End)
        $now = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()
        $old = $now - 100L * 86400000
        $plan = Get-RlrCollectionPlan -Store $script:Store -StartMs $old -EndMs $script:End -Settings $script:Settings -NowMs $now
        $plan.UnrecoverableMs | Should -BeGreaterThan (9L * 86400000)
        @($plan.Items | Where-Object { $_.StartMs -lt $script:End -and $_.EndMs -gt $script:Start }).Count | Should -Be 0
    }
}

Describe 'Compaction and retention' {
    BeforeEach {
        $script:Settings = New-TestSettings
        $script:Store = Open-RlrStore -Settings $script:Settings
        $script:Fake = [RlrTests.FakeGraph]::new()
        $script:T0 = [DateTimeOffset]::UtcNow.AddDays(-3).Date
        $script:Start = [DateTimeOffset]::new($script:T0, [TimeSpan]::Zero).ToUnixTimeMilliseconds()
        $script:End = $script:Start + 86400000
    }
    AfterEach { $script:Store.Dispose() }

    It 'keeps only the messages over the limit once the period has settled' {
        $lists = [Collections.Generic.Dictionary[string, string[]]]::new(); $lists['dl@contoso.com'] = Get-Addresses 'm' 40
        [void]$script:Fake.AddMessage('a@contoso.com', (Get-Addresses 'x' 30), [DateTimeOffset]::new($script:T0.AddHours(10), [TimeSpan]::Zero), 'Over')
        [void]$script:Fake.AddMessage('a@contoso.com', (Get-Addresses 'y' 3), [DateTimeOffset]::new($script:T0.AddHours(11), [TimeSpan]::Zero), 'Small')
        [void]$script:Fake.AddMessage('a@contoso.com', (Get-Addresses 'z' 2), [DateTimeOffset]::new($script:T0.AddHours(12), [TimeSpan]::Zero), 'List under', $lists)
        [void](Invoke-TestRun -Store $script:Store -Settings $script:Settings -Fake $script:Fake -StartMs $script:Start -EndMs $script:End)
        (Invoke-RlrCompaction -Store $script:Store -Settings $script:Settings) | Should -Be 2
        $stats = $script:Store.GetStatistics(25, 3)
        ($stats.Messages, $stats.Matches, $stats.CompactionLimit) | Should -Be @(1, 1, 25)
    }
    It 'keeps a message with distribution lists until its count is read' {
        $lists = [Collections.Generic.Dictionary[string, string[]]]::new(); $lists['dl@contoso.com'] = Get-Addresses 'm' 40
        [void]$script:Fake.AddMessage('a@contoso.com', (Get-Addresses 'z' 2), [DateTimeOffset]::new($script:T0.AddHours(12), [TimeSpan]::Zero), 'List', $lists)
        $runId = $script:Store.StartRun('Test', 'test', 'host', '1.0.0', $null)
        $plan = Get-RlrCollectionPlan -Store $script:Store -StartMs $script:Start -EndMs $script:End -Settings $script:Settings
        [void](Invoke-RlrCollection -Store $script:Store -Settings $script:Settings -Connection (New-TestConnection) -Items $plan.Items -RunId $runId -Options (New-TestOptions $script:Settings $script:Fake) 6>$null)
        (Invoke-RlrCompaction -Store $script:Store -Settings $script:Settings) | Should -Be 0
        $script:Store.GetStatistics(25, 3).Pending | Should -Be 1
    }
    It 'purges the messages and the coverage older than the retention' {
        [void]$script:Fake.AddMessage('a@contoso.com', (Get-Addresses 'x' 30), [DateTimeOffset]::new($script:T0.AddHours(10), [TimeSpan]::Zero), 'Over')
        [void](Invoke-TestRun -Store $script:Store -Settings $script:Settings -Fake $script:Fake -StartMs $script:Start -EndMs $script:End)
        $p = $script:Store.PurgeBefore($script:End)
        ($p.Messages, $p.Rows) | Should -Be @(1, 30)
        (Get-RlrCoverage -Store $script:Store -Settings $script:Settings -StartMs $script:Start -EndMs $script:End).Percent | Should -Be 0
    }
}

Describe 'Report files' {
    BeforeAll {
        $script:Settings = New-TestSettings -Name 'report'
        $script:Store = Open-RlrStore -Settings $script:Settings
        $script:Fake = [RlrTests.FakeGraph]::new()
        $script:T0 = [DateTimeOffset]::UtcNow.AddDays(-20).Date
        $script:Start = [DateTimeOffset]::new($script:T0, [TimeSpan]::Zero).ToUnixTimeMilliseconds()
        $script:End = $script:Start + 14L * 86400000
        $lists = [Collections.Generic.Dictionary[string, string[]]]::new(); $lists['dl-sales@contoso.com'] = Get-Addresses 'sales' 40
        $script:ListId = $script:Fake.AddMessage('bob@contoso.com', (Get-Addresses 'direct' 26), [DateTimeOffset]::new($script:T0.AddHours(30), [TimeSpan]::Zero), 'With list', $lists)
        [void]$script:Fake.AddMessage('=cmd@contoso.com', (Get-Addresses 'x' 27), [DateTimeOffset]::new($script:T0.AddHours(40), [TimeSpan]::Zero), '+SUM(1;2)')
        # The same Message ID under two trace IDs (for example back from an on-premises server): one row.
        [void]$script:Fake.AddMessage('carol@contoso.com', (Get-Addresses 'y' 28), [DateTimeOffset]::new($script:T0.AddHours(50), [TimeSpan]::Zero), 'Twice', $null, 'delivered', 'Mailbox', '<same@contoso.com>')
        [void]$script:Fake.AddMessage('carol@contoso.com', (Get-Addresses 'y' 28), [DateTimeOffset]::new($script:T0.AddHours(50).AddSeconds(5), [TimeSpan]::Zero), 'Twice', $null, 'delivered', 'Mailbox', '<same@contoso.com>')
        foreach ($d in 0..9) { [void]$script:Fake.AddMessage("weekly$d@contoso.com", (Get-Addresses 'w' 31), [DateTimeOffset]::new($script:T0.AddDays(3 + $d).AddHours(9), [TimeSpan]::Zero), "Weekly $d") }
        [void](Invoke-TestRun -Store $script:Store -Settings $script:Settings -Fake $script:Fake -StartMs $script:Start -EndMs $script:End)
        $script:Period = [pscustomobject]@{ Range = 'Custom'; StartMs = $script:Start; EndMs = $script:End }
        $script:Coverage = Get-RlrCoverage -Store $script:Store -Settings $script:Settings -StartMs $script:Start -EndMs $script:End
    }
    AfterAll { $script:Store.Dispose() }

    It 'writes one row per Message ID with the business columns and the count before expansion' {
        $r = New-RlrReport -Store $script:Store -Settings $script:Settings -Period $script:Period -Coverage $script:Coverage
        $r.Result.Messages | Should -Be 13
        $csv = Import-Csv (@($r.Result.Files | Where-Object Kind -eq 'CSV')[0].Path) -Delimiter ';'
        ($csv[0].PSObject.Properties.Name -join '|') | Should -Be 'Received time (Europe/Paris)|Sender|Recipients|Subject|Recipient count|Distribution lists|Message ID'
        $list = $csv | Where-Object Subject -eq 'With list'
        ($list.'Recipient count', $list.'Distribution lists') | Should -Be @('27', '1')
        # As addressed by the sender: the list, then the 26 direct recipients (the 40 members are left out).
        $list.Recipients | Should -BeLike 'dl-sales@contoso.com; direct001@contoso.com; *'
        @($list.Recipients -split '; ').Count | Should -Be 27
        @($csv | Where-Object 'Message ID' -eq '<same@contoso.com>').Count | Should -Be 1
    }
    It 'omits the recipient addresses when IncludeRecipientDetails is false and cuts long lists' {
        $s = New-TestSettings -Name 'report'
        $s.Report.IncludeRecipientDetails = $false
        $r = New-RlrReport -Store $script:Store -Settings $s -Period $script:Period -Coverage $script:Coverage
        $csv = Import-Csv (@($r.Result.Files | Where-Object Kind -eq 'CSV')[0].Path) -Delimiter ';'
        $csv[0].PSObject.Properties.Name | Should -Not -Contain 'Recipients'
        $s.Report.IncludeRecipientDetails = $true; $s.Report.MaxRecipientsListed = 10
        $r = New-RlrReport -Store $script:Store -Settings $s -Period $script:Period -Coverage $script:Coverage
        $csv = Import-Csv (@($r.Result.Files | Where-Object Kind -eq 'CSV')[0].Path) -Delimiter ';'
        ($csv | Where-Object Subject -eq 'With list').Recipients | Should -BeLike '*; (+17 more)'
    }
    It 'splits by week above MaxRowsPerFile and names the files by period' {
        $s = New-TestSettings -Name 'report'
        $s.Report.MaxRowsPerFile = 1000
        $request = [RecipientLimitReport.ReportRequest]::new()
        $request.MaxRowsPerFile = 5; $request.SplitBy = 'Week'; $request.Zone = $script:Paris; $request.StartMs = $script:Start; $request.EndMs = $script:End; $request.FilePrefix = 'X'
        $times = [Collections.Generic.List[long]]::new(); foreach ($d in 0..12) { $times.Add($script:Start + $d * 86400000L + 3600000) }
        $plan = [RecipientLimitReport.ReportPlanner]::Plan($times, $request)
        $plan.Count | Should -BeGreaterThan 2
        $plan[0].BaseName | Should -Match '^X_\d{4}-\d{2}-\d{2}'
    }
    It 'writes a self-contained HTML file with the data in compressed blocks' {
        $r = New-RlrReport -Store $script:Store -Settings $script:Settings -Period $script:Period -Coverage $script:Coverage -Unmeasured 2
        $html = Get-Content (@($r.Result.Files | Where-Object Kind -eq 'HTML')[0].Path) -Raw
        $html | Should -Match 'application/x-rlr-chunk'
        $html | Should -Not -Match '%%CHUNKS%%|%%META%%'
        $html | Should -Match '"limit":25'
        $html | Should -Match '"unmeasured":2'
        $html | Should -Match '"afterExpansion":0'
    }
    It 'protects CSV cells from formula injection' {
        $r = New-RlrReport -Store $script:Store -Settings $script:Settings -Period $script:Period -Coverage $script:Coverage
        $text = Get-Content (@($r.Result.Files | Where-Object Kind -eq 'CSV')[0].Path) -Raw
        $text | Should -Match "'=cmd@contoso.com"
        $text | Should -Match "'\+SUM\(1;2\)|`"'\+SUM\(1;2\)`""
    }
}
