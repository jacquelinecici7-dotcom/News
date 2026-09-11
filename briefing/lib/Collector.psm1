<#
    Collector.psm1 - fetching / parsing / validation helpers for the daily briefing collector.

    Design notes:
      * Everything goes through curl.exe. Windows PowerShell 5.1's Invoke-WebRequest
        defaults to the IE parsing engine and old TLS settings; curl.exe avoids both.
      * No source is ever allowed to kill a run. Every fetch returns a status object
        and the caller records OK / FAILED so failures show up in the output rather
        than silently shrinking the briefing.
      * HTTP 200 does NOT mean success. Several sources return a styled HTML 404 page
        with status 200, so feeds are validated by shape (parses as XML, has items).
#>

$script:UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36'

function Write-Utf8 {
    <# Writes UTF-8 without a BOM. PS 5.1's Set-Content/Out-File would add one. #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Content
    )
    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    $enc = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Content, $enc)
}

function Read-BytesAsText {
    <# Decodes a downloaded file, honouring a charset hint in the markup (GB2312 etc). #>
    param([Parameter(Mandatory)][string]$Path)
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -eq 0) { return '' }

    $probeLen = [Math]::Min(4096, $bytes.Length)
    $probe = [System.Text.Encoding]::ASCII.GetString($bytes, 0, $probeLen)
    $enc = [System.Text.Encoding]::UTF8
    $m = [regex]::Match($probe, '(?i)charset\s*=\s*["'']?([\w\-]+)')
    if ($m.Success -and $m.Groups[1].Value -notmatch '(?i)^utf-?8$') {
        try { $enc = [System.Text.Encoding]::GetEncoding($m.Groups[1].Value) } catch { $enc = [System.Text.Encoding]::UTF8 }
    }
    return ($enc.GetString($bytes)).TrimStart([char]0xFEFF)
}

function Invoke-Fetch {
    <# Single HTTP GET. Never throws; returns .Ok / .Http / .Text / .Error. #>
    param(
        [Parameter(Mandatory)][string]$Url,
        [int]$TimeoutSec = 25
    )
    # On Linux/macOS there's no %TEMP% env var; /tmp is the universal fallback.
    $tmpDir = if ($env:TEMP) { $env:TEMP } else { '/tmp' }
    $tmp = Join-Path $tmpDir ('brf_' + [guid]::NewGuid().ToString('N') + '.bin')
    # PowerShell aliases `curl` to Invoke-WebRequest on Windows but to native
    # curl on PS Core/Linux. Use the explicit executable name to avoid that.
    $curlCmd = if ($IsLinux -or $IsMacOS) { 'curl' } else { 'curl.exe' }
    $out = [ordered]@{ Url = $Url; Http = '000'; Ok = $false; Text = ''; Error = '' }
    try {
        $code = & $curlCmd -s -L -o $tmp -w '%{http_code}' --max-time $TimeoutSec -A $script:UA $Url 2>$null
        $out.Http = ([string]$code).Trim()
        if (Test-Path -LiteralPath $tmp) { $out.Text = Read-BytesAsText -Path $tmp }
        if ($out.Http -eq '000') {
            $out.Error = '连接失败或超时'
        } elseif ($out.Http -ne '200') {
            $out.Error = "HTTP $($out.Http)"
        } elseif ($out.Text.Length -eq 0) {
            $out.Error = '响应为空'
        } else {
            $out.Ok = $true
        }
    } catch {
        $out.Error = $_.Exception.Message
    } finally {
        if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
    }
    return [pscustomobject]$out
}

function ConvertTo-LocalDate {
    <# Parses RFC822 / ISO8601 / compact-digit publication dates into local time. Returns $null on failure. #>
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $s = $Text.Trim()

    # Compact all-digit stamps come from URLs (yyyyMMdd, yyyyMMddHHmmss). These must
    # be handled before TryParse, which interprets bare digit runs unpredictably.
    if ($s -match '^\d{14}$') {
        try { return [datetime]::ParseExact($s, 'yyyyMMddHHmmss', [Globalization.CultureInfo]::InvariantCulture) } catch { return $null }
    }
    if ($s -match '^\d{8}$') {
        try { return [datetime]::ParseExact($s, 'yyyyMMdd', [Globalization.CultureInfo]::InvariantCulture) } catch { return $null }
    }

    $dto = [datetimeoffset]::MinValue
    if ([datetimeoffset]::TryParse($s, [ref]$dto)) { return $dto.ToLocalTime().DateTime }
    foreach ($f in @('ddd, dd MMM yyyy HH:mm:ss zzz', 'ddd, d MMM yyyy HH:mm:ss zzz', 'yyyy-MM-dd HH:mm:ss', 'yyyy-MM-dd')) {
        try { return [datetime]::ParseExact($s, $f, [Globalization.CultureInfo]::InvariantCulture) } catch { }
    }
    return $null
}

function Test-InWindow {
    <#
        Freshness test. Sources that expose only a date (no clock time) yield
        midnight; comparing those against an intraday cutoff would wrongly drop
        same-day releases - e.g. a CPI release dated 2026-09-09 against a cutoff of
        2026-09-09 06:28 - so date-only items are compared at date granularity.
    #>
    param($Date, [Parameter(Mandatory)][datetime]$Cutoff)
    if ($null -eq $Date) { return $true }
    $d = [datetime]$Date
    if ($d.TimeOfDay -eq [TimeSpan]::Zero) { return ($d.Date -ge $Cutoff.Date) }
    return ($d -ge $Cutoff)
}

function Get-NodeText {
    <# XML value that may be plain text, CDATA, or an element. #>
    param($Node)
    if ($null -eq $Node) { return '' }
    if ($Node -is [string]) { return $Node }
    try { return [string]$Node.InnerText } catch { return [string]$Node }
}

function Format-Link {
    <#
        Normalises a feed link. Some feeds emit malformed markup inside <link>,
        e.g. 'https://example.com/x" target="blank', which would otherwise be
        published into the briefing as a broken URL.
    #>
    param([string]$Url)
    if ([string]::IsNullOrWhiteSpace($Url)) { return '' }
    return ((($Url.Trim() -split '["<>\s]')[0]).Trim())
}

function Get-ItemLink {
    param($Item)
    $l = $null
    try { $l = $Item.link } catch { }
    if ($null -eq $l) { try { $l = $Item.guid } catch { } }
    if ($null -eq $l) { return '' }

    $raw = ''
    if ($l -is [string]) {
        $raw = $l
    } elseif ($l -is [array]) {
        foreach ($c in $l) {
            if ($c -is [string]) { $raw = $c; break }
            $h = $null; try { $h = $c.href } catch { }
            if ($h) { $raw = [string]$h; break }
        }
    } else {
        $h = $null; try { $h = $l.href } catch { }
        if ($h) { $raw = [string]$h } else { $raw = Get-NodeText $l }
    }
    return (Format-Link -Url $raw)
}

function ConvertFrom-Feed {
    <#
        Parses RSS or Atom. Throws with a plain-language reason when the payload is
        not actually a feed - that reason is surfaced in the run's status table.
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Xml)

    $t = $Xml.Trim()
    if ($t.Length -eq 0) { throw '响应为空' }
    if ($t -notmatch '<rss|<feed|<\?xml') { throw '返回的不是 XML（很可能是 HTML 错误页）' }

    $doc = $null
    try { $doc = [xml]$t } catch { throw ('XML 解析失败：' + $_.Exception.Message.Split([char]10)[0]) }

    $items = @()
    if ($doc.rss) { $items = @($doc.rss.channel.item) }
    elseif ($doc.feed) { $items = @($doc.feed.entry) }
    $items = @($items | Where-Object { $_ -ne $null })
    if ($items.Count -eq 0) { throw 'XML 可解析，但没有任何条目' }

    $out = New-Object System.Collections.ArrayList
    foreach ($it in $items) {
        $title = (Get-NodeText $it.title).Trim()
        if ($title.Length -eq 0) { continue }
        $rawDate = ''
        foreach ($p in 'pubDate', 'published', 'updated', 'date') {
            if ($rawDate) { break }
            try { $rawDate = (Get-NodeText $it.$p).Trim() } catch { }
        }
        # Summary gives the writing model substance beyond the headline (figures,
        # attribution). Markup and entities are stripped; long bodies are clipped.
        $summary = ''
        foreach ($p in 'description', 'summary', 'content') {
            if ($summary) { break }
            try { $summary = (Get-NodeText $it.$p).Trim() } catch { }
        }
        if ($summary) {
            $summary = $summary -replace '(?s)<[^>]+>', ' '
            $summary = ([System.Net.WebUtility]::HtmlDecode($summary)).Trim() -replace '\s+', ' '
            if ($summary.Length -gt 300) { $summary = $summary.Substring(0, 300).TrimEnd() + '…' }
        }
        [void]$out.Add([pscustomobject]@{
            Title   = ($title -replace '\s+', ' ')
            Link    = Get-ItemLink -Item $it
            Date    = ConvertTo-LocalDate -Text $rawDate
            Summary = $summary
        })
    }
    if ($out.Count -eq 0) { throw '条目均缺少标题' }
    return $out.ToArray()
}

function Resolve-Url {
    param([Parameter(Mandatory)][string]$Base, [Parameter(Mandatory)][string]$Href)
    try { return (New-Object System.Uri([System.Uri]$Base, $Href)).AbsoluteUri } catch { return $Href }
}

function Get-AnchorItems {
    <#
        Best-effort headline extraction from a listing page: pull <a> tags, keep the
        ones whose href matches HrefFilter, optionally read a yyyyMMdd date out of
        the href (many government sites encode the publication date there).
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Html,
        [Parameter(Mandatory)][string]$BaseUrl,
        [string]$HrefFilter = '',
        [string]$DateFromHref = '',
        [int]$Max = 20
    )
    if ([string]::IsNullOrWhiteSpace($Html)) { throw '响应为空' }

    $out = New-Object System.Collections.ArrayList
    $seen = @{}
    foreach ($m in [regex]::Matches($Html, '(?is)<a\s[^>]*href\s*=\s*"([^"]+)"[^>]*>(.*?)</a>')) {
        $href = $m.Groups[1].Value.Trim()
        if ($HrefFilter -and $href -notmatch $HrefFilter) { continue }

        $text = $m.Groups[2].Value -replace '(?s)<[^>]+>', ''
        $text = ([System.Net.WebUtility]::HtmlDecode($text)).Trim() -replace '\s+', ' '
        if ($text.Length -lt 6) { continue }

        $abs = Resolve-Url -Base $BaseUrl -Href $href
        if ($seen.ContainsKey($abs)) { continue }
        $seen[$abs] = $true

        $d = $null
        if ($DateFromHref) {
            $dm = [regex]::Match($href, $DateFromHref)
            if ($dm.Success) { $d = ConvertTo-LocalDate -Text $dm.Groups[1].Value }
        }
        [void]$out.Add([pscustomobject]@{ Title = $text; Link = $abs; Date = $d; Summary = '' })
        if ($out.Count -ge $Max) { break }
    }
    if ($out.Count -eq 0) { throw '页面可访问，但没有匹配到条目链接（站点结构可能已变化）' }
    return $out.ToArray()
}

function Get-CsvTail {
    <# Returns the header row plus the last N data rows of a CSV dataset. #>
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Text,
        [int]$Rows = 6,
        [int]$HeaderLine = 1
    )
    $lines = @($Text -split "`r?`n" | Where-Object { $_.Trim().Length -gt 0 })
    if ($lines.Count -lt ($HeaderLine + 1)) { throw 'CSV 内容不足（无法取得表头与数据行）' }
    $header = $lines[$HeaderLine]
    $data = @($lines[($HeaderLine + 1)..($lines.Count - 1)])
    $take = [Math]::Min($Rows, $data.Count)
    return [pscustomobject]@{
        Title  = $lines[0].Trim(',').Trim()
        Header = $header
        Rows   = @($data[($data.Count - $take)..($data.Count - 1)])
    }
}

function Get-MarketQuote {
    <#
        Live snapshot from the Yahoo Finance chart endpoint. Returns an object with
        .Ok=$false and a reason instead of throwing, so a dead ticker is reported
        rather than silently omitted (the brief must never invent a price).
    #>
    param(
        [Parameter(Mandatory)][string]$Symbol,
        [Parameter(Mandatory)][string]$Label,
        [string]$Group = '',
        [string]$Suffix = ''
    )
    $q = [ordered]@{
        Label = $Label; Symbol = $Symbol; Group = $Group; Suffix = $Suffix
        Ok = $false; Price = $null; PrevClose = $null; ChangePct = $null
        Currency = ''; AsOf = $null; Error = ''
    }
    # query1 and query2 are independent hosts; trying both survives a partial outage.
    $r = $null
    foreach ($apiHost in @('query1', 'query2')) {
        $url = ('https://{0}.finance.yahoo.com/v8/finance/chart/{1}?range=5d&interval=1d' -f $apiHost, [uri]::EscapeDataString($Symbol))
        $r = Invoke-Fetch -Url $url -TimeoutSec 20
        if ($r.Ok) { break }
    }
    if (-not $r.Ok) { $q.Error = $r.Error; return [pscustomobject]$q }

    try {
        $j = $r.Text | ConvertFrom-Json
        $meta = $j.chart.result[0].meta
        if ($null -eq $meta -or $null -eq $meta.regularMarketPrice) { $q.Error = '接口未返回价格'; return [pscustomobject]$q }

        $q.Price = [double]$meta.regularMarketPrice
        $q.Currency = [string]$meta.currency
        if ($null -ne $meta.chartPreviousClose -and [double]$meta.chartPreviousClose -ne 0) {
            $q.PrevClose = [double]$meta.chartPreviousClose
            $q.ChangePct = [Math]::Round((($q.Price - $q.PrevClose) / $q.PrevClose) * 100, 2)
        }
        if ($meta.regularMarketTime) {
            $q.AsOf = ([datetimeoffset]::FromUnixTimeSeconds([long]$meta.regularMarketTime)).ToLocalTime().DateTime
        }
        $q.Ok = $true
    } catch {
        $q.Error = 'JSON 解析失败'
    }
    return [pscustomobject]$q
}

function Send-Toast {
    <#
        Windows toast notification. Best-effort: notification failure must never
        fail a run that already produced the briefing. On non-Windows (Linux/macOS
        runners) this is a silent no-op.
    #>
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$Message
    )
    if ($IsLinux -or $IsMacOS) { return $false }
    try {
        # WinRT 类型只加载一次；重复加载会抛 "type already exists"，导致第二次及以后
        # 的 toast 全部静默失败。第一次跑成功后用户再也看不到通知，就是这个原因。
        if (-not ('Windows.UI.Notifications.ToastNotificationManager' -as [type])) {
            [void][Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]
        }
        $tpl = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02)
        $nodes = $tpl.GetElementsByTagName('text')
        [void]$nodes.Item(0).AppendChild($tpl.CreateTextNode($Title))
        [void]$nodes.Item(1).AppendChild($tpl.CreateTextNode($Message))
        $aumid = '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe'
        [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($aumid).Show(
            [Windows.UI.Notifications.ToastNotification]::new($tpl))
        return $true
    } catch {
        return $false
    }
}

function Limit-LogFile {
    <# Keeps the run log from growing without bound. #>
    param([Parameter(Mandatory)][string]$Path, [int]$MaxLines = 2000)
    if (-not (Test-Path -LiteralPath $Path)) { return }
    try {
        $lines = @([System.IO.File]::ReadAllLines($Path, [System.Text.Encoding]::UTF8))
        if ($lines.Count -le $MaxLines) { return }
        $keep = $lines[($lines.Count - $MaxLines)..($lines.Count - 1)]
        Write-Utf8 -Path $Path -Content (($keep -join "`r`n") + "`r`n")
    } catch { }
}

function ConvertFrom-MarkdownLite {
    <#
        Minimal Markdown -> HTML for the delivered document. Handles the subset both
        the digest and the briefing templates actually use: headings, bold, links,
        inline code, bullets, tables, fenced code, blockquotes.
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Markdown)

    function Convert-Inline([string]$s) {
        $s = [System.Net.WebUtility]::HtmlEncode($s)
        $s = [regex]::Replace($s, '`([^`]+)`', '<code>$1</code>')
        $s = [regex]::Replace($s, '\*\*([^*]+)\*\*', '<strong>$1</strong>')
        $s = [regex]::Replace($s, '\[([^\]]+)\]\((https?://[^)\s]+)\)', '<a href="$2" target="_blank">$1</a>')
        # bare URLs on their own
        $s = [regex]::Replace($s, '(?<!["=>])\b(https?://[^\s<]+)', '<a href="$1" target="_blank">$1</a>')
        return $s
    }

    $sb = New-Object System.Text.StringBuilder
    $lines = @($Markdown -split "`r?`n")
    $inList = $false; $inCode = $false; $inTable = $false

    foreach ($raw in $lines) {
        $line = $raw.TrimEnd()

        if ($line -match '^\s*```') {
            if ($inCode) { [void]$sb.AppendLine('</pre>'); $inCode = $false }
            else {
                if ($inList) { [void]$sb.AppendLine('</ul>'); $inList = $false }
                if ($inTable) { [void]$sb.AppendLine('</tbody></table>'); $inTable = $false }
                [void]$sb.AppendLine('<pre>'); $inCode = $true
            }
            continue
        }
        if ($inCode) { [void]$sb.AppendLine([System.Net.WebUtility]::HtmlEncode($raw)); continue }

        if ($line.Trim().Length -eq 0) {
            if ($inList) { [void]$sb.AppendLine('</ul>'); $inList = $false }
            if ($inTable) { [void]$sb.AppendLine('</tbody></table>'); $inTable = $false }
            continue
        }

        # table
        if ($line -match '^\s*\|.*\|\s*$') {
            if ($line -match '^\s*\|[\s\-:|]+\|\s*$') { continue }   # separator row
            $cells = @(($line.Trim().Trim('|') -split '\|') | ForEach-Object { $_.Trim() })
            if (-not $inTable) {
                if ($inList) { [void]$sb.AppendLine('</ul>'); $inList = $false }
                [void]$sb.AppendLine('<table><thead><tr>')
                foreach ($c in $cells) { [void]$sb.AppendLine('<th>' + (Convert-Inline $c) + '</th>') }
                [void]$sb.AppendLine('</tr></thead><tbody>')
                $inTable = $true
                continue
            }
            [void]$sb.AppendLine('<tr>')
            foreach ($c in $cells) { [void]$sb.AppendLine('<td>' + (Convert-Inline $c) + '</td>') }
            [void]$sb.AppendLine('</tr>')
            continue
        } elseif ($inTable) {
            [void]$sb.AppendLine('</tbody></table>'); $inTable = $false
        }

        # heading
        $h = [regex]::Match($line, '^(#{1,4})\s+(.*)$')
        if ($h.Success) {
            if ($inList) { [void]$sb.AppendLine('</ul>'); $inList = $false }
            $lvl = $h.Groups[1].Value.Length
            [void]$sb.AppendLine("<h$lvl>" + (Convert-Inline $h.Groups[2].Value) + "</h$lvl>")
            continue
        }

        # blockquote
        if ($line -match '^\s*>\s?(.*)$') {
            if ($inList) { [void]$sb.AppendLine('</ul>'); $inList = $false }
            [void]$sb.AppendLine('<blockquote>' + (Convert-Inline $matches[1]) + '</blockquote>')
            continue
        }

        # bullet
        if ($line -match '^\s*[-*]\s+(.*)$') {
            if (-not $inList) { [void]$sb.AppendLine('<ul>'); $inList = $true }
            [void]$sb.AppendLine('<li>' + (Convert-Inline $matches[1]) + '</li>')
            continue
        }

        # continuation line inside a list (indented)
        if ($inList -and $raw -match '^\s{2,}\S') {
            [void]$sb.AppendLine('<div class="cont">' + (Convert-Inline $line.Trim()) + '</div>')
            continue
        }

        if ($inList) { [void]$sb.AppendLine('</ul>'); $inList = $false }
        [void]$sb.AppendLine('<p>' + (Convert-Inline $line) + '</p>')
    }
    if ($inCode) { [void]$sb.AppendLine('</pre>') }
    if ($inList) { [void]$sb.AppendLine('</ul>') }
    if ($inTable) { [void]$sb.AppendLine('</tbody></table>') }
    return $sb.ToString()
}

function Publish-GitHubPages {
    <#
        Pushes a single HTML file to a GitHub repo via the Contents API, so it
        serves on GitHub Pages. Token is read from state/github.token and never
        stored in code or config. If the token file is missing, this is a no-op.

        Uses System.Net.Http.HttpClient directly because PowerShell 5.1's
        Invoke-RestMethod chokes on Chinese paths in URLs and certain header
        encodings.
    #>
    param(
        [Parameter(Mandatory)][string]$LocalPath,
        [Parameter(Mandatory)][string]$RepoPath,
        [Parameter(Mandatory)][string]$CommitMessage
    )

    $cfgPath = Join-Path $PSScriptRoot '..\config\github.json'
    $tokPath = Join-Path $PSScriptRoot '..\state\github.token'
    if (-not (Test-Path -LiteralPath $cfgPath)) { return @{ Ok = $false; Skipped = 'no-config' } }
    if (-not (Test-Path -LiteralPath $tokPath)) { return @{ Ok = $false; Skipped = 'no-token' } }

    $cfg = Read-JsonFromPath -Path $cfgPath
    # Strip UTF-8 BOM if present (Out-File -Encoding UTF8 in PS 5.1 writes one),
    # then trim whitespace and any leftover line terminators.
    $raw = [System.IO.File]::ReadAllText($tokPath, [System.Text.Encoding]::UTF8)
    $token = ($raw -replace '^\xEF\xBB\xBF', '').Trim()
    if (-not $token -or $token.StartsWith('#')) { return @{ Ok = $false; Skipped = 'empty-token' } }

    $owner = $cfg.owner; $repo = $cfg.repo; $branch = if ($cfg.branch) { $cfg.branch } else { 'main' }
    $fullPath = if ($cfg.pathPrefix) { '{0}/{1}' -f $cfg.pathPrefix.Trim('/'), $RepoPath.Trim('/') } else { $RepoPath.Trim('/') }
    # Encode each path segment so Chinese filenames don't break Uri parsing in PS 5.1.
    $encodedPath = ($fullPath -split '/') | ForEach-Object { [Uri]::EscapeDataString($_) }
    $encodedPath = $encodedPath -join '/'
    $apiBase = "https://api.github.com/repos/$owner/$repo/contents/$encodedPath"

    $bytes = [System.IO.File]::ReadAllBytes($LocalPath)
    $b64 = [Convert]::ToBase64String($bytes)

    # PowerShell 5.1 runs on .NET Framework — System.Net.Http isn't loaded by
    # default. Without this, New-Object HttpClient throws PSArgumentException.
    try { Add-Type -AssemblyName 'System.Net.Http' } catch { }

    $client = New-Object System.Net.Http.HttpClient
    $client.Timeout = [TimeSpan]::FromSeconds(30)
    try {
        [void]$client.DefaultRequestHeaders.Accept.Add([System.Net.Http.Headers.MediaTypeWithQualityHeaderValue]::new('application/vnd.github+json'))
        [void]$client.DefaultRequestHeaders.TryAddWithoutValidation('X-GitHub-Api-Version', '2022-11-28')
        [void]$client.DefaultRequestHeaders.TryAddWithoutValidation('User-Agent', 'briefing-deployer')
        $client.DefaultRequestHeaders.Authorization =
            [System.Net.Http.Headers.AuthenticationHeaderValue]::new('Bearer', $token)

        # Step 1: fetch existing file's SHA (if updating).
        $sha = $null
        try {
            $getUrl = $apiBase + '?ref=' + [Uri]::EscapeDataString($branch)
            $getResp = $client.GetAsync($getUrl).Result
            if ($getResp.IsSuccessStatusCode) {
                $getJson = $getResp.Content.ReadAsStringAsync().Result | ConvertFrom-Json
                if ($getJson.sha) { $sha = $getJson.sha }
            }
        } catch { }

        # Step 2: PUT the file.
        $body = [ordered]@{ message = $CommitMessage; content = $b64; branch = $branch }
        if ($sha) { $body.sha = $sha }
        $json = $body | ConvertTo-Json -Depth 5
        $content = [System.Net.Http.StringContent]::new($json, [System.Text.Encoding]::UTF8, 'application/json')

        $resp = $client.PutAsync($apiBase, $content).Result
        $respBody = $resp.Content.ReadAsStringAsync().Result
        if ($resp.IsSuccessStatusCode) {
            return @{ Ok = $true; Status = [int]$resp.StatusCode }
        } else {
            return @{ Ok = $false; Status = [int]$resp.StatusCode; Error = $respBody }
        }
    } finally {
        $client.Dispose()
    }
}

function Read-JsonFromPath {
    param([string]$Path)
    try { return ([System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8) | ConvertFrom-Json) } catch { return $null }
}

function ConvertTo-SummaryBox {
    <# Builds the trailing "今日要点" callout. Pure HTML so it stays visually distinct. #>
    param([Parameter(Mandatory)]$Result)

    $cfg = $Result.Config
    $okCount = @($Result.Status | Where-Object { $_.State -eq 'OK' }).Count
    $totalCount = @($Result.Status).Count
    $failedNames = ($Result.Status | Where-Object { $_.State -ne 'OK' } | ForEach-Object { $_.Name }) -join '、'

    $movers = @($Result.Quotes | Where-Object { $_.Ok -and $null -ne $_.ChangePct } | Sort-Object -Property @{Expression='ChangePct'; Descending=$true})
    $newCount = 0
    foreach ($g in @($Result.Groups)) { $newCount += @($g.Items).Count }
    $cats = @($Result.Groups | Select-Object -ExpandProperty Category -Unique)
    $allItems = @()
    foreach ($g in @($Result.Groups)) { foreach ($it in @($g.Items)) { if ($it.Date) { $allItems += $it } } }
    $allItems = $allItems | Sort-Object -Property Date -Descending
    $top3 = $allItems | Select-Object -First 3

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<div class="summary">')
    [void]$sb.AppendLine('<h2>📌 今日要点</h2>')
    [void]$sb.AppendLine('<ul>')

    if ($okCount -eq $totalCount) {
        [void]$sb.AppendLine("<li><strong>数据健康</strong>：本次 $totalCount 个来源全部抓取成功，可放心引用。</li>")
    } else {
        $f = ($totalCount - $okCount)
        [void]$sb.AppendLine("<li><strong>数据健康</strong>：本次 $totalCount 个来源中有 $f 个失败（$failedNames），相关内容可能缺失。</li>")
    }

    if ($movers.Count -gt 0) {
        $top = $movers[0]; $bot = $movers[-1]
        $topPct = if ($null -ne $top.ChangePct) { ('{0}{1}%' -f $(if ([double]$top.ChangePct -gt 0) { '+' } else { '' }), ([double]$top.ChangePct).ToString('N2')) } else { 'n/a' }
        $botPct = if ($null -ne $bot.ChangePct) { ('{0}{1}%' -f $(if ([double]$bot.ChangePct -gt 0) { '+' } else { '' }), ([double]$bot.ChangePct).ToString('N2')) } else { 'n/a' }
        $topPx = $top.Price.ToString('N2')
        $botPx = $bot.Price.ToString('N2')
        [void]$sb.AppendLine("<li><strong>行情异动</strong>：最大涨幅 $($top.Label) $($topPx)$($top.Suffix)（$topPct），最大跌幅 $($bot.Label) $($botPx)$($bot.Suffix)（$botPct）。</li>")
    }

    if ($newCount -gt 0) {
        [void]$sb.AppendLine("<li><strong>新增资讯</strong>：$newCount 条，覆盖 $($cats.Count) 个类别（$($cats -join '、')）。</li>")
    } else {
        [void]$sb.AppendLine('<li><strong>新增资讯</strong>：本次窗口内无新增条目，可能是新闻空窗期或来源抓取失败。</li>')
    }
    [void]$sb.AppendLine('</ul>')

    if ($top3.Count -gt 0) {
        [void]$sb.AppendLine('<div class="latest"><strong>最新 3 条</strong>：<ul>')
        foreach ($it in $top3) {
            $d = $it.Date.ToString('MM-dd HH:mm')
            $title = ($it.Title -replace '[\r\n]+', ' ').Trim()
            $titleEnc = [System.Net.WebUtility]::HtmlEncode($title)
            if ($it.Link) {
                [void]$sb.AppendLine("<li><code>[$d]</code> <a href=`"$($it.Link)`" target=`"_blank`">$titleEnc</a></li>")
            } else {
                [void]$sb.AppendLine("<li><code>[$d]</code> $titleEnc</li>")
            }
        }
        [void]$sb.AppendLine('</ul></div>')
    }
    [void]$sb.AppendLine('</div>')
    return $sb.ToString()
}

function ConvertTo-BriefingHtml {
    <# Wraps rendered HTML in a self-contained, readable document. #>
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][AllowEmptyString()][string]$BodyHtml,
        [string]$Banner = '',
        [string]$SummaryBox = ''
    )
    $css = @'
:root{color-scheme:light dark}
body{margin:0 auto;padding:28px 22px 64px;max-width:900px;
 font-family:"Microsoft YaHei","Segoe UI",-apple-system,"PingFang SC",sans-serif;
 line-height:1.75;color:#1b1b1b;background:#fff;font-size:16px}
h1{font-size:25px;margin:0 0 6px;padding-bottom:10px;border-bottom:3px solid #1a56c4}
h2{font-size:20px;margin:30px 0 10px;padding-left:10px;border-left:4px solid #1a56c4}
h3{font-size:17px;margin:20px 0 8px;color:#1a3f7a}
h4{font-size:15px;margin:14px 0 6px;color:#444}
p{margin:8px 0}
a{color:#1a56c4;text-decoration:none;word-break:break-all}
a:hover{text-decoration:underline}
ul{margin:8px 0;padding-left:22px}
li{margin:5px 0}
.cont{margin:2px 0 8px 4px;color:#555;font-size:14px;line-height:1.6}
table{border-collapse:collapse;width:100%;margin:12px 0;font-size:14px}
th,td{border:1px solid #d7dde5;padding:7px 9px;text-align:left}
th{background:#eef3fb;font-weight:600}
tr:nth-child(even) td{background:#fafbfd}
pre{background:#f5f7fa;border:1px solid #e0e5ec;border-radius:5px;padding:11px;
 overflow-x:auto;font-size:12.5px;line-height:1.5;font-family:Consolas,monospace}
code{background:#f0f2f5;padding:1px 5px;border-radius:3px;font-size:13px}
blockquote{margin:10px 0;padding:9px 13px;background:#f7f9fc;border-left:3px solid #9bb4d8;color:#444;font-size:14.5px}
.banner{padding:12px 15px;border-radius:6px;margin:14px 0;font-size:15px}
.banner.warn{background:#fff4e5;border:1px solid #ffb84d;color:#7a4a00}
.banner.info{background:#eaf4ff;border:1px solid #8fc0f0;color:#12406e}
.banner.ok{background:#f0f8f0;border:1px solid #7ab87a;color:#2d5a2d}
.summary{margin:30px 0 18px;padding:18px 20px;background:linear-gradient(135deg,#f5f8ff 0%,#eef4ff 100%);border:1px solid #b8cdf0;border-left:5px solid #1a56c4;border-radius:6px;font-size:15px;box-shadow:0 1px 3px rgba(26,86,196,0.08)}
.summary h2{margin:0 0 10px;padding:0 0 8px;border-bottom:1px solid #cfdcef;border-left:none;font-size:18px;color:#1a3f7a}
.summary ul{margin:6px 0;padding-left:20px}
.summary li{margin:4px 0}
.summary .latest{margin-top:10px;padding-top:10px;border-top:1px dashed #cfdcef}
.foot{margin-top:40px;padding-top:14px;border-top:1px solid #e2e6ea;color:#888;font-size:12.5px}
.section{margin:32px 0;padding:22px 26px 18px;border-radius:10px;border:1px solid;border-left:6px solid;box-shadow:0 1px 4px rgba(0,0,0,0.04)}
.section h1{margin-top:0;font-size:22px;padding-bottom:8px}
.section h2{font-size:18px;padding-left:8px;border-left-width:3px;margin-top:24px}
.section-property{background:linear-gradient(135deg,#fff7eb 0%,#fff2e0 100%);border-color:#f0c590;color:#7a4a00}
.section-property h1{color:#b35900;border-bottom-color:#f0c590}
.section-property h2{color:#b35900;border-left-color:#f0c590}
.section-property h3{color:#a04800}
.section-property table{border-color:#f0c590}
.section-property th{background:#fff0d8}
.section-property tr:nth-child(even) td{background:#fff8ec}
.section-property blockquote{background:#fff5e2;border-left-color:#e0a060;color:#7a4a00}
.section-property code{background:#fff0d8;color:#7a4a00}
.section-finance{background:linear-gradient(135deg,#f0f5ff 0%,#e6efff 100%);border-color:#b8cdf0;color:#1a3f7a}
.section-finance h1{color:#1a56c4;border-bottom-color:#b8cdf0}
.section-finance h2{color:#1a56c4;border-left-color:#b8cdf0}
.section-finance h3{color:#1a3f7a}
.section-finance table{border-color:#b8cdf0}
.section-finance th{background:#e6efff}
.section-finance tr:nth-child(even) td{background:#f5f8ff}
.section-finance blockquote{background:#eaf2ff;border-left-color:#7ba0d8;color:#1a3f7a}
.section-finance code{background:#e6efff;color:#1a3f7a}
@media(prefers-color-scheme:dark){
 body{background:#16181c;color:#dfe3e8}
 h3{color:#8fb6f0} h4{color:#aaa}
 th{background:#22262e} td,th{border-color:#333941}
 tr:nth-child(even) td{background:#1b1e24}
 pre{background:#1c1f25;border-color:#333941}
 code{background:#242830}
 blockquote{background:#1c1f25;color:#b8bec6}
 .cont{color:#a8aeb6}
.banner.warn{background:#3a2c14;border-color:#7a5a20;color:#f0c380}
  .banner.info{background:#152436;border-color:#2d5480;color:#9cc4ee}
  .banner.ok{background:#152a14;border-color:#3d6a30;color:#9bd49b}
  .summary{background:linear-gradient(135deg,#1c2330 0%,#1a2030 100%);border-color:#2d4a7a;box-shadow:0 1px 3px rgba(100,150,255,0.1)}
  .summary h2{color:#8fb6f0;border-bottom-color:#2d4a7a}
  .summary .latest{border-top-color:#2d4a7a}
  .section-property{background:linear-gradient(135deg,#2a1f10 0%,#1f1808 100%);border-color:#5a4020;color:#f0c380}
  .section-property h1{color:#f0c380;border-bottom-color:#5a4020}
  .section-property h2{color:#f0c380;border-left-color:#5a4020}
  .section-property h3{color:#e0a060}
  .section-property table{border-color:#4a3520}
  .section-property th{background:#3a2810;color:#f0c380}
  .section-property tr:nth-child(even) td{background:#241808}
  .section-property blockquote{background:#1f1808;color:#e0a060;border-left-color:#7a5020}
  .section-property code{background:#3a2810;color:#f0c380}
  .section-finance{background:linear-gradient(135deg,#152436 0%,#0f1828 100%);border-color:#2d5480;color:#9cc4ee}
  .section-finance h1{color:#8fb6f0;border-bottom-color:#2d5480}
  .section-finance h2{color:#8fb6f0;border-left-color:#2d5480}
  .section-finance h3{color:#7ba0d8}
  .section-finance table{border-color:#2d4060}
  .section-finance th{background:#1c2a40;color:#8fb6f0}
  .section-finance tr:nth-child(even) td{background:#152030}
  .section-finance blockquote{background:#152030;color:#b0c8e0;border-left-color:#3d6090}
  .section-finance code{background:#1c2a40;color:#9cc4ee}}
'@
    $head = @"
<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>$([System.Net.WebUtility]::HtmlEncode($Title))</title><style>$css</style></head><body>
"@
    $foot = '<div class="foot">由每日资讯采集器自动生成，不构成投资建议。</div></body></html>'
    return $head + $Banner + $BodyHtml + $SummaryBox + $foot
}

function Test-Relevant {
    <# Keyword gate for broad feeds (e.g. all-of-government press releases). #>
    param([string]$Text, [string]$Include)
    if ([string]::IsNullOrWhiteSpace($Include)) { return $true }
    return ($Text -match $Include)
}

Export-ModuleMember -Function Write-Utf8, Read-BytesAsText, Invoke-Fetch, ConvertTo-LocalDate,
    Test-InWindow, Get-NodeText, Format-Link, Get-ItemLink, ConvertFrom-Feed, Resolve-Url,
    Get-AnchorItems, Get-CsvTail, Get-MarketQuote, Test-Relevant, Send-Toast, Limit-LogFile,
    ConvertFrom-MarkdownLite, ConvertTo-BriefingHtml, ConvertTo-SummaryBox, Publish-GitHubPages
