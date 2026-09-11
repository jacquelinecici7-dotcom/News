<#
.SYNOPSIS
    Daily briefing collector and publisher.

.DESCRIPTION
    Runs on a schedule and puts a readable briefing in front of the user every day.

    Pipeline:
      collect -> pack.md (source material + writing prompt)
              -> briefing.md   (model-written, only when ANTHROPIC_API_KEY is set)
              -> 今日资讯.html  (always produced; delivered to the Desktop + toast)

    The delivered document never depends on the model being available: without a key
    it contains the mechanically assembled digest, clearly labelled as raw material.
    With a key the model-written briefing goes on top and the digest becomes an
    appendix, so the sources behind every claim stay visible.

.EXAMPLE
    .\Invoke-Briefing.ps1
.EXAMPLE
    .\Invoke-Briefing.ps1 -Brief finance -Force
#>
[CmdletBinding()]
param(
    [ValidateSet('finance', 'hkproperty', 'all')]
    [string]$Brief = 'all',

    [string]$OutputRoot,
    [int]$FreshnessHours = 0,
    [switch]$NoSynthesis,
    [string]$Model = 'claude-opus-5',

    # Re-run even if today's briefing was already delivered.
    [switch]$Force,

    # Collect only; skip Desktop copy / toast / auto-open.
    [switch]$NoDeliver,

    # Also produce a combined single HTML containing both briefings as independent sections.
    [switch]$Merge,

    # Run inside GitHub Actions: skip Publish-GitHubPages (workflow commits),
    # skip desktop copy, write outputs to repo-root paths the workflow expects.
    [switch]$GitHubActions,

    [switch]$Open
)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
if ($GitHubActions) {
    # When this script lives at briefing/Invoke-Briefing.ps1 inside the repo, the
    # repo root is one level up. State, output, and logs all live at repo root so
    # the workflow can commit them without dragging the script source along.
    $repoRoot = (Resolve-Path (Join-Path $root '..')).Path
    $OutputRoot = Join-Path $repoRoot 'out'
    $NoDeliver = $true
}
Import-Module (Join-Path $root 'lib\Collector.psm1') -Force

$LogPath = if ($GitHubActions) { Join-Path $repoRoot 'logs\run.log' } else { Join-Path $root 'logs\run.log' }

# ---------------------------------------------------------------- helpers ----

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line
    $dir = Split-Path -Parent $LogPath
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
}

function Format-Number {
    param([double]$Value)
    if ([Math]::Abs($Value) -ge 100) { return $Value.ToString('N2') }
    if ([Math]::Abs($Value) -ge 1)   { return $Value.ToString('N3') }
    return $Value.ToString('N4')
}

function Format-Pct {
    param($Value)
    if ($null -eq $Value) { return 'n/a' }
    $sign = if ([double]$Value -gt 0) { '+' } else { '' }
    return ('{0}{1}%' -f $sign, ([double]$Value).ToString('N2'))
}

function Format-Cell {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    return ((($Text -replace '[\r\n]+', ' ') -replace '\|', '/').Trim())
}

function Read-Json {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    try { return ([System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8) | ConvertFrom-Json) } catch { return $null }
}

# -- dedupe state -------------------------------------------------------------

function Get-SeenMap {
    param([string]$Path)
    $map = @{}
    $arr = @(Read-Json -Path $Path)
    $cutoff = (Get-Date).AddDays(-30)
    foreach ($e in $arr) {
        if (-not $e -or -not $e.link) { continue }
        $d = ConvertTo-LocalDate -Text ([string]$e.firstSeen)
        if ($null -ne $d -and $d -lt $cutoff) { continue }   # prune
        $map[[string]$e.link] = [string]$e.firstSeen
    }
    return $map
}

function Save-SeenMap {
    param([string]$Path, [hashtable]$Map)
    $arr = foreach ($k in $Map.Keys) { [pscustomobject]@{ link = $k; firstSeen = $Map[$k] } }
    Write-Utf8 -Path $Path -Content ((@($arr) | ConvertTo-Json -Depth 4))
}

function Test-AlreadySeen {
    <#
        An item counts as "seen" only if it was first collected on an EARLIER day.
        Items first seen today stay eligible, so re-running the same day rebuilds a
        full pack instead of progressively thinning it out.
    #>
    param([hashtable]$Seen, [string]$Link, [datetime]$Today)
    if ([string]::IsNullOrWhiteSpace($Link)) { return $false }
    if (-not $Seen.ContainsKey($Link)) { return $false }
    $d = ConvertTo-LocalDate -Text ([string]$Seen[$Link])
    if ($null -ne $d -and $d.Date -eq $Today.Date) { return $false }
    return $true
}

# -- run metadata (adaptive window, failure streaks, delivery guard) -----------

function Get-Meta {
    param([string]$Path)
    $m = [pscustomobject]@{ lastSuccessAt = ''; lastDeliveredDate = ''; failStreak = @() }
    $j = Read-Json -Path $Path
    if ($j) {
        if ($j.lastSuccessAt)     { $m.lastSuccessAt = [string]$j.lastSuccessAt }
        if ($j.lastDeliveredDate) { $m.lastDeliveredDate = [string]$j.lastDeliveredDate }
        if ($j.failStreak)        { $m.failStreak = @($j.failStreak) }
    }
    return $m
}

function Get-Streak {
    param($Meta, [string]$Name)
    $e = @($Meta.failStreak | Where-Object { $_ -and $_.name -eq $Name })
    if ($e.Count -gt 0) { return [int]$e[0].count }
    return 0
}

# --------------------------------------------------------------- collect ----

function Invoke-Collect {
    param($Config, [datetime]$RunAt, $Meta)

    # Adaptive window: normally the configured span, but if the last successful run
    # was longer ago (weekend, holiday, machine off) stretch it to cover the gap so
    # nothing falls through unseen. Capped so a long outage can't pull in weeks.
    $baseHours = if ($null -ne $Config.freshnessHours) { [int]$Config.freshnessHours } else { 30 }
    $hours = if ($FreshnessHours -gt 0) { $FreshnessHours } else { $baseHours }
    $windowNote = ''
    if ($FreshnessHours -le 0 -and $Meta.lastSuccessAt) {
        $last = ConvertTo-LocalDate -Text $Meta.lastSuccessAt
        if ($null -ne $last) {
            $gap = [int][Math]::Ceiling(($RunAt - $last).TotalHours) + 2
            if ($gap -gt $hours) {
                $hours = [Math]::Min($gap, 96)
                $windowNote = ('（距上次成功采集 {0} 小时，已自动延长窗口）' -f [int]($RunAt - $last).TotalHours)
            }
        }
    }

    $cutoff = $RunAt.AddHours(-$hours)
    $cap = [int]$Config.maxItemsPerSource
    $status = New-Object System.Collections.ArrayList
    $groups = New-Object System.Collections.ArrayList
    $datasets = New-Object System.Collections.ArrayList

    $statePath = Join-Path $root ('state\seen.{0}.json' -f $Config.id)
    $seen = Get-SeenMap -Path $statePath
    $filteredOld = 0

    # -- quotes ------------------------------------------------------------
    $quotes = New-Object System.Collections.ArrayList
    foreach ($q in @($Config.quotes)) {
        $r = Get-MarketQuote -Symbol $q.symbol -Label $q.label -Group $q.group -Suffix $q.suffix
        [void]$quotes.Add($r)
        if (-not $r.Ok) { Write-Log ("行情失败 {0} ({1}): {2}" -f $q.label, $q.symbol, $r.Error) 'WARN' }
    }
    $qOk = @($quotes | Where-Object { $_.Ok }).Count
    [void]$status.Add([pscustomobject]@{
        Name = 'Yahoo Finance 行情'; Kind = '行情'
        State = $(if ($qOk -gt 0) { 'OK' } else { '失败' }); Count = $qOk
        Note = "$qOk/$($quotes.Count) 个标的取得报价"
    })

    # Shared item pipeline: window filter and cap are applied BEFORE the dedupe
    # check, so a busy day can't defer items into tomorrow (where they would fall
    # out of the window and be lost). The cap is a deterministic "top N of window".
    $sift = {
        param($items, $include)
        $inWindow = New-Object System.Collections.ArrayList
        foreach ($it in $items) {
            if (-not (Test-Relevant -Text $it.Title -Include $include)) { continue }
            if (-not (Test-InWindow -Date $it.Date -Cutoff $cutoff)) { continue }
            [void]$inWindow.Add($it)
            if ($inWindow.Count -ge $cap) { break }
        }
        return $inWindow.ToArray()
    }

    # -- feeds -------------------------------------------------------------
    foreach ($f in @($Config.feeds)) {
        $note = ''; $state = 'OK'; $kept = @(); $win = 0
        try {
            $resp = Invoke-Fetch -Url $f.url -TimeoutSec 25
            if (-not $resp.Ok) { throw $resp.Error }
            $windowed = & $sift (ConvertFrom-Feed -Xml $resp.Text) $f.include
            $win = $windowed.Count
            $kept = @($windowed | Where-Object { -not (Test-AlreadySeen -Seen $seen -Link $_.Link -Today $RunAt) })
            $filteredOld += ($win - $kept.Count)
            if ($kept.Count -eq 0) { $note = if ($win -gt 0) { "窗口内 $win 条，均为此前已采集" } else { '本次窗口内无新增相关条目' } }
        } catch {
            $state = '失败'; $note = $_.Exception.Message
            Write-Log ("来源失败 {0}: {1}" -f $f.name, $note) 'WARN'
        }
        [void]$status.Add([pscustomobject]@{ Name = $f.name; Kind = 'RSS'; State = $state; Count = $kept.Count; Note = $note })
        if ($kept.Count -gt 0) { [void]$groups.Add([pscustomobject]@{ Category = $f.category; Source = $f.name; Url = $f.url; Items = $kept }) }
    }

    # -- listing pages -----------------------------------------------------
    foreach ($p in @($Config.pages)) {
        $note = ''; $state = 'OK'; $kept = @(); $win = 0
        try {
            $resp = Invoke-Fetch -Url $p.url -TimeoutSec 30
            if (-not $resp.Ok) { throw $resp.Error }
            $max = if ($p.max) { [int]$p.max } else { 15 }
            $items = Get-AnchorItems -Html $resp.Text -BaseUrl $p.url -HrefFilter $p.hrefFilter -DateFromHref $p.dateFromHref -Max $max
            $windowed = & $sift $items $p.include
            $win = $windowed.Count
            $kept = @($windowed | Where-Object { -not (Test-AlreadySeen -Seen $seen -Link $_.Link -Today $RunAt) })
            $filteredOld += ($win - $kept.Count)
            if ($kept.Count -eq 0) { $note = if ($win -gt 0) { "窗口内 $win 条，均为此前已采集" } else { '本次窗口内无新增相关条目' } }
        } catch {
            $state = '失败'; $note = $_.Exception.Message
            Write-Log ("来源失败 {0}: {1}" -f $p.name, $note) 'WARN'
        }
        [void]$status.Add([pscustomobject]@{ Name = $p.name; Kind = '页面'; State = $state; Count = $kept.Count; Note = $note })
        if ($kept.Count -gt 0) { [void]$groups.Add([pscustomobject]@{ Category = $p.category; Source = $p.name; Url = $p.url; Items = $kept }) }
    }

    # -- official datasets (CSV) -------------------------------------------
    foreach ($d in @($Config.datasets)) {
        $note = ''; $state = 'OK'; $n = 0
        try {
            $resp = Invoke-Fetch -Url $d.url -TimeoutSec 30
            if (-not $resp.Ok) { throw $resp.Error }
            $csv = Get-CsvTail -Text $resp.Text -Rows $(if ($d.rows) { [int]$d.rows } else { 6 }) `
                               -HeaderLine $(if ($null -ne $d.headerLine) { [int]$d.headerLine } else { 1 })
            $n = $csv.Rows.Count
            [void]$datasets.Add([pscustomobject]@{
                Name = $d.name; Category = $d.category; Url = $d.url; Note = $d.note
                Title = $csv.Title; Header = $csv.Header; Rows = $csv.Rows
            })
        } catch {
            $state = '失败'; $note = $_.Exception.Message
            Write-Log ("数据集失败 {0}: {1}" -f $d.name, $note) 'WARN'
        }
        [void]$status.Add([pscustomobject]@{ Name = $d.name; Kind = '数据集'; State = $state; Count = $n; Note = $note })
    }

    # -- persist dedupe state ---------------------------------------------
    $stamp = $RunAt.ToString('yyyy-MM-dd HH:mm:ss')
    foreach ($g in $groups) { foreach ($it in $g.Items) { if ($it.Link -and -not $seen.ContainsKey($it.Link)) { $seen[$it.Link] = $stamp } } }
    Save-SeenMap -Path $statePath -Map $seen

    return [pscustomobject]@{
        Config = $Config; RunAt = $RunAt; Hours = $hours; Cutoff = $cutoff; WindowNote = $windowNote
        Quotes = $quotes.ToArray(); Status = $status.ToArray()
        Groups = $groups.ToArray(); Datasets = $datasets.ToArray(); FilteredOld = $filteredOld
    }
}

# ---------------------------------------------------------------- render ----

function Format-CompactDigest {
    <#
        Compact, news-like digest used for the delivered HTML when no model is
        available (or as the appendix under a model briefing). Groups items by
        category as one-line bullets with title + link + time; status is a small
        footer instead of a full table.
    #>
    param($Result)

    $cfg = $Result.Config
    $sb = New-Object System.Text.StringBuilder
    function add([string]$s = '') { [void]$sb.AppendLine($s) }

    add ("# {0} · {1}" -f $cfg.title, $Result.RunAt.ToString('yyyy年M月d日 HH:mm'))
    add ''
    add ("截至北京时间 {0}，" -f $Result.RunAt.ToString('yyyy年M月d日 HH:mm'))

    # One-line preview: count per category, and any chronic failures surfaced inline.
    $cats = @($Result.Groups | Select-Object -ExpandProperty Category -Unique)
    $catCounts = @{}
    foreach ($c in $cats) { $catCounts[$c] = (@($Result.Groups | Where-Object { $_.Category -eq $c } | ForEach-Object { $_.Items.Count } | Measure-Object -Sum).Sum) }
    $failed = @($Result.Status | Where-Object { $_.State -ne 'OK' })
    $failedNames = ($failed | ForEach-Object { $_.Name }) -join '、'
    if ($failedNames) {
        add ("本次采集覆盖 {0} 类共 {1} 条新增；{2} 抓取失败（{3}），相关内容可能缺失。" -f $cats.Count, ($catCounts.Values | Measure-Object -Sum).Sum, $failed.Count, $failedNames)
    } else {
        add ("本次采集覆盖 {0} 类共 {1} 条新增，来源全部正常。" -f $cats.Count, ($catCounts.Values | Measure-Object -Sum).Sum)
    }
    add ''

    # Compact market snapshot — single line per group, no full table.
    $okQuotes = @($Result.Quotes | Where-Object { $_.Ok })
    if ($okQuotes.Count -gt 0) {
        add '## 行情快照'
        add ''
        foreach ($grp in ($okQuotes | Select-Object -ExpandProperty Group -Unique)) {
            $bits = foreach ($q in ($okQuotes | Where-Object { $_.Group -eq $grp })) {
                $arrow = if ($null -ne $q.ChangePct -and [double]$q.ChangePct -gt 0) { '↑' }
                         elseif ($null -ne $q.ChangePct -and [double]$q.ChangePct -lt 0) { '↓' } else { '·' }
                '{0} {1}{2} {3}{4}' -f $q.Label, (Format-Number -Value $q.Price), $q.Suffix, $arrow, (Format-Pct -Value $q.ChangePct)
            }
            add ('- **{0}**：{1}' -f $grp, ($bits -join ' ｜ '))
        }
        add ''
    }

    # News items, grouped by category. Capped to newest N per category so a busy
    # day doesn't drown the reader; total count is still surfaced in the summary.
    add '## 新增资讯'
    add ''
    if ($Result.Groups.Count -eq 0) {
        add '_本次窗口内无新增条目。_'
        add ''
    } else {
        $maxPerCat = 10
        foreach ($cat in $cats) {
            $items = @($Result.Groups | Where-Object { $_.Category -eq $cat } | ForEach-Object { $_.Items })
            if ($items.Count -eq 0) { continue }
            # Sort newest first so the cap keeps the freshest, not the oldest.
            $sorted = @($items | Sort-Object -Property Date -Descending)
            $shown = $sorted | Select-Object -First $maxPerCat
            $total = $items.Count
            add ("### {0}（{1} 条 · 显示最新 {2} 条）" -f $cat, $total, $shown.Count)
            add ''
            foreach ($it in $shown) {
                $d = if ($null -ne $it.Date) { $it.Date.ToString('MM-dd HH:mm') } else { '日期未知' }
                $title = ($it.Title -replace '[\r\n]+', ' ').Trim()
                $line = if ($it.Link) { '- [{0}]({1}) · {2}' -f $title, $it.Link, $d }
                        else        { '- {0} · {1}' -f $title, $d }
                add $line
            }
            if ($total -gt $shown.Count) {
                add ('<small>_另有 {0} 条同类未列出，详见 raw.json。_</small>' -f ($total - $shown.Count))
            }
            add ''
        }
    }

    # Small status footer — full table is noise for a daily read.
    add '## 抓取状态'
    add ''
    add ('| 来源 | 结果 | 新增 |')
    add ('|---|---|---|')
    foreach ($s in $Result.Status) {
        $icon = if ($s.State -eq 'OK') { '✓' } else { '✗' }
        add ('| {0} | {1} | {2} |' -f (Format-Cell $s.Name), $icon, $s.Count)
    }
    add ''
    # 今日要点 — 自动汇总，不依赖模型。放在最末，让读者一眼收尾。
    add '## 今日要点'
    add ''

    # 1. 数据健康：来源全部 OK 还是有问题
    $okCount = @($Result.Status | Where-Object { $_.State -eq 'OK' }).Count
    $totalCount = @($Result.Status).Count
    if ($okCount -eq $totalCount) {
        add '- **数据健康**：本次 {0} 个来源全部抓取成功，可放心引用。' -f $totalCount
    } else {
        add ('- **数据健康**：本次 {0} 个来源中有 {1} 个失败（{2}），相关内容可能缺失。' -f $totalCount, ($totalCount - $okCount), (($Result.Status | Where-Object { $_.State -ne 'OK' } | ForEach-Object { $_.Name }) -join '、'))
    }

    # 2. 行情最大变动：哪个标的波动最大
    $movers = @($Result.Quotes | Where-Object { $_.Ok -and $null -ne $_.ChangePct } | Sort-Object -Property @{Expression='ChangePct'; Descending=$true})
    if ($movers.Count -gt 0) {
        $top = $movers[0]; $bot = $movers[-1]
        add ('- **行情异动**：今日最大涨幅 {0} {1}{2}（{3}），最大跌幅 {4} {5}{6}（{7}）。' -f `
            $top.Label, (Format-Number -Value $top.Price), $top.Suffix, (Format-Pct -Value $top.ChangePct), `
            $bot.Label, (Format-Number -Value $bot.Price), $bot.Suffix, (Format-Pct -Value $bot.ChangePct))
    }

    # 3. 资讯条数
    $newCount = 0
    foreach ($g in @($Result.Groups)) { $newCount += @($g.Items).Count }
    if ($newCount -gt 0) {
        $cats = @($Result.Groups | Select-Object -ExpandProperty Category -Unique)
        add ('- **新增资讯**：本次采集到 {0} 条新增，覆盖 {1} 个类别。' -f $newCount, $cats.Count)
    } else {
        add '- **新增资讯**：本次窗口内无新增条目，可能是新闻空窗期或来源抓取失败。'
    }

    # 4. 头部 3 条最新消息（按时间倒序）
    $allItems = @()
    foreach ($g in @($Result.Groups)) { foreach ($it in @($g.Items)) { if ($it.Date) { $allItems += $it } } }
    $allItems = $allItems | Sort-Object -Property Date -Descending
    if ($allItems.Count -gt 0) {
        add ''
        add '**最新 3 条**：'
        add ''
        $top3 = $allItems | Select-Object -First 3
        foreach ($it in $top3) {
            $d = $it.Date.ToString('MM-dd HH:mm')
            $title = ($it.Title -replace '[\r\n]+', ' ').Trim()
            if ($it.Link) { add ('- `[{0}]` [{1}]({2})' -f $d, $title, $it.Link) }
            else        { add ('- `[{0}]` {1}' -f $d, $title) }
        }
    }
    add ''
    # 结尾免责声明由 ConvertTo-BriefingHtml 的 .foot 渲染，避免重复
    return $sb.ToString()
}

function Format-Pack {
    <# -Human drops the writing prompt and titles the document for reading. #>
    param($Result, [AllowEmptyString()][string]$Prompt, [switch]$Human)

    $cfg = $Result.Config
    $sb = New-Object System.Text.StringBuilder
    function add([string]$s = '') { [void]$sb.AppendLine($s) }

    if ($Human) {
        add ("# {0} · {1}" -f $cfg.title, $Result.RunAt.ToString('yyyy年M月d日'))
    } else {
        add ("# {0} · 采集包" -f $cfg.title)
    }
    add ''
    add ("- 采集时间：{0}（{1}）" -f $Result.RunAt.ToString('yyyy年M月d日 HH:mm:ss'), [System.TimeZoneInfo]::Local.Id)
    add ("- 覆盖窗口：过去 {0} 小时（{1} 起）{2}" -f $Result.Hours, $Result.Cutoff.ToString('yyyy-MM-dd HH:mm'), $Result.WindowNote)
    add ("- 已过滤往日已采集条目：{0} 条" -f $Result.FilteredOld)
    add ''

    add '## 一、实时行情快照'
    add ''
    add '> 来源：Yahoo Finance。以下为**采集时点的快照**，不是收盘价；期货为连续合约。'
    add ''
    $okQuotes = @($Result.Quotes | Where-Object { $_.Ok })
    if ($okQuotes.Count -eq 0) {
        add '**本次未能取得任何实时行情，请勿引用任何具体价格。**'
        add ''
    } else {
        foreach ($grp in ($okQuotes | Select-Object -ExpandProperty Group -Unique)) {
            add ("### {0}" -f $grp)
            add ''
            add '| 指标 | 最新 | 前收 | 涨跌 | 涨跌幅 | 行情时间 |'
            add '|---|---|---|---|---|---|'
            foreach ($q in ($okQuotes | Where-Object { $_.Group -eq $grp })) {
                $prev = 'n/a'; $chg = 'n/a'
                if ($null -ne $q.PrevClose) {
                    $prev = (Format-Number -Value $q.PrevClose) + $q.Suffix
                    $delta = $q.Price - $q.PrevClose
                    $chg = $(if ($delta -gt 0) { '+' } else { '' }) + (Format-Number -Value $delta)
                }
                $asOf = if ($null -ne $q.AsOf) { $q.AsOf.ToString('MM-dd HH:mm') } else { 'n/a' }
                add ('| {0} | {1}{2} | {3} | {4} | {5} | {6} |' -f $q.Label, (Format-Number -Value $q.Price), $q.Suffix, $prev, $chg, (Format-Pct -Value $q.ChangePct), $asOf)
            }
            add ''
        }
        if (@($okQuotes | Where-Object { $_.Suffix -eq '%' }).Count -gt 0) {
            add '> 收益率类指标的单位为百分点：「涨跌 +0.041」表示收益率上升约 4.1 个基点，而非上涨 0.041%。'
            add ''
        }
    }
    $badQuotes = @($Result.Quotes | Where-Object { -not $_.Ok })
    if ($badQuotes.Count -gt 0) {
        add ('> 未能取得报价：' + (($badQuotes | ForEach-Object { '{0}（{1}）' -f $_.Label, $_.Error }) -join '；'))
        add ''
    }

    add '## 二、抓取状态'
    add ''
    add '| 来源 | 类型 | 状态 | 新增条目 | 说明 |'
    add '|---|---|---|---|---|'
    foreach ($s in $Result.Status) {
        add ('| {0} | {1} | {2} | {3} | {4} |' -f (Format-Cell $s.Name), $s.Kind, $s.State, $s.Count, (Format-Cell $s.Note))
    }
    add ''
    $failed = @($Result.Status | Where-Object { $_.State -ne 'OK' })
    if ($failed.Count -gt 0) {
        if ($Human) {
            add ('> 本次有 {0} 个来源抓取失败：{1}。这些领域的内容可能缺失。' -f $failed.Count, (($failed | ForEach-Object { $_.Name }) -join '、'))
        } else {
            add ('> **本次有 {0} 个来源抓取失败：{1}。** 相关领域如因此缺料，必须在简报中明确写出「本次未能获取×××」，不要用记忆或推测补齐。' -f $failed.Count, (($failed | ForEach-Object { $_.Name }) -join '、'))
        }
    } else {
        add '> 全部来源抓取成功。'
    }
    add ''

    if ($Result.Datasets.Count -gt 0) {
        add '## 三、官方数据集（最新数据行）'
        add ''
        foreach ($d in $Result.Datasets) {
            add ("### {0}" -f $d.Name)
            add ''
            add ("- 来源：{0}" -f $d.Url)
            if ($d.Note) { add ("- 说明：{0}" -f $d.Note) }
            add ''
            add '```'
            add $d.Header
            foreach ($r in $d.Rows) { add $r }
            add '```'
            add ''
        }
    }

    $n4 = if ($Result.Datasets.Count -gt 0) { '四' } else { '三' }
    add ("## {0}、新增资讯条目" -f $n4)
    add ''
    if ($Result.Groups.Count -eq 0) {
        add '**本次窗口内没有采集到任何新增条目。**'
        add ''
    } else {
        foreach ($cat in ($Result.Groups | Select-Object -ExpandProperty Category -Unique)) {
            add ("### {0}" -f $cat)
            add ''
            foreach ($g in ($Result.Groups | Where-Object { $_.Category -eq $cat })) {
                add ("**{0}**" -f $g.Source)
                add ''
                foreach ($it in $g.Items) {
                    $d = if ($null -ne $it.Date) { $it.Date.ToString('MM-dd HH:mm') } else { '日期未知' }
                    add ('- `[{0}]` {1}' -f $d, $it.Title)
                    if ($it.Summary) { add ('  {0}' -f $it.Summary) }
                    if ($it.Link) { add ('  {0}' -f $it.Link) }
                }
                add ''
            }
        }
    }

    if (-not $Human -and $Prompt) {
        $n5 = if ($Result.Datasets.Count -gt 0) { '五' } else { '四' }
        add ("## {0}、写稿指令" -f $n5)
        add ''
        add '> 将本文件整体交给模型即可成稿。'
        add ''
        add $Prompt
    }
    return $sb.ToString()
}

# ------------------------------------------------------------- synthesis ----

function Invoke-Synthesis {
    param([string]$Pack, [string]$ModelName, [string]$ApiKey)
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $body = @{ model = $ModelName; max_tokens = 16000; messages = @(@{ role = 'user'; content = $Pack }) }
    $json = $body | ConvertTo-Json -Depth 8 -Compress
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    $headers = @{ 'x-api-key' = $ApiKey; 'anthropic-version' = '2023-06-01' }
    $resp = Invoke-RestMethod -Uri 'https://api.anthropic.com/v1/messages' -Method Post `
        -Headers $headers -ContentType 'application/json' -Body $bytes -TimeoutSec 300
    return (@($resp.content | Where-Object { $_.type -eq 'text' } | ForEach-Object { $_.text }) -join "`n")
}

# --------------------------------------------------------------- deliver ----

function Save-Archive {
    <# Never overwrite an existing artefact: keep the previous copy alongside. #>
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return }
    $dir = Split-Path -Parent $Path
    $base = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    $ext = [System.IO.Path]::GetExtension($Path)
    # 用当前时间 + 毫秒，而不是被归档文件的 LastWriteTime；后者在 -Force 快速重跑时
    # 可能落到同一秒，导致后写的归档覆盖先写的。
    $stamp = Get-Date -Format 'HHmmssfff'
    Move-Item -LiteralPath $Path -Destination (Join-Path $dir ("{0}.{1}{2}" -f $base, $stamp, $ext)) -Force
}

function Publish-Briefing {
    param($Result, [AllowEmptyString()][string]$BriefingText, [string]$OutDir, $Delivery, $Meta)

    # Build the compact digest here, once. Used as the entire body when there's no
    # model briefing, or as the appendix under one.
    $Digest = Format-CompactDigest -Result $Result

    $cfg = $Result.Config
    $failed = @($Result.Status | Where-Object { $_.State -ne 'OK' })
    $desktopCopyOk = $false

    # Long-running breakage is worth shouting about: a source that has failed for
    # several days running needs a human, not another quiet note in a table.
    $chronic = @()
    foreach ($s in $Result.Status) {
        if ($s.State -eq 'OK') { continue }
        $n = (Get-Streak -Meta $Meta -Name $s.Name) + 1
        if ($n -ge 3) { $chronic += ('{0}（连续 {1} 次）' -f $s.Name, $n) }
    }

    $banner = ''
    if ($chronic.Count -gt 0) {
        $banner += '<div class="banner warn"><b>需要检修：</b>' + [System.Net.WebUtility]::HtmlEncode(($chronic -join '、')) + ' 已连续多次抓取失败，可能是网站改版。</div>'
    } elseif ($failed.Count -gt 0) {
        $banner += '<div class="banner warn"><b>注意：</b>本次有 ' + $failed.Count + ' 个来源抓取失败（' + [System.Net.WebUtility]::HtmlEncode((($failed | ForEach-Object { $_.Name }) -join '、')) + '），相关内容可能缺失。</div>'
    }
    if (-not $BriefingText) {
        $banner += '<div class="banner info">当前为<b>自动汇总的原始素材</b>，未经模型撰写成稿。设置环境变量 <code>ANTHROPIC_API_KEY</code> 后，这里会自动换成按模板写好的简报。</div>'
    }

    $md = if ($BriefingText) {
        $BriefingText + "`r`n`r`n---`r`n`r`n# 附：素材与来源`r`n`r`n" + $Digest
    } else { $Digest }

    $html = ConvertTo-BriefingHtml -Title ("{0} · {1}" -f $cfg.title, $Result.RunAt.ToString('yyyy年M月d日')) `
                                   -BodyHtml (ConvertFrom-MarkdownLite -Markdown $md) -Banner $banner `
                                   -SummaryBox (ConvertTo-SummaryBox -Result $Result)

    $htmlPath = Join-Path $OutDir '今日资讯.html'
    Save-Archive -Path $htmlPath
    Write-Utf8 -Path $htmlPath -Content $html

    $delivered = $htmlPath
    if (-not $NoDeliver -and $Delivery.desktopCopy) {
        try {
            $desk = [Environment]::GetFolderPath('Desktop')
            # Windows 文件名禁用字符做一次替换，避免配置里出现 < > : / \ | ? * 时 Copy-Item 抛错
            $safeTitle = ($cfg.title -replace '[<>:"/\\|?*\x00-\x1F]', '_')
            $target = Join-Path $desk ('{0}.html' -f $safeTitle)
            Copy-Item -LiteralPath $htmlPath -Destination $target -Force
            $delivered = $target
            $desktopCopyOk = $true
            Write-Log ("已送达桌面：{0}" -f $target)
        } catch { Write-Log ("桌面送达失败：{0}" -f $_.Exception.Message) 'WARN' }
    }

    if (-not $NoDeliver -and $Delivery.autoOpen) {
        try { Start-Process $delivered | Out-Null } catch { Write-Log ("自动打开失败：{0}" -f $_.Exception.Message) 'WARN' }
    }

    if (-not $NoDeliver -and $Delivery.toast) {
        $newCount = 0
        foreach ($g in @($Result.Groups)) { $newCount += @($g.Items).Count }
        $msg = if ($chronic.Count -gt 0) { "新增 $newCount 条，有来源需要检修" }
               elseif ($failed.Count -gt 0) { "新增 $newCount 条，$($failed.Count) 个来源失败" }
               else { "新增 $newCount 条，来源全部正常" }
        [void](Send-Toast -Title $cfg.title -Message $msg)
    }
    $deliveredOk = $NoDeliver -or (-not $Delivery.desktopCopy) -or $desktopCopyOk
    return [pscustomobject]@{ Path = $delivered; Ok = $deliveredOk }
}

function Publish-MergedBriefing {
    <#
        Combines multiple brief results into a single HTML with each brief as
        its own section. Aggregates banners, summary box, and meta-state from
        all briefs.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()]$Briefs,
        [Parameter(Mandatory)][string]$OutDir,
        $Delivery
    )
    if (-not $Briefs -or $Briefs.Count -eq 0) { return $null }

    $combinedQuotes = New-Object System.Collections.ArrayList
    $combinedStatus = New-Object System.Collections.ArrayList
    $combinedGroups = New-Object System.Collections.ArrayList
    $sectionHtmls = New-Object System.Collections.ArrayList

    foreach ($b in $Briefs) {
        foreach ($q in @($b.Result.Quotes)) { [void]$combinedQuotes.Add($q) }
        foreach ($s in @($b.Result.Status)) { [void]$combinedStatus.Add($s) }
        foreach ($g in @($b.Result.Groups)) { [void]$combinedGroups.Add($g) }

        # Map brief id to a CSS class so each section gets its own color theme.
        $cls = switch ($b.Result.Config.id) {
            'hkproperty' { 'section section-property' }
            'finance'    { 'section section-finance' }
            default      { 'section' }
        }
        $digest = Format-CompactDigest -Result $b.Result
        $body = if ($b.BriefingText) {
            (ConvertFrom-MarkdownLite -Markdown $b.BriefingText) + '<hr>' + (ConvertFrom-MarkdownLite -Markdown $digest)
        } else {
            ConvertFrom-MarkdownLite -Markdown $digest
        }
        [void]$sectionHtmls.Add(("<div class=`"{0}`">" -f $cls) + $body + '</div>')
    }

    # Aggregated banner: any chronic failure surfaces "needs repair"; failed sources aggregate.
    $chronicAny = @()
    $failedAny = @()
    foreach ($b in $Briefs) {
        $failed = @($b.Result.Status | Where-Object { $_.State -ne 'OK' })
        foreach ($f in $failed) { $failedAny += $f.Name }
        $bChronic = @()
        foreach ($s in $b.Result.Status) {
            if ($s.State -eq 'OK') { continue }
            $n = (Get-Streak -Meta $b.Meta -Name $s.Name) + 1
            if ($n -ge 3) { $bChronic += ('{0}（连续 {1} 次）' -f $s.Name, $n) }
        }
        $chronicAny += $bChronic
    }
    $banner = ''
    if ($chronicAny.Count -gt 0) {
        $banner += '<div class="banner warn"><b>需要检修：</b>' + [System.Net.WebUtility]::HtmlEncode(($chronicAny -join '、')) + ' 已连续多次抓取失败，可能是网站改版。</div>'
    } elseif ($failedAny.Count -gt 0) {
        $banner += '<div class="banner warn"><b>注意：</b>本次有 ' + $failedAny.Count + ' 个来源抓取失败（' + [System.Net.WebUtility]::HtmlEncode(($failedAny -join '、')) + '），相关内容可能缺失。</div>'
    }
    $hasAnyBriefing = @($Briefs | Where-Object { $_.BriefingText }).Count -gt 0
    if (-not $hasAnyBriefing) {
        $banner += '<div class="banner info">当前为<b>自动汇总的原始素材</b>，未经模型撰写成稿。设置环境变量 <code>ANTHROPIC_API_KEY</code> 后，各版块会自动换成按模板写好的简报。</div>'
    }

    # Build aggregated Result-like object for the combined summary box.
    $combined = [pscustomobject]@{
        Config = [pscustomobject]@{ title = '每日资讯' }
        RunAt = $Briefs[0].Result.RunAt
        Quotes = $combinedQuotes.ToArray()
        Status = $combinedStatus.ToArray()
        Groups = $combinedGroups.ToArray()
    }

    $html = ConvertTo-BriefingHtml -Title ('每日资讯 · {0}' -f $combined.RunAt.ToString('yyyy年M月d日')) `
                                   -BodyHtml ([string]::Join('<hr style="margin:40px 0;border:0;border-top:2px dashed #cfdcef">', $sectionHtmls.ToArray())) `
                                   -Banner $banner `
                                   -SummaryBox (ConvertTo-SummaryBox -Result $combined)

    $htmlPath = Join-Path $OutDir '今日资讯.html'
    Save-Archive -Path $htmlPath
    Write-Utf8 -Path $htmlPath -Content $html

    $delivered = $htmlPath
    $desktopOk = $false
    if (-not $NoDeliver -and $Delivery.desktopCopy) {
        try {
            $desk = [Environment]::GetFolderPath('Desktop')
            $safeTitle = ('每日资讯' -replace '[<>:"/\\|?*\x00-\x1F]', '_')
            $target = Join-Path $desk ('{0}.html' -f $safeTitle)
            Copy-Item -LiteralPath $htmlPath -Destination $target -Force
            $delivered = $target
            $desktopOk = $true
            Write-Log ("已送达桌面（合并版）：{0}" -f $target)
        } catch { Write-Log ("桌面送达失败：{0}" -f $_.Exception.Message) 'WARN' }
    }
    if (-not $NoDeliver -and $Delivery.autoOpen) {
        try { Start-Process $delivered | Out-Null } catch { }
    }
    if (-not $NoDeliver -and $Delivery.toast) {
        $newCount = 0
        foreach ($g in @($combined.Groups)) { $newCount += @($g.Items).Count }
        $msg = if ($chronicAny.Count -gt 0) { "合并简报：新增 $newCount 条，有来源需要检修" }
               elseif ($failedAny.Count -gt 0) { "合并简报：新增 $newCount 条，$($failedAny.Count) 个来源失败" }
               else { "合并简报：新增 $newCount 条，来源全部正常" }
        [void](Send-Toast -Title '每日资讯' -Message $msg)
    }
    return [pscustomobject]@{ Path = $delivered; Ok = ($NoDeliver -or (-not $Delivery.desktopCopy) -or $desktopOk) }
}

# ------------------------------------------------------------------ main ----

$RunAt = Get-Date
Limit-LogFile -Path $LogPath -MaxLines 2000
if (-not $OutputRoot) { $OutputRoot = Join-Path $root 'out' }
$dayDir = Join-Path $OutputRoot $RunAt.ToString('yyyy-MM-dd')

$delivery = Read-Json -Path (Join-Path $root 'config\delivery.json')
if (-not $delivery) { $delivery = [pscustomobject]@{ desktopCopy = $true; toast = $true; autoOpen = $true } }

$briefs = if ($Brief -eq 'all') { @('hkproperty', 'finance') } else { @($Brief) }
$mergedBriefs = New-Object System.Collections.ArrayList

foreach ($id in $briefs) {
    $cfgPath = Join-Path $root ('config\sources.{0}.json' -f $id)
    if (-not (Test-Path -LiteralPath $cfgPath)) { Write-Log "找不到配置：$cfgPath" 'ERROR'; continue }
    $cfg = Read-Json -Path $cfgPath
    if (-not $cfg) { Write-Log "配置解析失败：$cfgPath" 'ERROR'; continue }

    $metaPath = Join-Path $root ('state\meta.{0}.json' -f $id)
    $meta = Get-Meta -Path $metaPath

    # Two triggers (daily + at logon) mean this can fire twice; skip if today's
    # briefing was already delivered, unless explicitly forced.
    if (-not $Force -and $meta.lastDeliveredDate -eq $RunAt.ToString('yyyy-MM-dd')) {
        Write-Log ("{0}：今日已送达，跳过（需要重跑请加 -Force）" -f $cfg.title)
        continue
    }

    Write-Log ("开始采集：{0}" -f $cfg.title)
    $result = Invoke-Collect -Config $cfg -RunAt $RunAt -Meta $meta

    $promptPath = Join-Path $root ($cfg.promptFile -replace '/', '\')
    $prompt = if (Test-Path -LiteralPath $promptPath) { [System.IO.File]::ReadAllText($promptPath, [System.Text.Encoding]::UTF8) } else { '' }

    $outDir = Join-Path $dayDir $id
    $pack = Format-Pack -Result $result -Prompt $prompt

    $packPath = Join-Path $outDir 'pack.md'
    Save-Archive -Path $packPath
    Write-Utf8 -Path $packPath -Content $pack
    Write-Utf8 -Path (Join-Path $outDir 'raw.json') -Content ([pscustomobject]@{
        id = $cfg.id; title = $cfg.title; runAt = $RunAt.ToString('o'); windowHours = $result.Hours
        quotes = $result.Quotes; status = $result.Status; datasets = $result.Datasets; groups = $result.Groups
    } | ConvertTo-Json -Depth 8)

    $newItems = ($result.Groups | ForEach-Object { $_.Items.Count } | Measure-Object -Sum).Sum
    if (-not $newItems) { $newItems = 0 }
    $failCount = @($result.Status | Where-Object { $_.State -ne 'OK' }).Count
    Write-Log ("采集完成：{0} — 新增 {1} 条，来源失败 {2} 个" -f $cfg.title, $newItems, $failCount)

    # -- synthesis (optional) ---------------------------------------------
    $briefingText = ''
    $apiKey = $env:ANTHROPIC_API_KEY
    if (-not $NoSynthesis -and $apiKey) {
        try {
            Write-Log ("调用 Anthropic API 成稿（{0}）…" -f $Model)
            $briefingText = Invoke-Synthesis -Pack $pack -ModelName $Model -ApiKey $apiKey
            $bPath = Join-Path $outDir 'briefing.md'
            Save-Archive -Path $bPath
            Write-Utf8 -Path $bPath -Content $briefingText
            Write-Log ("成稿完成 -> {0}" -f $bPath)
        } catch {
            $briefingText = ''
            Write-Log ("成稿失败，改用原始素材送达：{0}" -f $_.Exception.Message) 'ERROR'
        }
    } elseif (-not $NoSynthesis) {
        Write-Log '未设置 ANTHROPIC_API_KEY，本次送达原始素材汇总。'
    }

    # -- deliver -----------------------------------------------------------
    $deliveryResult = Publish-Briefing -Result $result -BriefingText $briefingText `
                                  -OutDir $outDir -Delivery $delivery -Meta $meta
    Write-Log ("已送达：{0}" -f $deliveryResult.Path)

    # Collect for the optional merged briefing (built after this loop).
    if ($Merge) {
        [void]$mergedBriefs.Add([pscustomobject]@{ Result = $result; BriefingText = $briefingText; Meta = $meta })
    }

    # -- persist meta ------------------------------------------------------
    $streaks = New-Object System.Collections.ArrayList
    foreach ($s in $result.Status) {
        $n = if ($s.State -eq 'OK') { 0 } else { (Get-Streak -Meta $meta -Name $s.Name) + 1 }
        if ($n -gt 0) { [void]$streaks.Add([pscustomobject]@{ name = $s.Name; count = $n }) }
    }
    # lastDeliveredDate 只在送达成功时才写——桌面文件被占、Copy-Item 失败等情况下，
    # 下一次触发器（登录备用）还能再补一次，而不是被守卫误判为「今日已送达」跳过。
    Write-Utf8 -Path $metaPath -Content ([pscustomobject]@{
        lastSuccessAt = $RunAt.ToString('yyyy-MM-dd HH:mm:ss')
        lastDeliveredDate = $(if ($deliveryResult.Ok) { $RunAt.ToString('yyyy-MM-dd') } else { $meta.lastDeliveredDate })
        failStreak = @($streaks.ToArray())
    } | ConvertTo-Json -Depth 5)
}

# -- merged briefing (when -Merge) -----------------------------------------
if ($Merge -and $mergedBriefs.Count -gt 0) {
    $merged = Publish-MergedBriefing -Briefs $mergedBriefs.ToArray() -OutDir $dayDir -Delivery $delivery
    if ($merged) { Write-Log ("已送达合并版：{0}" -f $merged.Path) }

    # Push merged HTML to GitHub Pages so the phone sees the same content under
    # a stable URL. Silent no-op if state\github.token is missing or invalid.
    # On GitHub Actions the workflow handles commits itself, so skip the API push.
    if ($merged -and -not $GitHubActions) {
        $pushResult = Publish-GitHubPages -LocalPath $merged.Path `
                                          -RepoPath '今日资讯.html' `
                                          -CommitMessage ("每日资讯 · {0}" -f $RunAt.ToString('yyyy-MM-dd HH:mm'))
        if ($pushResult.Ok) {
            Write-Log '已推送到 GitHub Pages'
        } elseif ($pushResult.Skipped) {
            Write-Log ("GitHub 推送跳过：{0}" -f $pushResult.Skipped)
        } else {
            Write-Log ("GitHub 推送失败：{0}" -f $pushResult.Error) 'WARN'
        }
    } elseif ($merged -and $GitHubActions) {
        # On Actions, copy the merged HTML to the repo-root path the workflow expects
        # so the URL stays stable across runs (no /<date>/ in the path).
        $stablePath = Join-Path $repoRoot '今日资讯.html'
        if (Test-Path -LiteralPath $merged.Path) {
            Copy-Item -LiteralPath $merged.Path -Destination $stablePath -Force
            Write-Log ("已写入稳定路径：{0}" -f $stablePath)
        }
    }
}

Write-Log ("本次运行结束，输出目录：{0}" -f $dayDir)
if ($Open -and (Test-Path -LiteralPath $dayDir)) { Invoke-Item -LiteralPath $dayDir }
