#Requires -Version 5.1
<#
.SYNOPSIS
openslot の MCP 版バックエンド。カレンダー取得は行わず、Claude が Microsoft 365 コネクタの
outlook_calendar_search で集めた予定を受け取って空き時間を算出する。2 フェーズで動作する。

.DESCRIPTION
フェーズ1（既定 / -ResolveOnly）:
  引数を検証し、対象期間（-Week や既定値を含む）を解決して JSON で標準出力へ返す。
  Claude はこの JSON の fetch 範囲で outlook_calendar_search を呼ぶ。

フェーズ2（-PlanPath <p> -EventsPath <e>）:
  フェーズ1 の JSON（plan）と、outlook_calendar_search の結果を格納した配列（events）を読み、
  空き時間の行を標準出力へ出し、同じ全文をクリップボードへコピーする。

-Start / -End / -Week / -Duration / -NoClipboard は openslot.ps1 と同じ意味・既定値・検証。

.PARAMETER Start
開始日。yyyy-MM-dd または年を省略した MM-dd。省略時は実行日の翌日。

.PARAMETER End
終了日（両端を含む）。yyyy-MM-dd または MM-dd。省略時は開始日の「次の金曜日」。

.PARAMETER Week
対象期間を「Week 週後の月曜〜金曜」とする 0 以上の整数（カンマ区切りで複数可）。-Start / -End と排他。

.PARAMETER Duration
出力する空き枠の最小サイズ（分）。30 の倍数（30 以上）。既定 30。

.PARAMETER NoClipboard
クリップボードへのコピーを行わない（フェーズ2）。

.PARAMETER ResolveOnly
フェーズ1（期間解決）を明示する。指定しなくてもフェーズ2 用の引数が無ければフェーズ1 が動く。

.PARAMETER PlanPath
フェーズ2。フェーズ1 が出力した plan JSON のパス。

.PARAMETER EventsPath
フェーズ2。outlook_calendar_search の結果オブジェクトを格納した JSON 配列のパス。

.NOTES
終了コード: 0=正常 / 1=引数不正 / 3=処理失敗（plan・events JSON 不正など）。
認証は Claude の Microsoft 365 コネクタ側で確立済みのため、認証失敗コードは無い。
#>
[CmdletBinding()]
param(
    [string]$Start,
    [string]$End,
    [int[]]$Week,
    [int]$Duration = 30,
    [switch]$NoClipboard,
    [switch]$ResolveOnly,
    [string]$PlanPath,
    [string]$EventsPath
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Stop-WithError {
    param([string]$Message, [int]$Code)
    [Console]::Error.WriteLine($Message)
    exit $Code
}

Import-Module (Join-Path $PSScriptRoot 'src/OpenSlot/OpenSlot.psd1') -Force

$ci = [System.Globalization.CultureInfo]::InvariantCulture

# ============================================================================
# フェーズ2: plan + events から空き時間を算出
# ============================================================================
if ($PSBoundParameters.ContainsKey('PlanPath') -or $PSBoundParameters.ContainsKey('EventsPath')) {
    if (-not ($PSBoundParameters.ContainsKey('PlanPath') -and $PSBoundParameters.ContainsKey('EventsPath'))) {
        Stop-WithError "-PlanPath と -EventsPath は両方指定してください" 1
    }

    try {
        $plan = (Get-Content -LiteralPath $PlanPath -Raw -ErrorAction Stop) | ConvertFrom-Json
    }
    catch {
        Stop-WithError "plan JSON の読み込みに失敗しました: $($_.Exception.Message)" 3
    }

    try {
        # PS 5.1 の ConvertFrom-Json は JSON 配列をパイプラインに 1 オブジェクトとして流すため、
        # 一旦変数へ受けてから @() で正規化する（@(... | ConvertFrom-Json) だと入れ子になる）。
        $parsedEvents = (Get-Content -LiteralPath $EventsPath -Raw -ErrorAction Stop) | ConvertFrom-Json
        $rawEvents = @($parsedEvents)
    }
    catch {
        Stop-WithError "events JSON の読み込みに失敗しました: $($_.Exception.Message)" 3
    }

    $duration = if ($plan.PSObject.Properties['duration']) { [int]$plan.duration } else { 30 }
    $noClip = [bool]$NoClipboard -or ($plan.PSObject.Properties['noClipboard'] -and [bool]$plan.noClipboard)
    $periods = if ($plan.PSObject.Properties['periods'] -and $plan.periods) { @($plan.periods) } else { @() }

    try {
        $busy = @(ConvertFrom-CalendarSearchEvent -Event $rawEvents)

        $collected = New-Object System.Collections.Generic.List[string]
        foreach ($p in $periods) {
            $ps = [datetime]::ParseExact([string]$p.start, 'yyyy-MM-dd', $ci)
            $pe = [datetime]::ParseExact([string]$p.end,   'yyyy-MM-dd', $ci)
            foreach ($line in @(Get-FreeTime -Start $ps -End $pe -MinimumMinutes $duration -Events $busy)) {
                $collected.Add($line)
            }
        }
        $lines = @($collected)
    }
    catch {
        Stop-WithError "空き時間の算出に失敗しました: $($_.Exception.Message)" 3
    }

    if ($lines.Count -gt 0) {
        $lines | ForEach-Object { Write-Output $_ }
    }

    if (-not $noClip) {
        $text = ($lines -join [Environment]::NewLine)
        try {
            Set-Clipboard -Value $text
        }
        catch {
            [Console]::Error.WriteLine("クリップボードへのコピーに失敗しました: $($_.Exception.Message)")
        }
    }

    exit 0
}

# ============================================================================
# フェーズ1: 引数検証 + 期間解決 → plan JSON を出力
# ============================================================================
$today = [datetime]::Today
$startProvided = $PSBoundParameters.ContainsKey('Start')
$endProvided   = $PSBoundParameters.ContainsKey('End')
$weekProvided  = $PSBoundParameters.ContainsKey('Week')

if ($Duration -lt 30 -or $Duration % 30 -ne 0) {
    Stop-WithError "-Duration には 30 の倍数（30 以上）を分単位で指定してください: $Duration" 1
}

if ($weekProvided) {
    if ($startProvided -or $endProvided) {
        Stop-WithError "-Week は -Start / -End と同時に指定できません" 1
    }
    if (@($Week).Count -eq 0) {
        Stop-WithError "-Week には 1 つ以上の 0 以上の整数を指定してください" 1
    }
    if (@($Week | Where-Object { $_ -lt 0 }).Count -gt 0) {
        Stop-WithError "-Week には 0 以上の整数を指定してください: $($Week -join ',')" 1
    }
}

$startDate = [datetime]::MinValue
$endDate   = [datetime]::MinValue

if ($startProvided) {
    try { $startDate = ConvertTo-InputDate -Text $Start -Today $today }
    catch { Stop-WithError "開始日エラー — $($_.Exception.Message)" 1 }
}
if ($endProvided) {
    try { $endDate = ConvertTo-InputDate -Text $End -Today $today }
    catch { Stop-WithError "終了日エラー — $($_.Exception.Message)" 1 }
}
if ($startProvided -and $endProvided -and $endDate.Date -lt $startDate.Date) {
    Stop-WithError "終了日は開始日以降を指定してください: $Start..$End" 1
}

$resolveArgs = @{ Start = $startDate; End = $endDate; Today = $today }
if ($weekProvided) { $resolveArgs['Week'] = $Week }
$periods = @(Resolve-Period @resolveArgs | Where-Object { $_.Start -le $_.End })

$periodList = New-Object System.Collections.Generic.List[pscustomobject]
foreach ($p in $periods) {
    $periodList.Add([pscustomobject]@{
        start = $p.Start.ToString('yyyy-MM-dd')
        end   = $p.End.ToString('yyyy-MM-dd')
    })
}

$fetch = $null
if ($periods.Count -gt 0) {
    $minStart = ($periods | ForEach-Object { $_.Start } | Sort-Object)[0]
    $maxEnd   = ($periods | ForEach-Object { $_.End }   | Sort-Object)[-1]
    $fetch = [pscustomobject]@{
        start = $minStart.ToString('yyyy-MM-dd')
        end   = $maxEnd.AddDays(1).ToString('yyyy-MM-dd')
    }
}

[pscustomobject]@{
    periods     = $periodList.ToArray()
    fetch       = $fetch
    duration    = $Duration
    noClipboard = [bool]$NoClipboard
} | ConvertTo-Json -Depth 5

exit 0
