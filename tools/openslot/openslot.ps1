#Requires -Version 5.1
<#
.SYNOPSIS
MS365 予定表から指定期間の空き時間を抽出し、標準出力とクリップボードへ出力する。

.DESCRIPTION
開始日〜終了日（両端を含む）の平日について、予定（表示区分 busy / oof / tentative）が
入っていない時間帯を抽出する。早朝(8時前)・昼休み(12-13時)・夜(19時以降)と土日は対象外。
空き枠は 30 分刻みに内側丸めし、30 分未満は出力しない。タイムゾーンは JST 固定。

MS365 へは実行ユーザーの権限（委任認証）でアクセスする。
初回実行時に Microsoft Graph へのサインイン（スコープ Calendars.Read）を求められる。

.PARAMETER Start
開始日。yyyy-MM-dd 形式、または年を省略した MM-dd 形式（年は実行年）。
省略時はコマンド実行日の翌日。

.PARAMETER End
終了日。yyyy-MM-dd 形式、または年を省略した MM-dd 形式（年は実行年）。両端を含む。
省略時は開始日の「次の金曜日」
（開始日が月〜木なら同じ週の金曜、金曜または土日開始なら翌週の金曜）。

.PARAMETER Week
対象期間を「Week 週後の月曜日から、Week 週後の金曜日まで」とする。0 以上の整数。
基準は実行日を含む週（月曜起算）。ただし `-Week 0`（今週）の開始日はその週の月曜ではなく「翌日」。
`0,2,4` のようにカンマ区切りで複数の週を指定でき、その場合は週ごとに期間を出力する（重複は除き昇順）。
`-Start` / `-End` とは同時に指定できない。

.PARAMETER Duration
出力する空き枠の最小サイズ（分）。30 の倍数（30 以上）で指定する。既定は 30。

.PARAMETER NoClipboard
クリップボードへのコピーを行わない。

.EXAMPLE
.\openslot.ps1 -Start 2026-09-01 -End 2026-09-05

.EXAMPLE
.\openslot.ps1 -Week 1 -Duration 60
1 時間以上まとまって空いている枠だけを出力する。

.EXAMPLE
.\openslot.ps1
実行日の翌日から、その週の金曜日までの空き時間を出力する。

.EXAMPLE
.\openslot.ps1 -Start 2026-09-01
2026-09-01 から次の金曜日（2026-09-04）までの空き時間を出力する。

.EXAMPLE
.\openslot.ps1 -Week 2
2 週後の月曜日から金曜日までの空き時間を出力する。

.EXAMPLE
.\openslot.ps1 -Week 0,1,3
今週（翌日〜金曜）・翌週・3 週後の、それぞれの空き時間を出力する。

.NOTES
終了コード: 0=正常 / 1=引数不正 / 2=認証失敗 / 3=カレンダー取得失敗
#>
[CmdletBinding()]
param(
    [string]$Start,
    [string]$End,
    [int[]]$Week,
    [int]$Duration = 30,
    [switch]$NoClipboard
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Stop-WithError {
    param([string]$Message, [int]$Code)
    [Console]::Error.WriteLine($Message)
    exit $Code
}

function Copy-EmptyAndExit {
    param([switch]$NoClipboard)
    if (-not $NoClipboard) {
        try { Set-Clipboard -Value '' } catch { }
    }
    exit 0
}

Import-Module (Join-Path $PSScriptRoot 'src/OpenSlot/OpenSlot.psd1') -Force

# --- パラメータ検証（指定されたものだけパース）---
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

# --- 省略時の既定値／-Week を解決する（1 つ以上の期間）---
$resolveArgs = @{ Start = $startDate; End = $endDate; Today = $today }
if ($weekProvided) { $resolveArgs['Week'] = $Week }
$periods = @(Resolve-Period @resolveArgs)

# 開始日 > 終了日の期間（既定値のズレ、-Week 0 を金曜以降に実行、など）は対象外
$periods = @($periods | Where-Object { $_.Start -le $_.End })
if ($periods.Count -eq 0) {
    Copy-EmptyAndExit -NoClipboard:$NoClipboard
}

# --- 認証（実行ユーザーの委任権限）---
try {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    if (-not (Get-MgContext)) {
        Connect-MgGraph -Scopes 'Calendars.Read' -NoWelcome -ErrorAction Stop | Out-Null
    }
}
catch {
    Stop-WithError "MS365 への接続に失敗しました: $($_.Exception.Message)" 2
}

# --- 取得・整形（期間ごとに実行して連結）---
try {
    $collected = New-Object System.Collections.Generic.List[string]
    foreach ($p in $periods) {
        foreach ($line in @(Get-FreeTime -Start $p.Start -End $p.End -MinimumMinutes $Duration)) {
            $collected.Add($line)
        }
    }
    $lines = @($collected)
}
catch {
    Stop-WithError "カレンダーの取得に失敗しました: $($_.Exception.Message)" 3
}

if ($lines.Count -gt 0) {
    $lines | ForEach-Object { Write-Output $_ }
}

if (-not $NoClipboard) {
    $text = ($lines -join [Environment]::NewLine)
    try {
        Set-Clipboard -Value $text
    }
    catch {
        [Console]::Error.WriteLine("クリップボードへのコピーに失敗しました: $($_.Exception.Message)")
    }
}

exit 0
