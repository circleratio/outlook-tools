function Get-FreeTime {
    <#
    .SYNOPSIS
    指定期間（両端を含む）の平日について、MS365 予定表の空き時間を抽出し、
    "M月d日(w): h:mm-h:mm, ..." 形式の行を文字列配列で返す。
    空き時間が無い営業日は行を返さない。

    MinimumMinutes で出力する空き枠の最小分数を指定する（既定 30）。

    予定の取得元は 2 通り:
    - -Events 未指定: Get-CalendarEvents（Microsoft Graph）で取得する。
      前提: 呼び出し前に Connect-MgGraph -Scopes Calendars.Read 済みであること。
    - -Events 指定: 渡された区間配列 @{ Start; End } をそのまま「埋まっている区間」として使う
      （Graph 通信を行わない）。空配列も可。showAs フィルタや終日予定の切り出しは注入側の責務。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][datetime]$Start,
        [Parameter(Mandatory)][datetime]$End,
        [int]$MinimumMinutes = 30,
        [pscustomobject[]]$Events
    )

    $days = @(Resolve-DateRange -Start $Start -End $End)
    if ($days.Count -eq 0) { return @() }

    $rangeStart = $days[0].Date
    $rangeEnd   = $days[-1].Date.AddDays(1)
    if ($PSBoundParameters.ContainsKey('Events')) {
        $events = @($Events)
    }
    else {
        $events = @(Get-CalendarEvents -Start $rangeStart -End $rangeEnd)
    }

    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($day in $days) {
        $dayStart = $day.Date
        $dayEnd   = $dayStart.AddDays(1)
        $todays = @($events | Where-Object { $_.Start -lt $dayEnd -and $_.End -gt $dayStart })

        $free = @(Get-FreeInterval -Date $day -Busy $todays -MinimumMinutes $MinimumMinutes)
        $line = Format-FreeDay -Date $day -Free $free
        if ($line) { $lines.Add($line) }
    }

    $lines.ToArray()
}
