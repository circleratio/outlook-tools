function Format-FreeDay {
    <#
    .SYNOPSIS
    1 営業日分の空き区間を "M月d日(w): h:mm-h:mm, ..." に整形する。
    日付・時は 0 埋めしない。分は 2 桁 0 埋め。曜日は日本語。
    Free が空なら $null を返す（呼び出し側で当該日をスキップ）。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][datetime]$Date,
        [Parameter()][AllowNull()][pscustomobject[]]$Free
    )

    $items = @($Free | Where-Object { $_ })
    if ($items.Count -eq 0) { return $null }

    $wdays = @('日', '月', '火', '水', '木', '金', '土')
    $w = $wdays[[int]$Date.DayOfWeek]

    $slots = $items | ForEach-Object {
        '{0}:{1:00}-{2}:{3:00}' -f $_.Start.Hour, $_.Start.Minute, $_.End.Hour, $_.End.Minute
    }

    '{0}月{1}日({2}): {3}' -f $Date.Month, $Date.Day, $w, ($slots -join ', ')
}
