function Get-NextFriday {
    <#
    .SYNOPSIS
    From の「次の金曜日」を返す。
    月〜木 → 同じ週の金曜日 / 金曜 → 翌週の金曜日 / 土日 → 次に来る金曜日（翌週の金曜日）。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][datetime]$From
    )

    # 金曜(DayOfWeek=5)までの日数。From が金曜なら 0 → 翌週まわし
    $delta = (5 - [int]$From.DayOfWeek + 7) % 7
    if ($delta -eq 0) { $delta = 7 }

    $From.Date.AddDays($delta)
}
