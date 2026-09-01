function Resolve-Period {
    <#
    .SYNOPSIS
    省略された開始日・終了日、または -Week 指定から対象期間の配列（各要素 @{ Start; End }）を返す。

    .DESCRIPTION
    未指定の日付は [datetime]::MinValue で渡す（openslot.ps1 が省略時にそうする）。
    Today 未指定時は [datetime]::Today を用いる。戻り値は常に 1 要素以上の配列。

    - Week に 1 つ以上の値を指定した場合（週指定モード）:
      Today を含む週（月曜起算）の月曜日を基準に、各 w について
      End = 基準月曜 ＋ w×7 日 ＋ 4 日（その週の金曜）、
      Start = w が 0（今週）なら Today の翌日、w >= 1 なら 基準月曜 ＋ w×7 日。
      重複は除き、昇順に整列した週ごとに 1 期間を返す。Start / End 引数は無視する。
      （w が 0 で Today が金曜以降の場合、Start > End になる。逆転の除外は openslot.ps1 が行う。）
    - Week 未指定（空）:
      Start 未指定 → Today の翌日 / End 未指定 → 確定した Start の「次の金曜日」（Get-NextFriday）。
      1 要素の配列を返す。

    逆転（End < Start）や Week の負数・排他の検証は行わない（openslot.ps1 の責務）。
    #>
    [CmdletBinding()]
    param(
        [datetime]$Start = [datetime]::MinValue,
        [datetime]$End   = [datetime]::MinValue,
        [int[]]$Week     = @(),
        [datetime]$Today = [datetime]::MinValue
    )

    if ($Today -eq [datetime]::MinValue) { $Today = [datetime]::Today }

    if ($Week.Count -gt 0) {
        $daysSinceMonday = ([int]$Today.DayOfWeek + 6) % 7      # 月曜=0 … 日曜=6
        $baseMonday = $Today.Date.AddDays(-$daysSinceMonday)
        $tomorrow   = $Today.Date.AddDays(1)

        return @(
            $Week | Sort-Object -Unique | ForEach-Object {
                $monday = $baseMonday.AddDays(7 * $_)
                $start  = if ($_ -eq 0) { $tomorrow } else { $monday }
                [pscustomobject]@{ Start = $start; End = $monday.AddDays(4) }
            }
        )
    }

    if ($Start -eq [datetime]::MinValue) { $Start = $Today.Date.AddDays(1) }
    if ($End   -eq [datetime]::MinValue) { $End   = Get-NextFriday -From $Start }

    @( [pscustomobject]@{ Start = $Start.Date; End = $End.Date } )
}
