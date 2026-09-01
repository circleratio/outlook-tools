function Resolve-DateRange {
    <#
    .SYNOPSIS
    開始日・終了日（両端を含む）から、対象の平日（月〜金）を昇順で返す。
    土日は除外する。祝日は現行仕様では除外しない（将来オプション化する拡張点）。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][datetime]$Start,
        [Parameter(Mandatory)][datetime]$End
    )

    $s = $Start.Date
    $e = $End.Date
    if ($e -lt $s) {
        throw '終了日は開始日以降を指定してください'
    }

    $days = New-Object System.Collections.Generic.List[datetime]
    for ($d = $s; $d -le $e; $d = $d.AddDays(1)) {
        if ($d.DayOfWeek -eq [DayOfWeek]::Saturday -or $d.DayOfWeek -eq [DayOfWeek]::Sunday) {
            continue
        }
        # 祝日フィルタは将来ここに追加する（現行仕様では検索対象に含める）
        $days.Add($d)
    }

    $days.ToArray()
}
