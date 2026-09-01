function Merge-Interval {
    <#
    .SYNOPSIS
    区間の配列を Start 昇順にソートし、重なり・隣接するものを統合して返す。
    空入力・$null は空配列を返す。
    #>
    [CmdletBinding()]
    param(
        [Parameter()][AllowNull()][pscustomobject[]]$Interval
    )

    $items = @($Interval | Where-Object { $_ })
    if ($items.Count -eq 0) { return @() }

    $sorted = @($items | Sort-Object Start)
    $result = New-Object System.Collections.Generic.List[pscustomobject]
    $cur = [pscustomobject]@{ Start = $sorted[0].Start; End = $sorted[0].End }

    foreach ($iv in ($sorted | Select-Object -Skip 1)) {
        if ($iv.Start -le $cur.End) {
            if ($iv.End -gt $cur.End) { $cur.End = $iv.End }
        }
        else {
            $result.Add($cur)
            $cur = [pscustomobject]@{ Start = $iv.Start; End = $iv.End }
        }
    }
    $result.Add($cur)

    $result.ToArray()
}
