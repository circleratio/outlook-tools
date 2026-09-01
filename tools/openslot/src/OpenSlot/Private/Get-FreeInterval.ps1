function ConvertTo-FlooredSlot {
    param([datetime]$Time, [int]$SlotMinutes)
    $day = $Time.Date
    $slots = [math]::Floor((($Time - $day).TotalMinutes) / $SlotMinutes)
    $day.AddMinutes($slots * $SlotMinutes)
}

function ConvertTo-CeiledSlot {
    param([datetime]$Time, [int]$SlotMinutes)
    $day = $Time.Date
    $slots = [math]::Ceiling((($Time - $day).TotalMinutes) / $SlotMinutes)
    $day.AddMinutes($slots * $SlotMinutes)
}

function ConvertTo-RoundedFreeSlot {
    <#
    空き部分区間を SlotMinutes グリッドに内側丸めし、
    MinimumMinutes 以上あれば区間を、なければ $null を返す。
    #>
    param([datetime]$Start, [datetime]$End, [int]$SlotMinutes, [int]$MinimumMinutes)

    $rStart = ConvertTo-CeiledSlot  -Time $Start -SlotMinutes $SlotMinutes
    $rEnd   = ConvertTo-FlooredSlot -Time $End   -SlotMinutes $SlotMinutes
    if (($rEnd - $rStart).TotalMinutes -ge $MinimumMinutes) {
        return [pscustomobject]@{ Start = $rStart; End = $rEnd }
    }
    return $null
}

function Get-FreeInterval {
    <#
    .SYNOPSIS
    対象日の営業時間ウィンドウ（8:00-12:00 / 13:00-19:00）から Busy 区間を減算し、
    SlotMinutes グリッドに内側丸めした上で MinimumMinutes 以上の空き区間を昇順で返す。
    早朝(8時前)・昼休み(12-13時)・夜(19時以降)は最初から対象外。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][datetime]$Date,
        [Parameter()][AllowNull()][pscustomobject[]]$Busy,
        [int]$SlotMinutes = 30,
        [int]$MinimumMinutes = 30
    )

    $d = $Date.Date
    $windows = @(
        (New-Interval -Start $d.AddHours(8)  -End $d.AddHours(12)),
        (New-Interval -Start $d.AddHours(13) -End $d.AddHours(19))
    )

    $merged = @(Merge-Interval -Interval $Busy)
    $free = New-Object System.Collections.Generic.List[pscustomobject]

    foreach ($w in $windows) {
        $cursor = $w.Start
        foreach ($b in $merged) {
            if ($b.End -le $w.Start -or $b.Start -ge $w.End) { continue }

            $bStart = if ($b.Start -lt $w.Start) { $w.Start } else { $b.Start }
            $bEnd   = if ($b.End   -gt $w.End)   { $w.End }   else { $b.End }

            if ($bStart -gt $cursor) {
                $slot = ConvertTo-RoundedFreeSlot -Start $cursor -End $bStart -SlotMinutes $SlotMinutes -MinimumMinutes $MinimumMinutes
                if ($slot) { $free.Add($slot) }
            }
            if ($bEnd -gt $cursor) { $cursor = $bEnd }
        }

        if ($cursor -lt $w.End) {
            $slot = ConvertTo-RoundedFreeSlot -Start $cursor -End $w.End -SlotMinutes $SlotMinutes -MinimumMinutes $MinimumMinutes
            if ($slot) { $free.Add($slot) }
        }
    }

    $free.ToArray()
}
