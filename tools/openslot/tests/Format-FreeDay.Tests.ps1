$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$here/../src/OpenSlot/Private/New-Interval.ps1"
. "$here/../src/OpenSlot/Private/Format-FreeDay.ps1"

function iv($s, $e) { New-Interval -Start ([datetime]$s) -End ([datetime]$e) }

Describe 'Format-FreeDay' {

    It '日・時は 0 埋めせず、分は 2 桁、曜日は日本語' {
        $line = Format-FreeDay -Date ([datetime]'2026-09-01') -Free @((iv '2026-09-01 08:00' '2026-09-01 10:00'))
        $line | Should Be '9月1日(火): 8:00-10:00'
    }

    It '分ありの時刻を 2 桁で表示する' {
        $line = Format-FreeDay -Date ([datetime]'2026-09-01') -Free @((iv '2026-09-01 13:00' '2026-09-01 15:30'))
        $line | Should Be '9月1日(火): 13:00-15:30'
    }

    It '複数枠をカンマで連結する' {
        $free = @(
            (iv '2026-09-01 08:00' '2026-09-01 10:00'),
            (iv '2026-09-01 13:00' '2026-09-01 15:30'),
            (iv '2026-09-01 17:00' '2026-09-01 19:00')
        )
        $line = Format-FreeDay -Date ([datetime]'2026-09-01') -Free $free
        $line | Should Be '9月1日(火): 8:00-10:00, 13:00-15:30, 17:00-19:00'
    }

    It '各曜日を日本語表記する' {
        (Format-FreeDay -Date ([datetime]'2026-08-31') -Free @((iv '2026-08-31 08:00' '2026-08-31 09:00'))) | Should Be '8月31日(月): 8:00-9:00'
        (Format-FreeDay -Date ([datetime]'2026-09-04') -Free @((iv '2026-09-04 08:00' '2026-09-04 09:00'))) | Should Be '9月4日(金): 8:00-9:00'
    }

    It 'Free が空なら $null' {
        Format-FreeDay -Date ([datetime]'2026-09-01') -Free @() | Should BeNullOrEmpty
    }
}
