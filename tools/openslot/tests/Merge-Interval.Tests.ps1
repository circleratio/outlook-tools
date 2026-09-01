$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$here/../src/OpenSlot/Private/New-Interval.ps1"
. "$here/../src/OpenSlot/Private/Merge-Interval.ps1"

function iv($s, $e) { New-Interval -Start ([datetime]$s) -End ([datetime]$e) }

Describe 'Merge-Interval' {

    It '空入力は空配列' {
        (@(Merge-Interval -Interval @())).Count | Should Be 0
    }

    It '$null 入力は空配列' {
        (@(Merge-Interval -Interval $null)).Count | Should Be 0
    }

    It '重なる 2 区間を統合する' {
        $r = @(Merge-Interval -Interval @((iv '2026-09-01 09:00' '2026-09-01 10:00'), (iv '2026-09-01 09:30' '2026-09-01 11:00')))
        $r.Count | Should Be 1
        $r[0].Start | Should Be ([datetime]'2026-09-01 09:00')
        $r[0].End | Should Be ([datetime]'2026-09-01 11:00')
    }

    It '隣接する区間を統合する' {
        $r = @(Merge-Interval -Interval @((iv '2026-09-01 09:00' '2026-09-01 10:00'), (iv '2026-09-01 10:00' '2026-09-01 10:30')))
        $r.Count | Should Be 1
        $r[0].End | Should Be ([datetime]'2026-09-01 10:30')
    }

    It '内包される区間を統合する' {
        $r = @(Merge-Interval -Interval @((iv '2026-09-01 09:00' '2026-09-01 12:00'), (iv '2026-09-01 10:00' '2026-09-01 11:00')))
        $r.Count | Should Be 1
        $r[0].End | Should Be ([datetime]'2026-09-01 12:00')
    }

    It '離れた 2 区間は分かれたまま・順不同でもソートされる' {
        $r = @(Merge-Interval -Interval @((iv '2026-09-01 15:00' '2026-09-01 16:00'), (iv '2026-09-01 09:00' '2026-09-01 10:00')))
        $r.Count | Should Be 2
        $r[0].Start | Should Be ([datetime]'2026-09-01 09:00')
        $r[1].Start | Should Be ([datetime]'2026-09-01 15:00')
    }
}
