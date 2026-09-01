$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$here/../src/OpenSlot/Private/New-Interval.ps1"
. "$here/../src/OpenSlot/Private/Merge-Interval.ps1"
. "$here/../src/OpenSlot/Private/Get-FreeInterval.ps1"

function iv($s, $e) { New-Interval -Start ([datetime]$s) -End ([datetime]$e) }
function fmt($x) { '{0:HH:mm}-{1:HH:mm}' -f $x.Start, $x.End }

Describe 'Get-FreeInterval' {

    It '予定なしなら午前・午後の 2 枠' {
        $r = @(Get-FreeInterval -Date ([datetime]'2026-09-01') -Busy @())
        $r.Count | Should Be 2
        (fmt $r[0]) | Should Be '08:00-12:00'
        (fmt $r[1]) | Should Be '13:00-19:00'
    }

    It '午前を完全に埋める予定があれば午後のみ' {
        $r = @(Get-FreeInterval -Date ([datetime]'2026-09-01') -Busy @((iv '2026-09-01 08:00' '2026-09-01 12:00')))
        $r.Count | Should Be 1
        (fmt $r[0]) | Should Be '13:00-19:00'
    }

    It '予定境界を 30 分グリッドに内側丸めする' {
        # 09:15-09:45 の予定 -> 08:00-09:00 と 10:00-12:00 が残る
        $r = @(Get-FreeInterval -Date ([datetime]'2026-09-01') -Busy @((iv '2026-09-01 09:15' '2026-09-01 09:45')))
        $morning = @($r | Where-Object { $_.Start.Hour -lt 12 })
        $morning.Count | Should Be 2
        (fmt $morning[0]) | Should Be '08:00-09:00'
        (fmt $morning[1]) | Should Be '10:00-12:00'
    }

    It '丸めた結果 30 分未満になる空きは出力しない' {
        # 08:15 まで空き -> 切り上げ 08:30、その後 08:40-19:00 が予定 -> 08:00-08:15 は消える
        $r = @(Get-FreeInterval -Date ([datetime]'2026-09-01') -Busy @((iv '2026-09-01 08:15' '2026-09-01 19:00')))
        $r.Count | Should Be 0
    }

    It '昼をまたぐ予定は午前・午後それぞれで切られる' {
        $r = @(Get-FreeInterval -Date ([datetime]'2026-09-01') -Busy @((iv '2026-09-01 11:00' '2026-09-01 14:00')))
        $r.Count | Should Be 2
        (fmt $r[0]) | Should Be '08:00-11:00'
        (fmt $r[1]) | Should Be '14:00-19:00'
    }

    It '終日埋まっていれば空配列' {
        $r = @(Get-FreeInterval -Date ([datetime]'2026-09-01') -Busy @((iv '2026-09-01 00:00' '2026-09-02 00:00')))
        $r.Count | Should Be 0
    }

    It '-MinimumMinutes 60 は 30 分枠を除外する' {
        $busy = @((iv '2026-09-01 08:30' '2026-09-01 12:00'))  # 午前は 08:00-08:30 の 30 分だけ空く
        $r30 = @(Get-FreeInterval -Date ([datetime]'2026-09-01') -Busy $busy -MinimumMinutes 30)
        $r60 = @(Get-FreeInterval -Date ([datetime]'2026-09-01') -Busy $busy -MinimumMinutes 60)
        (@($r30 | Where-Object { $_.Start.Hour -lt 12 })).Count | Should Be 1
        (@($r60 | Where-Object { $_.Start.Hour -lt 12 })).Count | Should Be 0
    }

    It '昼休みの中だけの予定は無視される' {
        $r = @(Get-FreeInterval -Date ([datetime]'2026-09-01') -Busy @((iv '2026-09-01 12:15' '2026-09-01 12:45')))
        $r.Count | Should Be 2
        (fmt $r[0]) | Should Be '08:00-12:00'
        (fmt $r[1]) | Should Be '13:00-19:00'
    }
}
