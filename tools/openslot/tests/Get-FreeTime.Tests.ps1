$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Get-ChildItem "$here/../src/OpenSlot/Private/*.ps1" | ForEach-Object { . $_.FullName }
. "$here/../src/OpenSlot/Public/Get-FreeTime.ps1"

function iv($s, $e) { New-Interval -Start ([datetime]$s) -End ([datetime]$e) }

# Get-CalendarEvents をテスト用スタブで差し替える（Graph 通信を行わない）
$script:StubEvents = @()
function Get-CalendarEvents { param($Start, $End) return $script:StubEvents }

Describe 'Get-FreeTime' {

    It '土日を除いた平日のみ行を返す' {
        $script:StubEvents = @()
        $r = @(Get-FreeTime -Start '2026-09-04' -End '2026-09-07')  # 金・土・日・月
        $r.Count | Should Be 2
        $r[0] | Should Be '9月4日(金): 8:00-12:00, 13:00-19:00'
        $r[1] | Should Be '9月7日(月): 8:00-12:00, 13:00-19:00'
    }

    It '空きが無い営業日は行を出さない' {
        $script:StubEvents = @((iv '2026-09-01 08:00' '2026-09-01 19:00'))  # 火を終日埋める
        $r = @(Get-FreeTime -Start '2026-09-01' -End '2026-09-02')  # 火・水
        $r.Count | Should Be 1
        $r[0] | Should Match '^9月2日\(水\):'
    }

    It '期間内に平日が無ければ空配列' {
        $script:StubEvents = @()
        (@(Get-FreeTime -Start '2026-09-05' -End '2026-09-06')).Count | Should Be 0
    }

    It '該当日の予定だけを差し引く' {
        $script:StubEvents = @(
            (iv '2026-09-01 09:00' '2026-09-01 10:00'),
            (iv '2026-09-02 14:00' '2026-09-02 15:00')
        )
        $r = @(Get-FreeTime -Start '2026-09-01' -End '2026-09-02')
        $r[0] | Should Be '9月1日(火): 8:00-9:00, 10:00-12:00, 13:00-19:00'
        $r[1] | Should Be '9月2日(水): 8:00-12:00, 13:00-14:00, 15:00-19:00'
    }

    It '-MinimumMinutes で最小枠を絞る' {
        $script:StubEvents = @((iv '2026-09-01 08:30' '2026-09-01 19:00'))  # 火。午前 08:00-08:30 のみ空く
        $r30 = @(Get-FreeTime -Start '2026-09-01' -End '2026-09-01' -MinimumMinutes 30)
        $r60 = @(Get-FreeTime -Start '2026-09-01' -End '2026-09-01' -MinimumMinutes 60)
        $r30[0] | Should Be '9月1日(火): 8:00-8:30'
        $r60.Count | Should Be 0
    }

    It '-Events を渡すと Get-CalendarEvents を呼ばずその区間で算出する' {
        $script:StubEvents = @((iv '2026-09-01 08:00' '2026-09-01 19:00'))  # スタブ側は終日埋め
        $injected = @((iv '2026-09-01 09:00' '2026-09-01 10:00'))
        $r = @(Get-FreeTime -Start '2026-09-01' -End '2026-09-01' -Events $injected)
        $r[0] | Should Be '9月1日(火): 8:00-9:00, 10:00-12:00, 13:00-19:00'
    }

    It '-Events @() は予定なし扱い' {
        $script:StubEvents = @((iv '2026-09-01 08:00' '2026-09-01 19:00'))
        $r = @(Get-FreeTime -Start '2026-09-01' -End '2026-09-01' -Events @())
        $r[0] | Should Be '9月1日(火): 8:00-12:00, 13:00-19:00'
    }
}
