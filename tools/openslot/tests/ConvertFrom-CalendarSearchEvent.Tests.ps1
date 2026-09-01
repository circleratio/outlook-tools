$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$here/../src/OpenSlot/Private/New-Interval.ps1"
. "$here/../src/OpenSlot/Public/ConvertFrom-CalendarSearchEvent.ps1"

function ev {
    param(
        $showAs,
        $start,
        $end,
        $cancelled = $false,
        $allDay = $false,
        $tz = 'Tokyo Standard Time'
    )
    [pscustomobject]@{
        subject     = 'x'
        showAs      = $showAs
        isCancelled = $cancelled
        isAllDay    = $allDay
        start       = [pscustomobject]@{ dateTime = $start; timeZone = $tz }
        end         = [pscustomobject]@{ dateTime = $end;   timeZone = $tz }
    }
}

Describe 'ConvertFrom-CalendarSearchEvent' {

    It 'busy を採用し Start/End 区間を返す' {
        $r = @(ConvertFrom-CalendarSearchEvent -Event @(ev 'busy' '2026-09-01T09:00:00.0000000' '2026-09-01T10:00:00.0000000'))
        $r.Count | Should Be 1
        $r[0].Start | Should Be ([datetime]'2026-09-01 09:00')
        $r[0].End   | Should Be ([datetime]'2026-09-01 10:00')
    }

    It 'oof / tentative を採用する' {
        $r = @(ConvertFrom-CalendarSearchEvent -Event @(
                (ev 'oof'       '2026-09-01T09:00:00' '2026-09-01T10:00:00'),
                (ev 'tentative' '2026-09-01T11:00:00' '2026-09-01T12:00:00')
            ))
        $r.Count | Should Be 2
    }

    It 'free / workingElsewhere を除外する' {
        $r = @(ConvertFrom-CalendarSearchEvent -Event @(
                (ev 'free'            '2026-09-01T09:00:00' '2026-09-01T10:00:00'),
                (ev 'workingElsewhere' '2026-09-01T11:00:00' '2026-09-01T12:00:00')
            ))
        $r.Count | Should Be 0
    }

    It 'showAs の大文字小文字を無視する' {
        $r = @(ConvertFrom-CalendarSearchEvent -Event @(ev 'Busy' '2026-09-01T09:00:00' '2026-09-01T10:00:00'))
        $r.Count | Should Be 1
    }

    It 'isCancelled=true を除外する' {
        $r = @(ConvertFrom-CalendarSearchEvent -Event @(ev 'busy' '2026-09-01T09:00:00' '2026-09-01T10:00:00' $true))
        $r.Count | Should Be 0
    }

    It '終日予定 (isAllDay) を区間として扱う' {
        $r = @(ConvertFrom-CalendarSearchEvent -Event @(ev 'oof' '2026-09-01T00:00:00.0000000' '2026-09-02T00:00:00.0000000' $false $true))
        $r.Count | Should Be 1
        $r[0].Start | Should Be ([datetime]'2026-09-01 00:00')
        $r[0].End   | Should Be ([datetime]'2026-09-02 00:00')
    }

    It 'dateTime を JST 壁時計 (Kind=Unspecified) として解釈する' {
        $r = @(ConvertFrom-CalendarSearchEvent -Event @(ev 'busy' '2026-09-01T09:30:00.0000000' '2026-09-01T10:00:00.0000000'))
        $r[0].Start.Kind | Should Be 'Unspecified'
        $r[0].Start.Hour | Should Be 9
        $r[0].Start.Minute | Should Be 30
    }

    It 'timeZone が Tokyo 以外なら warning を出す (パースは継続)' {
        $w = $null
        $r = @(ConvertFrom-CalendarSearchEvent -Event @(ev 'busy' '2026-09-01T09:00:00' '2026-09-01T10:00:00' $false $false 'Pacific Standard Time') -WarningVariable w -WarningAction SilentlyContinue)
        $r.Count | Should Be 1
        @($w).Count | Should Not Be 0
    }

    It 'End <= Start の予定を除外する' {
        $r = @(ConvertFrom-CalendarSearchEvent -Event @(ev 'busy' '2026-09-01T10:00:00' '2026-09-01T10:00:00'))
        $r.Count | Should Be 0
    }

    It '空入力は空配列' {
        (@(ConvertFrom-CalendarSearchEvent -Event @())).Count | Should Be 0
    }

    It '$null 入力は空配列' {
        (@(ConvertFrom-CalendarSearchEvent -Event $null)).Count | Should Be 0
    }

    It '日時がパースできない要素は除外する' {
        $r = @(ConvertFrom-CalendarSearchEvent -Event @(
                (ev 'busy' 'not-a-date' 'also-bad'),
                (ev 'busy' '2026-09-01T09:00:00' '2026-09-01T10:00:00')
            ))
        $r.Count | Should Be 1
    }

    It 'ハッシュテーブル形式の予定も受け付ける' {
        $h = @{
            showAs      = 'busy'
            isCancelled = $false
            start       = @{ dateTime = '2026-09-01T09:00:00'; timeZone = 'Tokyo Standard Time' }
            end         = @{ dateTime = '2026-09-01T10:00:00'; timeZone = 'Tokyo Standard Time' }
        }
        $r = @(ConvertFrom-CalendarSearchEvent -Event @($h))
        $r.Count | Should Be 1
    }
}
