$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$here/../src/OpenSlot/Private/Resolve-DateRange.ps1"

Describe 'Resolve-DateRange' {

    It '平日単日はその日のみ返す' {
        $r = @(Resolve-DateRange -Start '2026-09-01' -End '2026-09-01')  # 火
        $r.Count | Should Be 1
        $r[0] | Should Be ([datetime]'2026-09-01')
    }

    It '土日単日は空配列' {
        $r = @(Resolve-DateRange -Start '2026-09-05' -End '2026-09-06')  # 土日
        $r.Count | Should Be 0
    }

    It '週をまたぐ期間から土日を除外する' {
        $r = @(Resolve-DateRange -Start '2026-08-31' -End '2026-09-07')  # 月〜翌月
        $r.Count | Should Be 6
        ($r | Where-Object { $_.DayOfWeek -eq [DayOfWeek]::Saturday -or $_.DayOfWeek -eq [DayOfWeek]::Sunday }).Count | Should Be 0
        $r[0] | Should Be ([datetime]'2026-08-31')
        $r[-1] | Should Be ([datetime]'2026-09-07')
    }

    It '終了日が開始日より前なら例外' {
        { Resolve-DateRange -Start '2026-09-05' -End '2026-09-01' } | Should Throw
    }
}
