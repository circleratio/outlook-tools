$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$here/../src/OpenSlot/Private/Get-NextFriday.ps1"
. "$here/../src/OpenSlot/Public/Resolve-Period.ps1"

Describe 'Get-NextFriday' {

    It '月曜からは同じ週の金曜' {
        (Get-NextFriday -From ([datetime]'2026-08-31')) | Should Be ([datetime]'2026-09-04')
    }
    It '木曜からは翌日の金曜' {
        (Get-NextFriday -From ([datetime]'2026-09-03')) | Should Be ([datetime]'2026-09-04')
    }
    It '金曜からは翌週の金曜' {
        (Get-NextFriday -From ([datetime]'2026-09-04')) | Should Be ([datetime]'2026-09-11')
    }
    It '土曜からは次に来る金曜（翌週）' {
        (Get-NextFriday -From ([datetime]'2026-09-05')) | Should Be ([datetime]'2026-09-11')
    }
    It '日曜からは次に来る金曜（翌週）' {
        (Get-NextFriday -From ([datetime]'2026-09-06')) | Should Be ([datetime]'2026-09-11')
    }
    It '時刻部分は無視される' {
        (Get-NextFriday -From ([datetime]'2026-08-31 23:30')) | Should Be ([datetime]'2026-09-04')
    }
}

Describe 'Resolve-Period' {

    $today = [datetime]'2026-09-01'  # 火

    It '戻り値は常に配列' {
        (@(Resolve-Period -Today $today)).Count | Should Be 1
    }

    It 'Start 省略時は実行日の翌日' {
        $p = @(Resolve-Period -End ([datetime]'2026-09-30') -Today $today)
        $p[0].Start | Should Be ([datetime]'2026-09-02')
        $p[0].End | Should Be ([datetime]'2026-09-30')
    }

    It 'End 省略時は開始日の次の金曜' {
        (@(Resolve-Period -Start ([datetime]'2026-09-01') -Today $today))[0].End | Should Be ([datetime]'2026-09-04')
    }

    It '両方省略時は 翌日 と その次の金曜' {
        $p = @(Resolve-Period -Today $today)
        $p[0].Start | Should Be ([datetime]'2026-09-02')
        $p[0].End | Should Be ([datetime]'2026-09-04')
    }

    It '金曜に実行して両方省略すると 開始=土曜・終了=翌週金曜' {
        $p = @(Resolve-Period -Today ([datetime]'2026-09-04'))
        $p[0].Start | Should Be ([datetime]'2026-09-05')
        $p[0].End | Should Be ([datetime]'2026-09-11')
    }

    It '両方指定時はそのまま（時刻は落とす）' {
        $p = @(Resolve-Period -Start ([datetime]'2026-09-01 09:30') -End ([datetime]'2026-09-10 15:00') -Today $today)
        $p[0].Start | Should Be ([datetime]'2026-09-01')
        $p[0].End | Should Be ([datetime]'2026-09-10')
    }
}

Describe 'Resolve-Period -Week' {

    $today = [datetime]'2026-09-01'  # 火 → その週の月曜は 08-31

    It 'Week 0（今週）の開始日は月曜ではなく翌日' {
        $p = @(Resolve-Period -Week 0 -Today $today)
        $p.Count | Should Be 1
        $p[0].Start | Should Be ([datetime]'2026-09-02')  # 翌日（火の翌日＝水）
        $p[0].End | Should Be ([datetime]'2026-09-04')    # その週の金曜
    }

    It 'Week >= 1 の開始日はその週の月曜' {
        (@(Resolve-Period -Week 1 -Today $today))[0].Start | Should Be ([datetime]'2026-09-07')
    }

    It 'Week 2（単一）は 2 週後の月〜金' {
        $p = @(Resolve-Period -Week 2 -Today $today)
        $p[0].Start | Should Be ([datetime]'2026-09-14')
        $p[0].End | Should Be ([datetime]'2026-09-18')
    }

    It 'Week 0 を金曜に実行すると Start > End（逆転はそのまま返す）' {
        $p = @(Resolve-Period -Week 0 -Today ([datetime]'2026-09-04'))
        $p[0].Start | Should Be ([datetime]'2026-09-05')
        $p[0].End | Should Be ([datetime]'2026-09-04')
    }

    It '複数 Week は週ごとの期間を昇順で返す（Week 0 は翌日開始）' {
        $p = @(Resolve-Period -Week 2, 0 -Today $today)
        $p.Count | Should Be 2
        $p[0].Start | Should Be ([datetime]'2026-09-02')
        $p[0].End | Should Be ([datetime]'2026-09-04')
        $p[1].Start | Should Be ([datetime]'2026-09-14')
        $p[1].End | Should Be ([datetime]'2026-09-18')
    }

    It '重複 Week は 1 期間にまとめる' {
        $p = @(Resolve-Period -Week 1, 1, 1 -Today $today)
        $p.Count | Should Be 1
        $p[0].Start | Should Be ([datetime]'2026-09-07')
    }

    It '日曜に実行した場合の基準週（前の月曜起算）' {
        $p = @(Resolve-Period -Week 1 -Today ([datetime]'2026-09-06'))  # 日
        $p[0].Start | Should Be ([datetime]'2026-09-07')
        $p[0].End | Should Be ([datetime]'2026-09-11')
    }

    It 'Week 指定時は Start / End 引数を無視する' {
        $p = @(Resolve-Period -Start ([datetime]'2020-01-01') -End ([datetime]'2020-01-02') -Week 1 -Today $today)
        $p[0].Start | Should Be ([datetime]'2026-09-07')
        $p[0].End | Should Be ([datetime]'2026-09-11')
    }
}
