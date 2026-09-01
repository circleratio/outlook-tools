function ConvertFrom-CalendarSearchEvent {
    <#
    .SYNOPSIS
    Claude の Microsoft 365 コネクタ（outlook_calendar_search）が返す予定オブジェクトの配列を、
    「埋まっている」区間だけの [pscustomobject] @{ Start; End } 配列へ変換する。
    戻り値の形式は Get-CalendarEvents（Graph 版）と同一。

    .DESCRIPTION
    採用条件（すべて満たすもののみ）:
    - showAs が busy / oof / tentative（大文字小文字は無視）。free / workingElsewhere は除外。
    - isCancelled が true でない。

    時刻は start.dateTime / end.dateTime を JST 壁時計（Kind=Unspecified）として解釈する。
    timeZone の値は検証しないが、'Tokyo Standard Time' 以外なら warning を出す（将来の変換拡張点）。
    isAllDay の終日予定も条件に合致すればそのまま区間として扱う（00:00〜翌 00:00 で返るため
    その日の営業時間全体を覆い、結果的に空き行が出力されない）。
    End <= Start（New-Interval が $null）や日時パース不能の要素は除外する。
    入力が空・$null の場合は空配列を返す。並び順は入力のまま（呼び出し側でマージ）。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object[]]$Event
    )

    function Get-Prop {
        param($Object, [string]$Name)
        if ($null -eq $Object) { return $null }
        if ($Object -is [System.Collections.IDictionary]) {
            if ($Object.Contains($Name)) { return $Object[$Name] }
            return $null
        }
        $p = $Object.PSObject.Properties[$Name]
        if ($p) { return $p.Value }
        return $null
    }

    $busyShowAs = @('busy', 'oof', 'tentative')
    $ci   = [System.Globalization.CultureInfo]::InvariantCulture
    $none = [System.Globalization.DateTimeStyles]::None
    $result = New-Object System.Collections.Generic.List[pscustomobject]

    foreach ($ev in @($Event)) {
        if ($null -eq $ev) { continue }

        $showAs = [string](Get-Prop $ev 'showAs')
        if ($busyShowAs -notcontains $showAs.ToLowerInvariant()) { continue }

        $isCancelled = Get-Prop $ev 'isCancelled'
        if ($isCancelled -is [bool] -and $isCancelled) { continue }
        if ("$isCancelled" -eq 'True') { continue }

        $start = Get-Prop $ev 'start'
        $end   = Get-Prop $ev 'end'
        $sText = [string](Get-Prop $start 'dateTime')
        $eText = [string](Get-Prop $end   'dateTime')
        if (-not $sText -or -not $eText) { continue }

        $sTz = [string](Get-Prop $start 'timeZone')
        if ($sTz -and $sTz -ne 'Tokyo Standard Time') {
            Write-Warning "予定のタイムゾーンが 'Tokyo Standard Time' ではありません ('$sTz')。JST 壁時計として解釈します。"
        }

        $s = [datetime]::MinValue
        $e = [datetime]::MinValue
        if (-not [datetime]::TryParse($sText, $ci, $none, [ref]$s)) { continue }
        if (-not [datetime]::TryParse($eText, $ci, $none, [ref]$e)) { continue }

        $iv = New-Interval -Start $s -End $e
        if ($iv) { $result.Add($iv) }
    }

    $result.ToArray()
}
