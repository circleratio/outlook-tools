function ConvertTo-InputDate {
    <#
    .SYNOPSIS
    日付文字列を [datetime]（日付のみ）に変換する。

    .DESCRIPTION
    受け付ける書式:
    - `yyyy-MM-dd` … その日付
    - `MM-dd`      … 年を省略した形式。年は Today の年（実行年）を補う
    どちらの書式にも一致しない、または存在しない日付（実行年の 2-29 など）の場合は terminating error。
    Today 未指定時は [datetime]::Today。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Text,
        [datetime]$Today = [datetime]::MinValue
    )

    if ($Today -eq [datetime]::MinValue) { $Today = [datetime]::Today }

    $ci     = [System.Globalization.CultureInfo]::InvariantCulture
    $none   = [System.Globalization.DateTimeStyles]::None
    $parsed = [datetime]::MinValue

    if ([datetime]::TryParseExact($Text, 'yyyy-MM-dd', $ci, $none, [ref]$parsed)) {
        return $parsed.Date
    }

    if ($Text -match '^\d{2}-\d{2}$') {
        $parts = $Text.Split('-')
        try {
            return ([datetime]::new($Today.Year, [int]$parts[0], [int]$parts[1])).Date
        }
        catch {
            throw "日付が存在しません: $Text（$($Today.Year) 年）"
        }
    }

    throw "日付の書式が不正です (yyyy-MM-dd または MM-dd): $Text"
}
