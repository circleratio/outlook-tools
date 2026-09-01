function ConvertTo-SignatureHtml {
    <#
    .SYNOPSIS
        プレーンテキストの署名本文を、Outlook 用の最小 HTML 署名に変換する。
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string] $Text
    )

    $escaped = [System.Net.WebUtility]::HtmlEncode($Text)

    # 改行を LF に正規化してから <br> + CRLF に変換
    $normalized = $escaped -replace "`r`n", "`n" -replace "`r", "`n"
    $body = $normalized -replace "`n", "<br>`r`n"

    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine('<!DOCTYPE html>')
    [void]$sb.AppendLine('<html xmlns="http://www.w3.org/1999/xhtml">')
    [void]$sb.AppendLine('<head>')
    [void]$sb.AppendLine('<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />')
    [void]$sb.AppendLine('</head>')
    [void]$sb.AppendLine('<body>')
    [void]$sb.AppendLine("<div>$body</div>")
    [void]$sb.AppendLine('</body>')
    [void]$sb.AppendLine('</html>')

    # 出力全体を CRLF 改行に統一
    ($sb.ToString() -replace "`r`n", "`n" -replace "`r", "`n") -replace "`n", "`r`n"
}
