function Write-SignatureFileSet {
    <#
    .SYNOPSIS
        署名のプレーンテキストから <Name>.htm (UTF-8 no BOM) と <Name>.txt (UTF-16LE BOM) を生成する。
    .OUTPUTS
        生成したファイルの絶対パス（string[]）。
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory)]
        [string] $SignaturesPath,

        [Parameter(Mandatory)]
        [string] $Name,

        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string] $Text
    )

    if (-not (Test-Path -LiteralPath $SignaturesPath)) {
        New-Item -ItemType Directory -Path $SignaturesPath -Force | Out-Null
    }

    $htmPath = Join-Path $SignaturesPath ($Name + '.htm')
    $txtPath = Join-Path $SignaturesPath ($Name + '.txt')

    $htmContent = ConvertTo-SignatureHtml -Text $Text

    # .txt は CRLF に正規化
    $txtContent = ($Text -replace "`r`n", "`n" -replace "`r", "`n") -replace "`n", "`r`n"

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($htmPath, $htmContent, $utf8NoBom)
    [System.IO.File]::WriteAllText($txtPath, $txtContent, [System.Text.Encoding]::Unicode)

    @(
        (Resolve-Path -LiteralPath $htmPath).Path
        (Resolve-Path -LiteralPath $txtPath).Path
    )
}
