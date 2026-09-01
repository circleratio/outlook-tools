function Set-AccountSignatureRegistry {
    <#
    .SYNOPSIS
        指定アカウントキーに New Signature / Reply-Forward Signature を REG_SZ で設定する。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string] $AccountKeyPath,

        [Parameter(Mandatory)]
        [string] $SignatureName
    )

    if (-not (Test-Path -LiteralPath $AccountKeyPath)) {
        throw "アカウントキーが見つかりません: $AccountKeyPath"
    }

    Set-ItemProperty -LiteralPath $AccountKeyPath -Name 'New Signature'           -Value $SignatureName -Type String
    Set-ItemProperty -LiteralPath $AccountKeyPath -Name 'Reply-Forward Signature' -Value $SignatureName -Type String
}
