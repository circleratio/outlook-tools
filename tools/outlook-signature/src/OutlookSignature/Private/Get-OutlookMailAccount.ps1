function Get-OutlookMailAccount {
    <#
    .SYNOPSIS
        アカウントキー配下から、メール送信系アカウントのみを列挙する。
    .OUTPUTS
        AccountName / ServiceName / RegistryPath / PreviousNew / PreviousReplyForward を持つ
        pscustomobject の配列。
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject[]])]
    param(
        [Parameter(Mandatory)]
        [string] $AccountsKeyPath
    )

    # メール送信系サービス名（アドレス帳 CONTAB / PST 等は除外）
    $mailServices = @('MSEMS', 'IMAP', 'POP3', 'SMTP', 'MAPI', 'EAS', 'EXHTTP', 'EXPOP', 'EXIMAP')

    if (-not (Test-Path -LiteralPath $AccountsKeyPath)) {
        return @()
    }

    $result = foreach ($sub in Get-ChildItem -LiteralPath $AccountsKeyPath -ErrorAction SilentlyContinue) {
        $serviceName = Get-RegistryValue -Path $sub.PSPath -Name 'Service Name'
        if ([string]::IsNullOrWhiteSpace($serviceName)) { continue }
        if ($mailServices -notcontains $serviceName) { continue }

        [pscustomobject]@{
            AccountName          = [string](Get-RegistryValue -Path $sub.PSPath -Name 'Account Name')
            ServiceName          = [string]$serviceName
            RegistryPath         = $sub.PSPath
            PreviousNew          = [string](Get-RegistryValue -Path $sub.PSPath -Name 'New Signature')
            PreviousReplyForward = [string](Get-RegistryValue -Path $sub.PSPath -Name 'Reply-Forward Signature')
        }
    }

    @($result)
}
