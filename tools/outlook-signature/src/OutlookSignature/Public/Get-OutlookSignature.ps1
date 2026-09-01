function Get-OutlookSignature {
    <#
    .SYNOPSIS
        Signatures フォルダの署名一覧と、アカウントごとの現在の既定署名を返す。

    .PARAMETER SignaturesPath
        Signatures フォルダのパス（テスト／上級者向け上書き）。

    .PARAMETER RegistryRoot
        Office のレジストリルート（テスト／上級者向け上書き）。

    .EXAMPLE
        Get-OutlookSignature
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [string] $SignaturesPath,

        [string] $RegistryRoot = 'HKCU:\Software\Microsoft\Office\16.0'
    )

    $env = Resolve-OutlookEnvironment -RegistryRoot $RegistryRoot -SignaturesPath $SignaturesPath

    $signatureNames = @()
    if (Test-Path -LiteralPath $env.SignaturesPath) {
        $signatureNames = @(
            Get-ChildItem -LiteralPath $env.SignaturesPath -Filter '*.htm' -ErrorAction SilentlyContinue |
                ForEach-Object { [System.IO.Path]::GetFileNameWithoutExtension($_.Name) } |
                Sort-Object -Unique
        )
    }

    $accounts = @(
        Get-OutlookMailAccount -AccountsKeyPath $env.AccountsKeyPath | ForEach-Object {
            [pscustomobject]@{
                AccountName          = $_.AccountName
                ServiceName          = $_.ServiceName
                NewSignature         = $_.PreviousNew
                ReplyForwardSignature = $_.PreviousReplyForward
            }
        }
    )

    [pscustomobject]@{
        SignaturesPath           = $env.SignaturesPath
        DefaultProfile           = $env.DefaultProfile
        Signatures               = $signatureNames
        Accounts                 = $accounts
        RoamingSignaturesEnabled = (Test-RoamingSignatureEnabled -RegistryRoot $RegistryRoot)
    }
}
