function Resolve-OutlookEnvironment {
    <#
    .SYNOPSIS
        レジストリルートと（任意の）Signatures パスから Outlook 環境情報を解決する。
    .OUTPUTS
        RegistryRoot / OutlookKeyPath / DefaultProfile / ProfileKeyPath / AccountsKeyPath / SignaturesPath
        を持つ pscustomobject。
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [string] $RegistryRoot,

        [string] $SignaturesPath
    )

    $outlookKey = Join-Path $RegistryRoot 'Outlook'
    if (-not (Test-Path -LiteralPath $outlookKey)) {
        throw "Outlook のレジストリキーが見つかりません: $outlookKey [OutlookNotFound]"
    }

    $pickLogon = Get-RegistryValue -Path $outlookKey -Name 'PickLogonProfile'
    if ($pickLogon -eq 1) {
        throw "起動時にプロファイル選択が有効です（PickLogonProfile=1）。対象外です。 [ProfilePromptEnabled]"
    }

    $defaultProfile = Get-RegistryValue -Path $outlookKey -Name 'DefaultProfile'
    if ([string]::IsNullOrWhiteSpace($defaultProfile)) {
        throw "既定プロファイル（DefaultProfile）が設定されていません。 [DefaultProfileNotFound]"
    }

    $profileKey = Join-Path $outlookKey ("Profiles\{0}" -f $defaultProfile)
    if (-not (Test-Path -LiteralPath $profileKey)) {
        throw "既定プロファイルのキーが見つかりません: $profileKey [DefaultProfileNotFound]"
    }

    $accountsKey = Join-Path $profileKey '9375CFF0413111d3B88A00104B2A6676'

    if ([string]::IsNullOrWhiteSpace($SignaturesPath)) {
        $appData = [Environment]::GetFolderPath('ApplicationData')
        $SignaturesPath = Join-Path $appData 'Microsoft\Signatures'
    }

    [pscustomobject]@{
        RegistryRoot    = $RegistryRoot
        OutlookKeyPath  = $outlookKey
        DefaultProfile  = $defaultProfile
        ProfileKeyPath  = $profileKey
        AccountsKeyPath = $accountsKey
        SignaturesPath  = $SignaturesPath
    }
}
