# テスト共通ヘルパー。各 *.Tests.ps1 の BeforeAll から dot-source する。

$script:ModulePath = Join-Path $PSScriptRoot '..\src\OutlookSignature\OutlookSignature.psd1'

function Import-TargetModule {
    Import-Module $script:ModulePath -Force
}

# 指定した scriptblock がその ErrorId（FullyQualifiedErrorId の先頭）で終了エラーを投げることを検証する。
function Assert-ThrowsErrorId {
    param(
        [Parameter(Mandatory)] [scriptblock] $ScriptBlock,
        [Parameter(Mandatory)] [string] $ErrorId
    )
    $err = $null
    try {
        & $ScriptBlock | Out-Null
    }
    catch {
        $err = $_
    }
    if ($null -eq $err) {
        throw "ErrorId '$ErrorId' の終了エラーを期待しましたが、例外は投げられませんでした。"
    }
    if ([string]$err.FullyQualifiedErrorId -notlike "$ErrorId*") {
        throw "ErrorId '$ErrorId*' を期待しましたが、実際は '$($err.FullyQualifiedErrorId)' でした。"
    }
}

# 疑似 Outlook レジストリツリーを作成する。
#   -RegistryRoot  例: 'TestRegistry:\Office\16.0'
#   -Accounts      @{ '00000002' = @{ ServiceName='MSEMS'; AccountName='a@b.c'; New='sig1'; Reply='sig1' } }
function New-FakeOutlookRegistry {
    param(
        [Parameter(Mandatory)] [string] $RegistryRoot,
        [string] $DefaultProfile = 'Outlook',
        [hashtable] $Accounts = @{},
        [switch] $NoDefaultProfileValue,
        [switch] $NoOutlookKey,
        [int] $PickLogonProfile = -1,
        [int] $DisableRoamingToggle = -1
    )

    if (-not $NoOutlookKey) {
        $outlookKey = Join-Path $RegistryRoot 'Outlook'
        New-Item -Path $outlookKey -Force | Out-Null

        if (-not $NoDefaultProfileValue) {
            New-ItemProperty -Path $outlookKey -Name 'DefaultProfile' -Value $DefaultProfile -PropertyType String -Force | Out-Null
        }
        if ($PickLogonProfile -ge 0) {
            New-ItemProperty -Path $outlookKey -Name 'PickLogonProfile' -Value $PickLogonProfile -PropertyType DWord -Force | Out-Null
        }
        if ($DisableRoamingToggle -ge 0) {
            $setupKey = Join-Path $RegistryRoot 'Outlook\Setup'
            New-Item -Path $setupKey -Force | Out-Null
            New-ItemProperty -Path $setupKey -Name 'DisableRoamingSignaturesTemporaryToggle' -Value $DisableRoamingToggle -PropertyType DWord -Force | Out-Null
        }

        $accountsKey = Join-Path $outlookKey ("Profiles\{0}\9375CFF0413111d3B88A00104B2A6676" -f $DefaultProfile)
        if ($Accounts.Count -gt 0) {
            New-Item -Path $accountsKey -Force | Out-Null
        }
        foreach ($id in $Accounts.Keys) {
            $a = $Accounts[$id]
            $k = Join-Path $accountsKey $id
            New-Item -Path $k -Force | Out-Null
            if ($a.ContainsKey('ServiceName') -and $a.ServiceName) {
                New-ItemProperty -Path $k -Name 'Service Name' -Value $a.ServiceName -PropertyType String -Force | Out-Null
            }
            if ($a.ContainsKey('AccountName')) {
                New-ItemProperty -Path $k -Name 'Account Name' -Value $a.AccountName -PropertyType String -Force | Out-Null
            }
            if ($a.ContainsKey('New')) {
                New-ItemProperty -Path $k -Name 'New Signature' -Value $a.New -PropertyType String -Force | Out-Null
            }
            if ($a.ContainsKey('Reply')) {
                New-ItemProperty -Path $k -Name 'Reply-Forward Signature' -Value $a.Reply -PropertyType String -Force | Out-Null
            }
        }
    }

    Join-Path $RegistryRoot 'Outlook'
}
