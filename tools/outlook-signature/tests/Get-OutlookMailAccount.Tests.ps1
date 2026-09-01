BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    Import-TargetModule
    $script:TestKeyBase = "HKCU:\Software\__OLSIG_TEST__"
}

AfterAll {
    if (Test-Path $script:TestKeyBase) { Remove-Item $script:TestKeyBase -Recurse -Force }
}

Describe 'Get-OutlookMailAccount' {
    BeforeEach {
        $script:Root = Join-Path $script:TestKeyBase ("{0}\16.0" -f [guid]::NewGuid().ToString('N'))
    }

    It 'メール送信系のみ抽出し、CONTAB（アドレス帳）は除外する' {
        $root = $script:Root
        New-FakeOutlookRegistry -RegistryRoot $root -Accounts @{
            '00000001' = @{ ServiceName = 'CONTAB'; AccountName = 'アドレス帳' }
            '00000002' = @{ ServiceName = 'MSEMS'; AccountName = 'a@b.c'; New = 'sig1'; Reply = 'sig2' }
            '00000003' = @{ ServiceName = 'MSPST MS'; AccountName = 'PST' }
        } | Out-Null

        $accountsKey = Join-Path $root 'Outlook\Profiles\Outlook\9375CFF0413111d3B88A00104B2A6676'
        $r = InModuleScope OutlookSignature -Parameters @{ k = $accountsKey } {
            param($k)
            Get-OutlookMailAccount -AccountsKeyPath $k
        }
        @($r).Count | Should -Be 1
        $r.AccountName | Should -Be 'a@b.c'
        $r.ServiceName | Should -Be 'MSEMS'
        $r.PreviousNew | Should -Be 'sig1'
        $r.PreviousReplyForward | Should -Be 'sig2'
    }

    It '現行署名が未設定なら PreviousNew は空文字' {
        $root = $script:Root
        New-FakeOutlookRegistry -RegistryRoot $root -Accounts @{
            '00000002' = @{ ServiceName = 'MSEMS'; AccountName = 'a@b.c' }
        } | Out-Null
        $accountsKey = Join-Path $root 'Outlook\Profiles\Outlook\9375CFF0413111d3B88A00104B2A6676'
        $r = InModuleScope OutlookSignature -Parameters @{ k = $accountsKey } {
            param($k)
            Get-OutlookMailAccount -AccountsKeyPath $k
        }
        $r.PreviousNew | Should -Be ''
        $r.PreviousReplyForward | Should -Be ''
    }

    It '該当 0 件で空配列（例外なし）' {
        $root = $script:Root
        New-FakeOutlookRegistry -RegistryRoot $root -Accounts @{
            '00000001' = @{ ServiceName = 'CONTAB' }
        } | Out-Null
        $accountsKey = Join-Path $root 'Outlook\Profiles\Outlook\9375CFF0413111d3B88A00104B2A6676'
        $r = InModuleScope OutlookSignature -Parameters @{ k = $accountsKey } {
            param($k)
            ,(Get-OutlookMailAccount -AccountsKeyPath $k)
        }
        @($r).Count | Should -Be 0
    }

    It 'キーが存在しない場合も空配列' {
        $r = InModuleScope OutlookSignature -Parameters @{ k = 'HKCU:\Software\__OLSIG_TEST__\nope' } {
            param($k)
            ,(Get-OutlookMailAccount -AccountsKeyPath $k)
        }
        @($r).Count | Should -Be 0
    }
}
