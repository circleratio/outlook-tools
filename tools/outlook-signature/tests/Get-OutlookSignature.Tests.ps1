BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    Import-TargetModule
    $script:TestKeyBase = "TestRegistry:"
}

Describe 'Get-OutlookSignature' {
    BeforeEach {
        $script:Root = Join-Path $script:TestKeyBase ("{0}\16.0" -f [guid]::NewGuid().ToString('N'))
        $script:Sp   = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    }

    It '署名ファイル一覧とアカウント別の現行値を返す' {
        $root = $script:Root; $sp = $script:Sp
        New-FakeOutlookRegistry -RegistryRoot $root -Accounts @{
            '00000002' = @{ ServiceName = 'MSEMS'; AccountName = 'a@b.c'; New = 'sigA'; Reply = 'sigB' }
        } | Out-Null
        New-Item -ItemType Directory -Path $sp -Force | Out-Null
        Set-Content (Join-Path $sp 'sigA.htm') 'x'
        Set-Content (Join-Path $sp 'English.htm') 'x'

        $r = InModuleScope OutlookSignature -Parameters @{ root = $root; sp = $sp } {
            param($root, $sp)
            Get-OutlookSignature -SignaturesPath $sp -RegistryRoot $root
        }

        $r.Signatures | Should -Contain 'sigA'
        $r.Signatures | Should -Contain 'English'
        $r.Accounts.Count | Should -Be 1
        $r.Accounts[0].NewSignature | Should -Be 'sigA'
        $r.Accounts[0].ReplyForwardSignature | Should -Be 'sigB'
    }

    It 'Set-OutlookSignature 実行直後に設定した署名名が全アカウントで返る' {
        $root = $script:Root; $sp = $script:Sp
        New-FakeOutlookRegistry -RegistryRoot $root -DisableRoamingToggle 1 -Accounts @{
            '00000002' = @{ ServiceName = 'MSEMS'; AccountName = 'a@b.c' }
            '00000004' = @{ ServiceName = 'IMAP'; AccountName = 'c@d.e' }
        } | Out-Null

        $r = InModuleScope OutlookSignature -Parameters @{ root = $root; sp = $sp } {
            param($root, $sp)
            Mock Test-OutlookRunning { $false }
            Set-OutlookSignature -Name '標準' -Text 'x' -SignaturesPath $sp -RegistryRoot $root | Out-Null
            Get-OutlookSignature -SignaturesPath $sp -RegistryRoot $root
        }

        $r.Accounts.NewSignature | Should -Not -Contain ''
        ($r.Accounts | Where-Object { $_.NewSignature -eq '標準' }).Count | Should -Be 2
    }
}
