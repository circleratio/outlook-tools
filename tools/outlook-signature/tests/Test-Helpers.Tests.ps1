BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    Import-TargetModule
    $script:TestKeyBase = "TestRegistry:"
}

Describe 'Test-RoamingSignatureEnabled' {
    BeforeEach {
        $script:Root = Join-Path $script:TestKeyBase ("{0}\16.0" -f [guid]::NewGuid().ToString('N'))
    }

    It 'トグル=1 なら $false（無効）' {
        $root = $script:Root
        New-FakeOutlookRegistry -RegistryRoot $root -DisableRoamingToggle 1 -Accounts @{ '1' = @{ ServiceName = 'MSEMS' } } | Out-Null
        InModuleScope OutlookSignature -Parameters @{ root = $root } {
            param($root)
            Test-RoamingSignatureEnabled -RegistryRoot $root | Should -BeFalse
        }
    }

    It 'トグル未設定なら $true（有効の疑い）' {
        $root = $script:Root
        New-FakeOutlookRegistry -RegistryRoot $root -Accounts @{ '1' = @{ ServiceName = 'MSEMS' } } | Out-Null
        InModuleScope OutlookSignature -Parameters @{ root = $root } {
            param($root)
            Test-RoamingSignatureEnabled -RegistryRoot $root | Should -BeTrue
        }
    }

    It 'トグル=0 なら $true' {
        $root = $script:Root
        New-FakeOutlookRegistry -RegistryRoot $root -DisableRoamingToggle 0 -Accounts @{ '1' = @{ ServiceName = 'MSEMS' } } | Out-Null
        InModuleScope OutlookSignature -Parameters @{ root = $root } {
            param($root)
            Test-RoamingSignatureEnabled -RegistryRoot $root | Should -BeTrue
        }
    }
}

Describe 'Test-OutlookRunning' {
    It 'Get-Process の結果に応じて bool を返す' {
        InModuleScope OutlookSignature {
            Mock Get-Process { [pscustomobject]@{ Name = 'OUTLOOK' } } -ParameterFilter { $Name -eq 'OUTLOOK' }
            Test-OutlookRunning | Should -BeTrue

            Mock Get-Process { $null } -ParameterFilter { $Name -eq 'OUTLOOK' }
            Test-OutlookRunning | Should -BeFalse
        }
    }
}
