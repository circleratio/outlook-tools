BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    Import-TargetModule
    $script:TestKeyBase = "HKCU:\Software\__OLSIG_TEST__"
}

AfterAll {
    if (Test-Path $script:TestKeyBase) { Remove-Item $script:TestKeyBase -Recurse -Force }
}

Describe 'Resolve-OutlookEnvironment' {
    BeforeEach {
        $script:Root = Join-Path $script:TestKeyBase ("{0}\16.0" -f [guid]::NewGuid().ToString('N'))
    }

    It 'Outlook キーが無ければ OutlookNotFound' {
        $root = $script:Root
        New-Item -Path $root -Force | Out-Null
        InModuleScope OutlookSignature -Parameters @{ root = $root } {
            param($root)
            { Resolve-OutlookEnvironment -RegistryRoot $root } | Should -Throw -ExpectedMessage '*OutlookNotFound*'
        }
    }

    It 'DefaultProfile が無ければ DefaultProfileNotFound' {
        $root = $script:Root
        New-FakeOutlookRegistry -RegistryRoot $root -NoDefaultProfileValue | Out-Null
        InModuleScope OutlookSignature -Parameters @{ root = $root } {
            param($root)
            { Resolve-OutlookEnvironment -RegistryRoot $root } | Should -Throw -ExpectedMessage '*DefaultProfileNotFound*'
        }
    }

    It 'PickLogonProfile=1 なら ProfilePromptEnabled' {
        $root = $script:Root
        New-FakeOutlookRegistry -RegistryRoot $root -PickLogonProfile 1 -Accounts @{ '1' = @{ ServiceName = 'MSEMS' } } | Out-Null
        InModuleScope OutlookSignature -Parameters @{ root = $root } {
            param($root)
            { Resolve-OutlookEnvironment -RegistryRoot $root } | Should -Throw -ExpectedMessage '*ProfilePromptEnabled*'
        }
    }

    It '正常時に各パスが期待どおり' {
        $root = $script:Root
        New-FakeOutlookRegistry -RegistryRoot $root -DefaultProfile 'MyProf' -Accounts @{ '1' = @{ ServiceName = 'MSEMS' } } | Out-Null
        InModuleScope OutlookSignature -Parameters @{ root = $root } {
            param($root)
            $r = Resolve-OutlookEnvironment -RegistryRoot $root -SignaturesPath 'C:\sig'
            $r.DefaultProfile  | Should -Be 'MyProf'
            $r.ProfileKeyPath  | Should -BeLike '*Profiles\MyProf'
            $r.AccountsKeyPath | Should -BeLike '*9375CFF0413111d3B88A00104B2A6676'
            $r.SignaturesPath  | Should -Be 'C:\sig'
        }
    }
}
