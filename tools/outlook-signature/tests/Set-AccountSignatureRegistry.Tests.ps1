BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    Import-TargetModule
    $script:TestKeyBase = "HKCU:\Software\__OLSIG_TEST__"
}

AfterAll {
    if (Test-Path $script:TestKeyBase) { Remove-Item $script:TestKeyBase -Recurse -Force }
}

Describe 'Set-AccountSignatureRegistry' {
    BeforeEach {
        $script:Key = Join-Path $script:TestKeyBase ("{0}\acct" -f [guid]::NewGuid().ToString('N'))
        New-Item -Path $script:Key -Force | Out-Null
    }

    It 'New Signature / Reply-Forward Signature を REG_SZ で設定する' {
        $key = $script:Key
        InModuleScope OutlookSignature -Parameters @{ key = $key } {
            param($key)
            Set-AccountSignatureRegistry -AccountKeyPath $key -SignatureName '標準'
        }
        (Get-ItemProperty -Path $key -Name 'New Signature').'New Signature' | Should -Be '標準'
        (Get-ItemProperty -Path $key -Name 'Reply-Forward Signature').'Reply-Forward Signature' | Should -Be '標準'
        $sub = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey(($key -replace '^HKCU:\\', ''))
        $sub.GetValueKind('New Signature') | Should -Be ([Microsoft.Win32.RegistryValueKind]::String)
    }

    It '既存値を上書きする' {
        $key = $script:Key
        New-ItemProperty -Path $key -Name 'New Signature' -Value 'old' -PropertyType String -Force | Out-Null
        InModuleScope OutlookSignature -Parameters @{ key = $key } {
            param($key)
            Set-AccountSignatureRegistry -AccountKeyPath $key -SignatureName 'new'
        }
        (Get-ItemProperty -Path $key -Name 'New Signature').'New Signature' | Should -Be 'new'
    }
}
