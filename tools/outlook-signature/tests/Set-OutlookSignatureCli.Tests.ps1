BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    $script:Cli = Join-Path $PSScriptRoot '..\scripts\Set-OutlookSignatureCli.ps1'
    $script:TestKeyBase = "HKCU:\Software\__OLSIG_TEST__"

    function Invoke-Cli {
        param([string[]] $CliArgs)
        $psExe = (Get-Process -Id $PID).Path
        & $psExe -NoProfile -NonInteractive -File $script:Cli @CliArgs *> $null
        return $LASTEXITCODE
    }
}

AfterAll {
    if (Test-Path $script:TestKeyBase) { Remove-Item $script:TestKeyBase -Recurse -Force }
}

Describe 'Set-OutlookSignatureCli.ps1' {
    BeforeEach {
        $script:Root = Join-Path $script:TestKeyBase ("{0}\16.0" -f [guid]::NewGuid().ToString('N'))
        $script:Sp   = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    }

    It '正常時（-Name 指定）は exit 0' {
        New-FakeOutlookRegistry -RegistryRoot $script:Root -DisableRoamingToggle 1 -Accounts @{
            '00000002' = @{ ServiceName = 'MSEMS'; AccountName = 'a@b.c' }
        } | Out-Null
        Invoke-Cli @('-Name', 's', '-Text', 'hi', '-Force', '-SignaturesPath', $script:Sp, '-RegistryRoot', $script:Root) |
            Should -Be 0
    }

    It '正常時（-Name 省略, in-place）は exit 0' {
        New-FakeOutlookRegistry -RegistryRoot $script:Root -DisableRoamingToggle 1 -Accounts @{
            '00000002' = @{ ServiceName = 'MSEMS'; AccountName = 'a@b.c'; New = 'cur'; Reply = 'cur' }
        } | Out-Null
        Invoke-Cli @('-Text', 'hi', '-Force', '-SignaturesPath', $script:Sp, '-RegistryRoot', $script:Root) |
            Should -Be 0
    }

    It 'プリフライト失敗（NoMailAccount）は exit 2' {
        New-FakeOutlookRegistry -RegistryRoot $script:Root -Accounts @{ '1' = @{ ServiceName = 'CONTAB' } } | Out-Null
        Invoke-Cli @('-Name', 's', '-Text', 'hi', '-Force', '-SignaturesPath', $script:Sp, '-RegistryRoot', $script:Root) |
            Should -Be 2
    }

    It 'NoCurrentSignature は exit 2' {
        New-FakeOutlookRegistry -RegistryRoot $script:Root -DisableRoamingToggle 1 -Accounts @{
            '00000002' = @{ ServiceName = 'MSEMS'; AccountName = 'a@b.c' }
        } | Out-Null
        Invoke-Cli @('-Text', 'hi', '-Force', '-SignaturesPath', $script:Sp, '-RegistryRoot', $script:Root) |
            Should -Be 2
    }

    It 'Outlook 起動中相当（-Force 無し・実プロセス）で exit 3 になりうる' -Skip:(-not (Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)) {
        New-FakeOutlookRegistry -RegistryRoot $script:Root -DisableRoamingToggle 1 -Accounts @{
            '00000002' = @{ ServiceName = 'MSEMS'; AccountName = 'a@b.c' }
        } | Out-Null
        Invoke-Cli @('-Name', 's', '-Text', 'hi', '-SignaturesPath', $script:Sp, '-RegistryRoot', $script:Root) |
            Should -Be 3
    }
}
