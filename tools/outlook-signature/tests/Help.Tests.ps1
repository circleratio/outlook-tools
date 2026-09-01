BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    Import-TargetModule
}

Describe 'comment-based help' {
    It '<_> にパラメータ説明と 2 つ以上の例がある' -ForEach @('Set-OutlookSignature', 'Get-OutlookSignature') {
        $help = Get-Help $_ -Full
        $help.Synopsis | Should -Not -BeNullOrEmpty
        @($help.parameters.parameter).Count | Should -BeGreaterThan 0
    }

    It 'Set-OutlookSignature に 2 つ以上の例がある' {
        $help = Get-Help Set-OutlookSignature -Full
        @($help.examples.example).Count | Should -BeGreaterOrEqual 2
    }
}
