BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    Import-TargetModule
    $script:TestKeyBase = "TestRegistry:"
}

Describe 'Set-OutlookSignature' {
    BeforeEach {
        $script:Root = Join-Path $script:TestKeyBase ("{0}\16.0" -f [guid]::NewGuid().ToString('N'))
        $script:Sp   = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $script:AcctKey = Join-Path $script:Root 'Outlook\Profiles\Outlook\9375CFF0413111d3B88A00104B2A6676\00000002'
        Mock -ModuleName OutlookSignature Test-OutlookRunning { $false }
    }

    Context 'CreateAndAssign（-Name 指定）' {
        It '全アカウントに両署名値を設定し、ファイルを生成する' {
            New-FakeOutlookRegistry -RegistryRoot $script:Root -DisableRoamingToggle 1 -Accounts @{
                '00000001' = @{ ServiceName = 'CONTAB' }
                '00000002' = @{ ServiceName = 'MSEMS'; AccountName = 'a@b.c' }
                '00000004' = @{ ServiceName = 'IMAP'; AccountName = 'c@d.e' }
            } | Out-Null

            $r = Set-OutlookSignature -Name '標準' -Text "本文`n2行目" -SignaturesPath $script:Sp -RegistryRoot $script:Root

            $r.Mode | Should -Be 'CreateAndAssign'
            $r.SignatureNames | Should -Be @('標準')
            $r.RegistryChanged | Should -BeTrue
            $r.RestartRequired | Should -BeTrue
            $r.Files.Count | Should -Be 2
            $r.Accounts.Count | Should -Be 2

            Test-Path (Join-Path $script:Sp '標準.htm') | Should -BeTrue
            (Get-ItemProperty -Path $script:AcctKey -Name 'New Signature').'New Signature' | Should -Be '標準'
            (Get-ItemProperty -Path $script:AcctKey -Name 'Reply-Forward Signature').'Reply-Forward Signature' | Should -Be '標準'
        }

        It '既存署名がある場合はバックアップして置き換える' {
            New-FakeOutlookRegistry -RegistryRoot $script:Root -DisableRoamingToggle 1 -Accounts @{
                '00000002' = @{ ServiceName = 'MSEMS'; AccountName = 'a@b.c' }
            } | Out-Null
            New-Item -ItemType Directory -Path $script:Sp -Force | Out-Null
            Set-Content (Join-Path $script:Sp '標準.htm') 'OLD'

            $r = Set-OutlookSignature -Name '標準' -Text 'NEW' -SignaturesPath $script:Sp -RegistryRoot $script:Root

            $r.BackupPaths.Count | Should -Be 1
            [System.IO.File]::ReadAllText((Join-Path $script:Sp '標準.htm'), [System.Text.Encoding]::UTF8) | Should -Match 'NEW'
            (Get-ChildItem (Join-Path $script:Sp '.backup') -Recurse -Filter '標準.htm').Count | Should -Be 1
        }

        It '-WhatIf ではファイルもレジストリも変更しない' {
            New-FakeOutlookRegistry -RegistryRoot $script:Root -DisableRoamingToggle 1 -Accounts @{
                '00000002' = @{ ServiceName = 'MSEMS'; AccountName = 'a@b.c' }
            } | Out-Null

            $r = Set-OutlookSignature -Name 'x' -Text 'y' -SignaturesPath $script:Sp -RegistryRoot $script:Root -WhatIf

            $r.Files.Count | Should -Be 0
            $r.RegistryChanged | Should -BeFalse
            Test-Path (Join-Path $script:Sp 'x.htm') | Should -BeFalse
            (Get-ItemProperty -Path $script:AcctKey -Name 'New Signature' -ErrorAction SilentlyContinue) | Should -BeNullOrEmpty
        }
    }

    Context 'UpdateInPlace（-Name 省略）' {
        It '現行署名のファイルのみ更新し、レジストリは変更しない' {
            New-FakeOutlookRegistry -RegistryRoot $script:Root -DisableRoamingToggle 1 -Accounts @{
                '00000002' = @{ ServiceName = 'MSEMS'; AccountName = 'a@b.c'; New = 'default (a@b.c)'; Reply = 'default (a@b.c)' }
            } | Out-Null

            $r = Set-OutlookSignature -Text '新しい本文' -SignaturesPath $script:Sp -RegistryRoot $script:Root

            $r.Mode | Should -Be 'UpdateInPlace'
            $r.SignatureNames | Should -Be @('default (a@b.c)')
            $r.RegistryChanged | Should -BeFalse
            Test-Path (Join-Path $script:Sp 'default (a@b.c).htm') | Should -BeTrue
        }

        It '現行署名が全アカウント空なら NoCurrentSignature' {
            New-FakeOutlookRegistry -RegistryRoot $script:Root -DisableRoamingToggle 1 -Accounts @{
                '00000002' = @{ ServiceName = 'MSEMS'; AccountName = 'a@b.c' }
            } | Out-Null

            Assert-ThrowsErrorId -ErrorId 'NoCurrentSignature' -ScriptBlock {
                Set-OutlookSignature -Text 'x' -SignaturesPath $script:Sp -RegistryRoot $script:Root
            }
        }
    }

    Context 'プリフライト検証' {
        It 'アカウント 0 件で NoMailAccount' {
            New-FakeOutlookRegistry -RegistryRoot $script:Root -Accounts @{ '1' = @{ ServiceName = 'CONTAB' } } | Out-Null
            Assert-ThrowsErrorId -ErrorId 'NoMailAccount' -ScriptBlock {
                Set-OutlookSignature -Name 's' -Text 'x' -SignaturesPath $script:Sp -RegistryRoot $script:Root
            }
        }

        It '無効文字を含む署名名で InvalidSignatureName' {
            New-FakeOutlookRegistry -RegistryRoot $script:Root -Accounts @{ '2' = @{ ServiceName = 'MSEMS' } } | Out-Null
            Assert-ThrowsErrorId -ErrorId 'InvalidSignatureName' -ScriptBlock {
                Set-OutlookSignature -Name 'a<b>c' -Text 'x' -SignaturesPath $script:Sp -RegistryRoot $script:Root
            }
        }

        It 'DefaultProfile 未設定で DefaultProfileNotFound' {
            New-FakeOutlookRegistry -RegistryRoot $script:Root -NoDefaultProfileValue | Out-Null
            Assert-ThrowsErrorId -ErrorId 'DefaultProfileNotFound' -ScriptBlock {
                Set-OutlookSignature -Name 's' -Text 'x' -SignaturesPath $script:Sp -RegistryRoot $script:Root
            }
        }
    }

    Context 'Outlook 起動中' {
        It '-Force 無しで OutlookRunning' {
            New-FakeOutlookRegistry -RegistryRoot $script:Root -Accounts @{ '2' = @{ ServiceName = 'MSEMS' } } | Out-Null
            Mock -ModuleName OutlookSignature Test-OutlookRunning { $true }
            Assert-ThrowsErrorId -ErrorId 'OutlookRunning' -ScriptBlock {
                Set-OutlookSignature -Name 's' -Text 'x' -SignaturesPath $script:Sp -RegistryRoot $script:Root
            }
        }

        It '-Force 有りで警告つき成功' {
            New-FakeOutlookRegistry -RegistryRoot $script:Root -DisableRoamingToggle 1 -Accounts @{
                '00000002' = @{ ServiceName = 'MSEMS'; AccountName = 'a@b.c' }
            } | Out-Null
            Mock -ModuleName OutlookSignature Test-OutlookRunning { $true }
            $r = Set-OutlookSignature -Name 's' -Text 'x' -SignaturesPath $script:Sp -RegistryRoot $script:Root -Force -WarningAction SilentlyContinue
            $r.Warnings | Should -Match 'Outlook が起動中'
            $r.Files.Count | Should -Be 2
        }
    }

    Context 'その他' {
        It '空 -Text でも成功する' {
            New-FakeOutlookRegistry -RegistryRoot $script:Root -DisableRoamingToggle 1 -Accounts @{
                '00000002' = @{ ServiceName = 'MSEMS'; AccountName = 'a@b.c' }
            } | Out-Null
            $r = Set-OutlookSignature -Name 's' -Text '' -SignaturesPath $script:Sp -RegistryRoot $script:Root
            $r.Files.Count | Should -Be 2
        }

        It '署名名 33 文字は警告つき成功' {
            New-FakeOutlookRegistry -RegistryRoot $script:Root -DisableRoamingToggle 1 -Accounts @{
                '00000002' = @{ ServiceName = 'MSEMS'; AccountName = 'a@b.c' }
            } | Out-Null
            $r = Set-OutlookSignature -Name ('a' * 33) -Text 'x' -SignaturesPath $script:Sp -RegistryRoot $script:Root -WarningAction SilentlyContinue
            $r.Warnings | Should -Match '32 文字'
        }

        It 'RoamingSignaturesEnabled が判定結果と一致する' {
            New-FakeOutlookRegistry -RegistryRoot $script:Root -Accounts @{
                '00000002' = @{ ServiceName = 'MSEMS'; AccountName = 'a@b.c' }
            } | Out-Null
            $r = Set-OutlookSignature -Name 's' -Text 'x' -SignaturesPath $script:Sp -RegistryRoot $script:Root -WarningAction SilentlyContinue
            $r.RoamingSignaturesEnabled | Should -BeTrue
        }
    }
}
