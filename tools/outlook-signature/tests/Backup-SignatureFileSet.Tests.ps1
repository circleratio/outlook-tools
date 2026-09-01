BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    Import-TargetModule
}

Describe 'Backup-SignatureFileSet' {
    It '既存ファイル群を .backup へ移動し、パスを返す' {
        $dir = Join-Path $TestDrive 'b1'
        New-Item -ItemType Directory -Path $dir | Out-Null
        Set-Content (Join-Path $dir 'sig.htm') 'h'
        Set-Content (Join-Path $dir 'sig.txt') 't'
        Set-Content (Join-Path $dir 'sig.rtf') 'r'

        $backup = InModuleScope OutlookSignature -Parameters @{ dir = $dir } {
            param($dir)
            Backup-SignatureFileSet -SignaturesPath $dir -Name 'sig'
        }

        $backup | Should -Not -BeNullOrEmpty
        Test-Path $backup | Should -BeTrue
        Test-Path (Join-Path $dir 'sig.htm') | Should -BeFalse
        Test-Path (Join-Path $backup 'sig.htm') | Should -BeTrue
        Test-Path (Join-Path $backup 'sig.txt') | Should -BeTrue
        Test-Path (Join-Path $backup 'sig.rtf') | Should -BeTrue
    }

    It '対象が無ければ $null を返し、何も作らない' {
        $dir = Join-Path $TestDrive 'b2'
        New-Item -ItemType Directory -Path $dir | Out-Null
        $r = InModuleScope OutlookSignature -Parameters @{ dir = $dir } {
            param($dir)
            Backup-SignatureFileSet -SignaturesPath $dir -Name 'none'
        }
        $r | Should -BeNullOrEmpty
        Test-Path (Join-Path $dir '.backup') | Should -BeFalse
    }

    It '.files フォルダと _files フォルダの両方を移動する' {
        $dir = Join-Path $TestDrive 'b3'
        New-Item -ItemType Directory -Path (Join-Path $dir 'sig.files') -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $dir 'sig_files') -Force | Out-Null
        Set-Content (Join-Path $dir 'sig.htm') 'h'

        $backup = InModuleScope OutlookSignature -Parameters @{ dir = $dir } {
            param($dir)
            Backup-SignatureFileSet -SignaturesPath $dir -Name 'sig'
        }

        Test-Path (Join-Path $backup 'sig.files') | Should -BeTrue
        Test-Path (Join-Path $backup 'sig_files') | Should -BeTrue
        Test-Path (Join-Path $dir 'sig.files') | Should -BeFalse
    }

    It '同名で連続実行してもパスが衝突しない' {
        $dir = Join-Path $TestDrive 'b4'
        New-Item -ItemType Directory -Path $dir | Out-Null
        $pair = InModuleScope OutlookSignature -Parameters @{ dir = $dir } {
            param($dir)
            Set-Content (Join-Path $dir 'sig.htm') 'h1'
            $b1 = Backup-SignatureFileSet -SignaturesPath $dir -Name 'sig'
            Set-Content (Join-Path $dir 'sig.htm') 'h2'
            $b2 = Backup-SignatureFileSet -SignaturesPath $dir -Name 'sig'
            [pscustomobject]@{ B1 = $b1; B2 = $b2 }
        }
        $pair.B1 | Should -Not -Be $pair.B2
        Test-Path $pair.B1 | Should -BeTrue
        Test-Path $pair.B2 | Should -BeTrue
    }
}
