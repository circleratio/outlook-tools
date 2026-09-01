BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    Import-TargetModule
}

Describe 'Write-SignatureFileSet' {
    It '.htm と .txt のみ生成し .rtf は作らない' {
        $dir = Join-Path $TestDrive 'sig1'
        InModuleScope OutlookSignature -Parameters @{ dir = $dir } {
            param($dir)
            $files = Write-SignatureFileSet -SignaturesPath $dir -Name 'テスト' -Text '山田 太郎'
            $files.Count | Should -Be 2
        }
        Test-Path (Join-Path $dir 'テスト.htm') | Should -BeTrue
        Test-Path (Join-Path $dir 'テスト.txt') | Should -BeTrue
        Test-Path (Join-Path $dir 'テスト.rtf') | Should -BeFalse
    }

    It '.htm は BOM 無し（先頭が "<"）、.txt は UTF-16LE BOM（FF FE）' {
        $dir = Join-Path $TestDrive 'sig2'
        InModuleScope OutlookSignature -Parameters @{ dir = $dir } {
            param($dir)
            Write-SignatureFileSet -SignaturesPath $dir -Name 's' -Text '日本語テスト' | Out-Null
        }
        $htmBytes = [System.IO.File]::ReadAllBytes((Join-Path $dir 's.htm'))
        $txtBytes = [System.IO.File]::ReadAllBytes((Join-Path $dir 's.txt'))
        $htmBytes[0] | Should -Be 0x3C
        $txtBytes[0] | Should -Be 0xFF
        $txtBytes[1] | Should -Be 0xFE
    }

    It '本文を正しいエンコーディングで読み戻せる（文字化けしない）' {
        $dir = Join-Path $TestDrive 'sig3'
        InModuleScope OutlookSignature -Parameters @{ dir = $dir } {
            param($dir)
            Write-SignatureFileSet -SignaturesPath $dir -Name 's' -Text "山田 太郎`n㈱テスト" | Out-Null
        }
        [System.IO.File]::ReadAllText((Join-Path $dir 's.htm'), [System.Text.Encoding]::UTF8) | Should -Match '山田 太郎'
        [System.IO.File]::ReadAllText((Join-Path $dir 's.txt'), [System.Text.Encoding]::Unicode) | Should -Match '㈱テスト'
    }

    It '.txt の改行は CRLF' {
        $dir = Join-Path $TestDrive 'sig4'
        InModuleScope OutlookSignature -Parameters @{ dir = $dir } {
            param($dir)
            Write-SignatureFileSet -SignaturesPath $dir -Name 's' -Text "a`nb" | Out-Null
        }
        $t = [System.IO.File]::ReadAllText((Join-Path $dir 's.txt'), [System.Text.Encoding]::Unicode)
        $t | Should -Be "a`r`nb"
    }

    It '存在しないサブフォルダを自動作成する' {
        $dir = Join-Path $TestDrive 'nested\deeper\sig'
        InModuleScope OutlookSignature -Parameters @{ dir = $dir } {
            param($dir)
            Write-SignatureFileSet -SignaturesPath $dir -Name 's' -Text 'x' | Out-Null
        }
        Test-Path (Join-Path $dir 's.htm') | Should -BeTrue
    }
}
