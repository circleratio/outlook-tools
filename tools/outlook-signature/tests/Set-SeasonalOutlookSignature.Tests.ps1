BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    $script:Script = Join-Path $PSScriptRoot '..\scripts\Set-SeasonalOutlookSignature.ps1'
    # 別プロセスで実行するため TestRegistry: (Pester のプロセス内ドライブ) が使えない。
    # 実 HKCU 配下に一意な一時キーを作り、AfterAll で削除する。
    $script:TestKeyBase = "HKCU:\Software\__OLSIG_SEASONAL_TEST__"

    function Invoke-Seasonal {
        param([string[]] $ScriptArgs)
        $psExe = (Get-Process -Id $PID).Path
        & $psExe -NoProfile -NonInteractive -File $script:Script @ScriptArgs *> $null
        return $LASTEXITCODE
    }

    function New-MessagesFile {
        param([string[]] $Lines, [switch] $Bom)
        $p = Join-Path $TestDrive ([guid]::NewGuid().ToString('N') + '.txt')
        [System.IO.File]::WriteAllText($p, ($Lines -join "`r`n"), (New-Object System.Text.UTF8Encoding($Bom.IsPresent)))
        return $p
    }

    function New-TemplateFile {
        param([string] $Content, [switch] $Bom)
        $p = Join-Path $TestDrive ([guid]::NewGuid().ToString('N') + '.txt')
        [System.IO.File]::WriteAllText($p, $Content, (New-Object System.Text.UTF8Encoding($Bom.IsPresent)))
        return $p
    }

    $script:TwelveMonths = 1..12 | ForEach-Object { "{0}月のあいさつ" -f $_ }

    function New-InPlaceEnv {
        New-FakeOutlookRegistry -RegistryRoot $script:Root -DisableRoamingToggle 1 -Accounts @{
            '00000002' = @{ ServiceName = 'MSEMS'; AccountName = 'a@b.c'; New = 'cur'; Reply = 'cur' }
        } | Out-Null
    }
}

AfterAll {
    if (Test-Path $script:TestKeyBase) { Remove-Item $script:TestKeyBase -Recurse -Force }
}

Describe 'Set-SeasonalOutlookSignature.ps1' {
    BeforeEach {
        $script:Root = Join-Path $script:TestKeyBase ("{0}\16.0" -f [guid]::NewGuid().ToString('N'))
        $script:Sp   = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $script:Msgs = New-MessagesFile -Lines $script:TwelveMonths
        $script:Tpl  = New-TemplateFile -Content "山田 太郎`r`n`r`n{{message}}`r`n`r`nyamada@example.com"
    }

    It '正常時（-Month 指定）は exit 0 で対象月メッセージが差し込まれる' {
        New-InPlaceEnv
        Invoke-Seasonal @('-MessagesPath', $script:Msgs, '-TemplatePath', $script:Tpl,
            '-Month', '3', '-Force', '-SignaturesPath', $script:Sp, '-RegistryRoot', $script:Root) |
            Should -Be 0
        $txt = Get-Content -LiteralPath (Join-Path $script:Sp 'cur.txt') -Raw
        $txt | Should -Match '3月のあいさつ'
        $txt | Should -Match '山田 太郎'
    }

    It '-Month 省略で当月のメッセージが使われる' {
        New-InPlaceEnv
        $m = (Get-Date).Month
        Invoke-Seasonal @('-MessagesPath', $script:Msgs, '-TemplatePath', $script:Tpl,
            '-Force', '-SignaturesPath', $script:Sp, '-RegistryRoot', $script:Root) |
            Should -Be 0
        (Get-Content -LiteralPath (Join-Path $script:Sp 'cur.txt') -Raw) | Should -Match ("{0}月のあいさつ" -f $m)
    }

    It 'カスタム -Placeholder を複数箇所とも置換する' {
        New-InPlaceEnv
        $tpl = New-TemplateFile -Content "%X% / start`r`n%X% / end"
        Invoke-Seasonal @('-MessagesPath', $script:Msgs, '-TemplatePath', $tpl,
            '-Month', '5', '-Placeholder', '%X%', '-Force',
            '-SignaturesPath', $script:Sp, '-RegistryRoot', $script:Root) |
            Should -Be 0
        $txt = Get-Content -LiteralPath (Join-Path $script:Sp 'cur.txt') -Raw
        ([regex]::Matches($txt, '5月のあいさつ')).Count | Should -Be 2
    }

    It 'メッセージファイルが存在しない → exit 2' {
        New-InPlaceEnv
        Invoke-Seasonal @('-MessagesPath', (Join-Path $TestDrive 'nope.txt'), '-TemplatePath', $script:Tpl,
            '-Month', '3', '-Force', '-SignaturesPath', $script:Sp, '-RegistryRoot', $script:Root) |
            Should -Be 2
    }

    It 'テンプレートファイルが存在しない → exit 2' {
        New-InPlaceEnv
        Invoke-Seasonal @('-MessagesPath', $script:Msgs, '-TemplatePath', (Join-Path $TestDrive 'nope.txt'),
            '-Month', '3', '-Force', '-SignaturesPath', $script:Sp, '-RegistryRoot', $script:Root) |
            Should -Be 2
    }

    It '-Month 範囲外（13）→ exit 2' {
        New-InPlaceEnv
        Invoke-Seasonal @('-MessagesPath', $script:Msgs, '-TemplatePath', $script:Tpl,
            '-Month', '13', '-Force', '-SignaturesPath', $script:Sp, '-RegistryRoot', $script:Root) |
            Should -Be 2
    }

    It '-Month 範囲外（0）→ exit 2' {
        New-InPlaceEnv
        Invoke-Seasonal @('-MessagesPath', $script:Msgs, '-TemplatePath', $script:Tpl,
            '-Month', '0', '-Force', '-SignaturesPath', $script:Sp, '-RegistryRoot', $script:Root) |
            Should -Be 2
    }

    It '対象月のメッセージが空白のみ → exit 2' {
        New-InPlaceEnv
        $msgs = New-MessagesFile -Lines @('1月', '   ', '3月')
        Invoke-Seasonal @('-MessagesPath', $msgs, '-TemplatePath', $script:Tpl,
            '-Month', '2', '-Force', '-SignaturesPath', $script:Sp, '-RegistryRoot', $script:Root) |
            Should -Be 2
    }

    It '行数不足で対象月が無い → exit 2' {
        New-InPlaceEnv
        $msgs = New-MessagesFile -Lines @('1月', '2月', '3月')
        Invoke-Seasonal @('-MessagesPath', $msgs, '-TemplatePath', $script:Tpl,
            '-Month', '7', '-Force', '-SignaturesPath', $script:Sp, '-RegistryRoot', $script:Root) |
            Should -Be 2
    }

    It 'プレースホルダを含まないテンプレート → exit 2' {
        New-InPlaceEnv
        $tpl = New-TemplateFile -Content "山田 太郎`r`nyamada@example.com"
        Invoke-Seasonal @('-MessagesPath', $script:Msgs, '-TemplatePath', $tpl,
            '-Month', '3', '-Force', '-SignaturesPath', $script:Sp, '-RegistryRoot', $script:Root) |
            Should -Be 2
    }

    It '下流のプリフライト失敗（NoMailAccount）→ exit 2' {
        New-FakeOutlookRegistry -RegistryRoot $script:Root -Accounts @{ '1' = @{ ServiceName = 'CONTAB' } } | Out-Null
        Invoke-Seasonal @('-MessagesPath', $script:Msgs, '-TemplatePath', $script:Tpl,
            '-Month', '3', '-Force', '-SignaturesPath', $script:Sp, '-RegistryRoot', $script:Root) |
            Should -Be 2
    }

    It '下流のプリフライト失敗（NoCurrentSignature）→ exit 2' {
        New-FakeOutlookRegistry -RegistryRoot $script:Root -DisableRoamingToggle 1 -Accounts @{
            '00000002' = @{ ServiceName = 'MSEMS'; AccountName = 'a@b.c' }
        } | Out-Null
        Invoke-Seasonal @('-MessagesPath', $script:Msgs, '-TemplatePath', $script:Tpl,
            '-Month', '3', '-Force', '-SignaturesPath', $script:Sp, '-RegistryRoot', $script:Root) |
            Should -Be 2
    }

    It '-WhatIf は exit 0 で署名ファイルを生成しない' {
        New-InPlaceEnv
        Invoke-Seasonal @('-MessagesPath', $script:Msgs, '-TemplatePath', $script:Tpl,
            '-Month', '3', '-Force', '-WhatIf', '-SignaturesPath', $script:Sp, '-RegistryRoot', $script:Root) |
            Should -Be 0
        (Test-Path -LiteralPath (Join-Path $script:Sp 'cur.txt')) | Should -BeFalse
    }

    It '13 行目以降は無視される' {
        New-InPlaceEnv
        $msgs = New-MessagesFile -Lines ($script:TwelveMonths + @('EXTRA'))
        Invoke-Seasonal @('-MessagesPath', $msgs, '-TemplatePath', $script:Tpl,
            '-Month', '12', '-Force', '-SignaturesPath', $script:Sp, '-RegistryRoot', $script:Root) |
            Should -Be 0
        $txt = Get-Content -LiteralPath (Join-Path $script:Sp 'cur.txt') -Raw
        $txt | Should -Match '12月のあいさつ'
        $txt | Should -Not -Match 'EXTRA'
    }

    It 'BOM 付き UTF-8 のメッセージ／テンプレートでも文字化けしない' {
        New-InPlaceEnv
        $msgs = New-MessagesFile -Lines @('あ', 'い', '髙島㈱ ①') -Bom
        $tpl  = New-TemplateFile -Content "拝啓`r`n{{message}}`r`n敬具" -Bom
        Invoke-Seasonal @('-MessagesPath', $msgs, '-TemplatePath', $tpl,
            '-Month', '3', '-Force', '-SignaturesPath', $script:Sp, '-RegistryRoot', $script:Root) |
            Should -Be 0
        (Get-Content -LiteralPath (Join-Path $script:Sp 'cur.txt') -Raw) | Should -Match '髙島㈱ ①'
    }
}
