$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$here/../src/OpenSlot/Public/ConvertTo-InputDate.ps1"

Describe 'ConvertTo-InputDate' {

    It 'yyyy-MM-dd をそのまま解釈する' {
        (ConvertTo-InputDate -Text '2024-09-01' -Today ([datetime]'2026-08-31')) | Should Be ([datetime]'2024-09-01')
    }

    It 'MM-dd は実行年を補う' {
        (ConvertTo-InputDate -Text '09-01' -Today ([datetime]'2026-08-31')) | Should Be ([datetime]'2026-09-01')
    }

    It 'MM-dd は Today の年に追従する' {
        (ConvertTo-InputDate -Text '09-01' -Today ([datetime]'2030-01-15')) | Should Be ([datetime]'2030-09-01')
    }

    It '時刻部分は落とす' {
        (ConvertTo-InputDate -Text '2026-09-01' -Today ([datetime]'2026-01-01')) | Should Be ([datetime]'2026-09-01')
    }

    It 'うるう日 MM-dd はうるう年なら通る' {
        (ConvertTo-InputDate -Text '02-29' -Today ([datetime]'2028-01-01')) | Should Be ([datetime]'2028-02-29')
    }

    It 'うるう日 MM-dd は非うるう年なら例外' {
        { ConvertTo-InputDate -Text '02-29' -Today ([datetime]'2026-01-01') } | Should Throw
    }

    It 'スラッシュ区切りは例外' {
        { ConvertTo-InputDate -Text '2026/09/01' -Today ([datetime]'2026-01-01') } | Should Throw
    }

    It '1 桁月日（M-d）は例外' {
        { ConvertTo-InputDate -Text '9-1' -Today ([datetime]'2026-01-01') } | Should Throw
    }

    It '不正文字列は例外' {
        { ConvertTo-InputDate -Text 'bad' -Today ([datetime]'2026-01-01') } | Should Throw
    }

    It '存在しない月日は例外' {
        { ConvertTo-InputDate -Text '13-40' -Today ([datetime]'2026-01-01') } | Should Throw
    }
}
