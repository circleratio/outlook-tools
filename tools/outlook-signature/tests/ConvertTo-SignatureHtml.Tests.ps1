BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    Import-TargetModule
}

Describe 'ConvertTo-SignatureHtml' {
    It 'HTML 特殊文字をエスケープする' {
        InModuleScope OutlookSignature {
            $h = ConvertTo-SignatureHtml -Text 'a<b> & "q"'
            $h | Should -Match '&lt;b&gt;'
            $h | Should -Match '&amp;'
            $h | Should -Match '&quot;'
            $h | Should -Not -Match '<b>'
        }
    }

    It 'CRLF / LF / CR の改行を br 要素に変換する' {
        InModuleScope OutlookSignature {
            (ConvertTo-SignatureHtml -Text "a`r`nb") | Should -Match 'a<br>'
            (ConvertTo-SignatureHtml -Text "a`nb")   | Should -Match 'a<br>'
            (ConvertTo-SignatureHtml -Text "a`rb")   | Should -Match 'a<br>'
        }
    }

    It 'charset=utf-8 と div 要素を含む' {
        InModuleScope OutlookSignature {
            $h = ConvertTo-SignatureHtml -Text 'x'
            $h | Should -Match 'charset=utf-8'
            $h | Should -Match '<div>'
        }
    }

    It '出力の改行は CRLF に統一される' {
        InModuleScope OutlookSignature {
            $h = ConvertTo-SignatureHtml -Text "a`nb"
            ([regex]::Matches($h, "(?<!`r)`n")).Count | Should -Be 0
        }
    }

    It '空文字入力でも例外なく最小 HTML を返す' {
        InModuleScope OutlookSignature {
            $h = ConvertTo-SignatureHtml -Text ''
            $h | Should -Match '<!DOCTYPE html>'
            $h | Should -Match '<div></div>'
        }
    }
}
