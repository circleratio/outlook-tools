<#
.SYNOPSIS
    月替わりメッセージを差し込んだ署名で Outlook を更新するラッパー。

.DESCRIPTION
    「1 行 1 メッセージ（1 行目＝1 月 … 12 行目＝12 月）」のメッセージファイルと、
    差し込み位置にプレースホルダ（既定 {{message}}）を含むテンプレートファイルを受け取り、
    対象月（既定は実行月。-Month で上書き可）のメッセージをテンプレートへ差し込んだ
    プレーンテキストで Set-OutlookSignature を呼び出す。

    scripts/Set-OutlookSignatureCli.ps1 は変更せず、別スクリプトとして動作する。
    終了コードは Set-OutlookSignatureCli.ps1 と同一体系。
        0 = 成功
        1 = 想定外エラー / 書き込み失敗（WriteFailed ほか）
        2 = プリフライト検証失敗（入力ファイル不備・月範囲外・メッセージ欠落・
            プレースホルダ不在、および Set-OutlookSignature のプリフライト失敗）
        3 = Outlook 起動中（-Force 未指定）

.PARAMETER MessagesPath
    月替わりメッセージファイル。1 行 1 メッセージ、上から 1〜12 月。
    文字コードは UTF-8（BOM 有無どちらでも可）。13 行目以降は無視する。

.PARAMETER TemplatePath
    署名テンプレート（プレーンテキスト、複数行可）。差し込み位置に -Placeholder を置く。

.PARAMETER Month
    対象月（1〜12）。省略時は実行時のシステムローカル日付の月。範囲外はエラー（exit 2）。

.PARAMETER Placeholder
    テンプレート内の差し込み位置マーカー。既定 "{{message}}"。
    複数箇所に現れる場合はすべて置換する。

.PARAMETER Name
    署名名。Set-OutlookSignature へ透過。省略時は現在の既定署名を in-place 更新。

.PARAMETER Force
    Outlook 起動中でも続行する。Set-OutlookSignature へ透過。

.PARAMETER SignaturesPath
    Signatures フォルダのパス（テスト／上級者向け）。Set-OutlookSignature へ透過。

.PARAMETER RegistryRoot
    Office のレジストリルート（テスト／上級者向け）。Set-OutlookSignature へ透過。

.EXAMPLE
    powershell -File scripts/Set-SeasonalOutlookSignature.ps1 -MessagesPath .\messages.txt -TemplatePath .\template.txt

    実行月のメッセージを差し込んで署名を in-place 更新する。

.EXAMPLE
    powershell -File scripts/Set-SeasonalOutlookSignature.ps1 -MessagesPath .\messages.txt -TemplatePath .\template.txt -Month 10 -WhatIf

    10 月分の差し込み結果を確認する（実際には変更しない）。
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string] $MessagesPath,

    [Parameter(Mandatory)]
    [string] $TemplatePath,

    [int] $Month,

    [string] $Placeholder = '{{message}}',

    [string] $Name,

    [switch] $Force,

    [string] $SignaturesPath,

    [string] $RegistryRoot = 'HKCU:\Software\Microsoft\Office\16.0'
)

$ErrorActionPreference = 'Stop'

function Read-TextFileOrThrow {
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [string] $ErrorId
    )
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "ファイルが見つかりません: $Path [$ErrorId]"
    }
    try {
        $full = (Resolve-Path -LiteralPath $Path).ProviderPath
        # .NET が UTF-8 / UTF-16 の BOM を判定し、BOM 無しは UTF-8 とみなす。
        return [System.IO.File]::ReadAllText($full)
    }
    catch {
        throw "ファイルを読み取れません: $Path [$ErrorId]"
    }
}

function Get-SeasonalMessageLine {
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [int] $Month
    )
    $content = Read-TextFileOrThrow -Path $Path -ErrorId 'MessagesFileNotFound'
    $lines = @($content -split "\r?\n")
    if ($lines.Count -lt $Month -or [string]::IsNullOrWhiteSpace($lines[$Month - 1])) {
        throw "$Month 月のメッセージが $Path にありません（$Month 行目が空、または行数不足）。 [NoMessageForMonth]"
    }
    return $lines[$Month - 1].Trim()
}

function Expand-SignatureTemplate {
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [string] $Placeholder,
        [Parameter(Mandatory)] [AllowEmptyString()] [string] $Message
    )
    $template = Read-TextFileOrThrow -Path $Path -ErrorId 'TemplateFileNotFound'
    if (-not $template.Contains($Placeholder)) {
        throw "テンプレート $Path にプレースホルダ '$Placeholder' が含まれていません。 [PlaceholderNotFound]"
    }
    return $template.Replace($Placeholder, $Message)
}

# ErrorId（FullyQualifiedErrorId に含まれる文字列）→ 終了コード 2 に写像する ID 一覧。
# Set-OutlookSignatureCli.ps1 のプリフライト ID にラッパー固有 ID を加えたもの。
$preflightIds = @(
    'MessagesFileNotFound', 'TemplateFileNotFound', 'MonthOutOfRange',
    'NoMessageForMonth', 'PlaceholderNotFound',
    'InvalidSignatureName', 'OutlookNotFound', 'DefaultProfileNotFound',
    'ProfilePromptEnabled', 'NoMailAccount', 'NoCurrentSignature', 'SignaturesFolderNotWritable'
)

try {
    $targetMonth = if ($PSBoundParameters.ContainsKey('Month')) { $Month } else { (Get-Date).Month }
    if ($targetMonth -lt 1 -or $targetMonth -gt 12) {
        throw "対象月は 1〜12 で指定してください: $targetMonth [MonthOutOfRange]"
    }

    $message = Get-SeasonalMessageLine -Path $MessagesPath -Month $targetMonth
    $text = Expand-SignatureTemplate -Path $TemplatePath -Placeholder $Placeholder -Message $message

    Import-Module (Join-Path $PSScriptRoot '..\src\OutlookSignature\OutlookSignature.psd1') -Force

    $params = @{
        Text         = $text
        RegistryRoot = $RegistryRoot
        WhatIf       = [bool]$WhatIfPreference
    }
    if ($PSBoundParameters.ContainsKey('Name') -and -not [string]::IsNullOrEmpty($Name)) { $params.Name = $Name }
    if ($Force) { $params.Force = $true }
    if ($PSBoundParameters.ContainsKey('SignaturesPath')) { $params.SignaturesPath = $SignaturesPath }

    Write-Host ("対象月        : {0}" -f $targetMonth)
    Write-Host ("使用メッセージ : {0}" -f $message)
    Write-Host "--- 署名本文 ---"
    Write-Host $text
    Write-Host "----------------"

    $result = Set-OutlookSignature @params
    $result | Format-List | Out-String | Write-Host
    exit 0
}
catch {
    $id = [string]$_.FullyQualifiedErrorId
    [Console]::Error.WriteLine($_.Exception.Message)

    if ($id -match 'OutlookRunning') { exit 3 }
    foreach ($pf in $preflightIds) {
        if ($id -match $pf) { exit 2 }
    }
    exit 1
}
