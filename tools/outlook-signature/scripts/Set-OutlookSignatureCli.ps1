<#
.SYNOPSIS
    Set-OutlookSignature を呼び出し、意味のある終了コードを返す薄いラッパー。

.DESCRIPTION
    powershell -File からの実行を想定。ErrorId を終了コードに写像する。
        0 = 成功
        1 = 想定外エラー / 書き込み失敗（WriteFailed ほか）
        2 = プリフライト検証失敗
        3 = Outlook 起動中（-Force 未指定）

.PARAMETER Name
    署名名。省略すると現在の既定署名を in-place 更新する。

.PARAMETER Text
    署名本文（プレーンテキスト、複数行可）。

.PARAMETER Force
    Outlook 起動中でも続行する。

.PARAMETER SignaturesPath
    Signatures フォルダのパス（テスト／上級者向け）。

.PARAMETER RegistryRoot
    Office のレジストリルート（テスト／上級者向け）。

.EXAMPLE
    powershell -File scripts/Set-OutlookSignatureCli.ps1 -Text "山田 太郎"
#>
[CmdletBinding()]
param(
    [string] $Name,

    [Parameter(Mandatory)]
    [AllowEmptyString()]
    [string] $Text,

    [switch] $Force,

    [string] $SignaturesPath,

    [string] $RegistryRoot = 'HKCU:\Software\Microsoft\Office\16.0'
)

$ErrorActionPreference = 'Stop'

$preflightIds = @(
    'InvalidSignatureName', 'OutlookNotFound', 'DefaultProfileNotFound',
    'ProfilePromptEnabled', 'NoMailAccount', 'NoCurrentSignature', 'SignaturesFolderNotWritable'
)

try {
    Import-Module (Join-Path $PSScriptRoot '..\src\OutlookSignature\OutlookSignature.psd1') -Force

    $params = @{
        Text         = $Text
        RegistryRoot = $RegistryRoot
    }
    if ($PSBoundParameters.ContainsKey('Name') -and -not [string]::IsNullOrEmpty($Name)) { $params.Name = $Name }
    if ($Force) { $params.Force = $true }
    if ($PSBoundParameters.ContainsKey('SignaturesPath')) { $params.SignaturesPath = $SignaturesPath }

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
