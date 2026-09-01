<#
.SYNOPSIS
    S13（実機確認）用スクリプト。既存署名を安全に退避してから in-place 更新を実行し、
    Outlook 起動後の目視確認に備える。

.DESCRIPTION
    実行前に Outlook を終了しておくこと。
    1. 現在の Signatures フォルダ全体をタイムスタンプ付きで別フォルダに複製（退避）
    2. Set-OutlookSignature -Text <日本語サンプル> を in-place 実行
    3. 生成された .htm / .txt のパスとエンコーディングを表示
    その後 Outlook を起動し、新規メール／返信で署名が文字化けなく出るか目視する。

    確認が終わったら -Restore を付けて再実行すると、最新の退避フォルダから元に戻す。

.PARAMETER Restore
    最新の退避フォルダから Signatures を復元する。

.EXAMPLE
    powershell -File scripts/Invoke-S13Check.ps1
    # → Outlook 起動して確認 → powershell -File scripts/Invoke-S13Check.ps1 -Restore
#>
[CmdletBinding()]
param(
    [switch] $Restore
)

$ErrorActionPreference = 'Stop'
$sigPath   = Join-Path ([Environment]::GetFolderPath('ApplicationData')) 'Microsoft\Signatures'
$stashRoot = Join-Path ([Environment]::GetFolderPath('ApplicationData')) 'Microsoft\Signatures.s13-stash'

if (Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue) {
    throw 'Outlook が起動中です。終了してから実行してください。'
}

if ($Restore) {
    $latest = Get-ChildItem $stashRoot -Directory | Sort-Object Name -Descending | Select-Object -First 1
    if (-not $latest) { throw "退避フォルダが見つかりません: $stashRoot" }
    Write-Host "復元元: $($latest.FullName)"
    Get-ChildItem $sigPath -Force | Where-Object { $_.Name -ne 'Signatures.s13-stash' } | Remove-Item -Recurse -Force
    Copy-Item (Join-Path $latest.FullName '*') $sigPath -Recurse -Force
    Write-Host '復元しました。Outlook を起動して確認してください。'
    return
}

# 1. 退避
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$stash = Join-Path $stashRoot $stamp
New-Item -ItemType Directory -Path $stash -Force | Out-Null
Copy-Item (Join-Path $sigPath '*') $stash -Recurse -Force
Write-Host "退避先: $stash"

# 2. in-place 更新（絵文字・機種依存文字を含むサンプル）
Import-Module (Join-Path $PSScriptRoot '..\src\OutlookSignature\OutlookSignature.psd1') -Force
$sample = @"
山田 太郎（やまだ たろう）
一般社団法人 ○○○○
㈱ / ① / ～ / — / 髙橋
yamada@example.com
"@
$result = Set-OutlookSignature -Text $sample
$result | Format-List Mode, SignatureNames, Files, BackupPaths, RegistryChanged, RoamingSignaturesEnabled

# 3. エンコーディング表示
foreach ($f in $result.Files) {
    $bytes = [System.IO.File]::ReadAllBytes($f)
    $head  = ($bytes[0..3] | ForEach-Object { $_.ToString('X2') }) -join ' '
    Write-Host ("{0}  先頭バイト: {1}" -f (Split-Path $f -Leaf), $head)
}

Write-Host ''
Write-Host '次の手順:'
Write-Host '  1) Outlook を起動'
Write-Host '  2) 新規メール／返信を作成し、署名が文字化けなく出るか確認'
Write-Host '  3) 数分後に再度 Outlook を再起動し、内容が保持されているか確認（ローミング同期）'
Write-Host '  4) 確認後: powershell -File scripts/Invoke-S13Check.ps1 -Restore'
