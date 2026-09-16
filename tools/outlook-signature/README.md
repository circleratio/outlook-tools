# OutlookSignature

クラシック版 Outlook デスクトップ（Windows / Microsoft 365 Apps）のメール署名を、CLI（PowerShell）から更新するモジュール。

新しい Outlook for Windows / Outlook on the web には対応していません（`doc/todo.md` 参照）。

## 必要環境

- Windows
- Windows PowerShell 5.1 以上、または PowerShell 7
- クラシック版 Outlook デスクトップ（レジストリキーは `16.0` 前提）
- 管理者権限は不要（`HKCU` と `%APPDATA%` のみ操作）

## インストール

リポジトリを取得し、モジュールを import します。

```powershell
Import-Module .\src\OutlookSignature\OutlookSignature.psd1
```

> ソースファイル（`.ps1` / `.psd1` / `.psm1`）は UTF-8 with BOM で保存されています。Windows PowerShell 5.1 が
> 日本語を含むファイルを正しく解析するために必要です。編集する場合はエンコーディングを維持してください。

## 使いかた

### 現在の署名の本文だけ差し替える（推奨）

`-Name` を省略すると、各アカウントが今使っている署名（レジストリの `New Signature` /
`Reply-Forward Signature` が指す署名）の `.htm` / `.txt` を書き換えます。レジストリは変更しません。
ローミング署名が有効な環境ではこの方法を使ってください。

```powershell
Set-OutlookSignature -Text @"
山田 太郎
一般社団法人 ○○
yamada@example.com
"@
```

### 新しい署名を作成して既定にする

`-Name` を指定すると、その名前で署名を作成／更新し、対象アカウントすべての新規メール・返信転送の
既定署名に設定します。

```powershell
Set-OutlookSignature -Name "標準" -Text "山田 太郎`n一般社団法人 ○○"
```

> ローミング署名が有効な環境では、新規作成した署名が Outlook の署名ピッカーに出ないことがあります。
> その場合は `-Name` を省略した in-place 更新を使ってください。

### よく使うオプション

| 操作 | コマンド |
| --- | --- |
| 変更内容の確認のみ（変更しない） | `Set-OutlookSignature -Text "..." -WhatIf` |
| Outlook 起動中でも実行 | `Set-OutlookSignature -Text "..." -Force` |
| 現在の署名・アカウント状態を表示 | `Get-OutlookSignature` |
| ヘルプ | `Get-Help Set-OutlookSignature -Full` |

### 反映と注意

- **反映には Outlook の再起動が必要です。** 可能なら Outlook を終了してから実行してください
  （起動中は中断します。`-Force` で続行できますが、終了時に設定が書き戻されることがあります）。
- 生成ファイルのエンコーディング: `.htm` = UTF-8（BOM なし、`charset=utf-8`）、`.txt` = UTF-16 LE（BOM 付き）。
- 上書き時は既存ファイルを `%APPDATA%\Microsoft\Signatures\.backup\<署名名>-<日時>\` に退避します。
  復元は手動です（専用コマンドはありません）。

## CLI ラッパー（終了コードが必要な場合）

`powershell -File` から実行し、意味のある終了コードを返します。

```powershell
powershell -File scripts\Set-OutlookSignatureCli.ps1 -Text "山田 太郎"
```

| 終了コード | 意味 |
| --- | --- |
| 0 | 成功 |
| 1 | 想定外エラー / 書き込み失敗 |
| 2 | プリフライト検証失敗（Outlook 未検出、プロファイル／アカウント無し、署名名不正 など） |
| 3 | Outlook 起動中（`-Force` 未指定） |

## 月替わりメッセージで署名を毎月更新する

`scripts\Set-SeasonalOutlookSignature.ps1` は、月替わりの一言メッセージをテンプレートに差し込み、
**実行した月**に合わせた署名で Outlook を更新するラッパーです。

### 入力ファイル

雛形として `message.example.txt` / `template.example.txt` を同梱しています。コピーして
`message.txt` / `template.txt`（内容は自分の署名・組織に合わせて編集）を作成してください。
これらの実データファイルは `.gitignore` 済みでリポジトリには含めません。

**メッセージファイル**（雛形 `message.example.txt`）— 1 行 1 メッセージ。1 行目＝1 月 … 12 行目＝12 月。
UTF-8（BOM 有無どちらでも可）。13 行目以降は無視します。

```
新春のお慶びを申し上げます。本年もよろしくお願いいたします。
寒さ厳しき折、ご自愛ください。
日ごとに春めいてまいりました。
（…4〜12 月分を続ける。全 12 行）
```

**テンプレートファイル**（雛形 `template.example.txt`）— 署名の本体。メッセージを差し込む位置に
`{{message}}` を置きます（`-Placeholder` で変更可、複数箇所あれば全て置換）。

```
山田 太郎
一般社団法人 ○○

{{message}}

yamada@example.com
```

### 実行

```powershell
# 実行月のメッセージで署名を更新（-Name 省略 = in-place）
powershell -File scripts\Set-SeasonalOutlookSignature.ps1 -MessagesPath .\messages.txt -TemplatePath .\template.txt

# 翌月分を事前確認（変更しない。差し込み結果が表示される）
powershell -File scripts\Set-SeasonalOutlookSignature.ps1 -MessagesPath .\messages.txt -TemplatePath .\template.txt -Month 10 -WhatIf

# Outlook 起動中でも実行
powershell -File scripts\Set-SeasonalOutlookSignature.ps1 -MessagesPath .\messages.txt -TemplatePath .\template.txt -Force
```

`-Name` / `-Force` / `-WhatIf` / `-SignaturesPath` / `-RegistryRoot` は `Set-OutlookSignature` にそのまま渡ります。
反映には Outlook の再起動が必要な点、バックアップの扱いは上記「反映と注意」と同じです。

### 毎月自動実行（タスクスケジューラ）

毎月 1 日 8:00 に実行する例（パスは実際のものに置き換えてください）:

```powershell
schtasks /Create /TN "OutlookSeasonalSignature" /SC MONTHLY /D 1 /ST 08:00 /TR `
  "powershell.exe -NoProfile -File \"C:\path\to\sig\scripts\Set-SeasonalOutlookSignature.ps1\" -MessagesPath \"C:\path\to\messages.txt\" -TemplatePath \"C:\path\to\template.txt\""
```

> Outlook 起動中に自動実行すると中断（exit 3）します。起動中でも更新したい場合は引数に `-Force` を加えてください
> （終了時に設定が書き戻されることがあります）。

### 終了コード

| 終了コード | 意味 |
| --- | --- |
| 0 | 成功 |
| 1 | 想定外エラー / 書き込み失敗 |
| 2 | プリフライト検証失敗（ファイル無し・`-Month` 範囲外・対象月のメッセージ欠落・プレースホルダ不在、および Outlook 未検出／アカウント無し等） |
| 3 | Outlook 起動中（`-Force` 未指定） |

## 開発

### テスト

Pester v5 系（開発環境は 6.1.0）が必要です。

```powershell
Install-Module Pester -Scope CurrentUser -MinimumVersion 5.1   # 未導入の場合（TestRegistry: を使うため 5.1 以上必須）
Invoke-Pester .\tests
```

レジストリ／ファイルは `TestDrive:` と一時キーで隔離しており、実環境の署名には触れません。

### 実機確認

`scripts\Invoke-S13Check.ps1` は、既存署名を退避してからサンプルで in-place 更新し、
Outlook 起動後の目視確認に備えます。`-Restore` で元に戻します（Outlook を終了して実行）。

## ドキュメント

| ファイル | 内容 |
| --- | --- |
| [doc/requirement.md](doc/requirement.md) | 要求仕様（確定事項） |
| [doc/spec.md](doc/spec.md) | 設計（モジュール構成、関数責務、エラー方針、実機確認結果） |
| [doc/plan.md](doc/plan.md) | 実装計画とステップ |
| [doc/todo.md](doc/todo.md) | 将来対応（新 Outlook / OWA 対応、ローミング署名の抑止オプション など） |
