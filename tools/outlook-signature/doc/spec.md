# 設計: Outlook 署名更新 CLI（クラシック版デスクトップ）

入力: `doc/requirement.md`（確定した要求仕様）
対象フェーズ: クラシック版 Outlook デスクトップのみ。新 Outlook / OWA は `doc/todo.md`。

## 1. ディレクトリ構成

```
tools/outlook-signature/
├─ doc/
│  ├─ requirement.md
│  ├─ spec.md          ← 本ファイル
│  ├─ plan.md
│  └─ todo.md
├─ src/
│  └─ OutlookSignature/
│     ├─ OutlookSignature.psd1        # マニフェスト。ModuleVersion をここで一元管理
│     ├─ OutlookSignature.psm1        # ルート。Public/Private をドットソースし公開関数を Export
│     ├─ Public/
│     │  ├─ Set-OutlookSignature.ps1
│     │  └─ Get-OutlookSignature.ps1
│     └─ Private/
│        ├─ Resolve-OutlookEnvironment.ps1
│        ├─ Get-OutlookMailAccount.ps1
│        ├─ Test-OutlookRunning.ps1
│        ├─ Test-RoamingSignatureEnabled.ps1
│        ├─ ConvertTo-SignatureHtml.ps1
│        ├─ Write-SignatureFileSet.ps1
│        ├─ Backup-SignatureFileSet.ps1
│        └─ Set-AccountSignatureRegistry.ps1
├─ scripts/
│  ├─ Set-OutlookSignatureCli.ps1        # 終了コードを返す薄いラッパー
│  └─ Set-SeasonalOutlookSignature.ps1   # 月替わりメッセージ差し込みラッパー（§10）
├─ tests/
│  ├─ ConvertTo-SignatureHtml.Tests.ps1
│  ├─ Write-SignatureFileSet.Tests.ps1
│  ├─ Backup-SignatureFileSet.Tests.ps1
│  ├─ Get-OutlookMailAccount.Tests.ps1
│  ├─ Set-OutlookSignature.Tests.ps1
│  └─ Set-SeasonalOutlookSignature.Tests.ps1
└─ README.md
```

### モジュール構成方針

- `OutlookSignature.psd1` … `ModuleVersion`（バージョンの唯一の情報源）、`RootModule = 'OutlookSignature.psm1'`、`PowerShellVersion = '5.1'`、`FunctionsToExport = @('Set-OutlookSignature','Get-OutlookSignature')`、依存モジュールなし。
- `OutlookSignature.psm1` … `Public\*.ps1` と `Private\*.ps1` をドットソースし、`Export-ModuleMember -Function 'Set-OutlookSignature','Get-OutlookSignature'`。
- Private 関数はエクスポートしない（テストは `InModuleScope` で参照する）。

## 2. モジュールごとの責務と主要シグネチャ

### Public / Set-OutlookSignature（メイン）

```powershell
function Set-OutlookSignature {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    [OutputType([pscustomobject])]
    param(
        # 省略時: 各アカウントの現行 New/Reply-Forward Signature が指す既存署名を更新（in-place）
        # 指定時: その名前で署名を作成/更新し、全対象アカウントの既定署名に設定
        [Parameter(Position = 0)]
        [string] $Name,

        [Parameter(Mandatory, Position = 1, ValueFromPipeline = $true)]
        [AllowEmptyString()]
        [string] $Text,

        # Outlook 起動中でも続行する
        [switch] $Force,

        # --- テスト/上級者向けオーバーライド（既定は実際のパス）---
        [string] $SignaturesPath,
        [string] $RegistryRoot = 'HKCU:\Software\Microsoft\Office\16.0'
    )
}
```

**2 つのモード**（ローミング署名が ON の環境が既定なので、既存署名の in-place 更新を基本形とする）:

| モード | 条件 | ファイル | レジストリ | 用途 |
| --- | --- | --- | --- | --- |
| `UpdateInPlace` | `-Name` 省略 | 対象アカウントの現行 `New Signature` ∪ `Reply-Forward Signature` に現れる各署名名の `.htm`/`.txt` を書き換え | 触らない | ローミング署名を含む「今使っている署名」の本文だけ差し替え |
| `CreateAndAssign` | `-Name` 指定 | `<Name>.htm`/`.txt` を作成/更新 | 全対象アカウントの `New Signature` と `Reply-Forward Signature` に `<Name>` を設定 | 新しい署名を作って既定にする／名前を切り替える |

処理順序:

1. **入力検証**:
   - `-Name` 指定時のみ: 空でない / `[IO.Path]::GetInvalidFileNameChars()` を含まない / 長さ 1..255。32 超は警告のみ。NG は `InvalidSignatureName`。
2. **環境解決** `Resolve-OutlookEnvironment -RegistryRoot $RegistryRoot -SignaturesPath $SignaturesPath` → `$env`。
3. **アカウント列挙** `Get-OutlookMailAccount -AccountsKeyPath $env.AccountsKeyPath` → `$accounts`（0 件なら `NoMailAccount`）。各要素は現行の `New Signature` / `Reply-Forward Signature` 値も保持。
4. **更新対象署名名の決定**:
   - `CreateAndAssign`: `@($Name)`。
   - `UpdateInPlace`: `$accounts` の `PreviousNew` と `PreviousReplyForward` の非空値の集合。空集合なら `NoCurrentSignature` で終了。
5. **Signatures フォルダ検証**: `$env.SignaturesPath` が作成可能か（親の書き込み可否を試す）。NG は `SignaturesFolderNotWritable`。
6. **Outlook 起動チェック** `Test-OutlookRunning`。起動中かつ `-Force` 無しなら `OutlookRunning` で終了。起動中かつ `-Force` は `Write-Warning`（設定が書き戻される可能性）。
7. **ローミング署名チェック** `Test-RoamingSignatureEnabled`。有効の疑いがあれば `Write-Warning`（続行）。`CreateAndAssign` かつローミング有効時は「新規署名がクラウド側リストに現れないことがある」旨も警告。
8. ここまでが**プリフライト（読み取りのみ）**。以降が変更。`$PSCmdlet.ShouldProcess()` で各変更をガード（`-WhatIf` 時は実行しない）。
9. **バックアップ** 更新対象の各署名名について `Backup-SignatureFileSet` → `$backupDirs`。
10. **ファイル生成** 更新対象の各署名名について `Write-SignatureFileSet -SignaturesPath ... -Name <sig> -Text $Text` → `$files`。
11. **レジストリ更新** `CreateAndAssign` のときのみ、各 `$accounts` に `Set-AccountSignatureRegistry -AccountKeyPath ... -SignatureName $Name`。
12. 途中失敗時は `try/catch` でベストエフォートのロールバック（生成ファイル削除 → バックアップから復元、レジストリは旧値へ戻す）、その後 `WriteFailed` で終了。
13. 結果オブジェクトを返す。

戻り値 `[pscustomobject]`:

| プロパティ | 型 | 内容 |
| --- | --- | --- |
| `Mode` | string | `UpdateInPlace` / `CreateAndAssign` |
| `SignatureNames` | string[] | 実際に書き換えた署名名 |
| `SignaturesPath` | string | 生成先フォルダ |
| `Files` | string[] | 生成したファイルの絶対パス |
| `BackupPaths` | string[] | バックアップフォルダ（対象が存在した署名の分のみ） |
| `Accounts` | pscustomobject[] | `AccountName`, `ServiceName`, `RegistryPath`, `PreviousNew`, `PreviousReplyForward` |
| `RegistryChanged` | bool | `CreateAndAssign` で実際に書いたら `$true` |
| `RoamingSignaturesEnabled` | bool | 判定結果 |
| `Warnings` | string[] | 出した警告 |
| `RestartRequired` | bool | 常に `$true`（Outlook 再起動で反映） |

### Public / Get-OutlookSignature（確認用・読み取り専用）

```powershell
function Get-OutlookSignature {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [string] $SignaturesPath,
        [string] $RegistryRoot = 'HKCU:\Software\Microsoft\Office\16.0'
    )
}
```

`Signatures` フォルダにある署名名の一覧と、アカウントごとの現在の `New Signature` / `Reply-Forward Signature` を返す。手動確認・テストの検証に使う。

### Private 関数

| 関数 | 入力 | 出力 / 責務 |
| --- | --- | --- |
| `Resolve-OutlookEnvironment` | `-RegistryRoot`, `-SignaturesPath`(任意) | `$env`: `RegistryRoot`, `OutlookKeyPath`(`...\Outlook`), `DefaultProfile`, `ProfileKeyPath`(`...\Profiles\<profile>`), `AccountsKeyPath`(`...\9375CFF0413111d3B88A00104B2A6676`), `SignaturesPath`。`Outlook` キー無し→`OutlookNotFound`。`DefaultProfile` 値無し/プロファイルキー無し→`DefaultProfileNotFound`。`...\Outlook` の `PickLogonProfile = 1`（起動時プロファイル選択）→`ProfilePromptEnabled`。`SignaturesPath` 未指定時は `[Environment]::GetFolderPath('ApplicationData') + '\Microsoft\Signatures'`。 |
| `Get-OutlookMailAccount` | `-AccountsKeyPath` | 各サブキーを読み、`Service Name` がメール送信系（下記）のものだけを返す。要素: `{AccountName=<Account Name>; ServiceName=<Service Name>; RegistryPath=<subkey>; PreviousNew=<New Signature>; PreviousReplyForward=<Reply-Forward Signature>}`。メール送信系 `Service Name` = `MSEMS`(Exchange/M365) / `IMAP` / `POP3` / `SMTP` / `MAPI` / `EAS` / `EXHTTP`。除外例: `CONTAB`(アドレス帳) / `MSPST MS`(PST)。（実機確認済み: §7-2） |
| `Test-OutlookRunning` | なし | `[bool]` `Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue` の有無。テストでは `Mock` する。 |
| `Test-RoamingSignatureEnabled` | `-RegistryRoot` | `[bool]`（疑いの有無）。`...\Outlook\Setup\DisableRoamingSignaturesTemporaryToggle`（REG_DWORD）が `1` なら `$false`。それ以外は `$true`（疑いあり＝警告）。実機はトグル未設定＝ローミング ON（§7-4）。緩い判定に留める。 |
| `ConvertTo-SignatureHtml` | `-Text` | プレーンテキスト → 最小 HTML 文字列（下記§3）。 |
| `Write-SignatureFileSet` | `-SignaturesPath`, `-Name`, `-Text` | `<Name>.htm`（`ConvertTo-SignatureHtml` の結果, UTF-8 no BOM）と `<Name>.txt`（UTF-16LE BOM）を書き出し、生成パス配列を返す。フォルダが無ければ作成。`.rtf` は作らない。 |
| `Backup-SignatureFileSet` | `-SignaturesPath`, `-Name` | 既存の `<Name>.htm/.txt/.rtf` と `<Name>.files\` / `<Name>_files\`（両方の綴りに対応。実機は `.files`＝§7-6）を `<SignaturesPath>\.backup\<Name>-<yyyyMMdd-HHmmss>\` へ移動し、そのパスを返す。対象が無ければ `$null`。 |
| `Set-AccountSignatureRegistry` | `-AccountKeyPath`, `-SignatureName` | `New Signature` と `Reply-Forward Signature` を REG_SZ で設定。 |

## 3. プレーンテキスト → HTML 変換（ConvertTo-SignatureHtml）

- HTML エスケープ: `[System.Net.WebUtility]::HtmlEncode($Text)`（`&`,`<`,`>`,`"` を実体参照化）。
- 改行正規化: `\r\n` / `\r` / `\n` を一旦 `\n` に統一 → `<br>\r\n` に置換。
- ラップ（フォント指定はしない。Outlook の作成時フォントに従わせる）:

```html
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head><meta http-equiv="Content-Type" content="text/html; charset=utf-8" /></head>
<body><div>{escaped-body-with-br}</div></body>
</html>
```

## 4. エンコーディング / 改行

| ファイル | エンコーディング | 改行 | 実装 |
| --- | --- | --- | --- |
| `<Name>.htm` | **UTF-8 no BOM** ＋ `<meta http-equiv="Content-Type" content="text/html; charset=utf-8">` | CRLF | `[IO.File]::WriteAllText($p,$s,(New-Object Text.UTF8Encoding($false)))` |
| `<Name>.txt` | UTF-16 LE with BOM（`FF FE`） | CRLF | `[IO.File]::WriteAllText($p,$s,[Text.Encoding]::Unicode)` |

- 実機の既存署名は `.htm`=Shift_JIS(BOMなし) / `.txt`=UTF-16LE+BOM だった（§7-1）。`.txt` はこれに合わせる。
- `.htm` は Shift_JIS だと絵文字・CP932 外の文字が壊れるため **UTF-8 no BOM ＋ charset meta** を採用する。現行ビルド（16.0.20326）が meta を尊重して表示するかを **S13 で実機確認**し、ダメなら Shift_JIS へ切替。
- `Set-Content -Encoding` は PS 5.1 と 7 で挙動差があるため `[IO.File]::WriteAllText` に統一する。

## 5. CLI 仕様

### モジュール関数（主）

`Import-Module` 後、`Set-OutlookSignature` / `Get-OutlookSignature` を使う。ヘルプは comment-based help（`Get-Help Set-OutlookSignature -Full`）。「バージョン確認」は `Get-Module OutlookSignature -ListAvailable`。

| パラメータ | 必須 | 既定 | 説明 |
| --- | --- | --- | --- |
| `-Name` | | （省略）| 省略時: 現行の既定署名を in-place 更新。指定時: その名前で作成し既定に設定 |
| `-Text` | ○ | — | 署名本文（プレーンテキスト、複数行可、パイプライン可） |
| `-Force` | | off | Outlook 起動中でも続行 |
| `-WhatIf` / `-Confirm` | | — | `SupportsShouldProcess` |
| `-SignaturesPath` | | 実際のパス | テスト/上級者向け上書き |
| `-RegistryRoot` | | `HKCU:\Software\Microsoft\Office\16.0` | テスト/上級者向け上書き |

### 薄いラッパー scripts/Set-OutlookSignatureCli.ps1

`powershell -File` からの実行で意味のある終了コードを返すためのもの。

```powershell
param(
  [string]$Name,
  [Parameter(Mandatory)][string]$Text,
  [switch]$Force
)
```

module を import → `Set-OutlookSignature` 呼び出し → 例外の `FullyQualifiedErrorId` を終了コードへ写像 → 結果を表示。

### 終了コード（ラッパーのみ）

| コード | 意味 | 対応 ErrorId |
| --- | --- | --- |
| 0 | 成功 | — |
| 1 | 想定外エラー / 書き込み失敗 | `WriteFailed` ほか |
| 2 | プリフライト検証失敗 | `InvalidSignatureName`, `OutlookNotFound`, `DefaultProfileNotFound`, `ProfilePromptEnabled`, `NoMailAccount`, `NoCurrentSignature`, `SignaturesFolderNotWritable` |
| 3 | Outlook 起動中（`-Force` 未指定） | `OutlookRunning` |

## 6. エラーハンドリング方針

- Private 関数は失敗時に `throw`（メッセージのみ）。Public 関数の入口で `try/catch` し、`$PSCmdlet.ThrowTerminatingError()` で **ErrorId を付けた** 終了エラーに変換する。
- ErrorId 一覧（`ErrorRecord.FullyQualifiedErrorId` に載る）:
  `InvalidSignatureName` / `OutlookNotFound` / `DefaultProfileNotFound` / `ProfilePromptEnabled` /
  `NoMailAccount` / `NoCurrentSignature`（`-Name` 省略だが現行署名が空）/ `SignaturesFolderNotWritable` /
  `OutlookRunning` / `WriteFailed`
- 非終了の警告（`Write-Warning`、処理は続行）:
  - ローミング署名が有効の疑い
  - `-Force` 指定で Outlook 起動中（書き戻しの可能性）
  - 署名名が 32 文字超
- **プリフライトを全て通過してからのみ**ファイル/レジストリを変更する（部分適用しない = requirement の確定事項）。
- プリフライト通過後の想定外失敗（IO 競合・権限）に限り、ベストエフォートで生成ファイル削除＋バックアップ復元を試みてから `WriteFailed`。
- `Common\MailSettings`（`NewSignature` / `ReplySignature`）は **書き込まない**（Outlook 2010+ は per-account キーが実効値。型が REG_BINARY のバージョンがあり誤設定リスクが高い）。→ `doc/todo.md`。
- **専用の復元コマンドは phase 1 では作らない**。バックアップは `<SignaturesPath>\.backup\` に残し、手動復元とする。→ 必要になれば `doc/todo.md`。

## 7. 実機確認事項

実機（山田氏環境: 日本語 Windows 11 / Windows PowerShell 5.1 / クラシック Outlook Microsoft 365 Apps `16.0.20326.20112` Current Channel / **ローミング署名 ON**）で 2026-09-01 に一次調査済み。

| # | 項目 | 一次調査結果 | 状態（実機確認 2026-09-01） |
| --- | --- | --- | --- |
| 1 | `.htm` のエンコーディング | 既存署名は Shift_JIS（BOM なし, `charset=shift_jis`）。`.txt` は UTF-16LE+BOM | ✅ **UTF-8 no BOM + charset meta で問題なし**。in-place 更新後、新規メール／返信で `㈱ ① ～ — 髙` を含む本文が文字化けせず表示。§4 の方針を確定。Shift_JIS への切替は不要 |
| 2 | 対象アカウント判定 | メールアカウントでも `SMTP Address` は空。`Service Name`（`MSEMS` / `CONTAB`）で判別 | ✅ MSEMS アカウントのみ抽出（`AccountName=yamada@example.com`） |
| 3 | per-account 署名値の型 | `New Signature` / `Reply-Forward Signature` = REG_SZ（`String`） | ✅ 現行値 `default (yamada@example.com)` を読めた |
| 4 | ローミング署名の判定 | `DisableRoamingSignaturesTemporaryToggle` 未設定、`Common\Roaming` キー有り、署名名が `<名前> (<メール>)` 形式 → ローミング ON | ✅ トグル未設定＝ON の緩い判定で確定。**in-place 更新は Outlook 再起動後も保持され、クラウド同期で戻らなかった** |
| 5 | `PickLogonProfile` | `HKCU:\...\16.0\Outlook` 直下、実機は未設定 | ✅ `DefaultProfile=Outlook` を解決、プロンプト無し |
| 6 | 付随フォルダの綴り | 実機は `<名前>.files`（`_files` ではない） | ✅ バックアップは両綴り対応。`Write-SignatureFileSet` は `.files` を生成しない |
| 7 | `CreateAndAssign` の可視性 | — | ⚠️ **未検証**（in-place で用が足りたため）。ローミング ON 環境では新規署名がピッカーに出ない可能性が残る。README は `-Name` 省略の in-place 更新を推奨とし、`-Name` 指定は「Outlook 再起動後にピッカーを確認」と注記。実挙動の確認は `doc/todo.md` に残す |

結論: §4 のエンコーディング方針（`.htm`=UTF-8 no BOM + meta、`.txt`=UTF-16LE+BOM）と in-place 更新方式で、
山田氏の実機（ローミング署名 ON）で正しく反映・保持されることを確認済み。

## 8. パッケージング方針

- バージョンは `OutlookSignature.psd1` の `ModuleVersion` のみで管理（コード中にハードコードしない）。ラッパーや README はマニフェストを参照。
- 外部モジュール依存なし。Windows PowerShell 5.1 と PowerShell 7 の両方で動作（`[IO.File]` / `[Text.Encoding]` ベースで差異を回避）。
- 配布は `src/OutlookSignature/` をそのまま `Import-Module` する形（PSGallery 公開は対象外）。
- **ソースファイルの文字コード**: `.ps1` / `.psd1` / `.psm1` は **UTF-8 with BOM** で保存する。Windows PowerShell 5.1 は BOM 無しファイルをシステム ANSI（日本語環境では CP932）として読むため、BOM が無いと日本語コメント／文字列を含むファイルの解析に失敗する（`.psd1` は「restricted language file として無効」エラー）。エディタ設定または `scripts/` の一括変換で担保する。
- テスト（Pester v6）は BOM 依存しないが、上記に合わせ同様に UTF-8 BOM で保存する。

## 9. テスト方針（Pester v5 系。実機導入済み: 6.1.0）

Private 関数は `InModuleScope OutlookSignature { ... }` で直接呼ぶ。レジストリは Pester の `TestRegistry:` ドライブ（テストごとに使い捨てキー）を `-RegistryRoot` に渡す。ファイルは `TestDrive:` を `-SignaturesPath` に渡す。`Test-OutlookRunning` は `Mock`。

担保する項目:

- **ConvertTo-SignatureHtml**: `<`,`>`,`&`,`"` のエスケープ／改行が `<br>` になる／`charset=utf-8` を含む／複数行／空文字入力で最小 HTML。
- **Write-SignatureFileSet**: `.htm` と `.txt` のみ生成（`.rtf` を作らない）／`.htm` 先頭が BOM でない（`3C`＝`<`）／`.txt` 先頭が `FF FE`／改行が CRLF ／フォルダ自動作成。
- **Backup-SignatureFileSet**: 既存ファイルが `.backup\<name>-<timestamp>\` へ移動されパスが返る／対象が無ければ `$null` を返し何もしない／`<name>.files\` と `<name>_files\` の両方を移動。
- **Get-OutlookMailAccount**: 疑似レジストリツリーで `Service Name=MSEMS` のみ抽出／`CONTAB`（アドレス帳を模擬）を除外／`PreviousNew` `PreviousReplyForward` に現行値が入る／該当 0 件で空配列。
- **Resolve-OutlookEnvironment**: `Outlook` キー無し→`OutlookNotFound`／`DefaultProfile` 無し→`DefaultProfileNotFound`／`PickLogonProfile=1`→`ProfilePromptEnabled`／正常時に各パスが期待どおり。
- **Set-AccountSignatureRegistry**: 指定キーに `New Signature` と `Reply-Forward Signature` が REG_SZ で入る／既存値は上書き。
- **Set-OutlookSignature（結合、オーバーライド使用）**:
  - `CreateAndAssign`（`-Name` 指定）: 疑似アカウント複数すべてに両署名値が `<Name>` に設定／`Files` に 2 パス／`Mode='CreateAndAssign'`／`RegistryChanged=$true`／`RestartRequired=$true`。
  - `UpdateInPlace`（`-Name` 省略）: 現行 `New/Reply-Forward Signature` が指す署名ファイルのみ書き換え／レジストリ不変（`RegistryChanged=$false`）／`SignatureNames` が現行値集合と一致。
  - `UpdateInPlace` で現行署名が全アカウント空 → `NoCurrentSignature`。
  - `-WhatIf`: ファイル 0、レジストリ変更なし、戻り値に変更予定が分かる情報。
  - プリフライト各失敗（アカウント 0 件 / 署名名不正 / プロファイル無し）が対応 ErrorId で終了。
  - `Test-OutlookRunning` を `$true` にモック → `-Force` 無しで `OutlookRunning`／`-Force` 有りで警告つき続行。
  - 上書き時に既存ファイルがバックアップされ、新内容で置き換わる。
  - `RoamingSignaturesEnabled` が `Test-RoamingSignatureEnabled` の結果と一致。
- **scripts/Set-OutlookSignatureCli.ps1**: ErrorId → 終了コード写像（2 / 3 / 1）と正常時 0。

境界値・異常系: 空 `-Text`（許可、空署名）、署名名 33 文字（警告つき成功）、無効文字入りの署名名（`InvalidSignatureName`）、`.backup` へ二重実行（タイムスタンプで衝突しない）。

## 10. 月替わりメッセージ署名ラッパー（`scripts/Set-SeasonalOutlookSignature.ps1`）

入力: `doc/requirement.md`「月替わりメッセージ署名ラッパー」節。対象は本体と同じくクラシック版デスクトップのみ。

### 10.1 位置づけ

- `scripts/` に置く 2 本目のラッパー。`Set-OutlookSignatureCli.ps1` は変更しない。
- モジュール（`OutlookSignature.psd1`）を `Import-Module` して `Set-OutlookSignature` を呼ぶ。終了コード写像は `Set-OutlookSignatureCli.ps1` と同じ表を各スクリプトが自前で持つ（重複は許容。3 本目の利用者が現れたら共有ヘルパーへ抽出 → `doc/todo.md`）。
- メッセージ抽出・テンプレート展開のロジックはスクリプト内のローカル関数に分けるが、テストは `Set-OutlookSignatureCli.ps1` と同方式で `powershell -File` のサブプロセス実行＋疑似環境（`New-FakeOutlookRegistry` / `-SignaturesPath`）で検証する。内部関数単体の Pester は行わない。

### 10.2 パラメータ

| パラメータ | 型 | 必須 | 既定 | 説明 |
| --- | --- | --- | --- | --- |
| `-MessagesPath` | string | ○ | — | 月替わりメッセージファイル（1 行 1 メッセージ、1 行目＝1 月 … 12 行目＝12 月） |
| `-TemplatePath` | string | ○ | — | 署名テンプレート（プレースホルダを含むプレーンテキスト、複数行可） |
| `-Month` | int | | 実行時のローカル月 | 対象月。1〜12。範囲外はエラー（exit 2） |
| `-Placeholder` | string | | `{{message}}` | テンプレート内の差し込み位置マーカー |
| `-Name` | string | | （省略）| `Set-OutlookSignature` へ透過。省略時は in-place 更新 |
| `-Force` | switch | | off | `Set-OutlookSignature` へ透過 |
| `-WhatIf` | switch | | — | `SupportsShouldProcess` の共通変数として受理し透過 |
| `-SignaturesPath` | string | | （省略）| `Set-OutlookSignature` へ透過（テスト／上級者向け） |
| `-RegistryRoot` | string | | `HKCU:\Software\Microsoft\Office\16.0` | `Set-OutlookSignature` へ透過（テスト／上級者向け） |

- `[CmdletBinding(SupportsShouldProcess = $true)]`。
- `-Month` に `[ValidateRange(1,12)]` は付けない（範囲外を自前で検出して exit 2 に写像するため。属性で弾くとパラメータバインドエラー＝exit 1 になる）。指定有無は `$PSBoundParameters.ContainsKey('Month')` で判定。

### 10.3 ローカル関数

| 関数 | 入力 | 出力 / 責務 | throw する ErrorId |
| --- | --- | --- | --- |
| `Get-SeasonalMessageLine` | `-Path`, `-Month` | ファイルを読み（下記エンコーディング）、`\r?\n` で分割。`$Month` 行目（1 始まり）を `.Trim()` して返す | `MessagesFileNotFound`（存在しない／読めない）, `NoMessageForMonth`（行が無い／空白のみ） |
| `Expand-SignatureTemplate` | `-Path`, `-Placeholder`, `-Message` | テンプレートを読み、`String.Replace($Placeholder, $Message)`（全出現を置換）した文字列を返す。改行・その他の文字は不変 | `TemplateFileNotFound`（存在しない／読めない）, `PlaceholderNotFound`（`-Placeholder` を 1 つも含まない） |

- 読み取り: 事前に `Test-Path -LiteralPath` で存在確認 →無ければ該当 ErrorId で `throw`。存在すれば `[System.IO.File]::ReadAllText((Resolve-Path -LiteralPath $Path))`。.NET が UTF-8 / UTF-16 の BOM を判定し、BOM 無しは UTF-8 とみなす（requirement の「UTF-8、BOM 有無どちらでも」を満たす）。相対パスは `Resolve-Path` で絶対化。
- `Get-SeasonalMessageLine`: 分割後 `$lines.Count -lt $Month` または `[string]::IsNullOrWhiteSpace($lines[$Month-1])` なら `NoMessageForMonth`。13 行目以降は参照しないので自然に無視される。
- `Expand-SignatureTemplate`: `$template.Contains($Placeholder)` が偽なら `PlaceholderNotFound`。真なら `$template.Replace($Placeholder, $Message)`。

### 10.4 処理順

1. `-Month` 決定: 指定時はその値、未指定時は `(Get-Date).Month`。`1..12 -notcontains $month` なら `MonthOutOfRange` で `throw`。
2. `$message = Get-SeasonalMessageLine -Path $MessagesPath -Month $month`。
3. `$text = Expand-SignatureTemplate -Path $TemplatePath -Placeholder $Placeholder -Message $message`。
4. `Import-Module (Join-Path $PSScriptRoot '..\src\OutlookSignature\OutlookSignature.psd1') -Force`。
5. `Set-OutlookSignature` を splat で呼ぶ。`Text = $text`、`RegistryRoot = $RegistryRoot`、`WhatIf = $WhatIfPreference`。`-Name`（指定かつ非空のときのみ）、`-Force`、`-SignaturesPath`（指定時のみ）を条件付きで追加。
6. 対象月・使用メッセージ・組み立てた本文・`Set-OutlookSignature` の戻り値を表示 → `exit 0`。
7. 1〜5 で例外 → `[Console]::Error.WriteLine($_.Exception.Message)`、`FullyQualifiedErrorId` を §10.5 に従って終了コードへ写像して `exit`。

### 10.5 終了コード写像

`Set-OutlookSignatureCli.ps1` の写像にラッパー固有 ID を加えたもの。

| コード | 意味 | 対応 ErrorId |
| --- | --- | --- |
| 0 | 成功 | — |
| 1 | 想定外エラー / 書き込み失敗 | `WriteFailed` ほか上記以外すべて |
| 2 | プリフライト検証失敗 | `MessagesFileNotFound`, `TemplateFileNotFound`, `MonthOutOfRange`, `NoMessageForMonth`, `PlaceholderNotFound`, および `Set-OutlookSignatureCli.ps1` のプリフライト ID 一式（`InvalidSignatureName` / `OutlookNotFound` / `DefaultProfileNotFound` / `ProfilePromptEnabled` / `NoMailAccount` / `NoCurrentSignature` / `SignaturesFolderNotWritable`） |
| 3 | Outlook 起動中（`-Force` 未指定） | `OutlookRunning` |

### 10.6 エラーハンドリング方針

- ローカル関数（`Get-SeasonalMessageLine` / `Expand-SignatureTemplate`）は失敗時に `throw`。ErrorId はメッセージ中ではなく、スクリプト入口の `try/catch` で `$_.FullyQualifiedErrorId` を見て写像する。`throw` は `Write-Error -ErrorId` ではなく、`$PSCmdlet.ThrowTerminatingError()` または `throw [ErrorRecord]` で ErrorId を確実に付ける（`Set-OutlookSignatureCli.ps1` が `-match` で緩く判定しているのに合わせ、ここも `-match` 判定で可）。
- `$ErrorActionPreference = 'Stop'`。プリフライト（手順 1〜3）を全通過してから手順 4〜5 の変更に進む（部分適用しない）。
- `Set-OutlookSignature` 側のロールバック／バックアップはそのまま活用する。ラッパーで独自のロールバックはしない。

### 10.7 出力（例）

```
対象月        : 9
使用メッセージ : 朝晩は過ごしやすくなってまいりました。
--- 署名本文 ---
山田 太郎
一般社団法人 ○○

朝晩は過ごしやすくなってまいりました。

yamada@example.com
----------------
（続けて Set-OutlookSignature の戻り値を Format-List）
```

`-WhatIf` 時は `Set-OutlookSignature` が変更予定のみ表示。ラッパーは常に組み立てた本文を表示するので、差し込み結果を実行前に確認できる。

### 10.8 テスト方針（`tests/Set-SeasonalOutlookSignature.Tests.ps1`）

`Set-OutlookSignatureCli.Tests.ps1` と同じくサブプロセス実行（`Invoke-Cli` 相当のヘルパー）。`TestDrive:` にメッセージ／テンプレートファイルを作り、`New-FakeOutlookRegistry` ＋ `-SignaturesPath`（`TestDrive` 配下）で疑似環境を用意する。

- 正常時 `exit 0`。生成された `<署名>.txt` を読み、対象月メッセージがテンプレートに差し込まれていること。
- `-Month 3` でその月（3 行目）が使われる／`-Month` 未指定で当月（`(Get-Date).Month` 行目）が使われる。
- `-Placeholder '%X%'` でカスタムマーカーが置換される／マーカーが複数箇所あるとき全置換される。
- `MessagesFileNotFound` / `TemplateFileNotFound`（存在しないパス）→ `exit 2`。
- `MonthOutOfRange`（`-Month 13`, `-Month 0`）→ `exit 2`。
- `NoMessageForMonth`（対象行が空白のみ／行数不足）→ `exit 2`。
- `PlaceholderNotFound`（マーカー無しテンプレート）→ `exit 2`。
- 下流のプリフライト（`NoMailAccount` 等）→ `exit 2`。`OutlookRunning`（実 Outlook 起動時のみ、`-Skip`）→ `exit 3`。
- `-WhatIf` で署名ファイルが生成されず `exit 0`。
- 13 行目以降が無視される。BOM 付き UTF-8 のメッセージ／テンプレートでも文字化けしない。

### 10.9 ドキュメント（→ 実装計画のドキュメントステップ）

`README.md` に「月替わりメッセージで署名を毎月更新する」節を追加する。`messages.txt` / `template.txt` の書式例、実行例、タスクスケジューラ（毎月 1 日など）への登録例、終了コード表を含める。
