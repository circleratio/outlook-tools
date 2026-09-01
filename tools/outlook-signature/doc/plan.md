# 実装計画: Outlook 署名更新 CLI（クラシック版デスクトップ）

入力: `doc/spec.md`

## 依存関係の概要

```
S1 モジュール土台
      │
      ├─► S2 ConvertTo-SignatureHtml ─┐
      ├─► S3 Write-SignatureFileSet ──┤
      ├─► S4 Backup-SignatureFileSet ─┤
      ├─► S5 Resolve-OutlookEnvironment┤   （S2〜S8 は相互に独立。並行実装可）
      ├─► S6 Get-OutlookMailAccount ──┤
      ├─► S7 Test-* 2関数 ────────────┤
      └─► S8 Set-AccountSignatureRegistry┘
                     │
                     ▼
             S9 Set-OutlookSignature（結合）
                     │
             ┌───────┴───────┐
             ▼               ▼
   S10 Get-OutlookSignature  S11 CLI ラッパー
             └───────┬───────┘
                     ▼
             S12 ヘルプ整備
                     ▼
             S13 実機確認（spec §7 の残タスク）
                     ▼
             S14 手動動作確認
                     ▼
             S15 最終確認
```

各ステップ「実装 → 該当 Pester テストを書く → そのステップのテストが緑」を1単位とする。

---

## S1. モジュール土台

**やること**
- `src/OutlookSignature/OutlookSignature.psd1`（`ModuleVersion = '0.1.0'`、`RootModule`、`PowerShellVersion = '5.1'`、`FunctionsToExport = @('Set-OutlookSignature','Get-OutlookSignature')`、依存なし）。
- `OutlookSignature.psm1`: `Public\*.ps1` と `Private\*.ps1` をドットソース → `Export-ModuleMember -Function Set-OutlookSignature,Get-OutlookSignature`。
- `Public/` `Private/` に空の関数スタブ（`throw 'not implemented'`）を配置。
- `tests/` に Pester 実行の共通設定（各テスト冒頭で `Import-Module ./src/OutlookSignature -Force`）。
- Pester v5 系の導入確認 → **実機に 6.1.0 導入済み**（2026-09-01）。テストは Pester v5/6 構文で書く。

**完了条件**
- `Import-Module ./src/OutlookSignature -Force` が成功し、`Get-Command -Module OutlookSignature` に 2 関数が出る。
- `Invoke-Pester ./tests` が「0 失敗（テスト未追加）」で完走する。

## S2. ConvertTo-SignatureHtml（Private）

**やること**: `doc/spec.md` §3 の変換。HTML エスケープ（`WebUtility::HtmlEncode`）→ 改行正規化 → `<br>` → 最小 HTML ラップ。

**完了条件**: `tests/ConvertTo-SignatureHtml.Tests.ps1` が緑。
- `< > & "` が実体参照化される
- `\r\n` / `\n` / `\r` が `<br>` になる
- 出力に `charset=utf-8` と `<div>` が含まれる
- 空文字入力で例外なく最小 HTML を返す

## S3. Write-SignatureFileSet（Private）

**やること**: `spec` §4 のエンコーディングで `<Name>.htm`（**UTF-8 no BOM** ＋ charset meta）と `<Name>.txt`（UTF-16LE+BOM）を CRLF で `[IO.File]::WriteAllText` 出力。フォルダが無ければ作成。生成パス配列を返す。`.rtf` は作らない。

**完了条件**: `tests/Write-SignatureFileSet.Tests.ps1` が緑。
- `TestDrive:` 配下に `.htm` と `.txt` の 2 ファイルのみ生成（`.rtf` が無い）
- `.htm` 先頭バイトが BOM でない（`3C`＝`<`）、`.txt` 先頭が `FF FE`
- 本文中の改行が CRLF
- 存在しないサブフォルダ指定時に自動作成される

## S4. Backup-SignatureFileSet（Private）

**やること**: 既存 `<Name>.htm/.txt/.rtf` と `<Name>.files\` / `<Name>_files\` を `<SignaturesPath>\.backup\<Name>-<yyyyMMdd-HHmmss>\` へ移動。パスを返す。対象なしなら `$null`。

**完了条件**: `tests/Backup-SignatureFileSet.Tests.ps1` が緑。
- 既存ファイル群が backup フォルダへ移動し、元の場所から消える。戻り値が backup パス
- 対象ゼロなら `$null` を返し、何も作らない
- `<Name>.files\` と `<Name>_files\` の両方が移動される
- 同名で連続実行してもタイムスタンプで衝突しない（1 秒以内の連続実行はサフィックス付与などで回避）

## S5. Resolve-OutlookEnvironment（Private）

**やること**: `spec` §2 の表どおり。`-RegistryRoot`（既定 `HKCU:\Software\Microsoft\Office\16.0`）と任意の `-SignaturesPath` から `$env` を組み立て。異常は `throw`（`OutlookNotFound` / `DefaultProfileNotFound` / `ProfilePromptEnabled`）。

**完了条件**: `tests/Resolve-OutlookEnvironment.Tests.ps1` が緑（`TestRegistry:` に疑似ツリーを作成）。
- `Outlook` キー無し → `OutlookNotFound`
- `DefaultProfile` 値無し／プロファイルキー無し → `DefaultProfileNotFound`
- `PickLogonProfile = 1` → `ProfilePromptEnabled`
- 正常時、`ProfileKeyPath` / `AccountsKeyPath` / `SignaturesPath` が期待値

## S6. Get-OutlookMailAccount（Private）

**やること**: `-AccountsKeyPath` 配下のサブキーを走査し、`Service Name` がメール送信系（`MSEMS`/`IMAP`/`POP3`/`SMTP`/`MAPI`/`EAS`/`EXHTTP`）のものを `{AccountName; ServiceName; RegistryPath; PreviousNew; PreviousReplyForward}` で返す。

**完了条件**: `tests/Get-OutlookMailAccount.Tests.ps1` が緑。
- `Service Name=MSEMS` サブキーのみ抽出、`CONTAB`（アドレス帳の模擬）は除外
- `PreviousNew` / `PreviousReplyForward` に現行の署名値が入る（未設定なら空文字）
- サブキー 0 個／該当 0 個で空配列（例外なし）

## S7. Test-OutlookRunning / Test-RoamingSignatureEnabled（Private）

**やること**
- `Test-OutlookRunning`: `Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue` の有無を `[bool]` で返すだけ。
- `Test-RoamingSignatureEnabled`: `-RegistryRoot` 配下 `Outlook\Setup\DisableRoamingSignaturesTemporaryToggle` が `1` なら `$false`、それ以外 `$true`。

**完了条件**: `tests/Test-Helpers.Tests.ps1` が緑。
- `Test-RoamingSignatureEnabled`: トグル `1` で `$false`、未設定で `$true`、`0` で `$true`
- `Test-OutlookRunning`: `Mock Get-Process` で `$true`/`$false` を切替確認

## S8. Set-AccountSignatureRegistry（Private）

**やること**: `-AccountKeyPath` に `New Signature` と `Reply-Forward Signature` を REG_SZ で `Set-ItemProperty -Type String`。

**完了条件**: `tests/Set-AccountSignatureRegistry.Tests.ps1` が緑。
- `TestRegistry:` の疑似アカウントキーに両値が文字列型で設定される
- 既存値がある場合は上書きされる

## S9. Set-OutlookSignature（Public・結合）

**やること**: `spec` §2 の処理順 1〜13（2 モード: `UpdateInPlace` / `CreateAndAssign`）。`[CmdletBinding(SupportsShouldProcess=$true)]`。プリフライト（S5/S6/S7 + 対象署名名決定 + フォルダ検証 + 入力検証）→ `ShouldProcess` ガード下で 対象署名ごとに Backup(S4)→Write(S3←S2)、`CreateAndAssign` のみ Registry(S8) → 戻り値 `[pscustomobject]`。入口 `try/catch` で ErrorId 付き `ThrowTerminatingError`。プリフライト通過後の失敗はベストエフォートでロールバック後 `WriteFailed`。警告を `Write-Warning`。

**完了条件**: `tests/Set-OutlookSignature.Tests.ps1` が緑（`-RegistryRoot`＝`TestRegistry:`、`-SignaturesPath`＝`TestDrive:`、`Mock Test-OutlookRunning`）。
- `CreateAndAssign`（`-Name` 指定）: 疑似アカウント複数すべてに両署名値が `<Name>`／`Files` に 2 パス／`Mode='CreateAndAssign'`／`RegistryChanged=$true`／`RestartRequired=$true`／既存時のみ `BackupPaths` 非空
- `UpdateInPlace`（`-Name` 省略）: 現行 `New/Reply-Forward Signature` が指す署名のみ書き換え／`RegistryChanged=$false`／`SignatureNames` が現行値集合と一致
- `UpdateInPlace` で現行署名が全アカウント空 → `NoCurrentSignature`
- `-WhatIf`: ファイル 0・レジストリ変更 0
- プリフライト失敗（アカウント 0 / 無効な署名名 / プロファイル無し）が対応 ErrorId で終了
- `Test-OutlookRunning=$true`: `-Force` 無しで `OutlookRunning`、`-Force` 有りで警告つき成功
- 上書き時: 既存ファイルが backup され新内容で置換
- 空 `-Text` 成功、署名名 33 文字は警告つき成功、無効文字で `InvalidSignatureName`
- `RoamingSignaturesEnabled` が `Test-RoamingSignatureEnabled` 結果と一致

## S10. Get-OutlookSignature（Public）

**やること**: `spec` §2。Signatures フォルダの署名名一覧＋アカウントごとの現在の `New Signature` / `Reply-Forward Signature` を返す。

**完了条件**: `tests/Get-OutlookSignature.Tests.ps1` が緑。
- 疑似環境で署名ファイル名一覧とアカウント別現行値が返る
- S9 実行直後に呼ぶと、設定した署名名が全アカウントで返る（S9 と組み合わせた検証）

## S11. scripts/Set-OutlookSignatureCli.ps1

**やること**: `param([string]$Name,[Parameter(Mandatory)][string]$Text,[switch]$Force)`（`-Name` は任意）→ `Import-Module` → `Set-OutlookSignature`（`$Name` 指定時のみ渡す）→ `catch` で `FullyQualifiedErrorId` を終了コード（`spec` §5 表）へ写像 → 結果表示。comment-based help 付き。

**完了条件**: `tests/Set-OutlookSignatureCli.Tests.ps1` が緑。
- 正常時 `exit 0`（`-Name` 有り／無しの両方）
- `NoMailAccount` / `NoCurrentSignature` 等 → `exit 2`
- `OutlookRunning` → `exit 3`
- `WriteFailed` → `exit 1`
（テストは `powershell -File` でサブプロセス実行し `$LASTEXITCODE` を確認。`Set-OutlookSignature` は `-RegistryRoot`/`-SignaturesPath` 越しに疑似環境で実行、または失敗系はモック用フックで）

## S12. ヘルプ整備

**やること**: `Set-OutlookSignature` / `Get-OutlookSignature` / ラッパーに comment-based help（`.SYNOPSIS` `.PARAMETER` `.EXAMPLE` `.NOTES`（再起動要・ローミング署名注意））。

**完了条件**: `Get-Help Set-OutlookSignature -Full` に全パラメータと 2 つ以上の例が出る。`tests` に「`Get-Help` の `.Examples.Example.Count -ge 2`」チェックを追加し緑。

## S13. 実機確認（`spec` §7 の残タスク）

一次調査で §7 の #2/#3/#5/#6 は確定済み。実機（山田氏環境, ローミング署名 ON）で残りを確認し、結果を `doc/spec.md` §7 と各関数へ反映する。**この確認は Outlook を一度終了して行う**。事前に既存署名（`default (yamada@example.com)` 等）を `Backup-SignatureFileSet` 相当で退避しておき、確認後に戻せるようにする。

1. **#1 `.htm` エンコーディング**: `UpdateInPlace` で既存署名を「UTF-8 no BOM + charset meta」の `.htm` に書き換え → Outlook 起動 → 新規メール／返信で**文字化けなく**表示されるか。NG なら `Write-SignatureFileSet` と `ConvertTo-SignatureHtml` を Shift_JIS 出力へ切替（`charset=shift_jis`、`[Text.Encoding]::GetEncoding(932)`）。
2. **#4 in-place 更新の同期**: 上記の書き換えが Outlook 再起動後に保持されるか（クラウド同期で元に戻らないか）。数分待って再確認。戻る場合は README に「反映されないことがある。確実には新 Outlook / OWA 側（`doc/todo.md`）」と明記。
3. **#6 `.files` フォルダ**: `Write-SignatureFileSet` の最小 HTML では `<Name>.files\` が新規生成されないこと。生成される場合はバックアップ／復元の対象に含める（既に両綴り対応済み）。
4. **#7 `CreateAndAssign` の可視性**: テスト用の名前（例 `zzz-test`）で新規署名を作り既定に設定 → Outlook 起動 → 署名ピッカーに出るか／新規メールに挿入されるか。出ない場合は README で `-Name` 省略運用を推奨とする。**確認後はこのテスト署名を削除し、元の設定へ戻す**。

**完了（2026-09-01）**: `scripts/Invoke-S13Check.ps1` で実施。
- #1 `.htm`=UTF-8 no BOM+meta で文字化けなし → §4 確定、Shift_JIS 切替不要
- #2/#3/#5/#6 実機確認済み
- #4 in-place 更新は Outlook 再起動後も保持（ローミング同期で戻らず）
- #7 `CreateAndAssign` の可視性のみ未検証 → `doc/todo.md` へ
- 実機の署名は `-Restore` で確認前の状態へ復元済み。コード変更は無し（テスト更新不要）。

## S14. 手動動作確認

**やること**: `doc/requirement.md` の「使いかた」のコマンドを実際に実行（実機がある場合）。無い場合は `-WhatIf` と `-SignaturesPath`/`-RegistryRoot` オーバーライドで疑似環境実行し出力を目視。

```powershell
Import-Module ./src/OutlookSignature -Force

# 1) WhatIf（疑似環境）: 変更予定の表示のみ
Set-OutlookSignature -Name "標準" -Text "山田 太郎`n一般社団法人 ○○" -WhatIf `
  -SignaturesPath "$env:TEMP\sigtest\Signatures" -RegistryRoot "HKCU:\Software\__sigtest__\16.0"

# 2) 疑似環境で実行 → Get で確認
#   （事前に __sigtest__ に Outlook/DefaultProfile/Profiles/<p>/<GUID>/<acct>(Service Name=MSEMS) を用意）
Set-OutlookSignature -Name "標準" -Text "..." -SignaturesPath ... -RegistryRoot ...   # CreateAndAssign
Set-OutlookSignature -Text "..." -SignaturesPath ... -RegistryRoot ...                # UpdateInPlace
Get-OutlookSignature -SignaturesPath ... -RegistryRoot ...

# 3) 実機: Outlook 終了 → in-place 更新 → 起動して署名確認（事前バックアップ必須）
Set-OutlookSignature -Text @"
山田 太郎
一般社団法人 ○○
yamada@example.com
"@

# 4) ラッパーの終了コード
powershell -File scripts/Set-OutlookSignatureCli.ps1 -Text "..."; $LASTEXITCODE

# 5) 異常系: Outlook 起動中に -Force 無し → 停止・exit 3
```

**完了条件**: 上記 1〜5 の各出力が `spec` の仕様（戻り値プロパティ、警告文、終了コード、再起動要メッセージ）と一致することを目視確認。実機分は署名が Outlook 上で正しく表示され、確認後に元へ復元。

## S15. 最終確認

**やること**
- `Invoke-Pester ./tests -Output Detailed` 全緑（失敗 0）。
- `doc/requirement.md` の「要求仕様（確定事項）」を 1 項目ずつコード/挙動と突合。乖離があれば修正、または requirement/spec/todo を実態へ更新。
- `doc/spec.md` §7 の各項目の状態（確認済み / 実機未確認で todo 移送）を最終化。
- 生成副産物の削除（テストの一時ファイル、`$env:TEMP\sigtest` 等。`.backup\` はリポジトリに含めない＝`.gitignore` 不要だが実機の実フォルダは触らない）。

**完了条件**
- Pester 全緑のログを提示。
- requirement の確定事項リストに対する「実装済み / 仕様どおり」の対応表を提示。
- 未解決事項が `doc/todo.md` に漏れなく載っている。
- 次段（S5: ドキュメント作成 = `README.md`）へ渡せる状態。

## S16. 月替わりメッセージ署名ラッパー（`doc/spec.md` §10）

**やること**
- `scripts/Set-SeasonalOutlookSignature.ps1`（UTF-8 BOM）を新規作成。`Set-OutlookSignatureCli.ps1` は変更しない。
  - `[CmdletBinding(SupportsShouldProcess = $true)]`、パラメータは spec §10.2 の表。
  - ローカル関数 `Get-SeasonalMessageLine` / `Expand-SignatureTemplate`（spec §10.3）。失敗時は `throw "... [ErrorId]"`（本体と同じ `[ErrorId]` サフィックス方式）。
  - 処理順 spec §10.4。終了コード写像 spec §10.5（`Set-OutlookSignatureCli.ps1` の写像＋固有 ID）。
  - comment-based help（`.SYNOPSIS` `.DESCRIPTION` `.PARAMETER` `.EXAMPLE`）。
- `tests/Set-SeasonalOutlookSignature.Tests.ps1`（UTF-8 BOM）を新規作成。spec §10.8 の項目。サブプロセス実行＋`New-FakeOutlookRegistry`＋`TestDrive:`。

**完了条件**
- `Invoke-Pester ./tests` 全緑（既存 + 新規）。
- 疑似環境で `-Month` 指定／未指定・カスタムプレースホルダ・各異常系（exit 2）・`-WhatIf`（exit 0・ファイル未生成）が仕様どおり。
- 生成された `<署名>.txt` に対象月メッセージが差し込まれていることを読み出しで確認。

**依存**: S9（`Set-OutlookSignature`）完了済み。他ステップと独立。

**完了（2026-09-01）**: `scripts/Set-SeasonalOutlookSignature.ps1` ＋ `tests/Set-SeasonalOutlookSignature.Tests.ps1`（15 件）実装。`Invoke-Pester ./tests` 66 passed / 0 failed。疑似環境で `-Month` 指定／未指定・カスタムプレースホルダ全置換・各異常系 exit 2・`-WhatIf`（exit 0・ファイル未生成）を確認。

## S17. ドキュメント更新（`README.md`）

**やること**: `README.md` に「月替わりメッセージで署名を毎月更新する」節（`messages.txt` / `template.txt` 書式、実行例、タスクスケジューラ登録例、終了コード表）。`doc/` リンク表は変更不要（ファイル増なし）。

**完了条件**: README の手順どおりに実行して署名が更新できる（疑似環境で追試）。

**完了（2026-09-01）**: README に「月替わりメッセージで署名を毎月更新する」節（入力ファイル書式・実行例・タスクスケジューラ `schtasks /SC MONTHLY` 例・終了コード表）を追加。

---

**完了（2026-09-01, S1〜S15）**: Pester 50 passed / 1 skipped（CLI exit-3 テストは実 Outlook 起動が前提のため、
未起動時はスキップ。Outlook 起動中に実行した回では pass 済み）/ 0 failed。副産物（一時キー・
`Signatures.s13-stash`・TEMP 配下）を削除。`README.md` 作成済み。requirement 対応表は下記。

| requirement の確定事項 | 実装 |
| --- | --- |
| クラシック版のみ・HKCU/%APPDATA% のみ・管理者不要 | `Resolve-OutlookEnvironment` ほか全て HKCU/AppData |
| `16.0` 固定・`DefaultProfile` で既定プロファイル特定・`PickLogonProfile=1` は対象外 | `Resolve-OutlookEnvironment`（`ProfilePromptEnabled`） |
| 対象アカウント = `Service Name` がメール送信系 | `Get-OutlookMailAccount`（allowlist） |
| ローミング署名は検出して警告のみ | `Test-RoamingSignatureEnabled` ＋ `Set-OutlookSignature` の警告 |
| PowerShell モジュール・comment-based help・パス/レジストリ差し替え可 | `OutlookSignature.psd1/.psm1`、`-SignaturesPath`/`-RegistryRoot` |
| プレーンテキスト入力・複数行・`.htm`+`.txt` 生成・`.rtf` は作らない | `Write-SignatureFileSet` / `ConvertTo-SignatureHtml` |
| `.htm`=UTF-8 no BOM+meta / `.txt`=UTF-16LE+BOM（実機確認済み） | `Write-SignatureFileSet` |
| `-Name` 省略で in-place、指定で作成＋既定化 | `Set-OutlookSignature`（2 モード） |
| 上書き時バックアップ（`.files`/`_files` 両対応）、復元は手動 | `Backup-SignatureFileSet` |
| `-Name` 指定時のみ両署名値を全アカウントに設定、`Common\MailSettings` は書かない | `Set-AccountSignatureRegistry` |
| プリフライト検証・部分適用しない・失敗時ロールバック | `Set-OutlookSignature` |
| Outlook 起動中は中断（`-Force` で続行＋警告） | `Test-OutlookRunning` ＋ `OutlookRunning` |
| `-WhatIf`/`-Confirm`・処理結果を出力 | `SupportsShouldProcess`＋戻り値 pscustomobject |
