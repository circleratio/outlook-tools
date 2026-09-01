# 要求仕様: Outlook 署名更新 CLI

## 概要

MS365 の Outlook で使うメール署名を、CLI（PowerShell）から更新できるようにする。

対象は 2 系統ある。

| 系統 | 署名の保存先 | 更新手段 | 本フェーズでの扱い |
| --- | --- | --- | --- |
| クラシック版 Outlook デスクトップ（Win32 / Microsoft 365 Apps） | ローカル（`%APPDATA%\Microsoft\Signatures\` ＋ レジストリ） | ファイル生成 ＋ レジストリ書き込み | **今回実装する** |
| 新しい Outlook for Windows / Outlook on the web | Exchange Online（メールボックス） | Exchange Online PowerShell（`Set-MailboxMessageConfiguration`） | 将来対応。`doc/todo.md` を参照 |

将来的には両系統を 1 つのツールから扱えるようにしたいが、最初のステップではクラシック版デスクトップのみを対象とする。

あわせて、月替わりの一言メッセージを差し込んだ署名で毎月更新する用途に特化したラッパースクリプトを提供する（「月替わりメッセージ署名ラッパー」節。クラシック版デスクトップのみが対象）。

## 要求仕様（確定事項）

### 対象・前提

- 対象 Outlook: クラシック版 Outlook デスクトップ（Win32）。新しい Outlook / OWA は対象外。
- 実行環境: Windows + Windows PowerShell 5.1 以上（または PowerShell 7）。管理者権限は不要（`HKCU` と `%APPDATA%` のみ操作）。
- 対象 Office バージョン: Microsoft 365 Apps（レジストリキーは `16.0` 固定とみなす。他バージョン対応は将来課題）。
- 対象プロファイル: 既定の Outlook プロファイル（`HKCU\Software\Microsoft\Office\16.0\Outlook` の `DefaultProfile` 値で特定）。起動時にプロファイル選択が有効な環境は対象外（プリフライトでエラー）。
- 対象アカウント: 既定プロファイル内の送信可能なメールアカウントすべて（`9375CFF0413111d3B88A00104B2A6676` 配下のうち `Service Name` がメール送信系＝`MSEMS`/`IMAP`/`POP3` 等のサブキー。アドレス帳(`CONTAB`)・PST 等は除外。実機では `SMTP Address` が空のため `Service Name` で判定）。
- **ローミング署名（クラウド署名）が有効な場合の扱い**: ツールは検出して警告する（抑止はしない）。`-Name` 省略時は現在の（ローミング）署名を in-place で更新するため、有効環境でも反映される見込み。`-Name` 指定の新規署名作成は反映されないことがある。抑止オプションと確実な更新経路は `doc/todo.md` を参照。
  - 山田氏の実機はローミング署名 **有効**（一次調査 2026-09-01）。

### 実装形態

- PowerShell モジュール（`.psm1` ＋ `.psd1`）として実装する。
- 公開関数から署名の更新を行う。関数名は設計段階で決める（`doc/spec.md`）。
- 各公開関数に comment-based help を付け、`Get-Help` で使いかたが分かるようにする。
- テスト容易性のため、Signatures フォルダのパスとレジストリのルートを差し替え可能にする（内部パラメータ or 環境変数）。既定は実際のパスを使う。

### 署名の内容

- 署名の中身は**プレーンテキストを引数で直接指定**する（HTML テンプレートや差し込みは今回対象外）。
- 複数行のテキストを渡せること。
- 渡されたプレーンテキストから、Outlook が使う以下のファイルを生成する。
  - `<署名名>.txt` … プレーンテキストそのまま
  - `<署名名>.htm` … HTML エスケープし、改行を `<br>` に変換して最小限の HTML でラップ。`<meta charset>` を含める。
  - `.rtf` は生成しない（RTF 形式で作文する場合、Outlook は `.txt` をフォールバック利用する）。
- 生成先は `%APPDATA%\Microsoft\Signatures\`。フォルダが無ければ作成する。
- **文字コード（実機確認済み 2026-09-01）**: `.htm` は UTF-8 no BOM ＋ charset メタ、`.txt` は UTF-16 LE + BOM。山田氏の実機（ローミング署名 ON）で新規メール・返信に文字化けなく表示・保持されることを確認（`doc/spec.md` §4/§7）。

### 署名の識別

- 更新対象の署名は**署名名を引数で指定**する。ただし引数を省略した場合は、各アカウントの現在の既定署名（レジストリの `New Signature` / `Reply-Forward Signature` が指す署名）を更新対象とする（ローミング署名が有効な環境向けの基本動作）。
- 指定した名前の署名が既に存在する場合は上書き、無ければ新規作成する。
- 上書き時は、既存の `<署名名>.txt` / `<署名名>.htm` / `<署名名>.files\` / `<署名名>_files\`（あれば）を、タイムスタンプ付きのバックアップ先へ退避してから置き換える。バックアップの保存場所・世代管理・復元コマンドの要否は設計段階で決める。
- 署名名に Windows ファイル名として使えない文字が含まれる場合はエラー。

### 既定署名の割り当て

- `-Name` を指定して署名を作成した場合は、その署名を**新規メールと返信・転送の両方**の既定署名として対象アカウントすべてに割り当てる。
- `-Name` 省略（in-place 更新）の場合はレジストリを変更しない（現在の既定署名の本文だけ差し替える）。
- 書き込むレジストリ:
  - per-account（現行 Outlook の実効値。実機で REG_SZ を確認済み）:
    `HKCU\Software\Microsoft\Office\16.0\Outlook\Profiles\<profile>\9375CFF0413111d3B88A00104B2A6676\<account>` の
    `New Signature` / `Reply-Forward Signature`
  - `HKCU\Software\Microsoft\Office\16.0\Common\MailSettings` の `NewSignature` / `ReplySignature` は**書き込まない**（`doc/spec.md` §6、`doc/todo.md`）。

### 異常系（プリフライト検証）

- **書き込み前に以下を検証し、いずれか NG なら一切書き込まずにエラー終了する**（部分適用しない）:
  - 既定プロファイルがレジストリに存在する
  - 対象アカウントが 1 つ以上見つかる
  - Signatures フォルダが作成可能
  - 署名名が有効（`-Name` 指定時）
  - `-Name` 省略時、更新対象にできる現在の既定署名が 1 つ以上ある
- **Outlook が起動中の場合は中断する**（`-Force` 指定時のみ続行）。理由: Outlook は終了時にメール設定をレジストリへ書き戻すため、起動中の変更が失われうる。`-Force` 続行時はその旨を警告する。
- ローミング署名が有効と判定される場合は警告を出す（処理は続行）。

### 付帯仕様

- `-WhatIf` / `-Confirm` に対応する（`SupportsShouldProcess`）。
- 処理結果（生成したファイルのパス、更新したレジストリ値、警告、Outlook 再起動が必要な旨）を出力する。

## 月替わりメッセージ署名ラッパー（`scripts/Set-SeasonalOutlookSignature.ps1`）

### 目的・前提

- `Set-OutlookSignature`（および `scripts/Set-OutlookSignatureCli.ps1`）を、月替わりの一言メッセージを差し込んだ署名で毎月更新する用途に特化して呼び出すラッパー。
- 既存の `scripts/Set-OutlookSignatureCli.ps1` は変更せず、**別ファイル**として追加する。
- 対象は本ドキュメントのメイン機能と同じ（クラシック版 Outlook デスクトップ、Windows PowerShell 5.1 以上 / PowerShell 7、管理者権限不要）。

### 入力

- **月替わりメッセージファイル**（`-MessagesPath`、必須）
  - 1 行 1 メッセージのプレーンテキスト。上から順に 1 月 → 12 月に対応する（1 行目＝1 月、… 12 行目＝12 月）。1 メッセージは 1 行（改行を含められない）。
  - 文字コードは UTF-8（BOM 有無どちらでも可）。改行は CRLF / LF どちらでも可。
  - 各行は前後の空白を除去して使う。行が空（空白のみ）または対象月の行が存在しない場合、その月は「メッセージ欠落」とみなす（異常系を参照）。
  - 13 行目以降は無視する。
- **テンプレートファイル**（`-TemplatePath`、必須）
  - 署名の本体となるプレーンテキスト（複数行可）。月替わりメッセージを差し込む位置にプレースホルダを置く。
  - 既定のプレースホルダは `{{message}}`。`-Placeholder` で変更できる。テンプレート内に複数回現れる場合はすべて置換する。
  - 文字コード・改行の扱いはメッセージファイルと同じ。

### 挙動

1. 対象月を決める。`-Month <1-12>` 指定時はその月、省略時は実行時のシステムローカル日付の月。
2. メッセージファイルから対象月（N 行目）のメッセージを取得する。
3. テンプレート中のプレースホルダを対象月メッセージで置換し、署名本文（プレーンテキスト、複数行）を組み立てる。テンプレートの他の文字・改行・空白はそのまま保持する。
4. 組み立てた本文で署名更新処理を呼び、クラシック版 Outlook の署名を更新する（既定は `-Name` 省略の in-place 更新）。
5. 結果（対象月、使用したメッセージ、更新処理の戻り値、終了コード）を出力する。

### 引き継ぐオプション

`scripts/Set-OutlookSignatureCli.ps1` と同じ意味で以下を受け取り、そのまま署名更新処理へ渡す。

- `-Name` … 省略時は in-place 更新。指定時はその名前で署名を作成し、対象アカウントすべての既定署名に設定。
- `-Force` … Outlook 起動中でも続行。
- `-WhatIf` … 変更内容の確認のみ（組み立てた署名本文を確認できるようにする）。
- `-SignaturesPath` / `-RegistryRoot` … テスト／上級者向けの上書き。

### 異常系（プリフライト。署名を書き込む前に検証し、NG なら一切更新せずエラー終了）

- `-MessagesPath` / `-TemplatePath` のファイルが存在しない、または読み取れない。
- `-Month` が 1〜12 の範囲外。
- 対象月のメッセージが空、または対象月の行が存在しない。
- テンプレートにプレースホルダが 1 つも含まれない。

### 終了コード

`scripts/Set-OutlookSignatureCli.ps1` と同一体系。ラッパー固有のプリフライト失敗（ファイル無し・月範囲外・メッセージ欠落・プレースホルダ無し）は `2` に写像する。

| コード | 意味 |
| --- | --- |
| 0 | 成功 |
| 1 | 想定外エラー / 書き込み失敗 |
| 2 | プリフライト検証失敗（ラッパー固有の入力エラーを含む） |
| 3 | Outlook 起動中（`-Force` 未指定） |

### 使いかた

```powershell
# messages.txt : 1 行 1 メッセージで 1〜12 月分（12 行）
# template.txt : 署名本文。差し込み位置に {{message}} を置く

# 実行月のメッセージを差し込んで署名を更新（in-place）
powershell -File scripts/Set-SeasonalOutlookSignature.ps1 `
  -MessagesPath .\messages.txt -TemplatePath .\template.txt

# 翌月分を事前確認（実際には変更しない）
powershell -File scripts/Set-SeasonalOutlookSignature.ps1 `
  -MessagesPath .\messages.txt -TemplatePath .\template.txt -Month 10 -WhatIf

# Outlook 起動中でも強制実行
powershell -File scripts/Set-SeasonalOutlookSignature.ps1 `
  -MessagesPath .\messages.txt -TemplatePath .\template.txt -Force
```

## 使いかた

> 具体的な関数名・パラメータ名は設計（`doc/spec.md`）確定後にコマンド例を実態へ合わせて更新する。

```powershell
Import-Module ./src/OutlookSignature

# 現在使っている署名の本文だけ差し替える（-Name 省略 = in-place。ローミング署名環境の推奨）
Set-OutlookSignature -Text @"
山田 太郎
一般社団法人 ○○
yamada@example.com
"@

# 署名 "標準" を新規作成し、新規・返信転送の既定署名に設定する
Set-OutlookSignature -Name "標準" -Text "..."

# 変更内容の確認のみ（実際には変更しない）
Set-OutlookSignature -Text "..." -WhatIf

# Outlook 起動中でも強制実行（書き戻しリスクの警告つき）
Set-OutlookSignature -Text "..." -Force
```
