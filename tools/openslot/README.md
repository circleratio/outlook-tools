# openslot

MS365（Microsoft 365）の予定表から、指定期間の**空き時間**を抽出するツール。
日程調整メールへの返信に貼り付ける候補づくりが主な用途。
結果は標準出力に表示され、同時にクリップボードにコピーされる。

フロントエンドは 2 つあり、抽出ロジック（`src/OpenSlot`）を共有する。

| 版 | 実体 | カレンダー取得 | 使う場面 |
|---|---|---|---|
| **Graph 版** | `openslot.ps1` | Microsoft Graph PowerShell SDK（対話サインイン） | 端末でコマンドを直接叩く |
| **MCP 版** | Claude Code スキル `/openslot`（＋ `openslot-mcp.ps1`） | Claude の Microsoft 365 コネクタ | Claude Code 上でメール返信中に呼ぶ |

以下の抽出ルール・出力フォーマットは両版共通。

---

# Graph 版（`openslot.ps1`）

- 対象は平日（月〜金）のみ。土日は除外。
- 早朝（8:00 より前）・昼休み（12:00〜13:00）・夜（19:00 以降）は対象外。実質の対象時間帯は **8:00〜12:00 / 13:00〜19:00**。
- 「予定あり」と見なす表示区分は `busy` / `oof`（外出中）/ `tentative`（仮の予定）。`free` は空きとして扱う。
- 空き枠は 30 分刻みに内側丸めし、30 分未満の空きは出力しない。
- タイムゾーンは JST 固定。
- 祝日は現時点では検索対象に含める（将来オプション化予定）。

## 必要環境

| 項目 | 要件 |
|---|---|
| OS | Windows |
| PowerShell | Windows PowerShell 5.1 以上 |
| モジュール | `Microsoft.Graph.Authentication` |
| アカウント | MS365 アカウント。予定表への読み取り権限（`Calendars.Read`）を委任同意できること |

MS365 へは**コマンドを実行したユーザーの権限**でアクセスする。初回実行時にサインイン画面が表示される。

## インストール

このリポジトリを任意の場所に配置し、`Microsoft.Graph.Authentication` を導入する。

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
```

（`Microsoft.Graph` 一式が既に入っていれば追加インストールは不要。）

## 使いかた

```powershell
.\openslot.ps1 -Start 2026-09-01 -End 2026-09-05
```

| パラメータ | 説明 |
|---|---|
| `-Start` | 開始日。`yyyy-MM-dd` または年を省略した `MM-DD`。**省略可** — 省略時はコマンド実行日の翌日 |
| `-End` | 終了日（両端を含む）。`yyyy-MM-dd` または年を省略した `MM-DD`。**省略可** — 省略時は開始日の「次の金曜日」 |
| `-Week <数字[,数字...]>` | 対象期間を「`<数字>` 週後の月曜日 〜 `<数字>` 週後の金曜日」にする（基準は実行日を含む週）。各要素 0 以上の整数。`-Start` / `-End` とは併用不可 |
| `-Duration <分>` | 出力する空き枠の最小サイズ（分）。30 の倍数（30 以上）。既定 30 |
| `-NoClipboard` | クリップボードへのコピーを行わない |

- `MM-DD` で指定した場合、年はコマンド実行年（今年）を補う。
- 「次の金曜日」は、開始日が月〜木なら同じ週の金曜日、金曜日または土日開始の場合は翌週の金曜日。
- `-Week 0`（今週）は「翌日 〜 今週の金曜日」。`-Week 1` 以降は「その週の月曜 〜 金曜」。
  - 金曜以降に `-Week 0` を実行するとその週は対象日なし（出力なし）。
- `-Week 0,2,4` のようにカンマ区切りで複数の週を指定できる（重複は除き、週の昇順で連結して出力）。

```powershell
# 引数なし … 実行日の翌日から、その週の金曜日まで
.\openslot.ps1

# 開始日だけ指定 … 2026-09-01（火）から次の金曜 2026-09-04 まで
.\openslot.ps1 -Start 2026-09-01

# 年を省略 … 実行年の 9/1〜9/5
.\openslot.ps1 -Start 09-01 -End 09-05

# 2 週後の月曜日〜金曜日
.\openslot.ps1 -Week 2

# 今週（翌日〜金曜）・翌週・3 週後をまとめて
.\openslot.ps1 -Week 0,1,3

# 60 分以上まとまって空いている枠だけ
.\openslot.ps1 -Week 1 -Duration 60
```

既定値を適用した結果、終了日が開始日より前になる場合は何も出力せず正常終了する（`exit 0`）。
`-Start` と `-End` の両方を指定して終了日が開始日より前の場合は引数エラー（`exit 1`）。

初回はサインインを求められる。`Calendars.Read` に同意すると以降はキャッシュされたトークンで動作する。

### 出力例

```
9月1日(火): 8:00-10:00, 13:00-15:30, 17:00-19:00
9月2日(水): 8:00-12:00, 13:00-19:00
9月4日(金): 9:00-12:00
```

- 日付・時は 0 埋めしない。分は 2 桁。曜日は日本語。
- 空き時間が全く無い営業日、および土日は行を出さない。
- 空き枠は 30 分刻み。既定では 30 分以上の枠を出力し、`-Duration <分>`（30 の倍数）で最小サイズを変えられる。
- 上記と同じ全文がクリップボードにもコピーされる。

### 終了コード

| コード | 意味 |
|---|---|
| 0 | 正常終了（空きが 0 行、既定値適用で対象日なしの場合も 0） |
| 1 | 引数不正（日付書式不正・存在しない日付、`-Start`/`-End` 両方指定で終了日が開始日より前、`-Week` の空・負数や `-Start`/`-End` との併用、`-Duration` が 30 の倍数でない） |
| 2 | 認証失敗（サインイン中断・同意拒否など） |
| 3 | カレンダー取得失敗（Graph API エラー・ネットワーク断） |

ヘルプは `Get-Help .\openslot.ps1 -Full` で参照できる。

---

# MCP 版（Claude Code スキル `/openslot`）

Claude Code 上で、日程調整メールに返信する際に空き時間候補を素早く得るためのスキル。
カレンダー取得は Claude の **Microsoft 365 コネクタ**（`outlook_calendar_search`）を使うため、
`Microsoft.Graph` モジュールのインストールも実行のたびのサインインも不要。

## 必要環境

| 項目 | 要件 |
|---|---|
| 実行環境 | Claude Code |
| コネクタ | Microsoft 365 コネクタが接続済みであること |
| OS | Windows（`Set-Clipboard` を使う）。Windows PowerShell 5.1 以上 |

## セットアップ

1. このリポジトリを配置する（`.claude/skills/openslot/` がスキル本体）。
2. Claude Code で Microsoft 365 コネクタを接続する。
3. リポジトリのディレクトリで Claude Code を起動すると `/openslot` が使える。

## 使いかた

Claude Code のプロンプトで、期間を自然言語かオプションで伝える。Claude が
`openslot.ps1` と同じ引数体系（`-Start` / `-End` / `-Week` / `-Duration`）に変換して実行する。

```
/openslot -Week 1 -Duration 60
```

```
来週の空き時間を1時間枠で出して
```

内部の流れ:

1. `openslot-mcp.ps1 -ResolveOnly <引数>` … 対象期間を解決（`-Week` や「次の金曜」を含む）。
2. Claude が解決範囲で `outlook_calendar_search` を呼び、予定を取得（ページングを辿る）。
3. `openslot-mcp.ps1 -PlanPath <plan> -EventsPath <events>` … 空き行を算出して標準出力＋クリップボード。

出力・出力フォーマット・空き 0 行時の挙動は Graph 版と同じ。スキルは**空き時間の抽出まで**を行い、
メール本文の作成・送信はしない（必要なら結果を見てから別途 Claude に依頼する）。

## 終了コード（`openslot-mcp.ps1`）

| コード | 意味 |
|---|---|
| 0 | 正常終了（空き 0 行、対象日なしも 0） |
| 1 | 引数不正（Graph 版の 1 と同じ条件、または `-PlanPath` / `-EventsPath` の片方のみ指定） |
| 3 | 処理失敗（plan / events JSON の不正など） |

Graph 版の `2`（認証失敗）は無い。認証は Claude のコネクタ側で完結する。

---

## 開発者向け

### 構成

- `openslot.ps1` … Graph 版 CLI エントリポイント（引数検証・既定値解決・認証・終了コード変換・出力）
- `openslot-mcp.ps1` … MCP 版バックエンド（`-ResolveOnly` = 期間解決 / `-PlanPath -EventsPath` = 空き算出）
- `.claude/skills/openslot/SKILL.md` … MCP 版スキル本体（Claude への手順書）
- `src/OpenSlot/` … ロジック本体のモジュール。公開関数は
  `Get-FreeTime` / `Resolve-Period` / `ConvertTo-InputDate` / `ConvertFrom-CalendarSearchEvent`
- `tests/` … Pester テスト

### テスト実行

Pester 3.x（Windows 同梱）で実行する。

```powershell
Invoke-Pester -Path .\tests
```

`src/OpenSlot/` の純関数（日付文字列パース・既定期間の解決・日付範囲・区間マージ・空き算出・整形・
`outlook_calendar_search` 予定の変換）と `Get-FreeTime` のオーケストレーションを対象とする。
実際の MS365 通信を行う `Get-CalendarEvents`（Graph）と `openslot-mcp.ps1` / スキルは
単体テストせず、`doc/plan.md` の手動確認で担保する。

## ドキュメント

| ファイル | 内容 |
|---|---|
| [doc/requirement.md](doc/requirement.md) | 要求仕様（両版） |
| [doc/spec.md](doc/spec.md) | 設計（§8 が MCP 版） |
| [doc/plan.md](doc/plan.md) | 実装計画（末尾に MCP 版ステップ） |
| [doc/mcp-proposal.md](doc/mcp-proposal.md) | MCP 版の方式検討メモ（経緯） |
