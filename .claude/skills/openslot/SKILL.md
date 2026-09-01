---
name: openslot
description: >-
  日程調整メールへの返信に貼り付ける「空き時間候補」を作る。MS365 予定表を Claude の
  Microsoft 365 コネクタ（outlook_calendar_search）経由で読み、指定期間の平日 8:00-12:00 /
  13:00-19:00 から空き枠を 30 分刻みで抽出して標準出力＋クリップボードへ出す。期間指定は
  openslot.ps1 と同じ引数体系（-Start / -End / -Week / -Duration）。ユーザーが「空き時間を出して」
  「来週の空きを教えて」「日程調整の返信」等と言ったら使う。メール本文の作成・送信は行わない。
---

# openslot（MCP 版）

MS365 予定表から空き時間を抽出する。カレンダー取得は `outlook_calendar_search`（Claude が実行）、
日付ロジックとクリップボードは `openslot-mcp.ps1`（PowerShell）が担う。抽出ロジックは Graph 版
`openslot.ps1` と共通（`src/OpenSlot`）。

- 前提: Microsoft 365 コネクタが接続済みであること。Windows（`Set-Clipboard` を使う）。
- スコープ: **空き時間の抽出まで**。メール本文への反映・下書き・送信はしない。
  ユーザーがそのあと明示的に頼んだ場合のみ、別途対応する。
- 作業ディレクトリはリポジトリの `tools/openslot/`（`openslot-mcp.ps1` がある場所）。
- 一時ファイルはセッションのスクラッチディレクトリに置き、完了後に削除する。

## 手順

### 1. 期間と最小枠を決める

ユーザーの言葉を `openslot.ps1` の引数体系に変換する。

| ユーザーの言い方（例） | 引数 |
|---|---|
| 指定なし / 「直近の空き」 | （引数なし）= 翌日〜その週の金曜 |
| 「9/1〜9/5」 | `-Start 09-01 -End 09-05` |
| 「9/1から」 | `-Start 09-01`（終了は次の金曜） |
| 「今週の残り」 | `-Week 0` |
| 「来週」 | `-Week 1` |
| 「再来週の月〜金」 | `-Week 2` |
| 「今週と来週」 | `-Week 0,1` |
| 「1時間以上まとまって」 | `-Duration 60` |

- 年は省略形（`MM-dd`）で渡すと実行年を補う。`yyyy-MM-dd` も可。
- `-Week` は `-Start` / `-End` と併用不可。
- どの期間か判断できないときは、勝手に決めずユーザーに確認する。

### 2. フェーズ1: 期間解決

`tools/openslot/` で実行する（`<引数>` は手順1で決めたもの）:

```
powershell -NoProfile -File .\openslot-mcp.ps1 -ResolveOnly <引数>
```

標準出力は次の形の JSON:

```json
{
  "periods":  [ { "start": "2026-09-07", "end": "2026-09-11" } ],
  "fetch":    { "start": "2026-09-07", "end": "2026-09-12" },
  "duration": 60,
  "noClipboard": false
}
```

- 終了コード 1 … 引数不正。エラーメッセージをユーザーに伝えて終了。
- `periods` が空配列（`fetch` が `null`）… 対象日なし。「対象期間に平日がありません」と伝えて終了。
- この JSON をスクラッチに `plan.json` として保存する。

### 3. 予定を取得する

`plan.json` の `fetch` 範囲で `outlook_calendar_search` を呼ぶ:

- `query`: `"*"`
- `afterDateTime`: `fetch.start`
- `beforeDateTime`: `fetch.end`
- `limit`: `25`
- 応答末尾に `nextOffset` があれば、それを `offset` に渡して**全ページ取得**する。

返ってきた予定オブジェクト（`showAs` / `isAllDay` / `isCancelled` / `start` / `end` を含む）を
**そのまま** JSON 配列にまとめ、スクラッチに `events.json` として保存する。1 件も無ければ `[]`。
（busy/oof/tentative の絞り込み・キャンセル除外・終日予定の扱いは `openslot-mcp.ps1` 側で行うので、
ここでは加工しない。）

### 4. フェーズ2: 空き算出

```
powershell -NoProfile -File .\openslot-mcp.ps1 -PlanPath <plan.json のパス> -EventsPath <events.json のパス>
```

- 標準出力の各行（`M月d日(w): h:mm-h:mm, ...` 形式）をそのままユーザーに提示する。
- 同じ全文がクリップボードにコピーされている旨を一言添える。
- 空き 0 行なら「指定期間に空き枠はありませんでした」と伝える。
- 終了コード 3 … 処理失敗。メッセージを伝える。

### 5. 後始末

`plan.json` / `events.json` を削除する。
