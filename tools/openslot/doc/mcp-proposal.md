# openslot MCP 版 — 方式案

`openslot.ps1` はカレンダー取得に Microsoft Graph PowerShell SDK（`Connect-MgGraph` /
`Invoke-MgGraphRequest`）を直接使う。これを「MCP を使う版」に置き換えるための方式案をまとめる。
実装の意思決定用メモであり、確定仕様ではない（採用後に `doc/requirement.md` へ反映する）。

## 1. 目的の整理（なぜ MCP 版か）

想定している狙いを先に確定させる。案の評価軸が変わる。

- **A. 認証の手間をなくす** — `Microsoft.Graph.Authentication` の導入と、実行のたびの
  対話サインイン（`Calendars.Read` 同意）をやめ、既に接続済みの
  Microsoft 365 コネクタ（`claude_ai_Microsoft_365`）の認証を再利用したい。
- **B. openslot を Claude から呼べる道具にする** — 「来週の空きを出して」で Claude が
  openslot をツールとして起動できるようにしたい。
- **C. カレンダー取得部を疎結合にする** — データソースを差し替え可能にし、
  純ロジック（既存 `src/OpenSlot`）を Graph 非依存で再利用したい。

以下の案は A / C を主目的、B を副次目的として書く。

## 2. 現状分析 — Graph に触るのは 1 か所だけ

`src/OpenSlot` は `Get-CalendarEvents`（Private・副作用あり）以外すべて純関数。
`Get-FreeTime` がその `Get-CalendarEvents` を内部で直接呼んでいる
（`src/OpenSlot/Public/Get-FreeTime.ps1:24`）のが唯一の結合点。

```
openslot.ps1
  └ Connect-MgGraph                 ← 認証（Graph SDK）
  └ Get-FreeTime
       └ Get-CalendarEvents          ← /me/calendarView（Graph REST）★ここだけ差し替えれば済む
       └ （以降すべて純関数）
```

→ **どの案でも共通で、まず「イベント取得」を `Get-FreeTime` から外へ出す**（3 章）。

## 3. 共通の前提リファクタ（全案で必要）

`Get-FreeTime` を「イベントを注入できる」形にする。既存 CLI は既定引数で無改修動作。

```
Get-FreeTime [-Start] <datetime> [-End] <datetime> [-MinimumMinutes] <int> = 30
             [-Events <pscustomobject[]>]        # 省略時は従来どおり Get-CalendarEvents を呼ぶ
   -> [string[]]
```

- `-Events` が渡されたらそれを使い、`Get-CalendarEvents` を呼ばない。
- `-Events` の要素は既存と同じ `@{ Start=[datetime]; End=[datetime] }`（JST 壁時計・Kind=Unspecified）。
  「busy / oof / tentative のみ」「1 日にまたがる予定の切り出し」は**注入側の責務**にする
  （= Graph の `showAs` フィルタ相当を呼び出し側でやる）。
- あるいは `-EventProvider <scriptblock>`（`param($Start,$End)` を受けて区間配列を返す）にして
  取得方法をまるごと注入する方式でもよい。テスト観点では `-Events`（値注入）が単純。
- 既存 `tests/Get-FreeTime.Tests.ps1` は `Get-CalendarEvents` をスタブ化しているので、
  `-Events` 経由のケースを追加する。

この 1 点だけで「Graph 版」と「MCP 版」が同じ純ロジックを共有できる。

## 4. データソースの選択肢と制約（重要）

「MCP でカレンダーを取る」と一口に言っても、使える口が現状の仕様と噛み合わない。

| 手段 | 取れるもの | 制約 |
|---|---|---|
| `mcp__claude_ai_Microsoft_365__outlook_calendar_search` | 予定の一覧（subject / start・end の `{dateTime,timeZone}` / 出席者ほか） | **`showAs`（busy/oof/tentative/free）が要約に含まれない可能性が高い**。詳細は `read_resource` を予定ごとに呼ぶ必要（N+1）。1 ページ最大 25 件 + `offset` ページング。終日予定の扱い要確認 |
| `mcp__claude_ai_Microsoft_365__outlook_find_available_time` / `find_meeting_availability` | 「空いている候補スロット」（Graph `findMeetingTimes` ベース） | 固定 duration のランク付き候補で、**空き区間の網羅列挙ではない**。勤務時間に限定されない。入力は UTC。openslot の「全空き枠を 30 分刻みで」には不向き |
| 理想: Graph `/me/calendarView` または `/me/calendar/getSchedule`（free/busy） | showAs 付き予定 / free-busy マップ | **Claude の M365 コネクタには該当ツールが無い** |
| 専用 MS365 MCP サーバー（例: OSS の `ms-365-mcp-server` 系）を別途立てる | `list-calendar-view` 等で calendarView 相当 | サーバー導入と、そのサーバー自身の OAuth（デバイスコード等）が必要。Graph SDK モジュールは不要になる |

**結論**: Claude の M365 コネクタだけで現行仕様（showAs = busy/oof/tentative のみ、
30 分刻み網羅）を完全再現するのは難しい。取りうる道は

- (a) `calendar_search` + 予定ごとに `read_resource` で `showAs` を取得（呼び出し数増）
- (b) 仕様を緩める: **返ってきた予定はすべて busy 扱い**（free/tentative の区別を捨てる）
- (c) 専用 MS365 MCP サーバーを使い calendarView 相当を得る（認証は別途 1 回だけ）

まず (a) が本命。`calendar_search` の応答に実際に何が入るかを実アカウントで 1 回確認して確定する（8 章）。

## 5. 方式案

### 案C（推奨・短期）: Claude Code スキル + 既存純関数モジュール

MCP サーバーを新設せず、**Claude Code のスキル**として `openslot` を実装する。

フロー:
1. スキルが引数（`-Start/-End/-Week/-Duration` 相当）を受け取る。
2. 期間解決は既存 `Resolve-Period`（PowerShell 純関数）をそのまま使う。
3. 各期間について `mcp__claude_ai_Microsoft_365__outlook_calendar_search` で予定を取得
   （`afterDateTime`/`beforeDateTime` に期間、ページングを辿る）。必要なら `read_resource` で `showAs`。
4. 取得予定を `@{Start;End}` 配列に整形（JST 壁時計へ正規化、busy 判定）。
5. `Get-FreeTime -Start .. -End .. -MinimumMinutes $Duration -Events $events` を呼び、行を得る。
6. 標準出力へ表示し、`Set-Clipboard` でコピー（Claude Code はローカル PS が使えるのでクリップボード可）。

- **長所**: 新規インフラ最小。既存の純ロジックとテストをそのまま活用。認証は接続済み
  M365 コネクタを再利用（目的 A 達成）。クリップボードも動く。目的 B も自然に満たす。
- **短所**: 「Claude Code + M365 コネクタ接続済み」環境でのみ動く。単体の再利用可能な
  MCP サーバーにはならない。取得の細部（showAs、ページング、タイムゾーン）をスキルの
  手順書で担保する必要がある。

### 案A（推奨・中期の到達点）: 「純粋計算 MCP サーバー」

openslot を **カレンダーを取りに行かない** MCP サーバーにする。ツール例:

```
openslot_free_time({
  start?, end?, week?: number[], duration?: number,   // 期間指定は既存 CLI と同じ意味
  busyIntervals: [{ start: "2026-09-01T10:00:00", end: "..." }, ...]  // JST 壁時計
}) -> { lines: string[], text: string }
```

- サーバーは期間解決 + 空き算出 + 整形だけ（＝ `src/OpenSlot` の純関数群）。
- 呼び出し側（Claude）が M365 コネクタで予定を取得し `busyIntervals` として渡す。
- 実装は PowerShell モジュールを薄い MCP ラッパ（Node もしくは PS 用 MCP SDK）で包む。

- **長所**: openslot が Graph/MCP データソースから完全に独立。純粋関数なので
  テスト容易・言語非依存・どのカレンダー源でも使える。設定を Claude Code / Claude Desktop
  双方の `mcpServers` に置ける。
- **短所**: 「予定取得 → busy 整形」の段取りが呼び出し側（プロンプト or 上位スキル）に残る。
  クリップボードはサーバー責務にできない（別ツール or スキル側で）。

### 案B（非推奨）: 「自己完結 MCP サーバー」

openslot MCP サーバーが自分で予定取得までやる。

- B1: サーバーが Graph をアプリ登録（client credentials / device code）で直接叩く。
  → 対話サインインは消えるが、アプリ登録・シークレット管理が増える。MCP である必然性が薄い。
- B2: サーバーが上流の MS365 MCP サーバーに MCP クライアントとして接続してチェーンする。
  → 構成が複雑（MCP のクライアント兼サーバー）。

- **短所**: どちらも重い。認証問題を別の形で抱え直すだけになりがち。今回の目的には過剰。

## 6. 比較表

| 観点 | 案C（スキル） | 案A（純計算サーバー） | 案B（自己完結サーバー） |
|---|---|---|---|
| 新規実装量 | 小（手順書 + 整形ヘルパ） | 中（MCP ラッパ） | 大 |
| 認証の手間の解消（目的A） | ◎ 既存コネクタ再利用 | ◎ 呼び出し側が担当 | △ アプリ登録が必要 |
| Claude から呼べる（目的B） | ◎ | ○（上位で取得段取りが要る） | ○ |
| 純ロジックの Graph 非依存化（目的C） | ○ | ◎ | ○ |
| 再利用性（他クライアント/CI） | × Claude Code 前提 | ◎ | ○ |
| クリップボード | ◎ ローカル PS | △ 別手段 | △ 別手段 |
| 現行 `showAs` 仕様の再現 | 取得手順に依存（4章 a/b） | 呼び出し側次第 | ◎（calendarView 相当を実装すれば） |

## 7. 推奨

1. **まず 3 章のリファクタ**（`Get-FreeTime -Events`）を入れる。既存 CLI は無改修で動く。低リスク。
2. **案C を実装**して目的 A/B を最短で満たす。データ取得手順（4 章 (a)）をスキルに明記。
3. 単体の再利用可能サーバーが欲しくなった時点で **案A** に発展させる（純関数はそのまま流用）。
4. 案B は必要になるまで作らない。

`openslot.ps1`（Graph 直叩き版）は当面残す。MCP 版と両立できる（データソース違いの 2 経路）。

## 8. 未確認事項 / 次アクション

- [ ] `outlook_calendar_search` の応答に `showAs` / `isAllDay` が含まれるか実アカウントで確認。
      無ければ `read_resource` 併用か「全件 busy 扱い」（仕様変更）かを決める。
- [ ] `calendar_search` の `{dateTime,timeZone}` が JST 前提でよいか（メールボックスの TZ 設定）。
      異なる場合の JST 正規化方針を決める。
- [ ] 終日予定（`isAllDay`）が `calendar_search` でどう返るか。
- [ ] スキルの入力インターフェース（`openslot.ps1` の引数をそのまま踏襲するか）。
- [ ] クリップボード仕様を MCP 版でも維持するか（案A なら要検討）。
- [ ] `spec-driven-dev` スキルに沿って、方向性確定後に `doc/requirement.md` / `doc/spec.md` を更新。
