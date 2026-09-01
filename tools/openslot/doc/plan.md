# 実装計画

`doc/spec.md` を入力とした実装ステップ。

## 環境前提（実測）

- Windows PowerShell 5.1 のみ（`pwsh` / PowerShell 7 は無し）。
- Pester は 3.4.0（v5 は無し）。→ テストは **Pester 3.x 構文**（`Should Be` 等）で書く。
- `Microsoft.Graph.Authentication` 2.39.0 あり。→ `Connect-MgGraph` / `Invoke-MgGraphRequest` を利用可。

## ステップ

### S1. モジュール骨組み
- やること: `src/OpenSlot/OpenSlot.psd1`（`ModuleVersion=0.1.0`）、`OpenSlot.psm1`（`Private/` `Public/` を dot-source し `Public/` の関数を Export）を作成。
- 完了条件: `Import-Module ./src/OpenSlot/OpenSlot.psd1 -Force` が警告なく成功し、`Get-Command Get-FreeTime` / `Resolve-Period` / `ConvertTo-InputDate` が引ける（`psd1` の `FunctionsToExport` と一致）。
- 依存: なし。

### S2. 純関数（区間ユーティリティ）
- やること: `New-Interval`、`Merge-Interval`、`Get-FreeInterval`（内部にスロット丸めヘルパ）を実装。
- 完了条件: `tests/Merge-Interval.Tests.ps1` `tests/Get-FreeInterval.Tests.ps1` が全 It 成功。
- 依存: S1。S3・S4 と並行可。

### S3. 純関数（日付範囲・整形）
- やること: `Resolve-DateRange`、`Format-FreeDay` を実装。
- 完了条件: `tests/Resolve-DateRange.Tests.ps1` `tests/Format-FreeDay.Tests.ps1` が全 It 成功。
- 依存: S1。S2 と並行可。

### S3b. 純関数（既定期間の解決）
- やること: `Get-NextFriday`（Private）と `Resolve-Period`（Public）を実装。未指定日付は `[datetime]::MinValue`、
  `Week` 未指定は空配列（`[int[]]`）で表す。`Week` に値があるとき、`Sort-Object -Unique` した各週について
  `End = 実行週の月曜＋週×7＋4 日`、`Start = 週が 0 なら翌日 / 1 以上なら実行週の月曜＋週×7` を要素とする配列を返す
  （Start/End 引数は無視）。戻り値は常に配列。逆転（Week 0 を金曜以降に実行）はそのまま返す。
- 完了条件: `tests/Resolve-Period.Tests.ps1` が全 It 成功（`Get-NextFriday` の曜日別、`Resolve-Period` の省略パターン、
  `Week 0` の翌日開始・`Week>=1` の月曜開始・`Week 0` 金曜実行の逆転・単一/複数/重複 `-Week`）。
- 依存: S1。S2・S3 と並行可。

### S3c. 純関数（日付文字列パース）
- やること: `ConvertTo-InputDate`（Public）を実装。`yyyy-MM-dd` は `TryParseExact`、`MM-dd` は `^\d{2}-\d{2}$` 判定＋`[datetime]::new($Today.Year, …)`。不一致・不正日付は throw。
- 完了条件: `tests/ConvertTo-InputDate.Tests.ps1` が全 It 成功（両書式・実行年補完・うるう日・異常系）。
- 依存: S1。他と並行可。

### S4. Graph 取得
- やること: `Get-CalendarEvents` を実装（`/me/calendarView`、`Prefer` ヘッダで JST、`@odata.nextLink` ページング、`showAs` フィルタ）。
- 完了条件: 構文エラーなく dot-source でき、`Get-FreeTime` から呼べる（実通信は S7 で確認）。
- 依存: S1、S2（`New-Interval`）。

### S5. 公開関数 Get-FreeTime
- やること: `Resolve-DateRange` → `Get-CalendarEvents`（期間一括）→ 日毎に該当予定抽出 →
  `Get-FreeInterval -MinimumMinutes $MinimumMinutes` → `Format-FreeDay`、空行を除外して文字列配列を返す。
  `-MinimumMinutes`（既定 30）で最小空き枠を制御する。
- 完了条件: `Get-CalendarEvents` を差し替えた `tests/Get-FreeTime.Tests.ps1` が全 It 成功（`-MinimumMinutes` の効果を含む）。
- 依存: S2, S3, S4。

### S6. CLI ラッパ openslot.ps1
- やること: `-Start` `-End` `-Week`（`[int[]]`）`-Duration`（`[int]`、既定 30）`-NoClipboard`、comment-based help。
  先にモジュールを読み込む → `$today` 確定 → `-Duration` が 30 の倍数（30 以上）でなければ `exit 1` →
  `-Week` 指定時は `-Start`/`-End` との同時指定・空・負数を `exit 1` →
  指定分のみ `ConvertTo-InputDate -Today $today` でパース（`yyyy-MM-dd` / `MM-dd`）→
  両方指定で `End < Start` は `exit 1` → `$periods = @(Resolve-Period -Today $today (+Week))` で期間確定 →
  `Start -le End` の期間だけ残し、0 件なら空出力で `exit 0` → 認証（`Connect-MgGraph -Scopes Calendars.Read`）→
  期間ごとに `Get-FreeTime -MinimumMinutes $Duration` を実行して行を連結 → 標準出力＋`Set-Clipboard`。終了コード変換（0/1/2/3）。
- 完了条件:
  - `-Start 2026-13-01` で `exit 1`、`-Start 2026-09-05 -End 2026-09-01` で `exit 1`。
  - `-Start 09-01 -End 09-05`（`MM-dd`）が実行年で解釈される。
  - `-Week 1 -Start 2026-09-01` で `exit 1`（排他）、`-Week -1` で `exit 1`（負数）。
  - `-Duration 45` / `-Duration 20` / `-Duration 0` で `exit 1`。
  - `-Week 0,2` が 2 期間に解決される（`Resolve-Period` 経由。実通信は S7）。
  - 引数なし実行で `Resolve-Period` の既定値（翌日〜次の金曜）が使われる（実通信は S7）。
  - `-End 2000-01-01` のみ指定 → 空出力・`exit 0`。
  - `-?` でヘルプ表示。
- 依存: S5, S3b, S3c。

### S7. 手動動作確認
- やること: 実アカウントで以下を実行し、出力・終了コード・クリップボードを目視確認。
  ```powershell
  # 正常系（doc/requirement.md の使用例）
  .\openslot.ps1 -Start 2026-09-01 -End 2026-09-05
  $LASTEXITCODE                     # -> 0
  Get-Clipboard                     # -> 標準出力と同一

  # 既定値（開始日省略＝翌日、終了日省略＝次の金曜）
  .\openslot.ps1 ; $LASTEXITCODE                                    # -> 0（対象範囲は実行日次第）
  .\openslot.ps1 -Start 2026-09-01 ; $LASTEXITCODE                  # -> 0（2026-09-01〜2026-09-04）

  # 年省略（MM-DD、実行年を補う）
  .\openslot.ps1 -Start 09-01 -End 09-05 ; $LASTEXITCODE            # -> 0（実行年の 9/1〜9/5）

  # -Week（N週後の月曜〜金曜。0 は翌日〜今週金曜。カンマ区切りで複数週）
  .\openslot.ps1 -Week 0 ; $LASTEXITCODE                            # -> 0（翌日〜今週金曜）
  .\openslot.ps1 -Week 2 ; $LASTEXITCODE                            # -> 0（2週後の月〜金）
  .\openslot.ps1 -Week 0,1,3 ; $LASTEXITCODE                        # -> 0（3 週分を連結出力。0 は翌日開始）
  .\openslot.ps1 -Week 1 -Start 2026-09-01 ; $LASTEXITCODE          # -> 1（排他）

  # -Duration（最小空き枠。30 の倍数）
  .\openslot.ps1 -Week 1 -Duration 60 ; $LASTEXITCODE               # -> 0（60 分以上の枠のみ）
  .\openslot.ps1 -Duration 45 ; $LASTEXITCODE                       # -> 1（30 の倍数でない）

  # 異常系
  .\openslot.ps1 -Start 2026-09-05 -End 2026-09-01 ; $LASTEXITCODE   # -> 1
  .\openslot.ps1 -Start 2026/09/01 -End 2026-09-05 ; $LASTEXITCODE   # -> 1
  ```
- 完了条件: 出力が `M月d日(w): h:mm-h:mm, ...` 形式（日・時 0 埋めなし、分 2 桁、曜日日本語、土日と空き無し日は行なし）。既定値・異常系の終了コードが一致。

### S8. 最終確認
- やること: `Invoke-Pester ./tests` を全件実行。`doc/requirement.md` の要求仕様と実装を突き合わせる。副産物（一時ファイル等）を削除。
- 完了条件: テスト全 It 成功。要求仕様の各項目に対応する実装/テストが存在。

---

## MCP 版（スキル `openslot`）追加ステップ

`doc/spec.md` §8 を入力とする。S1〜S8（Graph 版）は完了済み。既存テストを壊さないこと。

### M1. `Get-FreeTime` に `-Events` を追加（リファクタ）
- やること: `Get-FreeTime` に `[pscustomobject[]]$Events` を追加。`$PSBoundParameters.ContainsKey('Events')` の
  ときは `Get-CalendarEvents` を呼ばず `$Events` を使う。既存の取得経路・シグネチャの後方互換を維持。
- 完了条件:
  - 既存 `tests/Get-FreeTime.Tests.ps1`（`Get-CalendarEvents` スタブ）が全 It 成功のまま。
  - 追加ケース: `-Events` に区間配列を渡すと `Get-CalendarEvents` を呼ばずにその区間で算出する／
    `-Events @()` で「予定なし」の結果になる。
- 依存: なし（既存モジュールのみ）。

### M2. `ConvertFrom-CalendarSearchEvent`（Public・純関数）
- やること: `src/OpenSlot/Public/ConvertFrom-CalendarSearchEvent.ps1` を実装（`doc/spec.md` §2 の仕様）。
  `OpenSlot.psm1` は `Public/*.ps1` を自動 export するため追加作業は基本不要だが、`OpenSlot.psd1` の
  `FunctionsToExport` に `ConvertFrom-CalendarSearchEvent` を追記する。
- 完了条件: `tests/ConvertFrom-CalendarSearchEvent.Tests.ps1` が全 It 成功
  （`doc/spec.md` §7 のケース表：showAs 採否・大文字小文字・`isCancelled`・`isAllDay`・壁時計パース・
  `timeZone` 非 Tokyo で warning・`End<=Start` 除外・空/`$null`・戻り値形式が `Get-CalendarEvents` と一致）。
  `Import-Module ./src/OpenSlot/OpenSlot.psd1 -Force` 後に `Get-Command ConvertFrom-CalendarSearchEvent` が引ける。
- 依存: なし。M1 と並行可。

### M3. `openslot-mcp.ps1`（2 フェーズ）
- やること: `doc/spec.md` §8.2 の通り実装。
  - 共通: `-Start/-End/-Week/-Duration/-NoClipboard`（`openslot.ps1` と同じ検証。検証部は可能なら共通化、
    無理なら同等ロジックを複製し spec と一致させる）。
  - `-ResolveOnly`: 検証 → `Resolve-Period` → `Start -le End` で絞り込み → プラン JSON を標準出力、`exit 0`。
    不正引数は `exit 1`。
  - `-PlanPath <p> -EventsPath <e>`: プラン JSON・予定 JSON を読む →
    `ConvertFrom-CalendarSearchEvent` → 期間ごとに `Get-FreeTime -Events` → 連結して標準出力＋`Set-Clipboard`
    （`noClipboard` 尊重）。`periods` 空なら空出力（＋空文字コピー）。JSON 不正は `exit 3`。
- 完了条件（手動・スタブ確認。S7 と同様に単体テストはしない）:
  - `./openslot-mcp.ps1 -ResolveOnly -Week 1 -Duration 60` が
    `periods` / `fetch` / `duration:60` / `noClipboard:false` を含む JSON を返す。
  - `./openslot-mcp.ps1 -ResolveOnly -Start 2026-09-05 -End 2026-09-01` → `exit 1`。
  - `./openslot-mcp.ps1 -ResolveOnly -Week 1 -Start 2026-09-01` → `exit 1`（排他）。
  - 手書きのプラン JSON＋予定 JSON（`outlook_calendar_search` 応答を模した数件）を与えて
    `-PlanPath/-EventsPath` を実行し、`M月d日(w): ...` 形式の行が出て、同内容がクリップボードに入る。
  - `periods:[]` のプラン → 標準出力空・`exit 0`・クリップボード空文字。
  - 壊れた JSON → `exit 3`。
- 依存: M1, M2。

### M4. スキル `.claude/skills/openslot/SKILL.md`
- やること: `doc/spec.md` §8.3 の手順を Markdown で記述。フロントマター（`name` / `description`）、
  期間の決め方、フェーズ1 実行、`outlook_calendar_search`（ページング）、予定 JSON 保存、フェーズ2 実行、
  結果提示、スコープ境界（メール本文は作らない）、スクラッチ後始末。
- 完了条件: `/openslot` 起動を想定した手順が、この会話の Claude が迷わず追える粒度で書かれている
  （実際の end-to-end 確認は M5）。
- 依存: M3。

### M5. 手動 end-to-end 確認
- やること: 実際のカレンダーに対しスキル相当の手順を実行する。
  ```
  # 1) 期間解決
  ./openslot-mcp.ps1 -ResolveOnly -Week 0
  # 2) 返ってきた fetch 範囲で outlook_calendar_search（nextOffset を辿る）→ events.json へ保存
  # 3) 空き算出
  ./openslot-mcp.ps1 -PlanPath plan.json -EventsPath events.json
  ```
  - Graph 版 `./openslot.ps1 -Week 0` の出力と、MCP 版の出力が一致することを確認（同一カレンダー・同一時刻）。
  - `-Duration 60` / 複数 `-Week`（例 `-Week 0,1`）でも一致すること。
  - クリップボード内容が標準出力と一致すること。
- 完了条件: 上記が一致。差分があれば原因（showAs 判定・タイムゾーン・ページング漏れ）を特定して修正。
- 依存: M4。

### M6. ドキュメント更新
- やること: `README.md` に MCP 版（スキル）の必要環境・セットアップ（コネクタ接続）・使いかた・
  Graph 版との違いを追記。`doc/mcp-proposal.md` は検討経緯として残す（本文からリンク）。
- 完了条件: README だけ読んで MCP 版スキルを使い始められる。`Invoke-Pester ./tests` 全 It 成功。
  スクラッチの一時 JSON を削除。
- 依存: M5。
