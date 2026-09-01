# 設計

`doc/requirement.md` を入力とした実装方針。対象は PowerShell 5.1 / 7.x、Windows 限定。

Graph 版（`openslot.ps1`）と MCP 版（Claude Code スキル）の 2 フロントエンドが
抽出ロジック `src/OpenSlot` を共有する。§1〜§7 は共通・Graph 版の設計、§8 が MCP 版の追加設計。

## 1. ディレクトリ構成

```
tools/openslot/
  openslot.ps1                  # Graph 版 CLI エントリポイント（薄いラッパ）
  openslot-mcp.ps1              # MCP 版バックエンド（期間解決フェーズ / 空き算出フェーズ）
  .claude/skills/openslot/
    SKILL.md                    # MCP 版スキル本体（Claude への手順書）
  src/
    OpenSlot/
      OpenSlot.psd1             # モジュールマニフェスト（バージョンの一元管理）
      OpenSlot.psm1             # 下記 Public/Private を dot-source して公開関数を Export
      Public/
        Get-FreeTime.ps1        # オーケストレーション（公開関数）。-Events でイベント注入可
        Resolve-Period.ps1      # 省略された開始日・終了日に既定値を補う（純関数・公開関数）
        ConvertTo-InputDate.ps1 # 日付文字列（yyyy-MM-dd / MM-dd）を [datetime] に変換（純関数・公開関数）
        ConvertFrom-CalendarSearchEvent.ps1  # outlook_calendar_search の生予定 → busy 区間（純関数・公開関数。MCP 版）
      Private/
        New-Interval.ps1        # [datetime]Start/[datetime]End の pscustomobject 生成
        Resolve-DateRange.ps1   # 対象平日リストの算出
        Get-NextFriday.ps1      # 指定日の「次の金曜日」を算出（純関数）
        Get-CalendarEvents.ps1  # Graph から予定取得（副作用あり。Graph 版のみ）
        Merge-Interval.ps1      # 区間のマージ（純関数）
        Get-FreeInterval.ps1    # 1 日分の空き区間算出（純関数）
        Format-FreeDay.ps1      # 1 日分の出力行整形（純関数）
  tests/
    Resolve-DateRange.Tests.ps1
    Resolve-Period.Tests.ps1
    ConvertTo-InputDate.Tests.ps1
    ConvertFrom-CalendarSearchEvent.Tests.ps1
    Merge-Interval.Tests.ps1
    Get-FreeInterval.Tests.ps1
    Format-FreeDay.Tests.ps1
    Get-FreeTime.Tests.ps1
  doc/{requirement,spec,plan,mcp-proposal}.md
  README.md
```

- 純関数（`ConvertTo-InputDate` / `Resolve-Period` / `Get-NextFriday` / `Resolve-DateRange` /
  `ConvertFrom-CalendarSearchEvent` / `Merge-Interval` / `Get-FreeInterval` / `Format-FreeDay`）と
  副作用を持つ関数（`Get-CalendarEvents`）を分離し、前者を Pester で単体テストする。
- `openslot.ps1` / `openslot-mcp.ps1` はパラメータ定義・検証・終了コード変換・出力/クリップボードのみを担当する。

## 2. モジュールごとの責務と関数シグネチャ

### New-Interval （Private・純関数）
```
New-Interval [-Start] <datetime> [-End] <datetime> -> [pscustomobject] @{ Start; End }
```
区間の生成のみ。`End <= Start` の場合は $null を返す（呼び出し側で除外）。

### ConvertTo-InputDate （Public・純関数）
```
ConvertTo-InputDate [-Text] <string> [-Today] <datetime> -> [datetime]
```
- `Text` が `yyyy-MM-dd` に一致 → その日付（`.Date`）。
- `Text` が `^\d{2}-\d{2}$`（`MM-dd`）に一致 → `[datetime]::new($Today.Year, MM, DD)`。年は実行年。
- どちらにも一致しない、または存在しない日付（実行年の `02-29` など）→ terminating error。
- `Today` 未指定時は `[datetime]::Today`。1 桁月日（`M-d`）やスラッシュ区切りは不可。

### Get-NextFriday （Private・純関数）
```
Get-NextFriday [-From] <datetime> -> [datetime]
```
- `From` の「次の金曜日」を返す。金曜(5)までの日数 `delta = (5 - [int]From.DayOfWeek + 7) % 7` を用い、
  `delta -eq 0`（`From` が金曜）なら `delta = 7` とする。
- 結果: 月〜木 → 同じ週の金曜、金 → 翌週の金曜、土日 → 次に来る金曜（翌週の金曜）。

### Resolve-Period （Public・純関数）
```
Resolve-Period [-Start] <datetime> [-End] <datetime> [-Week] <int[]> [-Today] <datetime> -> [pscustomobject[]] @{ Start; End }
```
- 未指定の日付は `[datetime]::MinValue`、`Week` 未指定は空配列で表す（`openslot.ps1` から省略時に渡す）。
- `Today` 未指定時は `[datetime]::Today`。戻り値は常に 1 要素以上の配列。
- `Week.Count -gt 0`（週指定モード）: `Today` を含む週の月曜（`daysSinceMonday = ([int]Today.DayOfWeek + 6) % 7`）を基準に、
  `Week` を `Sort-Object -Unique` した各 `w` について
  `End = 基準月曜.AddDays(7 * w + 4)`、`Start = w -eq 0 ? Today の翌日 : 基準月曜.AddDays(7 * w)`。
  週ごとに 1 要素、週の昇順。`Start` / `End` 引数は無視する。
  `w -eq 0` かつ `Today` が金曜以降のとき `Start -gt End` となる（逆転の除外は `openslot.ps1`）。
- `Week` 空: `Start` 未指定 → `Today` の翌日 / `End` 未指定 → `Get-NextFriday -From <確定した Start>`。1 要素の配列。
- 返す `Start` / `End` はいずれも `.Date`（時刻を落とす）。
- 逆転（`End -lt Start`）や `Week` の負数・排他の検証は行わない（`openslot.ps1` の責務）。

### Resolve-DateRange （Private・純関数）
```
Resolve-DateRange [-Start] <datetime> [-End] <datetime> -> [datetime[]]
```
- `Start` / `End`（いずれも日付、時刻は 00:00 とみなす）から、両端を含む平日（月〜金）の
  日付配列を昇順で返す。土日は含めない。
- `End.Date -lt Start.Date` の場合は terminating error（メッセージ: `終了日は開始日以降を指定してください`）。
- 祝日は除外しない（現行仕様）。将来のオプション拡張点としてコメントを残す。

### Get-CalendarEvents （Private・副作用あり・Graph 版のみ）
```
Get-CalendarEvents [-Start] <datetime> [-End] <datetime> -> [pscustomobject[]] @{ Start; End }
```
- 前提: 呼び出し前に `Connect-MgGraph -Scopes Calendars.Read` 済み（接続は `openslot.ps1` が管理）。
- `Invoke-MgGraphRequest -Method GET` で `/v1.0/me/calendarView` を呼ぶ。
  - クエリ: `startDateTime={Start:o}`, `endDateTime={End:o}`, `$select=start,end,showAs,isAllDay`, `$top=100`
  - ヘッダ: `Prefer: outlook.timezone="Tokyo Standard Time"` … 応答の `start.dateTime` / `end.dateTime` を JST で受け取る
  - `@odata.nextLink` を辿って全ページ取得
- フィルタ: `showAs` が `busy` / `oof` / `tentative` のもののみ採用（`free` / `workingElsewhere` は除外）。
- `isAllDay` の終日予定も上記 `showAs` 条件に合致すればそのまま区間として扱う（終日 busy/oof なら
  その日は営業時間全体が埋まり、結果的に行が出力されない）。
- 返す区間の時刻は JST の `[datetime]`（Kind=Unspecified、JST 壁時計）。
- フォールバック: `Invoke-MgGraphRequest` が使えない環境では `Get-MgUserCalendarView`
  （`-UserId (Get-MgContext).Account`）に切り替える。実装時に前者で動けば後者は不要。

### ConvertFrom-CalendarSearchEvent （Public・純関数・MCP 版）
```
ConvertFrom-CalendarSearchEvent [-Event] <object[]> -> [pscustomobject[]] @{ Start; End }
```
- Claude の `outlook_calendar_search` が返す予定オブジェクト（`showAs` / `isAllDay` / `isCancelled` /
  `start` = `@{ dateTime; timeZone }` / `end` = 同）の配列を受け取り、「埋まっている」区間だけを
  `@{ Start=[datetime]; End=[datetime] }` 配列で返す。`Get-CalendarEvents` の戻り値と同一形式。
- 採用条件（すべて満たすもののみ）:
  - `showAs` が `busy` / `oof` / `tentative`（大文字小文字無視）。`free` / `workingElsewhere` は除外。
  - `isCancelled` が `true` でない。
- 時刻: `start.dateTime` / `end.dateTime` を `[datetime]::Parse(..., InvariantCulture, AssumeLocal→なし)` 相当で
  パースし、Kind=Unspecified の壁時計として扱う（JST 前提。`timeZone` の値は検証しない）。
  `timeZone` が `Tokyo Standard Time` 以外だった場合はそのままパースしつつ warning を出す（将来の変換拡張点）。
- `isAllDay` の終日予定も上記条件に合致すれば区間として扱う（`start`〜`end` が 00:00〜翌 00:00 で返るため
  その日の営業時間全体を覆い、結果的に行が出力されない）。
- `New-Interval` を通し、`End <= Start`（$null）や不正日時の要素は除外する。
- 入力が空・`$null` なら空配列を返す。並び順は保持（呼び出し側でマージ）。

### Merge-Interval （Private・純関数）
```
Merge-Interval [-Interval] <pscustomobject[]> -> [pscustomobject[]]
```
- 入力区間を `Start` 昇順にソートし、重なり・隣接（`次.Start -le 現.End`）を統合して返す。
- 空入力なら空配列を返す。

### Get-FreeInterval （Private・純関数）
```
Get-FreeInterval [-Date] <datetime> [-Busy] <pscustomobject[]>
                 [-SlotMinutes] <int> = 30 [-MinimumMinutes] <int> = 30 -> [pscustomobject[]]
```
- 対象日の営業時間ウィンドウ `Date+08:00 .. Date+12:00` と `Date+13:00 .. Date+19:00` を生成
  （昼休み 12:00–13:00・早朝・夜は最初から除外）。
- `Busy`（マージ済み前提でなくてよい。内部で `Merge-Interval` を呼ぶ）を各ウィンドウから減算し、
  空き部分区間を得る。
- 各空き部分区間を `SlotMinutes` グリッドに内側丸め（Start は切り上げ、End は切り捨て）。
- 丸め後に `End - Start` が `MinimumMinutes` 未満のものを除外。
- 残った区間を昇順で返す。

### Format-FreeDay （Private・純関数）
```
Format-FreeDay [-Date] <datetime> [-Free] <pscustomobject[]> -> [string]
```
- 書式: `M月d日(w): h:mm-h:mm, h:mm-h:mm, ...`
  - 日付・時は 0 埋めしない（`$d.Month`, `$d.Day`, `$t.Hour`）。分は 2 桁 0 埋め（`{0:00}`）。
  - 曜日はハードコード配列 `('日','月','火','水','木','金','土')[[int]$Date.DayOfWeek]` で日本語化
    （カルチャ非依存。これを主実装とする）。
- `Free` が空の場合は $null を返す（呼び出し側で当該日をスキップ）。

### Get-FreeTime （Public・オーケストレーション）
```
Get-FreeTime [-Start] <datetime> [-End] <datetime> [-MinimumMinutes] <int> = 30
             [-Events <pscustomobject[]>] -> [string[]]
```
1. `$days = Resolve-DateRange -Start $Start -End $End`
2. イベント取得:
   - `-Events` が渡された場合（`$PSBoundParameters.ContainsKey('Events')`、空配列も可）はそれを使う。
   - 渡されなかった場合は従来どおり `Get-CalendarEvents -Start $days[0] -End $days[-1].AddDays(1)`（Graph 版）。
   - `-Events` の要素は `@{ Start; End }`（JST 壁時計）。busy/oof/tentative フィルタや終日予定の
     切り出しは注入側の責務（MCP 版では `ConvertFrom-CalendarSearchEvent` が担う）。
3. 各 `$day` について、その日に重なる events を抽出 → `Get-FreeInterval -MinimumMinutes $MinimumMinutes` → `Format-FreeDay`
4. $null（空きなし）の行を除外して文字列配列で返す
- `$days` が空（期間内に平日なし）の場合は空配列を返す（エラーにしない）。
- 既存 `openslot.ps1` は `-Events` を渡さないため、挙動は不変。

## 3. CLI 仕様（openslot.ps1）

### パラメータ
| パラメータ | 型 | 必須 | 説明 |
|---|---|---|---|
| `-Start` | string | 任意 | 開始日。`yyyy-MM-dd` または `MM-dd`（年は実行年）。省略時は実行日の翌日 |
| `-End` | string | 任意 | 終了日。`yyyy-MM-dd` または `MM-dd`（年は実行年）。両端を含む。省略時は開始日の次の金曜日（`Resolve-Period`） |
| `-Week` | int[] | 任意 | 対象期間を「Week 週後の月曜 〜 Week 週後の金曜」とする（基準は実行週の月曜）。`0`（今週）のみ開始日は「翌日」。各要素 0 以上。`0,2,4` の複数指定可（週昇順・重複除去）。`-Start` / `-End` と排他 |
| `-Duration` | int | 任意 | 出力する空き枠の最小分数。既定 30。30 の倍数（30 以上）のみ。`Get-FreeTime -MinimumMinutes` に渡す |
| `-NoClipboard` | switch | 任意 | クリップボードへのコピーを抑止（CI・パイプ利用向け） |
| `-?` | switch | 任意 | comment-based help を表示（`Get-Help .\openslot.ps1 -Full` も可） |

- 指定の有無は `$PSBoundParameters.ContainsKey('Start' / 'End' / 'Week')` で判定する。
- 解決手順（`openslot.ps1` は最初にモジュールを読み込んでから検証する）:
  1. `$today = [datetime]::Today` を確定。
  2. `-Duration`（既定 30）: `30 未満` または `30 で割り切れない` なら `exit 1`。
  3. `-Week` 指定時: `-Start` / `-End` も指定されていれば `exit 1`（排他）、要素が空、または負数を含むなら `exit 1`。
  4. 指定された日付パラメータのみ `ConvertTo-InputDate -Text $v -Today $today` でパースする
     （`yyyy-MM-dd` / `MM-dd` を受理。失敗時 → `exit 1`）。
  5. `-Start` `-End` 両方指定かつ `End.Date -lt Start.Date` → `exit 1`（`終了日は開始日以降を指定してください`）。
  6. `$periods = @(Resolve-Period @{ Start; End; Today }(+ Week 指定時は Week))` で 1 つ以上の期間を確定する。
  7. `$periods` から `Start -le End` の期間だけ残す（既定値のズレ、`-Week 0` を金曜以降に実行、等を除外）。
     残りが 0 件 → 対象日なし。標準出力は空、`-NoClipboard` でなければ空文字列をコピーして `exit 0`。
  8. 各期間について `Get-FreeTime -Start $p.Start -End $p.End -MinimumMinutes $Duration` を実行し、結果行を順に連結する。
     クリップボードには連結後の全文をコピーする。
- 既定動作（`-NoClipboard` なし）では標準出力と同一の全文を `Set-Clipboard` でコピーする。
  空き結果が 0 行のときは空文字列をコピーする。

### 終了コード
| コード | 意味 |
|---|---|
| 0 | 正常終了（空き 0 行、または既定値適用で対象日なしの場合も 0） |
| 1 | 引数不正（日付書式不正、`-Start`/`-End` 両方指定で `End < Start`、`-Week` の空・負数・`-Start`/`-End` との同時指定、`-Duration` が 30 の倍数でない） |
| 2 | 認証失敗（`Connect-MgGraph` 例外、スコープ同意拒否） |
| 3 | カレンダー取得失敗（Graph API 例外・ネットワーク断） |

### 標準出力／標準エラー
- 結果行は標準出力（`Write-Output`）。
- エラーメッセージは標準エラー（`Write-Error` もしくは `[Console]::Error.WriteLine`）。

## 4. エラーハンドリング方針

- モジュール内関数は異常時に terminating error を投げる（`throw` / `-ErrorAction Stop`）。
- `openslot.ps1` が全体を `try/catch` で囲み、例外の発生源に応じて終了コードへ変換する。
  - パラメータ検証は `openslot.ps1` 内で行い、失敗時は `exit 1`。
  - `Connect-MgGraph` を専用の `try/catch` で囲み、失敗時 `exit 2`。
  - `Get-FreeTime`（内部で `Get-CalendarEvents`）の例外は `exit 3`。
- Graph の 429（スロットリング）は `Invoke-MgGraphRequest` 側の自動リトライに委ねる。
  それでも失敗した場合は `exit 3`。

## 5. 標準機能だけで満たせない項目とフォールバック

| 項目 | 主手段 | フォールバック |
|---|---|---|
| Graph 接続（Graph 版） | `Microsoft.Graph.Authentication` の `Connect-MgGraph` | なし（未インストール時は `exit 2` とし README で導線を示す） |
| 予定取得（Graph 版） | `Invoke-MgGraphRequest`（生 REST） | `Get-MgUserCalendarView` cmdlet |
| 予定取得（MCP 版） | Claude の `outlook_calendar_search`（Claude がツール実行） | なし（コネクタ未接続時はスキルが実行不可。README で導線を示す） |
| MCP 版のフェーズ間受け渡し | セッションのスクラッチディレクトリ上の JSON ファイル | （不要） |
| 応答の JST 変換 | `Prefer: outlook.timezone` ヘッダ | UTC で受け取り `[TimeZoneInfo]::FindSystemTimeZoneById('Tokyo Standard Time')` で変換 |
| 曜日の日本語化 | ハードコード配列（主実装） | （不要） |
| クリップボード | `Set-Clipboard` | `-NoClipboard` 指定で回避可能。CI では常にこれを付ける |

## 6. パッケージング方針

- `src/OpenSlot/OpenSlot.psd1` にモジュールバージョンを一元管理（`ModuleVersion`）。
- `OpenSlot.psm1` は `Public/*.ps1` `Private/*.ps1` を dot-source し、`Public/` の関数
  （`Get-FreeTime` / `Resolve-Period` / `ConvertTo-InputDate`）を Export。`psd1` の `FunctionsToExport` にも同じ一覧を記載する。
- `openslot.ps1` / `openslot-mcp.ps1` はいずれも `Import-Module "$PSScriptRoot/src/OpenSlot/OpenSlot.psd1" -Force` で読み込む。
- `ConvertFrom-CalendarSearchEvent` は **Public**（`openslot-mcp.ps1` から呼ぶため）。
  `OpenSlot.psm1` の Export と `psd1` の `FunctionsToExport` に追加する
  （公開関数は `Get-FreeTime` / `Resolve-Period` / `ConvertTo-InputDate` / `ConvertFrom-CalendarSearchEvent`）。
- 外部パッケージマネージャは使わない。配布はリポジトリごと。`Microsoft.Graph` の導入手順は README に記載。

## 7. テスト方針（Pester v5）

単体テスト対象は純関数のみ。Graph 呼び出し（`Get-CalendarEvents`）は単体テストせず、手動確認（plan.md）で担保する。

| 関数 | 主なケース（境界値・異常系を含む） |
|---|---|
| `ConvertTo-InputDate` | `yyyy-MM-dd`、`MM-dd`＋実行年、`MM-dd` が Today の年に追従、時刻を落とす、うるう日（うるう年可／非うるう年で例外）、`M-d`・スラッシュ・不正文字列で例外 |
| `Get-NextFriday` | 月→同週金、木→翌日金、金→翌週金、土→次の金、日→次の金 |
| `Resolve-Period` | `Start` のみ省略、`End` のみ省略、両方省略、両方指定（時刻を落とす）、金曜実行（開始=土・終了=翌週金）、`Week 0` は翌日開始・`Week>=1` は月曜開始、`Week 0` を金曜実行で Start>End、複数 `Week`（週昇順の期間配列）、重複除去、日曜実行での基準週、`Week` 指定時は Start/End 引数を無視、戻り値は常に配列 |
| `Resolve-DateRange` | 単日（平日／土日）、週またぎ、土日のみの期間→空配列、`End < Start`→例外 |
| `Merge-Interval` | 空入力、重複、隣接、内包、離れた 2 区間、順不同入力 |
| `Get-FreeInterval` | 予定なし→`8:00-12:00, 13:00-19:00`、昼休みで分割、予定が午前を完全に埋める、境界が :15 の予定→30分グリッドに内側丸め、残り 30 分未満→除外、`-MinimumMinutes 60` で 30 分枠を除外、終日 busy→空配列 |
| `Format-FreeDay` | 0 埋めなし（`9月1日`, `8:00`）、分の 2 桁化（`15:30`）、各曜日の日本語表記、複数枠のカンマ連結、`Free` 空→$null |
| `Get-FreeTime`（`Get-CalendarEvents` をスタブ化 / `-Events` 注入） | 土日除外、空きなし日は行なし、平日なし→空配列、該当日の予定のみ差し引き、`-MinimumMinutes` で最小枠を制御、`-Events @()` 明示で予定なし扱い、`-Events` 指定時は `Get-CalendarEvents` を呼ばない |
| `ConvertFrom-CalendarSearchEvent` | `busy`/`oof`/`tentative` を採用・`free`/`workingElsewhere` を除外、`showAs` の大文字小文字揺れ、`isCancelled=true` を除外、`isAllDay` の終日予定を区間化、`start.dateTime` の壁時計パース、`timeZone` が Tokyo 以外で warning、`End<=Start` を除外、空入力・`$null`→空配列、戻り値の形式が `Get-CalendarEvents` と一致 |

- テストは `Invoke-Pester ./tests` で全件実行。CI 相当の実行は `pwsh -c "Invoke-Pester ./tests -CI"`。
- `pwsh` / Pester v5 が無い環境では Windows PowerShell 5.1 同梱の Pester、もしくは
  各関数を dot-source した簡易アサーションスクリプトにフォールバックする（採用時に plan.md と README を実態へ更新）。
- `openslot-mcp.ps1` と `SKILL.md` は単体テストせず、`doc/plan.md` の手動確認で担保する。

## 8. MCP 版の設計（スキル `openslot` + `openslot-mcp.ps1`）

### 8.1 全体フロー

Claude Code スキルは「手順書」であり、MCP ツール呼び出し（`outlook_calendar_search`）は Claude が実行する。
日付ロジックとクリップボードは `openslot-mcp.ps1`（PowerShell）に閉じ込め、`src/OpenSlot` を再利用する。

```
1. ユーザー: 「来週の空きを1時間枠で」など（メール返信の下調べ）
2. Claude: 自然言語 → CLI 引数体系（-Week 1 -Duration 60 等）へ変換
3. Claude: openslot-mcp.ps1 -ResolveOnly <引数>        → 期間解決 JSON を得る（フェーズ1）
4. Claude: 期間解決 JSON の fetch 範囲で outlook_calendar_search を呼ぶ
          （nextOffset を辿り全件。取得予定を JSON 配列でファイル保存）
5. Claude: openslot-mcp.ps1 -PlanPath <p> -EventsPath <e>  → 空き行を算出・出力・クリップボード（フェーズ2）
6. Claude: 出力をそのままユーザーに提示（必要ならこの後ユーザー指示でメール本文へ）
```

フェーズを 2 つに分ける理由: 期間解決（`Resolve-Period`。`-Week` や「次の金曜」を含む）を PowerShell 側の
テスト済みロジックで行い、その結果に基づいて Claude が予定取得範囲を決めるため。

### 8.2 `openslot-mcp.ps1`

共通パラメータ（Graph 版 `openslot.ps1` と同一の意味・既定値・検証）:
`-Start` / `-End` / `-Week` / `-Duration` / `-NoClipboard`。

モード:

| モード | 起動 | 動作 |
|---|---|---|
| フェーズ1 | `-ResolveOnly` | 引数を検証（不正なら `exit 1`）→ `Resolve-Period` で期間確定 → `Start -le End` の期間のみ残す → 下記 JSON を標準出力へ。`exit 0` |
| フェーズ2 | `-PlanPath <p> -EventsPath <e>` | プラン JSON と予定 JSON を読む → `ConvertFrom-CalendarSearchEvent` で busy 区間化 → 期間ごとに `Get-FreeTime -Events` → 行を連結して標準出力＋`Set-Clipboard` → `exit 0` |

フェーズ1 の出力 JSON:
```json
{
  "periods":  [ { "start": "2026-09-07", "end": "2026-09-11" } ],
  "fetch":    { "start": "2026-09-07", "end": "2026-09-12" },
  "duration": 60,
  "noClipboard": false
}
```
- `periods` … 各期間の開始・終了日（両端含む、`yyyy-MM-dd`）。0 件なら空配列。
- `fetch` … `outlook_calendar_search` に渡す範囲。`start` = 全 `periods` の最小開始日、
  `end` = 全 `periods` の最大終了日 + 1 日（`beforeDateTime` 用の排他端）。`periods` 空なら `null`。
- `duration` / `noClipboard` … フェーズ2 へ引き継ぐ確定値。

フェーズ2:
- `periods` が空 → 標準出力は空、`noClipboard` でなければ空文字列を `Set-Clipboard`、`exit 0`。
- 予定 JSON は `outlook_calendar_search` の要素をそのまま格納した配列（`showAs` / `isAllDay` /
  `isCancelled` / `start` / `end` を含む）。パース不能なら `exit 3`。
- 期間ごとに `Get-FreeTime -Start $p.start -End $p.end -MinimumMinutes $duration -Events $events` を実行し、
  結果行を順に連結。クリップボードには連結後の全文（0 行なら空文字列）。

終了コード: `0`=正常 / `1`=引数不正 / `3`=処理失敗（プラン・予定 JSON 不正など）。
Graph 版の `2`（認証失敗）は無い（認証は Claude コネクタ側）。

### 8.3 `.claude/skills/openslot/SKILL.md`

- フロントマター: `name: openslot` / `description`（メール返信時に空き時間候補を作る、CLI 引数体系で期間指定、
  MS365 コネクタ使用、の要旨）。
- 本文に記す手順:
  1. 期間の決め方（自然言語 → `-Start/-End/-Week/-Duration`）。判断がつかなければユーザーに確認。
  2. フェーズ1 実行例（`openslot-mcp.ps1 -ResolveOnly ...`）。`periods` 空なら「対象日なし」で終了。
  3. `outlook_calendar_search` の呼び方: `query:"*"`、`afterDateTime`/`beforeDateTime` に `fetch` 範囲、
     `nextOffset` が返る限りページング。取得した要素をそのまま JSON 配列でスクラッチへ保存。
  4. フェーズ2 実行例。標準出力の行をそのままユーザーへ提示（クリップボードにも入っている旨を添える）。
  5. スコープ境界: メール本文の作成・送信は行わない。ユーザーが明示した場合のみ別途対応。
- スクラッチファイルの置き場所はセッションのスクラッチディレクトリ。完了後は残さない。
