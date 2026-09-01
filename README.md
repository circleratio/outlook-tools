# outlook-tools

Outlook / Microsoft 365 まわりの小さな社内ツールをまとめたリポジトリ。
各ツールは独立していて、共通の実行基盤は持たない（PowerShell モジュールと Claude Code アセットの寄せ集め）。

## ツール一覧

| ツール | 場所 | 種類 | 概要 |
|---|---|---|---|
| **mail-assistant** | [`tools/mail-assistant/`](tools/mail-assistant/) | Claude Code スラッシュコマンド | 受信箱一覧・返信ドラフト作成。MS365 コネクタ経由 |
| **openslot** | [`tools/openslot/`](tools/openslot/) | PowerShell CLI ＋ Claude スキル | MS365 予定表から指定期間の空き時間を抽出（日程調整メール用） |
| **outlook-signature** | [`tools/outlook-signature/`](tools/outlook-signature/) | PowerShell モジュール ＋ CLI ラッパー | クラシック版 Outlook デスクトップのメール署名を CLI から更新 |

各ツールの詳細・使いかた・終了コードはそれぞれの `README.md` を参照。

## Claude Code アセットの置き場所

Claude Code はリポジトリ直下の `.claude/` しか探索しないため、コマンド・スキルの実体はここに集約している。

| パス | 対応ツール |
|---|---|
| `.claude/commands/{inbox-today,reply-draft}.md` | mail-assistant |
| `.claude/skills/openslot/SKILL.md` | openslot（MCP 版） |

`tools/<name>/` 側にはドキュメントと付随ファイル（`hint.md` など）だけを置く。

## 環境要件

| ツール | OS | PowerShell | 追加要件 |
|---|---|---|---|
| mail-assistant | 任意（Claude Code 実行環境） | — | Microsoft 365 コネクタ接続済み |
| openslot（Graph 版） | Windows | 5.1+ | `Microsoft.Graph.Authentication`、`Calendars.Read` 委任同意 |
| openslot（MCP 版） | Windows | 5.1+ | Microsoft 365 コネクタ接続済み |
| outlook-signature | Windows | 5.1 または 7 | クラシック版 Outlook デスクトップ（レジストリ `16.0` 前提） |

## 開発

- リポジトリ規約（エンコーディング、テスト、spec-driven-dev の成果物パス）は [`CLAUDE.md`](CLAUDE.md) を参照。
- 全ツールのテストをまとめて実行: `pwsh -File build/Invoke-Tests.ps1`
  （openslot は Pester 3.x、outlook-signature は Pester 5.x を使うため、ツールごとに別プロセスで実行する）。

## ライセンス

GNU General Public License v3.0 以降（GPL-3.0-or-later）。全文は [`LICENSE`](LICENSE) を参照。

`tools/outlook-signature/` の `*.example.txt` は雛形です。実際の署名・組織情報を書いた
`message.txt` / `template.txt` はコミットしないでください（`.gitignore` 済み）。
