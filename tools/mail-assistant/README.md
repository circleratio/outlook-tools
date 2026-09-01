# mail-assistant

Microsoft 365 の受信メールを Claude Code から扱うためのスラッシュコマンド群。
Claude の **Microsoft 365 コネクタ**（`mcp__claude_ai_Microsoft_365__*`）を使うため、
PowerShell モジュールのインストールは不要。

コマンドの実体はリポジトリ直下の [`.claude/commands/`](../../.claude/commands/) にある
（Claude Code がそこしか探索しないため）。このディレクトリには付随ファイルと本ドキュメントのみを置く。

## コマンド

| コマンド | 実体 | 概要 |
|---|---|---|
| `/inbox-today` | `.claude/commands/inbox-today.md` | 本日受信した受信ボックスのメールを連番付きの表で一覧 |
| `/reply-draft` | `.claude/commands/reply-draft.md` | `hint.md` の雛形をもとに返信ドラフトを作成（送信はしない） |

`/inbox-today` で振った連番を `/reply-draft` の対象指定に使える。

## 付随ファイル

- [`hint.md`](hint.md) … `/reply-draft` が毎回読み込む返信文の雛形・敬称ルール。
  文面の方針を変えたいときはこのファイルを編集する。

## 前提

- Claude Code で Microsoft 365 コネクタが接続済みであること。
- 認証エラーが出たら `/mcp` を実行して「claude.ai Microsoft 365」を再認証する。

## 注意

- `/reply-draft` は Draft 作成までで、`outlook_send_draft` は呼ばない（自動送信しない）。
- MS365 へは Claude Code を実行しているユーザーの権限でアクセスする。
