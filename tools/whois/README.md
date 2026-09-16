# whois

Outlook カレンダーの予定の参加者が「誰なのか」を、手元の人物辞書 `whois.md` と照合して明らかにするツール。
PowerShell の実体は持たず、Claude Code スキルのみで完結する。

スキルの実体はリポジトリ直下の [`.claude/skills/meeting-attendees/`](../../.claude/skills/meeting-attendees/) にある
（Claude Code がそこしか探索しないため）。このディレクトリには付随ファイルと本ドキュメントのみを置く。

## スキル

| スキル | 実体 | 概要 |
|---|---|---|
| `meeting-attendees` | `.claude/skills/meeting-attendees/SKILL.md` | 時間帯から予定を特定 → 参加者取得 → `whois.md` で人物照合 |

## 付随ファイル

- `whois.md` … 人物辞書。メールアドレスと氏名・役割の対応表。
  **個人情報を含むため git 管理外**（`.gitignore` に `tools/whois/whois.md` を登録済み）。
  このリポジトリを新しい環境に配置する場合は、各自でこのファイルを作成する必要がある。

### whois.md の書式

`meeting-attendees` スキルは概ね次の形式を期待する。

```markdown
# 人物名
account@example.com
役割・所属などの説明

# 検索結果から除外する
- メールのドメインが example.org である。
- 除外したい個人のメールアドレス <address@example.com>
```

- `# 見出し` を人物単位の区切りとして扱う。
- 「検索結果から除外する」見出し配下は、ユーザー本人のアカウント（ドメイン単位・個別アドレス）を列挙する。
  ここに該当する参加者は照合結果の表に出さない。

## 前提

- Claude Code で Microsoft 365 コネクタが接続済みであること。
- 認証エラーが出たら `/mcp` を実行して「claude.ai Microsoft 365」を再認証する。

## 注意

- 予定の参照のみを行い、予定の作成・変更・返信は行わない。
- `whois.md` はローカル管理のファイルであり、リポジトリ間で共有されない。
