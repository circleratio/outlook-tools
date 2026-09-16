# リポジトリ規約

Outlook / Microsoft 365 まわりの小さなツールを寄せ集めた monorepo。
各ツールは `tools/<name>/` に独立して置き、相互依存はしない。

## ディレクトリ構成

```
.claude/
  commands/            # スラッシュコマンド（mail-assistant）
  skills/openslot/     # openslot MCP 版スキル
tools/
  mail-assistant/      # README + hint.md（コマンド実体は .claude/commands/）
  openslot/            # openslot.ps1 / openslot-mcp.ps1 / src/OpenSlot / doc / tests
  outlook-signature/   # src/OutlookSignature / scripts / doc / tests / message.example.txt / template.example.txt
build/
  Invoke-Tests.ps1     # 全ツールのテストを実行（ツールごとに別プロセス）
```

Claude Code のコマンド・スキルはリポジトリ直下の `.claude/` でしか認識されないため、
実体は `.claude/` に置き、`tools/<name>/` にはドキュメントと付随ファイルのみを置く。

## エンコーディング

PowerShell ソース（`.ps1` / `.psd1` / `.psm1`）は **UTF-8 with BOM** で保存する。
Windows PowerShell 5.1 が日本語を含むファイルを正しく解析するために必要。
Markdown・YAML は UTF-8（BOM なし）。`.editorconfig` に定義済み。

## テスト

ツールごとに必要な Pester のメジャーバージョンが異なる。混在環境で 1 セッションに
両方を読み込めないため、`build/Invoke-Tests.ps1` がツールごとに子プロセスを分けて実行する。

| ツール | Pester | 実行 |
|---|---|---|
| openslot | 3.x（Windows 同梱） | `Invoke-Pester -Path tools/openslot/tests` |
| outlook-signature | 5.x（要 `Install-Module Pester -MinimumVersion 5.1`。`TestRegistry:` を使うため 5.1 未満不可） | `Invoke-Pester tools/outlook-signature/tests` |

レジストリ／ファイルを触るテストは `TestDrive:` と Pester の `TestRegistry:` で隔離済み。
ただし CLI／ラッパースクリプトを別プロセスとして実行するテスト
（`Set-OutlookSignatureCli.Tests.ps1`、`Set-SeasonalOutlookSignature.Tests.ps1`）は、
`TestRegistry:` ドライブが子プロセスから参照できないため、実 HKCU 配下に一意な一時キーを作り
`AfterAll` で削除する方式を取る。

## 機能追加・変更フロー（spec-driven-dev）

`spec-driven-dev` スキルの成果物は **対象ツールのディレクトリ配下** に置く。

- `tools/<name>/doc/requirement.md` … 要求仕様
- `tools/<name>/doc/spec.md` … 設計
- `tools/<name>/doc/plan.md` … 実装計画
- `tools/<name>/README.md` … 利用者向けドキュメント

新しいツールを足すときは `tools/<new-name>/` を作り、同じ 4 点セットを揃える。
Claude Code アセットを伴う場合は実体を `.claude/` に追加する。
