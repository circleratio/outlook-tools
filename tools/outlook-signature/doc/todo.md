# TODO / 将来対応

## クラシック版：ローミング署名の抑止オプション

今回はローミング署名（クラウド署名）が有効な場合、**検出して警告するのみ**（レジストリ変更等の抑止はしない）。
将来、必要なら以下を検討する。

- `HKCU\Software\Microsoft\Office\16.0\Outlook\Setup\DisableRoamingSignaturesTemporaryToggle = 1` を
  セットするオプション（`-DisableRoaming` 等）。既定では触らない。
- 組織単位の抑止 `Set-OrganizationConfig -PostponeRoamingSignaturesUntilLater $true`（管理者操作）。
- ローミング署名有効時の「クラウド側の署名」を直接更新する経路（＝新 Outlook / OWA 側と同じ Exchange Online 経由）。

## クラシック版：CreateAndAssign（-Name 指定）の実挙動確認

ローミング署名 ON の環境で `-Name` を指定して新規作成した署名が、Outlook の署名ピッカー／
既定署名として実際に見えるか未検証（phase 1 の実機確認では in-place 更新で用が足りたため）。
出ない場合の対応案:
- README の注記を強める（ローミング環境では in-place のみ推奨）
- `-Name` 指定時にクラウド側へ登録する経路（新 Outlook / OWA と同じ Exchange Online 経由）を追加

## ラッパー：終了コード写像の共有

`scripts/Set-OutlookSignatureCli.ps1` と `scripts/Set-SeasonalOutlookSignature.ps1` が
ErrorId → 終了コードの写像表を各自で持っている（phase 1 は重複を許容）。
3 本目の利用者が現れたら、共有のドット取り込みヘルパー（例 `scripts/Common/Get-CliExitCode.ps1`）へ抽出する。

## クラシック版：他 Office バージョン対応

今回はレジストリキーを `16.0` 固定とみなす。将来、複数バージョン検出が必要になったら対応。

## 新しい Outlook for Windows / Outlook on the web の署名更新（CLI）

### 背景

署名の更新は Outlook の系統によって手段が完全に分かれる。

- クラシック版 Outlook デスクトップ … ローカルファイル ＋ レジストリ（本リポジトリで今回実装）
- 新しい Outlook for Windows / Outlook on the web … Exchange Online のメールボックス側に保存

Microsoft Graph API には署名を読み書きする API が無い（EWS も 2026-10-01 に既定でブロック）。
そのため新 Outlook / OWA 側は **Exchange Online PowerShell** が唯一の現実的な CLI 経路。

### やること（将来）

- `Set-MailboxMessageConfiguration` を使った署名更新機能を追加する。
  ```powershell
  Connect-ExchangeOnline -UserPrincipalName yamada@example.com
  Set-MailboxMessageConfiguration -Identity yamada@example.com `
    -SignatureHTML "<div>...</div>" `
    -AutoAddSignature $true -AutoAddSignatureOnReply $true
  ```
- クラシック版と新 Outlook を 1 つのモジュール／コマンドから切り替えて更新できるようにする
  （例: `-Target Classic|Cloud|Both`）。
- プレーンテキスト入力を HTML へ変換して `-SignatureHTML` に渡す処理を共通化する。

### 調査・確認が必要な事項

- **ローミング署名（roaming signatures）との競合**: ローミング署名が有効だと
  `Set-MailboxMessageConfiguration -SignatureHTML` が正しく反映されない既知の不具合がある。
  回避には管理者が組織単位で `Set-OrganizationConfig -PostponeRoamingSignaturesUntilLater $true`
  を実行する必要がある。自組織の Microsoft 365 テナントでの現状（ローミング署名の有効/無効）を要確認。
- 自分のメールボックスに対する `Set-MailboxMessageConfiguration` の実行に必要な権限
  （既定 RBAC ロール `MyBaseOptions` で可能なはずだが要検証）。
- `ExchangeOnlineManagement` モジュールの導入と、非対話（アプリ登録 / 証明書認証）での接続可否。
- 反映先の範囲（新 Outlook デスクトップ、OWA。クラシック版・モバイルには反映されない）。

### 参考

- Set-MailboxMessageConfiguration: https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/set-mailboxmessageconfiguration
- Roaming Signatures と本 cmdlet の不具合: https://learn.microsoft.com/en-us/answers/questions/1372754/outlook-roaming-signatures-feature-id-60371-set-ma
- Graph API に署名 API が無い件: https://learn.microsoft.com/en-us/answers/questions/1093518/user-email-signature-management-via-graph-api
- 横断ツールの参考実装 Set-OutlookSignatures: https://github.com/Set-OutlookSignatures/Set-OutlookSignatures
