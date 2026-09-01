function Set-OutlookSignature {
    <#
    .SYNOPSIS
        クラシック版 Outlook デスクトップのメール署名を更新する。

    .DESCRIPTION
        指定したプレーンテキストから署名ファイル（.htm / .txt）を生成し、既定プロファイルの
        メールアカウントに反映する。2 つのモードがある。

        - UpdateInPlace  : -Name 省略時。各アカウントが現在使っている署名（レジストリの
                           New Signature / Reply-Forward Signature が指す署名）の本文だけを差し替える。
                           レジストリは変更しない。ローミング署名が有効な環境で推奨。
        - CreateAndAssign: -Name 指定時。その名前で署名を作成／更新し、対象アカウントすべての
                           新規メール・返信転送の既定署名に設定する。

        変更前にプリフライト検証を行い、いずれか失敗した場合は一切書き込まずに終了する。
        Outlook が起動中の場合は中断する（-Force で続行可能）。反映には Outlook の再起動が必要。

    .PARAMETER Name
        署名名。省略すると現在の既定署名を in-place 更新する。

    .PARAMETER Text
        署名本文（プレーンテキスト、複数行可）。

    .PARAMETER Force
        Outlook 起動中でも続行する。

    .PARAMETER SignaturesPath
        Signatures フォルダのパス（テスト／上級者向け上書き）。既定は %APPDATA%\Microsoft\Signatures。

    .PARAMETER RegistryRoot
        Office のレジストリルート（テスト／上級者向け上書き）。既定は
        HKCU:\Software\Microsoft\Office\16.0。

    .EXAMPLE
        Set-OutlookSignature -Text "山田 太郎`n一般社団法人 ○○"

        現在使っている署名の本文を差し替える（in-place）。

    .EXAMPLE
        Set-OutlookSignature -Name "標準" -Text "山田 太郎" -WhatIf

        署名 "標準" を作成し既定にする操作の予定を表示する（変更しない）。

    .NOTES
        - 反映には Outlook の再起動が必要。
        - ローミング署名が有効な場合、-Name 指定の新規署名は Outlook のピッカーに出ないことがある。
          その場合は -Name を省略した in-place 更新を使う。
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Position = 0)]
        [string] $Name,

        [Parameter(Mandatory, Position = 1, ValueFromPipeline = $true)]
        [AllowEmptyString()]
        [string] $Text,

        [switch] $Force,

        [string] $SignaturesPath,

        [string] $RegistryRoot = 'HKCU:\Software\Microsoft\Office\16.0'
    )

    process {
        $warnings = [System.Collections.Generic.List[string]]::new()
        try {
            $isCreateMode = -not [string]::IsNullOrEmpty($Name)

            # 1. 入力検証（-Name 指定時のみ）
            if ($isCreateMode) {
                $invalid = [System.IO.Path]::GetInvalidFileNameChars()
                if ($Name.IndexOfAny($invalid) -ge 0) {
                    throw "署名名にファイル名として使えない文字が含まれています: '$Name' [InvalidSignatureName]"
                }
                if ($Name.Length -gt 255) {
                    throw "署名名が長すぎます（255 文字以内）: '$Name' [InvalidSignatureName]"
                }
                if ($Name.Length -gt 32) {
                    $warnings.Add("署名名が 32 文字を超えています。Outlook 側で切り詰められる可能性があります。")
                }
            }

            # 2. 環境解決
            $env = Resolve-OutlookEnvironment -RegistryRoot $RegistryRoot -SignaturesPath $SignaturesPath

            # 3. アカウント列挙
            $accounts = @(Get-OutlookMailAccount -AccountsKeyPath $env.AccountsKeyPath)
            if ($accounts.Count -eq 0) {
                throw "対象となるメールアカウントが見つかりません。 [NoMailAccount]"
            }

            # 4. 更新対象署名名の決定
            if ($isCreateMode) {
                $mode = 'CreateAndAssign'
                $targetNames = @($Name)
            }
            else {
                $mode = 'UpdateInPlace'
                $targetNames = @(
                    $accounts | ForEach-Object { $_.PreviousNew; $_.PreviousReplyForward } |
                        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                        Select-Object -Unique
                )
                if ($targetNames.Count -eq 0) {
                    throw "現在の既定署名が設定されていないため、更新対象を特定できません。-Name を指定してください。 [NoCurrentSignature]"
                }
            }

            # 5. Signatures フォルダ検証（作成可能か）。プリフライトなので -WhatIf の影響を受けない .NET API を使う。
            try {
                [void][System.IO.Directory]::CreateDirectory($env.SignaturesPath)
                $probe = Join-Path $env.SignaturesPath ('.probe-' + [guid]::NewGuid().ToString('N'))
                [System.IO.File]::WriteAllText($probe, 'x')
                [System.IO.File]::Delete($probe)
            }
            catch {
                throw "Signatures フォルダに書き込めません: $($env.SignaturesPath) [SignaturesFolderNotWritable]"
            }

            # 6. Outlook 起動チェック
            if (Test-OutlookRunning) {
                if (-not $Force) {
                    throw "Outlook が起動中です。終了してから再実行するか -Force を指定してください。 [OutlookRunning]"
                }
                $warnings.Add("Outlook が起動中です。-Force で続行しますが、終了時に設定が書き戻される可能性があります。")
            }

            # 7. ローミング署名チェック
            $roamingEnabled = Test-RoamingSignatureEnabled -RegistryRoot $RegistryRoot
            if ($roamingEnabled) {
                $warnings.Add("ローミング署名が有効の可能性があります。変更がクラウド同期で上書きされることがあります。")
                if ($mode -eq 'CreateAndAssign') {
                    $warnings.Add("ローミング署名環境では、新規作成した署名が Outlook のピッカーに現れないことがあります。-Name を省略した in-place 更新の利用を検討してください。")
                }
            }

            # --- ここまでプリフライト（読み取りのみ）---

            $result = [pscustomobject]@{
                Mode                     = $mode
                SignatureNames           = $targetNames
                SignaturesPath           = $env.SignaturesPath
                Files                    = @()
                BackupPaths              = @()
                Accounts                 = $accounts
                RegistryChanged          = $false
                RoamingSignaturesEnabled = $roamingEnabled
                Warnings                 = @($warnings)
                RestartRequired          = $true
            }

            $actionText = if ($mode -eq 'CreateAndAssign') {
                "署名 '$Name' を作成し、$($accounts.Count) 個のアカウントの既定署名に設定"
            }
            else {
                "署名 [$($targetNames -join ', ')] の本文を更新"
            }

            if (-not $PSCmdlet.ShouldProcess($env.SignaturesPath, $actionText)) {
                foreach ($w in $warnings) { Write-Warning $w }
                return $result
            }

            # --- 変更フェーズ ---
            $writtenFiles = [System.Collections.Generic.List[string]]::new()
            $backupPaths  = [System.Collections.Generic.List[string]]::new()
            $registryUndo = [System.Collections.Generic.List[pscustomobject]]::new()

            try {
                foreach ($sig in $targetNames) {
                    $backup = Backup-SignatureFileSet -SignaturesPath $env.SignaturesPath -Name $sig
                    if ($backup) { $backupPaths.Add($backup) }
                    $files = Write-SignatureFileSet -SignaturesPath $env.SignaturesPath -Name $sig -Text $Text
                    foreach ($f in $files) { $writtenFiles.Add($f) }
                }

                if ($mode -eq 'CreateAndAssign') {
                    foreach ($acct in $accounts) {
                        $registryUndo.Add([pscustomobject]@{
                            Path         = $acct.RegistryPath
                            PreviousNew  = $acct.PreviousNew
                            PreviousRfwd = $acct.PreviousReplyForward
                        })
                        Set-AccountSignatureRegistry -AccountKeyPath $acct.RegistryPath -SignatureName $Name
                    }
                    $result.RegistryChanged = $true
                }
            }
            catch {
                $rollbackErr = $_
                # ベストエフォートのロールバック
                foreach ($f in $writtenFiles) {
                    if (Test-Path -LiteralPath $f) { Remove-Item -LiteralPath $f -Force -ErrorAction SilentlyContinue }
                }
                foreach ($b in $backupPaths) {
                    Get-ChildItem -LiteralPath $b -ErrorAction SilentlyContinue | ForEach-Object {
                        Move-Item -LiteralPath $_.FullName -Destination $env.SignaturesPath -Force -ErrorAction SilentlyContinue
                    }
                }
                foreach ($u in $registryUndo) {
                    if ([string]::IsNullOrEmpty($u.PreviousNew)) {
                        Remove-ItemProperty -LiteralPath $u.Path -Name 'New Signature' -Force -ErrorAction SilentlyContinue
                    }
                    else {
                        Set-ItemProperty -LiteralPath $u.Path -Name 'New Signature' -Value $u.PreviousNew -Type String -ErrorAction SilentlyContinue
                    }
                    if ([string]::IsNullOrEmpty($u.PreviousRfwd)) {
                        Remove-ItemProperty -LiteralPath $u.Path -Name 'Reply-Forward Signature' -Force -ErrorAction SilentlyContinue
                    }
                    else {
                        Set-ItemProperty -LiteralPath $u.Path -Name 'Reply-Forward Signature' -Value $u.PreviousRfwd -Type String -ErrorAction SilentlyContinue
                    }
                }
                throw "署名の書き込み中にエラーが発生しました（ロールバックを試行済み）: $($rollbackErr.Exception.Message) [WriteFailed]"
            }

            $result.Files       = @($writtenFiles)
            $result.BackupPaths  = @($backupPaths)

            foreach ($w in $warnings) { Write-Warning $w }
            if ($result.RestartRequired) {
                Write-Verbose "反映には Outlook の再起動が必要です。"
            }
            return $result
        }
        catch {
            $msg = $_.Exception.Message
            $errorId = 'SignatureUpdateError'
            if ($msg -match '\[(\w+)\]\s*$') {
                $errorId = $Matches[1]
                $msg = ($msg -replace '\s*\[\w+\]\s*$', '')
            }
            $exception = [System.InvalidOperationException]::new($msg)
            $record = [System.Management.Automation.ErrorRecord]::new(
                $exception, $errorId, [System.Management.Automation.ErrorCategory]::InvalidOperation, $Name)
            $PSCmdlet.ThrowTerminatingError($record)
        }
    }
}
