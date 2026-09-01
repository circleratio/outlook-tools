function Backup-SignatureFileSet {
    <#
    .SYNOPSIS
        既存の署名ファイル群を <SignaturesPath>\.backup\<Name>-<timestamp>\ へ退避する。
    .OUTPUTS
        バックアップ先フォルダのパス。退避対象が無ければ $null。
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [string] $SignaturesPath,

        [Parameter(Mandatory)]
        [string] $Name
    )

    if (-not (Test-Path -LiteralPath $SignaturesPath)) {
        return $null
    }

    $candidates = @(
        (Join-Path $SignaturesPath ($Name + '.htm'))
        (Join-Path $SignaturesPath ($Name + '.txt'))
        (Join-Path $SignaturesPath ($Name + '.rtf'))
        (Join-Path $SignaturesPath ($Name + '.files'))
        (Join-Path $SignaturesPath ($Name + '_files'))
    )
    $existing = @($candidates | Where-Object { Test-Path -LiteralPath $_ })
    if ($existing.Count -eq 0) {
        return $null
    }

    $stamp    = Get-Date -Format 'yyyyMMdd-HHmmss'
    $backupRoot = Join-Path $SignaturesPath '.backup'
    $backupDir  = Join-Path $backupRoot ("{0}-{1}" -f $Name, $stamp)
    $suffix = 1
    while (Test-Path -LiteralPath $backupDir) {
        $backupDir = Join-Path $backupRoot ("{0}-{1}-{2}" -f $Name, $stamp, $suffix)
        $suffix++
    }
    New-Item -ItemType Directory -Path $backupDir -Force | Out-Null

    foreach ($item in $existing) {
        Move-Item -LiteralPath $item -Destination $backupDir -Force
    }

    (Resolve-Path -LiteralPath $backupDir).Path
}
