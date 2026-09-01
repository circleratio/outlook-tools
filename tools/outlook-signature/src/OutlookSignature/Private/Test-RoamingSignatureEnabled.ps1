function Test-RoamingSignatureEnabled {
    <#
    .SYNOPSIS
        ローミング署名（クラウド署名）が有効の疑いがあるかを返す（緩い判定）。
    .DESCRIPTION
        Outlook\Setup\DisableRoamingSignaturesTemporaryToggle が 1 のときのみ「無効」と判断し $false。
        それ以外は $true（＝有効の疑いあり、警告対象）。
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)]
        [string] $RegistryRoot
    )

    $setupKey = Join-Path $RegistryRoot 'Outlook\Setup'
    $toggle = Get-RegistryValue -Path $setupKey -Name 'DisableRoamingSignaturesTemporaryToggle'
    if ($toggle -eq 1) {
        return $false
    }
    return $true
}
