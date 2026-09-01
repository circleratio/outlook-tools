function Test-OutlookRunning {
    <#
    .SYNOPSIS
        OUTLOOK.EXE プロセスが動作中かを返す。
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    [bool](Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue)
}
