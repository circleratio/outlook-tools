function New-Interval {
    <#
    .SYNOPSIS
    [datetime]Start / [datetime]End を持つ区間オブジェクトを生成する。
    End が Start 以下の場合は $null を返す。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][datetime]$Start,
        [Parameter(Mandatory)][datetime]$End
    )

    if ($End -le $Start) { return $null }

    [pscustomobject]@{
        Start = $Start
        End   = $End
    }
}
