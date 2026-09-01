function Get-CalendarEvents {
    <#
    .SYNOPSIS
    Microsoft Graph の /me/calendarView から指定期間の予定を取得し、
    「埋まっている」区間（showAs = busy / oof / tentative）だけを [Start,End] 区間として返す。

    前提: 呼び出し前に Connect-MgGraph -Scopes Calendars.Read 済みであること。
    応答時刻は Prefer ヘッダで JST（Tokyo Standard Time）として受け取る。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][datetime]$Start,
        [Parameter(Mandatory)][datetime]$End
    )

    $busyShowAs = @('busy', 'oof', 'tentative')

    $uri = '/v1.0/me/calendarView?startDateTime={0}&endDateTime={1}&$select=start,end,showAs,isAllDay&$top=100' -f `
        $Start.ToString('yyyy-MM-ddTHH:mm:ss'), $End.ToString('yyyy-MM-ddTHH:mm:ss')
    $headers = @{ Prefer = 'outlook.timezone="Tokyo Standard Time"' }

    $result = New-Object System.Collections.Generic.List[pscustomobject]

    while ($uri) {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers $headers

        foreach ($ev in @($resp['value'])) {
            $showAs = [string]$ev['showAs']
            if ($busyShowAs -notcontains $showAs) { continue }

            $s = [datetime]::Parse([string]$ev['start']['dateTime'])
            $e = [datetime]::Parse([string]$ev['end']['dateTime'])
            $iv = New-Interval -Start $s -End $e
            if ($iv) { $result.Add($iv) }
        }

        $uri = $resp['@odata.nextLink']
    }

    $result.ToArray()
}
