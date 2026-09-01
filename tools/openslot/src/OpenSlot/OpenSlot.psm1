Set-StrictMode -Version 2.0

$private = @(Get-ChildItem -Path (Join-Path $PSScriptRoot 'Private/*.ps1') -ErrorAction SilentlyContinue)
$public  = @(Get-ChildItem -Path (Join-Path $PSScriptRoot 'Public/*.ps1')  -ErrorAction SilentlyContinue)

foreach ($file in @($private + $public)) {
    . $file.FullName
}

Export-ModuleMember -Function @($public | ForEach-Object { $_.BaseName })
