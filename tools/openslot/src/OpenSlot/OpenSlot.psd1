@{
    RootModule        = 'OpenSlot.psm1'
    ModuleVersion     = '0.1.0'
    GUID              = '53c0986e-3adb-42be-8fdc-e63d8feba82b'
    Author            = 'circleratio'
    Copyright         = '(c) 2026 circleratio. Licensed under GPL-3.0-or-later.'
    Description       = 'MS365 予定表から指定期間の空き時間を抽出する'
    PowerShellVersion = '5.1'
    FunctionsToExport = @('Get-FreeTime', 'Resolve-Period', 'ConvertTo-InputDate', 'ConvertFrom-CalendarSearchEvent')
    CmdletsToExport   = @()
    AliasesToExport   = @()
    VariablesToExport = @()
    PrivateData       = @{
        PSData = @{
            LicenseUri = 'https://www.gnu.org/licenses/gpl-3.0.html'
        }
    }
}
