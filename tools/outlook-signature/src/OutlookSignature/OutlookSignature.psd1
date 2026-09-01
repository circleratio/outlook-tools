@{
    RootModule        = 'OutlookSignature.psm1'
    ModuleVersion     = '0.1.0'
    GUID              = 'b8d1e4a2-6c3f-4b7a-9e21-2f5c7a0d9e11'
    Author            = 'circleratio'
    CompanyName       = 'Unknown'
    Copyright         = '(c) 2026 circleratio. Licensed under GPL-3.0-or-later.'
    Description       = 'クラシック版 Outlook デスクトップのメール署名を CLI から更新する。'
    PowerShellVersion = '5.1'
    FunctionsToExport = @('Set-OutlookSignature', 'Get-OutlookSignature')
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
    PrivateData       = @{
        PSData = @{
            Tags = @('Outlook', 'Signature', 'Email', 'Windows')
            LicenseUri = 'https://www.gnu.org/licenses/gpl-3.0.html'
        }
    }
}
