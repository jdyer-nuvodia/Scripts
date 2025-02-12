@{
    RootModule = 'DC-Configuration.psm1'
    ModuleVersion = '1.7'
    GUID = 'f9e8d7c6-b5a4-4321-9876-543210fedcba'
    Author = 'jdyer-nuvodia'
    Description = 'Module for Domain Controller configuration in Azure'
    PowerShellVersion = '7.0'
    FunctionsToExport = @('Write-Log', 'Initialize-DCConfiguration')
    PrivateData = @{
        PSData = @{
            Tags = @('Azure', 'DomainController', 'Configuration')
            LastUpdated = '2025-02-12 17:49:17'
            UpdatedBy = 'jdyer-nuvodia'
        }
    }
}