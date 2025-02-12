@{
    RootModule = 'DC-Configuration.psm1'
    ModuleVersion = '1.6'
    GUID = '9p8o7n6m-5l4k-3j2i-1h0g-f9e8d7c6b5a4'
    Author = 'jdyer-nuvodia'
    Description = 'Module for Domain Controller configuration in Azure'
    PowerShellVersion = '7.0'
    FunctionsToExport = @('Write-Log', 'Initialize-DCConfiguration')
    PrivateData = @{
        PSData = @{
            Tags = @('Azure', 'DomainController', 'Configuration')
            LastUpdated = '2025-02-12 17:32:27'
            UpdatedBy = 'jdyer-nuvodia'
        }
    }
}