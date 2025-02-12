@{
    RootModule = 'DC-Deployment.psm1'
    ModuleVersion = '1.8'
    GUID = '1a2b3c4d-5e6f-4321-9876-543210abcdef'
    Author = 'jdyer-nuvodia'
    Description = 'Module for Domain Controller deployment in Azure'
    PowerShellVersion = '7.0'
    FunctionsToExport = @('New-DCEnvironment')
    PrivateData = @{
        PSData = @{
            Tags = @('Azure', 'DomainController', 'Deployment')
            LastUpdated = '2025-02-12 18:24:19'
            UpdatedBy = 'jdyer-nuvodia'
        }
    }
}