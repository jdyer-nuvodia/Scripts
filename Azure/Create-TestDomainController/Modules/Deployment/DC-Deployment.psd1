@{
    RootModule = 'DC-Deployment.psm1'
    ModuleVersion = '3.5'
    GUID = 'f8b0e1c0-5c1a-4e1b-9b1a-1c2b3d4e5f6a'  # This GUID is already valid
    Author = 'jdyer-nuvodia'
    Description = 'Module for deploying and configuring Domain Controller VMs in Azure'
    PowerShellVersion = '7.0'
    FunctionsToExport = @('New-DCEnvironment', 'Set-VMAutoShutdown')
    PrivateData = @{
        PSData = @{
            Tags = @('Azure', 'DomainController', 'VM')
            LastUpdated = '2025-02-12 17:50:17'
            UpdatedBy = 'jdyer-nuvodia'
        }
    }
}