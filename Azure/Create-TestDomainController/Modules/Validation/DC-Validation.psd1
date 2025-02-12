@{
    RootModule = 'DC-Validation.psm1'
    ModuleVersion = '1.8'
    GUID = 'a1b2c3d4-e5f6-4321-9876-543210abcdef'
    Author = 'jdyer-nuvodia'
    Description = 'Module for Domain Controller validation in Azure'
    PowerShellVersion = '7.0'
    FunctionsToExport = @('Test-DCPrerequisites')
    PrivateData = @{
        PSData = @{
            Tags = @('Azure', 'DomainController', 'Validation')
            LastUpdated = '2025-02-12 18:24:19'
            UpdatedBy = 'jdyer-nuvodia'
        }
    }
}