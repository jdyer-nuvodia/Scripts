@{
    RootModule = 'DC-Validation.psm1'
    ModuleVersion = '1.6'
    GUID = '7a1b2c3d-4e5f-48aa-9b1a-0c1d2e3f4a5b'
    Author = 'jdyer-nuvodia'
    Description = 'Module for validating Domain Controller prerequisites in Azure'
    PowerShellVersion = '7.0'
    FunctionsToExport = @('Test-DCPrerequisites')
    PrivateData = @{
        PSData = @{
            Tags = @('Azure', 'DomainController', 'Validation')
            LastUpdated = '2025-02-12 17:50:17'
            UpdatedBy = 'jdyer-nuvodia'
        }
    }
}