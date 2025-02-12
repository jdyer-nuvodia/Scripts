@{
    RootModule = 'DC-Validation.psm1'
    ModuleVersion = '1.5'
    GUID = '7a1b2c3d-4e5f-6g7h-8i9j-0k1l2m3n4o5p'
    Author = 'jdyer-nuvodia'
    Description = 'Module for validating Domain Controller prerequisites in Azure'
    PowerShellVersion = '7.0'
    FunctionsToExport = @('Test-DCPrerequisites')
    PrivateData = @{
        PSData = @{
            Tags = @('Azure', 'DomainController', 'Validation')
        }
    }
}