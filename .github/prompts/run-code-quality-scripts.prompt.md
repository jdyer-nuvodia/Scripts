# Run Code Quality Scripts

Run both CodeQuality scripts on the specified PowerShell script file:

1. First run `Convert-WriteHostToColorOutput.ps1` to convert any Write-Host calls to Write-ColorOutput
2. Then run `Invoke-PowerShellCodeCleanup.ps1` to perform whitespace and formatting cleanup

Please run these scripts in sequence on the target PowerShell file to ensure complete code quality compliance.
