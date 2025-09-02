# Run Code Quality Scripts

Run all four CodeQuality scripts on the specified PowerShell script file:

1. First run `Align-OperatorsVertically.ps1 -FilePath <ScriptPath>` to identify operator alignment opportunities for improved readability
2. Then run `Convert-WriteHostToColorOutput.ps1` to convert any Write-Host calls to Write-ColorOutput
3. Next run `Invoke-PowerShellCodeCleanup.ps1` to perform whitespace and formatting cleanup
4. Finally run `Invoke-PSScriptAnalyzerCheck.ps1` to perform comprehensive PSScriptAnalyzer analysis and identify any remaining code quality issues

## Fix All PSScriptAnalyzer Issues

After running the analysis, **fix ALL issues identified** by the PSScriptAnalyzer check script according to copilot-instructions.md requirements:

### Error Priority (Fix Immediately)
- **Errors**: Fix all PSScriptAnalyzer errors immediately - these can cause script failures
- Address syntax issues, undefined variables, and critical problems

### Warning Priority (Must Address)
- **PSAlignAssignmentStatement**: Align assignment statements within code blocks
- **PSPlaceCloseBrace**: Ensure closing braces follow proper placement rules (usually on new lines)
- **PSUseConsistentWhitespace**: Fix all whitespace issues including spaces before closing braces in hash tables
- **PSUseCorrectCasing**: Fix PowerShell keyword casing (e.g., `Foreach` → `foreach`)
- **PSUseBOMForUnicodeEncodedFile**: Address file encoding issues if present
- All other PSScriptAnalyzer warnings for best practices and maintainability

### Information Priority (Should Address)
- **Information**: Review and address informational items for code style improvements
- Apply PowerShell best practices and coding standards

## Final Validation
After all fixes:
- Re-run `Invoke-PSScriptAnalyzerCheck.ps1` to verify **ZERO** issues remain
- Ensure script passes all PSScriptAnalyzer checks with no errors, warnings, or critical information items
- **DO NOT** ignore any PSScriptAnalyzer issues - fix them all for complete compliance

The goal is **complete PSScriptAnalyzer AND copilot-instructions.md compliance** with all issues resolved and proper header versioning maintained.
