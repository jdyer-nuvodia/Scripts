---
description: "Run Invoke-PowerShellCodeCleanup.ps1 to fix whitespace and formatting issues"
mode: "agent"
---

# Run PowerShell Code Cleanup

Execute the `Invoke-PowerShellCodeCleanup.ps1` script to automatically fix whitespace, newline, and formatting issues in PowerShell scripts.

## Cleanup Process:

### Pre-Cleanup Assessment:
1. **Identify target files**: ${input:targetFiles:Which PowerShell files need cleanup? (provide file paths or use current selection)}
2. **Backup verification**: Ensure files are committed to git or backed up
3. **Current issues**: Identify existing formatting problems

### Execute Cleanup:
Use -FilePath to specify the target script directory and run the cleanup tool:

```powershell
& "c:\Users\jdyer\OneDrive - Nuvodia\Documents\GitHub\Scripts\Development\CodeQuality\Invoke-PowerShellCodeCleanup.ps1"
-FilePath "${input:scriptDirectory:Directory containing the target scripts}"
```

**OR**

Run the cleanup script from the target script's directory:
```powershell
Set-Location "${input:scriptDirectory:Directory containing the target scripts}"
& "c:\Users\jdyer\OneDrive - Nuvodia\Documents\GitHub\Scripts\Development\CodeQuality\Invoke-PowerShellCodeCleanup.ps1"
```

### Post-Cleanup Validation:
1. **Review changes**: Check what formatting changes were made
2. **Verify functionality**: Ensure script functionality is preserved
3. **PSScriptAnalyzer check**: Confirm cleanup resolved formatting-related warnings
4. **Git diff review**: Review all changes before committing

## Target Files:
${selection}

## Execution Steps:
1. **Navigate to script directory** and run the cleanup tool
2. **Review all changes** made by the cleanup process
3. **Test script functionality** to ensure nothing was broken
4. **Run PSScriptAnalyzer** to verify formatting issues are resolved
5. **Provide summary** of changes made and any remaining issues

## Safety Notes:
- Always have files backed up or committed to git before running cleanup
- The cleanup tool fixes whitespace, indentation, and newline issues
- Review changes carefully before committing to ensure no functionality was affected
- If any issues arise, the original files can be restored from git

## Expected Fixes:
- Trailing whitespace removal
- Consistent indentation
- Proper newline placement
- Missing newlines that cause PSScriptAnalyzer warnings
- Tab to space conversion (if configured)

Execute the cleanup process and report on all changes made.
