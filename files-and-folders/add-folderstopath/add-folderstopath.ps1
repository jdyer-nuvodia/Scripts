[Previous header comments remain the same...]

function Convert-ToPascalCase {
    param([string]$text)
    
    Write-Verbose "Converting to PascalCase: $text"
    
    # Split by common delimiters
    $words = $text -split '[-_\s]'
    
    # Convert each word to proper case
    $words = $words | ForEach-Object { 
        if ($_.Length -gt 0) {
            $_.Substring(0,1).ToUpper() + $_.Substring(1).ToLower()
        }
    }
    
    # Rejoin with hyphens for PowerShell convention
    $result = $words -join '-'
    Write-Verbose "Converted to: $result"
    return $result
}

function Rename-FolderWithCase {
    param(
        [string]$folderPath
    )
    
    try {
        $folder = Get-Item -LiteralPath $folderPath
        $parentPath = Split-Path -Path $folderPath -Parent
        $currentName = Split-Path -Path $folderPath -Leaf
        
        # Skip if it's a file
        if (!$folder.PSIsContainer) {
            return
        }

        # Convert name to proper case
        $newName = Convert-ToPascalCase -text $currentName
        
        # Skip if name wouldn't change (FIXED: Now compares with actual new name)
        if ($newName.Equals($currentName, [StringComparison]::Ordinal)) {
            Write-Host "Skipping '$currentName' - already in correct case" -ForegroundColor Yellow
            return
        }
        
        Write-Host "Need to rename '$currentName' to '$newName'" -ForegroundColor Cyan
        $newPath = Join-Path -Path $parentPath -ChildPath $newName
        
        # Handle case where only case is different (needs temp rename)
        if ($newPath.ToLower() -eq $folderPath.ToLower()) {
            $tempName = "_temp_" + [Guid]::NewGuid().ToString().Substring(0,8)
            $tempPath = Join-Path -Path $parentPath -ChildPath $tempName
            
            if ($PSCmdlet.ShouldProcess($folderPath, "Rename to temp folder '$tempPath'")) {
                Write-Verbose "Temporary rename: '$folderPath' -> '$tempPath'"
                Rename-Item -LiteralPath $folderPath -NewName $tempName -ErrorAction Stop
                
                Write-Verbose "Final rename: '$tempPath' -> '$newPath'"
                Rename-Item -LiteralPath $tempPath -NewName $newName -ErrorAction Stop
                Write-Host "Successfully renamed: '$currentName' -> '$newName'" -ForegroundColor Green
            }
        }
        else {
            if ($PSCmdlet.ShouldProcess($folderPath, "Rename to '$newPath'")) {
                Rename-Item -LiteralPath $folderPath -NewName $newName -ErrorAction Stop
                Write-Host "Successfully renamed: '$currentName' -> '$newName'" -ForegroundColor Green
            }
        }
    }
    catch {
        Write-Error "Error renaming folder '$folderPath': $_"
        Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
    }
}

[Rest of the script remains the same...]