# =============================================================================
# Script: Get-NTFSFolderPermissions.ps1
# Created: 2025-03-06 21:06:43 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-06 21:06:43 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0
# Additional Info: Script to extract and report NTFS permissions for folders
# =============================================================================

<#
.SYNOPSIS
    Extracts and reports NTFS permissions for a specified folder and its subfolders.
.DESCRIPTION
    This script retrieves NTFS permissions for a specified folder path and all its subfolders.
    It provides a detailed report including identity references, file system rights, access control types,
    and inheritance settings. The results are displayed in the console and exported to a CSV file
    in the same directory as the script.
    
    - Traverses the specified folder structure recursively
    - Captures all NTFS permission entries for each folder
    - Exports results to a CSV file in the script's directory
.PARAMETER FolderPath
    The path to the folder for which permissions will be extracted. This parameter is mandatory.
.EXAMPLE
    .\Get-NTFSFolderPermissions.ps1 -FolderPath "C:\Important\Data"
    Retrieves NTFS permissions for C:\Important\Data and all subfolders, and exports the results to CSV.
.EXAMPLE
    .\Get-NTFSFolderPermissions.ps1 -FolderPath "\\server\share\folder"
    Retrieves NTFS permissions for the specified network share and displays results.
.NOTES
    Security Level: Low
    Required Permissions: Read access to the folders being scanned
    Validation Requirements: Verify FolderPath exists and is accessible
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
    [ValidateScript({Test-Path $_ -PathType Container})]
    [string]$FolderPath
)

# Get the script's directory to use for output files
$ScriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$OutputCsv = Join-Path -Path $ScriptDirectory -ChildPath "NTFSPermissions_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').csv"

# Initialize an array to store results
$Results = @()

# Display start message
Write-Host "Starting NTFS permissions analysis for: $FolderPath" -ForegroundColor Cyan

try {
    # Get all folders and subfolders recursively
    Write-Host "Retrieving folder structure..." -ForegroundColor Cyan
    $Folders = Get-ChildItem -Path $FolderPath -Recurse -Directory -ErrorAction Stop
    
    # Include the root folder in the list
    $Folders += Get-Item -Path $FolderPath -ErrorAction Stop
    
    $TotalFolders = ($Folders | Measure-Object).Count
    Write-Host "Found $TotalFolders folders to process" -ForegroundColor Cyan
    
    $CurrentFolder = 0
    
    # Loop through each folder and retrieve NTFS permissions
    foreach ($Folder in $Folders) {
        $CurrentFolder++
        Write-Host "Processing folder ($CurrentFolder/$TotalFolders): $($Folder.FullName)" -ForegroundColor DarkGray
        
        try {
            $Acl = Get-Acl -Path $Folder.FullName -ErrorAction Stop
            
            foreach ($Access in $Acl.Access) {
                # Create a custom object for each permission entry
                $Result = [PSCustomObject]@{
                    FolderPath       = $Folder.FullName
                    IdentityReference = $Access.IdentityReference
                    FileSystemRights  = $Access.FileSystemRights
                    AccessControlType = $Access.AccessControlType
                    IsInherited       = $Access.IsInherited
                    InheritanceFlags  = $Access.InheritanceFlags
                    PropagationFlags  = $Access.PropagationFlags
                }
                # Add the result to the array
                $Results += $Result
            }
        }
        catch {
            Write-Host "Error processing folder: $($Folder.FullName)" -ForegroundColor Yellow
            Write-Host "Error details: $_" -ForegroundColor Yellow
        }
    }
    
    # Display completion message
    Write-Host "Analysis completed. Found $($Results.Count) permission entries." -ForegroundColor Green
    
    # Display results in a human-readable table format
    $Results | Format-Table -AutoSize
    
    # Export results to CSV
    $Results | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
    Write-Host "NTFS permissions exported to: $OutputCsv" -ForegroundColor Green
}
catch {
    Write-Host "An error occurred during the NTFS permissions analysis:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}
