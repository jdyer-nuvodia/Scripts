# get-foldersizes.ps1

# Check for elevated privileges and restart if necessary
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Elevated privileges required. Attempting to restart script as Administrator..."
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" `"$Path`"" -Verb RunAs
    Exit
}

param (
    [string]$Path = "C:\",
    [int]$MaxDepth = 10
)

# Set global error action preference
$ErrorActionPreference = 'SilentlyContinue'

# Enable required privileges
$privilege = @"
using System;
using System.Runtime.InteropServices;

public class Privileges {
    [DllImport("advapi32.dll", SetLastError = true)]
    private static extern bool AdjustTokenPrivileges(IntPtr TokenHandle, 
        bool DisableAllPrivileges, 
        ref TOKEN_PRIVILEGES NewState, 
        uint BufferLength, 
        IntPtr PreviousState, 
        IntPtr ReturnLength);

    [DllImport("advapi32.dll", SetLastError = true)]
    private static extern bool LookupPrivilegeValue(string lpSystemName, 
        string lpName, 
        ref LUID lpLuid);

    [DllImport("advapi32.dll", SetLastError = true)]
    private static extern bool OpenProcessToken(IntPtr ProcessHandle, 
        uint DesiredAccess, 
        out IntPtr TokenHandle);

    [StructLayout(LayoutKind.Sequential)]
    private struct LUID {
        public uint LowPart;
        public int HighPart;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct TOKEN_PRIVILEGES {
        public uint PrivilegeCount;
        public LUID Luid;
        public uint Attributes;
    }

    public const uint SE_PRIVILEGE_ENABLED = 0x00000002;
    public const uint TOKEN_ADJUST_PRIVILEGES = 0x00000020;
    public const uint TOKEN_QUERY = 0x00000008;

    public static bool EnablePrivilege(string privilegeName) {
        IntPtr tokenHandle;
        if (!OpenProcessToken(System.Diagnostics.Process.GetCurrentProcess().Handle, 
            TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, out tokenHandle)) {
            return false;
        }

        TOKEN_PRIVILEGES tokenPrivileges;
        tokenPrivileges.PrivilegeCount = 1;
        tokenPrivileges.Luid = new LUID();
        tokenPrivileges.Attributes = SE_PRIVILEGE_ENABLED;

        if (!LookupPrivilegeValue(null, privilegeName, ref tokenPrivileges.Luid)) {
            return false;
        }

        return AdjustTokenPrivileges(tokenHandle, false, ref tokenPrivileges, 0, IntPtr.Zero, IntPtr.Zero);
    }
}
"@

Add-Type $privilege

# Enable backup privileges to access all directories
[Privileges]::EnablePrivilege("SeBackupPrivilege")
[Privileges]::EnablePrivilege("SeRestorePrivilege")
[Privileges]::EnablePrivilege("SeTakeOwnershipPrivilege")

Write-Host "Analyzing folders in: $Path with elevated privileges"

function Get-FolderSizes {
    param (
        [string]$FolderPath,
        [int]$CurrentDepth = 0
    )

    if ($CurrentDepth -ge $MaxDepth) {
        return $null
    }

    # Create a backup token for accessing protected directories
    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $context = $identity.Impersonate()

    try {
        # Get largest file in current directory first
        $currentFiles = Get-ChildItem -Path $FolderPath -File -Force
        $largestCurrentFile = $currentFiles | Sort-Object -Property Length -Descending | Select-Object -First 1
        if ($largestCurrentFile) {
            Write-Host "`nLargest file in $FolderPath :"
            Write-Host "Name: $($largestCurrentFile.Name)"
            Write-Host "Size: $([math]::round($largestCurrentFile.Length / 1GB, 2)) GB ($([math]::round($largestCurrentFile.Length / 1MB, 2)) MB)"
        } else {
            Write-Host "`nNo files found directly in $FolderPath"
        }

        $folders = Get-ChildItem -Path $FolderPath -Directory -Force
        $folderSizes = @()
        $totalItems = ($folders | Measure-Object).Count
        Write-Host "`nFound $totalItems subfolders to process..."
        $processedCount = 0
        
        foreach ($folder in $folders) {
            try {
                # Get immediate files in the current directory (non-recursive)
                $currentFiles = Get-ChildItem -Path $folder.FullName -File -Force
                $largestCurrentFile = $currentFiles | Sort-Object -Property Length -Descending | Select-Object -First 1

                # Get all files recursively for total size calculation
                $allFiles = Get-ChildItem -Path $folder.FullName -File -Recurse -Force
                $subfolders = Get-ChildItem -Path $folder.FullName -Directory -Recurse -Force
                $folderSize = ($allFiles | Measure-Object -Property Length -Sum).Sum
                
                $folderSizes += [PSCustomObject]@{
                    Folder = $folder.FullName
                    SizeGB = [math]::round($folderSize / 1GB, 2)
                    TotalSubfolders = ($subfolders | Measure-Object).Count
                    TotalFiles = ($allFiles | Measure-Object).Count
                    LargestFile = if ($largestCurrentFile) {
                        [PSCustomObject]@{
                            Name = $largestCurrentFile.Name
                            Path = $largestCurrentFile.FullName
                            SizeGB = [math]::round($largestCurrentFile.Length / 1GB, 2)
                            SizeMB = [math]::round($largestCurrentFile.Length / 1MB, 2)
                        }
                    } else { $null }
                }
                $processedCount++
                Write-Host "`rProcessed $processedCount of $totalItems folders..." -NoNewline
            } catch {
                Write-Warning "Access to the path '$($folder.FullName)' is denied despite elevated privileges. Error: $($_.Exception.Message)"
            }
        }
        Write-Host "`nCompleted processing $processedCount folders."
        return $folderSizes
    }
    finally {
        if ($context) {
            $context.Undo()
        }
    }
}

[Rest of the script remains the same as previous version]