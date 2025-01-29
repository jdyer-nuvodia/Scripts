function Remove-FromPath {
    param(
        [Parameter(Mandatory=$true)]
        [string[]]$DirectoriesToRemove
    )

    $currentPath = [System.Environment]::GetEnvironmentVariable("Path", "Machine")
    $pathArray = $currentPath -split ';'

    $newPath = $pathArray | Where-Object { $dir = $_; -not ($DirectoriesToRemove | Where-Object { $dir -eq $_ })} | Join-String -Separator ';'

    if ($newPath -ne $currentPath) {
        [System.Environment]::SetEnvironmentVariable("Path", $newPath, "Machine")
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine")
        Write-Host "Specified directories have been removed from the PATH."
    } else {
        Write-Host "No changes were made to the PATH. Specified directories were not found."
    }
}

# Example usage:
$pathsToRemove = @(
    "C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\1my-scripts\copy-filestoazurestorage"
	"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\1my-scripts\delete-oldscreenshots"
	"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\1my-scripts\mount-azurestorage"
	"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\1my-scripts\transfer-filestodownloadsfolder"
)

Remove-FromPath -DirectoriesToRemove $pathsToRemove
