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
    "C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\365"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\azure"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\domain"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\files-and-folders"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\local"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\network"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\nuvodiaonly"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\software"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\365\add-userlistto365group"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\365\delete-mailboxes"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\365\diagnose-mailboxfolderassistant"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\365\get-allmailboxforwardingrules"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\365\get-calendarpermissions"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\365\get-mailboxesexist"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\365\get-mailboxfolderlist"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\365\getfullmailboxattributes"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\365\grant-calendarpermissions"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\365\grant-rmtomailboxeditcalendarpermissions"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\365\remediate-365account"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\365\remove-allmailboxpermissions"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\azure\create-testdomaincontroller"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\azure\delete-testdomaincontroller"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\domain\change-aduserpassword"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\domain\copy-aduser"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\domain\create-aduser"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\domain\remove-groupsfromdisabledusers"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\files-and-folders\add-folderstopath"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\files-and-folders\delete-allfilesindirectory"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\files-and-folders\delete-oldfiles"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\files-and-folders\get-foldersizes"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\files-and-folders\get-ntfspermissions"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\files-and-folders\unprotect-rmsfile"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\local\change-localuserpassword"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\local\create-localaccount"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\local\repair-windowsos"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\network\test-pingextended"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\nuvodiaonly\create-nuvodialocalaccount"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\software\get-installedsoftware"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\software\get-wiztreeportable"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\software\reinstall-onedrive"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\software\remove-adobeacrobatreader"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\my-scripts"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\files-and-folders\get-machinepath"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\files-and-folders\get-userpath"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\files-and-folders\remove-folderfrompath"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\my-scripts\copy-filestoazurestorage"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\my-scripts\delete-oldscreenshots"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\my-scripts\mount-azurestorage"
"C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\my-scripts\transfer-filestodownloadsfolder"
)

Remove-FromPath -DirectoriesToRemove $pathsToRemove
