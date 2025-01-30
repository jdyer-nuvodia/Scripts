# Variables
$storageAccountName = "jbteststorage0"
$storageAccountKey = 's+gZISuRt6q2Ugxc9MyumGmVjGalAmwHo3+6yr6XD3+887P+0Zq5WOWyoDWhuV2zkpPwEQ+c8Z2Z+ASttP+eIQ=='
$fileShareName = "jbteststorage0"
$localDirectoryPath = "C:\Users\jdyer\OneDrive - Nuvodia\Documents\GitHub\Scripts"
$destinationPath = "scripts"
$subscriptionId = "2d9a0b3b-de9d-4acf-baad-af240553bcc7"  # Optional, specify if you have multiple subscriptions

# Set the default subscription (optional, specify if you have multiple subscriptions)
if ($subscriptionId) {
    Write-Host "Setting subscription..."
    & az account set --subscription $subscriptionId
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Subscription set successfully." -ForegroundColor Green
    } else {
        Write-Host "Failed to set subscription. Please check your subscription ID." -ForegroundColor Red
        exit
    }
}

# Check if file share exists and create it if it doesn't
Write-Host "Checking if file share exists..."
$shareExists = & az storage share exists --name $fileShareName --account-name $storageAccountName --account-key $storageAccountKey --output tsv
if ($shareExists -eq "True") {
    Write-Host "File share already exists." -ForegroundColor Yellow
} else {
    Write-Host "Creating file share..."
    & az storage share create --name $fileShareName --account-name $storageAccountName --account-key $storageAccountKey
    if ($LASTEXITCODE -eq 0) {
        Write-Host "File share created successfully." -ForegroundColor Green
    } else {
        Write-Host "Error creating file share. Please check your parameters and try again." -ForegroundColor Red
        exit
    }
}

# Upload all files from the local directory, including directory structure, to the 'scripts' directory in the Azure file share
Write-Host "Uploading files to Azure file share..."
& az storage file upload-batch --account-name $storageAccountName --account-key $storageAccountKey --destination "$fileShareName/$destinationPath" --source $localDirectoryPath

if ($LASTEXITCODE -eq 0) {
    Write-Host "Files uploaded successfully." -ForegroundColor Green
} else {
    Write-Host "Error uploading files. Please check your parameters and try again." -ForegroundColor Red
}
