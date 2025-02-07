$targetDir = "./clouddrive/scripts"

# Remove all files
Get-ChildItem -Path $targetDir -File -Recurse | ForEach-Object {
Remove-Item -Path $_.FullName -Force
}

# Remove all folders (empty ones will be removed first)
Get-ChildItem -Path $targetDir -Directory -Recurse | Sort-Object -Property FullName -Descending | ForEach-Object {
Remove-Item -Path $_.FullName -Force
}