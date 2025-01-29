#Sets Script Execution Policy to RemoteSigned.
Set-ExecutionPolicy RemoteSigned

#Installs the AIPService Powrshell Module.
Install-Module -Name AIPService

#Imports the AIPService module.
Import-Module AIPService

#Connects to AIPService, will prompt for credentials.
Connect-AIPService

#Enables SuperUserFeature for the AIPService Module. 
Enable-AipServiceSuperUserFeature

#Sets the environment variable for the downloads folder.
$downloadsFolder = Join-Path $env:USERPROFILE "Downloads"

#Creates a folder in Downloads for storing the input file.
$inputFolder = Join-Path $downloadsFolder "Input"
if (-not (Test-Path -Path $inputFolder)) {
    New-Item -Path $inputFolder -ItemType Directory
	Write-Output "Folder created successfully: $inputFolder"
} else {
    Write-Output "Folder already exists: $inputFolder"
}
	
#Creates a folder in Downloads for the output file.
$outputFolder = Join-Path $downloadsFolder "Output"
if (-not (Test-Path -Path $outputFolder)) {
    New-Item -Path $outputFolder -ItemType Directory
	Write-Output "Folder created successfully: $outputFolder"
} else {
    Write-Output "Folder already exists: $outputFolder"
}
	
#Unprotects the file(s).
Unprotect-RMSFile -Folder $inputFolder -OutputFolder $outputFolder -Recurse

