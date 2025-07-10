# Define the resource group and VM name
$resourceGroupName = "JB-TEST-RG2"
$vmName = "JB-TEST-DC01"

# Get the auto-shutdown schedule for the VM
$autoShutdown = Get-AzResource
-ResourceGroupName $resourceGroupName `
    -ResourceType "Microsoft.DevTestLab/schedules" `
    -ResourceName "$vmName/auto-shutdown" `
    -ErrorAction SilentlyContinue

# Check if auto-shutdown is enabled
if ($autoShutdown) {
    Write-Output "Auto-shutdown is enabled for VM: $vmName"
    Write-Output "Details: $($autoShutdown.Properties | ConvertTo-Json -Depth 2)"
} else {
    Write-Output "Auto-shutdown is NOT enabled for VM: $vmName"
}
