# Install Azure PowerShell if not already installed
# Install-Module -Name Az

# Sign in to Azure (if not already signed in)
Connect-AzAccount

# Set the context to the correct subscription if needed
# Set-AzContext -SubscriptionId "YourSubscriptionId"

# Deploy ARM template using Azure PowerShell
New-AzResourceGroupDeployment `
-Name AGDeployment `
-ResourceGroupName JB-TEST-RG2 `
-TemplateFile "C:\Users\jdyer\OneDrive - Nuvodia\Documents\GitHub\Scripts\Azure\Deploy-ActionGroup\AGTemplate-JB-TEST-RG2\template.json" `
-TemplateParameterFile "C:\Users\jdyer\OneDrive - Nuvodia\Documents\GitHub\Scripts\Azure\Deploy-ActionGroup\AGTemplate-JB-TEST-RG2\parameters.json"
