# Import the SharePoint PnP PowerShell module
Import-Module SharePointPnPPowerShellOnline

# Path to the creds.json file
$creds_file = ".\creds.json"  # creds.json file should not have any spaces or tabs

# Example of the creds.json file (Copy Below line and paste it to creds.json, replace the values with your own values)
# {"sharePointUrl": "https://your_org.sharepoint.com/sites/some_sharepoint_site_name","Username": "example@example.com","Password": "SomeVeryStrongPassword"}

# Check if the creds.json file exists
if (-not (Test-Path $creds_file)) {
    Write-Host "The creds.json file does not exist."
    return
}

# Load the credentials from the creds.json file
$creds = Get-Content -Raw -Path  $creds_file | ConvertFrom-Json

try {
    # Connect to the SharePoint site using the provided credentials
    Connect-PnPOnline -Url $creds.sharePointUrl -Credentials (New-Object System.Management.Automation.PSCredential($creds.Username, (ConvertTo-SecureString $creds.Password -AsPlainText -Force)))
}
catch {
    Write-Host "Error connecting to SharePoint site: $_"
}
# Connect to the SharePoint site using the provided credentials
# Connect-PnPOnline -Url $creds.sharePointUrl -Credentials (New-Object System.Management.Automation.PSCredential($creds.Username, (ConvertTo-SecureString $creds.Password -AsPlainText -Force)))
