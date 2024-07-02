.\connect.ps1

# Import the Custome Module
Import-Module .\SharePoint-Lists.psm1

# Lists to be deleted
$ListNames = 
    @(
        "Leave Types", 
        "Leave Requests",
        "Leave Balance"
    )

# Delete the Lists
Remove-SharePoint-Lists -ListNames $ListNames   

# Disconnect from SharePoint
Disconnect-PnPOnline