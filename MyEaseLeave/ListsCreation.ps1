.\connect.ps1

# Import the Custome Module
Import-Module .\SharePoint-Lists.psm1

# List1: Leave Types List
$ListName = "Leave Types"

# Fields
$Leave_Type_Columns = @{
    Leave_Type  = "Text"
    Max_Leaves  = "Number"
    Description = "Note"
}

# Create the List
Create-SharePoint-Fields -ListName $ListName -Columns $Leave_Type_Columns

# Change the Title Column to Leave_Code  
Set-PnPField -List $ListName -Identity "Title" -Values @{Title = "Leave_Code" }

# Add Items to the List
$Items = @(
    @{
        Title = "AL"
        Leave_Type = "Annual Leave"
        Max_Leaves = 20
        Description = "Annual Leave"
    },
    @{
        Title = "SL"
        Leave_Type = "Sick Leave"
        Max_Leaves = 10
        Description = "Sick Leave"
    },
    @{
        Title = "PL"
        Leave_Type = "Paternity Leave"
        Max_Leaves = 5
        Description = "Paternity Leave"
    }
)

Add-SharePoint-Items -ListName $ListName -Items $Items

'''
-------------------------------------------------------------------------------
'''

# List2: Leave Requests List
$ListName = "Leave Requests"

# Fields
$Columns = @{
    Leave_Type       = "Lookup"
    Leave_Start_Date = "DateTime"
    Leave_End_Date   = "DateTime"
    Leave_Duration   = "Number"
    Leave_Status     = "Choice"
    Leave_Reason     = "Note"
    Document         = "Attachments"
}

# Choice Options
$Choice_Options = @{
    "Leave_Status" = @("Pending", "Approved", "Rejected", "Depreciated")
}

# LookUp Attributes
$LookUpAttributes = @{
    "Leave_Type" = @("Leave Types", "Leave_Type")
}

# Create the List
Create-SharePoint-Fields -ListName $ListName -Columns $Columns -Choice_Options $Choice_Options -LookUpAttributes $LookUpAttributes

'''
-------------------------------------------------------------------------------
'''

# List3: Leave Balance List
$ListName = "Leave Balance"

# Fields
$Columns = @{
    Employee_Email = "Text"
    Total_Leaves   = "Number" 
}

# Add the Leave Types as columns in the Leave Balance List
foreach ($Item in $Items) {
    $Columns.Add($Item.Title, "Number")
}

# Create the List
Create-SharePoint-Fields -ListName $ListName -Columns $Columns


# Disconnect from SharePoint
Disconnect-PnPOnline


