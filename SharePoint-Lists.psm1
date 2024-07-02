# Just Import as Module: Import-Module -Name .\SharePoint-Lists.psm1

# Example for every Function are provided below

# Create-SharePoint-Fields
function Create-SharePoint-Fields {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ListName,
        [hashtable]$Columns,
        [hashtable]$Choice_Options = $null,
        [hashtable]$LookUpAttributes = $null  
    )
    
    # Create a new List if not exists
    $List = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue

    if ($null -eq $List) {
        Write-Host "Creating a new list..."
        New-PnPList -Title $ListName -Template GenericList
    }

    # Add a new column to the list if not exists
    foreach ($column in $Columns.Keys) {
        $Column = $column
        $Type = $Columns.$Column
        $ColumnExists = Get-PnPField -List $ListName -Identity $Column -ErrorAction SilentlyContinue
        if ($null -eq $ColumnExists) {
            Write-Host "Adding column $Column..."
            if ($Type -eq "Choice") {
                Add-PnPField -List $ListName -Type $Type -DisplayName $Column -InternalName $Column -Choices $Choice_Options.$Column -AddToDefaultView 
            }
            elseif ($Type -eq "Lookup") {
                Add-PnPField -List $ListName -Type $Type -DisplayName $Column -InternalName $Column -AddToDefaultView
                Set-PnPField -List $ListName -Identity $Column -Values @{LookupList = (Get-PnPList $LookUpAttributes.$Column[0]).Id.ToString(); LookupField = $LookUpAttributes.$Column[1] }
            }
            else {
                Add-PnPField -List $ListName -Type $Type -DisplayName $Column -InternalName $Column -AddToDefaultView
            }
            Write-Host "Column $Column added."
        }
        else {
            Write-Host "Column $Column already exists."
        }
    }

    Write-Host "List $ListName created successfully."

    # Pause for 2 seconds
    Start-Sleep -Seconds 2


}
# Example:
# Create-SharePoint-Fields -ListName "Leave Requests" -Columns $Columns -Choice_Options $Choice_Options -LookUpAttributes $LookUpAttributes
# Where $Columns, $Choice_Options, and $LookUpAttributes are defined as:
# $Columns = @{
#     Leave_Type       = "Lookup"
#     Leave_Start_Date = "DateTime"
#     Leave_Status     = "Choice"
# }
#
# $Choice_Options = @{
#     "Leave_Status" = @("Pending", "Approved", "Rejected", "Depreciated")
# }
#
# $LookUpAttributes = @{
#     Leave_Type = @("Leave Types", "Title")
# }

'''
-------------------------------------------------------------------------------
'''

# Add-SharePoint-Items Function 
function Add-SharePoint-Items {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ListName,
        [array]$Items
    )

    # Check if the list exists
    $List = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue

    if ($null -eq $List) {
        Write-Host "List $ListName does not exist."
        return
    }

    # Add items to the list
    foreach ($item in $Items) {
        $ItemExists = Get-PnPListItem -List $ListName -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($item.Title)</Value></Eq></Where></Query></View>"
        if ($null -eq $ItemExists) {
            Write-Host "Adding item $($item.Title)..."
            Add-PnPListItem -List $ListName -Values $item
            Write-Host "Item $($item.Title) added."
        }
        else {
            Write-Host "Item $($item.Title) already exists."
        }
    }

    Write-Host "Items added successfully."

    # Pause for 2 seconds
    Start-Sleep -Seconds 2

}

# Example:
# Add-SharePoint-Items -ListName "Leave Types" -Items $Items
# Where $Items is defined as:
# $Items = @(
#     @{
#          Title = "Annual Leave" 
#          "Days" = 21
#     },
#     @{
#          Title = "Sick Leave" 
#          "Days" = 14
#     }
# )

'''
-------------------------------------------------------------------------------
'''

# Remove-SharePoint-List Function
function Remove-SharePoint-List {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ListName
    )

    # Check if the list exists
    $List = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue

    if ($null -eq $List) {
        Write-Host "List $ListName does not exist."
        return
    }

    # Remove the list
    Remove-PnPList -Identity $ListName -Force

    Write-Host "List $ListName deleted successfully."

    # Pause for 2 seconds
    Start-Sleep -Seconds 2

}

# Example:
# Remove-SharePoint-List -ListName "Leave Types"

'''
-------------------------------------------------------------------------------
'''

# Remove-SharePoint-Lists Function
function Remove-SharePoint-Lists {
    param (
        [Parameter(Mandatory = $true)]
        [array]$ListNames
    )

    foreach ($ListName in $ListNames) {
        Remove-SharePoint-List -ListName $ListName
    }

    Write-Host "All lists deleted successfully."

    # Pause for 2 seconds
    Start-Sleep -Seconds 2

}

# Example:
# Remove-SharePoint-Lists -ListNames @("Leave Types", "Leave Requests")

'''
-------------------------------------------------------------------------------
'''

# Export the functions
Export-ModuleMember -Function *