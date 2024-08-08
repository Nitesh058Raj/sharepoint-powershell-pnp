# Just Import as Module: Import-Module -Name .\Modules\SharePoint-Lists.psm1
# Example for every Function are provided after each Function

# Create-SharePoint-Fields
function Create-SharePoint-Fields {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ListName,
        [hashtable]$Columns,
        [hashtable]$Choice_Options = $null,
        [hashtable]$LookUpAttributes = $null,
        [hashtable]$Formulas = $null
    )
    
    try {

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
                if ($Type -eq "Choice" -or $Type -eq "MultiChoice") {
                    Add-PnPField -List $ListName -Type $Type -DisplayName $Column -InternalName $Column -Choices $Choice_Options.$Column -AddToDefaultView 
                }
                elseif ($Type -eq "Lookup") {
                    Add-PnPField -List $ListName -Type $Type -DisplayName $Column -InternalName $Column -AddToDefaultView
                    Set-PnPField -List $ListName -Identity $Column -Values @{LookupList = (Get-PnPList $LookUpAttributes.$Column[0]).Id.ToString(); LookupField = $LookUpAttributes.$Column[1] }
                } elseif ($Type -eq "Calculated") {
                    Add-PnPField -List $ListName -Type $Type -DisplayName $Column -InternalName $Column -Formula $Formulas.$Column -AddToDefaultView
                } else {
                    Add-PnPField -List $ListName -Type $Type -DisplayName $Column -InternalName $Column -AddToDefaultView
                }
                Write-Host "Column $Column added."
            }
            else {
                Write-Host "Column $Column already exists."
            }

            Start-Sleep -Seconds 5

        }

        Write-Host "Columns added successfully."
        # Pause for 2 seconds
        Start-Sleep -Seconds 2

    } catch {
        Write-Host "Error: $_"
    }


}
# Example:
# Create-SharePoint-Fields -ListName $ListName -Columns $Columns -Choice_Options $Choice_Options -LookUpAttributes $LookUpAttributes -Formulas $Formulas
# Where $ListName, $Columns, $Choice_Options, $LookUpAttributes and $Formulas are defined as:
#
# $ListName = "Test List"
#
# $Columns = @{
#   "Title" = "Text"                            # Pre-Defined (Single Line of Text)
#   "Description" = "Note"                      # Works Fine (Multiple Lines of Text)
#   "Favorite_Color" = "Choice"                 # Works Fine (Single Choice: Select One Option)
#   "BirthDate" = "DateTime"                    # Works Fine (Date and Time)
#   "Attched_Person" = "Lookup"                 # Works Fine (Lookup: Choose from another list)
#   "Favorite_Number" = "Number"                # Works Fine (Integer)
#   "Id2" = "Counter"                           # Not Supported (Need more information)
#   "Is_Active" = "Boolean"                     # Works Fine (Yes/No)
#   "Balance" = "Currency"                      # Works Fine (Currency)
#   "Link" = "URL"                              # Works Fine (Hyperlink)
#   "T1" = "Threading"                          # Error (Not Supported)
#   "G1" = "Guid"                               # Works Fine (Guid)
#   "Some_Lauguage" = "MultiChoice"             # Works Fine (MultiChoice: Select Multiple Options)
#   "Person" = "User"                           # Works Fine (User or Group)
#   "Grid" = "GridChoice"                       # Added But Not Working (Or Not know how to use it)
#   "Location" = "Location"                     # Works Fine (Location)
#   "File" = "File"                             # Not Supported (Need more information)
#   "Image" = "Image"                           # Not Supported
#   "Recurrence" = "Recurrence"                 # Not Supported For Form (Need more information) (Recurrence: 0/1)
#   "CrossProjectLink" = "CrossProjectLink"     # Not Supported For Form (Need more information) (CrossProjectLink: 0/1)
#   "ModStat" = "ModStat"                       # Works Fine (Moderation Status: Approved, Rejected, Pending)
#   "Error" = "Error"                           # Error (Not Supported)
#   "ContentType" = "ContentTypeId"             # Already Exists
#   "PageSeparator" = "PageSeparator"           # For Survey List Only (Need more information)
#   "ThreadIndex" = "ThreadIndex"               # Error (Not Supported)    
#   "WorkflowStatus" = "WorkflowStatus"         # Error (Not Supported)
#   "OutcomeChoice" = "OutcomeChoice"           # Exists But Don't know the use cases | WorkFlow Outcome (Approved, Rejected)
#   "AllDayEvent" = "AllDayEvent"               # Works But Not supported to Form (Yes/No)
#   "WorkflowEventType" = "WorkflowEventType"   # Exists But Don't know the use cases | WorkFlow Event Type (Item Added, Item Updated, Item Deleted, etc.)
#   "Geolocation" = "Geolocation"               # Not Supported (Need more information)
#   "Outcome" = "OutcomeChoice"                 # Exists But Don't know the use cases | WorkFlow Outcome (Approved, Rejected)
#   "Thumbnail" = "Thumbnail"                   # Works Fine (Thumbnail: Image)
#   "MaxItems" = "MaxItems"                     # Error (Not Supported)
#   "Calculated_1" = "Calculated"               # Works Fine (Calculated) 
# }
#
# $Choice_Options = @{
#   "Favorite_Color" = @("Red", "Green", "Blue", "Yellow", "Cyan", "Magenta", "Lime", "Orange", "Purple", "Pink")
#   "Some_Lauguage" = @("English", "Spanish", "French", "German", "Italian", "Dutch", "Portuguese", "Russian", "Chinese", "Japanese")
# }
#
# $LookUpAttributes = @{
#   "Attched_Person" = @("Test1", "Title")
# }
# 
# $Formulas = @{
#   "Calculated_1" = "=[Title]"
# }


#########################################################################################################

# Add-SharePoint-Items Function 
function Add-SharePoint-Items {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ListName,
        [array]$Items,
        [bool]$isLookup = $false,
        [string]$LookupList = $null, 
        [string]$LookupField = "Title"
    )


    try {

        # Check if the list exists
        $List = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue

        if ($null -eq $List) {
            Write-Host "List $ListName does not exist."
            return
        }

        if($isLookup) {

            # Check if the lookup list exists
            $LookupList = Get-PnPList -Identity $LookupList -ErrorAction SilentlyContinue
            if($null -eq $LookupList) {
                Write-Host "Lookup List $LookupList does not exist."
                return
            }

            $LookupList_Data = Get-PnPListItem -List $LookupList -Fields $LookupField

            if($null -eq $LookupList_Data) {
                Write-Host "Lookup List $LookupList is empty."
                return
            }

            foreach($item in $Items) {
                foreach($lookupItem in $LookupList_Data) {
                    if($item[1] -eq $lookupItem[$LookupField]) {
                        $item[1] = $lookupItem.Id
                        break
                    }
                }
            }

            foreach($item in $Items) {
                $ItemExists = Get-PnPListItem -List $ListName -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($item.Title)</Value></Eq></Where></Query></View>"
                if ($null -eq $ItemExists) {
                    Write-Host "Adding item $($item[0])..."
                    Add-PnPListItem -List $ListName -Values @{"Title" = $item.Title; $item.Keys[1] = $item[1]}
                    Write-Host "Item $($item[0]) added."
                }
                else {
                    Write-Host "Item $($item[0]) already exists."
                }
            }

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

    } catch {
        Write-Host "Error: $_"
    }

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


#########################################################################################################


# Remove-SharePoint-List Function
function Remove-SharePoint-List {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ListName
    )

    try {
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

    } catch {
        Write-Host "Error: $_"
    }

}

# Example:
# Remove-SharePoint-List -ListName "Leave Types"

#########################################################################################################

# Remove-SharePoint-Lists Function
function Remove-SharePoint-Lists {
    param (
        [Parameter(Mandatory = $true)]
        [array]$ListNames
    )

    try {

        foreach ($ListName in $ListNames) {
            Remove-SharePoint-List -ListName $ListName
        }

        Write-Host "All lists deleted successfully."

        # Pause for 2 seconds
        Start-Sleep -Seconds 2

    } catch {
        Write-Host "Error: $_"
    }

}

# Example:
# Remove-SharePoint-Lists -ListNames @("Leave Types", "Leave Requests")

#########################################################################################################

# Add-SharePoint-List Function
function Add-SharePoint-List {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ListName
    )

    try {
        # Check if the list exists
        $List = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue

        if ($null -ne $List) {
            Write-Host "List $ListName already exists."
            return
        }

        # Add a new list
        New-PnPList -Title $ListName -Template GenericList -Url "lists/$ListName"

        Write-Host "List $ListName created successfully."

        # Pause for 2 seconds
        Start-Sleep -Seconds 2

    } catch {
        Write-Host "Error: $_"
    }

}

# Example:
# Add-SharePoint-List -ListName "Leave Types"

#########################################################################################################

# Remove-SharePoint-List-Fields Function
function Remove-SharePoint-List-Fields {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ListName,
        [array]$Columns
    )

    try {

        # Check if the list exists
        $List = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue

        if ($null -eq $List) {
            Write-Host "List $ListName does not exist."
            return
        }

        # Remove columns from the list
        foreach ($column in $Columns) {
            $ColumnExists = Get-PnPField -List $ListName -Identity $column -ErrorAction SilentlyContinue
            if ($null -ne $ColumnExists) {
                Write-Host "Removing column $column..."
                Remove-PnPField -List $ListName -Identity $column -Force
                Write-Host "Column $column removed."
            }
            else {
                Write-Host "Column $column does not exist."
            }
        }

        Write-Host "Columns removed successfully."

        # Pause for 2 seconds
        Start-Sleep -Seconds 2

    } catch {
        Write-Host "Error: $_"
    }

}

# Break-Inheritance-List-Item Function
function Break-Inheritance-List-Item {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ListItemName,
        [Parameter(Mandatory = $true)]
        [string]$ListName
    )

    try {

        $query = "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$ListItemName</Value></Eq></Where></Query></View>"
        $ListItemDetails = Get-PnPListItem -List $ListName -Query $query -ErrorAction SilentlyContinue
        $ListItem = Get-PnPListItem -List $ListName -Id $ListItemDetails.Id -ErrorAction SilentlyContinue

        if ($null -eq $ListItem) {
            Write-Host "List Item $ListItemName not found in list $ListName."
            return
        }    
        else {
            # Write-Host "List Item $ListItemName found in list $ListName."

            $ListItem.BreakRoleInheritance($true, $true)
            $ListItem.Update()
            Invoke-PnPQuery

            Write-Host "Inheritance broken successfully for list item $ListItemName in list $ListName."
            return  
        }
    }
    catch {
        Write-Host "Error: $_"
    }

}

# Example:
# Break-Inheritance-List-Item -ListItemName "Annual Leave" -ListName "Leave Types"

#########################################################################################################

# Break-Inheritance-List Function
function Break-Inheritance-List {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ListName
    )

    try {

        $List = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue

        if ($null -eq $List) {
            Write-Host "List $ListName not found."
            return
        }    
        else {
            # Write-Host "List $ListName found."

            $List.BreakRoleInheritance($true, $true)
            $List.Update()
            Invoke-PnPQuery

            Write-Host "Inheritance broken successfully for list $ListName."
            return  
        }
    }
    catch {
        Write-Host "Error: $_"
    }

}

# Example:
# Break-Inheritance-List -ListName "Leave Types"

#########################################################################################################

# Get-List-Size Function
function Get-List-Size {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ListName
    )

    try {

        $List = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue

        if ($null -eq $List) {
            Write-Host "List $ListName not found."
            return
        }    
        else {
            # Write-Host "List $ListName found."

            $ListSize = $List.ItemCount

            Write-Host "List Size: $ListSize"
            return  
        }
    }
    catch {
        Write-Host "Error: $_"
    }

}

# Export the functions
Export-ModuleMember -Function *