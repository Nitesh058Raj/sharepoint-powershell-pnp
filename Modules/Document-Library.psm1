# Just Import as Module: Import-Module -Name .\Modules\Document-Library.psm1
# Example for every Function are provided below

# Create-DocumentLibrary Function
function Create-DocumentLibrary {
    param (
        [Parameter(Mandatory = $true)]
        [string]$LibraryName,
        [string]$TemplateType = "DocumentLibrary"
    )

    $library = Get-PnPList -Identity $LibraryName -ErrorAction SilentlyContinue

    if ($null -eq $library) {
        Write-Host "Creating a new document library..."
        New-PnPList -Title $LibraryName -Template $TemplateType -OnQuickLaunch 
        Write-Host "Document Library $LibraryName created successfully."
    }
    else {
        Write-Host "Document Library $LibraryName already exists."
    }

}

# Example:
# Create-DocumentLibrary -LibraryName "Test Library" -TemplateType "DocumentLibrary"

#########################################################################################################

# Break-Inheritance-File Function
function Break-Inheritance-File {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FileName,
        [Parameter(Mandatory = $true)]
        [string]$LibraryName,
        [Parameter(Mandatory = $true)]
        [string]$SiteName
    )


    $FileUrl = "/sites/$SiteName/$LibraryName/$FileName"

    try {
        $file = Get-PnPFile  -Url $FileUrl -ErrorAction SilentlyContinue

        if ($null -eq $file) {
            Write-Host "File $FileName not found in library $LibraryName."
            return
        }    
        else {
            # Write-Host "File $FileName found in library $LibraryName."

            $item = $file.ListItemAllFields
            $item.BreakRoleInheritance($true, $true)
            $item.Update()
            Invoke-PnPQuery

            Write-Host "Inheritance broken successfully for file $FileName in library $LibraryName."
            return  
        }
    }
    catch {
        Write-Host "Error: $_"
    }
}

# Example:
# Break-Inheritance-File -SiteUrl "your_site_name" -FileName "Random_Data.xlsx" -LibraryName "Test Library"

#########################################################################################################

# Break-Inheritance-Folder Function
function Break-Inheritance-Folder {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FolderName,
        [Parameter(Mandatory = $true)]
        [string]$LibraryName,
        [Parameter(Mandatory = $true)]
        [string]$SiteName
    )

    $FolderUrl = "/sites/$SiteName/$LibraryName/$FolderName"

    try {
        $folder = Get-PnPFolder -Url $FolderUrl -ErrorAction SilentlyContinue

        if ($null -eq $folder) {
            Write-Host "Folder $FolderName not found in library $LibraryName."
            return
        }    
        else {
            # Write-Host "Folder $FolderName found in library $LibraryName."

            $item = $folder.ListItemAllFields
            $item.BreakRoleInheritance($true, $true)
            $item.Update()
            Invoke-PnPQuery

            Write-Host "Inheritance broken successfully for folder $FolderName in library $LibraryName."
            return  
        }
    }
    catch {
        Write-Host "Error: $_"
    }
}

# Example:
# Break-Inheritance-Folder -SiteUrl "your_site_name" -FolderName "Test Folder" -LibraryName "Test Library"

#########################################################################################################

# Break-Inheritance-DocSet Function
function Break-Inheritance-DocSet {
    param (
        [Parameter(Mandatory = $true)]
        [string]$DocSetName,
        [Parameter(Mandatory = $true)]
        [string]$LibraryName,
        [Parameter(Mandatory = $true)]
        [string]$SiteName
    )

    $DocSetUrl = "/sites/$SiteName/$LibraryName"

    try {

        $query = "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$DocSetName</Value></Eq></Where></Query></View>"
        $DocSetDetails = Get-PnPListItem -List $LibraryName -Query $query -ErrorAction SilentlyContinue
        $docSet = Get-PnPListItem -List $LibraryName -Id $DocSetDetails.Id -ErrorAction SilentlyContinue
        
        # All the details of the document set as a table
        # $DocSetDetails | Format-Table
        # $docSet | Format-Table

        if ($null -eq $docSet) {
            Write-Host "Document Set $DocSetName not found in library $LibraryName."
            return
        }    
        else {
            # Write-Host "Document Set $DocSetName found in library $LibraryName."

            $docSet.BreakRoleInheritance($true, $true)
            $docSet.Update()
            Invoke-PnPQuery

            Write-Host "Inheritance broken successfully for document set $DocSetName in library $LibraryName."
            return  
        }
    }
    catch {
        Write-Host "Error: $_"
    }
}

# Example:
# Break-Inheritance-DocSet -SiteUrl "your_site_name" -DocSetName "Test_DocumentSet" -LibraryName "Test Library"

#########################################################################################################

# Break-Inheritance-Library Function
function Break-Inheritance-Library {
    param (
        [Parameter(Mandatory = $true)]
        [string]$LibraryName,
        [Parameter(Mandatory = $true)]
        [string]$SiteName
    )

    $LibraryUrl = "/sites/$SiteName/$LibraryName"

    try {
        $library = Get-PnPList -Identity $LibraryName -ErrorAction SilentlyContinue

        if ($null -eq $library) {
            Write-Host "Library $LibraryName not found."
            return
        }    
        else {
            # Write-Host "Library $LibraryName found."

            $library.BreakRoleInheritance($true, $true)
            $library.Update()
            Invoke-PnPQuery

            Write-Host "Inheritance broken successfully for library $LibraryName."
            return  
        }
    }
    catch {
        Write-Host "Error: $_"
    }
}

# Example:
# Break-Inheritance-Library -SiteUrl "your_site_name" -LibraryName "Test Library"

#########################################################################################################

# Export all functions
Export-ModuleMember -Function *