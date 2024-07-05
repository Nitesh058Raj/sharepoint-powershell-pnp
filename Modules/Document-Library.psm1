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

# Export all functions
Export-ModuleMember -Function *