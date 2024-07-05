# Just Import as Module: Import-Module -Name .\Modules\Content-Types.psm1
# Example for every Function are provided below

# Create-ContentType
function Create-ContentType {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [Parameter(Mandatory = $true)]
        [string]$Description,
        [string]$Group = "Custom Content Types"
    )

    $contentType = Get-PnPContentType -Identity $Name -ErrorAction SilentlyContinue

    if ($null -eq $contentType) {
        Write-Host "Creating a new content type..."
        Add-PnPContentType -Name $Name -Description $Description -Group $Group
        Write-Host "Content Type $Name created successfully."
    }
    else {
        Write-Host "Content Type $Name already exists."
    }

    # Pause for 2 seconds
    Start-Sleep -Seconds 2
}

# Example:
# Create-ContentType -Name "TestContentType" -Description "Test Content Type" -Group "Test Group"

# Add-SiteColumnToContentType
# Use the Add-PnPFieldToContentType cmdlet to add a site column to a content type, because for some reason the Add-PnPFieldToContentType cmdlet is not working as expected in this function.
function Add-SiteColumnToContentType {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SiteColumn, # Internal Name of the Site Column 
        [Parameter(Mandatory = $true)]
        [string]$ContentType
    )

    $siteColumn = Get-PnPField -Identity $SiteColumn -ErrorAction SilentlyContinue
    $contentType = Get-PnPContentType -Identity $ContentType -ErrorAction SilentlyContinue

    if ($null -ne $siteColumn -and $null -ne $contentType) {
        Write-Host "Adding site column $SiteColumn to content type $ContentType..."
        Add-PnPFieldToContentType -Field $siteColumn -ContentType $contentType
        Write-Host "Site Column $SiteColumn added to Content Type $ContentType successfully."
    }
    else {
        Write-Host "Site Column $SiteColumn or Content Type $ContentType does not exist."
    }

    # Pause for 2 seconds
    Start-Sleep -Seconds 2
}

# Example:
# Create-ContentType -Name "TestContentType" -Description "Test Content Type" -Group "Test Group"
# Use below command to add site column to content type instead of Add-SiteColumnToContentType function
# Add-SiteColumnToContentType -SiteColumn "Test_Site_Column_v3" -ContentType "TestContentType"

# Export all functions
Export-ModuleMember -Function *