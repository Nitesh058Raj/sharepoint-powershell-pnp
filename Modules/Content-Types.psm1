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
function Add-SiteColumnToContentType {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SiteColumn,
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

# Export all functions
Export-ModuleMember -Function *