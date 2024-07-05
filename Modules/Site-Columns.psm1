# Just Import as Module: Import-Module -Name .\Modules\Site-Columns.psm1
# Example for every Function are provided below

# Create-SiteColumn
function Create-SiteColumn {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [Parameter(Mandatory = $true)]
        [string]$DisplayName,
        [Parameter(Mandatory = $true)]
        [string]$Type,
        [string]$Group = "Custom Columns",
        [bool]$Required = $false
    )

    $siteColumn = Get-PnPField -Identity $Name -ErrorAction SilentlyContinue

    if ($null -eq $siteColumn) {
        Write-Host "Creating a new site column..."
        Add-PnPField -DisplayName $DisplayName -Type $Type -InternalName $Name -Group $Group -Required:$Required 
        Write-Host "Site Column $Name created successfully."
    }
    else {
        Write-Host "Site Column $Name already exists."
    }

    # Pause for 2 seconds
    Start-Sleep -Seconds 2
}

# Example:
# Create-SiteColumn -Name "TestSiteColumn" -DisplayName "Test Site Column" -Type "Text" -Group "Test Group" -Required $true

# Export all functions
Export-ModuleMember -Function *