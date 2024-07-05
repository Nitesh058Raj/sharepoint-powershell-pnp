# Just Import as Module: Import-Module -Name .\Modules\Permissions_Groups.psm1
# Example for every Function are provided below

# Create-Group Function
function Create-Group {
    param (
        [Parameter(Mandatory = $true)]
        [string]$GroupName,
        [string]$Description = "Group created by PowerShell"
    )

    try {

        $group = Get-PnPGroup -Identity $GroupName -ErrorAction SilentlyContinue

        if ($null -eq $group) {
            Write-Host "Creating a new group..."
            New-PnPGroup -Title $GroupName -Description $Description
            Write-Host "Group $GroupName created successfully."
        }
        else {
            Write-Host "Group $GroupName already exists."
        }
    }
    catch {
        Write-Host "Error: $_"
    }
}

# Example:
# Create-Group -GroupName "Test Group" -Description "Test Group Description"

#########################################################################################################

# Create-Group-And-Add-Members Function
function Create-Group-And-Add-Members {
    param (
        [Parameter(Mandatory = $true)]
        [string]$GroupName,
        [string]$Description = "Group created by PowerShell",
        [string]$Owner = "",
        [string]$Members = ""
    )

    try{

        $group = Get-PnPGroup -Identity $GroupName -ErrorAction SilentlyContinue

        if ($null -eq $group) {
            Write-Host "Creating a new group..."
            New-PnPGroup -Title $GroupName -Description $Description -Owner $Owner
            Write-Host "Group $GroupName created successfully."
        }
        else {
            Write-Host "Group $GroupName already exists."
        }

        if($Members -ne "") {
            Add-MembersToGroup -GroupName $GroupName -Members $Members
        }

    } catch {
        Write-Host "Error: $_"
    }
}

# Example:
# Create-Group-And-Add-Members -GroupName "Test Group" -Description "Test Group Description" -Owner "

#########################################################################################################

# Add-MembersToGroup Function
function Add-MembersToGroup {
    param (
        [Parameter(Mandatory = $true)]
        [string]$GroupName,
        [string]$Members
    )

    try {

        $group = Get-PnPGroup -Identity $GroupName -ErrorAction SilentlyContinue

        if ($null -ne $group) {
            $members = $Members -split ","

            foreach ($member in $members) {
                Add-PnPUserToGroup -LoginName $member -Identity $GroupName
                Write-Host "User $member added to group $GroupName."
            }
        }
        else {
            Write-Host "Group $GroupName does not exist."
        }

    }
    catch {
        Write-Host "Error: $_"
    }
}
# Example:
# Create-Group -GroupName "Test Group" -Description "Test Group Description" -Owner "112@1m.msw" -Members "23@23n.ra, 98@12l.ra"

#########################################################################################################

# Create-PermissionLevel Function
function Create-PermissionLevel {
    param (
        [Parameter(Mandatory = $true)]
        [string]$PermissionLevelName,
        [string]$Description = "Permission Level created by PowerShell",
        [string]$Clone = "",
        [string]$BasePermissions = "",
        [string]$ExcludePermission = ""
    )

    try {

        $roleDefinition = Get-PnPRoleDefinition -Identity $PermissionLevelName -ErrorAction SilentlyContinue

        if ($null -eq $roleDefinition) {
            Write-Host "Creating a new permission level..."
            
            if ($Clone -ne "") {
                Add-PnPRoleDefinition -RoleName $PermissionLevelName -Description $Description
            } elseif ($ExcludePermission -ne "" -and $BasePermissions -ne "") {
                Add-PnPRoleDefinition -RoleName $PermissionLevelName -Description $Description -Clone $Clone 
            } elseif ($ExcludePermission -ne "") {
                $BasePermissionsArray = $BasePermissions -split ","
                Add-PnPRoleDefinition -RoleName $PermissionLevelName -Description $Description -Clone $Clone -Include $BasePermissionsArray
            } elseif ($BasePermissions -ne "") {

                $ExcludePermissionArray = $ExcludePermission -split ","
                Add-PnPRoleDefinition -RoleName $PermissionLevelName -Description $Description -Clone $Clone -Exclude $ExcludePermissionArray
            } else {
                $BasePermissionsArray = $BasePermissions -split ","
                $ExcludePermissionArray = $ExcludePermission -split ","
                Add-PnPRoleDefinition -RoleName $PermissionLevelName -Description $Description -Clone $Clone -Include $BasePermissionsArray -Exclude $ExcludePermissionArray
            }
            
            Write-Host "Permission Level $PermissionLevelName created successfully."
        }
    else {
        Write-Host "Permission Level $PermissionLevelName already exists."
    }

    } catch {
        Write-Host "Error: $_"
    }

}

# Example:
# Create-PermissionLevel -PermissionLevelName "TestCustomPermissionLevel" -Description "Test Custom Permission Level" -BasePermissions "EditListItems,DeleteListItems" -Clone "Contribute" -ExcludePermission "ManagePermissions"

#########################################################################################################

# Set-PermissionLevel Function
function Set-PermissionLevel {
    param (
        [Parameter(Mandatory = $true)]
        [string]$GroupName,
        [string]$PermissionLevel
    )

    try {
        $group = Get-PnPGroup -Identity $GroupName -ErrorAction SilentlyContinue

        if ($null -ne $group) {
            $roleDefinition = Get-PnPRoleDefinition -Identity $PermissionLevel -ErrorAction SilentlyContinue

            if ($null -ne $roleDefinition) {
                Set-PnPGroupPermissions -Identity $GroupName -AddRole $PermissionLevel
                Write-Host "Permission Level $PermissionLevel set for group $GroupName."
            }
            else {
                Write-Host "Permission Level $PermissionLevel does not exist."
            }
        }
        else {
            Write-Host "Group $GroupName does not exist."
        }
    } catch {
        Write-Host "Error: $_"
    }
    
}

# Example: 
# Set-PermissionLevel -GroupName "Test Group" -PermissionLevel "Contribute"

#########################################################################################################

# Remove-Group Function
function Remove-Group {
    param (
        [Parameter(Mandatory = $true)]
        [string]$GroupName
    )

    try {
        $group = Get-PnPGroup -Identity $GroupName -ErrorAction SilentlyContinue

        if ($null -ne $group) {
            Remove-PnPGroup -Identity $GroupName -Force
            Write-Host "Group $GroupName removed successfully."
        }
        else {
            Write-Host "Group $GroupName does not exist."
        }
    } catch {
        Write-Host "Error: $_"
    }

}

# Example:
# Remove-Group -GroupName "Test Group"

#########################################################################################################

# Remove-MembersFromGroup Function
function Remove-MembersFromGroup {
    param (
        [Parameter(Mandatory = $true)]
        [string]$GroupName,
        [string]$Members
    )

    try {

        $group = Get-PnPGroup -Identity $GroupName -ErrorAction SilentlyContinue

        if ($null -ne $group) {
            $members = $Members -split ","

            foreach ($member in $members) {
                Remove-PnPUserFromGroup -LoginName $member -Identity $GroupName
                Write-Host "User $member removed from group $GroupName."
            }
        }
        else {
            Write-Host "Group $GroupName does not exist."
        }

    } catch {
        Write-Host "Error: $_"
    }
}

# Example:
# Remove-MembersFromGroup -GroupName "Test Group" -Members "abc@ii.ss"

#########################################################################################################

# Remove-PermissionLevel Function
function Remove-PermissionLevel {
    param (
        [Parameter(Mandatory = $true)]
        [string]$PermissionLevelName
    )

    try {
        $roleDefinition = Get-PnPRoleDefinition -Identity $PermissionLevelName -ErrorAction SilentlyContinue

        if ($null -ne $roleDefinition) {
            Remove-PnPRoleDefinition -Identity $PermissionLevelName -Force
            Write-Host "Permission Level $PermissionLevelName removed successfully."
        }
        else {
            Write-Host "Permission Level $PermissionLevelName does not exist."
        }
    
    } catch {
        Write-Host "Error: $_"
    }
}

# Example:
# Remove-PermissionLevel -PermissionLevelName "TestCustomPermissionLevel"

#########################################################################################################

# Export all functions
Export-ModuleMember -Function *

