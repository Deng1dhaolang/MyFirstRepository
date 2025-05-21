# Import SharePoint Online Management Shell module
Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking

# SharePoint site URL and credentials
$siteUrl = "https://yourtenant.sharepoint.com/sites/yoursite"
Connect-PnPOnline -Url $siteUrl -Interactive

# Get all document libraries
$libraries = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 }

foreach ($library in $libraries) {
    Write-Host "`nPermissions for library: $($library.Title)" -ForegroundColor Green
    
    # Get role assignments for the library
    $roleAssignments = Get-PnPRoleAssignment -List $library.Title
    
    foreach ($role in $roleAssignments) {
        Write-Host "User/Group: $($role.Member.Title)"
        Write-Host "Permission Level: $($role.RoleDefinitionBindings.Name)"
        Write-Host "-------------------"
    }
}

# Disconnect from SharePoint
Disconnect-PnPOnline