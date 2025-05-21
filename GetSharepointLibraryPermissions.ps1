# Parameters
$SiteUrl = "https://https://m365cpi78458766.sharepoint.com/sites/0-Site1-NoTeams-Nogroup"
$ReportOutput = "C:\NoUploadToCloud\LibraryPermissions.csv"
$LibraryName = "IC Documents"
 
#Connect to PnP Online
Connect-PnPOnline -Url $SiteUrl -Interactive # -Credentials (Get-Credential)
 
# Get the document library
$Library = Get-PnpList -Identity $LibraryName -Includes RoleAssignments
 
# Get all users and groups who has access
$RoleAssignments = $Library.RoleAssignments
$PermissionCollection = @()
Foreach ($RoleAssignment in $RoleAssignments)
{
    #Get the Permission Levels assigned and Member
    Get-PnPProperty -ClientObject $roleAssignment -Property RoleDefinitionBindings, Member
 
    #Get the Principal Type: User, SP Group, AD Group
    $PermissionType = $RoleAssignment.Member.PrincipalType
    $PermissionLevels = $RoleAssignment.RoleDefinitionBindings | Select -ExpandProperty Name
     
    #Get all permission levels assigned (Excluding:Limited Access)
    $PermissionLevels = ($PermissionLevels | Where { $_ -ne "Limited Access"}) -join ","
    If($PermissionLevels.Length -eq 0) {Continue}
 
    #Get SharePoint group members
    If($PermissionType -eq "SharePointGroup")
    {
        #Get Group Members
        $GroupMembers = Get-PnPGroupMember -Identity $RoleAssignment.Member.LoginName                 
        #Leave Empty Groups
        If($GroupMembers.count -eq 0){Continue}
 
        ForEach($User in $GroupMembers)
        {
            #Add the Data to Object
            $Permissions = New-Object PSObject
            $Permissions | Add-Member NoteProperty User($User.Title)
            $Permissions | Add-Member NoteProperty Type($PermissionType)
            $Permissions | Add-Member NoteProperty Permissions($PermissionLevels)
            $Permissions | Add-Member NoteProperty GrantedThrough("SharePoint Group: $($RoleAssignment.Member.LoginName)")
            $PermissionCollection += $Permissions
        }
    }
    Else
    {
        #Add the Data to Object
        $Permissions = New-Object PSObject
        $Permissions | Add-Member NoteProperty User($RoleAssignment.Member.Title)
        $Permissions | Add-Member NoteProperty Type($PermissionType)
        $Permissions | Add-Member NoteProperty Permissions($PermissionLevels)
        $Permissions | Add-Member NoteProperty GrantedThrough("Direct Permissions")
        $PermissionCollection += $Permissions
    }
}
#Export Permissions to CSV File
$PermissionCollection
$PermissionCollection | Export-CSV $ReportOutput -NoTypeInformation
Write-host -f Green "Permission Report Generated Successfully!"
