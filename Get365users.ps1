###############################################################################
# This script retrieves flowing log from M365 thorough Microsoft Graph API.
# 1- 365 User Information
# 2- M365 User activity Report
# 3- M365 User Copilot Audit Log
###############################################################################
#This script based on the following script by Andres Bohren from GitHub:
#https://github.com/BohrenAn/GitHub_PowerShellScripts/blob/a442ef1ab0bd18ec7b96e76056058d0b9ebbb11f/ExchangeOnline/SearchAuditLog.ps1

###############################################################################
# 2025-04-16 - Yibo Yang initial version
# 
###############################################################################

#Begin Scpript


#Connect to Graph - Interactive
Connect-MgGraph -Scopes AuditLogsQuery.Read.All -NoWelcome

#Azure AD App > Icewolf
#$TenantId = "46bbad84-29f0-4e03-8d34-f6841a5071ad"
#$AppID = "99d8df8d-67b6-4a3a-b915-5cfc835fbfc7"
#$CertificateThumbprint = "07EFF3918F47995EB53B91848F69B5C0E78622FD" 
#Connect-MgGraph -ClientId $AppID -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint -NoWelcome

$Start = Get-Date

###############################################################################
# Create Search
###############################################################################
Write-Output "Create Array"
$OperationsArray = @()
$OperationsArray += "MailItemsAccessed"

Write-Output "Create Search"
$DisplayName = "DemoSearch_" + (Get-Date -Format "yyyyMMdd_HHmm")
[String]$StartDate = [datetime]::parseexact("2025-04-16", "yyyy-MM-dd", $null).Tostring("yyyy-MM-ddT00:00:00Z")
[String]$EndDate = [datetime]::parseexact("2025-04-18", "yyyy-MM-dd", $null).Tostring("yyyy-MM-ddT00:00:00Z")

$Uri = "https://graph.microsoft.com/beta/security/auditLog/queries"
$SearchParameters = @{
	displayName 		= "$DisplayName"
	filterStartDateTime = "$StartDate"
	filterEndDateTime 	= "$EndDate"
	recordTypeFilters 	= @("CopilotInteraction")
	#operationFilters	= @("MailItemsAccessed")
}

Write-Output "Invoke Search"
$SearchQuery = Invoke-MgGraphRequest -Method POST -Uri $Uri -Body $SearchParameters
$SearchId = $SearchQuery.Id
Write-Output "Searchid: $SearchId"

If ($SearchId -eq $null -or $SearchId -eq "")
{
    Write-Output "No SearchId > Aborting Script"
    #Exit
}

###############################################################################
# Check if SearchQuery Suceeded
###############################################################################
Write-Output "Wait for Search to complete"
#$AuditSearch = Get-MgBetaSecurityAuditLogQuery -AuditLogQueryId $SearchId | fl
#$AuditSearch = Get-MgBetaSecurityAuditLogQuery -AuditLogQueryId $SearchId
$URI = "https://graph.microsoft.com/beta/security/auditLog/queries/$searchId"
$AuditSearch = Invoke-MgGraphRequest -Method "GET" -Uri $Uri
$AuditSearchStatus = $AuditSearch.Status
Write-Output "Status: $AuditSearchStatus"
While ($AuditSearch.Status -ne "succeeded")
{
	#Status: notStarted, running, succeeded, failed
	$AuditSearch = Get-MgBetaSecurityAuditLogQuery -AuditLogQueryId $SearchId
    $AuditSearchStatus = $AuditSearch.Status
	Write-Output "Status: $AuditSearchStatus"
	Start-Sleep -Seconds 60

	If ($AuditSearchStatus -eq "failed")
	{
		Write-Output "Audit Search failed - aborting Script"
		Exit
	}
}

###############################################################################
# Get Data from SearchQuery
###############################################################################
Write-Output "Loop through results"
$Uri = ("https://graph.microsoft.com/beta/security/auditLog/queries/{0}/records" -f $SearchId)
[array]$SearchRecords = Invoke-MgGraphRequest -Uri $Uri -Method GET -OutputType json
$AuditRecords += $SearchRecords.value

# Paginate to fetch all available audit records
$NextLink = $SearchRecords.'@Odata.NextLink'
While ($null -ne $NextLink) {
    $SearchRecords = $null
    [array]$SearchRecords = Invoke-MgGraphRequest -Uri $NextLink -Method GET 
    $AuditRecords += $SearchRecords.value
    Write-Host ("{0} audit records fetched so far..." -f $AuditRecords.count)
    $NextLink = $SearchRecords.'@odata.NextLink' 
} 

$AuditRecordCount = $AuditRecords.Count
Write-Output "Audit Records found: $AuditRecordCount"


$End = Get-Date
$Timespan = New-Timespan -Start $Start -End $End
$Timespan