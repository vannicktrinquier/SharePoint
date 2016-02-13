# Get information about SharePoint Add-Ins already deployed
# Among the information retrieve, we have the expiration date for SharePoint client secret
# PS> .\GetSPAddInsInformation.ps1 -addIn <APPNAME>
# Done by Vannick Trinquier 
# On 13/02/2016


[CmdletBinding()]
Param(
   [Parameter(Mandatory=$True)]
   [string]$addIn
)

Connect-MsolService
$applist = Get-MsolServicePrincipal -all  |Where-Object -FilterScript { ($_.DisplayName -like "*$addIn*") }
foreach ($appentry in $applist)
{
    $principalId = $appentry.AppPrincipalId
    $principalName = $appentry.DisplayName
    Write-Host "----------------------------------`n"
    Write-Host "Name: $principalName"
	Write-Host "Client Id: $principalId"
    
    Get-MsolServicePrincipalCredential -AppPrincipalId $principalId -ReturnKeyValues $false | Where-Object { ($_.Type -ne "Other") -and ($_.Type -ne "Asymmetric") }
} 