Connect-MsolService
$applist = Get-MsolServicePrincipal -all  |Where-Object -FilterScript { ($_.DisplayName -like "*Kulikat*") }
foreach ($appentry in $applist)
{
    $principalId = $appentry.AppPrincipalId
    $principalName = $appentry.DisplayName
    Write-Host "Name: $principalName"
	Write-Host "Client Id: $principalId"
    
    Get-MsolServicePrincipalCredential -AppPrincipalId $principalId -ReturnKeyValues $false | Where-Object { ($_.Type -ne "Other") -and ($_.Type -ne "Asymmetric") }
} 