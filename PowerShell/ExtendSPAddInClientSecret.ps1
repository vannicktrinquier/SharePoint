# Extend SPAddIn client secret
# The new client secret will be valid for a period of 3 years from now
# PS> .\ExtendSPAddInClientSecret.ps1 -clientId <CLIENT_ID>
# Done by Vannick Trinquier 
# On 13/02/2016


[CmdletBinding()]
Param(
   [Parameter(Mandatory=$True)]
   [string]$clientId
)

Connect-MsolService
$bytes = New-Object Byte[] 32
$rand = [System.Security.Cryptography.RandomNumberGenerator]::Create()
$rand.GetBytes($bytes)
$rand.Dispose()
$newClientSecret = [System.Convert]::ToBase64String($bytes)

$startDate= [System.DateTime]::Now
$endDate = $startDate.AddYears(3)
New-MsolServicePrincipalCredential -AppPrincipalId $clientId -Type Symmetric -Usage Sign -Value $newClientSecret -StartDate $startDate -EndDate $endDate
New-MsolServicePrincipalCredential -AppPrincipalId $clientId -Type Symmetric -Usage Verify -Value $newClientSecret -StartDate $startDate -EndDate $endDate 
New-MsolServicePrincipalCredential -AppPrincipalId $clientId -Type Password -Usage Verify -Value $newClientSecret -StartDate $startDate -EndDate $endDate
$newClientSecret