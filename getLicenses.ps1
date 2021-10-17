Import-Module AzureAD

$Credentials = Get-Credential
Connect-AzureAD -Credential $Credentials

Get-MsolUser -UserPrincipalName test.schueler@sbs-herzogenaurach.de | Format-List DisplayName,Licenses | Out-File -FilePath .\licenseList.txt

Get-AzureADSubscribedSku | Select SkuPartNumber |  Add-Content -Path .\licenseList.txt