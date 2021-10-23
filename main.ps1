Import-Module AzureAD

$Credentials = Get-Credential
Connect-AzureAD -Credential $Credentials

$path = $PSScriptRoot+"\addUsers.ps1"
. "$path"

$path = $PSScriptRoot+"\addUsersToGroups.ps1"
. "$path"
