$cert = Get-AutomationCertificate -Name 'WindowsAutomationApp'

$AppID = Get-AutomationVariable -Name 'App- AutomationApplication-AppId'

 

Import-Module -Name ActiveDirectory

Import-Module MicrosoftTeams -RequiredVersion 4.9.3

Connect-MicrosoftTeams -Certificate $cert -ApplicationId $AppID -TenantId "tenantid"

 

$Names = Get-ADGroup sg-Licensing-TeamsPhoneStandard -Properties Member | Select-Object -ExpandProperty Member | Get-ADUser -Properties Name | Sort-Object

 

ForEach ($Name in $Names){

    $number = Get-CsPhoneNumberAssignment -AssignedPstnTargetId $Name.userPrincipalName | Select-Object -ExpandProperty TelephoneNumber

    $office = Get-ADUser $Name -Properties OfficePhone | Select-Object -ExpandProperty OfficePhone

    if ($office -eq $null) {

        Get-ADUser -Identity $Name | Set-ADUser -OfficePhone "$number"

    }

}

 

Disconnect-MicrosoftTeams
