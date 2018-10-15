# Sets test mode on
$varTest = $fal

# Remove any existing sessions
Get-PSSession | Remove-PSSession

# Connect to the local AD network, prefix all the commands with company to allow multiple commands of the same name. Must remember all the command become <verb>-company<noun>
# (E.g. Get-ADuser is for Get-ADUser)
Remove-Module ActiveDirectory
Import-Module ActiveDirectory

$objAllDomainUsers = Get-ADUser -Filter * -SearchBase "OU=Staff,DC=company,DC=com,DC=au"

$VarCompanyAddress1 = "@company.com.au"
$VarOffice365 = "@company.onmicrosoft.com"
$VarCompanyAddress2 = "@companyname2.com"

# Work around to input a colon character into a string.
$varColon = [char]58

ForEach ($varUser in $objAllDomainUsers) {
    
	# Creates the email addresses for all the users.
	$VarSIP = "SIP" + $varColon + $varUser.SamAccountName + "@company.com.au"
	$VarEmail = $varUser.SamAccountName + "@company.com.au"
	$VarCompanyAddress1 = "SMTP" + $varColon + $varUser.SamAccountName + "@company.com"
	$VarCompanyAddress2 = "smtp" + $varColon + $varUser.SamAccountName + "@companyname2.com"
	$VarOffice365 = "smtp" + $varColon + $varUser.SamAccountName + "@company.onmicrosoft.com"
	
	Write-Host -NoNewLine "Working on user '"
	Write-Host $varUser.Name
	
	if ($varTest -eq $true) {
		Write-Host -NoNewLine "Set-ADuser -Identity "
		Write-Host -NoNewLine $varUser
		Write-Host " -Clear proxyAddresses,mailNickname"
		
		Write-Host -NoNewLine "Set-ADuser -Identity "
		Write-Host -NoNewLine $varUser
		Write-Host -NoNewLine " -EmailAddress "
		Write-Host $VarEmail
		
		Write-Host -NoNewLine "Set-ADuser -Identity "
		Write-Host -NoNewLine $varUser
		Write-Host -NoNewLine " -Add @{proxyAddresses="
		Write-Host -NoNewLine $VarSIP
		Write-Host "}"
		
		Write-Host -NoNewLine "Set-ADuser -Identity "
		Write-Host -NoNewLine $varUser
		Write-Host -NoNewLine " -Add @{proxyAddresses="
		Write-Host -NoNewLine $VarCompanyAddress1
		Write-Host "}"
		
		Write-Host -NoNewLine "Set-ADuser -Identity "
		Write-Host -NoNewLine $varUser
		Write-Host -NoNewLine " -Add @{proxyAddresses="
		Write-Host -NoNewLine $VarCompanyAddress2
		Write-Host "}"
		
		Write-Host -NoNewLine "Set-ADuser -Identity "
		Write-Host -NoNewLine $varUser
		Write-Host -NoNewLine " -Add @{proxyAddresses="
		Write-Host -NoNewLine $VarOffice365
		Write-Host "}"
		
		Write-Host -NoNewLine "Set-ADuser -Identity "
		Write-Host -NoNewLine $varUser
		Write-Host -NoNewLine " -Add @{mailNickname="
		Write-Host -NoNewLine $varUser.SAMAccountName
		Write-Host "}"
	} else {
		Set-ADuser -Identity $varUser -Clear proxyAddresses,mailNickname
		Set-ADuser -Identity $varUser -EmailAddress $VarEmail
		Set-ADuser -Identity $varUser -Add @{proxyAddresses=$VarSIP}
		Set-ADuser -Identity $varUser -Add @{proxyAddresses=$VarCompanyAddress1}
		Set-ADuser -Identity $varUser -Add @{proxyAddresses=$VarCompanyAddress2}
		Set-ADuser -Identity $varUser -Add @{proxyAddresses=$VarOffice365}
		Set-ADuser -Identity $varUser -Add @{mailNickname=$varUser.SAMAccountName}
	}
}