#Import the Microsoft Hosted Services modules
Import-Module MSOnline
Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking;

# Set user credentials for O365
$username = "admin@company.com"
$password = Get-Content D:\Scripts\password.txt
$secstr = New-Object -TypeName System.Security.SecureString
$password.ToCharArray() | ForEach-Object {$secstr.AppendChar($_)}
$O365Cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $username, $secstr

# Prompt for user email address AND Out of Office message
$VAR_OoOUser = Read-Host "Please enter the user's email address "
$VAR_OoOMessage = Read-Host "Please enter the Out of Office message for the user "

#Create's a new session to the Microsoft Hosted Services
$O365Session = New-PSSession –ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $O365Cred -Authentication Basic -AllowRedirection 
Import-PSSession $O365Session -AllowClobber 

Connect-MsolService –Credential $O365Cred

Connect-SPOService -Url https://company-admin.sharepoint.com -Credential $O365Cred

Set-MailboxAutoReplyConfiguration $VAR_OoOUser -AutoReplyState enabled -ExternalAudience all -InternalMessage "$VAR_OoOMessage" -ExternalMessage "$VAR_OoOMessage"