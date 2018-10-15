# Created by Nathaniel Mitchell

# **** Quick Doco
# You MUST run this command for each user that this script uses so it can
# make a remote powershell connection to each exchange server
#
# Set-User <USERNAME> -RemotePowerShellEnabled $True
#
# You MUST run this command to enable the user to export user mailboxes
#
# New-ManagementRoleAssignment –Role "Mailbox Import Export" –User <USERNAME> 

# *******************************************************
# *******************************************************
# Set some variables

$VAR_CONFIG_REDIRECT = $True
$VAR_CONFIG_REDIRECT_KEEP = $False
$VAR_CONFIG_DISABLE_OLD = $True
$VAR_CONFIG_CHANGE_PASSWORD = $True
$VAR_CONFIG_COPY_FILES = $True
$VAR_CONFIG_EXPORT_MAIL = $True
$VAR_CONFIG_IMPORT_MAIL = $True
$VAR_CONFIG_PASSWORD_OLD = "Once Upon a Night - Part 1"
$VAR_CONFIG_PAUSE = $false

$VAR_SCRIPT_SOURCE = "C:\Scripts\Domain Migration"

# Below are the variables for forced domain replication. From local DC to DC near OCS server
# DSACLS must be located in the search path
$VAR_REP_OCS_DC = "DC.NEW-DOMAIN.LOCAL"
$VAR_REP_DOMAIN_PART = "DC=new-domain,DC=local"
$VAR_REP_DOMAIN_GUID = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
$VAR_REP_CONFIG_PART = "CN=Configuration,DC=new-domain,DC=local"
$VAR_REP_CONFIG_GUID = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
$VAR_REP_SCHEMA_PART = "CN=Schema,CN=Configuration,DC=new-domain,DC=local"
$VAR_REP_SCHEMA_GUID = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
$VAR_REP_FRSDNS_PART = "DC=ForestDnsZones,DC=new-domain,DC=local"
$VAR_REP_FRSDNS_GUID = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
$VAR_REP_DOMDNS_PART = "DC=DomainDnsZones,DC=new-domain,DC=local"
$VAR_REP_DOMDNS_GUID = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

# These source / destination locations MUST be a network share and accessible by both domain users.
$VAR_DESTINATION_PST = "\\FS.NEW-DOMAIN.LOCAL\psts$"
$VAR_SOURCE_PST = "\\FS.NEW-DOMAIN.LOCAL\psts$\blank.pst"

# Settings for DOMAIN1
$VAR_Domain1 = "OLD-DOMAIN.LOCAL"
$VAR_Domain1_OU = "OU=Batch 1,OU=Exporting Users,OU=Company Users"
#$VAR_Domain1_OU = "OU=Company Users"
$VAR_Domain1_Username = "OLD-DOMAIN\admin"
$VAR_Domain1_PASSWORD = "password"
$VAR_Domain1_Server = "DC.OLD-DOMAIN.LOCAL"
$VAR_Domain1_HOMEPATH = "\\FS.OLD-DOMAIN.LOCAL\Users"
$VAR_Domain1_Exchange_PS = "http://EXCHANGE.OLD-DOMAIN.LOCAL/PowerShell/"
$VAR_Domain1_Contact_OU = "Mail Contacts"

# Settings for DOMAIN2
$VAR_Domain2 = "NEW-DOMAIN.LOCAL"
$VAR_Domain2_OU = "OU=Imported Users,OU=Users,OU=Enterprise Tree"
$VAR_Domain2_Username = "NEW-DOMAIN\admin"
$VAR_Domain2_PASSWORD = "password"
$VAR_Domain2_Server = "DC.NEW-DOMAIN.LOCAL"
$VAR_Domain2_Exchange_PS = "http://EXCHANGE.NEW-DOMAIN.LOCAL/PowerShell/"
$VAR_Domain2_Exchange_Suffix = "@new-domain.net.au"
# *** Settings for OCS 2007 R2
$VAR_Domain2_POOL = "DOMAIN-CS"
# If Office Communicator 2007 R2 has been installed as a 'system' container
$VAR_Domain2_HomeServerDN = "CN=LC Services,CN=Microsoft,CN=DOMAIN-CS,CN=Pools,CN=RTC Service,CN=Microsoft,CN=System,DC=new-domain,DC=local"
# If Office Communicator 2007 R2 has been installed as a 'configuration' container
#$VAR_Domain2_HomeServerDN = "CN=LC Services,CN=Microsoft,CN=DOMAIN-CS,CN=Pools,CN=RTC Service,CN=Microsoft,CN=Configuration,DC=new-domain,DC=local"

# Sets some variables that are going to be used on all new user accounts.
$City = "xxxx"
$Company = "xxxx"
$Country = "AU"
$Department = "xxxx"
$HomePage = "xxxx"
$fax = "xxxx"
$Office = "xxxx"
$PostalCode = "xxxx"
$State = "xxxx"
$StreetAddress = "xxxx"
$HOMEDRIVE = "H:"
$HOMEPATH = "\\FS.NEW-DOMAIN.LOCAL\HOME$"
$LOGINPATH = "login.bat"
#$Password = "This is a Really Long & Complex Password. YOU SHOULDN'T BE USING THIS ACCOUNT!"
$Password = "xxxxxxxxxxxx"

# *******************************************************
# *******************************************************
# **                                                   **
# **            DO NOT EDIT BELOW THIS LINE            **
# **                                                   **
# *******************************************************
# *******************************************************

# *******************************************************
# *******************************************************
# Quick clean up
NET USE J: /DELETE
NET USE K: /DELETE
Remove-PSSession *

# *******************************************************
# *******************************************************
# Set the scripts colours

$VAR_OBJ_CONSOLE = (Get-Host).UI.RawUI
$VAR_OBJ_CONSOLE.BackgroundColor = "black"
$VAR_OBJ_CONSOLE.ForegroundColor = "white"
$VAR_OBJ_CONSOLE.windowtitle = "User Migration - Script 1"

$VAR_OBJ_CON_NEW_BUFFER = $VAR_OBJ_CONSOLE.buffersize
$VAR_OBJ_CON_NEW_BUFFER.height = 3000
$VAR_OBJ_CON_NEW_BUFFER.width = 160
$VAR_OBJ_CONSOLE.buffersize = $VAR_OBJ_CON_NEW_BUFFER

$VAR_OBJ_CON_NEW_WINSIZE = $VAR_OBJ_CONSOLE.windowsize
$VAR_OBJ_CON_NEW_WINSIZE.height = 50
$VAR_OBJ_CON_NEW_WINSIZE.width = 160
$VAR_OBJ_CONSOLE.windowsize = $VAR_OBJ_CON_NEW_WINSIZE

Clear-Host

# *******************************************************
# *******************************************************
# Import required modules
#[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null 
Import-Module ActiveDirectory

# *******************************************************
# *******************************************************
# Create objects for later use
$VAR_OBJ_CURRENT_PROCCESS = [System.Diagnostics.Process]::GetCurrentProcess()

# *******************************************************
# *******************************************************
# Converts the passwords above into a 'secured string' (as required for storing passwords in PowerShell scripts)
$New_Password = $Password | ConvertTo-SecureString -AsPlainText -Force
$OLD_Password = $VAR_CONFIG_PASSWORD_OLD | ConvertTo-SecureString -AsPlainText -Force
$VAR_DOMAIN1_PASSWORD_SECURED = $VAR_Domain1_PASSWORD | ConvertTo-SecureString -AsPlainText -Force 
$VAR_DOMAIN2_PASSWORD_SECURED = $VAR_Domain2_PASSWORD | ConvertTo-SecureString -AsPlainText -Force 

# *******************************************************
# *******************************************************
# Creates an object for each domain that has the username and password pre-typed when password authentication requests come up
$VAR_DOMAIN1_CREDENTIALS = New-Object System.Management.Automation.PSCredential -ArgumentList $VAR_Domain1_Username, $VAR_DOMAIN1_PASSWORD_SECURED
$VAR_DOMAIN2_CREDENTIALS = New-Object System.Management.Automation.PSCredential -ArgumentList $VAR_Domain2_Username, $VAR_DOMAIN2_PASSWORD_SECURED

# *******************************************************
# *******************************************************
# Changes the above variables into usable variables
$VAR_DOMAIN1_PARTS = [ARRAY]$VAR_Domain1.split(".")

# Buils a variable that has the domain re-structured to be in LDAP format (IE DC=domainname,DC=domainname)
# uses a simple counter to make sure we dont have a comma to the end of the new variable
$VAR_TMP_COUNT1 = 0
foreach ($VAR_TMP_I in $VAR_DOMAIN1_PARTS) {
     $VAR_DOMAIN1_FQDN = $VAR_DOMAIN1_FQDN + "DC=" + $VAR_TMP_I
     $VAR_TMP_COUNT1 = $VAR_TMP_COUNT1 + 1
     IF ($VAR_TMP_COUNT1 -lt $VAR_DOMAIN1_PARTS.length) {$VAR_DOMAIN1_FQDN = $VAR_DOMAIN1_FQDN + ","}
}

# Repeats above for domain #2
$VAR_DOMAIN2_PARTS = [ARRAY]$VAR_Domain2.split(".")
$VAR_TMP_COUNT2 = 0
foreach ($VAR_TMP_J in $VAR_DOMAIN2_PARTS) {
     $VAR_DOMAIN2_FQDN = $VAR_DOMAIN2_FQDN + "DC=" + $VAR_TMP_J
     $VAR_TMP_COUNT2 = $VAR_TMP_COUNT2 + 1
     IF ($VAR_TMP_COUNT2 -lt $VAR_DOMAIN2_PARTS.length) {$VAR_DOMAIN2_FQDN = $VAR_DOMAIN2_FQDN + ","}
}

Write-Host " "
Write-Host "Creating mapped drives to the AD domain"
# Create temporary drives to the AD domains to increase network speed and authentication
New-PSDrive -Name DOMAIN1 -PSProvider ActiveDirectory -Root "" -Server $VAR_Domain1_Server -Credential $VAR_DOMAIN1_CREDENTIALS | Out-Null 
New-PSDrive -Name DOMAIN2 -PSProvider ActiveDirectory -Root "" -Server $VAR_Domain2_Server -Credential $VAR_DOMAIN2_CREDENTIALS | Out-Null

# Create links to connect to remote Exchange servers
$VAR_MX_Session_Domain1 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $VAR_Domain1_Exchange_PS -Authentication Kerberos -Credential $VAR_DOMAIN1_CREDENTIALS
$VAR_MX_Session_Domain2 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $VAR_Domain2_Exchange_PS -Authentication Kerberos -Credential $VAR_DOMAIN2_CREDENTIALS

Write-Host " "
Write-Host "Creating mapped drives to the new and old FS"
# Create temporary drives to the new server for the new home drive location
New-PSDrive -Name NEW_FS -PSProvider FileSystem -Root $HOMEPATH | Out-Null
$VAR_NET_DRIVE_HME = new-object -ComObject WScript.Network
$VAR_NET_DRIVE_HME.MapNetworkDrive("J:", "$VAR_Domain1_HOMEPATH", $false, "$VAR_Domain1_Username", "$VAR_Domain1_PASSWORD")
$VAR_NET_DRIVE_PST = new-object -ComObject WScript.Network
$VAR_NET_DRIVE_PST.MapNetworkDrive("K:", "$VAR_DESTINATION_PST", $false, "$VAR_Domain2_Username", "$VAR_Domain2_PASSWORD")

# Moves into the new drive for DOMAIN1
CD DOMAIN1:

# Get Domain 1's list of users to migrate
$VAR_DOMAIN1_EXPORT_USERS = Get-ADUser -Filter * -SearchBase "$VAR_Domain1_OU,$VAR_DOMAIN1_FQDN" -Properties *

If ($VAR_DOMAIN1_EXPORT_USERS.length -lt 2) {
     Write-Host " "
     Write-Host -NoNewLine "There is 1 user to mirgrate"
} Else {
     Write-Host " "
     Write-Host -NoNewLine "There are "
     Write-Host -NoNewLine $VAR_DOMAIN1_EXPORT_USERS.length
     Write-Host " users to mirgrate"
}

foreach ($VAR_TMP_DOMAIN1_USER in $VAR_DOMAIN1_EXPORT_USERS) {
     CD DOMAIN1:
     $VAR_USER_CREATION = $true
     
     Write-Host " "
     Write-Host -foregroundcolor "blue" "********************"
     Write-Host -foregroundcolor "blue" "********************"
     Write-Host -foregroundcolor "blue" "********************"
     Write-Host " "
     Write-Host -NoNewLine "The current user is "
     Write-Host -foregroundcolor "magenta" $VAR_TMP_DOMAIN1_USER
     Write-Host -NoNewLine "Their display name is "
     Write-Host -foregroundcolor "magenta" $VAR_TMP_DOMAIN1_USER.DisplayName
     Write-Host -NoNewLine "Their given name is "
     Write-Host -foregroundcolor "magenta" $VAR_TMP_DOMAIN1_USER.GivenName
     Write-Host -NoNewLine "Their last name is "
     Write-Host -foregroundcolor "magenta" $VAR_TMP_DOMAIN1_USER.Surname
     Write-Host " "
     Write-Host -NoNewLine "Their OLD homedrive is "
     Write-Host -foregroundcolor "magenta" $VAR_TMP_DOMAIN1_USER.HomeDirectory
     
     # *******************************************************
     # *******************************************************
     # Creates some variables to be used for user creation and moving
     
     # Generate the new username based on 
     #$NEW_USERNAME_TMP = $VAR_TMP_DOMAIN1_USER.GivenName.Substring(0,1) + $VAR_TMP_DOMAIN1_USER.Surname.Replace(" ","")
     $NEW_USERNAME_TMP = $VAR_TMP_DOMAIN1_USER.GivenName.Replace(" ","") + "." + $VAR_TMP_DOMAIN1_USER.Surname.Replace(" ","")
     if ($NEW_USERNAME_TMP.length -gt 20) {
          $NEW_USERNAME = $NEW_USERNAME_TMP.Substring(0,20)
     } ELSE {
          $NEW_USERNAME = $NEW_USERNAME_TMP
     }
     
     # Because the new user is created in the "User" folder at the root of the AD tree. This folder is a container (CN) *NOT* an organisation unit (OU)
     $NEW_TEMP_IDENTITY = "CN=" + $VAR_TMP_DOMAIN1_USER.DisplayName + ",CN=Users," + $VAR_DOMAIN2_FQDN
     
     $VAR_Export_location_PST = $NEW_USERNAME + ".pst"
     $VAR_Export_location = $VAR_DESTINATION_PST + "\" + $VAR_Export_location_PST
     $NEW_UPN = $NEW_USERNAME + "@" + $VAR_Domain2
     $NEW_HOMEPATH = $HOMEPATH + "\" + $NEW_USERNAME
     $NEW_FQ_USERNAME = $VAR_Domain2 + "\" + $NEW_USERNAME
     $OLD_HOMEPATH = "J:\" + $VAR_TMP_DOMAIN1_USER.sAMAccountName
     $VAR_DOMAIN1_USER_GROUPS=((Get-ADUser $VAR_TMP_DOMAIN1_USER.sAMAccountName -Properties *).MemberOf -split (",") | Select-String -SimpleMatch "CN=") -replace "cn=",""
     
     Write-Host -NoNewLine "Their OLD username is "
     Write-Host -foregroundcolor "magenta" $VAR_TMP_DOMAIN1_USER.sAMAccountName
     Write-Host " "
     Write-Host -NoNewLine "Their NEW username is "
     Write-Host -foregroundcolor "magenta" $NEW_USERNAME
     Write-Host -NoNewLine "Their NEW UPN is "
     Write-Host -foregroundcolor "magenta" $NEW_UPN
     Write-Host -NoNewLine "Their NEW home drive is "
     Write-Host -foregroundcolor "magenta" $NEW_HOMEPATH
     
     # Moves into the new drive for DOMAIN2 to do work
     CD DOMAIN2:
     
     $VAR_TEMP_ERROR_SETTING = $ErrorActionPreference
     $ErrorActionPreference = "SilentlyContinue"
     $VAR_OBJ_SEARCH = Get-ADUser -Identity $NEW_USERNAME
     $ErrorActionPreference = $VAR_TEMP_ERROR_SETTING
     If ($VAR_OBJ_SEARCH -eq $Null) {
          $VAR_USER_CREATION = $True
     } Else {
          $VAR_USER_CREATION = $False
          Write-Host " "
          Write-Host  -NoNewLine -foregroundcolor "darkred" -backgroundcolor "darkgray" "                                         "
          for ($VAR_TEMP_z = 0; $VAR_TEMP_z -lt ($NEW_USERNAME.length - 1); $VAR_TEMP_z++) {
               Write-Host -NoNewLine -foregroundcolor "red" -backgroundcolor "darkgray" " "
          }
          Write-Host -foregroundcolor "red" -backgroundcolor "darkgray" " "   
          Write-Host -NoNewLine -foregroundcolor "darkred" -backgroundcolor "darkgray" " Skipping the user (username found) - "
          Write-Host -NoNewLine -foregroundcolor "red" -backgroundcolor "black" " "
          Write-Host -NoNewLine -foregroundcolor "red" -backgroundcolor "black" $NEW_USERNAME
          Write-Host -NoNewLine -foregroundcolor "red" -backgroundcolor "black" " "
          Write-Host -foregroundcolor "red" -backgroundcolor "darkgray" " "
          Write-Host  -NoNewLine -foregroundcolor "darkred" -backgroundcolor "darkgray" "                                        "
          for ($VAR_TEMP_z = 0; $VAR_TEMP_z -lt ($NEW_USERNAME.length - 1); $VAR_TEMP_z++) {
               Write-Host -NoNewLine -foregroundcolor "red" -backgroundcolor "darkgray" " "
          }
          Write-Host -foregroundcolor "red" -backgroundcolor "darkgray" "  "
          $VAR_USER_CREATION = $False
     }
     
     $VAR_OBJ_SEARCH = Get-ADUser -Filter {(GivenName -eq $VAR_TMP_DOMAIN1_USER.GivenName) -and (Surname -eq $VAR_TMP_DOMAIN1_USER.Surname)}
     If (($VAR_OBJ_SEARCH -eq $Null) -and ($VAR_USER_CREATION -eq $True)) {
          $VAR_USER_CREATION = $True
     } Else {
               $VAR_USER_CREATION = $False
               Write-Host " "
               Write-Host  -NoNewLine -foregroundcolor "darkred" -backgroundcolor "darkgray" "                                     "
               for ($VAR_TEMP_z = 0; $VAR_TEMP_z -lt ($NEW_USERNAME.length - 1); $VAR_TEMP_z++) {
                    Write-Host -NoNewLine -foregroundcolor "red" -backgroundcolor "darkgray" " "
               }
               Write-Host -foregroundcolor "red" -backgroundcolor "darkgray" " "   
               Write-Host -NoNewLine -foregroundcolor "darkred" -backgroundcolor "darkgray" " Skipping the user (name found) - "
               Write-Host -NoNewLine -foregroundcolor "red" -backgroundcolor "black" " "
               Write-Host -NoNewLine -foregroundcolor "red" -backgroundcolor "black" $NEW_USERNAME
               Write-Host -NoNewLine -foregroundcolor "red" -backgroundcolor "black" " "
               Write-Host -foregroundcolor "red" -backgroundcolor "darkgray" " "
               Write-Host  -NoNewLine -foregroundcolor "darkred" -backgroundcolor "darkgray" "                                    "
               for ($VAR_TEMP_z = 0; $VAR_TEMP_z -lt ($NEW_USERNAME.length - 1); $VAR_TEMP_z++) {
                    Write-Host -NoNewLine -foregroundcolor "red" -backgroundcolor "darkgray" " "
               }
               Write-Host -foregroundcolor "red" -backgroundcolor "darkgray" "  "
          }
     
     If ($VAR_USER_CREATION -eq $True) {
          $VAR_TEMP_SUBJECT = "Starting the migration for " + $NEW_USERNAME
          $VAR_TEMP_OBJ_SMTP = new-object net.Mail.SmtpClient
          $VAR_TEMP_OBJ_SMTP.Host = "mx.NEW-DOMAIN.LOCAL"
          $VAR_TEMP_OBJ_MSG = new-object Net.Mail.MailMessage
          $VAR_TEMP_OBJ_MSG.From = "scripts@NEW-DOMAIN.LOCAL"
          $VAR_TEMP_OBJ_MSG.ReplyTo = "it@new-domain.net.au"
          $VAR_TEMP_OBJ_MSG.To.Add("it@new-domain.net.au")
          $VAR_TEMP_OBJ_MSG.Subject = "DOMAIN MIGRATION NOTICE - Starting new user migration"
          $VAR_TEMP_OBJ_MSG.Subject = $VAR_TEMP_SUBJECT
          $VAR_TEMP_OBJ_MSG.IsBodyHtml = $false
          $VAR_TEMP_OBJ_SMTP.send($VAR_TEMP_OBJ_MSG)
          Clear-Variable VAR_TEMP_OBJ_SMTP
          Clear-Variable VAR_TEMP_OBJ_MSG
          Clear-Variable VAR_TEMP_SUBJECT
          
          # *******************************************************
          # *******************************************************
          # Start user creation
          
          Write-Host " "
          Write-Host "Creating the user"
          New-ADUser $VAR_TMP_DOMAIN1_USER.DisplayName -SamAccountName $NEW_USERNAME -Enabled $true -AccountPassword $New_Password -Department $Department -Description $VAR_TMP_DOMAIN1_USER.Description -DisplayName $VAR_TMP_DOMAIN1_USER.DisplayName -GivenName $VAR_TMP_DOMAIN1_USER.GivenName -HomePhone $VAR_TMP_DOMAIN1_USER.HomePhone -MobilePhone $VAR_TMP_DOMAIN1_USER.MobilePhone -OfficePhone $VAR_TMP_DOMAIN1_USER.OfficePhone -Surname $VAR_TMP_DOMAIN1_USER.Surname -Title $VAR_TMP_DOMAIN1_USER.Title -City $City -Company $Company -Country $Country -HomePage $HomePage -fax $fax -Office $Office -PostalCode $PostalCode -State $State -StreetAddress $StreetAddress -UserPrincipalName $NEW_UPN -HomeDrive $HOMEDRIVE -HomeDirectory $NEW_HOMEPATH -ScriptPath $LOGINPATH
          # Move the new user to the correct OU
          Move-ADObject -Identity "$NEW_TEMP_IDENTITY" -TargetPath "$VAR_Domain2_OU,$VAR_DOMAIN2_FQDN" -Server $VAR_Domain2_Server
          
          # *******************************************************
          # *******************************************************
          # Add new user to standard set of groups
          
          Write-Host "Adding user to a standard set of groups"
          Add-ADGroupMember -Identity "dg_Office - Office" -Members $NEW_USERNAME
          Add-ADGroupMember -Identity "sg_Location - All  Staff" -Members $NEW_USERNAME
          Add-ADGroupMember -Identity "Company" -Members $NEW_USERNAME
          
          # *******************************************************
          # *******************************************************
          # Convert user from OLD DOMAIN's signature group to NEW DOMAIN's signature group
          
          Write-Host -NoNewLine "Adding user to their correct email signature - "
          Foreach ($VAR_TEMP_DOMAIN1_USER_GROUPS in $VAR_DOMAIN1_USER_GROUPS) {
               if ($VAR_TEMP_DOMAIN1_USER_GROUPS -eq "Signature__Mobile") {
                    Write-Host -foregroundcolor "magenta" "Location with mobile"
                    Add-ADGroupMember -Identity "Signature_Location_Mobile" -Members $NEW_USERNAME
               }
               if ($VAR_TEMP_DOMAIN1_USER_GROUPS -eq "Signature__noMobile") {
                    Write-Host -foregroundcolor "magenta" "Location without mobile"
                    Add-ADGroupMember -Identity "Signature_Location_noMobile" -Members $NEW_USERNAME
               }
               if ($VAR_TEMP_DOMAIN1_USER_GROUPS -eq "Signature__Location2_Mobile") {
                    Write-Host -foregroundcolor "magenta" "Location2 with mobile"
                    Add-ADGroupMember -Identity "Signature_Location2_Mobile" -Members $NEW_USERNAME
               }
          }
          
          # *******************************************************
          # *******************************************************
          # Convert user from OLD DOMAIN's groups to NEW DOMAIN's groups
          
          Write-Host -NoNewLine "Adding user to their correct groups - "
          Foreach ($VAR_TEMP_DOMAIN1_USER_GROUPS in $VAR_DOMAIN1_USER_GROUPS) {
               if ($VAR_TEMP_DOMAIN1_USER_GROUPS -eq "EH_UniData") {
                    Write-Host -foregroundcolor "magenta" -NoNewLine "Unidata, "
                    Add-ADGroupMember -Identity "UniData" -Members $NEW_USERNAME
               }
               if ($VAR_TEMP_DOMAIN1_USER_GROUPS -eq "NationalProjects") {
                    Write-Host -foregroundcolor "magenta" -NoNewLine "National Group, "
                    Add-ADGroupMember -Identity "NationalGroup" -Members $NEW_USERNAME
               }
               if ($VAR_TEMP_DOMAIN1_USER_GROUPS -eq "NationalProjects_NationalComercial") {
                    Write-Host -foregroundcolor "magenta" -NoNewLine "National Projects - Commercial, "
                    Add-ADGroupMember -Identity "NationalProject_NationalComercial" -Members $NEW_USERNAME
               }
          }
          Write-Host " "
          
          # *******************************************************
          # *******************************************************
          # Convert user from OLD DOMAIN's VDI group to NEW DOMAIN's VDI group
          
          Write-Host -NoNewLine "Adding user to their correct VDI Pool - "
          Foreach ($VAR_TEMP_DOMAIN1_USER_GROUPS in $VAR_DOMAIN1_USER_GROUPS) {
               if ($VAR_TEMP_DOMAIN1_USER_GROUPS -eq "VP100") {
                         Write-Host -foregroundcolor "magenta" "100 (Administration)"
                         Add-ADGroupMember -Identity "sg_VDI Pool - 100" -Members $NEW_USERNAME
               }
               if ($VAR_TEMP_DOMAIN1_USER_GROUPS -eq "VP200") {
                         Write-Host -foregroundcolor "magenta" "200 (Design, Drawings & Estimators"
                    Add-ADGroupMember -Identity "sg_VDI Pool - 200" -Members $NEW_USERNAME
               }
               if ($VAR_TEMP_DOMAIN1_USER_GROUPS -eq "VP300") {
                         Write-Host -foregroundcolor "magenta" "300 (Project Managers & Contract Administrators)"
                    Add-ADGroupMember -Identity "sg_VDI Pool - 300" -Members $NEW_USERNAME
               }
               if ($VAR_TEMP_DOMAIN1_USER_GROUPS -eq "VP400") {
                         Write-Host -foregroundcolor "magenta" "400 (Foremen & WHSO)"
                    Add-ADGroupMember -Identity "sg_VDI Pool - 400" -Members $NEW_USERNAME
               }
          }
          
          # *******************************************************
          # *******************************************************
          # Fix permission issues on the new user (arcaic problem)
          
          $VAR_OBJ_NEW_USERNAME = Get-ADUser -Identity $NEW_USERNAME -properties *
          dsacls "$VAR_OBJ_NEW_USERNAME.DistinguishedName" /P:N | Out-Null
          
          # *******************************************************
          # *******************************************************
          # Start copying the user's home drive to new fileserver
          
          Write-Host " "
          Write-Host "Creating user's home drive"
          # Create the new home drive for the user
          New-Item -Path NEW_FS: -Name $NEW_USERNAME -ItemType directory | Out-Null
          Start-Sleep -Seconds 5
          
          # *******************************************************
          # *******************************************************
          # Change the permissions to give user modify rights on their home drive
          
          Write-Host "Setting permissions on the user's home drive"
          # Run the ACL setter to enable the security on the user's new home drive.
          $VAR_colRights = [System.Security.AccessControl.FileSystemRights]"Read, Write, Modify, DeleteSubdirectoriesAndFiles"
          $VAR_InheritanceFlag = [System.Security.AccessControl.InheritanceFlags]::ContainerInherit
          $VAR_PropagationFlag = [System.Security.AccessControl.PropagationFlags]::None
          $VAR_objType =[System.Security.AccessControl.AccessControlType]::Allow
          $VAR_objUser = New-Object System.Security.Principal.NTAccount($NEW_FQ_USERNAME)
          $VAR_objACE = New-Object System.Security.AccessControl.FileSystemAccessRule($VAR_objUser, $VAR_colRights, $VAR_InheritanceFlag, $VAR_PropagationFlag, $VAR_objType)
          $VAR_objACL = Get-ACL "NEW_FS:\$NEW_USERNAME"
          $VAR_objACL.AddAccessRule($VAR_objACE)
          Set-ACL "NEW_FS:\$NEW_USERNAME" $VAR_objACL
          
          Write-Host "Setting permissions on the user's home drive #2"
          # Rerun the ACL setter to enable object/file inheritance
          $VAR_colRights = [System.Security.AccessControl.FileSystemRights]"Read, Write, Modify, DeleteSubdirectoriesAndFiles"
          $VAR_InheritanceFlag = [System.Security.AccessControl.InheritanceFlags]::ObjectInherit
          $VAR_PropagationFlag = [System.Security.AccessControl.PropagationFlags]::None
          $VAR_objType =[System.Security.AccessControl.AccessControlType]::Allow
          $VAR_objUser = New-Object System.Security.Principal.NTAccount($NEW_FQ_USERNAME)
          $VAR_objACE = New-Object System.Security.AccessControl.FileSystemAccessRule($VAR_objUser, $VAR_colRights, $VAR_InheritanceFlag, $VAR_PropagationFlag, $VAR_objType)
          $VAR_objACL = Get-ACL "NEW_FS:\$NEW_USERNAME"
          $VAR_objACL.AddAccessRule($VAR_objACE)
          Set-ACL "NEW_FS:\$NEW_USERNAME" $VAR_objACL
          
          If ($VAR_CONFIG_COPY_FILES -eq $True) {
               Write-Host "Copying user's home drive (with over writing enabled)"
               # Copy old home drive to the new home drive
               $VAR_items = get-childitem $OLD_HOMEPATH -recurse -exclude *.db
               if ($VAR_items.length -gt 0)
               {
                    $VAR_TEMP_COUNT_1 = 0
                    foreach ($VAR_item in $VAR_items)
                    {
                         $VAR_target = join-path NEW_FS:\$NEW_USERNAME $VAR_item.FullName.Substring($OLD_HOMEPATH.Length)
                         if (-not($VAR_item.PSIsContainer -and (test-path($VAR_target))))
                         {
                              $VAR_TEMP_COUNT_1++
                              $VAR_TEMP_PROGRESS_STATUS = "Amount done: " + (($VAR_TEMP_COUNT_1 / $VAR_items.length)  * 100) + "%"
                              Write-Progress -activity "Copying files" -status $VAR_TEMP_PROGRESS_STATUS -SecondsRemaining (Get-Random -minimum 10 -maximum 3000) -percentComplete (($VAR_TEMP_COUNT_1 / $VAR_items.length)  * 100)
                              $VAR_TEMP_ERROR_SETTING = $ErrorActionPreference
                              $ErrorActionPreference = "SilentlyContinue"
                              copy-item -force -path $VAR_item.FullName -destination $VAR_target -ErrorVariable VAR_TMP_OBJ_ERRORS -ErrorAction SilentlyContinue
                              $ErrorActionPreference = $VAR_TEMP_ERROR_SETTING
                              foreach ($VAR_TMP_ERRORS in $VAR_TMP_OBJ_ERRORS) {
                                   $VAR_TEMP_SUBJECT = "ERROR copying for " + $NEW_USERNAME
                                   $VAR_TEMP_BODY = "File - " + $VAR_item.FullName
                                   $VAR_TEMP_OBJ_SMTP = new-object net.Mail.SmtpClient
                                   $VAR_TEMP_OBJ_SMTP.Host = "mx.NEW-DOMAIN.LOCAL"
                                   $VAR_TEMP_OBJ_MSG = new-object Net.Mail.MailMessage
                                   $VAR_TEMP_OBJ_MSG.From = "scripts@NEW-DOMAIN.LOCAL"
                                   $VAR_TEMP_OBJ_MSG.ReplyTo = "it@new-domain.net.au"
                                   $VAR_TEMP_OBJ_MSG.To.Add("it@new-domain.net.au")
                                   $VAR_TEMP_OBJ_MSG.Subject = $VAR_TEMP_SUBJECT
                                   $VAR_TEMP_OBJ_MSG.Body = $VAR_TEMP_BODY
                                   $VAR_TEMP_OBJ_MSG.IsBodyHtml = $false
                                   $VAR_TEMP_OBJ_SMTP.send($VAR_TEMP_OBJ_MSG)
                                   Clear-Variable VAR_TEMP_OBJ_SMTP
                                   Clear-Variable VAR_TEMP_OBJ_MSG
                                   Clear-Variable VAR_TEMP_SUBJECT
                                   Clear-Variable VAR_TEMP_BODY
                              }
                         }
                    }
                    
                    Write-Progress -activity "Copying files" -status "Percent copied: " -percentComplete (100)
                    # Clears the variables for the next item
                    Clear-Variable VAR_TEMP_COUNT_1
                    Clear-Variable VAR_items
                    Clear-Variable VAR_item
                    Clear-Variable VAR_target
                    Write-Progress -activity "Copying files" -status "Percent copied: " -Completed
               }
          }
          
          If ($VAR_CONFIG_EXPORT_MAIL -eq $True) {
               # *******************************************************
               # *******************************************************
               # Start user mailbox export
               
               Write-Host " "
               Write-Host "Connecting to OLD mail server"
               # Create a session to the OLD domain's mail server
               Remove-PSSession $VAR_MX_Session_Domain1
               Remove-PSSession $VAR_MX_Session_Domain2
               $VAR_MX_Session_Domain1 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $VAR_Domain1_Exchange_PS -Authentication Kerberos -Credential $VAR_DOMAIN1_CREDENTIALS
               Import-PSSession $VAR_MX_Session_Domain1 -DisableNameChecking | Out-Null
               
               $VAR_Mailbox_OLD_Status = Get-User -Identity $VAR_TMP_DOMAIN1_USER.sAMAccountName | Select-Object RecipientType
               
               if ($VAR_Mailbox_OLD_Status.RecipientType -match "Mailbox") { 
                    $VAR_TMP_USER_ALIAS = Get-User -Identity $VAR_TMP_DOMAIN1_USER.sAMAccountName | Get-Mailbox
                    
                    # Copy an empty PST file to be used later
                    #   use a dirty dirty hack to pre-create the file so XCOPY will stop prompting if its a file or a directory. XCOPY is ignoring the /I option as well.
                    CD K:
                    ECHO 0 > $VAR_Export_location_PST
                    XCOPY /Y $VAR_SOURCE_PST $VAR_Export_location  | Out-Null
                    CD DOMAIN2:
                    
                    Write-Host "Exporting old mailbox to a PST"
                    # Start the Mailbox export
                    #Get-MailboxExportRequest -Status Completed | Remove-MailboxExportRequest -Confirm:$false | Out-Null
                    New-MailboxExportRequest -Mailbox $VAR_TMP_DOMAIN1_USER.sAMAccountName -FilePath $VAR_Export_location -BadItemLimit 490 -AcceptLargeDataLoss:$true -WarningAction:SilentlyContinue -Name $VAR_OBJ_CURRENT_PROCCESS.Id | Out-Null
                    
                    $VAR_TEMP_EXPORT_NAME = $VAR_TMP_USER_ALIAS.Alias + "\" + $VAR_OBJ_CURRENT_PROCCESS.Id
                    
                    $VAR_Export_Status = Get-MailboxExportRequest -Identity $VAR_TEMP_EXPORT_NAME
                    $VAR_EXPORT_Status_Statistics = Get-MailboxExportRequest -Identity $VAR_TEMP_EXPORT_NAME | Get-MailboxExportRequestStatistics
                    $VAR_TEMP_LINE_CLEARED = $false
                    $VAR_TEMP_EXPORT_WORKED = $true
                    
                    :EXPORTMAIL while ($VAR_Export_Status.status -ne "Completed") {
                         $VAR_TEMP_PROGRESS_STATUS = "Amount done: " + $VAR_EXPORT_Status_Statistics.PercentComplete + "%"
                         Write-Progress -activity "Exporting data" -Id 123 -status $VAR_TEMP_PROGRESS_STATUS -percentComplete ($VAR_EXPORT_Status_Statistics.PercentComplete)
                         If ($VAR_Export_Status.status -eq "Failed") {
                              $VAR_TEMP_SUBJECT = "Failed to export " + $NEW_USERNAME
                              $VAR_TEMP_OBJ_SMTP = new-object net.Mail.SmtpClient
                              $VAR_TEMP_OBJ_SMTP.Host = "mx.NEW-DOMAIN.LOCAL"
                              $VAR_TEMP_OBJ_MSG = new-object Net.Mail.MailMessage
                              $VAR_TEMP_OBJ_MSG.From = "scripts@NEW-DOMAIN.LOCAL"
                              $VAR_TEMP_OBJ_MSG.ReplyTo = "it@new-domain.net.au"
                              $VAR_TEMP_OBJ_MSG.To.Add("it@new-domain.net.au")
                              $VAR_TEMP_OBJ_MSG.Subject = "DOMAIN MIGRATION NOTICE - Failure to export"
                              $VAR_TEMP_OBJ_MSG.Subject = $VAR_TEMP_SUBJECT
                              $VAR_TEMP_OBJ_MSG.IsBodyHtml = $false
                              $VAR_TEMP_OBJ_SMTP.send($VAR_TEMP_OBJ_MSG)
                              Clear-Variable VAR_TEMP_OBJ_SMTP
                              Clear-Variable VAR_TEMP_OBJ_MSG
                              Clear-Variable VAR_TEMP_SUBJECT
                              $VAR_TEMP_EXPORT_WORKED = $false
                              Break EXPORTMAIL
                         }
                         
                         While ($VAR_Export_Status.status -eq "Queued") {
                              Write-Host -NoNewLine "                              "
                              Start-Sleep -Milliseconds 50
                              Write-Host -NoNewLine "`rWaiting for process to start /"
                              Start-Sleep -Milliseconds 50
                              Write-Host -NoNewLine "`rWaiting for process to start - "
                              Write-Host -NoNewLine -foregroundcolor "magenta" $VAR_Export_Status.status
                              Start-Sleep -Milliseconds 50
                              Write-Host -NoNewLine "`rWaiting for process to start `|"
                              Start-Sleep -Milliseconds 50
                              Write-Host -NoNewLine "`rWaiting for process to start - "
                              Write-Host -NoNewLine -foregroundcolor "magenta" $VAR_Export_Status.status
                              Start-Sleep -Milliseconds 50
                              Write-Host -NoNewLine "`rWaiting for process to start `\"
                              Start-Sleep -Milliseconds 50
                              Write-Host -NoNewLine "`rWaiting for process to start - "
                              Write-Host -NoNewLine -foregroundcolor "magenta" $VAR_Export_Status.status
                              Start-Sleep -Milliseconds 50
                              Write-Host -NoNewLine "`rWaiting for process to start `|"
                              Start-Sleep -Milliseconds 50
                              Write-Host -NoNewLine "`rWaiting for process to start - "
                              Write-Host -NoNewLine -foregroundcolor "magenta" $VAR_Export_Status.status
                              Start-Sleep -Milliseconds 100
                              $VAR_Export_Status = Get-MailboxExportRequest -Identity $VAR_TEMP_EXPORT_NAME
                         }
                         
                         If ($VAR_Export_Status.status -ne "Queued" -AND $VAR_TEMP_LINE_CLEARED -eq $false) {
                              Write-Host "`rWaiting for process to start - Has started"
                              $VAR_TEMP_LINE_CLEARED = $true
                         }
                         
                         Start-Sleep -Seconds 15
                         $VAR_Export_Status = Get-MailboxExportRequest -Identity $VAR_TEMP_EXPORT_NAME
                         $VAR_EXPORT_Status_Statistics = Get-MailboxExportRequest -Identity $VAR_TEMP_EXPORT_NAME | Get-MailboxExportRequestStatistics
                    }
                    
                    Write-Progress -activity "Exporting data" -Id 123 -status "Amount done:" -percentComplete (100)
                    Write-Progress -activity "Exporting data" -Id 123 -status "Amount done:" -Completed
                    Write-Host -NoNewLine "Current Status is "
                    Write-Host -foregroundcolor "magenta" $VAR_Export_Status.status
                    Get-MailboxExportRequest -Identity $VAR_TEMP_EXPORT_NAME | Remove-MailboxExportRequest -Confirm:$false | Out-Null
                    
                    Clear-Variable VAR_Export_Status
               }
               
               Write-Host "Disconnecting from OLD mail server"
               Remove-PSSession $VAR_MX_Session_Domain1
          }
          
          # *******************************************************
          # *******************************************************
          # Start user mailbox creation
          
          Write-Host " "
          Write-Host "Connecting to NEW mail server"
          Remove-PSSession $VAR_MX_Session_Domain1
          Remove-PSSession $VAR_MX_Session_Domain2
          $VAR_MX_Session_Domain2 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $VAR_Domain2_Exchange_PS -Authentication Kerberos -Credential $VAR_DOMAIN2_CREDENTIALS
          Import-PSSession $VAR_MX_Session_Domain2 -DisableNameChecking | Out-Null
          
          # Creat the mailbox for the new user
               Write-Host -NoNewLine "Creating mailbox on "
          IF ($VAR_TMP_DOMAIN1_USER.Surname.Substring(0,1) -imatch "^[A-K]*$")
          {    
               Write-Host -foregroundcolor "magenta" "Mailbox Database  (A-K)"
               Enable-Mailbox -Identity $NEW_FQ_USERNAME -Alias $NEW_USERNAME -Database 'Mailbox Database  (A-K)' | Out-Null
          } ELSE {
               Write-Host -foregroundcolor "magenta" "Mailbox Database  (L-Z)"
               Enable-Mailbox -Identity $NEW_FQ_USERNAME -Alias $NEW_USERNAME -Database 'Mailbox Database  (L-Z)' | Out-Null
          }
          
          Write-Host "Checking to see if mailbox was created"
          $VAR_Mailbox_Status = Get-User -Identity $NEW_FQ_USERNAME | Select-Object RecipientType
          while ($VAR_Mailbox_Status.RecipientType -ne "UserMailbox") {
               Write-Host "Sleeping for 10 seconds while waiting for the mailbox to be created"
               for ($VAR_TEMP_i = 0; $VAR_TEMP_i -lt 11; $VAR_TEMP_i++)
               {
                    Write-Progress -activity "Sleeping" -status "Amount done:" -SecondsRemaining (10 - $VAR_TEMP_i) -percentComplete (($VAR_TEMP_i / 10)  * 100)
                    Start-Sleep -Seconds 1
               }
               
               Write-Progress -activity "Sleeping" -status "Amount done:" -SecondsRemaining 0 -percentComplete (100)
               Write-Progress -activity "Sleeping" -status "Amount done:" -Completed
               $VAR_Mailbox_Status = Get-User -Identity $NEW_FQ_USERNAME | Select-Object RecipientType
          }
          
          if ($VAR_Mailbox_Status.RecipientType -match "Mailbox") {
               Write-Host -NoNewLine "Mailbox has been "
               Write-Host -foregroundcolor "magenta" "created"
          }
          
          $VAR_TMP_NEWUSER_ALIAS = Get-User -Identity $NEW_FQ_USERNAME | Get-Mailbox
          
          # *******************************************************
          # *******************************************************
          # Start a domain replication to ensure user details are on the DC near the OCS server
          
          Write-Host " "
          Write-Host -NoNewLine "Starting replication on "
          Write-Host -foregroundcolor "magenta" $VAR_Domain2
          Write-Host "Replicating changes to the domain partition"
          REPADMIN /sync $VAR_REP_DOMAIN_PART $VAR_REP_OCS_DC $VAR_REP_DOMAIN_GUID | Out-Null
          for ($VAR_TEMP_i = 0; $VAR_TEMP_i -lt 31; $VAR_TEMP_i++)
          {
               Write-Progress -activity "Replicating" -status "Amount done:" -SecondsRemaining (30 - $VAR_TEMP_i) -percentComplete (($VAR_TEMP_i / 30)  * 100)
               Start-Sleep -Seconds 1
          }
          
          Write-Progress -activity "Replicating" -status "Amount done:" -SecondsRemaining 0 -percentComplete (100)
          Write-Progress -activity "Replicating" -status "Amount done:" -Completed
          
          REPADMIN /sync $VAR_REP_CONFIG_PART $VAR_REP_OCS_DC $VAR_REP_CONFIG_GUID | Out-Null
          Write-Host "Replicating changes to the configuration partition"
          for ($VAR_TEMP_i = 0; $VAR_TEMP_i -lt 31; $VAR_TEMP_i++)
          {
               Write-Progress -activity "Replicating" -status "Amount done:" -SecondsRemaining (30 - $VAR_TEMP_i) -percentComplete (($VAR_TEMP_i / 30)  * 100)
               Start-Sleep -Seconds 1
          }
          
          Write-Progress -activity "Replicating" -status "Amount done:" -SecondsRemaining 0 -percentComplete (100)
          Write-Progress -activity "Replicating" -status "Amount done:" -Completed
          
          CD DOMAIN2:
          $VAR_OBJ_NEW_USERNAME = Get-ADUser -Identity $NEW_USERNAME -properties *
          
          If ($VAR_CONFIG_IMPORT_MAIL -eq $True) {
               # *******************************************************
               # *******************************************************
               # Start user mailbox import
               
               # Reload the VAR_OBJ_NEW_USERNAME to have new details
               #$VAR_OBJ_NEW_USERNAME = Get-ADUser -Identity $NEW_USERNAME -properties *
               
               Write-Host " "
               Write-Host "Importing from a PST into the new mailbox"
               # Start the Mailbox export
               # Get-MailboxImportRequest -Status Completed | Remove-MailboxImportRequest -Confirm:$false | Out-Null
               New-MailboxImportRequest -Mailbox $NEW_USERNAME -FilePath $VAR_Export_location -BadItemLimit 490 -AcceptLargeDataLoss:$true -WarningAction:SilentlyContinue -Name $VAR_OBJ_CURRENT_PROCCESS.Id | Out-Null
               
               $VAR_TEMP_IMPORT_NAME = $VAR_TMP_NEWUSER_ALIAS.Alias + "\" + $VAR_OBJ_CURRENT_PROCCESS.Id

               $VAR_IMPORT_Status = Get-MailboxImportRequest -Identity $VAR_TEMP_IMPORT_NAME
               $VAR_IMPORT_Status_Statistics = Get-MailboxImportRequest -Identity $VAR_TEMP_IMPORT_NAME | Get-MailboxImportRequestStatistics
               $VAR_TEMP_LINE_CLEARED = $false
               
               $VAR_TEMP_IMPORT = $false
               if ($VAR_IMPORT_Status.status -eq "Completed") {$VAR_TEMP_IMPORT = $true}
               if ($VAR_IMPORT_Status.status -eq "InProgress") {$VAR_TEMP_IMPORT = $true}
               if ($VAR_IMPORT_Status.status -eq "Queued") {$VAR_TEMP_IMPORT = $true}
               if ($VAR_TEMP_IMPORT -eq $false) {break}
               
               :IMPORTMAIL while ($VAR_IMPORT_Status.status -ne "Completed") {
                    $VAR_TEMP_PROGRESS_STATUS = "Amount done: " + $VAR_IMPORT_Status_Statistics.PercentComplete + "%"
                    Write-Progress -activity "Importing data" -Id 123 -status $VAR_TEMP_PROGRESS_STATUS -percentComplete ($VAR_IMPORT_Status_Statistics.PercentComplete)
                    If ($VAR_Export_Status.status -eq "Failed") {
                         $VAR_TEMP_SUBJECT = "Failed to import " + $NEW_USERNAME
                         $VAR_TEMP_OBJ_SMTP = new-object net.Mail.SmtpClient
                         $VAR_TEMP_OBJ_SMTP.Host = "mx.NEW-DOMAIN.LOCAL"
                         $VAR_TEMP_OBJ_MSG = new-object Net.Mail.MailMessage
                         $VAR_TEMP_OBJ_MSG.From = "scripts@NEW-DOMAIN.LOCAL"
                         $VAR_TEMP_OBJ_MSG.ReplyTo = "it@new-domain.net.au"
                         $VAR_TEMP_OBJ_MSG.To.Add("it@new-domain.net.au")
                         $VAR_TEMP_OBJ_MSG.Subject = "DOMAIN MIGRATION NOTICE - Failure to import"
                         $VAR_TEMP_OBJ_MSG.Subject = $VAR_TEMP_SUBJECT
                         $VAR_TEMP_OBJ_MSG.IsBodyHtml = $false
                         $VAR_TEMP_OBJ_SMTP.send($VAR_TEMP_OBJ_MSG)
                         Clear-Variable VAR_TEMP_OBJ_SMTP
                         Clear-Variable VAR_TEMP_OBJ_MSG
                         Clear-Variable VAR_TEMP_SUBJECT
                         Break IMPORTMAIL
                    }
                    
                    While ($VAR_IMPORT_Status.status -eq "Queued") {
                         $VAR_TEMP_LINE_CLEARED = $false
                         Write-Host -NoNewLine "                              "
                         Start-Sleep -Milliseconds 50
                         Write-Host -NoNewLine "`rWaiting for process to start /"
                         Start-Sleep -Milliseconds 50
                         Write-Host -NoNewLine "`rWaiting for process to start - "
                         Write-Host -NoNewLine -foregroundcolor "magenta" $VAR_IMPORT_Status.status
                         Start-Sleep -Milliseconds 50
                         Write-Host -NoNewLine "`rWaiting for process to start `|"
                         Start-Sleep -Milliseconds 50
                         Write-Host -NoNewLine "`rWaiting for process to start - "
                         Write-Host -NoNewLine -foregroundcolor "magenta" $VAR_IMPORT_Status.status
                         Start-Sleep -Milliseconds 50
                         Write-Host -NoNewLine "`rWaiting for process to start `\"
                         Start-Sleep -Milliseconds 50
                         Write-Host -NoNewLine "`rWaiting for process to start - "
                         Write-Host -NoNewLine -foregroundcolor "magenta" $VAR_IMPORT_Status.status
                         Start-Sleep -Milliseconds 50
                         Write-Host -NoNewLine "`rWaiting for process to start `|"
                         Start-Sleep -Milliseconds 50
                         Write-Host -NoNewLine "`rWaiting for process to start - "
                         Write-Host -NoNewLine -foregroundcolor "magenta" $VAR_IMPORT_Status.status
                         Start-Sleep -Milliseconds 100
                         $VAR_IMPORT_Status = Get-MailboxExportRequest
                    }
                    
                    If ($VAR_IMPORT_Status.status -ne "Queued" -AND $VAR_TEMP_LINE_CLEARED -eq $false) {
                         Write-Host "`rWaiting for process to start - Has started"
                         $VAR_TEMP_LINE_CLEARED = $true
                    }
                    
                    Start-Sleep -Seconds 1
                    $VAR_IMPORT_Status = Get-MailboxImportRequest -Identity $VAR_TEMP_IMPORT_NAME
                    $VAR_IMPORT_Status_Statistics = Get-MailboxImportRequest -Identity $VAR_TEMP_IMPORT_NAME | Get-MailboxImportRequestStatistics
               }
               
               Write-Progress -activity "Importing data" -Id 123 -status "Amount done:" -percentComplete (100)
               Write-Progress -activity "Importing data" -Id 123 -status "Amount done:" -Completed
               Write-Host -NoNewLine "Current Status is "
               Write-Host -foregroundcolor "magenta" $VAR_IMPORT_Status.status
               Get-MailboxImportRequest -Identity $VAR_TEMP_IMPORT_NAME | Remove-MailboxImportRequest -Confirm:$false
               
               Write-Host "Disconnecting from OLD mail server"
               Remove-PSSession $VAR_MX_Session_Domain2
          }
          
          # *******************************************************
          # *******************************************************
          # Do work on old account
          
          Write-Host " "
          Write-Host "Making changes to the old user account"
          CD DOMAIN1:
          
          if ($VAR_CONFIG_DISABLE_OLD -eq $true) {
               Write-Host "Disabling the old account"
               Disable-ADAccount -Identity $VAR_TMP_DOMAIN1_USER.sAMAccountName
          }
          
          if ($VAR_CONFIG_CHANGE_PASSWORD -eq $true) {
               Write-Host "Changing the password on the old account"
               Set-ADAccountPassword -Identity $VAR_TMP_DOMAIN1_USER.sAMAccountName -NewPassword $OLD_Password
          }
          
          # Sets redirect on the old account for emails
          if ($VAR_CONFIG_REDIRECT -eq $true) {
               Write-Host "Connecting to OLD mail server"
               # Create a session to the OLD domain's mail server
               Remove-PSSession $VAR_MX_Session_Domain1
               Remove-PSSession $VAR_MX_Session_Domain2
               $VAR_MX_Session_Domain1 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $VAR_Domain1_Exchange_PS -Authentication Kerberos -Credential $VAR_DOMAIN1_CREDENTIALS
               Import-PSSession $VAR_MX_Session_Domain1 -DisableNameChecking | Out-Null
               
               $VAR_TEMP_OVERRIDE_EMAIL = $NEW_USERNAME + "@new-domain.net.au"
               
               Write-Host -NoNewLine "Redirecting email from old account to "
               Write-Host -NoNewLine -foregroundcolor "magenta" $VAR_TEMP_OVERRIDE_EMAIL
               if ($VAR_CONFIG_REDIRECT_KEEP -eq $true) {
                    Write-Host " and keeping a copy in old mailbox"
               }else{
               Write-Host " and not keeping a copy in old mailbox"
               }
               
               $VAR_TEMP_CONTACT_NAME = "EX_" + $VAR_OBJ_NEW_USERNAME.sAMAccountName
               New-MailContact -name $VAR_TEMP_CONTACT_NAME -ExternalEmailAddress $VAR_TEMP_OVERRIDE_EMAIL -OrganizationalUnit $VAR_Domain1_Contact_OU -Alias $VAR_TEMP_CONTACT_NAME | Out-Null
               Set-Mailbox $VAR_TMP_DOMAIN1_USER.sAMAccountName -ForwardingAddress $VAR_TEMP_OVERRIDE_EMAIL -DeliverToMailboxAndForward $VAR_CONFIG_REDIRECT_KEEP | Out-Null
               
               Clear-Variable VAR_TEMP_OVERRIDE_EMAIL
               
               Write-Host "Disconnecting from OLD mail server"
               Remove-PSSession $VAR_MX_Session_Domain1
          }
          
          CD DOMAIN2:
          
          <# *** CURRENTLY BROKEN ****
          
          # *******************************************************
          # *******************************************************
          # Ins     tall the OCS components
          # Enable OCS on the user
               
          Import-Module $VAR_SCRIPT_SOURCE\OCS-ALL.ps1
          CD DOMAIN2:
          
          $VAR_TEMP_URI = "SIP:" + $VAR_OBJ_NEW_USERNAME.EmailAddress
          Get-adUser -Identity $NEW_USERNAME | New-OcsUser -URI $VAR_TEMP_URI -homeServer ($VAR_Domain2_POOL)
          
          $VAR_TEMP_OCS_QUERY = "Select * from MSFT_SIPESUserSetting where PrimaryURI = `'sip:" + $VAR_OBJ_NEW_USERNAME.EmailAddress + "`'"
          Write-Host ""
          Write-Host -NoNewLine "Query = "
          Write-Host -foregroundcolor "magenta" $VAR_TEMP_OCS_QUERY
          
          Write-Host "Press any key to continue ..."
          $VAR_TEMP_PAUSE = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
          
          # Updates the OCS object for the user
          $OCSUser = Get-WmiObject -Query $VAR_TEMP_OCS_QUERY
          $OCSUser.HomeServerDN = $VAR_Domain2_HomeServerDN
          $OCSUser.EnabledForEnhancedPresence = $true 
          #Enabled for federation.
          $OCSUser.EnabledForFederation = $true
          #Enabled for remote user access.
          $OCSUser.EnabledForInternetAccess = $true
          #Enabled for public instant messaging (IM) connectivity.
          $OCSUser.PublicNetworkEnabled = $true
          $OCSUser.UCEnabled = $true
          $OCSUser.LineURI = ""
          #The string for Distinguished Name of the phone usage policy.
          $OCSUser.UCPolicy = ""
          $OCSUser.put() | out-null
          #>
          
          # *******************************************************
          # *******************************************************
          # Clears the variables for the next item
          
          If ($VAR_colRights -ne $Null) {Clear-Variable VAR_colRights}
          If ($VAR_InheritanceFlag -ne $Null) {Clear-Variable VAR_InheritanceFlag}
          If ($VAR_PropagationFlag -ne $Null) {Clear-Variable VAR_PropagationFlag}
          If ($VAR_objType -ne $Null) {Clear-Variable VAR_objType}
          If ($VAR_objUser -ne $Null) {Clear-Variable VAR_objUser}
          If ($VAR_objACE -ne $Null) {Clear-Variable VAR_objACE}
          If ($VAR_objACL -ne $Null) {Clear-Variable VAR_objACL}
          If ($VAR_items -ne $Null) {Clear-Variable VAR_items}
          If ($VAR_TEMP_COUNT_1 -ne $Null) {Clear-Variable VAR_TEMP_COUNT_1}
          If ($VAR_target -ne $Null) {Clear-Variable VAR_target}
          If ($VAR_Mailbox_OLD_Status -ne $Null) {Clear-Variable VAR_Mailbox_OLD_Status}
          If ($VAR_Export_Status -ne $Null) {Clear-Variable VAR_Export_Status}
          If ($VAR_Mailbox_Status -ne $Null) {Clear-Variable VAR_Mailbox_Status}
          If ($VAR_TEMP_i -ne $Null) {Clear-Variable VAR_TEMP_i}
          If ($VAR_IMPORT_Status -ne $Null) {Clear-Variable VAR_IMPORT_Status}
          If ($VAR_TEMP_IMPORT -ne $Null) {Clear-Variable VAR_TEMP_IMPORT}
     }
     
     Write-Host " "
     if ($VAR_USER_CREATION -eq $true) {
          Write-Host -NoNewLine "Completed the current user - "
          Write-Host -foregroundcolor "magenta" $NEW_USERNAME
          $VAR_TEMP_SUBJECT = "Finished processing " + $NEW_USERNAME
          $VAR_TEMP_OBJ_SMTP = new-object net.Mail.SmtpClient
          $VAR_TEMP_OBJ_SMTP.Host = "mx.NEW-DOMAIN.LOCAL"
          $VAR_TEMP_OBJ_MSG = new-object Net.Mail.MailMessage
          $VAR_TEMP_OBJ_MSG.From = "scripts@NEW-DOMAIN.LOCAL"
          $VAR_TEMP_OBJ_MSG.ReplyTo = "it@new-domain.net.au"
          $VAR_TEMP_OBJ_MSG.To.Add("it@new-domain.net.au")
          $VAR_TEMP_OBJ_MSG.Subject = "DOMAIN MIGRATION NOTICE - User process"
          $VAR_TEMP_OBJ_MSG.Subject = $VAR_TEMP_SUBJECT
          $VAR_TEMP_OBJ_MSG.IsBodyHtml = $false
          $VAR_TEMP_OBJ_SMTP.send($VAR_TEMP_OBJ_MSG)
          Clear-Variable VAR_TEMP_OBJ_SMTP
          Clear-Variable VAR_TEMP_OBJ_MSG
          Clear-Variable VAR_TEMP_SUBJECT
     } Else {
          Write-Host -foregroundcolor "darkred" -backgroundcolor "darkgray" "                  "
          Write-Host -foregroundcolor "darkred" -backgroundcolor "darkgray" " Skipped the user "
          Write-Host -foregroundcolor "darkred" -backgroundcolor "darkgray" "                  "
          $VAR_TEMP_SUBJECT = "Skipped " + $NEW_USERNAME
          $VAR_TEMP_OBJ_SMTP = new-object net.Mail.SmtpClient
          $VAR_TEMP_OBJ_SMTP.Host = "mx.NEW-DOMAIN.LOCAL"
          $VAR_TEMP_OBJ_MSG = new-object Net.Mail.MailMessage
          $VAR_TEMP_OBJ_MSG.From = "scripts@NEW-DOMAIN.LOCAL"
          $VAR_TEMP_OBJ_MSG.ReplyTo = "it@new-domain.net.au"
          $VAR_TEMP_OBJ_MSG.To.Add("it@new-domain.net.au")
          $VAR_TEMP_OBJ_MSG.Subject = "DOMAIN MIGRATION NOTICE - User process"
          $VAR_TEMP_OBJ_MSG.Subject = $VAR_TEMP_SUBJECT
          $VAR_TEMP_OBJ_MSG.IsBodyHtml = $false
          $VAR_TEMP_OBJ_SMTP.send($VAR_TEMP_OBJ_MSG)
          Clear-Variable VAR_TEMP_OBJ_SMTP
          Clear-Variable VAR_TEMP_OBJ_MSG
          Clear-Variable VAR_TEMP_SUBJECT
     }
     
     If ($VAR_CONFIG_PAUSE -eq $true) {
          Write-Host " "
          Write-Host -foregroundcolor "Yellow" -backgroundcolor "Blue" "                                                 "
          Write-Host -foregroundcolor "Yellow" -backgroundcolor "Blue" " Press any key to continue with the next user... "
          Write-Host -foregroundcolor "Yellow" -backgroundcolor "Blue" "                                                 "
          $VAR_TEMP_PAUSE = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
     }
     
     If ($NEW_USERNAME_TMP -ne $Null) {Clear-Variable NEW_USERNAME_TMP}
     If ($NEW_USERNAME -ne $Null) {Clear-Variable NEW_USERNAME}
     If ($NEW_TEMP_IDENTITY -ne $Null) {Clear-Variable NEW_TEMP_IDENTITY}
     If ($VAR_Export_location_PST -ne $Null) {Clear-Variable VAR_Export_location_PST}
     If ($VAR_Export_location -ne $Null) {Clear-Variable VAR_Export_location}
     If ($NEW_UPN -ne $Null) {Clear-Variable NEW_UPN}
     If ($NEW_HOMEPATH -ne $Null) {Clear-Variable NEW_HOMEPATH}
     If ($NEW_FQ_USERNAME -ne $Null) {Clear-Variable NEW_FQ_USERNAME}
     If ($OLD_HOMEPATH -ne $Null) {Clear-Variable OLD_HOMEPATH}
     If ($VAR_DOMAIN1_USER_GROUPS -ne $Null) {Clear-Variable VAR_DOMAIN1_USER_GROUPS}
}

# *******************************************************
# *******************************************************
# Cleanup

# Changes out of the DOMAIN drives, so they can be closed by PowerShell
# Do extra clean up
CD C:
$VAR_NET_DRIVE_HME.RemoveNetworkDrive("J:","true","true")
$VAR_NET_DRIVE_PST.RemoveNetworkDrive("K:","true","true")
#Remove-PSSession *
Remove-Module ActiveDirectory
#Remove-Module $VAR_SCRIPT_SOURCE\OCS-ALL.ps1

Write-Host " "
Write-Host " "
Write-Host " "
