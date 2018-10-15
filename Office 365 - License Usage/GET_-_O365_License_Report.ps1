$username = "admin@office365.com"
$passwordFile = "D:\Scripts\password.txt"

$date= get-date -format "yyyyMMdd"
$LicenseReport="D:\Scripts\Licenses_$date.xlsx" 

$fromaddress = "it.scripts@company.com" 
$toaddress = "boss1@company.com" 
$CCaddress = "it.team@company.com" 
$Subject = "Office 365 license report" 
$body = "See attached file for the report"
$smtpserver = "mail.isp.net" 

#####################################################
#####################################################
##                                                 ##
##                DO NOT EDIT BELOW                ##
##                                                 ##
#####################################################
#####################################################

# Create an array of all license types for Office 365
$Sku = @{ 
    "EXCHANGEDESKLESS" = "Office 365 (Kiosk)"
	"DESKLESSPACK" = "Office 365 (Plan K1)"
    "DESKLESSWOFFPACK" = "Office 365 (Plan K2)"
    "LITEPACK" = "Office 365 (Plan P1)"
    "EXCHANGESTANDARD" = "Office 365 Exchange Online Only"
    "STANDARDPACK" = "Office 365 (Plan E1)"
    "STANDARDWOFFPACK" = "Office 365 (Plan E2)"
    "ENTERPRISEPACK" = "Office 365 (Plan E3)"
    "ENTERPRISEPACKLRG" = "Office 365 (Plan E3)"
    "ENTERPRISEWITHSCAL" = "Office 365 (Plan E4)"
    "STANDARDPACK_STUDENT" = "Office 365 (Plan A1) for Students"
    "STANDARDWOFFPACKPACK_STUDENT" = "Office 365 (Plan A2) for Students"
    "ENTERPRISEPACK_STUDENT" = "Office 365 (Plan A3) for Students"
    "ENTERPRISEWITHSCAL_STUDENT" = "Office 365 (Plan A4) for Students"
    "STANDARDPACK_FACULTY" = "Office 365 (Plan A1) for Faculty"
    "STANDARDWOFFPACKPACK_FACULTY" = "Office 365 (Plan A2) for Faculty"
    "ENTERPRISEPACK_FACULTY" = "Office 365 (Plan A3) for Faculty"
    "ENTERPRISEWITHSCAL_FACULTY" = "Office 365 (Plan A4) for Faculty"
    "ENTERPRISEPACK_B_PILOT" = "Office 365 (Enterprise Preview)"
    "STANDARD_B_PILOT" = "Office 365 (Small Business Preview)"
}

$VarSku = @{ 
    "syndication-account:EXCHANGEDESKLESS" = "Office 365 (Kiosk)"
	"syndication-account:DESKLESSPACK" = "Office 365 (Plan K1)"
    "syndication-account:DESKLESSWOFFPACK" = "Office 365 (Plan K2)"
    "syndication-account:LITEPACK" = "Office 365 (Plan P1)"
    "syndication-account:EXCHANGESTANDARD" = "Office 365 Exchange Online Only"
    "syndication-account:STANDARDPACK" = "Office 365 (Plan E1)"
    "syndication-account:STANDARDWOFFPACK" = "Office 365 (Plan E2)"
    "syndication-account:ENTERPRISEPACK" = "Office 365 (Plan E3)"
    "syndication-account:ENTERPRISEPACKLRG" = "Office 365 (Plan E3)"
    "syndication-account:ENTERPRISEWITHSCAL" = "Office 365 (Plan E4)"
    "syndication-account:STANDARDPACK_STUDENT" = "Office 365 (Plan A1) for Students"
    "syndication-account:STANDARDWOFFPACKPACK_STUDENT" = "Office 365 (Plan A2) for Students"
    "syndication-account:ENTERPRISEPACK_STUDENT" = "Office 365 (Plan A3) for Students"
    "syndication-account:ENTERPRISEWITHSCAL_STUDENT" = "Office 365 (Plan A4) for Students"
    "syndication-account:STANDARDPACK_FACULTY" = "Office 365 (Plan A1) for Faculty"
    "syndication-account:STANDARDWOFFPACKPACK_FACULTY" = "Office 365 (Plan A2) for Faculty"
    "syndication-account:ENTERPRISEPACK_FACULTY" = "Office 365 (Plan A3) for Faculty"
    "syndication-account:ENTERPRISEWITHSCAL_FACULTY" = "Office 365 (Plan A4) for Faculty"
    "syndication-account:ENTERPRISEPACK_B_PILOT" = "Office 365 (Enterprise Preview)"
    "syndication-account:STANDARD_B_PILOT" = "Office 365 (Small Business Preview)"
}

#Import the Microsoft Hosted Services modules
Import-Module MSOnline 

#Prompt for user credentials
$password = Get-Content $passwordFile
$secstr = New-Object -TypeName System.Security.SecureString
$password.ToCharArray() | ForEach-Object {$secstr.AppendChar($_)}
$O365Cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $username, $secstr

# Create a new session to the Microsoft Hosted Services
$O365Session = New-PSSession –ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $O365Cred -Authentication Basic -AllowRedirection 
Import-PSSession $O365Session -AllowClobber 

# Connect to the Microsoft Hosted Services
Connect-MsolService –Credential $O365Cred

# Get the information we are after
$LicenseCount = Get-MsolAccountSku | Select-Object AccountSkuID, ConsumedUnits, ActiveUnits
$LicenseUsage = Get-MsolUser -all | where {$_.isLicensed -eq "True"} | Sort-Object DisplayName

#Create an Excel object to use
$ExcelObject = new-Object -comobject Excel.Application  
$ExcelObject.visible = $false
$ExcelObject.DisplayAlerts =$false

# Create Excel file
#   The code below fills out the Excel document with data. 
$ActiveWorkbook = $ExcelObject.Workbooks.Add()  
$ActiveWorksheet = $ActiveWorkbook.Worksheets.Item(1)

$ActiveWorksheet.Cells.Item(1,1) = "License Type"  
$ActiveWorksheet.cells.item(1,2) = "# of licenses used" 
$ActiveWorksheet.cells.item(1,3) = "Total amount of licenses"

for ($i=1;$i -lt 4; $i++){
    $ActiveWorksheet.Cells.Item(1,$i).Interior.ColorIndex = 19
    $ActiveWorksheet.Cells.Item(1,$i).Font.ColorIndex = 11
    $ActiveWorksheet.Cells.Item(1,$i).Font.Bold = "True"
}

$varI = 2
ForEach ($j in $LicenseCount){
    #if ($j.AccountSkuID -eq "syndication-account:ENTERPRISEPACK"){
    #    $ActiveWorksheet.Cells.Item($varI,1) = "Office 365 (Plan E3)"
    #}ELSEIF ($j.AccountSkuID -eq "syndication-account:STANDARDWOFFPACK") {
    #    $ActiveWorksheet.Cells.Item($varI,1) = "Office 365 (Plan E2)"
    #}ELSE {
    #    $ActiveWorksheet.Cells.Item($varI,1) = "Office 365 (Kiosk)"
	#}
	$ActiveWorksheet.Cells.Item($varI,1) = $VarSku.Item($j.AccountSkuID)
    $ActiveWorksheet.Cells.Item($varI,2) = $j.ConsumedUnits
    $ActiveWorksheet.Cells.Item($varI,3) = $j.ActiveUnits
    $varI++
}

$varJ = $varI + 2
$ActiveWorksheet.Cells.Item($varJ,1) = "Username"  
$ActiveWorksheet.cells.item($varJ,2) = "License Type" 

for ($j=1;$j -lt 3; $j++){
    $ActiveWorksheet.Cells.Item($varJ,$j).Interior.ColorIndex = 19
    $ActiveWorksheet.Cells.Item($varJ,$j).Font.ColorIndex = 11
    $ActiveWorksheet.Cells.Item($varJ,$j).Font.Bold = "True"
}

$varK = $varJ + 1
$varUserCount = 0
ForEach ($j in $LicenseUsage){
    $ActiveWorksheet.Cells.Item($varK,1) = $j.DisplayName
    $ActiveWorksheet.Cells.Item($varK,2) = $Sku.Item($j.licenses[0].AccountSku.SkuPartNumber)
    $varK++
    $varUserCount++
}

$ActiveWorksheet.cells.item($varJ,4) = "Total Users"
$ActiveWorksheet.Cells.Item($varJ,4).Interior.ColorIndex = 19
$ActiveWorksheet.Cells.Item($varJ,4).Font.ColorIndex = 11
$ActiveWorksheet.Cells.Item($varJ,4).Font.Bold = "True"
$ActiveWorksheet.cells.item($varJ,5) = $varUserCount

# Hack to auto-fit the columns on the Excel document
#    Some reason it will write "True" to the screen
for ($i=1;$i -lt 5; $i++){
    $ActiveWorksheet.Cells.Item(1,$i).EntireColumn.Autofit()
}

# Save the document and quit
$ActiveWorkbook.SaveAs($LicenseReport)
$ExcelObject.Quit()

Start-Sleep -Seconds 5

# Email report 
$message = new-object System.Net.Mail.MailMessage 
$message.From = $fromaddress 
$message.To.Add($toaddress) 
$message.CC.Add($CCaddress) 
$message.IsBodyHtml = $True 
$message.Subject = $Subject 
$attach = new-object Net.Mail.Attachment($LicenseReport) 
$message.Attachments.Add($attach) 
$message.body = $body 
$smtp = new-object Net.Mail.SmtpClient($smtpserver) 
$smtp.Send($message) 

get-PsSession | Remove-PSSession