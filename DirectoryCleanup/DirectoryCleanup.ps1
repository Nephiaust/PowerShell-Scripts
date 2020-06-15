#Requires -Version 3.0

# **********************************************************
# **********************************************************
# *                                                        *
# * Script to clean up folders in a directory              *
# * Created by Nathaniel Mitchell                          *
# *                                                        *
# *               YYYY-MM-DD                               *
# * Version 0.1 - 2020-06-10 (Nathaniel Mitchell)          *
# *             * Initial creation of script               *
# * Version 1.0 - 2020-06-10 (Bill Couper)                 *
# *             * Corrected issue with JSON import *PS 4.0)*
# *                                                        *
# **********************************************************
# **********************************************************

# Set the location of the JSON file for the directory cleanup
$FILE_DirectoryCleanup = "C:\temp\DirectoryCleanup\DirectoryCleanup.json"

# ****************************************
# ****************************************
# ****************************************
# ***                                  ***
# ***  DO NOT EDIT BELOW THIS SECTION  ***
# ***                                  ***
# ****************************************
# ****************************************
# ****************************************
$VAR_CurrentLocation = Get-Location

$PSVersionMinimum = [version]'3.0.0.0'
if ($PSVersionMinimum -gt $PSVersionTable.PSVersion) { throw "This script requires PowerShell $PSVersionMinimum" }

#Temporarily create the variable
$OBJ_CurrentTime = Get-Date
$OBJ_OlderThanTime = $OBJ_CurrentTime

#Get the JSON for the process
$OBJ_DirectoryCleanup = Get-Content $FILE_DirectoryCleanup | Out-String | ConvertFrom-Json

ForEach ($OBJ_Directory in $OBJ_DirectoryCleanup) {

	$NewValue = $OBJ_Directory.Time * -1
	Switch ($OBJ_Directory.Unit) {
		"Years" {$OBJ_ReferenceTime = $OBJ_CurrentTime.AddYears($NewValue); break}
		"Months" {$OBJ_ReferenceTime = $OBJ_CurrentTime.AddMonths($NewValue); break}
		"Days" {$OBJ_ReferenceTime = $OBJ_CurrentTime.AddDays($NewValue); break}
		"Hours" {$OBJ_ReferenceTime = $OBJ_CurrentTime.AddHours($NewValue); break}
		"Minutes" {$OBJ_ReferenceTime = $OBJ_CurrentTime.AddMinutes($NewValue); break}
		"Seconds" {$OBJ_ReferenceTime = $OBJ_CurrentTime.AddSeconds($NewValue); break}
		"Milliseconds" {$OBJ_ReferenceTime = $OBJ_CurrentTime.AddMilliseconds($NewValue); break}
		default {$OBJ_ReferenceTime = $OBJ_CurrentTime.AddYears(-1); break}
	}
	Set-Location $OBJ_Directory.Directory
	
	If ([string]::IsNullOrWhiteSpace($OBJ_Directory.Excludes)) {
		$OBJ_Files = Get-ChildItem -recurse | Where-Object {$_.LastWriteTime -lt $OBJ_ReferenceTime}
	} else {
		$OBJ_Files = Get-ChildItem -recurse -exclude $OBJ_Directory.Excludes | Where-Object {$_.LastWriteTime -lt $OBJ_ReferenceTime}
	}
	
	$OBJ_ToDeleteDirectories = @()
	ForEach ($OBJ_File in $OBJ_Files) {
		If ($OBJ_File.PSIsContainer) {
			$OBJ_ToDeleteDirectories += $OBJ_File
		} else {
			If ($OBJ_Directory.Force -eq "Yes") {
				Remove-Item -LiteralPath $OBJ_File.FullName -force
			} else {
				Remove-Item -LiteralPath $OBJ_File.FullName 
			}
		}
	}
	ForEach ($OBJ_ToDeleteDirectory in $OBJ_ToDeleteDirectories) {
		Write-Host $OBJ_ToDeleteDirectory
		If ($OBJ_Directory.Force -eq "Yes") {
			Remove-Item -LiteralPath $OBJ_ToDeleteDirectory.FullName -force
		} else {
			Remove-Item -LiteralPath $OBJ_ToDeleteDirectory.FullName
		}
	}
	Remove-Variable OBJ_ToDeleteDirectories
}
Set-Location $VAR_CurrentLocation