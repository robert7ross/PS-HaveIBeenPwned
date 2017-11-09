# AUTHOR: Robert Ross
# LASTEDIT: 20171108
# KEYWORDS: HaveIBeenPwned
# DESCRIPTION: A simple script for ingesting data from the HaveIBeenPwned project
#   and enriching it with data from Active Directory
# LICENSE: MIT License, Copyright (c) 2017 Robert Ross

$BreachedAccounts = Import-Excel 'C:\accounts.xlsx' -WorksheetName 'Breached email accounts'

$Output = @()
foreach($Account in $BreachedAccounts){
	try{
		$Output += $AccountInfo = Get-ADUser $Account.Email.split("@")[0]
	}
	catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]{
	}
}

#CSV Formatted output
#$Output | Export-Csv 'C:\enriched.csv' -NoTypeInformation

#Excel format (requires the ImportExcel module)
$ExcelParams = @{
    Path = "C:\enriched.xlsx"
    BoldTopRow = $true
}

Remove-Item -Path $ExcelParams.Path -Force -EA Ignore

$Output| Export-Excel @ExcelParams -WorkSheetname Compromised -NoNumberConversion *