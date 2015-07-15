		<######################################################################
		365 Powershell Administration System 
		Copyright (C) 2015  Ashley Unwin, www.AshleyUnwin.com/powershell
		
		It is requested that you leave this notice in place when using the
		Menu System.

		This program is free software: you can redistribute it and/or modify
		it under the terms of the GNU General Public License as published by
		the Free Software Foundation, either version 3 of the License, or
		the latest version.

		This program is distributed in the hope that it will be useful,
		but WITHOUT ANY WARRANTY; without even the implied warranty of
		MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
		GNU General Public License for more details.

		You should have received a copy of the GNU General Public License
		along with this program.  If not, see <http://www.gnu.org/licenses/>.
		
		######################################################################
		Known Bugs:
		- Cannot accept company names with space - Cause: Line 61 $xMenuHash.add($_.Company,"fSetupCompany -xCompany "+$_.company) - Resolution: 
		- Cannot Switch company by re-running qqq
		- Rename Account/email
		- Remove Account
		- fEditUserAccountName might not change the name of the mailbox itself
		- hide from gal
		- external auth
		######################################################################>


function global:start-login{
	#This script requires the Multi Layered Dynamic Menu System Module from www.AshleyUnwin.com/Powershell_Multi_Layered_Dynamic_Menu_System

	if (get-module -name MenuSystem){}else{
		$source = "https://raw.githubusercontent.com/manicd/Powershell-Multi-Layered-Dynamic-Menu-System/master/MenuSystem.psm1"
		$destination = ".\MenuSystem.psm1"
		Invoke-WebRequest $source -OutFile $destination
		#$destination = "Z:\~Tools\Powershell\MenuSystem.psm1" #Temp Line to test MenuSystem Edit
		Import-Module $destination
	}
	if (get-module -name MenuSystem) {} else {
		fDisplayInfo -xText "MenuSystem Module not avalible, unable to contuinue" -xColor "red" -xTime 3
		Return $false
	}
	Import-Module MSOnline
	if (get-module -name MSOnline) {} else {
		fDisplayInfo -xText "MS Online Module not avalible, unable to contuinue" -xColor "red" -xTime 3
		Return $false
	}
	fclear-login
	cls
	fLoginMenu
}

function global:fLoginMenu{
	# Requires a CSV File in with the columns company,adminuser
	if (test-path Z:\~Tools\Powershell\companys.csv) {	
		$global:csv = import-csv Z:\~Tools\Powershell\companys.csv
	} else {
		if (test-path c:\PowerShell\companys.csv) {
			$global:csv = import-csv C:\PowerShell\companys.csv
		}
	}
	
	#Create Hash Table Object	
	$xMenuHash = New-Object System.Collections.HashTable
	#Create Menu Structure Hash Table
	$global:csv | sort -property company | foreach-object {
			$xMenuHash.add($_.Company,"fSetupCompany -xCompany "+$_.company)
		}
	#Call the Menu and pass the Hash Table and Title, Return the Selected Company, User and Pass
	$xReturn = Use-Menu -MenuHash $xMenuHash -Title "365 Login Menu" -NoSplash 1
	[string]$xCompany = $xReturn.xCompany
	#If the Pass is Set, run the Login Script
	if ($xReturn.xPass)	{fLoginto365 -xCompany $xCompany -xPass $xReturn.xPass -xAdminUser $xReturn.xAdminUser}
}

function global:fSetupCompany{

PARAM(
[string]$xCompany 
)
	$xAdminUser = $global:csv | where-object {$_.company -eq $xCompany} | select adminuser
	$xAdminUser = $xAdminUser.adminuser
	$global:xCompany = $global:csv | where-object {$_.company -eq $xCompany} | select company
	$global:xCompany = $global:xCompany.company
	$global:xDomain = $global:csv | where-object {$_.company -eq $xCompany} | select domain
	$global:xDomain = $global:xDomain.domain
	
	$passfile = "c:\O365\" + $global:xCompany + "365pass.txt"
	if (test-path $passfile) {
		} else {
		$string = Read-Host "Enter the Password"
		cls 
		if (test-path c:\O365) {} Else {
			new-item C:\O365 -type Directory
		}
		fcreate-sstring $string | out-file $passfile 
		if (test-path $passfile){
			} else {
			return
			} 
		}
	$xPass = Get-Content $passfile

	[hashtable]$xReturn = @{} 
	$xReturn.add("xAdminUser", $xAdminUser)
	$xReturn.add("xPass",$xPass)
	$xReturn.add("xCompany",$xCompany)
	return $xReturn
}

function global:fLoginTo365{

PARAM(
[string]$xAdminUser,
[string]$xPass,
[string]$xCompany
)

	# If username has been set, login
    if ($xPass)	{
		Write-host "Connecting to"$xCompany -Fore Green
		Write-host "Creating Credential Object" -Fore Green
		$O365Cred=New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $xAdminUser, ($xPass | ConvertTo-SecureString)
		Write-host "Creating Session Object" -Fore Green
		$O365Session = New-PSSession –ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $O365Cred -Authentication Basic -AllowRedirection
		write-host "Importing Session" -Fore Green
		Import-PSSession $O365Session
		write-host "Connecting to MSOL Service" -Fore Green
		Connect-MsolService –Credential $O365Cred
		cls
		write-host "`nYou are now logged in to"$xCompany". Type 'use-admin' to access the menu." -Fore Green
	}else{
		write-host "`n`tNo Account Set - Not Attempting Login to 365 `n" -Fore Red
	}
	return
}

function global:fcreate-sstring{

PARAM(
[STRING[]]$text = "Test String"
)
	$text | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString
	return
	
}	

function global:fclear-login {
	Import-Module MSOnline
    if (Get-PSSession) {remove-pssession *}
    if ($user){$global:user = $false}
    if ($pass){$global:pass = $false}
    if ($company){$global:company = $false}
	write-host "365 Logins Cleared"
    }
	
function global:clear-passwords{
	Remove-item C:\O365\* -confirm
}	

function global:qqq{start-login}

function global:www{use-admin}

function global:fUserPrompt {
PARAM(
[parameter(Mandatory=$True,
Position=0,
HelpMessage='Enter a String value to display as a question')]
[string]$xQuestion,
[parameter(Position=1,
HelpMessage='Enter a String value to display as the Prompt')]
$xPrompt = "=>"
)
	$count = $measureObject = $xQuestion | Measure-Object -Character | select Characters
	$count = $count.Characters
	$i=0
	while ($i -lt ($count + 18)) {
		[string]$xStars = $xStars+"*"
		$i++
		}
	Write-Host
	Write-Host $xStars -Fore Green 
	Write-Host "`t"$xQuestion -Fore Yellow
	Write-Host $xStars -Fore Green
	Write-Host
	[string]$xAnswer = Read-Host $xPrompt
	Write-Host
	Return $xAnswer
}

function global:fDisplayInfo {

PARAM(
[parameter(Mandatory=$True,
Position=0,
HelpMessage='Enter a String value to display as info')]
[string]$xText,
[parameter(Position=1,
HelpMessage='Enter the colour of the main text')]
[string]$xColor = "Cyan",
[parameter(Position=2,
HelpMessage='Time for message to be displayed')]
[int]$xTime = 1,
[string]$xText2,
[string]$xColor2 = "Cyan",
[string]$xText3,
[string]$xColor3 = "Cyan"
)
$count = $measureObject = $xText | Measure-Object -Character | select Characters
$count = $count.Characters
$i=0
while ($i -lt ($count + 18)) {[string]$xStars = $xStars+"*"; $i++}
Write-Host
Write-Host $xStars -Fore Green 
Write-Host "`t"$xText -Fore $xColor
if ($xText2) {Write-Host "`t"$xText2 -Fore $xColor2}
if ($xText3) {Write-Host "`t"$xText3 -Fore $xColor3}
Write-Host $xStars -Fore Green
Write-Host
start-sleep -s $xTime
Return
}

function global:Use-Admin {
	if (get-module -name MenuSystem){
	}elseif (Test-Path c:\powershell\MenuSystem.psm1) {
		Import-Module c:\powershell\MenuSystem.psm1
	}else{
		Import-Module Z:\~Tools\Powershell\MenuSystem.psm1
	}

	[bool]$global:UseAdminLoaded=$true

	$global:MenuHash2=@{ "Users"=@{		"Password Reset"="fResetUserPasswords"
										"Licence Names"="fGetMsolAccountSku"
										"New User"="fAddNewUser"
										"List Users"="fListUsers"
										"Edit User Account Name"="fEditUserAccountName"
										}
			"Mailboxes"=@{				"Grant User Access to Mailbox Folder"="fAddMailboxFolderPerm"
										"Add Email Alias"="fAddUserEmailAlias"
										"Remove User Access from Mailbox Folder"="fRemoveMailboxFolderPerm"
										"Grant Full Access to Mailbox"="fGrantFullAccessMailbox"
										"Remove Full Access to Mailbox"="fRemoveFullAccessMailbox"
										"List Mailboxes"="fListMailboxes"
										"List Mailbox Statistics"="fListMailboxStats"
										"List Email Forwarding Status"="fCheckForwarding"
										"Disable Access to Services"=@{	"Disable Outlook Anywhere Access"="fDisableOutlookAnywhere"
																		"Disable OWA Access"="fDisableOWA"
																		"Disable IMAP Access"="fDisableImap"
																		"Disable POP Access"="fDisablePop"
																		}
										"Hide/Unhide from GAL"="fToggleHideFromGAL"
										}
			"Dist Groups"=@{			"List Dist Groups and Members"="fListDistMembers"
										"Add User to Dist Group"="fadduserdistgroup"
										"Remove User from Dist Group" = "fremoveuserdistgroup"
										"Add New Dist Group"="fAddNewDistGroup"
										"Add Email Alias to Dist Group"="fAddDistGroupEmailAlias"}
			"MSOnline Org"=@{			"List Partner Information"="fViewPartnerInfo"
										"List Domain Info"="fVeiwDomain"
										"List Licencing Status"="fGetMsolAccountSku"
										}
			"X-Experimental Function"="fExperimentalFunction"							
										
			}
	$global:title="Office 365 Menu"	
	[bool]$global:quitmenu = $false
	[bool]$global:xProceed = $false
	while ($global:quitmenu -ne $true) {
		if ($global:xProceed -ne $true) {
				Use-Menu -MenuHash $MenuHash2 -Title $title -NoSplash 1
				$global:xProceed = $true
			} else {
				Use-Menu -MenuHash $MenuHash2 -Title $title -NoSplash 1 -SelectionHist $SelectionHist -xContinue $true
			}
		
	}
}

function global:fStoreMainMenu {
PARAM(
[bool]$xRestore
)
	if ($xRestore) {
		$global:SelectionHist = $xTempSelectionHist
		$global:Quit = $xTempQuit
	} else {
		$xTempSelectionHist = $global:SelectionHist
		$xTempQuit = $global:Quit
	}
}

function global:fCheckUPN {
PARAM(
[string]$xUPN,
[bool]$xCurrent
)
#With just xUPN - checks if xUPN is a valid email format and valid domain
#With xUPN and xCurrent  - Checks if xUPN is valid email, valid domain and is a existing UPN
#returns true if valid
#returns false if invalid

	if (($xUPN -like "*@*") -AND ($xUPN -like "*.*")) {
		$pos = $xUPN.IndexOf("@")
		$Dom = $xUPN.Substring($pos+1)
		get-msoldomain | foreach-object {
			if ($Dom -eq $_.name) {
				if ($xCurrent) {
					get-msoluser | foreach-object {
						if ($_.UserPrincipalName -eq $xUPN) {
							return $true
						}
					}
				} else {
					return $true
				}
			}
		}
	}
	return $false
}

function global:fCollectUPN {
PARAM(
[string]$xText,
[bool]$xCurrent = $false
)
	while (!$xVar) {
		$xInput = fUserPrompt -xQuestion $xText+" (Type 'QUIT' to exit)"
		if (fCheckUPN -xUPN $xInput -xCurrent $xCurrent) {
			$xVar = $xInput
			remove-variable -name xInput
		} elseif ($xInput -eq "quit") {
			Return $false
			# You must include if ($Var -eq $false) {Return $false} after calling the function to fully quit the function
		} else 	{
			fDisplayInfo -xText "Invalid Selection" -xColor "red"
		}
	}
	Return $xVar
}

function global:fCheckIdentity {
PARAM(
[string]$id
)
#Function to check if an identity specified exists
	if (Get-Mailbox -identity $id -ErrorAction 'silentlycontinue') {
		Return $true
	} elseif (Get-DistributionGroup -identity $id -ErrorAction 'silentlycontinue') {
		Return $true
	} elseif (Get-Contact -identity $id -ErrorAction 'silentlycontinue') {
		Return $true
	} 
	Return $false	
}

function global:fCollectIdentity {
PARAM(
[string]$xText
)
	while (!$xVar) {
		$xInput = fUserPrompt -xQuestion $xText" (Type 'QUIT' to exit)"
		if (fCheckIdentity -Id $xInput) {
			$xVar = $xInput
			remove-variable -name xInput
		} elseif ($xInput -eq "quit") {
			Return $false
			# You must include if ($Var -eq $false) {Return $false} after calling the function to fully quit the function
		} else 	{
			fDisplayInfo -xText "Invalid Selection" -xColor "red"
		}
	}
	Return $xVar
}


# Below this line are the functions called by the menu values



#Users =======================================================================================
function global:fResetUserPasswords {
	fStoreMainMenu -xRestore 0
	#Create a function to actually change the password
	function global:fResetUserPasswordsCollectPass {
	PARAM(
	[string]$xUser
	)
		$xString =  "Please enter the new password for "+$xUser+" or type [quit] to quit."
		$xPass = fUserPrompt -xQuestion $xString -xPrompt "Password"
		if ($xPass -ne "quit") {
			if (($xPass -ne "") -AND ($xPass -ne $null)) {
				write-host (DisplayInfo -xText "Setting New Password")
				Write-host
				$xPass = Set-MsolUserPassword -UserPrincipalName $xUser -ForceChangePassword $false -NewPassword $xPass
				write-host (fDisplayInfo -xText "Password now set" -xColor "Red" -xTime 3)
				$xPass = $null	
				Cls
				} else {
				write-host (fDisplayInfo -xText "Password not entered, Nothing has been changed." -xColor "Red")
				fResetUserPasswordsCollectPass -xUser $xUser				
			}
		}else{
		cls
		write-host (fDisplayInfo -xText "Quitting....Nothing has been changed." -xColor "Red" -xTime 3)
		Cls
		}
	}

	#Create the Menu Hash Table Object	
	$xMenuHash = New-Object System.Collections.HashTable
	#Create Menu Structure Hash Table and set values to be function with UPN as input
	get-msoluser | sort-object UserPrincipalName | select UserPrincipalName, FirstName, Lastname | foreach-object {
			$xMenuHash.add($_.UserPrincipalName+" - "+$_.FirstName+" "+$_.LastName,"fResetUserPasswordsCollectPass -xUser "+$_.UserPrincipalName)
		}
	#Call the Menu	
	use-menu -MenuHash $xMenuHash -Title "Reset User Password" -NoSplash $True
	fStoreMainMenu -xRestore 1
	
}

function global:fAddNewUser {
	fStoreMainMenu -xRestore 0
	$xFirstName = fUserPrompt -xQuestion "First Name"
	$xLastName = fUserPrompt -xQuestion "Last Name"
	$xUPN = fCollectUPN -xText "User Principal Name:"
	if ($xUPN -eq $false) {Return $false}
	$xLicMenu = New-Object System.Collections.HashTable
	Get-MsolAccountSku | sort-object AccountSkuId | select AccountSkuId | foreach-object {
			$xLicMenu.add($_.AccountSkuId,"write-output "+$_.AccountSkuId)
		}
	$xLic  = use-menu -MenuHash $xLicMenu -Title "Select a Licence" -NoSplash 1
		
	$xPass = fUserPrompt -xQuestion "Password"

	$xDisplayName = $xFirstName+" "+$xLastname
	
	Cls
	
	function fAddNewUserLicCheck {
		$xSKU = Get-MsolAccountSku | where-object { $_.AccountSkuId -eq $xLic } 
		
		if ($xSKU.ConsumedUnits -le $xSKU.ActiveUnits) {
			write-host (New-MsolUser -DisplayName $xDisplayName -FirstName $xFirstName -LastName $xLastName -UserPrincipalName $xUPN -LicenseAssignment $xLic -Password $xPass -UsageLocation GB -PreferredLanguage "en-GB" -ForceChangePassword $False | format-table | out-string)
			pause
			fStoreMainMenu -xRestore 1
			return
		} else {
			fDisplayInfo -xText "You Do Not currently have enough Licenses to proceed." -xText2 "Please Login to Office 365 and purchase more licences" -xText3 "before proceeding." -xTime 5
			$xTryAgain = fUserPrompt -xQuestion "Try Again? (y/n)"
			if ($xTryAgain -eq "n") {
				fStoreMainMenu -xRestore 1
				return
			} else {
				fAddNewUserLicCheck
				fStoreMainMenu -xRestore 1
				return
			}
		}
	}
	fAddNewUserLicCheck
	fStoreMainMenu -xRestore 1
}

function global:fListUsers {
	write-host (get-msoluser | format-table | out-string)
}

function global:fEditUserAccountName {
		
	$xOldUPN = fCollectUPN -xText "Enter Old User UPN: (Type 'QUIT' to exit)"
	if ($xOldUPN -eq $false) {Return $false}	
	
	$xNewUPN = fCollectUPN -xText "Enter New User UPN: (Type 'QUIT' to exit)"
	if ($xNewUPN -eq $false) {Return $false}	

	$xNewFirstName = fUserPrompt -xQuestion "What is the New Users First Name"
	$xNewLastName = fUserPrompt -xQuestion "What is the New Users Last Name"
	$xNewDisplayName = $xNewFirstName+" "+$xNewLastName

	set-msoluserprincipalname -UserPrincipalName $xOldUPN -NewUserPrincipalName $xNewUPN
	set-msoluser -UserPrincipalName $xNewUPN -Firstname $xNewUserName -LastName $xNewLastName -DisplayName $xNewDisplayName
	write-host (get-msoluser -UserPrincipalName $xNewUPN | fl UserPrincipalName, FirstName, LastName, ProxyAddresses | out-string)
	#This might not rename mailbox - investigate
}




#Mailboxes =======================================================================================
function global:fAddMailboxFolderPerm {
	fStoreMainMenu -xRestore 0
	cls
	$xUser = fCollectIdentity -xText "Enter the User who would like the access"
	if ($xUser -eq $false) {Return $false}
	$xMailbox = fCollectIdentity -xText "Enter the Mailbox they would like access to"
	if ($xMailbox -eq $false) {Return $false}
	$xFolder = fUserPrompt -xQuestion "Enter the Folder the would like access to"
	$xLevelMenuHash = @{ 
		"A_Owner" = "Write-Output Owner"
		"B_PublishingEditor" = "Write-Output PublishingEditor"
		"C_Editor" = "Write-Output Editor"
		"D_PublishingAuthor" = "Write-Output PublishingAuthor"
		"E_Author" = "Write-Output Author"
		"F_NonEditingAuthor" = "Write-Output NonEditingAuthor"
		"G_Reviewer" = "Write-Output Reviewer"
		"H_Contributor" = "Write-Output Contributor"
		"I_AvailabilityOnly" = "Write-Output AvailabilityOnly"
		"J_LimitedDetails" = "Write-Output LimitedDetails"
	}
	$xLevel = Use-Menu -MenuHash $xLevelMenuHash -NoSplash 1 -Title "Choose a Permissions Level"

	$xIdString = $xMailbox+":\"+$xFolder
	cls
	$xTextString = "Granting "+$xLevel+" to "+$xUser+" for "+$xFolder+" in "+$xMailbox+"'s Mailbox" 
	fDisplayInfo -xText $xTextString

	write-host (Add-MailboxFolderPermission -Identity $xIdString -User $xUser -AccessRight $xLevel | format-table | out-string)
	pause
	fStoreMainMenu -xRestore 1
}

function global:fGrantFullAccessMailbox {
	
	$xUser = fCollectIdentity -xText "Enter the User who would like the access"
	if ($xUser -eq $false) {Return $false}
	$xMailbox = fCollectIdentity -xText "Enter the Mailbox they would like access to"
	if ($xMailbox -eq $false) {Return $false}
	$xAutoMapYN = fUserPrompt -xQuestion "Would you like to enable AutoMapping? (y/n)"
	if ($xAutoMapYN -eq "y") {
		$xAutoMap = $true
	}elseif ($xAutoMapYN -eq "n") {
		$xAutoMap = $false
	}else{
		#Default
		$xAutoMap = $true
	}
	
	Add-MailboxPermission -identity $xMailbox -User $xUser -AccessRight fullaccess -InheritanceType all -Automapping $xAutoMap
	write-host (Get-MailboxPermission -identity $xMailbox | format-table | out-string)
	pause
}

function global:fRemoveFullAccessMailbox {

	$xUser = fCollectIdentity -xText "Enter the User who no longer requires the access"
	if ($xUser -eq $false) {Return $false}
	$xMailbox = fCollectIdentity -xText "Enter the Mailbox they no longer need"
	if ($xMailbox -eq $false) {Return $false}
	
	Remove-MailboxPermission -identity $xMailbox -User $xUser -AccessRight fullaccess
	write-host (Get-MailboxPermission -identity $xMailbox | format-table | out-string)
	pause
}

function global:fListMailboxes {
	write-host (get-mailbox | select DisplayName, Alias, UserPrincipalName, PrimarySmtpAddress | format-table | out-string)
	pause
}

function global:fListMailboxStats {
	write-host (get-mailbox | foreach-object { get-mailboxstatistics -identity $_.Identity | select DisplayName, TotalItemSize, LastLogonTime }  | format-table | out-string)
	pause
}

function global:fCheckForwarding {
	write-host (get-mailbox | select DisplayName, PrimarySMTPAddress, forwardingaddress, forwardingsmtpaddress, DeliverToMailboxAndForward | format-table | out-string)
	pause
}

function global:fAddUserEmailAlias {

	$xUser = fCollectIdentity -xText "Enter User ID:"
	if ($xUser -eq $false) {Return $false}

	$xVar = get-mailbox -id $xUser
	$xEmails = $xVar.EmailAddresses
	$i = 1
	fDisplayInfo -xText "The current emails attached to this user are"
	foreach ($email in $xEmails) {
		Write-Host "`t`t"$i" - "$email; $i++
	}
	write-host
	
	$xNewEmailAddress = fCollectUPN -xText "Enter the New Email Address to add:"
	if ($xNewEmailAddress -eq $false) {Return $false}

	$xNewEmailAddress = "smtp:"+$xNewEmailAddress
	Set-Mailbox -id $xUPN -emailAddresses @{Add=$xNewEmailAddress}

	$xVar = get-mailbox -id $xUser
	$xEmails = $xVar.EmailAddresses
	$i = 1
	fDisplayInfo -xText "The current emails attached to this user are"
	foreach ($email in $xEmails) {
		Write-Host "`t`t"$i" - "$email 
		$i++
	}
	write-host
	pause
}

function global:fToggleHideFromGAL {
	$xUser = fCollectIdentity -xText "Enter the User who would like to hide"
	if ($xUser -eq $false) {Return $false}
	$xStatus = Get-mailbox -identity $xUser | select HiddenFromAddressListsEnabled
	if ($xStatus.HiddenFromAddressListsEnabled) {
		$xUnhide = "n"
		while ($xUnhide -ne "y") {
			$xUnhide = fUserPrompt -xQuestion "Would you like to unhide the mailbox? (y/n)"
		}
		if ($xUnHide -eq "y") {$xHidden = $false}
	} else {
		$xhide = "n"
		$xhide = fUserPrompt -xQuestion "Would you like to hide the mailbox? (y/n)"
		if ($xHide -eq "y") {$xHidden = $true}
	}
	Set-mailbox -identity $xUser -HiddenFromAddressListsEnabled $xHidden
	write-host ( get-mailbox -identity $xUser | select Displayname, HiddenFromAddressListsEnabled | format-list | out-string )
	pause
}

#Mailbox Services =======================================================================================

function global:fDisableOutlookAnywhere {
 	fDisplayInfo -text "Disable Outlook Anywhere"
	while (!$xIdentity) {
		$xInput = fUserPrompt -xQuestion "Enter User ID. (Type 'ALL' to set globally or type 'QUIT' to exit)"
		if ((fCheckIdentity -id $xInput) -OR ($xInput -eq "all")) {
			$xIdentity = $xInput
			remove-variable -name xInput
		} elseif ($xInput -eq "quit") {
			Return
		} else 	{
			fDisplayInfo -Text "Invalid Selection" -xColour "red"
		}
	}
	if ($xIdentity -eq "all") {
		Get-Mailbox -ResultSize Unlimited | Set-CASMailbox -MAPIEnabled $False
	} else {
		Set-CASMailbox -identity $xIdentity -MAPIEnabled $False
	}
	pause	
}

function global:fDisableOWA {
 	fDisplayInfo -text "Disable Outlook Web Access"
	
	$xIdentity = fCollectIdentity -xText "Enter User ID: (Type 'All' to select globally)"
	if ($xIdentity -eq $false) {Return $false}
	
	if ($xIdentity -eq "all") {
		Get-Mailbox -ResultSize Unlimited | Set-CASMailbox -ActiveSyncEnabled $False
		write-host (get-casmailbox | format-list | out-string) 
	} else {
		Set-CASMailbox -identity $xIdentity -ActiveSyncEnabled $False
		write-host (get-casmailbox -identity $xIdentity | format-list | out-string) 
	}	
	pause
}

function global:fDisableImap {
 	fDisplayInfo -text "Disable IMAP Access"
	while (!$xIdentity) {
		$xInput = fUserPrompt -xQuestion "Enter User ID. (Type 'ALL' to set globally or type 'QUIT' to exit)"
		if ((fCheckIdentity -id $xInput) -OR ($xInput -eq "all")) {
			$xIdentity = $xInput
			remove-variable -name xInput
		} elseif ($xInput -eq "quit") {
			Return
		} else 	{
			fDisplayInfo -Text "Invalid Selection" -xColour "red"
		}
	}
	if ($xIdentity -eq "all") {
		Get-Mailbox -ResultSize Unlimited | Set-CASMailbox -imapenabled $False 
		write-host (get-casmailbox | format-list | out-string) 
	} else {
		Set-CASMailbox -identity $xIdentity -imapenabled $False 
		write-host (get-casmailbox -identity $xIdentity | format-list | out-string) 
	}	
	pause
}

function global:fDisablePop {
 	fDisplayInfo -text "Disable POP Access"
	while (!$xIdentity) {
		$xInput = fUserPrompt -xQuestion "Enter User ID. (Type 'ALL' to set globally or type 'QUIT' to exit)"
		if ((fCheckIdentity -id $xInput) -OR ($xInput -eq "all")) {
			$xIdentity = $xInput
			remove-variable -name xInput
		} elseif ($xInput -eq "quit") {
			Return
		} else 	{
			fDisplayInfo -Text "Invalid Selection" -xColour "red"
		}
	}
	if ($xIdentity -eq "all") {
		Get-Mailbox -ResultSize Unlimited | Set-CASMailbox -popenabled $False
		write-host (get-casmailbox | format-list | out-string)
	} else {
		Set-CASMailbox -identity $xIdentity -popenabled $False
		write-host (get-casmailbox -identity $xIdentity | format-list | out-string)
	}	
	pause
}



#Dist Groups =======================================================================================
function global:fListDistMembers {
PARAM(
[string]$xGroupName
)
	if (!$xGroupName) {
		Get-DistributionGroup | sort DisplayName | foreach-object {
			Write-host $($_.Displayname)"`n===========" 
			Get-DistributionGroupMember $($_.DisplayName) | foreach-object {
				write-host $_.DisplayName
			}
					write-host "`n" 
		}
	}else{
		write-host $xGroupName"`n==========="
		write-host (Get-DistributionGroupMember $xGroupName | select DisplayName | format-table | out-string)
	}	
	pause
}

function global:fRemoveMailboxFolderPerm {
	cls
	$xUser = fCollectIdentity -xText "Enter the User who would like the access"
	if ($xUser -eq $false) {Return $false}
	$xMailbox = fCollectIdentity -xText "Enter the Mailbox they would like access to"
	if ($xMailbox -eq $false) {Return $false}
	$xFolder = fUserPrompt -xQuestion "Enter the Folder the would no longer like access to"
	$xIdString = $xMailbox+":\"+$xFolder
	cls
	$xTextString = "Removing "+$xUser+" from "+$xFolder+" in "+$xMailbox+"'s Mailbox" 
	fDisplayInfo -xText $xTextString
	Remove-MailboxFolderPermission -Identity $xIdString -User $xUser
	write-host (get-MailboxFolderPermission -Identity $xIdString -User $xUser | Format-Table | out-string)
	pause
}

function global:fAddUserDistGroup {
	fStoreMainMenu -xRestore 0
	$xDistMenuHash = New-Object System.Collections.HashTable
	Get-DistributionGroup | sort-object Name | select Name | foreach-object {
			$xDistMenuHash.add($_.Name,"write-output "+$_.Name)
		}
	[string]$xGroupName  = use-menu -MenuHash $xDistMenuHash -Title "Select a group" -NoSplash 1
	
	$xAdd = "y"
	while ($xAdd -eq "y") {
		$xMember = fCollectIdentity -xText "Who would you like to add:"
		if ($xMember -ne $false) {
			Add-DistributionGroupMember $xGroupName -Member $xMember -BypassSecurityGroupManagerCheck
		} else { 
			fDisplayInfo -xText "Quitting" 
			fStoreMainMenu -xRestore 1
			pause
			Return $false
		}
		$xAdd = fUserPrompt -xQuestion "Would you like to add another? (y/n)"
	}
	fListDistMembers -xGroupName $xGroupName
	pause
	fStoreMainMenu -xRestore 1
}

function global:fRemoveUserDistGroup {
	fStoreMainMenu -xRestore 0
	$xDistMenuHash = New-Object System.Collections.HashTable
	Get-DistributionGroup | sort-object Name | select Name | foreach-object {
			$xDistMenuHash.add($_.Name,"write-output "+$_.Name)
		}
	[string]$xGroupName  = use-menu -MenuHash $xDistMenuHash -Title "Select a group" -NoSplash 1
	#$xMember  = fUserPrompt -xQuestion "Who would you like to remove"
	
	$xDistMemMenuHash = New-Object System.Collections.HashTable
	Get-DistributionGroupMember $xGroupName | sort-object Name | foreach-object {
			$xDistMemMenuHash.add($_.Name,"write-output "+$_.Name)
		}
	[string]$xMember  = use-menu -MenuHash $xDistMemMenuHash -Title "Select a Member" -NoSplash 1
	
	Remove-DistributionGroupMember $xGroupName -Member $xMember -BypassSecurityGroupManagerCheck
	fListDistMembers -xGroupName $xGroupName
	fStoreMainMenu -xRestore 1
	pause
}

function global:fAddNewDistGroup {
	fStoreMainMenu -xRestore 0
	$xgroupname = fUserPrompt -xQuestion "Enter the Alias: "  
		
	function global:fNewDistGroup {
	PARAM(
	[string]$xGroupName,
	[string]$xSelectedDomain
	)
		New-DistributionGroup -Name $xGroupName -DisplayName $xgroupname -Alias $xgroupname -PrimarySmtpAddress $xgroupname"@"$xSelectedDomain  
		Set-DistributionGroup -Identity $xGroupName -RequireSenderAuthenticationEnabled $false -HiddenFromAddressListsEnable $false 
	}
		
	#Create the Menu Hash Table Object	
	$xMenuHash = New-Object System.Collections.HashTable
	#Create Menu Structure Hash Table and set values to be function with UPN as input
	get-msoldomain | select name | sort-object Name | foreach-object {
			$xMenuHash.add($_.Name,"fNewDistGroup -xGroupName "+$xGroupName+" -xSelectedDomain "+$_.Name)
		}
	#Call the Menu	
	use-menu -MenuHash $xMenuHash -Title "Select Domain" -NoSplash $True
	
	Write-Host (get-DistributionGroup -Identity $xGroupName | format-table | out-string)
	fStoreMainMenu -xRestore 1
	pause
}

function global:fAddDistGroupEmailAlias {

	$xGroupName = fCollectIdentity -xText "Enter Group Name:"
	if ($xGroupName -eq $false) {Return $false}
	if ($xGroupName -eq $false) {
		Return
	}
	
	$xVar = get-distributiongroup -identity $xGroupName 
	$xEmails = $xVar.EmailAddresses
	$i = 1
	fDisplayInfo -xText "The current emails attached to this group are"
	foreach ($email in $xEmails) {
		Write-Host "`t`t"$i" - "$email
	}
	write-host 

	
	$xNewEmailAddress = fCollectUPN -xText "Enter the additional email address" -xCurrent $false
	if ($xNewEmailAddress -eq $false) {Return}
	
	$xNewEmailAddress = fuserPrompt -xQuestion "Enter the additional email address"
	Set-DistributionGroup $xGroupName -emailaddresses @{Add=$xNewEmailAddress}
	
	$xVar = get-distributiongroup -identity $xGroupName 
	$xEmails = $xVar.EmailAddresses
	$i = 1
	fDisplayInfo -xText "The current emails attached to this group are now"
	foreach ($email in $xEmails) {
		Write-Host "`t`t"$i" - "$email
	}
	pause
}



#Organisation =======================================================================================
	
function global:fViewPartnerInfo {
	Write-Host (Get-MsolPartnerInformation | Format-List | Out-String)
	pause 
}

function global:fVeiwDomain {
	Write-Host (Get-MsolDomain | Format-Table | Out-String)
	pause 
}

function global:fGetMsolAccountSku {
	write-host (Get-MsolAccountSku | Format-Table | Out-String)
	pause
}


# Other functions =======================================================================================

function global:fExperimentalFunction{

}
