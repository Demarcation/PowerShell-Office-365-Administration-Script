		<######################################################################
		365 Powershell Administration System 
		Copyright (C) 2019  Ashley Unwin, www.AshleyUnwin.com/powershell
		
		It is requested that you leave this notice in place when using the
		Menu System.

		This work is licensed under the Creative Commons Attribution-ShareAlike
		4.0 International License. To view a copy of this license, 
		visit http://creativecommons.org/licenses/by-sa/4.0/.
		
		Author: Ashley Unwin
		Website: http://www.ashleyunwin.com/powershell-office-365-admin-script/
		
		######################################################################
		Known Bugs and Feature Requests:
		- BUG STATUS-CONFIRMED: Cannot accept company names with space - Cause: Line 61 $xMenuHash.add($_.Company,"fSetupCompany -xCompany "+$_.company) - Resolution: 
		- BUG STATUS-CONFIRMED: fSetDefaultEmailAlias does not work
		
		- BUG STATUS-UNKNOWN: fEditUserAccountName might not change the name of the mailbox itself
		
		- FEATURE-COMPLETE: List all mailboxes a user has access to
		- FEATURE-COMPLETE: Added Send-As to Faull Access on all mailboxes
		- FEATURE-COMPLETE: Added Warning to Faull Access on all mailboxes
		- FEATURE-COMPLETE: Message Trace Sender in Mail Flow and Spam

		- FEATURE-REQUEST: Rename Account/email
		- FEATURE-REQUEST: Remove Account
		- FEATURE-REQUEST: Remove Dis Group
		- FEATURE-REQUEST: List all emails - Get-Recipient | Select Name -ExpandProperty EmailAddresses | Select Name,  SmtpAddress
		- FEATURE-REQUEST: Find Specfic email - Get-Recipient | Select Name -ExpandProperty Emailaddresses | ?{ $_.Smtpaddress -eq $xFindAddress} | select name, smtpaddress
		- Feature-REQUEST: make exchange on-site compatible
		######################################################################>


#Layout the Menu ================================================================
		
$global:MenuHash2=@{ "Users"=@{		"Password Reset"="fResetUserPasswords"
								"New User"="fAddNewUser"
								"List Users (All)"="fListUsers"
								"List Users (Licensed)"="fListLicensedUsers"
								"Edit User Account Name"="fEditUserAccountName"
								"Set Password Never Expires (All)"="fSetPasswordNeverExpireAll"
								}
	"Mailboxes"=@{				"Mailbox Permissions"=@{
															"Grant Full Access to SINGLE Mailbox"="fGrantFullAccessMailbox"
															"Remove Full Access from SINGLE Mailbox"="fRemoveFullAccessMailbox"
															"Grant Full Access to ALL Mailboxes"="fGrantFullAccessMailboxAllMailboxes"
															"Show Permissions on a Mailbox"="fShowMailboxPerms"
															"Find Mailboxes User can Access"="fFindUserPermissions"
															"Add Send As Permissions"="fGrantSendAsPerms"
															"Folder Access"=@{
																				"Grant User Access to Mailbox Folder"="fAddMailboxFolderPerm"
																				"Remove User Access from Mailbox Folder"="fRemoveMailboxFolderPerm"
																				}
														}
								"List Mailboxes (All)"="fListMailboxes"
								"List Mailboxes (User Only)"="fListUserMailboxes"
								"List Mailbox Statistics"="fListMailboxStats"
								"List Mailbox Statistics With Shared"="fListMailboxStatsWithShared"
								"Set 30day Deleted Items Recovery (All)"="fRecoverDelItem30days"
								"Toggle Shared Mailbox"="fConvertSharedMailbox"
								"List Email Forwarding Status"="fCheckForwarding"
								"Toggle Access to Services"=@{	"Display Mailbox Access Status"="fDisplayCASMailboxStatus"
																"Toggle MAPI Access"="fToggleMAPI"
																"Toggle OWA Access"="fToggleOWA"
																"Toggle IMAP Access"="fToggleImap"
																"Toggle POP Access"="fTogglePop"
																"Toggle ActiveSync"="fToggleActiveSync"
																}
								"Hide/Unhide from GAL"="fToggleMailboxHideFromGAL"
								"Email Alias for Mailboxes"=@{
															"Remove Mailbox Email Alias"="fRemoveMailboxEmailAlias"
															"Add Mailbox Email Alias"="fAddMailboxEmailAlias"
															"Set Default Mailbox Alias"="fSetDefaultEmailAlias"
															}
								"Change Forwarding Status"="fSetMailboxForwarding"
								"Add Shared Mailbox"="fAddSharedMailbox"
								"Shared Mailbox Copy to Sent (Single Mailbox)"="fSingleSharedSentItemsCopy"
								"Shared Mailbox Copy to Sent (All Mailboxes)"="fAllSharedSentItemsCopy"
								"Manage Clutter"=@{
													"List Clutter Status"="fshowclutterall"
													"Disable Clutter for ALL mailboxes"="fdisableclutterall"
													"Enable Clutter for ALL mailboxes"="fenableclutterall"
													}
								"Manage Focused Inbox"=@{
													"Disable Focused Inbox for ALL mailboxes"="fdisablefocusedall"
													"Enable Focused Inbox for ALL mailboxes"="fenablefocusedall"
													}													
								}
	"Dist Groups"=@{			"List Dist Groups and Members"="fListDistMembers"
								"Edit Group Members"=@{
														"Add User to Dist Group"="fadduserdistgroup"
														"Remove User from Dist Group" = "fremoveuserdistgroup"
														}
								"Add New Dist Group"="fAddNewDistGroup"
								"Email Alias for Dist Groups"=@{
																"Add Group Email Alias"="fAddDistGroupEmailAlias"
																"Remove Group Email Alias"="fRemoveDistGroupEmailAlias"
																}	
								"Hide/Unhide from GAL"="fToggleDistHideFromGAL"
								"View External Auth Settings"="fViewDistExtAuth"
								"Toggle External Auth Settings"="fToggleDistExtAuth"
								}
	"MSOnline Org"=@{			"List Partner Information"="fViewPartnerInfo"
								"List Domain Info"="fVeiwDomain"
								"List Licencing Status"="fGetMsolAccountSku"
								}
	"Mail Flow and Spam"=@{		"List Transport Rule Status"="fGetTranStatus"
								"Toggle Rule Status"="fToggleTransportRule"
								"Enable Connection Filter Safe List"="fSetSafelistEnable"
								"Create Warning Rules"="fcreatewarningrules"
								"Message Trace (Sender)"="fSenderMessageTrace"								
								}
	"Partner"=@{				"List All Active Accounts"="fListAllActive"
								}
		
								
	}
	
		
# Control the login process ================================================================
$global:xLocalUserPath = $env:UserProfile+"\Office365Data" #Define the local path to store user data in - Should NOT end with a '\'
$global:xCompanyFilePath = "Z:\~Tools\Powershell\companys.csv" #Allow central company.csv file for multi users
$ErrorActionPreference = 'Stop'




function global:start-login{
	$global:ForceLoginFirst = $true
	
	#If required create the user local path directory
	if (test-path $global:xLocalUserPath) {} Else {
		new-item $global:xLocalUserPath -type Directory
		cls
	}
	
	#This script requires the Multi Layered Dynamic Menu System Module from www.AshleyUnwin.com/Powershell_Multi_Layered_Dynamic_Menu_System
	Import-Module $global:xLocalUserPath"\MenuSystem.psm1" -ErrorAction silentlycontinue
	$i = 0
	while ((get-module -Name MenuSystem) -eq $null) {
		$source = "https://raw.githubusercontent.com/manicd/Powershell-Multi-Layered-Dynamic-Menu-System/master/MenuSystem.psm1"
		$destination = $global:xLocalUserPath+"\MenuSystem.psm1"
		if ($PSversiontable.PSVersion.Major -lt 3) {
			$web=New-Object System.Net.WebClient
			$web.DownloadFile($source,$destination)
		} else {
			Invoke-WebRequest $source -OutFile $destination
		}
		if (test-path $destination) {Import-Module $destination} else {Return "Error Menu System Download Failed"}
		if ((get-module -name MenuSystem) -ne $null) {fDisplayInfo -xText "Menu System Installed"} else { $i++; write-host "Loop "$i}
		if ($i -eq 3) {Return "FATAL ERROR: Failed to Install Menu System"}
	}
	
		
	fClear-Login	
	
	cls
	fLoginMenu	
}

function global:fLoginMenu{
	# Requires a CSV File in with the columns company,adminuser
	if (test-path $global:xLocalUserPath) {} Else {
		new-item $global:xLocalUserPath -type Directory
	}	
	
	if ((test-path $global:xCompanyFilePath ) -ne $true)  {
		#If the file is missing locally and on the share download the demo data from github
		$source = "https://raw.githubusercontent.com/manicd/PowerShell-Office-365-Administration-Script/master/companys.csv"
		$destination = $global:xCompanyFilePath
		Invoke-WebRequest $source -OutFile $destination
		fDisplayInfo -xText "Please edit the companys.csv file for your own use." -xText2 "At this time please ensure you do NOT include spaces in company names" -xText3 "Tip: Try using hypenated names instead e.g. Johns-Cleaning-Company " -xTime 5
		Invoke-Item $destination
		pause
	} 
	
	
	#Die if company file still not there
	if ((test-path $global:xCompanyFilePath ) -ne $true ) {
		Throw "Cannot Locate Company File at $global:xCompanyFilePath. Check file location and permissions and try again."
	}
	
	$global:csv = import-csv $global:xCompanyFilePath
	
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
	$global:xService = $global:csv | where-object {$_.company -eq $xCompany} | select Service
	$global:xService = $global:xService.Service
	
	$passfile = $global:xLocalUserPath+"\"+$global:xCompany+"365pass.txt"
	if (test-path $passfile) {
		} else {
		$string = fUserPrompt -xQuestion "Enter the Password"
		cls 
		if (test-path $global:xLocalUserPath) {} Else {
			new-item $global:xLocalUserPath -type Directory
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
	$xReturn.add("xService",$xService)
	return $xReturn
}

function global:fLoginTo365{

PARAM(
[string]$xAdminUser,
[string]$xPass,
[string]$xCompany
)
	
	#This script requires Microsoft Online Services Sign-In Assistant for IT Professionals installed
	$i = 0
	while ((test-path $env:programfiles'\Common Files\microsoft shared\Microsoft Online Services') -ne $true) {
		fDisplayInfo -xText "It appears you don't have Microsoft Online Services Sign-In Assistant for IT Professionals installed" -xText2 "Let's Install that now"
		$source = "http://download.microsoft.com/download/5/0/1/5017D39B-8E29-48C8-91A8-8D0E4968E6D4/en/msoidcli_64.msi" 
		$destination = $global:xLocalUserPath+"\Microsoft Online Services Sign-In Assistant for IT Professionals RTW.msi"
		if ($PSversiontable.PSVersion.Major -lt 3) {
			$web=New-Object System.Net.WebClient
			$web.DownloadFile($source,$destination)
		} else {
			Invoke-WebRequest $source -OutFile $destination
		}
		if (test-path $destination) {Invoke-Item $destination} else {Return "Error Online Services Sign-In Assistant Download Failed"}
		if ((test-path $env:programfiles+'\Common Files\microsoft shared\Microsoft Online Services') -ne $true) {
			fDisplayInfo -xText "You are required to install Microsoft Online Services Sign-In Assistant to run this script" -xtext2 "Please Complete the Installer before continuing"	
			$i++
			if ($i -eq 3) {Return "FATAL ERROR: Failed to Install Microsoft Online Services Sign-In Assistant"}
			pause
		}
	}
	
	#This script requires Azure Active Directory Module for Windows PowerShell installed
	Import-Module MSOnline -ErrorAction SilentlyContinue
	$i = 0
	while ((Get-Module -Name MSOnline) -eq $null) {
		fDisplayInfo -xText "It appears you don't have Azure Active Directory Module for Windows PowerShell installed" -xText2 "Let's Install that now" -xTime 3
		$source = "https://bposast.vo.msecnd.net/MSOPMW/Current/amd64/AdministrationConfig-en.msi" 
		$destination = $global:xLocalUserPath+"\Azure Active Directory Module for Windows PowerShell.msi"
		if ($PSversiontable.PSVersion.Major -lt 3) {
			$web=New-Object System.Net.WebClient
			$web.DownloadFile($source,$destination)
		} else {
			Invoke-WebRequest $source -OutFile $destination
		}
		if (test-path $destination) {Invoke-Item $destination} else {Return "Error Azure Active Directory Module Download Failed"}
		fDisplayInfo -xText "Please complete the setup before continuing" -xtime 5
		pause
		Import-Module MSOnline  -ErrorAction SilentlyContinue
		if ((Get-Module -Name MSOnline) -eq $null) {
			fDisplayInfo -xText "You are required to install Microsoft Online Services Sign-In Assistant to run this script" -xtext2 "Please Complete the Installer before continuing"	-xTime 3
			$i++
			if ($i -eq 3) {Return "FATAL ERROR: Failed to Install Azure Active Directory Module"}
			pause
		}
		
	}
	
 	# If username has been set, login
    if ($xPass)	{
		Write-host "Connecting to"$xCompany -Fore Green
		Write-host "Creating Credential Object" -Fore Green
		$global:O365Cred=New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $xAdminUser, ($xPass | ConvertTo-SecureString) -ErrorAction stop
		Write-host "Creating Session Object" -Fore Green
		$O365Session = New-PSSession –ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $O365Cred -Authentication Basic -AllowRedirection -ErrorAction stop
		write-host "Importing Session" -Fore Green
		Import-PSSession $O365Session -ErrorAction stop
		write-host "Connecting to MSOL Service" -Fore Green
		Connect-MsolService –Credential $O365Cred -ErrorAction stop
		cls
		write-host "`nYou are now logged in to"$xCompany". Type 'Use-Admin' to access the menu." -Fore Green
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
	Remove-item $global:xLocalUserPath+"\*" -confirm
}	

# Global Functions ============================================================

function global:qq{start-login; use-admin}
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
$xCount = $xText | Measure-Object -Character | select Characters
$xCount = $xCount.Characters

$xCount2 = $xText2 | Measure-Object -Character | select Characters
$xCount2 = $xCount3.Characters

$xCount3 = $xText3 | Measure-Object -Character | select Characters
$xCount3 = $xCount3.Characters

$xCountArray=@($xCount,$xCount2,$xCount3)
$xCountFinal = $xCountArray | Measure-Object -Maximum


$i=0
while ($i -lt ($xCountFinal.Maximum + 18)) {[string]$xStars = $xStars+"*"; $i++}
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
	if ($global:ForceLoginFirst -eq $false) {
		Return "`n`n`tPlease run 'Start-Login' to Login to Office 365 First`n`n`tSome Shortcuts for you: `n`t'qqq' will quickly run start-login, `n`t'www' will quickly run Use-Admin, `n`t'qq' to run Login and Admin together!`n`n "
	}

	if (get-module -name MenuSystem){
	}elseif (Test-Path c:\powershell\MenuSystem.psm1) {
		Import-Module c:\powershell\MenuSystem.psm1
	}else{
		Import-Module Z:\~Tools\Powershell\MenuSystem.psm1
	}

	[bool]$global:UseAdminLoaded=$true


		#OLD MENU LOCATION
	
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
		#Restores Settings
		$global:SelectionHist = $global:xTempSelectionHist
		$global:QuitMenu = $false
	} else {
		#Saves Settings
		$global:xTempSelectionHist = $global:SelectionHist
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
#This function collects a valid UPN and will ask again if the entered information does not confirm to the correct layout and valid domains. 
PARAM(
[string]$xText,
[bool]$xCurrent = $false
)
	while (!$xfCollectUPNvUPN) {
		$xInput = fUserPrompt -xQuestion $xText" (Type 'QUIT' to exit)"
		if (fCheckUPN -xUPN $xInput -xCurrent $xCurrent) {
			$xfCollectUPNvUPN = $xInput
			remove-variable -name xInput
		} elseif ($xInput -eq "quit") {
			Return $false
			# You must include if ($Var -eq $false) {Return $false} after calling the function to fully quit the function
		} else 	{
			fDisplayInfo -xText "Invalid Selection" -xColor "red"
		}
	}
	Return $xfCollectUPNvUPN
}

function global:fCheckIdentity {
# This function checks if a given Alias the correct alias assigned to an object
PARAM(
[string]$id,
[string]$xDefine = "any"
)
#Function to check if an identity specified exists
	if ($xDefine -eq "any") {
		if (Get-Mailbox -identity $id -ErrorAction 'silentlycontinue') {
			Return $true
		} elseif (Get-DistributionGroup -identity $id -ErrorAction 'silentlycontinue') {
			Return $true
		} elseif (Get-Contact -identity $id -ErrorAction 'silentlycontinue') {
			Return $true
		} 
		Return $false
	}elseif ($xDefine -eq "mailbox") {
		if (Get-Mailbox -identity $id -ErrorAction 'silentlycontinue') {
			Return $true
		}
	}elseif ($xDefine -eq "group") {
		if (Get-DistributionGroup -identity $id -ErrorAction 'silentlycontinue') {
			Return $true
		}
	}elseif ($xDefine -eq "contact") {
		if (Get-Contact -identity $id -ErrorAction 'silentlycontinue') {
			Return $true
		}
	}
	write-error "Unable to Determine the status in function global:fCheckIdentity"
	Return $false
}

function global:fCollectIdentity {
# This function Collects an Alias and checks that it is valid and currently in use in order to be used in modifying existing objects. If the Alias is invalid, it will ask again.
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
			fDisplayInfo -xText "Invalid Selection" -xColor "red" -xTime 1
			fDidYouMean	-xInput $xInput
			
		}
	}
	Return $xVar
}

function global:fDidYouMean {
PARAM(
[string]$xInput
)
	fDisplayInfo -xText "Lets check if i can find that for you..."
	$xGetRecipient = get-recipient *
	$xPossible = $xGetRecipient | ?{ ($_.Alias -match $xInput) -OR ($_.DisplayName -match $xInput) } -ErrorAction silentlycontinue
	
	if ($xPossible -eq $null) {
		#if the first option fails, try searching just the first few char's of the string
		$xPossible2 = $xGetRecipient | ?{ ($_.Alias -match $xInput.substring(0,5)) -OR ($_.DisplayName -match $xInput.substring(0,5)) }  -ErrorAction silentlycontinue
	} 
	if (($xPossible -eq $null) -and ($xPossible2 -eq $null)) {
		#if that fails, try searching just the last few char's of the string
		$xPossible3 = $xGetRecipient | ?{ ($_.Alias -match $xInput.substring($xInput.length - 5,5)) -OR ($_.DisplayName -match $xInput.substring($xInput.length - 5,5)) } -ErrorAction silentlycontinue
	} 
	
	
	if ($xPossible -ne $null) {
	fDisplayInfo -xText "Did you mean one of these?"
	write-host ($xPossible | ft DisplayName, Alias | out-string)
	} elseif ($xPossible2 -ne $null) {
	fDisplayInfo -xText "You might have meant one of these?"
	write-host ($xPossible2 | ft DisplayName, Alias | out-string)
	} elseif ($xPossible3 -ne $null) {
	fDisplayInfo -xText "Maybe one of these is what your looking for?"
	write-host ($xPossible3 | ft DisplayName, Alias | out-string)
	} else {
	fDisplayInfo -xText "Sorry, I'm not sure who you might mean, Try entering the alias again."
	}
}

function global:fCollectAlias {		
		# This Function checks if the user is entering an AVLIABLE alias for a new object and if not requests the user re-enters the alias
		$xAlias = fUserPrompt -xQuestion "Alias: (Type QUIT to exit)"
		$xAliasValid = $false
		while ($xAliasValid -eq $false) {
			if ( $xAlias -eq "QUIT") {
				Return $false
			}elseif ((fCheckIdentity -id $xAlias) -eq $true) {
				fDisplayInfo -xText "Alias already in use"
				$xAlias = fUserPrompt -xQuestion "Alias"
			} elseif ((fCheckIdentity -id $xAlias) -eq $false){
				$xAliasValid = $true
			}else {
				Write-host "The fCollectAlias function has errored....QUITTING"
				Pause
				Return $false
			}
		}
		Return $xAlias
	}

function global:fExportCSV {
PARAM(
$xInput,
[string]$xFilename
)

	$xExpCSV = fUserPrompt -xQuestion "Would you like to export this as CSV? (y/n)"
	if ($xExpCSV -eq "y") {
		$xInput | export-csv $xLocalUserPath"\"$xFilename".csv"
		write-host "Generating CSV, Please Wait..."
		timeout 3
		Invoke-item $xLocalUserPath"\"$xFilename".csv"
	}
}

function global:fExportTXT {
PARAM(
$xInput,
[string]$xFilename
)

	$xExpCSV = fUserPrompt -xQuestion "Would you like to export this as TXT? (y/n)"
	if ($xExpCSV -eq "y") {
		$xInput | out-file $xLocalUserPath"\"$xFilename".TXT"
		write-host "Generating TXT, Please Wait..."
		timeout 3
		Invoke-item $xLocalUserPath"\"$xFilename".TXT"
	}
}

# Below this line are the functions called by the menu values=================================

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
				fDisplayInfo -xText "Setting New Password"
				Write-host
				$xPass = Set-MsolUserPassword -UserPrincipalName $xUser -ForceChangePassword $false -NewPassword $xPass
				fDisplayInfo -xText "Password now set" -xColor "green" -xTime 3
				$xPass = $null	
				Cls
				} else {
				fDisplayInfo -xText "Password not entered, Nothing has been changed." -xColor "Red"
				fResetUserPasswordsCollectPass -xUser $xUser				
			}
		}else{
		cls
		fDisplayInfo -xText "Quitting....Nothing has been changed." -xColor "Red" -xTime 3
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
	
		
	$xAlias = fCollectAlias
	#Create the Menu Hash Table Object	
	$xDomainMenuHash = New-Object System.Collections.HashTable
	#Create Menu Structure Hash Table and set values to be function with UPN as input
	get-msoldomain | select name | sort-object Name | foreach-object {
			$xDomainMenuHash.add($_.Name,'$Global:xUPN = "'+$xAlias+'@'+$_.Name+'"')
		}
	#Call the Menu	
	use-menu -MenuHash $xDomainMenuHash -Title "Select Domain" -NoSplash $True
	
	
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
			fDisplayInfo -xText "You Do Not currently have enough Licenses to proceed." -xText2 "Please Login to Office 365 Website and purchase more" -xText3 "licences before proceeding." -xTime 5
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

	$xUserList = (get-msoluser | select DisplayName, UserPrincipalName, Licenses)
	write-host ($xUserList | sort DisplayName | format-table | out-string)
	
	$xExpCSV = fUserPrompt -xQuestion "Would you like to export this as CSV? (y/n)"
	if ($xExpCSV -eq "y") {
		$xUserList | export-csv $xLocalUserPath"\UserList.csv"
		write-host "Generating CSV, Please Wait..."
		timeout 5
		Invoke-item $xLocalUserPath"\UserList.csv"
	}
}

function global:fListLicensedUsers {

	$xUserList = (get-msoluser | ? { $_.islicensed -eq $true } | select DisplayName, UserPrincipalName, Licenses)
	write-host ($xUserList | sort DisplayName | format-table | out-string)
	
	$xExpCSV = fUserPrompt -xQuestion "Would you like to export this as CSV? (y/n)"
	if ($xExpCSV -eq "y") {
		$xUserList | export-csv $xLocalUserPath"\UserList.csv"
		write-host "Generating CSV, Please Wait..."
		timeout 5
		Invoke-item $xLocalUserPath"\UserList.csv"
	}
}

function global:fEditUserAccountName {
		
	$xOldUPN = fCollectUPN -xText "Enter Old User UPN:"
	if ($xOldUPN -eq $false) {Return $false}	
	
	$xNewUPN = fCollectUPN -xText "Enter New User UPN:"
	if ($xNewUPN -eq $false) {Return $false}	

	$xNewFirstName = fUserPrompt -xQuestion "What is the New Users First Name"
	$xNewLastName = fUserPrompt -xQuestion "What is the New Users Last Name"
	$xNewDisplayName = $xNewFirstName+" "+$xNewLastName

	set-msoluserprincipalname -UserPrincipalName $xOldUPN -NewUserPrincipalName $xNewUPN
	set-msoluser -UserPrincipalName $xNewUPN -Firstname $xNewUserName -LastName $xNewLastName -DisplayName $xNewDisplayName
	write-host (get-msoluser -UserPrincipalName $xNewUPN | fl UserPrincipalName, FirstName, LastName, ProxyAddresses | out-string)
	#This might not rename mailbox - investigate
}

function global:fSetPasswordNeverExpireAll {
			
			fDisplayInfo -xText "Collecting User List"
			$xUsers = Get-msoluser
			
			fDisplayInfo -xText "Adjusting Passwords to Never Expire"
			$xNull = $xUsers | Set-MsolUser -PasswordNeverExpires $true
			
			fDisplayInfo -xText "Gathering Results"
			$xResults = Get-msoluser | Select UserPrincipalName, PasswordNeverExpires | sort UserPrincipalName
			
			write-host ( $xResults | format-table | out-string)
			fExportCSV -xInput $xResults -xFilename "PassNeverExpire"
}

#Mailboxes =======================================================================================

function global:fListMailboxes {
	
	$xMList = get-mailbox | select DisplayName, Alias, UserPrincipalName, PrimarySmtpAddress, RecipientTypeDetails | where-object {$_.Alias -NOTMATCH 'DiscoverySearch'} | sort DisplayName
	
	write-host ( $xMList | format-table | out-string)
	
	fExportCSV -xInput $xMList -xFilename "MailboxList"
}

function global:fListUserMailboxes {
	
	$xMList = get-mailbox | ? { $_.IsShared -eq $false } | ? { $_.IsResource -eq $false } | ? { $_.IsLinked -eq $false } | ? { $_.IsRootPublicFolderMailbox -eq $false } | select DisplayName, Alias, UserPrincipalName, PrimarySmtpAddress | where-object {$_.Alias -NOTMATCH 'DiscoverySearch'} | sort DisplayName
	
	write-host ( $xMList | format-table | out-string)
	
	fExportCSV -xInput $xMList -xFilename "MailboxList"
}

function global:fListMailboxStatsWithShared {
	
	Write-host "Generating Statistics, Please Wait..."
	$xMStats = get-mailbox | where-object {$_.Alias -NOTMATCH 'DiscoverySearch'} | foreach-object { get-mailboxstatistics -identity $_.UserPrincipalName | select DisplayName, TotalItemSize, LastLogonTime, RecipientTypeDetails | sort TotalItemSize} 
	write-host ( $xMStats | sort TotalItemSize | format-table | out-string)
		
	fExportCSV -xInput $xMStats -xFilename "MailboxStats"
}

function global:fListMailboxStats {
	
	Write-host "Generating Statistics, Please Wait..."
	$xMStats = get-mailbox | where-object {$_.Alias -NOTMATCH 'DiscoverySearch' -and $_.RecipientTypeDetails -NOTMATCH 'SharedMailbox'} | foreach-object { get-mailboxstatistics -identity $_.UserPrincipalName | select DisplayName, TotalItemSize, LastLogonTime | sort TotalItemSize} 
	write-host ( $xMStats | sort TotalItemSize | format-table | out-string)
		
	fExportCSV -xInput $xMStats -xFilename "MailboxStats"
}

function global:fCheckForwarding {
	
	$xCheckFwd = get-mailbox | where-object {$_.Alias -NOTMATCH 'DiscoverySearch'} | select DisplayName, PrimarySMTPAddress, ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward | sort PrimarySMTPAddress
	
	write-host ($xCheckFwd | format-table | out-string)
	
	fExportCSV -xInput $xCheckFwd -xFilename "ForwardList"
	
}

function global:fToggleMailboxHideFromGAL {
	$xUser = fCollectIdentity -xText "Enter the User who you would like to hide"
	if ($xUser -eq $false) {Return $false}
	$xStatus = Get-mailbox -identity $xUser | select HiddenFromAddressListsEnabled
	if ($xStatus.HiddenFromAddressListsEnabled) {
		$xUnhide = "x"
		$xUnhide = fUserPrompt -xQuestion "Would you like to unhide the mailbox? (y/n)"
		if ($xUnhide -eq "n") {Return $false}
		if ($xUnHide -eq "y") {$xHidden = $false}
	} else {
		$xhide = "x"
		$xhide = fUserPrompt -xQuestion "Would you like to hide the mailbox? (y/n)"
		if ($xHide -eq "n") {Return $false}
		if ($xHide -eq "y") {$xHidden = $true}
	}
	Set-mailbox -identity $xUser -HiddenFromAddressListsEnabled $xHidden
	write-host ( get-mailbox -identity $xUser | select Displayname, HiddenFromAddressListsEnabled | format-list | out-string )
	pause
}

function global:fSetMailboxForwarding {
	$xIdentity = fCollectIdentity -xText "Enter Alias of User to Forward:"
	if ($xIdentity -eq $false) {Return $false}
	
	$xMailbox = get-mailbox -identity $xIdentity 
	
	fDisplayInfo -xText "The current setup is:"
	write-host ($xMailbox | select DisplayName, PrimarySMTPAddress, ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward | format-table | out-string)
	
	if (($xMailbox.ForwardingAddress -ne $null) -OR ($xMailbox.ForwardingSMTPAddress -ne $null)    ) {
		
		$xFwdCancel = fUserPrompt -xQuestion "Would you like to cancel the Current Forwarding? (y/n)"
			
		if ($xFwdCancel -eq "y") {
			set-mailbox -identity $xIdentity -ForwardingAddress $null -ForwardingSMTPAddress $null -DeliverToMailboxAndForward $false
		}elseif ($xFwdCancel -eq "n") {
			$xMailbox = get-mailbox -identity $xIdentity 
			fDisplayInfo -xText "Your configuration has not changed" -xColor Yellow
			pause
			Return $false
		}else{
			Return $false
		}
			
	} else {
	
		$xFwdAddress = fCollectIdentity -xText "Who would you like to forward email to?"
		if ($xFwdAddress -eq $false) {Return $false}
		
		$xDelToMailbox = fUserPrompt -xQuestion "Would you like to continue delivering to the Mailbox? (y/n)"
		
		if ($xDelToMailbox -eq "y") {
			set-mailbox -identity $xIdentity -ForwardingAddress $xFwdAddress -DeliverToMailboxAndForward $true
		}elseif ($xDelToMailbox -eq "n") {
			set-mailbox -identity $xIdentity -ForwardingAddress $xFwdAddress -DeliverToMailboxAndForward $false 
		}else{
			Return $false
		}
	
	}
	
	$xMailbox = get-mailbox -identity $xIdentity 
	fDisplayInfo -xText "The New setup is:"
	write-host ($xMailbox | select DisplayName, PrimarySMTPAddress, ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward | format-table | out-string)
	pause
}

function global:fAddSharedMailbox {
	fStoreMainMenu -xRestore 0
	
	$xDisplayName = fUserPrompt -xQuestion "Display Name"
	
	$xAlias = fCollectAlias	
	if ($xAlias -eq $false) {Return $false}
	
	#Create the Menu Hash Table Object	
	$xDomainMenuHash = New-Object System.Collections.HashTable
	#Create Menu Structure Hash Table and set values to be function with UPN as input
	get-msoldomain | select name | sort-object Name | foreach-object {
			$xDomainMenuHash.add($_.Name,'$global:xPrimarySMTPAddress = "'+$xAlias+'@'+$_.Name+'"')
		}
	#Call the Menu	
	use-menu -MenuHash $xDomainMenuHash -Title "Select Domain" -NoSplash $True
	
	New-Mailbox -Name $xAlias –Shared -PrimarySmtpAddress $xPrimarySMTPAddress -DisplayName $xDisplayName
	
	write-host (Get-Mailbox $xAlias | Select Name, Alias, IsShared, PrimarySMTPAddressSelect | format-table | out-string )

	pause

	fStoreMainMenu -xRestore 1
}

function global:fConvertSharedMailbox {

	$xUser = fCollectIdentity -xText "Enter the mailbox you would like to convert:"
	if ($xUser -eq $false) {Return $false}

	$xStatus = Get-mailbox -identity $xUser | select RecipientTypeDetails
	
	if (($xStatus.RecipientTypeDetails) -eq "UserMailbox") {
		
		fDisplayInfo -xText "Converting to Shared Mailbox"
		Set-mailbox -identity $xUser -Type Shared	
			
	} else { 
	
		fDisplayInfo -xText "Converting to User Mailbox"
		Set-mailbox -identity $xUser -Type Regular
		
	}
	
	start-sleep 3
	write-host ( get-mailbox -identity $xUser | select Displayname, RecipientTypeDetails | format-list | out-string )
	pause

}

function global:fShowMailboxPerms {
	
	$xAlias = fCollectIdentity -xText "Who's permissions would you like to check?"
	
	write-host (Get-MailboxPermission -identity $xAlias | where {$_.user -NOTMATCH "NT Authority"} | where {$_.User -NOTMATCH "Domain Admins"}| where {$_.user -NOTMATCH "Enterprise Admins"} | where {$_.user -NOTMATCH "Organization Management"} | where {$_.user -NOTMATCH "Public Folder Management"} | where {$_.user -NOTMATCH "Exchange Servers"} | where {$_.user -NOTMATCH "Exchange trusted Subsystem"} | where {$_.user -NOTMATCH "Managed Availability Servers"} | where {$_.user -NOTMATCH "Administrator"} | select User, AccessRights | Format-Table | Out-String)
	
	pause
	
}

function global:fRecoverDelItem30days {
			
			fDisplayInfo -xText "Collecting Mailbox List"
			$xMailboxes = Get-Mailbox -ResultSize unlimited -Filter {(RecipientTypeDetails -eq 'UserMailbox')} 
			fDisplayInfo -xText "Adjusting Deleted Items Recovery to 30 days"
			$xNull = $xMailboxes.UserPrincipalName | Set-Mailbox -RetainDeletedItemsFor 30
			fDisplayInfo -xText "Gathering Results"
			$xResults = Get-Mailbox -ResultSize unlimited -Filter {(RecipientTypeDetails -eq 'UserMailbox')} | Select Displayname, RetainDeletedItemsFor | sort displayname
			write-host ( $xResults | format-table | out-string)
			fExportCSV -xInput $xResults -xFilename "RecoverDelItemsTime"
}

function global:fSingleSharedSentItemsCopy {
	
	$xSMbox = fCollectIdentity -xText "Enter the Shared Mailbox name"
	if ($xSMBox -eq $false) {Return $false}
	
	if ( (get-mailbox $xSMbox | select RecipientTypeDetails).recipienttypedetails -eq "SharedMailbox" ) {
	
		Set-Mailbox -identity $xSMbox -MessageCopyForSentAsEnabled $true -MessageCopyForSendOnBehalfEnabled $true
		write-host (Get-Mailbox -identity $xSMbox | select identity,MessageCopyForSentAsEnabled,MessageCopyForSendOnBehalfEnabled | format-table | out-string)
		pause
		
	}
	else{
	
		fDisplayInfo -xText "Mailbox specified is NOT a Shared Mailbox"
		pause
		
	}
	
}

function global:fAllSharedSentItemsCopy {
	

	$SMBoxList = Get-Mailbox -ResultSize unlimited -Filter {(RecipientTypeDetails -eq 'SharedMailbox')}
	$SMBoxList	| Set-Mailbox -MessageCopyForSentAsEnabled $true -MessageCopyForSendOnBehalfEnabled $true
	$SMBoxchart = Get-Mailbox -ResultSize unlimited -Filter {(RecipientTypeDetails -eq 'SharedMailbox')} | select identity,MessageCopyForSentAsEnabled,MessageCopyForSendOnBehalfEnabled
	write-host ( $SMBoxchart | format-table | out-string)
	
	fExportCSV -xInput $SMBoxchart -xFilename "SharedMailboxCopySentItemsList"
		
}



	#Alias
function global:fAddMailboxEmailAlias {

	$xIdentity = fCollectIdentity -xText "Enter identity:"
	if ($xIdentity -eq $false) {Return $false}
		
	$xMailbox = get-mailbox -identity $xIdentity 
	$xEmails = $xMailbox.EmailAddresses
	$i = 1
	fDisplayInfo -xText "The current emails attached to this mailbox are"
	foreach ($email in $xEmails) {
		Write-Host "`t`t"$i" - "$email
	}
	write-host 
	
	$xNewEmailAddress = fCollectUPN -xText "Enter the additional email address:" -xCurrent $false
	if ($xNewEmailAddress -eq $false) {Return}

	Set-Mailbox -identity $xIdentity -emailaddresses @{Add=$xNewEmailAddress}

	$xMailbox = Get-Mailbox -identity $xIdentity 
	$xEmails = $xMailbox.EmailAddresses
	$i = 1
	fDisplayInfo -xText "The current emails attached to this mailbox are now"
	foreach ($email in $xEmails) {
		Write-Host "`t`t"$i" - "$email
	}
	pause
}

function global:fRemoveMailboxEmailAlias {

	$xIdentity = fCollectIdentity -xText "Enter identity:"
	if ($xIdentity -eq $false) {Return $false}
		
	$xMailbox = get-mailbox -identity $xIdentity 
	$xEmails = $xMailbox.EmailAddresses
	$i = 1
	fDisplayInfo -xText "The current emails attached to this mailbox are"
	foreach ($email in $xEmails) {
		Write-Host "`t`t"$i" - "$email
	}
	write-host 
	
	$xNewEmailAddress = fCollectUPN -xText "Enter the email address to remove:" -xCurrent $false
	if ($xNewEmailAddress -eq $false) {Return}

	Set-Mailbox -identity $xIdentity -emailaddresses @{Remove=$xNewEmailAddress}

	$xIdentity = Get-Mailbox -identity $xIdentity 
	$xEmails = $xMailbox.EmailAddresses
	$i = 1
	fDisplayInfo -xText "The current emails attached to this mailbox are now"
	foreach ($email in $xEmails) {
		Write-Host "`t`t"$i" - "$email
	}
	pause
}

function global:fSetDefaultEmailAlias {


	$xIdentity = fCollectIdentity -xText "Enter Identity:"
	if ($xIdentity -eq $false) {Return $false}
	
	$xMailbox = get-mailbox $xIdentity
	$i = 0; 
	
	foreach ($email in $xMailbox.EmailAddresses) {
		if ($email -cmatch "SMTP:") {
			$xMailBox.EmailAddresses[$i] = "smtp:ash@sussexcommunity.org.uk"
		} 
		$i++ 
	}

	set-mailbox -identity $xIdentity -EmailAddresses $xMailBox.EmailAddresses
	
	Pause
}

	#MailboxPermissions
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
		fDisplayInfo -xText "Defaulting Auto Mapping to True"
		$xAutoMap = $true
	}
	$xSendAsYN = fUserPrompt -xQuestion "Would you like to include Send As Permissions? (y/n)"
	$xSendAs = $false
	if ( ($xSendAsYN -eq "y") -OR ($xSendAsYN -eq "yes")) {
		$xSendAs = $true
	}elseif ( ($xSendAsYN -eq "n") -OR ($xSendAsYN -eq "no") ) {
				$xSendAs = $false
	}else{
		#Default
		fDisplayInfo -xText "Defaulting Send As to True"
		$xSendAs = $true
	}
	
	fDisplayInfo -xText "Adding Mailbox Permission"
	Add-MailboxPermission -identity $xMailbox -User $xUser -AccessRight fullaccess -InheritanceType all -Automapping $xAutoMap
	
	if ($xSendAs -eq $true) {
		fDisplayInfo -xText "Setting Send As Permission"
		Add-RecipientPermission $xMailbox -AccessRights SendAs -Trustee $xUser
	}
	
	fDisplayInfo -xText "Mailbox Permissions:"
	$xNewMBPerms = Get-MailboxPermission -identity $xMailbox
	write-host ( $xNewMBPerms | format-table | out-string)
	
	fDisplayInfo -xText "Send As Permissions:"
	$xNewSendAsPerms = Get-RecipientPermission $xMailbox
	write-host ($xNewSendAsPerms | format-table | out-string)
	
	fDisplayInfo -xText "Would you like to export Mailbox Permissions List"
	fExportCSV -xInput $xNewMBPerms -xFilename "MailboxPerms"
	fDisplayInfo -xText "Would you like to export Send As List"
	fExportCSV -xInput $xNewSendAsPerms -xFilename "SendAsPerms"
}

function global:fGrantSendAsPerms {
	
	$xUser = fCollectIdentity -xText "Enter the User who would like the access"
	if ($xUser -eq $false) {Return $false}
	$xMailbox = fCollectIdentity -xText "Enter the Mailbox they would like access for"
	if ($xMailbox -eq $false) {Return $false}
	
	fDisplayInfo -xText "Setting Send As Permission"
	Add-RecipientPermission $xMailbox -AccessRights SendAs -Trustee $xUser
	
	fDisplayInfo -xText "Send As Permissions:"
	$xNewSendAsPerms = Get-RecipientPermission $xMailbox
	write-host ($xNewSendAsPerms | format-table | out-string)
	
	fExportCSV -xInput $xNewSendAsPerms -xFilename "SendAsPerms"
}

function global:fGrantFullAccessMailboxAllMailboxes {
	
	fDisplayInfo -xText "WARNING: You are about to grant access to ALL mailboxes" -xColor "red"
	
	$xUser = fCollectIdentity -xText "Enter the User who would like the access"
	if ($xUser -eq $false) {Return $false}
	$xAutoMapYN = fUserPrompt -xQuestion "Would you like to enable AutoMapping? (y/n)"
	if ($xAutoMapYN -eq "y") {
		$xAutoMap = $true
	}elseif ($xAutoMapYN -eq "n") {
		$xAutoMap = $false
	}else{
		#Default
		$xAutoMap = $true
	}
		
	$xMailboxes = get-mailbox 
	$xUserPermList = @()
	
	foreach ( $xMailbox in $xMailboxes) { 
		if ($xMailbox.name -NOTMATCH "Discovery") { 
			add-mailboxpermission -identity $xMailbox.PrimarySMTPAddress -User $xUser -accessright fullaccess -inheritancetype all -automapping $false 
			$xUserPermList += get-mailboxpermission -identity $xMailbox.PrimarySMTPAddress -User $xUser
			Add-RecipientPermission -identity $xMailbox.UserPrincipalName -AccessRights SendAs -Trustee $xUser -Confirm:$false
		} 
	}
	
	clear
	
	write-host ( $xUserPermList | sort Identity | format-table | out-string)
	
	$xExpCSV = fUserPrompt -xQuestion "Would you like to export this as CSV? (y/n)"
	if ($xExpCSV -eq "y") {
		$xUserPermList | export-csv $xLocalUserPath"\UserPermList.csv"
		write-host "Generating CSV, Please Wait..."
		timeout 5
		Invoke-item $xLocalUserPath"\UserPermList.csv"
	}
	
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

function global:fFindUserPermissions {
write-host (get-msoluser | select UserPrincipalName | sort UserPrincipalName | format-table | out-string)

$xUsername = fCollectUPN -xCurrent $true -xText "Enter the User Login email address to search on"

$xPermsList = get-mailbox | Get-MailboxPermission | ? { $_.user -eq $xUsername}

write-host ($xPermsList| format-table | out-string)

fExportCSV -xInput $xPermsList -xFilename "PermsList"

}


	#Clutter
function global:fshowclutterall {

fdisplayinfo -xtext "Collecting Mailbox Data"
$mailboxes=get-mailbox
fdisplayinfo -xtext "Collecting Clutter Data...This may take a while"
$xhash=$null
$xhash=@{}
foreach($xmailbox in $mailboxes) {
		if 	( $xmailbox.alias -Like "DiscoverySearch*" ) {
				#write-host "SKIPPING: "$xmailbox.UserPrincipalName
			} else {
				#write-host "Processing: "$xmailbox.UserPrincipalName
				$xhash.add($xmailbox.UserPrincipalName,(get-clutter -identity $xmailbox.UserPrincipalName).isenabled)
			}
	}

$xOutObject = $xhash.getenumerator() | foreach {new-object psobject -Property @{Name = $_.name;Enabled=$_.value}}

clear

write-host ($xhash | sort Name | ft | out-string )

fExportCSV -xInput $xoutobject -xFilename "ClutterList"

}

function global:fdisableclutterall {

fdisplayinfo -xtext "Disabling Clutter for all users...this may take a while"

$mailboxes=get-mailbox

foreach($xmailbox in $mailboxes) {
		$null = set-clutter -Identity $xmailbox -Enable $false
		write-host "Disabling: "$xmailbox.UserPrincipalName
	}

fshowclutterall

}

function global:fenableclutterall {

fdisplayinfo -xtext "Enabling Clutter for all users...this may take a while"
$null = get-mailbox | set-clutter -Enable $true

fshowclutterall

}

	#Focused Inbox

function global:fdisablefocusedall {

fdisplayinfo -xtext "Disabling Focused Inbox for all users...this may take a while"
$null = Set-OrganizationConfig -FocusedInboxOn $false

write-host (Get-OrganizationConfig | select FocusedInboxOn | fl | out-string)
pause
}

function global:fenablefocusedall {

fdisplayinfo -xtext "Enabling Focused Inbox for all users...this may take a while"
$null = Set-OrganizationConfig -FocusedInboxOn $true

write-host (Get-OrganizationConfig | select FocusedInboxOn | fl | out-string)
pause

}


	#Mailbox Services 
function global:fToggleMAPI {
 	fDisplayInfo -xtext "Toggle MAPI Access"
	while (!$xIdentity) {
		$xInput = fUserPrompt -xQuestion "Enter either a User ID, 'EnableAll', 'DisableAll' or 'Quit'"
		if ((fCheckIdentity -id $xInput) -OR ($xInput -eq "enableall") -OR ($xInput -eq "disableall")) {
			$xIdentity = $xInput
			remove-variable -name xInput
		} elseif ($xInput -eq "quit") {
			Return
		} else 	{
			fDisplayInfo -xText "Invalid Selection" -xColour "red"
		}
	}
	
		
	if ($xIdentity -eq "disableall") {
		Get-Mailbox -ResultSize Unlimited | Set-CASMailbox -MAPIenabled $False
		write-host (get-casmailbox | select name, MAPIenabled | ft | out-string)
	} elseif ($xIdentity -eq "enableall") {
		Get-Mailbox -ResultSize Unlimited | Set-CASMailbox -MAPIenabled $true
		write-host (get-casmailbox | select name, MAPIenabled | ft | out-string)
	} else {
	
		write-host (get-casmailbox -identity $xIdentity | select name, MAPIenabled | ft | out-string)
		
		if ((get-casmailbox $xIdentity | select MAPIenabled).MAPIenabled -eq $true) {
			$xToggle = fUserPrompt -xQuestion "Would you like to disable MAPI Access? (y/n)"
			if ($xToggle -eq "y") {
				Set-CASMailbox -identity $xIdentity -MAPIenabled $False
			} else {
				fDisplayInfo -xText "No changes have been made"
			}
		} elseif ((get-casmailbox $xIdentity | select MAPIenabled).MAPIenabled -eq $false) {
			$xToggle = fUserPrompt -xQuestion "Would you like to enable MAPI Access? (y/n)"
			if ($xToggle -eq "y") {
				Set-CASMailbox -identity $xIdentity -MAPIenabled $true
			} else {
				fDisplayInfo -xText "No changes have been made"
			}
		}
	
		write-host (get-casmailbox -identity $xIdentity | select name, MAPIenabled | ft | out-string)
		pause
		
	}	
}

function global:fToggleActiveSync {
 	fDisplayInfo -xtext "Toggle ActiveSync Access"
	while ((!$xIdentity) -AND (!$xResponce)) {
		$xInput = fUserPrompt -xQuestion "Enter either a User ID, 'EnableAll', 'DisableAll' or 'Quit'"
		if (($xInput -eq "enableall") -OR ($xInput -eq "disableall")) {
			$xResponce = $xInput
		} elseif ((fCheckIdentity -id $xInput)){
			$xIdentity = $xInput
			remove-variable -name xInput		
		} elseif ($xInput -eq "quit") {
			Return
		} else 	{
			fDisplayInfo -xText "Invalid Selection" -xColor "red" -xTime 1
			fDidYouMean -xInput $xInput
			remove-variable -name xInput
		}
	}
	
		
	if ($xResponce -eq "disableall") {
		remove-variable -name xInput
		Get-Mailbox -ResultSize Unlimited | Set-CASMailbox -ActiveSyncenabled $False
		write-host (get-casmailbox | select name, ActiveSyncenabled | ft | out-string)
	} elseif ($xResponce -eq "enableall") {
		remove-variable -name xInput
		Get-Mailbox -ResultSize Unlimited | Set-CASMailbox -ActiveSyncenabled $true
		write-host (get-casmailbox | select name, ActiveSyncenabled | ft | out-string)
	} else {
	
		write-host (get-casmailbox -identity $xIdentity | select name, ActiveSyncenabled | ft | out-string)
		
		if ((get-casmailbox $xIdentity | select ActiveSyncenabled).ActiveSyncenabled -eq $true) {
			$xToggle = fUserPrompt -xQuestion "Would you like to disable ActiveSync Access? (y/n)"
			if ($xToggle -eq "y") {
				Set-CASMailbox -identity $xIdentity -ActiveSyncenabled $False
			} else {
				fDisplayInfo -xText "No changes have been made"
			}
		} elseif ((get-casmailbox $xIdentity | select ActiveSyncenabled).ActiveSyncenabled -eq $false) {
			$xToggle = fUserPrompt -xQuestion "Would you like to enable ActiveSync Access? (y/n)"
			if ($xToggle -eq "y") {
				Set-CASMailbox -identity $xIdentity -ActiveSyncenabled $true
			} else {
				fDisplayInfo -xText "No changes have been made"
			}
		}
	
		write-host (get-casmailbox -identity $xIdentity | select name, ActiveSyncenabled | ft | out-string)
		pause
		
	}	
}

function global:fToggleOWA {
 	fDisplayInfo -xtext "Toggle OWA Access"
	while (!$xIdentity) {
		$xInput = fUserPrompt -xQuestion "Enter either a User ID, 'EnableAll', 'DisableAll' or 'Quit'"
		if ((fCheckIdentity -id $xInput) -OR ($xInput -eq "enableall") -OR ($xInput -eq "disableall")) {
			$xIdentity = $xInput
			remove-variable -name xInput
		} elseif ($xInput -eq "quit") {
			Return
		} else 	{
			fDisplayInfo -xText "Invalid Selection" -xColour "red"
		}
	}
	
		
	if ($xIdentity -eq "disableall") {
		Get-Mailbox -ResultSize Unlimited | Set-CASMailbox -OWAenabled $False
		write-host (get-casmailbox | select name, OWAenabled | ft | out-string)
	} elseif ($xIdentity -eq "enableall") {
		Get-Mailbox -ResultSize Unlimited | Set-CASMailbox -OWAenabled $true
		write-host (get-casmailbox | select name, OWAenabled | ft | out-string)
	} else {
	
		write-host (get-casmailbox -identity $xIdentity | select name, OWAenabled | ft | out-string)
		
		if ((get-casmailbox $xIdentity | select OWAenabled).OWAenabled -eq $true) {
			$xToggle = fUserPrompt -xQuestion "Would you like to disable OWA Access? (y/n)"
			if ($xToggle -eq "y") {
				Set-CASMailbox -identity $xIdentity -OWAenabled $False
			} else {
				fDisplayInfo -xText "No changes have been made"
			}
		} elseif ((get-casmailbox $xIdentity | select OWAenabled).OWAenabled -eq $false) {
			$xToggle = fUserPrompt -xQuestion "Would you like to enable OWA Access? (y/n)"
			if ($xToggle -eq "y") {
				Set-CASMailbox -identity $xIdentity -OWAenabled $true
			} else {
				fDisplayInfo -xText "No changes have been made"
			}
		}
	
		write-host (get-casmailbox -identity $xIdentity | select name, OWAenabled | ft | out-string)
		pause
		
	}	
}

function global:fToggleImap {
 	fDisplayInfo -xtext "Toggle IMAP Access"
	while (!$xIdentity) {
		$xInput = fUserPrompt -xQuestion "Enter either a User ID, 'EnableAll', 'DisableAll' or 'Quit'"
		if ((fCheckIdentity -id $xInput) -OR ($xInput -eq "enableall") -OR ($xInput -eq "disableall")) {
			$xIdentity = $xInput
			remove-variable -name xInput
		} elseif ($xInput -eq "quit") {
			Return
		} else 	{
			fDisplayInfo -xText "Invalid Selection" -xColour "red"
		}
	}
	
		
	if ($xIdentity -eq "disableall") {
		Get-Mailbox -ResultSize Unlimited | Set-CASMailbox -imapenabled $False
		write-host (get-casmailbox | select name, imapenabled | ft | out-string)
	} elseif ($xIdentity -eq "enableall") {
		Get-Mailbox -ResultSize Unlimited | Set-CASMailbox -imapenabled $true
		write-host (get-casmailbox | select name, imapenabled | ft | out-string)
	} else {
	
		write-host (get-casmailbox -identity $xIdentity | select name, imapenabled | ft | out-string)
		
		if ((get-casmailbox $xIdentity | select imapenabled).imapenabled -eq $true) {
			$xToggle = fUserPrompt -xQuestion "Would you like to disable IMAP Access? (y/n)"
			if ($xToggle -eq "y") {
				Set-CASMailbox -identity $xIdentity -imapenabled $False
			} else {
				fDisplayInfo -xText "No changes have been made"
			}
		} elseif ((get-casmailbox $xIdentity | select imapenabled).imapenabled -eq $false) {
			$xToggle = fUserPrompt -xQuestion "Would you like to enable IMAP Access? (y/n)"
			if ($xToggle -eq "y") {
				Set-CASMailbox -identity $xIdentity -imapenabled $true
			} else {
				fDisplayInfo -xText "No changes have been made"
			}
		}
	
		write-host (get-casmailbox -identity $xIdentity | select name, imapenabled | ft | out-string)
		pause
		
	}	
}

function global:fTogglePop {
 	fDisplayInfo -xtext "Toggle POP Access"
	while (!$xIdentity) {
		$xInput = fUserPrompt -xQuestion "Enter either a User ID, 'EnableAll', 'DisableAll' or 'Quit'"
		if ((fCheckIdentity -id $xInput) -OR ($xInput -eq "enableall") -OR ($xInput -eq "disableall")) {
			$xIdentity = $xInput
			remove-variable -name xInput
		} elseif ($xInput -eq "quit") {
			Return
		} else 	{
			fDisplayInfo -xText "Invalid Selection" -xColour "red"
		}
	}
	
		
	if ($xIdentity -eq "disableall") {
		Get-Mailbox -ResultSize Unlimited | Set-CASMailbox -popenabled $False
		write-host (get-casmailbox | select name, popenabled | ft | out-string)
	} elseif ($xIdentity -eq "enableall") {
		Get-Mailbox -ResultSize Unlimited | Set-CASMailbox -popenabled $true
		write-host (get-casmailbox | select name, popenabled | ft | out-string)
	} else {
	
		write-host (get-casmailbox -identity $xIdentity | select name, popenabled | ft | out-string)
		
		if ((get-casmailbox $xIdentity | select popenabled).popenabled -eq $true) {
			$xToggle = fUserPrompt -xQuestion "Would you like to disable Pop Access? (y/n)"
			if ($xToggle -eq "y") {
				Set-CASMailbox -identity $xIdentity -popenabled $False
			} else {
				fDisplayInfo -xText "No changes have been made"
			}
		} elseif ((get-casmailbox $xIdentity | select popenabled).popenabled -eq $false) {
			$xToggle = fUserPrompt -xQuestion "Would you like to enable Pop Access? (y/n)"
			if ($xToggle -eq "y") {
				Set-CASMailbox -identity $xIdentity -popenabled $true
			} else {
				fDisplayInfo -xText "No changes have been made"
			}
		}
	
		write-host (get-casmailbox -identity $xIdentity | select name, popenabled | ft | out-string)
		pause
		
	}	
}

function global:fDisplayCASMailboxStatus {

	$xIdentity = fCollectIdentity -xText "Who would you like to see the status for?"
	write-host (get-casmailbox -identity $xIdentity | ft | out-string)
	pause
}


#Distribution Groups =======================================================================================

function global:fAddNewDistGroup {
	fStoreMainMenu -xRestore 0
	$xgroupname = fUserPrompt -xQuestion "Enter the group alias: "  
			
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
	$xvar = use-menu -MenuHash $xMenuHash -Title "Select Domain" -NoSplash $True
	
	$xDisExtAcc = fUserPrompt -xQuestion "Would you like to disable external mail to this address? (y/n)" 
	if ($xDisExtAcc -eq "y") { Set-DistributionGroup -Identity $xGroupName -RequireSenderAuthenticationEnabled $true }
	
	Write-Host (get-DistributionGroup -Identity $xGroupName | select Name, RequireSenderAuthenticationEnabled, PrimarySmtpAddress | format-table | out-string)
	remove-variable xgroupname, xmenuhash, xvar, xDisExtAcc
	fStoreMainMenu -xRestore 1
	fDisplayInfo -xText "You will now be taken to add new members to the group"
	pause
	fAddUserDistGroup
}

function global:fListDistMembers {
PARAM(
[string]$xGroupName
)
	if (!$xGroupName) {
		Get-DistributionGroup | sort DisplayName | foreach-object {
			Write-host $($_.Displayname)
			if ($($_.RequireSenderAuthenticationEnabled) -eq $true) {Write-host "*Internal Emails Only*"}
			write-host "===========" 
			#Get-DistributionGroupMember $($_.DisplayName) | foreach-object {
			#	write-host $_.DisplayName
			#}
			write-host (Get-DistributionGroupMember $($_.PrimarySmtpAddress) | select DisplayName | sort DisplayName | format-table | out-string)
					write-host "`n" 
		}
		pause
	}else{
		write-host $xGroupName"`n==========="
		write-host (Get-DistributionGroupMember $xGroupName | select DisplayName | sort DisplayName | format-table | out-string)
	}	
	
}

function global:fToggleDistHideFromGAL {
	$xName = fCollectIdentity -xText "Enter the group you would like to hide"
	if ($xName -eq $false) {Return $false}
	$xStatus = Get-DistributionGroup -identity $xName | select HiddenFromAddressListsEnabled
	if ($xStatus.HiddenFromAddressListsEnabled) {
		$xUnhide = "x"
		$xUnhide = fUserPrompt -xQuestion "Would you like to unhide the Group? (y/n)"
		if ($xUnhide -eq "n") {Return $false}
		if ($xUnHide -eq "y") {$xHidden = $false}
	} else {
		$xhide = "x"
		$xhide = fUserPrompt -xQuestion "Would you like to hide the Group? (y/n)"
		if ($xHide -eq "n") {Return $false}
		if ($xHide -eq "y") {$xHidden = $true}
	}
	Set-DistributionGroup -identity $xName -HiddenFromAddressListsEnabled $xHidden
	write-host ( Get-DistributionGroup -identity $xName | select Displayname, HiddenFromAddressListsEnabled | format-list | out-string )
	pause
}

function global:fViewDistExtAuth {

	$xExtAuth = Get-DistributionGroup  | ft identity, RequireSenderAuthenticationEnabled
	write-host ( $xExtAuth | out-string)
	fExportCSV -xInput $xExtAuth -XFilename "DistExtAuth"

}

function global:fToggleDistExtAuth {

	$xName = fCollectIdentity -xText "Enter the group you would like to toggle"
	if ($xName -eq $false) {Return $false}
	
	$xStatus = Get-DistributionGroup -identity $xName | select Displayname, RequireSenderAuthenticationEnabled
	
	fDisplayInfo -xText "The Current Status is:"
	write-host ( $xStatus | format-list | out-string )
	
	if ($xStatus.RequireSenderAuthenticationEnabled) {
		$xExt = "x"
		$xExt = fUserPrompt -xQuestion "Would you like to ALLOW external users? (y/n)"
		if ($xExt -eq "n") {Return $false}
		if ($xExt -eq "y") {$xEnabled = $false}
	} else {
		$xhide = "x"
		$xhide = fUserPrompt -xQuestion "Would you like to BLOCK external users? (y/n)"
		if ($xHide -eq "n") {Return $false}
		if ($xHide -eq "y") {$xEnabled = $true}
	}
	Set-DistributionGroup -identity $xName -RequireSenderAuthenticationEnabled $xEnabled
	
	fDisplayInfo -xText "The New Status is:"
	write-host ( Get-DistributionGroup -identity $xName | select Displayname, RequireSenderAuthenticationEnabled | format-list | out-string )
	pause
	
}


	#Add remove user
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
	
	#Alias
function global:fAddDistGroupEmailAlias {

	$xGroupName = fCollectIdentity -xText "Enter Group Name:"
	if ($xGroupName -eq $false) {Return $false}
		
	$xDistGroup = get-distributiongroup -identity $xGroupName 
	$xEmails = $xDistGroup.EmailAddresses
	$i = 1
	fDisplayInfo -xText "The current emails attached to this group are"
	foreach ($email in $xEmails) {
		Write-Host "`t`t"$i" - "$email
	}
	write-host 

	$xNewEmailAddress = fCollectUPN -xText "Enter the additional email address:" -xCurrent $false
	if ($xNewEmailAddress -eq $false) {Return}
		
	Set-DistributionGroup $xGroupName -emailaddresses @{Add=$xNewEmailAddress}
	
	$xDistGroup = get-distributiongroup -identity $xGroupName 
	$xEmails = $xDistGroup.EmailAddresses
	$i = 1
	fDisplayInfo -xText "The current emails attached to this group are now"
	foreach ($email in $xEmails) {
		Write-Host "`t`t"$i" - "$email
	}
	pause
}

function global:fRemoveDistGroupEmailAlias {

	$xGroupName = fCollectIdentity -xText "Enter Group Name:"
	if ($xGroupName -eq $false) {Return $false}
	if ($xGroupName -eq $false) {
		Return
	}
	
	$xDistGroup = get-distributiongroup -identity $xGroupName 
	$xEmails = $xDistGroup.EmailAddresses
	$i = 1
	fDisplayInfo -xText "The current emails attached to this group are"
	foreach ($email in $xEmails) {
		Write-Host "`t`t"$i" - "$email
	}
	write-host 
	
	$xNewEmailAddress = fCollectUPN -xText "Enter the email address to remove" -xCurrent $false
	if ($xNewEmailAddress -eq $false) {Return}
	
	Set-DistributionGroup $xGroupName -emailaddresses @{Remove=$xNewEmailAddress}
	
	$xDistGroup = get-distributiongroup -identity $xGroupName 
	$xEmails = $xDistGroup.EmailAddresses
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

#Mail Flow and Spam =======================================================================================

function global:fGetTranStatus {
	write-host (get-TransportRule | select name, state | format-table | out-string)
	pause

}

function global:fToggleTransportRule {
	fStoreMainMenu -xRestore 0

	function global:fToggleTransportRuleRun {
	param(
	[string]$xIdentity
	)
		fDisplayInfo -xText "Toggling Transport Rule"
		$xTR = Get-TransportRule -Identity $xIdentity 
		
		if ($xTR.State -eq "Enabled") {
			Disable-TransportRule -Identity $xIdentity
			write-host (Get-TransportRule -Identity $xIdentity | select name, state | format-table | out-string)	
			pause
		}elseif ($xTR.State -eq "Disabled") {
			Enable-TransportRule -Identity $xIdentity
			write-host (Get-TransportRule -Identity $xIdentity | select name, state | format-table | out-string)	
			pause
		}else{
			Write-error "Failed to Establish if $xTRState was enabled or disabled in function global:fToggleTransportRuleRun"
			pause
		}
		fStoreMainMenu -xRestore 1
	}
	
	$xTRMenuHash = New-Object System.Collections.HashTable
	Get-TransportRule | sort-object Name | foreach-object {
			$xTRMenuHash.add($_.Name+" - "+$_.State,'fToggleTransportRuleRun -xIdentity "'+$_.Name+'"')
		}
	use-menu -MenuHash $xTRMenuHash -Title "Select a Transport Rule" -NoSplash 1
	fStoreMainMenu -xRestore 1
}

function global:fSetSafelistEnable {

	fDisplayInfo -xText "Setting the Organisation Connection Filter"

	Set-HostedConnectionFilterPolicy -Identity Default -EnableSafeList $true
	write-host (Get-HostedConnectionFilterPolicy -Identity Default | select Identity, EnableSafeList | format-table | out-string)
	pause

	fDisplayInfo -xText "Setting each Mailbox to Accept Mail from Contacts"

	get-mailbox | set-MailboxJunkEmailConfiguration -ContactsTrusted $true
	write-host (get-mailbox | get-MailboxJunkEmailConfiguration | ft identity, ContactsTrusted | format-table | out-string)
	pause

}

function global:fcreatewarningrules {

	fDisplayInfo -xText "Setting Warning Rules. Please review in Exchange Admin Centre after setup"
	
	$bodywords = "forms.office.com", "validate email", "office team", "verify email", "confirm your account", "office 365"

	$headerwords = "onmicrosoft", "microsoft", "outlook"

	$disclaimer = '<font color="red">***External Email***<br/> This message may or may not be legitimate, please treat links/attachments with caution, especially if they ask you to enter login credentials.<br/> If you have any doubts, please forward to support@tetrabyte.com for verification.</font><br/> ---<br/><br/>'

	new-transportrule -Name "External Warning - Body" -FromScope NotInOrganization -SubjectOrBodyContainsWords $bodywords -ApplyHtmlDisclaimerLocation Prepend -ApplyHtmlDisclaimerText $disclaimer -ApplyHtmlDisclaimerFallbackAction wrap 

	new-transportrule -Name "External Warning - Header" -FromScope NotInOrganization -HeaderContainsMessageHeader from -HeaderContainsWords $headerwords -ApplyHtmlDisclaimerLocation Prepend -ApplyHtmlDisclaimerText $disclaimer -ApplyHtmlDisclaimerFallbackAction wrap

	fGetTranStatus

}

function global:fSenderMessageTrace {

fDisplayInfo -xText "Limited to 50,000 results over 7 days"

$dateEnd = get-date

$xMinusDays = fUserPrompt -xQuestion "How many day past do you wish to search (Max 7): "
if ($xMinusDays -gt 7) { 
	$xMinusDays = 7
	fDisplayInfo -xText "Limited to 50,000 results over 7 days - Setting 7 Days"
	}
	
$dateStart = $dateEnd.AddDays(0-$xMinusDays)

$page = 1
$SenderAddress = "*"
$SenderAddress = fUserPrompt -xQuestion "Enter the Sender Address in Format: sender@domain.com or *@domain.com"

while ($page -lt 10) {
	fDisplayInfo -xText "Processing Part $page of 10"
	$xVar = Get-MessageTrace -pagesize 5000 -page $page -StartDate $dateStart -EndDate $dateEnd | where{$_.SenderAddress -like $SenderAddress} | Select-Object Received, SenderAddress, RecipientAddress, Subject, Status, FromIP, Size
	$xResult += $xVar
	$page++
	}

$xResult | Out-GridView

fDisplayInfo -xText "Complete"

fExportCSV -xInput $xResult -xFilename "MessageTrace"

}

# Partner ==========================

function global:fListAllActive {


$tenids = (get-msolpartnercontract).tenantid.guid

$xInput = foreach ($ten in $tenids) {get-msoluser -tenantid $ten | ?{$_.licenses -ne $null} | sort licenses | select UserPrincipalName, Licenses| ft UserPrincipalName, Licenses}
	
	y
	
	write-host ($xInput | format-table | out-string)
	
	fExportTXT -xInput $xInput -xFilename "UserList"

}





if (!$global:ForceLoginFirst) {$global:ForceLoginFirst = $false}
