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
	######################################################################>


function global:start-login{
#This script requires the Multi Layered Dynamic Menu System Module from www.AshleyUnwin.com/Powershell_Multi_Layered_Dynamic_Menu_System

if (get-module -name MenuSystem){}else{
	if (Test-Path c:\powershell\MenuSystem.psm1) {
		Import-Module c:\powershell\MenuSystem.psm1
	}elseif (test-path Z:\~Tools\Powershell\MenuSystem.psm1) {
		Import-Module Z:\~Tools\Powershell\MenuSystem.psm1
	}else{
		$source = "https://raw.githubusercontent.com/manicd/Powershell-Multi-Layered-Dynamic-Menu-System/master/MenuSystem.psm1"
		if (test-path c:\powershell\) {
			$destination = "c:\powershell\MenuSystem.psm1"
		}else{
			$destination = "Z:\~Tools\Powershell\MenuSystem.psm1"
		}
		Invoke-WebRequest $source -OutFile $destination
		Import-Module $destination
	}
}

Import-Module MSOnline
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
		write-host "`nYou are now logged in to"$xCompany -Fore Green
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



	$MenuHash2=@{ "Users"=@{			"Reset User Password"="fResetUserPasswords"
										"Discover Licence Names"="fGetMsolAccountSku"
										"Add New User"="fAddNewUser"}
			"Mailboxes"=@{				"Add User to Mailbox Folder"="fAddMailboxFolderPerm"
										"Remove User from Mailbox Folder"="fRemoveMailboxFolderPerm"}
			"Dist Groups"=@{			"List all Dist Groups and Members"="fListDistMembers"
										"Add User to Dist Group"="fadduserdistgroup"
										"Remove User from Dist Group" = "fremoveuserdistgroup"}
			"MSOnline Org"=@{			"View Partner Information"="fViewPartnerInfo"
										"View Domain Info"="fVeiwDomain"}
			"X-Experimental Function"="fExperimentalFunction"							
										
			}
	$title="Office 365 Menu"								
	Use-Menu -MenuHash $MenuHash2 -Title $title -NoSplash 1

}

# Below this line are the functions called by the menu values

function global:fResetUserPasswords {
	
	#Create a function to actually change the password
	function global:fResetUserPasswordsCollectPass {
	PARAM(
	[string]$xUser
	)
		$xString =  "Please enter the new password for "+$xUser+"or type [quit] to quit."
		$xPass = fUserPrompt -xQuestion $xString -xPrompt "Password"
		if ($xPass -ne "quit") {
			if (($xPass -ne "") -AND ($xPass -ne $null)) {
				fDisplayInfo -xText "Setting New Password"
				Write-host
				$xPass = Set-MsolUserPassword -UserPrincipalName $xUser -ForceChangePassword $false -NewPassword $xPass
				fDisplayInfo -xText "Password now set" -xColor "Red" -xTime 3
				$xPass = $null	
				Cls
				} else {
				fDisplayInfo -xText "Password not entered, Nothing has been changed." -xColor "Red"
				fResetUserPasswordsCollectPass -xUser $xUser				
			}
		}else{
		cls
		fDisplayInfo -xText "Quitting....Nothing has been changed." -xColor "Red" -xTime 2
		Cls
		}
	}

	#Create the Menu Hash Table Object	
	$xMenuHash = New-Object System.Collections.HashTable
	#Create Menu Structure Hash Table and set values to be function with UPN as input
	get-msoluser | sort-object UserPrincipalName | select UserPrincipalName | foreach-object {
			$xMenuHash.add($_.UserPrincipalName,"fResetUserPasswordsCollectPass -xUser "+$_.UserPrincipalName)
		}
	#Call the Menu	
	use-menu -MenuHash $xMenuHash -Title "Reset User Password" -NoSplash $True
}
		
function global:fViewPartnerInfo {
	Get-MsolPartnerInformation
}

function global:fVeiwDomain {
	Get-MsolDomain
}

function global:fListDistMembers {
PARAM(
[string]$xGroupName
)
	if (!$xGroupName) {
		Get-DistributionGroup | sort name | foreach-object {
			Write-host $($_.name)"`n===========" 
			Get-DistributionGroupMember $($_.name) | foreach-object {
				write-host $_.Name
			}
					write-host "`n" 
		}
	}else{
		write-host $xGroupName"`n==========="
		Get-DistributionGroupMember $xGroupName | select Name
	}	
}

function global:fAddMailboxFolderPerm {
	cls
	$xUser = fUserPrompt -xQuestion "Enter the User who would like the access"
	$xMailbox = fUserPrompt -xQuestion "Enter the Mailbox they would like access to"
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

	Add-MailboxFolderPermission -Identity $xIdString -User $xUser -AccessRight $xLevel
}

function global:fRemoveMailboxFolderPerm {
	cls
	$xUser = fUserPrompt -xQuestion "Enter the User who would like the access"
	$xMailbox = fUserPrompt -xQuestion "Enter the Mailbox they would like access to"
	$xFolder = fUserPrompt -xQuestion "Enter the Folder the would like access to"
	$xIdString = $xMailbox+":\"+$xFolder
	cls
	$xTextString = "Removing "+$xUser+" from "+$xFolder+" in "+$xMailbox+"'s Mailbox" 
	fDisplayInfo -xText $xTextString
	Remove-MailboxFolderPermission -Identity $xIdString -User $xUser
}

function global:fGetMsolAccountSku {
	Get-MsolAccountSku
}

function global:fAddUserDistGroup {
	$xDistMenuHash = New-Object System.Collections.HashTable
	Get-DistributionGroup | sort-object Name | select Name | foreach-object {
			$xDistMenuHash.add($_.Name,"write-output "+$_.Name)
		}
	[string]$xGroupName  = use-menu -MenuHash $xDistMenuHash -Title "Select a group" -NoSplash 1
	function fAddUserDistGroupWho {
		[string]$xMember  = fUserPrompt -xQuestion "Who would you like to add"
		Add-DistributionGroupMember $xGroupName -Member $xMember -BypassSecurityGroupManagerCheck
	}
	fAddUserDistGroupWho
	$xAddAnother = fUserPrompt -xQuestion "Would you like to add another? (y/n)"
	if ($xAddAnother -eq "y") {
		fAddUserDistGroupWho
		return
	}else{
		fListDistMembers -xGroupName $xGroupName
		return
	}
	
	
}

function global:fRemoveUserDistGroup {
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
	$xMember  = use-menu -MenuHash $xDistMemMenuHash -Title "Select a Member" -NoSplash 1
	
	Remove-DistributionGroupMember $xGroupName -Member $xMember -BypassSecurityGroupManagerCheck
	fListDistMembers -xGroupName $xGroupName
}

function global:fAddNewUser {
	
	$xFirstName = fUserPrompt -xQuestion "First Name"
	$xLastName = fUserPrompt -xQuestion "Last Name"
	$xUPN = fUserPrompt -xQuestion "UserPrincipalName"
	
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
			New-MsolUser -DisplayName $xDisplayName -FirstName $xFirstName -LastName $xLastName -UserPrincipalName $xUPN -LicenseAssignment $xLic -UsageLocation GB -PreferredLanguage "en-GB" –Password $xPass -ForceChangePassword $False
			return
		} else {
			fDisplayInfo -xText "You Do Not currently have enough Licenses to proceed." -xText2 "Please Login to Office 365 and purchase more licences" -xText3 "before proceeding." -xTime 5
			$xTryAgain = fUserPrompt -xQuestion "Try Again? (y/n)"
			if ($xTryAgain -eq "n") {
				return
			} else {
				fAddNewUserLicCheck
				return
			}
		}
	}
	fAddNewUserLicCheck
	
}


# Other functions

function global:fExperimentalFunction{

}
