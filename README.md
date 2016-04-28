# PowerShell-Office-365-Administration-Script
This is an administration system for office 365 through powershell.
It can be used for quick login with stored details or full administrative tasks from the menu

Just download the 365Login.ps1 file and run with powershell

Once you have run the script, just type 'Start-Login' to get connected to Office 365 Powershell.
You can use 'Use-Admin' to access the main menu.

==Quick Loading of the script:


For ease of access try saving a shortcut
"Start powershell.exe -executionpolicy unrestricted -File C:\PowershellFiles\365Login.ps1"

Else, if you already have 'executionpolicy unrestricted' set, you can run 'Invoke-Item $profile' and add 'C:\PowershellFiles\365Login.ps1' into your profile to be automatically loaded every time you open Powershell.


==Quick Access Tips:

The command 'qqq' is quicker than typing 'start-login'
The command 'www' is quicker than typing 'Use-Admin'
and finally the command 'qq' runs 'start-login' then 'Use-admin' to take to straight to the menu.
