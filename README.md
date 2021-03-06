# O365DelegCalMgmt

## Overview

This utility is a standalone PowerShell script that presents a TUI to the user. Built around Microsoft Office 365 PowerShell cmdlets, its purpose is to simplify delegated calendar management. It takes a frequent 5 minute task and accomplishes it in 5 seconds. As usual, here's the [relevant xkcd.](https://xkcd.com/1205/)

## Requirements

Since Microsoft has begun phasing out basic authentication to Office 365 tenants, it is necessary to use the new "Modern Authentication". This method requires a working web browser to sign in to your O365 tenant. In addition that, PowerShell 7 is now required by the EXO-V2 module provided by Microsoft. Make sure to execute this script from the v7 shell window - otherwise it will not work!

- Get PowerShell 7+ from here: [PowerShell Documentation](https://docs.microsoft.com/en-us/powershell/)
- EXO v2 Module documentation: [https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps](https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps)

## How it works:
- Open PowerShell prompt, navigate to directory that contains the script and execute `.\Manage-DelegateCalendarPermissions.ps1`
- Connects to O365 
  - Make sure to use youremail@yourorg.com for the username (when connecting)
- A TUI is presented with four options:
  - Add
  - Check
  - Remove
  - Exit
- Changes are made upon pressing enter, limited error handling is in place to catch any incorrect usernames
- If you want to abort a change, just type exit and press enter (it'll throw an error, but will safely abort)
- When changes are finished, type exit
- The utility will attempt to close the O365 session, however you can do this manually if you type `Get-PSSession` and remove the correct session ID using `Remove-PSSession`

## Additional Notes

If your O365 environment requires using a different URI to connect to it (such as Government, Germany, etc.), this will not work out of the box for you. You'll need to edit the Start-O365 function to use the correct URI. Check the Microsoft documentation for more details on this. 
