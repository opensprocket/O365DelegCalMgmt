# O365DelegCalMgmt
Tool to manage delegate access to Office 365 calendars written in PowerShell. 

## Overview

This utility is a standalone PowerShell script that presents a TUI to the user. Built around Microsoft Office 365 PowerShell cmdlets, its purpose is to simplify delegated calendar management. It takes a frequent 5 minute task and accomplishes it in 5 seconds. As usual, here's the [relevant xkcd.](https://xkcd.com/1205/)

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

## 
