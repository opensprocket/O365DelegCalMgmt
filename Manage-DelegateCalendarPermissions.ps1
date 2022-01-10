# Last updated 10 Jan 2022
# Requires Powershell 7 - will not work otherwise!

# Global Variables
[bool]$continue = $true

# Functions
function Get-CalendarPermissions($Identity) {
    Get-MailboxFolderPermission -Identity "${Identity}:\calendar" | Format-Table
}

function Set-DelegateCalendarPermission($Identity, $DelegateIdentity) {
    try {
        Set-MailboxFolderPermission -Identity "${Identity}:\calendar" -User $DelegateIdentity -AccessRights Editor -SharingPermissionFlags Delegate -SendNotificationToUser #-Confirm   
    }
    catch {
        Write-Host "An error has occured while trying to add Delegate access for $DelegateIdentity to the Calendar of $Identity"  
    }
    
}

function Remove-CalendarDelegatePermission($Identity, $DelegateIdentity) {
    try {
        Remove-MailboxFolderPermission -Identity ${Identity} -User $DelegateIdentity -Confirm -ErrorAction SilentlyContinue #-Confirm
    }
    catch {
        Write-Host -Object "$DelegateIdentity does not have Delegate access to the calendar of $Identity"
    }
    
}

Function Start-O365ConnectionModernAuth {
    # Inform user of prerequisites if applicable
    if ($psversiontable.PSVersion -lt 7) {
      Write-Host -Object "Requires Powershell 7 or newer!" -ForegroundColor Green -BackgroundColor Black
      Write-Host -Object "Requires the Exchange Online Management Powershell Module!" -ForegroundColor Green -BackgroundColor Black
      Write-Host -Object "Connection aborted." -ForegroundColor Green -BackgroundColor Black
    }
    
    # Pass the version check, import the module
    else {
      Import-Module ExchangeOnlineManagement
      $username = Read-Host -Prompt "Please input your full email address"
      Connect-ExchangeOnline -UserPrincipalName $username
    } 
}
  
function Stop-O365Connection {
    # Write-Host -Object "Which session do you want to disconnect?" -ForegroundColor Green -BackgroundColor Black
    # Get-PSSession | Format-Table
    # Write-Host -Object "ID of session to disconnect" -ForegroundColor Green -BackgroundColor Black
    # $SessionID = Read-Host 
    try {
      Get-PSSession | Remove-PSSession -Id $SessionID
      Write-Host
      Write-Host -Object "The Session ID $SessionID has been disconnected."
    }
    catch {
        Write-Host "Unable to remove PSSession. Please check and clear manually."
        Write-Host $_  
    } 
}

function Show-CommandMenu {
    Write-Host -Object "`n---- Commands ----"
    Write-Host -Object "add - Add delegate `nchk - Check access `nrm - Remove delegate `nstop/exit - Exit"
    Write-Host -Object "------------------`n"

}

# Begin TUI

# Start O365 Connection
Start-O365ConnectionModernAuth
Clear-Host

# Loop until terminated
do {
    Show-CommandMenu
    [string]$action = Read-Host -Prompt "What would you like to do?"

    if ($action.ToLower() -notlike "stop" -and $action.ToLower() -notlike "exit") {

        # Begin Switch
        switch ($action.ToLower()) {
            "add" {
                $identity = Read-Host -Prompt "Please input the username of the person you want to modify"
                $delegate = Read-Host -Prompt "Please input the username of the person that will recieve delegate access"

                Write-Host -Object "This will grant ${delegate} delegated access to the calendar of ${identity}."
                $verify = Read-Host -Prompt "Do you want to continue? y/n"
                if ($verify.ToLower() -notlike "y") {
                    Write-Host "Changes aborted."
                }
                else {
                    Set-DelegateCalendarPermission($identity, $delegate)
                    Write-Host -Object "Changes committed. ${delegate} now has delegate access to the calendar of ${identity}"
                }
            }
            "a" {
                $identity = Read-Host -Prompt "Please input the username of the person you want to modify"
                $delegate = Read-Host -Prompt "Please input the username of the person that will recieve delegate access"

                Write-Host -Object "This will grant ${delegate} delegated access to the calendar of ${identity}."
                $verify = Read-Host -Prompt "Do you want to continue? y/n"
                if ($verify.ToLower() -notlike "y") {
                    Write-Host "Changes aborted."
                }
                else {
                    Set-DelegateCalendarPermission($identity, $delegate)
                    Write-Host -Object "Changes committed. ${delegate} now has delegate access to the calendar of ${identity}"
                }
            }
            "chk" {
                $identity = Read-Host -Prompt "Please input the username of the person you want to check"
                Get-CalendarPermissions($identity)
            }
            "c" {
                $identity = Read-Host -Prompt "Please input the username of the person you want to check"
                Get-CalendarPermissions($identity)
            }
            "rm" {
                $identity = Read-Host -Prompt "Please input the username of the person you want to modify"
                $delegate = Read-Host -Prompt "Please input the username of the person that will have delegate access removed"

                Write-Host -Object "This will remove ${delegate}'s access to the calendar of ${identity}."
                $verify = Read-Host -Prompt "Do you want to continue? y/n"
                if ($verify.ToLower() -notlike "y") {
                    Write-Host "Changes aborted."
                }
                else {
                    Set-DelegateCalendarPermission($identity, $delegate)
                    Write-Host -Object "Changes committed. ${delegate}'s delegate access has been removed from the calendar of ${identity}"
                }
            }
            "r" {
                $identity = Read-Host -Prompt "Please input the username of the person you want to modify"
                $delegate = Read-Host -Prompt "Please input the username of the person that will have delegate access removed"

                Write-Host -Object "This will remove ${delegate}'s access to the calendar of ${identity}."
                $verify = Read-Host -Prompt "Do you want to continue? y/n"
                if ($verify.ToLower() -notlike "y") {
                    Write-Host "Changes aborted."
                }
                else {
                    Set-DelegateCalendarPermission($identity, $delegate)
                    Write-Host -Object "Changes committed. ${delegate}'s delegate access has been removed from the calendar of ${identity}"
                }
            }
            Default {
                Write-Host -Object "Command not recognized. Please try again."
            }
        }

    }
    else {
        $continue = $false
        Write-Host "Exiting..."

        # Stop O365 Connection
        Stop-O365Connection

    }

} while ($continue)

## Created by opensprocket ##