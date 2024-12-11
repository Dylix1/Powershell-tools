# Function to connect to Exchange Online
function Connect-ExchangeOnlineSession {
    param (
        [string]$AdminUser
    )
    try {
        # Connect to Exchange Online with specified admin UPN and without showing progress
        Connect-ExchangeOnline -UserPrincipalName $AdminUser -ShowProgress $false
        Write-Host "Successfully connected to Exchange Online as $AdminUser." -ForegroundColor Green
    } catch {
        Write-Host "Error connecting to Exchange Online: $_" -ForegroundColor Red
        exit
    }
}

# Function to hide users from the Global Address List
function Hide-UsersFromGAL {
    param (
        [string]$InputFile
    )
    try {
        # Check if the file exists
        if (-Not (Test-Path -Path $InputFile)) {
            Write-Host "Input file '$InputFile' not found!" -ForegroundColor Red
            exit
        }

        # Read the file and process each user
        $users = Get-Content -Path $InputFile
        Write-Host "Processing $($users.Count) users from the file '$InputFile'." -ForegroundColor Cyan

        $onPremUsers = @() # To store users for whom the operation failed due to on-prem sync
        $alreadyHidden = @() # To store users already hidden from GAL

        foreach ($user in $users) {
            try {
                # Get the user's mailbox properties
                $mailbox = Get-Mailbox -Identity $user -ErrorAction Stop

                # Check if the user is already hidden
                if ($mailbox.HiddenFromAddressListsEnabled -eq $true) {
                    Write-Host "User '$user' is already hidden from the GAL. Skipping..." -ForegroundColor Yellow
                    $alreadyHidden += $user
                    continue
                }

                # Set the user property to hide them from the GAL
                Set-Mailbox -Identity $user -HiddenFromAddressListsEnabled $true -ErrorAction Stop
                Write-Host "Successfully hid user '$user' from the GAL." -ForegroundColor Green
            } catch {
                Write-Host "Error hiding user '$user': $_" -ForegroundColor Red
                if ($_ -match "it's out of the current user's write scope") {
                    $onPremUsers += $user
                }
            }
        }

        # If there are on-prem synced users, generate a new script
        if ($onPremUsers.Count -gt 0) {
            $scriptContent = @"
# This script is for use in an on-premises Active Directory environment
# It updates the 'msExchHideFromAddressLists' attribute for the specified users

Import-Module ActiveDirectory

# List of users to update
\$users = @(
    `"$($onPremUsers -join '", "')`
)

foreach (\$email in \$users) {
    # Find user in Active Directory
    \$user = Get-ADUser -Filter {EmailAddress -eq \$email}

    if (\$user) {
        # Update the attribute
        Set-ADUser -Identity \$user -Replace @{msExchHideFromAddressLists = \$true}
        Write-Host "Successfully updated user: \$email" -ForegroundColor Green
    } else {
        Write-Host "User not found: \$email" -ForegroundColor Red
    }
}
"@

            # Copy the script to the clipboard
            $scriptContent | Set-Clipboard
            Write-Host "`nA new script has been copied to the clipboard for updating on-premises users." -ForegroundColor Yellow
            Write-Host "Paste it into your on-prem PowerShell session and run it." -ForegroundColor Cyan
        }

        # Notify about already hidden users
        if ($alreadyHidden.Count -gt 0) {
            Write-Host "`nThe following users were already hidden from the GAL:" -ForegroundColor Cyan
            $alreadyHidden | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }
        }
    } catch {
        Write-Host "Error processing users from file: $_" -ForegroundColor Red
    }
}

# Function to disconnect from Exchange Online
function Disconnect-ExchangeOnlineSession {
    try {
        Disconnect-ExchangeOnline -Confirm:$false
        Write-Host "Disconnected from Exchange Online." -ForegroundColor Green
    } catch {
        Write-Host "Error disconnecting from Exchange Online: $_" -ForegroundColor Red
    }
}

# Get the directory of the current script
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$inputFile = Join-Path -Path $scriptDirectory -ChildPath "users.txt"

Write-Host "`n=== Exchange Online Hide Users From GAL Script ===" -ForegroundColor Cyan
$AdminUPN = Read-Host "`nEnter your Exchange Online admin email address"

# Main script execution
Write-Host "`nConnecting to Exchange Online..." -ForegroundColor Cyan
Connect-ExchangeOnlineSession -AdminUser $AdminUPN

Write-Host "`nHiding users from the Global Address List..." -ForegroundColor Cyan
Hide-UsersFromGAL -InputFile $inputFile

Write-Host "`nCleaning up..." -ForegroundColor Cyan
Disconnect-ExchangeOnlineSession

Write-Host "`nScript execution completed." -ForegroundColor Green