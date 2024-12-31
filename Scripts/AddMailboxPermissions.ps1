# Function to connect to Exchange Online
function Connect-ExchangeOnlineSession {
    param (
        [string]$AdminUser
    )
    try {
        Connect-ExchangeOnline -UserPrincipalName $AdminUser -ShowProgress $false
        Write-Host "Successfully connected to Exchange Online as $AdminUser." -ForegroundColor Green
    } catch {
        Write-Host "Error connecting to Exchange Online: $_" -ForegroundColor Red
        exit
    }
}

# Function to validate mailbox existence
function Test-MailboxExists {
    param (
        [string]$MailboxIdentity
    )
    try {
        $mailbox = Get-Mailbox -Identity $MailboxIdentity -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

# Function to validate user existence
function Test-UserExists {
    param (
        [string]$UserIdentity
    )
    try {
        $user = Get-User -Identity $UserIdentity -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

# Function to add mailbox permission
function Add-CustomMailboxPermission {
    param (
        [string]$MailboxIdentity,
        [string]$UserIdentity,
        [bool]$AutoMapping = $true
    )
    try {
        # Validate mailbox existence
        if (-not (Test-MailboxExists -MailboxIdentity $MailboxIdentity)) {
            Write-Host "Error: Mailbox '$MailboxIdentity' does not exist." -ForegroundColor Red
            return
        }

        # Validate user existence
        if (-not (Test-UserExists -UserIdentity $UserIdentity)) {
            Write-Host "Error: User '$UserIdentity' does not exist." -ForegroundColor Red
            return
        }

        # Check if permission already exists
        $existingPermission = Get-MailboxPermission -Identity $MailboxIdentity -User $UserIdentity -ErrorAction SilentlyContinue
        if ($existingPermission) {
            Write-Host "User '$UserIdentity' already has permissions on mailbox '$MailboxIdentity'." -ForegroundColor Yellow
            return
        }

        # Add permission
        Add-MailboxPermission -Identity $MailboxIdentity -User $UserIdentity -AccessRights FullAccess -AutoMapping $AutoMapping
        Write-Host "Successfully granted FullAccess permission to '$UserIdentity' on mailbox '$MailboxIdentity'." -ForegroundColor Green

        # Verify the permission was added
        $verifyPermission = Get-MailboxPermission -Identity $MailboxIdentity -User $UserIdentity
        if ($verifyPermission) {
            Write-Host "Permission verification successful." -ForegroundColor Green
        } else {
            Write-Host "Warning: Permission was added but verification failed. Please check manually." -ForegroundColor Yellow
        }
    } catch {
        Write-Host "Error adding mailbox permission: $_" -ForegroundColor Red
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

# Main script
Write-Host "`n=== Add Exchange Online Mailbox Permission ===" -ForegroundColor Cyan

# Get all required inputs from user
$mailboxIdentity = Read-Host "Enter the mailbox email address to grant access to"
$userIdentity = Read-Host "Enter the user email address who needs access"
$autoMappingInput = Read-Host "Enable automapping? (Y/N)"
$autoMapping = $autoMappingInput -eq 'Y' -or $autoMappingInput -eq 'y'

# Get admin credentials
$AdminUPN = Read-Host "`nEnter your Exchange Online admin email address"

# Execute the script
Write-Host "`nConnecting to Exchange Online..." -ForegroundColor Cyan
Connect-ExchangeOnlineSession -AdminUser $AdminUPN

Write-Host "`nAdding mailbox permission..." -ForegroundColor Cyan
Add-CustomMailboxPermission -MailboxIdentity $mailboxIdentity -UserIdentity $userIdentity -AutoMapping $autoMapping

Write-Host "`nCleaning up..." -ForegroundColor Cyan
Disconnect-ExchangeOnlineSession

Write-Host "`nScript execution completed." -ForegroundColor Green