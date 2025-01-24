function Test-MailboxExists {
    param ([string]$MailboxIdentity)
    try {
        $null = Get-Mailbox -Identity $MailboxIdentity -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

function Test-UserExists {
    param ([string]$UserIdentity)
    try {
        $null = Get-User -Identity $UserIdentity -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

function Add-CustomMailboxPermission {
    param (
        [string]$MailboxIdentity,
        [string]$UserIdentity,
        [bool]$AutoMapping = $true
    )
    try {
        if (-not (Test-MailboxExists -MailboxIdentity $MailboxIdentity)) {
            Write-Host "Error: Mailbox '$MailboxIdentity' does not exist." -ForegroundColor Red
            return
        }

        if (-not (Test-UserExists -UserIdentity $UserIdentity)) {
            Write-Host "Error: User '$UserIdentity' does not exist." -ForegroundColor Red
            return
        }

        $existingPermission = Get-MailboxPermission -Identity $MailboxIdentity -User $UserIdentity -ErrorAction SilentlyContinue
        if ($existingPermission) {
            Write-Host "User '$UserIdentity' already has permissions on mailbox '$MailboxIdentity'." -ForegroundColor Yellow
            return
        }

        Add-MailboxPermission -Identity $MailboxIdentity -User $UserIdentity -AccessRights FullAccess -AutoMapping $AutoMapping
        Write-Host "Successfully granted FullAccess permission to '$UserIdentity' on mailbox '$MailboxIdentity'." -ForegroundColor Green

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

# Main script
Clear-Host
Write-Host "=== Add Exchange Online Mailbox Permission ===" -ForegroundColor Cyan
Write-Host "`nConnected to: $($script:ExchangeConnection.OrganizationName)" -ForegroundColor Green
Write-Host "User: $($script:ExchangeConnection.CurrentUser)`n" -ForegroundColor Green

do {
    Write-Host "`nOptions:" -ForegroundColor Cyan
    Write-Host "1. Add Mailbox Permission"
    Write-Host "2. Return to Main Menu"
    
    $choice = Read-Host "`nEnter your choice (1-2)"
    
    if ($choice -eq '2') { return }
    
    if ($choice -eq '1') {
        $mailboxIdentity = Read-Host "`nEnter the mailbox email address to grant access to"
        $userIdentity = Read-Host "Enter the user email address who needs access"
        $autoMappingInput = Read-Host "Enable automapping? (Y/N)"
        $autoMapping = $autoMappingInput -eq 'Y' -or $autoMappingInput -eq 'y'

        Write-Host "`nAdding mailbox permission..." -ForegroundColor Cyan
        Add-CustomMailboxPermission -MailboxIdentity $mailboxIdentity -UserIdentity $userIdentity -AutoMapping $autoMapping
    }
} while ($true)