# Global variable to track connection state
$script:ExchangeConnection = @{
    IsConnected = $false
    CurrentUser = $null
    OrganizationName = $null
}

function Connect-ExchangeOnlineSession {
    param (
        [Parameter(Mandatory=$false)]
        [string]$AdminUser
    )
    
    try {
        # Verify Exchange Online module
        if (!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
            Write-Host "Exchange Online PowerShell module is not installed." -ForegroundColor Red
            Write-Host "Please install it by running: Install-Module -Name ExchangeOnlineManagement" -ForegroundColor Yellow
            return $false
        }

        # Get admin credentials if not provided
        if (-not $AdminUser) {
            $AdminUser = Read-Host "`nEnter your Exchange Online admin email address"
        }

        # Import module and connect
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        Connect-ExchangeOnline -UserPrincipalName $AdminUser -ShowProgress $false -ShowBanner:$false
        
        # Test connection
        try {
            $org = Get-OrganizationConfig -ErrorAction Stop
            $script:ExchangeConnection.IsConnected = $true
            $script:ExchangeConnection.CurrentUser = $AdminUser
            $script:ExchangeConnection.OrganizationName = $org.DisplayName
            Write-Host "`nSuccessfully connected to Exchange Online:" -ForegroundColor Green
            Write-Host "Organization: $($org.DisplayName)" -ForegroundColor Green
            Write-Host "User: $AdminUser" -ForegroundColor Green
            return $true
        } catch {
            Write-Host "`nFailed to verify Exchange connection: $_" -ForegroundColor Red
            return $false
        }
    } catch {
        Write-Host "`nError connecting to Exchange Online:" -ForegroundColor Red
        Write-Host $_ -ForegroundColor Red
        return $false
    }
}

function Get-UserMailboxPermissions {
    param (
        [string]$UserIdentity,
        [string]$OutputFile
    )
    try {
        # Verify user exists
        Write-Host "Verifying user exists..." -ForegroundColor Yellow
        $user = Get-User -Identity $UserIdentity -ErrorAction Stop
        Write-Host "Found user: $($user.DisplayName)" -ForegroundColor Green

        $results = @()
        
        # Get all mailboxes
        Write-Host "`nGetting all mailboxes in the tenant..." -ForegroundColor Yellow
        $mailboxes = Get-Mailbox -ResultSize Unlimited
        Write-Host "Found $($mailboxes.Count) mailboxes to check." -ForegroundColor Green

        foreach ($mailbox in $mailboxes) {
            Write-Host "`nChecking permissions on mailbox: $($mailbox.DisplayName)" -ForegroundColor Yellow
            
            # Check Full Access permissions
            $fullAccess = Get-MailboxPermission -Identity $mailbox.Identity | 
                Where-Object {
                    $_.User -eq $UserIdentity -and 
                    $_.IsInherited -eq $false
                }
            
            # Check Send As permissions
            $sendAs = Get-RecipientPermission -Identity $mailbox.Identity | 
                Where-Object {
                    $_.Trustee -eq $UserIdentity
                }
            
            # Check Send on Behalf permissions
            $sendOnBehalf = $mailbox.GrantSendOnBehalfTo | 
                Where-Object { $_ -eq $UserIdentity }
            
            # Check Calendar permissions
            $calendarPermissions = $null
            try {
                $calendarFolder = Get-MailboxFolderStatistics -Identity $mailbox.Identity -FolderScope Calendar | 
                    Select-Object -First 1
                if ($calendarFolder) {
                    $calendarPath = "$($mailbox.Identity):\$($calendarFolder.Name)"
                    $calendarPermissions = Get-MailboxFolderPermission -Identity $calendarPath |
                        Where-Object { $_.User -eq $UserIdentity }
                }
            } catch {
                Write-Host "Could not check calendar permissions for $($mailbox.DisplayName)" -ForegroundColor Yellow
            }

            # Add permissions to results if any found
            if ($fullAccess) {
                $results += [PSCustomObject]@{
                    Mailbox = $mailbox.DisplayName
                    MailboxType = $mailbox.RecipientTypeDetails
                    EmailAddress = $mailbox.PrimarySmtpAddress
                    AccessType = "Full Access"
                    Details = "Full mailbox access"
                }
            }
            
            if ($sendAs) {
                $results += [PSCustomObject]@{
                    Mailbox = $mailbox.DisplayName
                    MailboxType = $mailbox.RecipientTypeDetails
                    EmailAddress = $mailbox.PrimarySmtpAddress
                    AccessType = "Send As"
                    Details = "Can send as this mailbox"
                }
            }
            
            if ($sendOnBehalf) {
                $results += [PSCustomObject]@{
                    Mailbox = $mailbox.DisplayName
                    MailboxType = $mailbox.RecipientTypeDetails
                    EmailAddress = $mailbox.PrimarySmtpAddress
                    AccessType = "Send on Behalf"
                    Details = "Can send on behalf of this mailbox"
                }
            }

            if ($calendarPermissions) {
                $results += [PSCustomObject]@{
                    Mailbox = $mailbox.DisplayName
                    MailboxType = $mailbox.RecipientTypeDetails
                    EmailAddress = $mailbox.PrimarySmtpAddress
                    AccessType = "Calendar"
                    Details = "Calendar permission level: $($calendarPermissions.AccessRights -join ', ')"
                }
            }
        }

        # Display and export results
        if ($results.Count -eq 0) {
            Write-Host "`nNo mailbox permissions found for user: $($user.DisplayName)" -ForegroundColor Yellow
        } else {
            Write-Host "`nFound $($results.Count) permissions for user: $($user.DisplayName)" -ForegroundColor Green
            $results | Format-Table -AutoSize
        }

        if ($OutputFile -and $results.Count -gt 0) {
            $results | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
            Write-Host "Permissions exported to '$OutputFile'." -ForegroundColor Green
        }

        return $results
    } catch {
        Write-Host "Error retrieving mailbox permissions: $_" -ForegroundColor Red
    }
}

# Main script
Clear-Host
Write-Host "=== Exchange Online User Mailbox Permissions Report ===" -ForegroundColor Cyan

# First, connect with admin account
$adminEmail = Read-Host "Enter your Exchange Online admin email address"
$connected = Connect-ExchangeOnlineSession -AdminUser $adminEmail

if (-not $connected) {
    Write-Host "Failed to connect to Exchange Online. Exiting..." -ForegroundColor Red
    return
}

Write-Host "`nConnected to: $($script:ExchangeConnection.OrganizationName)" -ForegroundColor Green
Write-Host "User: $($script:ExchangeConnection.CurrentUser)`n" -ForegroundColor Green

# Prompt for user
$userIdentity = Read-Host "Enter the email address of the user to check permissions for"

$exportChoice = Read-Host "Do you want to export the results to a CSV file? (Y/N)"
if ($exportChoice -eq 'Y' -or $exportChoice -eq 'y') {
    $defaultPath = "C:\UserMailboxPermissions.csv"
    Write-Host "Default export path is: $defaultPath"
    $customPath = Read-Host "Press Enter to use default path or type a custom path"
    $outputFile = if ($customPath) { $customPath } else { $defaultPath }
} else {
    $outputFile = $null
}

Write-Host "`nRetrieving user's mailbox permissions..." -ForegroundColor Cyan
Get-UserMailboxPermissions -UserIdentity $userIdentity -OutputFile $outputFile

# Disconnect session
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "`nDisconnected from Exchange Online." -ForegroundColor Green