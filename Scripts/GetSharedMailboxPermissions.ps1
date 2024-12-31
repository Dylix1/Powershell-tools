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

# Function to get shared mailbox permissions
function Get-SharedMailboxPermissions {
    param (
        [string]$OutputFile
    )
    try {
        # Get all shared mailboxes
        $sharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited
        Write-Host "Found $($sharedMailboxes.Count) shared mailboxes." -ForegroundColor Green

        $results = @()
        foreach ($mailbox in $sharedMailboxes) {
            Write-Host "Processing mailbox: $($mailbox.DisplayName)" -ForegroundColor Yellow
            
            # Get full access permissions
            $fullAccess = Get-MailboxPermission -Identity $mailbox.Identity | 
                Where-Object {$_.User -notlike "NT AUTHORITY\*" -and $_.User -notlike "S-1-5*" -and $_.IsInherited -eq $false}
            
            # Get Send As permissions
            $sendAs = Get-RecipientPermission -Identity $mailbox.Identity | 
                Where-Object {$_.Trustee -notlike "NT AUTHORITY\*" -and $_.Trustee -notlike "S-1-5*"}
            
            # Get Send on Behalf permissions
            $sendOnBehalf = $mailbox.GrantSendOnBehalfTo
            
            # Combine permissions into results
            foreach ($access in $fullAccess) {
                $results += [PSCustomObject]@{
                    SharedMailbox = $mailbox.DisplayName
                    EmailAddress = $mailbox.PrimarySmtpAddress
                    User = $access.User
                    AccessType = "Full Access"
                }
            }
            
            foreach ($access in $sendAs) {
                $results += [PSCustomObject]@{
                    SharedMailbox = $mailbox.DisplayName
                    EmailAddress = $mailbox.PrimarySmtpAddress
                    User = $access.Trustee
                    AccessType = "Send As"
                }
            }
            
            foreach ($user in $sendOnBehalf) {
                $results += [PSCustomObject]@{
                    SharedMailbox = $mailbox.DisplayName
                    EmailAddress = $mailbox.PrimarySmtpAddress
                    User = $user
                    AccessType = "Send on Behalf"
                }
            }
        }

        # Display results in console
        $results | Format-Table -AutoSize

        # Export to CSV if specified
        if ($OutputFile) {
            $results | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
            Write-Host "Permissions exported to '$OutputFile'." -ForegroundColor Green
        }

        return $results
    } catch {
        Write-Host "Error retrieving shared mailbox permissions: $_" -ForegroundColor Red
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
Write-Host "`n=== Exchange Online Shared Mailbox Access Report ===" -ForegroundColor Cyan
$AdminUPN = Read-Host "`nEnter your Exchange Online admin email address"

# Ask if user wants to export to CSV
$exportChoice = Read-Host "Do you want to export the results to a CSV file? (Y/N)"
if ($exportChoice -eq 'Y' -or $exportChoice -eq 'y') {
    $defaultPath = "C:\SharedMailboxPermissions.csv"
    Write-Host "Default export path is: $defaultPath"
    $customPath = Read-Host "Press Enter to use default path or type a custom path"
    $outputFile = if ($customPath) { $customPath } else { $defaultPath }
} else {
    $outputFile = $null
}

# Execute the script
Write-Host "`nConnecting to Exchange Online..." -ForegroundColor Cyan
Connect-ExchangeOnlineSession -AdminUser $AdminUPN

Write-Host "`nRetrieving shared mailbox permissions..." -ForegroundColor Cyan
Get-SharedMailboxPermissions -OutputFile $outputFile

Write-Host "`nCleaning up..." -ForegroundColor Cyan
Disconnect-ExchangeOnlineSession

Write-Host "`nScript execution completed." -ForegroundColor Green