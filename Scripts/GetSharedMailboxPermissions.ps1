function Get-UserMailboxPermissions {
    param (
        [string]$UserIdentity,
        [string]$DomainFilter,
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
        
        # Apply domain filter if specified
        if (-not [string]::IsNullOrWhiteSpace($DomainFilter)) {
            Write-Host "Filtering mailboxes for domain: $DomainFilter" -ForegroundColor Cyan
            $mailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object { $_.PrimarySmtpAddress -like "*@$DomainFilter" }
        } else {
            $mailboxes = Get-Mailbox -ResultSize Unlimited
        }
        
        Write-Host "Found $($mailboxes.Count) mailboxes to check." -ForegroundColor Green
        
        # Counter for progress display
        $counter = 0
        $totalMailboxes = $mailboxes.Count

        foreach ($mailbox in $mailboxes) {
            $counter++
            $percentComplete = [math]::Round(($counter / $totalMailboxes) * 100)
            Write-Progress -Activity "Checking mailbox permissions" -Status "Processing $counter of $totalMailboxes ($percentComplete%)" -PercentComplete $percentComplete
            
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
                # Silently continue if calendar permissions cannot be checked
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
        
        # Clear progress bar
        Write-Progress -Activity "Checking mailbox permissions" -Completed

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
Write-Host "`nConnected to: $($script:ExchangeConnection.OrganizationName)" -ForegroundColor Green
Write-Host "User: $($script:ExchangeConnection.CurrentUser)`n" -ForegroundColor Green

do {
    Write-Host "`nOptions:" -ForegroundColor Cyan
    Write-Host "1. Check User's Mailbox Permissions (All Domains)"
    Write-Host "2. Check User's Mailbox Permissions (Specific Domain)"
    Write-Host "3. Return to Main Menu"
    
    $choice = Read-Host "`nEnter your choice (1-3)"
    
    if ($choice -eq '3') { return }
    
    if ($choice -in '1','2') {
        $userIdentity = Read-Host "`nEnter the email address of the user to check permissions for"
        
        $domainFilter = $null
        if ($choice -eq '2') {
            $domainFilter = Read-Host "Enter domain to filter mailboxes (e.g., contoso.com)"
        }

        $exportChoice = Read-Host "Do you want to export the results to a CSV file? (Y/N)"
        if ($exportChoice -eq 'Y' -or $exportChoice -eq 'y') {
            $defaultPath = "C:\UserMailboxPermissions_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
            Write-Host "Default export path is: $defaultPath"
            $customPath = Read-Host "Press Enter to use default path or type a custom path"
            $outputFile = if ($customPath) { $customPath } else { $defaultPath }
        } else {
            $outputFile = $null
        }

        Write-Host "`nRetrieving user's mailbox permissions..." -ForegroundColor Cyan
        Get-UserMailboxPermissions -UserIdentity $userIdentity -DomainFilter $domainFilter -OutputFile $outputFile
    }
} while ($true)