function Get-SharedMailboxPermissions {
    param ([string]$OutputFile)
    try {
        $sharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited
        Write-Host "Found $($sharedMailboxes.Count) shared mailboxes." -ForegroundColor Green

        $results = @()
        foreach ($mailbox in $sharedMailboxes) {
            Write-Host "Processing mailbox: $($mailbox.DisplayName)" -ForegroundColor Yellow
            
            $fullAccess = Get-MailboxPermission -Identity $mailbox.Identity | 
                Where-Object {$_.User -notlike "NT AUTHORITY\*" -and $_.User -notlike "S-1-5*" -and $_.IsInherited -eq $false}
            
            $sendAs = Get-RecipientPermission -Identity $mailbox.Identity | 
                Where-Object {$_.Trustee -notlike "NT AUTHORITY\*" -and $_.Trustee -notlike "S-1-5*"}
            
            $sendOnBehalf = $mailbox.GrantSendOnBehalfTo
            
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

        $results | Format-Table -AutoSize

        if ($OutputFile) {
            $results | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
            Write-Host "Permissions exported to '$OutputFile'." -ForegroundColor Green
        }

        return $results
    } catch {
        Write-Host "Error retrieving shared mailbox permissions: $_" -ForegroundColor Red
    }
}

# Main script
Clear-Host
Write-Host "=== Exchange Online Shared Mailbox Access Report ===" -ForegroundColor Cyan
Write-Host "`nConnected to: $($script:ExchangeConnection.OrganizationName)" -ForegroundColor Green
Write-Host "User: $($script:ExchangeConnection.CurrentUser)`n" -ForegroundColor Green

$exportChoice = Read-Host "Do you want to export the results to a CSV file? (Y/N)"
if ($exportChoice -eq 'Y' -or $exportChoice -eq 'y') {
    $defaultPath = "C:\SharedMailboxPermissions.csv"
    Write-Host "Default export path is: $defaultPath"
    $customPath = Read-Host "Press Enter to use default path or type a custom path"
    $outputFile = if ($customPath) { $customPath } else { $defaultPath }
} else {
    $outputFile = $null
}

Write-Host "`nRetrieving shared mailbox permissions..." -ForegroundColor Cyan
Get-SharedMailboxPermissions -OutputFile $outputFile