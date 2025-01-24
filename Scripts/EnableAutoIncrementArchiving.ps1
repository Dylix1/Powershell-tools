function Get-MailboxArchiveStatus {
    param (
        [Parameter(Mandatory)]
        [string]$Mailbox
    )
    try {
        $mailboxInfo = Get-Mailbox -Identity $Mailbox -ErrorAction Stop | 
            Select-Object DisplayName, UserPrincipalName, ArchiveStatus, ArchiveDatabase,
                        AutoExpandingArchiveEnabled, ArchiveQuota, ArchiveWarningQuota

        Write-Host "`nMailbox Archive Information:" -ForegroundColor Cyan
        Write-Host "Display Name: $($mailboxInfo.DisplayName)"
        Write-Host "Email: $($mailboxInfo.UserPrincipalName)"
        Write-Host "Archive Status: $($mailboxInfo.ArchiveStatus)"
        Write-Host "Archive Database: $($mailboxInfo.ArchiveDatabase)"
        Write-Host "Auto-Expanding Archive: $($mailboxInfo.AutoExpandingArchiveEnabled)"
        Write-Host "Archive Quota: $($mailboxInfo.ArchiveQuota)"
        Write-Host "Archive Warning Quota: $($mailboxInfo.ArchiveWarningQuota)"

        if ($mailboxInfo.ArchiveStatus -eq "Active" -or $mailboxInfo.ArchiveDatabase) {
            $archiveStats = Get-MailboxStatistics -Identity $Mailbox -Archive -ErrorAction Stop |
                Select-Object DisplayName, TotalItemSize, ItemCount

            Write-Host "`nArchive Statistics:" -ForegroundColor Cyan
            Write-Host "Total Items: $($archiveStats.ItemCount)"
            Write-Host "Total Size: $($archiveStats.TotalItemSize)"
        }

        return $mailboxInfo
    } catch {
        Write-Host "Error checking archive status: $_" -ForegroundColor Red
        return $null
    }
}

function Start-RetentionPolicyProcessing {
    param (
        [Parameter(Mandatory)]
        [string]$Mailbox
    )
    try {
        $mailboxInfo = Get-Mailbox -Identity $Mailbox -ErrorAction Stop
        Write-Host "Mailbox found. Starting retention policy processing..." -ForegroundColor Yellow
        
        Start-ManagedFolderAssistant -Identity $Mailbox -ErrorAction Stop
        $stats = Get-MailboxFolderStatistics -Identity $Mailbox -FolderScope RecoverableItems
        
        Write-Host "`nRetention Policy Processing initiated successfully." -ForegroundColor Green
        Write-Host "Current recoverable items status:" -ForegroundColor Cyan
        $stats | Format-Table Name, ItemsInFolder, FolderAndSubfolderSize
        
        return $true
    } catch {
        Write-Host "Error processing retention policy: $_" -ForegroundColor Red
        return $false
    }
}

function Enable-CustomAutoExpandingArchive {
    param (
        [Parameter(Mandatory)]
        [string]$Mailbox
    )
    try {
        $mailboxInfo = Get-Mailbox -Identity $Mailbox -ErrorAction Stop
        
        if ($mailboxInfo.ArchiveStatus -ne "Active" -and -not $mailboxInfo.ArchiveDatabase) {
            Write-Host "Archive is not active for this mailbox. Please enable archive first." -ForegroundColor Yellow
            return $false
        }

        Enable-Mailbox -Identity $Mailbox -AutoExpandingArchive -ErrorAction Stop
        Write-Host "Successfully enabled auto-expanding archive for $Mailbox" -ForegroundColor Green
        return $true
    } catch {
        Write-Host "Error enabling auto-expanding archive: $_" -ForegroundColor Red
        return $false
    }
}

# Main script
Clear-Host
Write-Host "=== Exchange Online Archive Manager ===" -ForegroundColor Cyan
Write-Host "`nConnected to: $($script:ExchangeConnection.OrganizationName)" -ForegroundColor Green
Write-Host "User: $($script:ExchangeConnection.CurrentUser)`n" -ForegroundColor Green

do {
    Write-Host "`nOptions:" -ForegroundColor Cyan
    Write-Host "1. Check Archive Status"
    Write-Host "2. Enable Auto-Expanding Archive"
    Write-Host "3. Force Retention Processing"
    Write-Host "4. Return to Main Menu"
    
    $choice = Read-Host "`nEnter your choice (1-4)"
    if ($choice -eq '4') { return }
    
    if ($choice -in '1','2','3') {
        $Mailbox = Read-Host "`nEnter mailbox to manage"
        switch ($choice) {
            '1' { Get-MailboxArchiveStatus -Mailbox $Mailbox }
            '2' { Enable-CustomAutoExpandingArchive -Mailbox $Mailbox }
            '3' { Start-RetentionPolicyProcessing -Mailbox $Mailbox }
        }
    }
} while ($true)