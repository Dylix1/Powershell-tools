function Get-CalendarFolderName {
    param (
        [Parameter(Mandatory)]
        [string]$Mailbox
    )
    try {
        $folders = Get-MailboxFolderStatistics -Identity $Mailbox | 
            Where-Object { $_.FolderType -eq "Calendar" } |
            Select-Object Name, FolderPath
        return $folders[0].Name
    } catch {
        $errorMsg = $Error[0].Exception.Message
        Write-Host "Error getting calendar folder name" -ForegroundColor Red
        Write-Host "Reason: $errorMsg" -ForegroundColor Red
        return $null
    }
}

function Show-AllCalendars {
    param (
        [Parameter(Mandatory)]
        [string]$Mailbox
    )
    try {
        $calendars = Get-MailboxFolderStatistics -Identity $Mailbox | 
            Where-Object { 
                $_.FolderType -eq "Calendar" -or 
                $_.ContainerClass -eq "IPF.Calendar" -or 
                $_.FolderPath -like "*/Calendar" -or
                $_.FolderPath -like "*/Agenda"
            } |
            Select-Object Name, FolderPath, ItemsInFolder, FolderSize
        
        Write-Host "`nAll Calendars for $Mailbox" -ForegroundColor Cyan
        Write-Host "----------------------------------------" -ForegroundColor Cyan
        
        foreach ($calendar in $calendars) {
            Write-Host "`nCalendar Name: $($calendar.Name)" -ForegroundColor Yellow
            Write-Host "Folder Path: $($calendar.FolderPath)"
            Write-Host "Items in Calendar: $($calendar.ItemsInFolder)"
            Write-Host "Folder Size: $($calendar.FolderSize)"
            
            # Display permissions for each calendar
            try {
                # Convert the folder path to the correct format
                $folderPath = $calendar.FolderPath -replace '^/', ''  # Remove leading slash
                $folderPath = $folderPath -replace '/', '\'  # Replace forward slashes with backslashes
                $permissions = Get-MailboxFolderPermission -Identity "$($Mailbox):\$folderPath" -ErrorAction Stop
                Write-Host "`nPermissions:" -ForegroundColor Green
                $permissions | Format-Table User, AccessRights -AutoSize
            } catch {
                $errorMsg = $Error[0].Exception.Message
                Write-Host "Unable to retrieve permissions for this calendar" -ForegroundColor Red
                Write-Host "Reason: $errorMsg" -ForegroundColor Red
            }
            Write-Host "----------------------------------------" -ForegroundColor Cyan
        }
    } catch {
        $errorMsg = $Error[0].Exception.Message
        Write-Host "Error retrieving calendars" -ForegroundColor Red
        Write-Host "Reason: $errorMsg" -ForegroundColor Red
    }
}

function Show-CalendarPermissions {
    param (
        [Parameter(Mandatory)]
        [string]$Mailbox,
        [Parameter(Mandatory)]
        [string]$CalendarName
    )
    try {
        $permissions = Get-MailboxFolderPermission -Identity "$($Mailbox):\$CalendarName"
        Write-Host "`nCurrent Calendar Permissions:" -ForegroundColor Cyan
        $permissions | Format-Table User, AccessRights -AutoSize
    } catch {
        $errorMsg = $Error[0].Exception.Message
        Write-Host "Error getting calendar permissions" -ForegroundColor Red
        Write-Host "Reason: $errorMsg" -ForegroundColor Red
    }
}

function Set-CalendarPermission {
    param (
        [Parameter(Mandatory)]
        [string]$Mailbox,
        [Parameter(Mandatory)]
        [string]$CalendarName,
        [Parameter(Mandatory)]
        [string]$User,
        [Parameter(Mandatory)]
        [ValidateSet("Owner", "PublishingEditor", "Editor", "PublishingAuthor", "Author", "NonEditingAuthor", "Reviewer", "Contributor", "None")]
        [string]$AccessRight
    )
    try {
        $existingPermission = Get-MailboxFolderPermission -Identity "$($Mailbox):\$CalendarName" -User $User -ErrorAction SilentlyContinue
        
        if ($existingPermission) {
            Set-MailboxFolderPermission -Identity "$($Mailbox):\$CalendarName" -User $User -AccessRights $AccessRight -ErrorAction Stop
            Write-Host "Successfully updated permissions for $User" -ForegroundColor Green
        } else {
            Add-MailboxFolderPermission -Identity "$($Mailbox):\$CalendarName" -User $User -AccessRights $AccessRight -ErrorAction Stop
            Write-Host "Successfully added permissions for $User" -ForegroundColor Green
        }
        return $true
    } catch {
        $errorMsg = $Error[0].Exception.Message
        Write-Host "Error setting calendar permission" -ForegroundColor Red
        Write-Host "Reason: $errorMsg" -ForegroundColor Red
        return $false
    }
}

function Set-CalendarPermissionForDomain {
    param (
        [Parameter(Mandatory)]
        [string]$Mailbox,
        [Parameter(Mandatory)]
        [string]$CalendarName,
        [Parameter(Mandatory)]
        [string]$Domain,
        [Parameter(Mandatory)]
        [ValidateSet("Owner", "PublishingEditor", "Editor", "PublishingAuthor", "Author", "NonEditingAuthor", "Reviewer", "Contributor", "None")]
        [string]$AccessRight
    )
    try {
        # Get all users with the specified domain
        Write-Host "Searching for users with domain: $Domain" -ForegroundColor Cyan
        $users = Get-Mailbox -ResultSize Unlimited | Where-Object {$_.PrimarySmtpAddress -like "*@$Domain"}
        
        if ($users.Count -eq 0) {
            Write-Host "No users found with the domain: $Domain" -ForegroundColor Yellow
            return $false
        }
        
        Write-Host "Found $($users.Count) users with domain $Domain" -ForegroundColor Cyan
        $confirmation = Read-Host "Are you sure you want to grant '$AccessRight' permissions to all $($users.Count) users? (Y/N)"
        
        if ($confirmation -ne "Y" -and $confirmation -ne "y") {
            Write-Host "Operation cancelled." -ForegroundColor Yellow
            return $false
        }
        
        $successCount = 0
        $failCount = 0
        
        foreach ($user in $users) {
            $userEmail = $user.PrimarySmtpAddress
            Write-Host "Setting permission for $userEmail..." -ForegroundColor Gray
            
            try {
                $existingPermission = Get-MailboxFolderPermission -Identity "$($Mailbox):\$CalendarName" -User $userEmail -ErrorAction SilentlyContinue
                
                if ($existingPermission) {
                    Set-MailboxFolderPermission -Identity "$($Mailbox):\$CalendarName" -User $userEmail -AccessRights $AccessRight -ErrorAction Stop
                } else {
                    Add-MailboxFolderPermission -Identity "$($Mailbox):\$CalendarName" -User $userEmail -AccessRights $AccessRight -ErrorAction Stop
                }
                $successCount++
            } catch {
                $errorMsg = $Error[0].Exception.Message
                Write-Host "Error setting permission for $userEmail" -ForegroundColor Red
                Write-Host "Reason: $errorMsg" -ForegroundColor Red
                $failCount++
            }
        }
        
        Write-Host "`nPermission assignment complete:" -ForegroundColor Green
        Write-Host "- Successfully set permissions for $successCount users" -ForegroundColor Green
        if ($failCount -gt 0) {
            Write-Host "- Failed to set permissions for $failCount users" -ForegroundColor Red
        }
        
        return $true
    } catch {
        $errorMsg = $Error[0].Exception.Message
        Write-Host "Error setting domain permissions" -ForegroundColor Red
        Write-Host "Reason: $errorMsg" -ForegroundColor Red
        return $false
    }
}

function Show-AccessRightsHelp {
    $help = @"
Available Access Rights:
- Owner: Full control, including modifying permissions
- PublishingEditor: Can create, read, modify, delete items and create subfolders
- Editor: Can create, read, modify and delete items
- PublishingAuthor: Can create, read items and create subfolders; can delete own items
- Author: Can create, read and delete own items
- NonEditingAuthor: Can create and delete own items, but not read
- Reviewer: Read-only access
- Contributor: Can only create new items
- None: No access
"@
    Write-Host $help -ForegroundColor Yellow
}

# Main script
Clear-Host
Write-Host "=== Calendar Permissions Manager ===" -ForegroundColor Cyan
Write-Host "`nConnected to: $($script:ExchangeConnection.OrganizationName)" -ForegroundColor Green
Write-Host "User: $($script:ExchangeConnection.CurrentUser)`n" -ForegroundColor Green

do {
    Write-Host "`nOptions:" -ForegroundColor Cyan
    Write-Host "1. Manage Calendar Permissions"
    Write-Host "2. View All Calendars"
    Write-Host "3. Show Access Rights Help"
    Write-Host "4. Add Domain Permissions"
    Write-Host "5. Return to Main Menu"

    $choice = Read-Host "`nEnter your choice (1-5)"
    
    switch ($choice) {
        "1" {
            $Mailbox = Read-Host "`nEnter mailbox to manage"
            $CalendarName = Get-CalendarFolderName -Mailbox $Mailbox
            if (!$CalendarName) { continue }

            do {
                Write-Host "`nOptions for $($Mailbox)'s calendar:" -ForegroundColor Cyan
                Write-Host "1. View current permissions"
                Write-Host "2. Add/Update permission"
                Write-Host "3. Back to main options"
                
                $subChoice = Read-Host "`nEnter your choice (1-3)"
                
                switch ($subChoice) {
                    "1" { Show-CalendarPermissions -Mailbox $Mailbox -CalendarName $CalendarName }
                    "2" {
                        $User = Read-Host "`nEnter user email to grant permission"
                        Show-AccessRightsHelp
                        $AccessRight = Read-Host "`nEnter access right"
                        if ([string]::IsNullOrWhiteSpace($AccessRight)) { continue }
                        Set-CalendarPermission -Mailbox $Mailbox -CalendarName $CalendarName -User $User -AccessRight $AccessRight
                    }
                    "3" { break }
                }
            } while ($subChoice -ne "3")
        }
        "2" {
            $Mailbox = Read-Host "`nEnter mailbox to view all calendars"
            Show-AllCalendars -Mailbox $Mailbox
        }
        "3" { Show-AccessRightsHelp }
        "4" {
            $Mailbox = Read-Host "`nEnter mailbox to manage"
            $CalendarName = Get-CalendarFolderName -Mailbox $Mailbox
            if (!$CalendarName) { continue }
            
            $Domain = Read-Host "`nEnter domain (e.g., example.com)"
            if ([string]::IsNullOrWhiteSpace($Domain)) { continue }
            
            Show-AccessRightsHelp
            $AccessRight = Read-Host "`nEnter access right"
            if ([string]::IsNullOrWhiteSpace($AccessRight)) { continue }
            
            Set-CalendarPermissionForDomain -Mailbox $Mailbox -CalendarName $CalendarName -Domain $Domain -AccessRight $AccessRight
        }
        "5" { return }
    }
} while ($true)