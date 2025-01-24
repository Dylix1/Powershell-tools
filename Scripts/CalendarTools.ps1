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
        Write-Host "Error getting calendar folder name: $_" -ForegroundColor Red
        return $null
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
        Write-Host "Error getting calendar permissions: $_" -ForegroundColor Red
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
        Write-Host "Error setting calendar permission: $_" -ForegroundColor Red
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
    Write-Host "2. Show Access Rights Help"
    Write-Host "3. Return to Main Menu"

    $choice = Read-Host "`nEnter your choice (1-3)"
    
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
        "2" { Show-AccessRightsHelp }
        "3" { return }
    }
} while ($true)