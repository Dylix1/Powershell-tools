function Connect-ExchangeOnlineSession {
    param (
        [Parameter(Mandatory)]
        [string]$AdminUser,
        [switch]$UseBasicAuth
    )
    try {
        $params = @{
            UserPrincipalName = $AdminUser
            ShowProgress = $false
            ShowBanner = $false
        }
        if ($UseBasicAuth) {
            $params.Add('BasicAuthentication', $true)
        }
        Connect-ExchangeOnline @params
        Write-Host "Successfully connected to Exchange Online as $AdminUser." -ForegroundColor Green
        return $true
    } catch {
        Write-Host "Error connecting to Exchange Online: $_" -ForegroundColor Red
        return $false
    }
}

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
try {
    $ErrorActionPreference = 'Stop'
    Write-Host "=== Exchange Online Calendar Permissions Manager ===" -ForegroundColor Cyan

    $AdminUPN = Read-Host "Enter Exchange Online admin email"
    
    $maxRetries = 3
    $connected = $false
    for ($i = 1; $i -le $maxRetries -and !$connected; $i++) {
        $connected = Connect-ExchangeOnlineSession -AdminUser $AdminUPN
        if (!$connected -and $i -lt $maxRetries) {
            Write-Host "Retrying connection ($i/$maxRetries)..." -ForegroundColor Yellow
            Start-Sleep -Seconds 3
        }
    }
    if (!$connected) { exit 1 }

    do {
        $Mailbox = Read-Host "`nEnter mailbox to manage (or 'exit' to quit)"
        if ($Mailbox -eq 'exit') { break }

        $CalendarName = Get-CalendarFolderName -Mailbox $Mailbox
        if (!$CalendarName) { continue }

        do {
            Write-Host "`nOptions for $($Mailbox)'s calendar:" -ForegroundColor Cyan
            Write-Host "1. View current permissions"
            Write-Host "2. Add/Update permission"
            Write-Host "3. Show access rights help"
            Write-Host "4. Back to mailbox selection"
            
            $choice = Read-Host "`nEnter your choice (1-4)"
            
            switch ($choice) {
                "1" {
                    Show-CalendarPermissions -Mailbox $Mailbox -CalendarName $CalendarName
                }
                "2" {
                    $User = Read-Host "Enter user email to grant permission"
                    Show-AccessRightsHelp
                    $AccessRight = Read-Host "Enter access right"
                    if ([string]::IsNullOrWhiteSpace($AccessRight)) { continue }
                    Set-CalendarPermission -Mailbox $Mailbox -CalendarName $CalendarName -User $User -AccessRight $AccessRight
                }
                "3" {
                    Show-AccessRightsHelp
                }
                "4" {
                    break
                }
            }
        } while ($choice -ne "4")
        
    } while ($true)

} catch {
    Write-Host "Critical error: $_" -ForegroundColor Red
} finally {
    if ($connected) {
        Disconnect-ExchangeOnline -Confirm:$false
        Write-Host "Successfully disconnected from Exchange Online." -ForegroundColor Green
    }
}