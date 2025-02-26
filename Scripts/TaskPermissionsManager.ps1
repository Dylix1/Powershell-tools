function Get-TaskFolderName {
    param (
        [Parameter(Mandatory)]
        [string]$Mailbox
    )
    try {
        $folders = Get-MailboxFolderStatistics -Identity $Mailbox | 
            Where-Object { $_.FolderType -eq "Tasks" -or $_.ContainerClass -eq "IPF.Task" } |
            Select-Object Name, FolderPath
        if ($folders.Count -gt 0) {
            return $folders[0].Name
        } else {
            Write-Host "No Task folder found for $Mailbox." -ForegroundColor Yellow
            return "Tasks" # Default task folder name
        }
    } catch {
        Write-Host "Error getting task folder name: $_" -ForegroundColor Red
        return "Tasks" # Default task folder name
    }
}

function Show-AllTaskFolders {
    param (
        [Parameter(Mandatory)]
        [string]$Mailbox
    )
    try {
        $taskFolders = Get-MailboxFolderStatistics -Identity $Mailbox | 
            Where-Object { 
                $_.FolderType -eq "Tasks" -or 
                $_.ContainerClass -eq "IPF.Task" -or 
                $_.FolderPath -like "*/Tasks" 
            } |
            Select-Object Name, FolderPath, ItemsInFolder, FolderSize
        
        if ($taskFolders.Count -eq 0) {
            Write-Host "No task folders found for $Mailbox." -ForegroundColor Yellow
            return
        }
        
        Write-Host "`nAll Task Folders for $Mailbox" -ForegroundColor Cyan
        Write-Host "----------------------------------------" -ForegroundColor Cyan
        
        foreach ($folder in $taskFolders) {
            Write-Host "`nTask Folder Name: $($folder.Name)" -ForegroundColor Yellow
            Write-Host "Folder Path: $($folder.FolderPath)"
            Write-Host "Items in Folder: $($folder.ItemsInFolder)"
            Write-Host "Folder Size: $($folder.FolderSize)"
            
            # Display permissions for each task folder
            try {
                # Convert the folder path to the correct format
                $folderPath = $folder.FolderPath -replace '^/', ''  # Remove leading slash
                $folderPath = $folderPath -replace '/', '\'  # Replace forward slashes with backslashes
                $permissions = Get-MailboxFolderPermission -Identity "$($Mailbox):\$folderPath" -ErrorAction Stop
                Write-Host "`nPermissions:" -ForegroundColor Green
                $permissions | Format-Table User, AccessRights -AutoSize
            } catch {
                Write-Host "Unable to retrieve permissions for this task folder: $_" -ForegroundColor Red
            }
            Write-Host "----------------------------------------" -ForegroundColor Cyan
        }
    } catch {
        Write-Host "Error retrieving task folders: $_" -ForegroundColor Red
    }
}

function Show-TaskFolderPermissions {
    param (
        [Parameter(Mandatory)]
        [string]$Mailbox,
        [Parameter(Mandatory)]
        [string]$TaskFolderName
    )
    try {
        $permissions = Get-MailboxFolderPermission -Identity "$($Mailbox):\$TaskFolderName"
        Write-Host "`nCurrent Task Folder Permissions:" -ForegroundColor Cyan
        $permissions | Format-Table User, AccessRights -AutoSize
        
        # Count users with access
        $usersWithAccess = $permissions | Where-Object { 
            $_.User.ToString() -ne "Default" -and 
            $_.User.ToString() -ne "Anonymous" -and 
            $_.AccessRights -ne "None" 
        }
        
        Write-Host "`nUsers with Access:" -ForegroundColor Green
        Write-Host "----------------------------------------" -ForegroundColor Green
        Write-Host "Total users with access: $($usersWithAccess.Count)" -ForegroundColor Yellow
        
        if ($usersWithAccess.Count -gt 0) {
            foreach ($user in $usersWithAccess) {
                Write-Host "$($user.User) - $($user.AccessRights)" -ForegroundColor White
            }
        } else {
            Write-Host "No users have been granted specific access to this task folder." -ForegroundColor Yellow
        }
        Write-Host "----------------------------------------" -ForegroundColor Green
    } catch {
        Write-Host "Error getting task folder permissions: $_" -ForegroundColor Red
    }
}

function Set-TaskFolderPermission {
    param (
        [Parameter(Mandatory)]
        [string]$Mailbox,
        [Parameter(Mandatory)]
        [string]$TaskFolderName,
        [Parameter(Mandatory)]
        [string]$User,
        [Parameter(Mandatory)]
        [ValidateSet("Owner", "PublishingEditor", "Editor", "PublishingAuthor", "Author", "NonEditingAuthor", "Reviewer", "Contributor", "None")]
        [string]$AccessRight
    )
    try {
        $existingPermission = Get-MailboxFolderPermission -Identity "$($Mailbox):\$TaskFolderName" -User $User -ErrorAction SilentlyContinue
        
        if ($existingPermission) {
            Set-MailboxFolderPermission -Identity "$($Mailbox):\$TaskFolderName" -User $User -AccessRights $AccessRight -ErrorAction Stop
            Write-Host "Successfully updated permissions for $User" -ForegroundColor Green
        } else {
            Add-MailboxFolderPermission -Identity "$($Mailbox):\$TaskFolderName" -User $User -AccessRights $AccessRight -ErrorAction Stop
            Write-Host "Successfully added permissions for $User" -ForegroundColor Green
        }
        return $true
    } catch {
        Write-Host "Error setting task folder permission: $_" -ForegroundColor Red
        return $false
    }
}

function Remove-TaskFolderPermission {
    param (
        [Parameter(Mandatory)]
        [string]$Mailbox,
        [Parameter(Mandatory)]
        [string]$TaskFolderName,
        [Parameter(Mandatory)]
        [string]$User
    )
    try {
        $existingPermission = Get-MailboxFolderPermission -Identity "$($Mailbox):\$TaskFolderName" -User $User -ErrorAction SilentlyContinue
        
        if ($existingPermission) {
            Remove-MailboxFolderPermission -Identity "$($Mailbox):\$TaskFolderName" -User $User -Confirm:$false -ErrorAction Stop
            Write-Host "Successfully removed permissions for $User" -ForegroundColor Green
            return $true
        } else {
            Write-Host "No permissions found for $User on this task folder" -ForegroundColor Yellow
            return $false
        }
    } catch {
        Write-Host "Error removing task folder permission: $_" -ForegroundColor Red
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
- NonEditingAuthor: Can create and read items, but cannot edit
- Reviewer: Read-only access
- Contributor: Can only create new items
- None: No access
"@
    Write-Host $help -ForegroundColor Yellow
}

# Main script
Clear-Host
Write-Host "=== Exchange Online Task Permissions Manager ===" -ForegroundColor Cyan
Write-Host "`nConnected to: $($script:ExchangeConnection.OrganizationName)" -ForegroundColor Green
Write-Host "User: $($script:ExchangeConnection.CurrentUser)`n" -ForegroundColor Green

do {
    Write-Host "`nOptions:" -ForegroundColor Cyan
    Write-Host "1. Manage Task Folder Permissions"
    Write-Host "2. View All Task Folders"
    Write-Host "3. Show Users with Task Access"
    Write-Host "4. Show Access Rights Help"
    Write-Host "5. Return to Main Menu"

    $choice = Read-Host "`nEnter your choice (1-5)"
    
    switch ($choice) {
        "1" {
            $Mailbox = Read-Host "`nEnter mailbox to manage"
            $TaskFolderName = Get-TaskFolderName -Mailbox $Mailbox
            if (!$TaskFolderName) { continue }

            do {
                Write-Host "`nOptions for $($Mailbox)'s task folder:" -ForegroundColor Cyan
                Write-Host "1. View current permissions"
                Write-Host "2. Add/Update permission"
                Write-Host "3. Remove permission"
                Write-Host "4. Back to main options"
                
                $subChoice = Read-Host "`nEnter your choice (1-4)"
                
                switch ($subChoice) {
                    "1" { Show-TaskFolderPermissions -Mailbox $Mailbox -TaskFolderName $TaskFolderName }
                    "2" {
                        $User = Read-Host "`nEnter user email to grant permission"
                        Show-AccessRightsHelp
                        $AccessRight = Read-Host "`nEnter access right"
                        if ([string]::IsNullOrWhiteSpace($AccessRight)) { continue }
                        Set-TaskFolderPermission -Mailbox $Mailbox -TaskFolderName $TaskFolderName -User $User -AccessRight $AccessRight
                    }
                    "3" {
                        $User = Read-Host "`nEnter user email to remove permission"
                        Remove-TaskFolderPermission -Mailbox $Mailbox -TaskFolderName $TaskFolderName -User $User
                    }
                    "4" { break }
                }
            } while ($subChoice -ne "4")
        }
        "2" {
            $Mailbox = Read-Host "`nEnter mailbox to view all task folders"
            Show-AllTaskFolders -Mailbox $Mailbox
        }
        "3" {
            $Mailbox = Read-Host "`nEnter mailbox to view users with task access"
            $TaskFolderName = Get-TaskFolderName -Mailbox $Mailbox
            if (!$TaskFolderName) { continue }
            Show-TaskFolderPermissions -Mailbox $Mailbox -TaskFolderName $TaskFolderName
        }
        "4" { Show-AccessRightsHelp }
        "5" { return }
    }
} while ($true)