function Test-UserExists {
    param (
        [Parameter(Mandatory)]
        [string]$UserIdentity
    )
    try {
        $null = Get-User -Identity $UserIdentity -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

function Get-UserGroups {
    param (
        [Parameter(Mandatory)]
        [string]$UserIdentity
    )
    try {
        # Get distribution group memberships
        $distributionGroups = Get-DistributionGroup -ResultSize Unlimited | 
            Where-Object { (Get-DistributionGroupMember -Identity $_.Identity).PrimarySmtpAddress -contains $UserIdentity }

        # Get mail-enabled security groups if connected to Exchange Online
        $securityGroups = Get-UnifiedGroup -ResultSize Unlimited |
            Where-Object { (Get-UnifiedGroupLinks -Identity $_.Identity -LinkType Members).PrimarySmtpAddress -contains $UserIdentity }

        return @{
            DistributionGroups = $distributionGroups
            SecurityGroups = $securityGroups
        }
    } catch {
        Write-Host "Error retrieving group memberships: $_" -ForegroundColor Red
        return $null
    }
}

function Add-UserToGroups {
    param (
        [Parameter(Mandatory)]
        [string]$UserIdentity,
        [Parameter(Mandatory)]
        [object]$Groups
    )
    try {
        $successCount = 0
        $failCount = 0

        # Process distribution groups
        foreach ($group in $Groups.DistributionGroups) {
            try {
                Add-DistributionGroupMember -Identity $group.Identity -Member $UserIdentity -ErrorAction Stop
                Write-Host "Successfully added to distribution group: $($group.DisplayName)" -ForegroundColor Green
                $successCount++
            } catch {
                if ($_.Exception.Message -match "is already a member") {
                    Write-Host "User is already a member of distribution group: $($group.DisplayName)" -ForegroundColor Yellow
                } else {
                    Write-Host "Failed to add to distribution group $($group.DisplayName): $_" -ForegroundColor Red
                    $failCount++
                }
            }
        }

        # Process security groups
        foreach ($group in $Groups.SecurityGroups) {
            try {
                Add-UnifiedGroupLinks -Identity $group.Identity -LinkType Members -Links $UserIdentity -ErrorAction Stop
                Write-Host "Successfully added to security group: $($group.DisplayName)" -ForegroundColor Green
                $successCount++
            } catch {
                if ($_.Exception.Message -match "is already a member") {
                    Write-Host "User is already a member of security group: $($group.DisplayName)" -ForegroundColor Yellow
                } else {
                    Write-Host "Failed to add to security group $($group.DisplayName): $_" -ForegroundColor Red
                    $failCount++
                }
            }
        }

        Write-Host "`nSummary:" -ForegroundColor Cyan
        Write-Host "Successfully added to $successCount groups" -ForegroundColor Green
        if ($failCount -gt 0) {
            Write-Host "Failed to add to $failCount groups" -ForegroundColor Red
        }
    } catch {
        Write-Host "Error adding user to groups: $_" -ForegroundColor Red
    }
}

# Main script
Clear-Host
Write-Host "=== Copy Group Memberships ===" -ForegroundColor Cyan
Write-Host "`nConnected to: $($script:ExchangeConnection.OrganizationName)" -ForegroundColor Green
Write-Host "User: $($script:ExchangeConnection.CurrentUser)`n" -ForegroundColor Green

do {
    Write-Host "`nOptions:" -ForegroundColor Cyan
    Write-Host "1. Copy Group Memberships"
    Write-Host "2. Return to Main Menu"
    
    $choice = Read-Host "`nEnter your choice (1-2)"
    
    if ($choice -eq '2') { return }
    
    if ($choice -eq '1') {
        $sourceUser = Read-Host "`nEnter the source user's email address"
        if (-not (Test-UserExists -UserIdentity $sourceUser)) {
            Write-Host "Source user not found!" -ForegroundColor Red
            continue
        }

        $targetUser = Read-Host "Enter the target user's email address"
        if (-not (Test-UserExists -UserIdentity $targetUser)) {
            Write-Host "Target user not found!" -ForegroundColor Red
            continue
        }

        Write-Host "`nRetrieving group memberships for $sourceUser..." -ForegroundColor Cyan
        $groups = Get-UserGroups -UserIdentity $sourceUser

        if ($groups) {
            $totalGroups = ($groups.DistributionGroups | Measure-Object).Count + ($groups.SecurityGroups | Measure-Object).Count
            Write-Host "Found $totalGroups groups" -ForegroundColor Green
            
            $confirm = Read-Host "`nDo you want to copy these group memberships to $targetUser? (Y/N)"
            if ($confirm -eq 'Y' -or $confirm -eq 'y') {
                Write-Host "`nCopying group memberships..." -ForegroundColor Cyan
                Add-UserToGroups -UserIdentity $targetUser -Groups $groups
            }
        }
    }
} while ($true)