function Add-MultipleGroupOwners {
    param (
        [Parameter(Mandatory)]
        [string]$GroupIdentity,
        [Parameter(Mandatory)]
        [string[]]$Owners,
        [Parameter(Mandatory=$false)]
        [switch]$DistributionGroup
    )
    try {
        # Verify group exists
        $groupExists = if ($DistributionGroup) {
            $group = Get-DistributionGroup -Identity $GroupIdentity -ErrorAction Stop
            $true
        } else {
            $group = Get-ADGroup -Identity $GroupIdentity -ErrorAction Stop
            $true
        }
        
        if (-not $groupExists) {
            Write-Host "Group '$GroupIdentity' does not exist." -ForegroundColor Red
            return $false
        }
        
        $successCount = 0
        $errorCount = 0
        $alreadyOwnerCount = 0
        
        foreach ($owner in $Owners) {
            try {
                if ($DistributionGroup) {
                    # Check if user exists in Exchange
                    try {
                        $null = Get-Recipient -Identity $owner -ErrorAction Stop
                    } catch {
                        Write-Host "User '$owner' does not exist in Exchange." -ForegroundColor Red
                        $errorCount++
                        continue
                    }
                    
                    # Check if already an owner
                    $groupInfo = Get-DistributionGroup -Identity $GroupIdentity
                    if ($groupInfo.ManagedBy -contains $owner) {
                        Write-Host "User '$owner' is already an owner of the group." -ForegroundColor Yellow
                        $alreadyOwnerCount++
                        continue
                    }
                    
                    # Add as owner
                    Set-DistributionGroup -Identity $GroupIdentity -ManagedBy @{Add=$owner} -ErrorAction Stop
                    Write-Host "Added '$owner' as an owner of the group." -ForegroundColor Green
                    $successCount++
                } else {
                    # Active Directory groups
                    # Check if user exists in AD
                    try {
                        $user = Get-ADUser -Identity $owner -ErrorAction Stop
                    } catch {
                        Write-Host "User '$owner' does not exist in Active Directory." -ForegroundColor Red
                        $errorCount++
                        continue
                    }
                    
                    # For AD groups, we need to use different techniques - first make them a member if not already
                    if (-not (Get-ADGroupMember -Identity $GroupIdentity | Where-Object { $_.SamAccountName -eq $user.SamAccountName })) {
                        Add-ADGroupMember -Identity $GroupIdentity -Members $owner -ErrorAction Stop
                        Write-Host "Added '$owner' as a member of the group." -ForegroundColor Green
                    }
                    
                    # Then set managedBy attribute (owner)
                    $currentOwner = (Get-ADGroup -Identity $GroupIdentity -Properties ManagedBy).ManagedBy
                    if ($currentOwner -eq $user.DistinguishedName) {
                        Write-Host "User '$owner' is already an owner of the group." -ForegroundColor Yellow
                        $alreadyOwnerCount++
                        continue
                    }
                    
                    Set-ADGroup -Identity $GroupIdentity -ManagedBy $user -ErrorAction Stop
                    Write-Host "Set '$owner' as the owner of the group." -ForegroundColor Green
                    $successCount++
                }
            } catch {
                Write-Host "Error adding '$owner' as owner: $_" -ForegroundColor Red
                $errorCount++
            }
        }
        
        Write-Host "`nSummary:" -ForegroundColor Cyan
        Write-Host "Total owners processed: $($Owners.Count)" -ForegroundColor White
        Write-Host "Successfully added: $successCount" -ForegroundColor Green
        Write-Host "Already owners: $alreadyOwnerCount" -ForegroundColor Yellow
        Write-Host "Failed to add: $errorCount" -ForegroundColor Red
        
        return $true
    } catch {
        Write-Host "Error adding group owners: $_" -ForegroundColor Red
        return $false
    }
}

function Add-MultipleGroupMembers {
    param (
        [Parameter(Mandatory)]
        [string]$GroupIdentity,
        [Parameter(Mandatory)]
        [string[]]$Members,
        [Parameter(Mandatory=$false)]
        [switch]$DistributionGroup
    )
    try {
        # Verify group exists
        $groupExists = if ($DistributionGroup) {
            Test-GroupExists -GroupIdentity $GroupIdentity -DistributionGroup
        } else {
            Test-GroupExists -GroupIdentity $GroupIdentity
        }
        
        if (-not $groupExists) {
            Write-Host "Group '$GroupIdentity' does not exist." -ForegroundColor Red
            return $false
        }
        
        $successCount = 0
        $errorCount = 0
        $alreadyMemberCount = 0
        
        foreach ($member in $Members) {
            try {
                if ($DistributionGroup) {
                    # Check if user exists in Exchange
                    try {
                        $null = Get-Recipient -Identity $member -ErrorAction Stop
                    } catch {
                        Write-Host "User '$member' does not exist in Exchange." -ForegroundColor Red
                        $errorCount++
                        continue
                    }
                    
                    # Check if already a member
                    $isAlreadyMember = Get-DistributionGroupMember -Identity $GroupIdentity | 
                        Where-Object { $_.PrimarySmtpAddress -eq $member -or $_.Alias -eq $member -or $_.Name -eq $member }
                    
                    if ($isAlreadyMember) {
                        Write-Host "User '$member' is already a member of the group." -ForegroundColor Yellow
                        $alreadyMemberCount++
                        continue
                    }
                    
                    # Add as member
                    Add-DistributionGroupMember -Identity $GroupIdentity -Member $member -ErrorAction Stop
                    Write-Host "Added '$member' as a member of the group." -ForegroundColor Green
                    $successCount++
                } else {
                    # Active Directory groups
                    # Check if user exists in AD
                    try {
                        $null = Get-ADUser -Identity $member -ErrorAction Stop
                    } catch {
                        Write-Host "User '$member' does not exist in Active Directory." -ForegroundColor Red
                        $errorCount++
                        continue
                    }
                    
                    # Check if already a member
                    $isAlreadyMember = Get-ADGroupMember -Identity $GroupIdentity | 
                        Where-Object { $_.SamAccountName -eq $member -or $_.DistinguishedName -eq $member }
                    
                    if ($isAlreadyMember) {
                        Write-Host "User '$member' is already a member of the group." -ForegroundColor Yellow
                        $alreadyMemberCount++
                        continue
                    }
                    
                    # Add as member
                    Add-ADGroupMember -Identity $GroupIdentity -Members $member -ErrorAction Stop
                    Write-Host "Added '$member' as a member of the group." -ForegroundColor Green
                    $successCount++
                }
            } catch {
                Write-Host "Error adding '$member' as member: $_" -ForegroundColor Red
                $errorCount++
            }
        }
        
        Write-Host "`nSummary:" -ForegroundColor Cyan
        Write-Host "Total members processed: $($Members.Count)" -ForegroundColor White
        Write-Host "Successfully added: $successCount" -ForegroundColor Green
        Write-Host "Already members: $alreadyMemberCount" -ForegroundColor Yellow
        Write-Host "Failed to add: $errorCount" -ForegroundColor Red
        
        return $true
    } catch {
        Write-Host "Error adding group members: $_" -ForegroundColor Red
        return $false
    }
}

function Test-GroupExists {
    param (
        [Parameter(Mandatory)]
        [string]$GroupIdentity,
        [Parameter(Mandatory=$false)]
        [switch]$DistributionGroup
    )
    try {
        if ($DistributionGroup) {
            $null = Get-DistributionGroup -Identity $GroupIdentity -ErrorAction Stop
        } else {
            $null = Get-ADGroup -Identity $GroupIdentity -ErrorAction Stop
        }
        return $true
    } catch {
        return $false
    }
}

function Import-UsersFromFile {
    param (
        [Parameter(Mandatory)]
        [string]$FilePath
    )
    try {
        if (-not (Test-Path -Path $FilePath)) {
            Write-Host "File not found: $FilePath" -ForegroundColor Red
            return $null
        }
        
        $extension = [System.IO.Path]::GetExtension($FilePath).ToLower()
        
        if ($extension -eq ".csv") {
            $users = Import-Csv -Path $FilePath
            
            # Try to determine which property contains user identifiers
            $potentialColumns = @("UserPrincipalName", "EmailAddress", "Email", "User", "Identity", "SamAccountName", "Name")
            $userColumn = $null
            
            foreach ($column in $potentialColumns) {
                if ($column -in $users[0].PSObject.Properties.Name) {
                    $userColumn = $column
                    break
                }
            }
            
            if ($userColumn) {
                return $users | ForEach-Object { $_.$userColumn } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            } else {
                Write-Host "Could not determine user column in CSV file. Please ensure the file has one of these columns: $($potentialColumns -join ', ')" -ForegroundColor Red
                return $null
            }
        } else {
            # Assume it's a text file with one user per line
            $users = Get-Content -Path $FilePath | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            return $users
        }
    } catch {
        Write-Host "Error importing users from file: $_" -ForegroundColor Red
        return $null
    }
}

# Main script
Clear-Host
Write-Host "=== Add Multiple Users to Groups Tool ===" -ForegroundColor Cyan

# Determine environment and check connection
$isConnectedToExchange = $false
$isConnectedToAD = $false

# Check Exchange connection
try {
    # First check if the global variable exists and is set
    if ($script:ExchangeConnection -and $script:ExchangeConnection.IsConnected) {
        $isConnectedToExchange = $true
        Write-Host "`nConnected to Exchange Online:" -ForegroundColor Green
        Write-Host "Organization: $($script:ExchangeConnection.OrganizationName)" -ForegroundColor Green
        Write-Host "User: $($script:ExchangeConnection.CurrentUser)" -ForegroundColor Green
    }
    # As a fallback, try to detect connection by testing an Exchange cmdlet
    elseif (Get-Command "Get-OrganizationConfig" -ErrorAction SilentlyContinue) {
        $org = Get-OrganizationConfig -ErrorAction Stop
        $isConnectedToExchange = $true
        Write-Host "`nConnected to Exchange Online." -ForegroundColor Green
        # If we can determine the connection details
        $session = Get-PSSession | Where-Object {
            ($_.ConfigurationName -eq "Microsoft.Exchange" -or $_.Name -like "ExchangeOnline*") -and 
            $_.State -eq "Opened"
        }
        if ($session) {
            Write-Host "User: $($session.Runspace.ConnectionInfo.Credential.UserName)" -ForegroundColor Green
        }
    }
} catch {
    # Not connected to Exchange
    Write-Host "Unable to detect Exchange connection: $_" -ForegroundColor Yellow -Verbose
}

# Check AD connection (simple test)
try {
    $null = Get-ADDomain -ErrorAction Stop
    $isConnectedToAD = $true
    Write-Host "`nConnected to Active Directory." -ForegroundColor Green
} catch {
    # Not connected to AD
}

# If neither is connected, show warning
if (-not $isConnectedToExchange -and -not $isConnectedToAD) {
    Write-Host "`nWarning: Not connected to Exchange Online or Active Directory." -ForegroundColor Yellow
    Write-Host "Limited functionality will be available." -ForegroundColor Yellow
}

do {
    Write-Host "`nSelect Group Type:" -ForegroundColor Cyan
    Write-Host "1. Active Directory Security/Distribution Group"
    Write-Host "2. Exchange Distribution Group"
    Write-Host "3. Return to Main Menu"
    
    $groupTypeChoice = Read-Host "`nEnter your choice (1-3)"
    
    if ($groupTypeChoice -eq '3') { return }
    
    if ($groupTypeChoice -in '1','2') {
        $isExchangeGroup = ($groupTypeChoice -eq '2')
        
        # Verify connection for selected group type
        if ($isExchangeGroup -and -not $isConnectedToExchange) {
            Write-Host "`nNot connected to Exchange Online. Please connect first." -ForegroundColor Red
            continue
        }
        
        if (-not $isExchangeGroup -and -not $isConnectedToAD) {
            Write-Host "`nNot connected to Active Directory. Please connect first." -ForegroundColor Red
            continue
        }
        
        # Get group name
        $groupName = Read-Host "`nEnter group name"
        
        # Verify group exists
        $groupExists = if ($isExchangeGroup) {
            Test-GroupExists -GroupIdentity $groupName -DistributionGroup
        } else {
            Test-GroupExists -GroupIdentity $groupName
        }
        
        if (-not $groupExists) {
            Write-Host "Group '$groupName' does not exist." -ForegroundColor Red
            continue
        }
        
        # Select operation
        Write-Host "`nSelect Operation:" -ForegroundColor Cyan
        Write-Host "1. Add multiple users as members"
        Write-Host "2. Add multiple users as owners"
        Write-Host "3. Back to group type selection"
        
        $operationChoice = Read-Host "`nEnter your choice (1-3)"
        
        if ($operationChoice -eq '3') { continue }
        
        if ($operationChoice -in '1','2') {
            $isOwnerOperation = ($operationChoice -eq '2')
            
            # Select input method
            Write-Host "`nSelect Input Method:" -ForegroundColor Cyan
            Write-Host "1. Enter users manually (comma-separated)"
            Write-Host "2. Import from file (CSV or TXT)"
            Write-Host "3. Back to operation selection"
            
            $inputMethodChoice = Read-Host "`nEnter your choice (1-3)"
            
            if ($inputMethodChoice -eq '3') { continue }
            
            $users = @()
            
            if ($inputMethodChoice -eq '1') {
                # Manual entry
                $userInput = Read-Host "`nEnter comma-separated list of users (email addresses or SAM account names)"
                $users = $userInput -split ',' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            } elseif ($inputMethodChoice -eq '2') {
                # File import
                $filePath = Read-Host "`nEnter path to CSV or TXT file containing users"
                $users = Import-UsersFromFile -FilePath $filePath
                
                if (-not $users) {
                    Write-Host "Failed to import users from file." -ForegroundColor Red
                    continue
                }
            } else {
                continue
            }
            
            # Confirm operation
            $roleText = if ($isOwnerOperation) { "owners" } else { "members" }
            Write-Host "`nReady to add $($users.Count) users as $roleText to group '$groupName'." -ForegroundColor Yellow
            $confirm = Read-Host "Continue? (Y/N)"
            
            if ($confirm -ne 'Y' -and $confirm -ne 'y') { continue }
            
            # Perform operation
            if ($isOwnerOperation) {
                if ($isExchangeGroup) {
                    Add-MultipleGroupOwners -GroupIdentity $groupName -Owners $users -DistributionGroup
                } else {
                    Add-MultipleGroupOwners -GroupIdentity $groupName -Owners $users
                }
            } else {
                if ($isExchangeGroup) {
                    Add-MultipleGroupMembers -GroupIdentity $groupName -Members $users -DistributionGroup
                } else {
                    Add-MultipleGroupMembers -GroupIdentity $groupName -Members $users
                }
            }
            
            Write-Host "`nOperation completed." -ForegroundColor Green
        }
    }
} while ($true)