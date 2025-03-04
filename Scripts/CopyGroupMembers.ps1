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

function New-ADSecurityGroup {
    param (
        [Parameter(Mandatory)]
        [string]$Name,
        [Parameter(Mandatory=$false)]
        [string]$Description,
        [Parameter(Mandatory=$false)]
        [string]$Path,
        [Parameter(Mandatory=$false)]
        [ValidateSet("Global", "DomainLocal", "Universal")]
        [string]$Scope = "Global",
        [Parameter(Mandatory=$false)]
        [ValidateSet("Security", "Distribution")]
        [string]$Type = "Security"
    )
    try {
        # Check if group already exists
        if (Test-GroupExists -GroupIdentity $Name) {
            Write-Host "Group '$Name' already exists in Active Directory." -ForegroundColor Yellow
            return $null
        }
        
        # Create the new group
        $params = @{
            Name = $Name
            GroupScope = $Scope
            GroupCategory = $Type
        }
        
        if ($Description) {
            $params.Add("Description", $Description)
        }
        
        if ($Path) {
            $params.Add("Path", $Path)
        }
        
        $newGroup = New-ADGroup @params -PassThru
        Write-Host "Successfully created group '$Name'." -ForegroundColor Green
        return $newGroup
    } catch {
        Write-Host "Error creating AD group: $_" -ForegroundColor Red
        return $null
    }
}

function New-ExchangeDistributionGroup {
    param (
        [Parameter(Mandatory)]
        [string]$Name,
        [Parameter(Mandatory=$false)]
        [string]$Alias,
        [Parameter(Mandatory=$false)]
        [string]$Description,
        [Parameter(Mandatory=$false)]
        [switch]$SecurityEnabled
    )
    try {
        # Check if group already exists
        if (Test-GroupExists -GroupIdentity $Name -DistributionGroup) {
            Write-Host "Group '$Name' already exists in Exchange." -ForegroundColor Yellow
            return $null
        }
        
        # Create the new group
        $params = @{
            Name = $Name
            Type = if ($SecurityEnabled) { "Security" } else { "Distribution" }
        }
        
        if ($Alias) {
            $params.Add("Alias", $Alias)
        } else {
            $params.Add("Alias", $Name.Replace(" ", ""))
        }
        
        if ($Description) {
            $params.Add("Description", $Description)
        }
        
        $newGroup = New-DistributionGroup @params
        Write-Host "Successfully created distribution group '$Name'." -ForegroundColor Green
        return $newGroup
    } catch {
        Write-Host "Error creating distribution group: $_" -ForegroundColor Red
        return $null
    }
}

function Copy-ADGroupMembers {
    param (
        [Parameter(Mandatory)]
        [string]$SourceGroup,
        [Parameter(Mandatory)]
        [string]$TargetGroup
    )
    try {
        # Verify both groups exist
        $sourceExists = Test-GroupExists -GroupIdentity $SourceGroup
        if (-not $sourceExists) {
            Write-Host "Source group '$SourceGroup' does not exist in Active Directory." -ForegroundColor Red
            return $false
        }
        
        $targetExists = Test-GroupExists -GroupIdentity $TargetGroup
        if (-not $targetExists) {
            Write-Host "Target group '$TargetGroup' does not exist in Active Directory." -ForegroundColor Red
            return $false
        }
        
        # Get members from source group
        Write-Host "Getting members from source group '$SourceGroup'..." -ForegroundColor Yellow
        $members = Get-ADGroupMember -Identity $SourceGroup -ErrorAction Stop
        
        if ($members.Count -eq 0) {
            Write-Host "Source group has no members." -ForegroundColor Yellow
            return $true
        }
        
        Write-Host "Found $($members.Count) members in source group." -ForegroundColor Green
        
        # Copy members to target group
        $successCount = 0
        $errorCount = 0
        $alreadyExistCount = 0
        
        foreach ($member in $members) {
            try {
                # Check if member is already in target group
                $isMember = Get-ADGroupMember -Identity $TargetGroup | Where-Object { $_.distinguishedName -eq $member.distinguishedName }
                
                if ($isMember) {
                    Write-Host "Member $($member.Name) is already in target group." -ForegroundColor Yellow
                    $alreadyExistCount++
                    continue
                }
                
                Add-ADGroupMember -Identity $TargetGroup -Members $member.distinguishedName -ErrorAction Stop
                Write-Host "Added member $($member.Name) to target group." -ForegroundColor Green
                $successCount++
            } catch {
                Write-Host "Error adding member $($member.Name): $_" -ForegroundColor Red
                $errorCount++
            }
        }
        
        Write-Host "`nSummary:" -ForegroundColor Cyan
        Write-Host "Total members in source group: $($members.Count)" -ForegroundColor White
        Write-Host "Successfully added: $successCount" -ForegroundColor Green
        Write-Host "Already members: $alreadyExistCount" -ForegroundColor Yellow
        Write-Host "Failed to add: $errorCount" -ForegroundColor Red
        
        return $true
    } catch {
        Write-Host "Error copying AD group members: $_" -ForegroundColor Red
        return $false
    }
}

function Copy-DistributionGroupMembers {
    param (
        [Parameter(Mandatory)]
        [string]$SourceGroup,
        [Parameter(Mandatory)]
        [string]$TargetGroup
    )
    try {
        # Verify both groups exist
        $sourceExists = Test-GroupExists -GroupIdentity $SourceGroup -DistributionGroup
        if (-not $sourceExists) {
            Write-Host "Source group '$SourceGroup' does not exist in Exchange." -ForegroundColor Red
            return $false
        }
        
        $targetExists = Test-GroupExists -GroupIdentity $TargetGroup -DistributionGroup
        if (-not $targetExists) {
            Write-Host "Target group '$TargetGroup' does not exist in Exchange." -ForegroundColor Red
            return $false
        }
        
        # Get members from source group
        Write-Host "Getting members from source group '$SourceGroup'..." -ForegroundColor Yellow
        $members = Get-DistributionGroupMember -Identity $SourceGroup -ErrorAction Stop
        
        if ($members.Count -eq 0) {
            Write-Host "Source group has no members." -ForegroundColor Yellow
            return $true
        }
        
        Write-Host "Found $($members.Count) members in source group." -ForegroundColor Green
        
        # Copy members to target group
        $successCount = 0
        $errorCount = 0
        $alreadyExistCount = 0
        
        # Get current target group members for comparison
        $currentMembers = Get-DistributionGroupMember -Identity $TargetGroup
        $currentMemberAddresses = $currentMembers | ForEach-Object { $_.PrimarySmtpAddress.ToString() }
        
        foreach ($member in $members) {
            try {
                # Check if member is already in target group
                if ($currentMemberAddresses -contains $member.PrimarySmtpAddress.ToString()) {
                    Write-Host "Member $($member.Name) is already in target group." -ForegroundColor Yellow
                    $alreadyExistCount++
                    continue
                }
                
                Add-DistributionGroupMember -Identity $TargetGroup -Member $member.Identity -ErrorAction Stop
                Write-Host "Added member $($member.Name) to target group." -ForegroundColor Green
                $successCount++
            } catch {
                Write-Host "Error adding member $($member.Name): $_" -ForegroundColor Red
                $errorCount++
            }
        }
        
        Write-Host "`nSummary:" -ForegroundColor Cyan
        Write-Host "Total members in source group: $($members.Count)" -ForegroundColor White
        Write-Host "Successfully added: $successCount" -ForegroundColor Green
        Write-Host "Already members: $alreadyExistCount" -ForegroundColor Yellow
        Write-Host "Failed to add: $errorCount" -ForegroundColor Red
        
        return $true
    } catch {
        Write-Host "Error copying distribution group members: $_" -ForegroundColor Red
        return $false
    }
}

# Main script
Clear-Host
Write-Host "=== Group Member Copy Tool ===" -ForegroundColor Cyan

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
        
        # Show operation options
        Write-Host "`nSelect Operation:" -ForegroundColor Cyan
        Write-Host "1. Copy members to existing group"
        Write-Host "2. Copy members to new group"
        Write-Host "3. Back to group type selection"
        
        $operationChoice = Read-Host "`nEnter your choice (1-3)"
        
        if ($operationChoice -eq '3') { continue }
        
        if ($operationChoice -in '1','2') {
            $createNew = ($operationChoice -eq '2')
            
            # Get source group
            $sourceGroup = Read-Host "`nEnter source group name"
            
            # Verify source group exists
            $sourceExists = if ($isExchangeGroup) {
                Test-GroupExists -GroupIdentity $sourceGroup -DistributionGroup
            } else {
                Test-GroupExists -GroupIdentity $sourceGroup
            }
            
            if (-not $sourceExists) {
                Write-Host "Source group '$sourceGroup' does not exist." -ForegroundColor Red
                continue
            }
            
            # Handle target group based on operation
            $targetGroup = $null
            
            if ($createNew) {
                $targetGroup = Read-Host "Enter new group name"
                
                if ($isExchangeGroup) {
                    # Create new Exchange distribution group
                    $alias = Read-Host "Enter alias (leave blank to auto-generate)"
                    $description = Read-Host "Enter description (optional)"
                    $securityEnabled = (Read-Host "Enable security for this group? (Y/N)") -eq 'Y'
                    
                    $newGroup = New-ExchangeDistributionGroup -Name $targetGroup -Alias $alias -Description $description -SecurityEnabled:$securityEnabled
                    if (-not $newGroup) { continue }
                } else {
                    # Create new AD group
                    $description = Read-Host "Enter description (optional)"
                    Write-Host "`nSelect Group Scope:" -ForegroundColor Cyan
                    Write-Host "1. Global (default)"
                    Write-Host "2. DomainLocal"
                    Write-Host "3. Universal"
                    $scopeChoice = Read-Host "`nEnter your choice (1-3)"
                    
                    $scope = switch ($scopeChoice) {
                        "2" { "DomainLocal" }
                        "3" { "Universal" }
                        default { "Global" }
                    }
                    
                    Write-Host "`nSelect Group Type:" -ForegroundColor Cyan
                    Write-Host "1. Security (default)"
                    Write-Host "2. Distribution"
                    $typeChoice = Read-Host "`nEnter your choice (1-2)"
                    
                    $type = if ($typeChoice -eq "2") { "Distribution" } else { "Security" }
                    
                    $path = Read-Host "Enter OU path (optional, e.g., 'OU=Groups,DC=contoso,DC=com')"
                    
                    $newGroup = New-ADSecurityGroup -Name $targetGroup -Description $description -Path $path -Scope $scope -Type $type
                    if (-not $newGroup) { continue }
                }
            } else {
                $targetGroup = Read-Host "Enter target group name"
                
                # Verify target group exists
                $targetExists = if ($isExchangeGroup) {
                    Test-GroupExists -GroupIdentity $targetGroup -DistributionGroup
                } else {
                    Test-GroupExists -GroupIdentity $targetGroup
                }
                
                if (-not $targetExists) {
                    Write-Host "Target group '$targetGroup' does not exist." -ForegroundColor Red
                    continue
                }
            }
            
            # Confirm before proceeding
            $confirm = Read-Host "`nCopy members from '$sourceGroup' to '$targetGroup'? (Y/N)"
            
            if ($confirm -eq 'Y' -or $confirm -eq 'y') {
                Write-Host "`nCopying group members..." -ForegroundColor Cyan
                
                if ($isExchangeGroup) {
                    Copy-DistributionGroupMembers -SourceGroup $sourceGroup -TargetGroup $targetGroup
                } else {
                    Copy-ADGroupMembers -SourceGroup $sourceGroup -TargetGroup $targetGroup
                }
                
                Write-Host "`nOperation completed." -ForegroundColor Green
            }
        }
    }
} while ($true)