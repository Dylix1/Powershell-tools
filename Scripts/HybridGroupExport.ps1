# All functions remain the same, but we modify the main execution block
function Test-ADGroupExists {
    param([string]$GroupName)
    try {
        $null = Get-ADGroup -Identity $GroupName
        return $true
    } catch {
        return $false
    }
}

function Test-O365GroupExists {
    param([string]$GroupName)
    try {
        $null = Get-UnifiedGroup -Identity $GroupName -ErrorAction Stop
        return $true
    } catch {
        try {
            $null = Get-DistributionGroup -Identity $GroupName -ErrorAction Stop
            return $true
        } catch {
            return $false
        }
    }
}

function Get-SafeFilePath {
    param(
        [string]$BasePath,
        [string]$GroupName,
        [string]$Source
    )
    $safeGroupName = [RegEx]::Replace($GroupName, "[{0}]" -f ([RegEx]::Escape([String][System.IO.Path]::GetInvalidFileNameChars())), '_')
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    return Join-Path -Path $BasePath -ChildPath "GroupMembers_${Source}_${safeGroupName}_${timestamp}.csv"
}

function Export-ADGroupMembers {
    param([string]$GroupName)
    try {
        if (-not (Test-ADGroupExists -GroupName $GroupName)) {
            throw "Group '$GroupName' does not exist in Active Directory."
        }

        $OutputFile = Get-SafeFilePath -BasePath $env:TEMP -GroupName $GroupName -Source "AD"

        Write-Host "Retrieving AD group members..." -ForegroundColor Yellow
        $GroupMembers = Get-ADGroupMember -Identity $GroupName -Recursive | 
            Where-Object { $_.objectClass -eq 'user' }

        if (-not $GroupMembers) {
            Write-Host "No members found in the AD group '$GroupName'." -ForegroundColor Yellow
            return
        }

        Write-Host "Processing member details..." -ForegroundColor Yellow
        $Results = foreach ($Member in $GroupMembers) {
            try {
                Get-ADUser -Identity $Member.distinguishedName -Properties DisplayName, EmailAddress, SamAccountName |
                    Select-Object DisplayName, EmailAddress, SamAccountName, @{
                        Name='Retrieved'; 
                        Expression={(Get-Date).ToString("yyyy-MM-dd HH:mm:ss")}
                    }
            } catch {
                Write-Warning "Could not retrieve details for user: $($Member.distinguishedName)"
                continue
            }
        }

        $Results | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8

        if (Test-Path $OutputFile) {
            Write-Host "Export complete. File saved at: $OutputFile" -ForegroundColor Green
            Write-Host "Total members exported: $($Results.Count)" -ForegroundColor Green
        } else {
            throw "Failed to create output file."
        }
    } catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Export-O365GroupMembers {
    param([string]$GroupName)
    try {
        if (-not (Test-O365GroupExists -GroupName $GroupName)) {
            throw "Group '$GroupName' does not exist in Office 365."
        }

        $OutputFile = Get-SafeFilePath -BasePath $env:TEMP -GroupName $GroupName -Source "O365"
        
        Write-Host "Retrieving Office 365 group members..." -ForegroundColor Yellow
        
        # Try as Unified Group first
        try {
            $Members = Get-UnifiedGroupLinks -Identity $GroupName -LinkType Members -ErrorAction Stop
            $GroupType = "Unified"
        } catch {
            # If not Unified, try as Distribution Group
            try {
                $Members = Get-DistributionGroupMember -Identity $GroupName -ErrorAction Stop
                $GroupType = "Distribution"
            } catch {
                throw "Failed to retrieve group members. Check if you have appropriate permissions."
            }
        }

        if (-not $Members) {
            Write-Host "No members found in the Office 365 $GroupType group '$GroupName'." -ForegroundColor Yellow
            return
        }

        Write-Host "Processing member details..." -ForegroundColor Yellow
        $Results = $Members | Select-Object DisplayName, 
            @{Name='EmailAddress'; Expression={$_.PrimarySmtpAddress}},
            @{Name='UserPrincipalName'; Expression={$_.WindowsLiveID}},
            @{Name='GroupType'; Expression={$GroupType}},
            @{Name='Retrieved'; Expression={(Get-Date).ToString("yyyy-MM-dd HH:mm:ss")}}

        $Results | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8

        if (Test-Path $OutputFile) {
            Write-Host "Export complete. File saved at: $OutputFile" -ForegroundColor Green
            Write-Host "Total members exported: $($Results.Count)" -ForegroundColor Green
        } else {
            throw "Failed to create output file."
        }
    } catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Copy-ADScriptToClipboard {
    param([string]$GroupName)
    try {
        if ([string]::IsNullOrWhiteSpace($GroupName)) {
            throw "No group name provided."
        }

        $ScriptContent = @"
# Import the Active Directory module
Import-Module ActiveDirectory -ErrorAction Stop

# Define the Active Directory group name
`$ADGroupName = '$GroupName'

# Verify group exists
try {
    `$null = Get-ADGroup -Identity `$ADGroupName
} catch {
    Write-Host "Error: Group '`$ADGroupName' does not exist in Active Directory." -ForegroundColor Red
    exit 1
}

# Create output file path with timestamp
`$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
`$OutputFile = Join-Path -Path `$env:TEMP -ChildPath "GroupMembers_AD_`$(`$ADGroupName)_`${timestamp}.csv"

try {
    Write-Host "Retrieving group members..." -ForegroundColor Yellow
    `$GroupMembers = Get-ADGroupMember -Identity `$ADGroupName -Recursive | 
        Where-Object { `$_.objectClass -eq 'user' }

    if (-not `$GroupMembers) {
        Write-Host "No members found in the group '`$ADGroupName'." -ForegroundColor Yellow
        exit 0
    }

    Write-Host "Processing member details..." -ForegroundColor Yellow
    `$Results = foreach (`$Member in `$GroupMembers) {
        try {
            Get-ADUser -Identity `$Member.distinguishedName -Properties DisplayName, EmailAddress, SamAccountName |
                Select-Object DisplayName, EmailAddress, SamAccountName, @{
                    Name='Retrieved'; 
                    Expression={(Get-Date).ToString("yyyy-MM-dd HH:mm:ss")}
                }
        } catch {
            Write-Warning "Could not retrieve details for user: `$(`$Member.distinguishedName)"
            continue
        }
    }

    `$Results | Export-Csv -Path `$OutputFile -NoTypeInformation -Encoding UTF8

    if (Test-Path `$OutputFile) {
        Write-Host "Export complete. File saved at: `$OutputFile" -ForegroundColor Green
        Write-Host "Total members exported: `$(`$Results.Count)" -ForegroundColor Green
    } else {
        throw "Failed to create output file."
    }
} catch {
    Write-Host "Error: `$(`$_.Exception.Message)" -ForegroundColor Red
    exit 1
}
"@

        Set-Clipboard -Value $ScriptContent
        Write-Host "Script copied to clipboard successfully with group name placeholder set to '$GroupName'." -ForegroundColor Green
    } catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Copy-O365ScriptToClipboard {
    param([string]$GroupName)
    try {
        if ([string]::IsNullOrWhiteSpace($GroupName)) {
            throw "No group name provided."
        }

        $ScriptContent = @"
# Import the Exchange Online Management module
Import-Module ExchangeOnlineManagement -ErrorAction Stop

# Define the Office 365 group name
`$GroupName = '$GroupName'

# Connect to Exchange Online
Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
Connect-ExchangeOnline -ShowBanner:`$false

try {
    # Create output file path with timestamp
    `$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    `$OutputFile = Join-Path -Path `$env:TEMP -ChildPath "GroupMembers_O365_`$(`$GroupName)_`${timestamp}.csv"
    
    # Try as Unified Group first
    try {
        `$Members = Get-UnifiedGroupLinks -Identity `$GroupName -LinkType Members -ErrorAction Stop
        `$GroupType = "Unified"
    } catch {
        # If not Unified, try as Distribution Group
        try {
            `$Members = Get-DistributionGroupMember -Identity `$GroupName -ErrorAction Stop
            `$GroupType = "Distribution"
        } catch {
            throw "Failed to retrieve group members. Check if you have appropriate permissions."
        }
    }

    if (-not `$Members) {
        Write-Host "No members found in the Office 365 `$GroupType group '`$GroupName'." -ForegroundColor Yellow
        exit 0
    }

    Write-Host "Processing member details..." -ForegroundColor Yellow
    `$Results = `$Members | Select-Object DisplayName, 
        @{Name='EmailAddress'; Expression={`$_.PrimarySmtpAddress}},
        @{Name='UserPrincipalName'; Expression={`$_.WindowsLiveID}},
        @{Name='GroupType'; Expression={`$GroupType}},
        @{Name='Retrieved'; Expression={(Get-Date).ToString("yyyy-MM-dd HH:mm:ss")}}

    `$Results | Export-Csv -Path `$OutputFile -NoTypeInformation -Encoding UTF8

    if (Test-Path `$OutputFile) {
        Write-Host "Export complete. File saved at: `$OutputFile" -ForegroundColor Green
        Write-Host "Total members exported: `$(`$Results.Count)" -ForegroundColor Green
    } else {
        throw "Failed to create output file."
    }
} catch {
    Write-Host "Error: `$(`$_.Exception.Message)" -ForegroundColor Red
    exit 1
} finally {
    # Disconnect from Exchange Online
    Disconnect-ExchangeOnline -Confirm:`$false
}
"@

        Set-Clipboard -Value $ScriptContent
        Write-Host "Script copied to clipboard successfully with group name placeholder set to '$GroupName'." -ForegroundColor Green
    } catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Main execution block - Modified to work with the menu system
Clear-Host
Write-Host "=== Hybrid Group Member Export Tool ===" -ForegroundColor Cyan

Write-Host "`nSelect group source:" -ForegroundColor Cyan
Write-Host "1. Active Directory (On-premises)" -ForegroundColor Yellow
Write-Host "2. Office 365" -ForegroundColor Yellow
Write-Host "3. Return to Main Menu" -ForegroundColor Yellow

$SourceChoice = Read-Host -Prompt "Enter your choice (1-3)"

if ($SourceChoice -in "1","2") {
    Write-Host "`nSelect execution mode:" -ForegroundColor Cyan
    Write-Host "1. Local device execution" -ForegroundColor Yellow
    Write-Host "2. Copy script to clipboard" -ForegroundColor Yellow
    $ExecutionMode = Read-Host -Prompt "Enter your choice (1-2)"

    if ($ExecutionMode -notin "1","2") {
        Write-Host "Invalid execution mode selection." -ForegroundColor Red
        Start-Sleep -Seconds 2
        return
    }

    switch ($SourceChoice) {
        "1" {
            # Check if AD module is available
            if (!(Get-Module -ListAvailable -Name ActiveDirectory)) {
                Write-Host "Active Directory PowerShell module is not installed." -ForegroundColor Red
                Start-Sleep -Seconds 2
                return
            }
            
            $GroupName = Read-Host -Prompt "Enter the Active Directory group name"
            if (-not [string]::IsNullOrWhiteSpace($GroupName)) {
                if ($ExecutionMode -eq "1") {
                    Export-ADGroupMembers -GroupName $GroupName
                } else {
                    Copy-ADScriptToClipboard -GroupName $GroupName
                }
            }
        }
        "2" {
            $GroupName = Read-Host -Prompt "Enter the Office 365 group name or email address"
            if (-not [string]::IsNullOrWhiteSpace($GroupName)) {
                if ($ExecutionMode -eq "1") {
                    if (Connect-O365) {
                        Export-O365GroupMembers -GroupName $GroupName
                    }
                } else {
                    Copy-O365ScriptToClipboard -GroupName $GroupName
                }
            }
        }
    }
} elseif ($SourceChoice -eq "3") {
    return
} else {
    Write-Host "Invalid selection." -ForegroundColor Red
    Start-Sleep -Seconds 2
}