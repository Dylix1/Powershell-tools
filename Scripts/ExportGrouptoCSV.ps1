function Test-ADGroupExists {
    param([string]$GroupName)
    try {
        $null = Get-ADGroup -Identity $GroupName
        return $true
    } catch {
        return $false
    }
}

function Get-SafeFilePath {
    param([string]$BasePath, [string]$GroupName)
    $safeGroupName = [RegEx]::Replace($GroupName, "[{0}]" -f ([RegEx]::Escape([String][System.IO.Path]::GetInvalidFileNameChars())), '_')
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    return Join-Path -Path $BasePath -ChildPath "GroupMembers_${safeGroupName}_${timestamp}.csv"
}

function Open-Locally {
    [CmdletBinding()]
    param()
    try {
        $GroupName = Read-Host -Prompt "Enter the Active Directory group name"

        if ([string]::IsNullOrWhiteSpace($GroupName)) {
            throw "No group name provided."
        }

        if (-not (Test-ADGroupExists -GroupName $GroupName)) {
            throw "Group '$GroupName' does not exist in Active Directory."
        }

        $OutputFile = Get-SafeFilePath -BasePath $env:TEMP -GroupName $GroupName

        Write-Host "Retrieving group members..." -ForegroundColor Yellow
        $GroupMembers = Get-ADGroupMember -Identity $GroupName -Recursive | 
            Where-Object { $_.objectClass -eq 'user' }

        if (-not $GroupMembers) {
            Write-Host "No members found in the group '$GroupName'." -ForegroundColor Yellow
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

function Copy-ScriptToClipboard {
    [CmdletBinding()]
    param()
    try {
        $GroupName = Read-Host -Prompt "Enter the Active Directory group name for the placeholder"
        
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
`$OutputFile = Join-Path -Path `$env:TEMP -ChildPath "GroupMembers_`$(`$ADGroupName)_`${timestamp}.csv"

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

# Main script
Clear-Host
Write-Host "=== Active Directory Group Member Export ===" -ForegroundColor Cyan

Write-Host "Select execution mode:" -ForegroundColor Cyan
Write-Host "1. Local device execution" -ForegroundColor Yellow
Write-Host "2. Copy script to clipboard" -ForegroundColor Yellow
$ExecutionMode = Read-Host -Prompt "Enter 1 or 2"

if ($ExecutionMode -notin @("1", "2")) {
    Write-Host "Invalid selection. Exiting." -ForegroundColor Red
    return
}

switch ($ExecutionMode) {
    "1" { Open-Locally }
    "2" { Copy-ScriptToClipboard }
}