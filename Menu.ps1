# Import connection manager and set error preference
$ErrorActionPreference = 'Stop'
. $PSScriptRoot\Scripts\ConnectionManager.ps1

function Show-MainMenu {
    Clear-Host
    Write-Host "=== PowerShell Administrative Tools ===" -ForegroundColor Cyan
    
    if ($script:ExchangeConnection.IsConnected) {
        Write-Host "`nConnected to: $($script:ExchangeConnection.OrganizationName)" -ForegroundColor Green
        Write-Host "User: $($script:ExchangeConnection.CurrentUser)`n" -ForegroundColor Green
    } else {
        Write-Host "`nNot connected to Exchange Online`n" -ForegroundColor Yellow
    }

    Write-Host "0. Connect/Switch Exchange Online Account" -ForegroundColor Yellow
    Write-Host "1. Exchange Online Archive Manager" -ForegroundColor Yellow
    Write-Host "2. Active Directory Group Member Export" -ForegroundColor Yellow
    Write-Host "3. Distribution Group Member Export" -ForegroundColor Yellow
    Write-Host "4. Calendar Permissions Manager" -ForegroundColor Yellow
    Write-Host "5. User Shared-Mailbox Access" -ForegroundColor Yellow
    Write-Host "6. Add Mailbox Permissions" -ForegroundColor Yellow
    Write-Host "7. Copy Group Memberships from user to user" -ForegroundColor Yellow
    Write-Host "8. Hybrid Group Member Export" -ForegroundColor Yellow
    Write-Host "9. Sharepoint Site Permissions" -ForegroundColor Yellow
    Write-Host "10. Exit" -ForegroundColor Yellow
    Write-Host "=====================================`n" -ForegroundColor Cyan
}

function Invoke-Tool {
    param (
        [Parameter(Mandatory)][string]$ScriptPath,
        [Parameter(Mandatory)][string]$ToolName
    )
    try {
        # Check for Exchange connection if needed
        if ($ScriptPath -notmatch "ExportGrouptoCSV|HybridGroupExport" -and -not $script:ExchangeConnection.IsConnected) {
            Write-Host "`nExchange Online connection required." -ForegroundColor Yellow
            if (-not (Connect-ExchangeOnlineSession)) {
                return
            }
        }

        Clear-Host
        Write-Host "=== $ToolName ===" -ForegroundColor Cyan
        
        $scriptsFolder = Join-Path -Path $PSScriptRoot -ChildPath "Scripts"
        $fullPath = Join-Path -Path $scriptsFolder -ChildPath $ScriptPath

        if (!(Test-Path $scriptsFolder)) {
            New-Item -ItemType Directory -Path $scriptsFolder | Out-Null
        }
        
        if (Test-Path $fullPath) {
            Clear-Host
            . $fullPath
            Write-Host "`nPress any key to return to main menu..."
            $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        } else {
            Write-Host "Script not found: $fullPath" -ForegroundColor Red
            Start-Sleep -Seconds 3
        }
    }
    catch {
        Write-Host "Error loading script: $_" -ForegroundColor Red
        Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
        Start-Sleep -Seconds 3
    }
}

# Main loop
do {
    Show-MainMenu
    $choice = Read-Host "Select an option (0-10)"
    
    switch ($choice) {
        "0" { Connect-ExchangeOnlineSession -Force }
        "1" { Invoke-Tool -ScriptPath "EnableAutoIncrementArchiving.ps1" -ToolName "Exchange Archive Manager" }
        "2" { Invoke-Tool -ScriptPath "ExportGrouptoCSV.ps1" -ToolName "AD Group Export" }
        "3" { Invoke-Tool -ScriptPath "ExportDistributionGroupMembers.ps1" -ToolName "Distribution Group Export" }
        "4" { Invoke-Tool -ScriptPath "CalendarTools.ps1" -ToolName "Calendar Permissions Manager" }
        "5" { Invoke-Tool -ScriptPath "GetSharedMailboxPermissions.ps1" -ToolName "User Shared-Mailbox Access" }
        "6" { Invoke-Tool -ScriptPath "AddMailboxPermissions.ps1" -ToolName "Add Mailbox Permissions" }
        "7" { Invoke-Tool -ScriptPath "CopyGroupMemberships.ps1" -ToolName "Copy Group Memberships" }
        "8" { Invoke-Tool -ScriptPath "HybridGroupExport.ps1" -ToolName "Hybrid Group Member Export" }
        "9" { Invoke-Tool -ScriptPath "GetSharepointPermissions.ps1" -ToolName "Sharepoint Site Permissions" }
        "10" { 
            if ($script:ExchangeConnection.IsConnected) {
                Disconnect-ExchangeOnlineSession
            }
            Clear-Host
            Write-Host "`nExiting..." -ForegroundColor Green
            exit 
        }
        default { 
            Write-Host "`nInvalid option. Please try again." -ForegroundColor Red
            Start-Sleep -Seconds 2
        }
    }
} while ($true)