# Main menu script for administrative tools
function Show-MainMenu {
    Clear-Host
    Write-Host "=== PowerShell Administrative Tools ===" -ForegroundColor Cyan
    Write-Host "1. Exchange Online Archive Manager" -ForegroundColor Yellow
    Write-Host "2. Active Directory Group Member Export" -ForegroundColor Yellow
    Write-Host "3. Distribution Group Member Export" -ForegroundColor Yellow
    Write-Host "4. Exit" -ForegroundColor Yellow
    Write-Host "=====================================`n" -ForegroundColor Cyan
}

function Invoke-Tool {
    param (
        [string]$ScriptPath,
        [string]$ToolName
    )
    try {
        Write-Host "`nLoading $ToolName..." -ForegroundColor Cyan
        . $ScriptPath
        Write-Host "`nPress any key to return to main menu..."
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    }
    catch {
        Write-Host "Error loading script: $_" -ForegroundColor Red
        Start-Sleep -Seconds 3
    }
}

# Main loop
do {
    Show-MainMenu
    $choice = Read-Host "Select an option (1-4)"
    
    switch ($choice) {
        "1" { Invoke-Tool -ScriptPath ".\ExchangeArchiveManager.ps1" -ToolName "Exchange Archive Manager" }
        "2" { Invoke-Tool -ScriptPath ".\ADGroupExport.ps1" -ToolName "AD Group Export" }
        "3" { Invoke-Tool -ScriptPath ".\DistributionGroupExport.ps1" -ToolName "Distribution Group Export" }
        "4" { 
            Write-Host "`nExiting..." -ForegroundColor Green
            exit 
        }
        default { 
            Write-Host "`nInvalid option. Please try again." -ForegroundColor Red
            Start-Sleep -Seconds 2
        }
    }
} while ($true)