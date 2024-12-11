# Main menu script for administrative tools
function Show-MainMenu {
    Clear-Host
    Write-Host "=== PowerShell Administrative Tools ===" -ForegroundColor Cyan
    Write-Host "1. Exchange Online Archive Manager" -ForegroundColor Yellow
    Write-Host "2. Active Directory Group Member Export" -ForegroundColor Yellow
    Write-Host "3. Distribution Group Member Export" -ForegroundColor Yellow
    Write-Host "4. Calendar Permissions Manager" -ForegroundColor Yellow
    Write-Host "5. Exit" -ForegroundColor Yellow
    Write-Host "=====================================`n" -ForegroundColor Cyan
}

function Invoke-Tool {
    param (
        [Parameter(Mandatory)]
        [string]$ScriptPath,
        [Parameter(Mandatory)]
        [string]$ToolName
    )
    try {
        Clear-Host  # Clear screen before loading tool
        Write-Host "=== $ToolName ===" -ForegroundColor Cyan
        Write-Host "Loading..." -ForegroundColor Gray
        
        # Get the current script's directory, fallback to current directory if not available
        $scriptDir = if ($PSScriptRoot) {
            $PSScriptRoot
        } else {
            $PWD.Path
        }
        
        # Construct full path to the target script
        $fullPath = Join-Path -Path $scriptDir -ChildPath $ScriptPath
        
        if (Test-Path $fullPath) {
            Clear-Host  # Clear loading message before executing tool
            . $fullPath
            Write-Host "`nPress any key to return to main menu..."
            $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        } else {
            Write-Host "Script not found at: $fullPath" -ForegroundColor Red
            Write-Host "Please ensure the following script exists: $ScriptPath" -ForegroundColor Yellow
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
    $choice = Read-Host "Select an option (1-5)"
    
    switch ($choice) {
        "1" { Invoke-Tool -ScriptPath "EnableAutoIncrementArchiving.ps1" -ToolName "Exchange Archive Manager" }
        "2" { Invoke-Tool -ScriptPath "ExportGrouptoCSV.ps1" -ToolName "AD Group Export" }
        "3" { Invoke-Tool -ScriptPath "ExportDistributionGroupMembers" -ToolName "Distribution Group Export" }
        "4" { Invoke-Tool -ScriptPath "CalendarTools.ps1" -ToolName "Calendar Permissions Manager" }
        "5" { 
            Write-Host "`nExiting..." -ForegroundColor Green
            exit 
        }
        default { 
            Write-Host "`nInvalid option. Please try again." -ForegroundColor Red
            Start-Sleep -Seconds 2
        }
    }
} while ($true)