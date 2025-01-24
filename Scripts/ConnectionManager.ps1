# Global variable to track connection state
$script:ExchangeConnection = @{
    IsConnected = $false
    CurrentUser = $null
    OrganizationName = $null
}

function Get-ExchangeConnectionStatus {
    try {
        $session = Get-PSSession | Where-Object {
            ($_.ConfigurationName -eq "Microsoft.Exchange" -or $_.Name -like "ExchangeOnline*") -and 
            $_.State -eq "Opened"
        }
        
        if ($session) {
            $orgInfo = Get-OrganizationConfig
            $script:ExchangeConnection.IsConnected = $true
            $script:ExchangeConnection.CurrentUser = $session.Runspace.ConnectionInfo.Credential.UserName
            $script:ExchangeConnection.OrganizationName = $orgInfo.DisplayName
            return $true
        }
        
        $script:ExchangeConnection.IsConnected = $false
        $script:ExchangeConnection.CurrentUser = $null
        $script:ExchangeConnection.OrganizationName = $null
        return $false
    } catch {
        $script:ExchangeConnection.IsConnected = $false
        return $false
    }
}

function Connect-ExchangeOnlineSession {
    param (
        [Parameter(Mandatory=$false)]
        [string]$AdminUser,
        [switch]$Force
    )
    
    try {
        # Check if already connected
        if ((Get-ExchangeConnectionStatus) -and -not $Force) {
            Write-Host "`nCurrently connected to Exchange Online:" -ForegroundColor Cyan
            Write-Host "Organization: $($script:ExchangeConnection.OrganizationName)" -ForegroundColor Green
            Write-Host "User: $($script:ExchangeConnection.CurrentUser)" -ForegroundColor Green
            
            $switch = Read-Host "`nDo you want to switch to a different account? (Y/N)"
            if ($switch -ne 'Y' -and $switch -ne 'y') {
                return $true
            }
            
            # Disconnect current session if switching
            Disconnect-ExchangeOnlineSession
        }

        # Verify Exchange Online module
        if (!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
            Write-Host "Exchange Online PowerShell module is not installed." -ForegroundColor Red
            Write-Host "Please install it by running: Install-Module -Name ExchangeOnlineManagement" -ForegroundColor Yellow
            return $false
        }

        # Get admin credentials if not provided
        if (-not $AdminUser) {
            $AdminUser = Read-Host "`nEnter your Exchange Online admin email address"
        }

        # Import module and connect
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        Connect-ExchangeOnline -UserPrincipalName $AdminUser -ShowProgress $false -ShowBanner:$false
        
        # Test connection by running a simple Exchange command
        try {
            $org = Get-OrganizationConfig -ErrorAction Stop
            $script:ExchangeConnection.IsConnected = $true
            $script:ExchangeConnection.CurrentUser = $AdminUser
            $script:ExchangeConnection.OrganizationName = $org.DisplayName
            Write-Host "`nSuccessfully connected to Exchange Online:" -ForegroundColor Green
            Write-Host "Organization: $($org.DisplayName)" -ForegroundColor Green
            Write-Host "User: $AdminUser" -ForegroundColor Green
            return $true
        } catch {
            Write-Host "`nFailed to verify Exchange connection: $_" -ForegroundColor Red
            Write-Host "Press Enter to continue..." -ForegroundColor Yellow
            Read-Host
            return $false
        }
    } catch {
        Write-Host "`nError connecting to Exchange Online:" -ForegroundColor Red
        Write-Host $_ -ForegroundColor Red
        Write-Host "`nPress Enter to continue..." -ForegroundColor Yellow
        Read-Host
        return $false
    }
}

function Disconnect-ExchangeOnlineSession {
    try {
        if (Get-ExchangeConnectionStatus) {
            Disconnect-ExchangeOnline -Confirm:$false
            $script:ExchangeConnection.IsConnected = $false
            $script:ExchangeConnection.CurrentUser = $null
            $script:ExchangeConnection.OrganizationName = $null
            Write-Host "Successfully disconnected from Exchange Online." -ForegroundColor Green
            return $true
        }
        return $false
    } catch {
        Write-Host "Error disconnecting from Exchange Online: $_" -ForegroundColor Red
        return $false
    }
}