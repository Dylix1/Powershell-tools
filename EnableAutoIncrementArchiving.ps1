function Connect-ExchangeOnlineSession {
    param (
        [Parameter(Mandatory)]
        [string]$AdminUser,
        [switch]$UseBasicAuth
    )
    try {
        $params = @{
            UserPrincipalName = $AdminUser
            ShowProgress = $false
            ShowBanner = $false
        }
        if ($UseBasicAuth) {
            $params.Add('BasicAuthentication', $true)
        }
        Connect-ExchangeOnline @params
        Write-Host "Successfully connected to Exchange Online as $AdminUser." -ForegroundColor Green
        return $true
    } catch {
        Write-Host "Error connecting to Exchange Online: $_" -ForegroundColor Red
        return $false
    }
}

function Check-AutoExpandingArchive {
    param (
        [Parameter(Mandatory)]
        [string]$Mailbox
    )
    try {
        $mailboxInfo = Get-Mailbox -Identity $Mailbox -ErrorAction Stop | 
            Select-Object AutoExpandingArchiveEnabled, ArchiveStatus, ArchiveQuota, ArchiveWarningQuota

        $output = @{
            AutoExpanding = $mailboxInfo.AutoExpandingArchiveEnabled
            Status = $mailboxInfo.ArchiveStatus
            Quota = $mailboxInfo.ArchiveQuota
            WarningQuota = $mailboxInfo.ArchiveWarningQuota
        }
        return $output
    } catch {
        Write-Host "Error checking archive status for $Mailbox" -ForegroundColor Red
        return $null
    }
}

function Enable-AutoExpandingArchive {
    param (
        [Parameter(Mandatory)]
        [string]$Mailbox,
        [switch]$Force
    )
    try {
        $params = @{
            Identity = $Mailbox
            AutoExpandingArchive = $true
            ErrorAction = 'Stop'
        }
        if ($Force) {
            $params.Add('Confirm', $false)
        }
        Enable-Mailbox @params
        Write-Host "Successfully enabled AutoExpandingArchive for $Mailbox." -ForegroundColor Green
        return $true
    } catch {
        Write-Host "Error enabling AutoExpandingArchive for $Mailbox" -ForegroundColor Red
        return $false
    }
}

function Disconnect-ExchangeOnlineSession {
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
        Write-Host "Successfully disconnected from Exchange Online." -ForegroundColor Green
    } catch {
        Write-Host "Error disconnecting from Exchange Online" -ForegroundColor Red
    }
}

function Show-ArchiveStatus {
    param($status)
    Write-Host "`nArchive Status:" -ForegroundColor Cyan
    Write-Host "Auto-Expanding: $($status.AutoExpanding)"
    Write-Host "Status: $($status.Status)"
    Write-Host "Quota: $($status.Quota)"
    Write-Host "Warning Quota: $($status.WarningQuota)`n"
}

# Main script
try {
    $ErrorActionPreference = 'Stop'
    Write-Host "=== Exchange Online Archive Manager ===" -ForegroundColor Cyan

    $AdminUPN = Read-Host "Enter Exchange Online admin email"
    
    $maxRetries = 3
    $connected = $false
    for ($i = 1; $i -le $maxRetries -and !$connected; $i++) {
        $connected = Connect-ExchangeOnlineSession -AdminUser $AdminUPN
        if (!$connected -and $i -lt $maxRetries) {
            Write-Host "Retrying connection ($i/$maxRetries)..." -ForegroundColor Yellow
            Start-Sleep -Seconds 3
        }
    }
    if (!$connected) { exit 1 }

    do {
        $Mailbox = Read-Host "`nEnter mailbox to manage (or 'exit' to quit)"
        if ($Mailbox -eq 'exit') { break }

        $status = Check-AutoExpandingArchive -Mailbox $Mailbox
        if ($status) {
            Show-ArchiveStatus -status $status
            if (!$status.AutoExpanding) {
                $enable = Read-Host "Enable AutoExpandingArchive? (Y/N)"
                if ($enable -eq 'Y') {
                    Enable-AutoExpandingArchive -Mailbox $Mailbox -Force
                }
            }
        }
    } while ($true)

} catch {
    Write-Host "Critical error: $_" -ForegroundColor Red
} finally {
    if ($connected) {
        Disconnect-ExchangeOnlineSession
    }
}