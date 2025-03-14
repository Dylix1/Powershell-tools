<#
.SYNOPSIS
    Shows which SharePoint groups a user belongs to across sites.
.DESCRIPTION
    This script uses the SharePoint Online Management Shell to identify all the SharePoint groups 
    a specified user is a member of across all accessible site collections.
.NOTES
    File Name      : Get-SPOUserGroupMembership.ps1
    Author         : Your Name
    Prerequisite   : SharePoint Online Management Shell
    Version        : 1.0
#>

function Connect-ToSharePointOnline {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$AdminUrl,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential
    )
    
    try {
        Write-Host "Connecting to SharePoint Online" -ForegroundColor Cyan
        
        if ($null -eq $Credential) {
            # Use interactive login
            Connect-SPOService -Url $AdminUrl -ErrorAction Stop
        }
        else {
            # Use provided credentials
            Connect-SPOService -Url $AdminUrl -Credential $Credential -ErrorAction Stop
        }
        
        Write-Host "Successfully connected to SharePoint Online" -ForegroundColor Green
        return $true
    }
    catch {
        $errorMsg = $Error[0].Exception.Message
        Write-Host "Error connecting to SharePoint Online" -ForegroundColor Red
        Write-Host "Reason $errorMsg" -ForegroundColor Red
        return $false
    }
}

function Get-SPOSites {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [int]$MaxSites = 0
    )
    
    try {
        Write-Host "Retrieving SharePoint site collections" -ForegroundColor Cyan
        $sites = Get-SPOSite -Limit All -ErrorAction Stop
        
        if ($MaxSites -gt 0 -and $sites.Count -gt $MaxSites) {
            Write-Host "Limiting to first $MaxSites sites" -ForegroundColor Yellow
            $sites = $sites | Select-Object -First $MaxSites
        }
        
        Write-Host "Found $($sites.Count) site collections" -ForegroundColor Green
        return $sites
    }
    catch {
        $errorMsg = $Error[0].Exception.Message
        Write-Host "Error retrieving site collections" -ForegroundColor Red
        Write-Host "Reason $errorMsg" -ForegroundColor Red
        return $null
    }
}

function Find-UserGroupMemberships {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,
        
        [Parameter(Mandatory = $false)]
        [int]$MaxSites = 0
    )
    
    try {
        # Get all sites
        $sites = Get-SPOSites -MaxSites $MaxSites
        if ($null -eq $sites) {
            return $null
        }
        
        $results = @()
        $processedSites = 0
        $totalSites = $sites.Count
        $userGroups = 0
        
        foreach ($site in $sites) {
            $processedSites++
            $percentComplete = [math]::Round(($processedSites / $totalSites) * 100)
            
            Write-Progress -Activity "Searching SharePoint Groups" -Status "Processing site $processedSites of $totalSites ($percentComplete%)" `
                -PercentComplete $percentComplete -CurrentOperation $site.Url
            
            try {
                # Get site users
                Write-Host "Checking site $($site.Url)" -ForegroundColor Gray
                $siteUsers = Get-SPOUser -Site $site.Url -ErrorAction SilentlyContinue
                
                # Look for our user
                $user = $siteUsers | Where-Object { 
                    $_.LoginName -like "*$UserEmail*" -or 
                    $_.Email -like "*$UserEmail*" -or 
                    $_.DisplayName -like "*$UserEmail*" 
                }
                
                if ($user) {
                    # Get groups for this user
                    foreach ($group in $user.Groups) {
                        $userGroups++
                        
                        # Get all site groups to find group details
                        $siteGroups = Get-SPOSiteGroup -Site $site.Url -ErrorAction SilentlyContinue
                        $groupDetails = $siteGroups | Where-Object { $_.Title -eq $group }
                        
                        $result = [PSCustomObject]@{
                            SiteUrl = $site.Url
                            SiteTitle = $site.Title
                            GroupName = $group
                            PermissionLevels = ($groupDetails.Roles -join ", ")
                            UserLoginName = $user.LoginName
                            UserEmail = $user.Email
                            UserDisplayName = $user.DisplayName
                        }
                        
                        $results += $result
                        Write-Host "Found user in group $group on site $($site.Title)" -ForegroundColor Green
                    }
                }
            }
            catch {
                $errorMsg = $Error[0].Exception.Message
                Write-Host "Error processing site $($site.Url)" -ForegroundColor Red
                Write-Host "Reason $errorMsg" -ForegroundColor Red
            }
        }
        
        Write-Progress -Activity "Searching SharePoint Groups" -Completed
        
        if ($results.Count -gt 0) {
            Write-Host "`nUser '$UserEmail' is a member of $userGroups SharePoint groups across $($results.Count) sites" -ForegroundColor Cyan
            return $results
        }
        else {
            Write-Host "`nUser '$UserEmail' was not found in any SharePoint groups" -ForegroundColor Yellow
            return $null
        }
    }
    catch {
        $errorMsg = $Error[0].Exception.Message
        Write-Host "Error finding user group memberships" -ForegroundColor Red
        Write-Host "Reason $errorMsg" -ForegroundColor Red
        return $null
    }
}

function Export-ResultsToCsv {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [array]$Results,
        
        [Parameter(Mandatory = $false)]
        [string]$FilePath
    )
    
    if ([string]::IsNullOrEmpty($FilePath)) {
        $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $FilePath = "SharePointUserGroups_$timestamp.csv"
    }
    
    try {
        $Results | Export-Csv -Path $FilePath -NoTypeInformation -ErrorAction Stop
        Write-Host "Results exported to $FilePath" -ForegroundColor Green
        return $true
    }
    catch {
        $errorMsg = $Error[0].Exception.Message
        Write-Host "Error exporting results to CSV" -ForegroundColor Red
        Write-Host "Reason $errorMsg" -ForegroundColor Red
        return $false
    }
}

function Show-Menu {
    Clear-Host
    Write-Host "=== SharePoint User Group Membership Tool ===" -ForegroundColor Cyan
    Write-Host "Using SharePoint Online Management Shell" -ForegroundColor Cyan
    Write-Host
    
    # Check if running in SharePoint Online Management Shell
    if (-not (Get-Command Connect-SPOService -ErrorAction SilentlyContinue)) {
        Write-Host "This script requires the SharePoint Online Management Shell" -ForegroundColor Red
        Write-Host "Please run this script from the SharePoint Online Management Shell or install it from" -ForegroundColor Yellow
        Write-Host "https://www.microsoft.com/en-us/download/details.aspx?id=35588" -ForegroundColor Yellow
        Write-Host
        Read-Host "Press Enter to exit"
        return
    }
    
    $adminUrl = Read-Host "Enter SharePoint admin URL (e.g., https://contoso-admin.sharepoint.com)"
    
    # Ask for credential method
    $useCredential = $false
    $credChoice = Read-Host "Use saved credentials? (Y/N)"
    
    if ($credChoice -eq "Y" -or $credChoice -eq "y") {
        $useCredential = $true
        $cred = Get-Credential -Message "Enter credentials for SharePoint Online"
        $connected = Connect-ToSharePointOnline -AdminUrl $adminUrl -Credential $cred
    }
    else {
        $connected = Connect-ToSharePointOnline -AdminUrl $adminUrl
    }
    
    if (-not $connected) {
        Read-Host "Press Enter to continue"
        return
    }
    
    do {
        Write-Host
        Write-Host "Options" -ForegroundColor Cyan
        Write-Host "1. Find user's SharePoint group memberships"
        Write-Host "2. Exit"
        Write-Host
        
        $choice = Read-Host "Enter your choice (1-2)"
        
        switch ($choice) {
            "1" {
                $userToFind = Read-Host "Enter user email, login name, or display name"
                $maxSitesStr = Read-Host "Enter maximum number of sites to search (0 for all)"
                $maxSites = 0
                
                if ([int]::TryParse($maxSitesStr, [ref]$maxSites)) {
                    $results = Find-UserGroupMemberships -UserEmail $userToFind -MaxSites $maxSites
                    
                    if ($results -and $results.Count -gt 0) {
                        Write-Host
                        Write-Host "User Group Memberships" -ForegroundColor Cyan
                        $results | Format-Table -Property SiteTitle, GroupName, PermissionLevels -AutoSize
                        
                        $exportChoice = Read-Host "Export results to CSV? (Y/N)"
                        if ($exportChoice -eq "Y" -or $exportChoice -eq "y") {
                            Export-ResultsToCsv -Results $results
                        }
                    }
                }
                else {
                    Write-Host "Invalid number of sites entered" -ForegroundColor Red
                }
                
                Read-Host "Press Enter to continue"
            }
            "2" {
                Disconnect-SPOService -ErrorAction SilentlyContinue
                return
            }
        }
    } while ($true)
}

# Run the menu
Show-Menu