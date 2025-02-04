# Function to verify SharePoint connection
function Test-SPOConnection {
    try {
        $null = Get-SPOTenant -ErrorAction Stop
        return $true
    } catch {
        Write-Host "Not connected to SharePoint Online. Please connect first using Connect-SPOService." -ForegroundColor Red
        return $false
    }
}

function Get-UserSPOSitePermissions {
    param (
        [string]$UserEmail,
        [string]$OutputFile
    )
    try {
        # Verify user exists and get all sites
        Write-Host "Getting all SharePoint sites..." -ForegroundColor Yellow
        $sites = Get-SPOSite -Limit All
        Write-Host "Found $($sites.Count) sites to check." -ForegroundColor Green

        $results = @()
        
        foreach ($site in $sites) {
            Write-Host "`nChecking permissions for site: $($site.Url)" -ForegroundColor Yellow
            
            try {
                # Get direct permissions
                $users = Get-SPOUser -Site $site.Url -ErrorAction Stop
                $userPermissions = $users | Where-Object { $_.LoginName -like "*$UserEmail*" }
                
                if ($userPermissions) {
                    foreach ($perm in $userPermissions) {
                        $results += [PSCustomObject]@{
                            SiteUrl = $site.Url
                            SiteTitle = $site.Title
                            Template = $site.Template
                            PermissionType = "Direct"
                            PermissionLevel = ($perm.Groups -join ", ")
                            GrantedThrough = "Direct Assignment"
                        }
                    }
                }

                # Get group memberships
                $groups = Get-SPOSiteGroup -Site $site.Url -ErrorAction Stop
                foreach ($group in $groups) {
                    $groupUsers = Get-SPOSiteGroup -Site $site.Url -Group $group.Title -ErrorAction Stop | 
                        Select-Object -ExpandProperty Users
                    
                    if ($groupUsers -like "*$UserEmail*") {
                        $results += [PSCustomObject]@{
                            SiteUrl = $site.Url
                            SiteTitle = $site.Title
                            Template = $site.Template
                            PermissionType = "Group"
                            PermissionLevel = $group.Roles -join ", "
                            GrantedThrough = "Group: $($group.Title)"
                        }
                    }
                }
            } catch {
                Write-Host "Error checking site $($site.Url): $_" -ForegroundColor Red
                continue
            }
        }

        # Display and export results
        if ($results.Count -eq 0) {
            Write-Host "`nNo site permissions found for user: $UserEmail" -ForegroundColor Yellow
        } else {
            Write-Host "`nFound permissions in $($results.Count) sites for user: $UserEmail" -ForegroundColor Green
            $results | Format-Table -AutoSize -Wrap
        }

        if ($OutputFile -and $results.Count -gt 0) {
            $results | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
            Write-Host "Permissions exported to '$OutputFile'." -ForegroundColor Green
        }

        return $results
    } catch {
        Write-Host "Error retrieving site permissions: $_" -ForegroundColor Red
    }
}

# Main script
Clear-Host
Write-Host "=== SharePoint Online User Site Permissions Report ===" -ForegroundColor Cyan

# Verify existing connection
if (-not (Test-SPOConnection)) {
    return
}

# Prompt for user
$userEmail = Read-Host "Enter the email address of the user to check permissions for"

$exportChoice = Read-Host "Do you want to export the results to a CSV file? (Y/N)"
if ($exportChoice -eq 'Y' -or $exportChoice -eq 'y') {
    $defaultPath = "C:\UserSharePointPermissions.csv"
    Write-Host "Default export path is: $defaultPath"
    $customPath = Read-Host "Press Enter to use default path or type a custom path"
    $outputFile = if ($customPath) { $customPath } else { $defaultPath }
} else {
    $outputFile = $null
}

Write-Host "`nRetrieving user's SharePoint site permissions..." -ForegroundColor Cyan
Get-UserSPOSitePermissions -UserEmail $userEmail -OutputFile $outputFile