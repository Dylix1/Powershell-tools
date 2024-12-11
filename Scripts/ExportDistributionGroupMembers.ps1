# Function to connect to Exchange Online
function Connect-ExchangeOnlineSession {
    param (
        [string]$AdminUser
    )
    try {
        # Connect to Exchange Online with specified admin UPN and without showing progress
        Connect-ExchangeOnline -UserPrincipalName $AdminUser -ShowProgress $false
        Write-Host "Successfully connected to Exchange Online as $AdminUser." -ForegroundColor Green
    } catch {
        Write-Host "Error connecting to Exchange Online: $_" -ForegroundColor Red
        exit
    }
}

# Function to retrieve and display distribution group members
function Get-DistributionGroupMembers {
    param (
        [string]$GroupName,
        [string]$OutputFile
    )
    try {
        # Get members of the distribution group
        $members = Get-DistributionGroupMember -Identity $GroupName -ErrorAction Stop
        Write-Host "Found $($members.Count) members in the distribution group '$GroupName'." -ForegroundColor Green

        # Display members in the console
        $members | Select-Object Name, PrimarySmtpAddress | Format-Table -AutoSize

        # Export to CSV if specified
        if ($OutputFile) {
            $members | Select-Object Name, PrimarySmtpAddress | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
            Write-Host "Members exported to '$OutputFile'." -ForegroundColor Green
        }
    } catch {
        Write-Host "Error retrieving members: $_" -ForegroundColor Red
    }
}

# Function to disconnect from Exchange Online
function Disconnect-ExchangeOnlineSession {
    try {
        Disconnect-ExchangeOnline -Confirm:$false
        Write-Host "Disconnected from Exchange Online." -ForegroundColor Green
    } catch {
        Write-Host "Error disconnecting from Exchange Online: $_" -ForegroundColor Red
    }
}

# Get parameters from user
Write-Host "`n=== Exchange Online Distribution Group Member Export ===" -ForegroundColor Cyan
$AdminUPN = Read-Host "`nEnter your Exchange Online admin email address"
$groupName = Read-Host "Enter the distribution group name or alias"

# Ask if user wants to export to CSV
$exportChoice = Read-Host "Do you want to export the results to a CSV file? (Y/N)"
if ($exportChoice -eq 'Y' -or $exportChoice -eq 'y') {
    $defaultPath = "C:\DistributionGroupMembers.csv"
    Write-Host "Default export path is: $defaultPath"
    $customPath = Read-Host "Press Enter to use default path or type a custom path"
    $outputFile = if ($customPath) { $customPath } else { $defaultPath }
} else {
    $outputFile = $null
}

# Main script execution
Write-Host "`nConnecting to Exchange Online..." -ForegroundColor Cyan
Connect-ExchangeOnlineSession -AdminUser $AdminUPN

Write-Host "`nRetrieving distribution group members..." -ForegroundColor Cyan
Get-DistributionGroupMembers -GroupName $groupName -OutputFile $outputFile

Write-Host "`nCleaning up..." -ForegroundColor Cyan
Disconnect-ExchangeOnlineSession

Write-Host "`nScript execution completed." -ForegroundColor Green