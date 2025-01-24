function Get-DistributionGroupMembers {
    param (
        [string]$GroupName,
        [string]$OutputFile
    )
    try {
        $members = Get-DistributionGroupMember -Identity $GroupName -ErrorAction Stop
        Write-Host "Found $($members.Count) members in the distribution group '$GroupName'." -ForegroundColor Green

        $members | Select-Object Name, PrimarySmtpAddress | Format-Table -AutoSize

        if ($OutputFile) {
            $members | Select-Object Name, PrimarySmtpAddress | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
            Write-Host "Members exported to '$OutputFile'." -ForegroundColor Green
        }
    } catch {
        Write-Host "Error retrieving members: $_" -ForegroundColor Red
    }
}

# Main script
Clear-Host
Write-Host "=== Exchange Online Distribution Group Member Export ===" -ForegroundColor Cyan
Write-Host "`nConnected to: $($script:ExchangeConnection.OrganizationName)" -ForegroundColor Green
Write-Host "User: $($script:ExchangeConnection.CurrentUser)`n" -ForegroundColor Green

$groupName = Read-Host "Enter the distribution group name or alias"

$exportChoice = Read-Host "Do you want to export the results to a CSV file? (Y/N)"
if ($exportChoice -eq 'Y' -or $exportChoice -eq 'y') {
    $defaultPath = "C:\DistributionGroupMembers.csv"
    Write-Host "Default export path is: $defaultPath"
    $customPath = Read-Host "Press Enter to use default path or type a custom path"
    $outputFile = if ($customPath) { $customPath } else { $defaultPath }
} else {
    $outputFile = $null
}

Write-Host "`nRetrieving distribution group members..." -ForegroundColor Cyan
Get-DistributionGroupMembers -GroupName $groupName -OutputFile $outputFile