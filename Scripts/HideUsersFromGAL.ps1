function Hide-UsersFromGAL {
    param ([string]$InputFile)
    try {
        if (-Not (Test-Path -Path $InputFile)) {
            Write-Host "Input file '$InputFile' not found!" -ForegroundColor Red
            return
        }

        $users = Get-Content -Path $InputFile
        Write-Host "Processing $($users.Count) users from file." -ForegroundColor Cyan

        $onPremUsers = @()
        $alreadyHidden = @()

        foreach ($user in $users) {
            try {
                $mailbox = Get-Mailbox -Identity $user -ErrorAction Stop

                if ($mailbox.HiddenFromAddressListsEnabled -eq $true) {
                    Write-Host "User '$user' is already hidden from GAL." -ForegroundColor Yellow
                    $alreadyHidden += $user
                    continue
                }

                Set-Mailbox -Identity $user -HiddenFromAddressListsEnabled $true -ErrorAction Stop
                Write-Host "Successfully hid user '$user' from GAL." -ForegroundColor Green
            } catch {
                Write-Host "Error hiding user '$user': $_" -ForegroundColor Red
                if ($_ -match "it's out of the current user's write scope") {
                    $onPremUsers += $user
                }
            }
        }

        if ($onPremUsers.Count -gt 0) {
            $scriptContent = @"
Import-Module ActiveDirectory

`$users = @(
    `"$($onPremUsers -join '", "')`"
)

foreach (`$email in `$users) {
    `$user = Get-ADUser -Filter {EmailAddress -eq `$email}
    if (`$user) {
        Set-ADUser -Identity `$user -Replace @{msExchHideFromAddressLists = `$true}
        Write-Host "Updated user: `$email" -ForegroundColor Green
    } else {
        Write-Host "User not found: `$email" -ForegroundColor Red
    }
}
"@
            Set-Clipboard -Value $scriptContent
            Write-Host "`nOn-premises update script copied to clipboard." -ForegroundColor Yellow
        }

        if ($alreadyHidden.Count -gt 0) {
            Write-Host "`nUsers already hidden from GAL:" -ForegroundColor Cyan
            $alreadyHidden | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }
        }
    } catch {
        Write-Host "Error: $_" -ForegroundColor Red
    }
}

# Main script
Clear-Host
Write-Host "=== Hide Users From GAL Tool ===" -ForegroundColor Cyan
Write-Host "`nConnected to: $($script:ExchangeConnection.OrganizationName)" -ForegroundColor Green
Write-Host "User: $($script:ExchangeConnection.CurrentUser)`n" -ForegroundColor Green

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$defaultInputFile = Join-Path -Path $scriptDir -ChildPath "users.txt"

do {
    Write-Host "`nOptions:" -ForegroundColor Cyan
    Write-Host "1. Hide users using default file (users.txt)"
    Write-Host "2. Specify custom input file"
    Write-Host "3. Return to Main Menu"

    $choice = Read-Host "`nEnter your choice (1-3)"

    switch ($choice) {
        "1" {
            if (Test-Path $defaultInputFile) {
                Hide-UsersFromGAL -InputFile $defaultInputFile
            } else {
                Write-Host "Default file 'users.txt' not found in script directory." -ForegroundColor Red
                Write-Host "Please create the file with one email address per line." -ForegroundColor Yellow
            }
        }
        "2" {
            $customPath = Read-Host "Enter the full path to your input file"
            if ($customPath) {
                Hide-UsersFromGAL -InputFile $customPath
            }
        }
        "3" { return }
    }
} while ($true)