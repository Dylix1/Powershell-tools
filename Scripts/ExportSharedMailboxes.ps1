function Export-SharedMailboxes {
    <#
    .SYNOPSIS
    Exporteert alle gedeelde postvakken naar een CSV-bestand.
    
    .DESCRIPTION
    Deze functie exporteert alle gedeelde postvakken met hun eigenschappen naar een CSV-bestand.
    De machtigingen (VolledigeToegangMachtigingen, VerzendenAlsMachtigingen, VerzendenNamens) worden 
    correct opgehaald en geëxporteerd.
    
    .PARAMETER OutputFile
    Het pad waar het CSV-bestand moet worden opgeslagen.
    
    .PARAMETER IncludePermissions
    Voegt machtigingsinformatie toe aan de export.
    
    .PARAMETER IncludeForwarding
    Voegt doorstuurinformatie toe aan de export.
    
    .EXAMPLE
    Export-SharedMailboxes -OutputFile "C:\GedeeldePostvakken.csv" -IncludePermissions -IncludeForwarding
    #>
    param (
        [string]$OutputFile,
        [switch]$IncludePermissions,
        [switch]$IncludeForwarding
    )
    try {
        # Retrieve all shared mailboxes
        Write-Host "Gedeelde postvakken ophalen..." -ForegroundColor Yellow
        $sharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited
        
        Write-Host "$($sharedMailboxes.Count) gedeelde postvakken gevonden." -ForegroundColor Green
        
        if ($sharedMailboxes.Count -eq 0) {
            Write-Host "Geen gedeelde postvakken gevonden. Afsluiten." -ForegroundColor Yellow
            return
        }
        
        # Create results array
        $results = @()
        
        # Process each mailbox
        $counter = 0
        foreach ($mailbox in $sharedMailboxes) {
            $counter++
            Write-Progress -Activity "Gedeelde Postvakken Verwerken" -Status "Postvak verwerken: $($mailbox.DisplayName)" -PercentComplete (($counter / $sharedMailboxes.Count) * 100)
            
            # Create mailbox info object with Dutch column headers
            $mailboxInfo = [PSCustomObject]@{
                Weergavenaam = $mailbox.DisplayName
                PrimairEmailadres = $mailbox.PrimarySmtpAddress
                AlleEmailadressen = ($mailbox.EmailAddresses | Where-Object { $_ -like "smtp:*" } | ForEach-Object { $_.ToString().Replace("smtp:", "") }) -join "; "
                AangemaaktOp = $mailbox.WhenCreated
                LaatstAangemeld = $null
                IsDirSynced = $mailbox.IsDirSynced
                DoorgestuurdNaar = $mailbox.ForwardingAddress
                DoorgestuurdEmailadres = $mailbox.ForwardingSmtpAddress
                VerzendenNamens = ($mailbox.GrantSendOnBehalfTo -join "; ")
                VolledigeToegangMachtigingen = ""
                VerzendenAlsMachtigingen = ""
                VerborgenVoorGAL = $mailbox.HiddenFromAddressListsEnabled
            }
            
            # Get last logon time
            try {
                $stats = Get-MailboxStatistics -Identity $mailbox.Identity -ErrorAction SilentlyContinue
                if ($stats) {
                    $mailboxInfo.LaatstAangemeld = $stats.LastLogonTime
                }
            } catch {
                # Continue if we can't get statistics
            }
            
            # Include permissions if requested
            if ($IncludePermissions) {
                try {
                    # Get FullAccess permissions
                    $fullAccess = Get-MailboxPermission -Identity $mailbox.Identity | 
                        Where-Object { $_.User -notlike "NT AUTHORITY\*" -and $_.User -notlike "S-1-5*" -and $_.IsInherited -eq $false }
                    
                    $mailboxInfo.VolledigeToegangMachtigingen = ($fullAccess | ForEach-Object { $_.User }) -join "; "
                    
                    # Get SendAs permissions
                    $sendAs = Get-RecipientPermission -Identity $mailbox.Identity | 
                        Where-Object { $_.Trustee -notlike "NT AUTHORITY\*" -and $_.Trustee -notlike "S-1-5*" }
                    
                    $mailboxInfo.VerzendenAlsMachtigingen = ($sendAs | ForEach-Object { $_.Trustee }) -join "; "
                } catch {
                    Write-Host "Kan machtigingen voor $($mailbox.DisplayName) niet ophalen: $_" -ForegroundColor Yellow
                }
            }
            
            # Get forwarding info if requested
            if ($IncludeForwarding) {
                # Forwarding information is already included in the base mailbox object
            }
            
            $results += $mailboxInfo
        }
        
        # Clear the progress bar
        Write-Progress -Activity "Gedeelde Postvakken Verwerken" -Completed
        
        # Export results to CSV
        if ($OutputFile) {
            $results | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
            Write-Host "Geëxporteerd: $($results.Count) gedeelde postvakken naar '$OutputFile'." -ForegroundColor Green
        }
        
        # Return results to console if not too many
        if ($results.Count -le 10) {
            Write-Host "`nSamenvatting Gedeelde Postvakken:" -ForegroundColor Cyan
            $results | Format-Table Weergavenaam, PrimairEmailadres, LaatstAangemeld -AutoSize
        } else {
            Write-Host "`nTe veel postvakken om in de console weer te geven. Controleer het geëxporteerde CSV-bestand." -ForegroundColor Yellow
        }
        
        return $results
    } catch {
        Write-Host "Fout bij ophalen van gedeelde postvakken: $_" -ForegroundColor Red
    }
}

# Main script
Clear-Host
Write-Host "=== Gedeelde Postvakken Export Tool ===" -ForegroundColor Cyan
Write-Host "`nVerbonden met: $($script:ExchangeConnection.OrganizationName)" -ForegroundColor Green
Write-Host "Gebruiker: $($script:ExchangeConnection.CurrentUser)`n" -ForegroundColor Green

do {
    Write-Host "`nOpties:" -ForegroundColor Cyan
    Write-Host "1. Exporteer Basis Gedeelde Postvak Informatie"
    Write-Host "2. Exporteer Gedetailleerde Gedeelde Postvak Informatie (met Machtigingen)"
    Write-Host "3. Terug naar Hoofdmenu"
    
    $choice = Read-Host "`nVoer uw keuze in (1-3)"
    
    if ($choice -eq '3') { return }
    
    if ($choice -in '1','2') {
        $includePermissions = $choice -eq '2'
        $includeForwarding = $choice -eq '2'
        
        # Default file path
        $defaultPath = "C:\GedeeldePostvakken_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        Write-Host "Standaard exportpad is: $defaultPath"
        $customPath = Read-Host "Druk op Enter om het standaardpad te gebruiken of typ een aangepast pad"
        $outputFile = if ($customPath) { $customPath } else { $defaultPath }
        
        Write-Host "`nGedeelde postvakken ophalen. Dit kan even duren..." -ForegroundColor Cyan
        
        Export-SharedMailboxes -OutputFile $outputFile -IncludePermissions:$includePermissions -IncludeForwarding:$includeForwarding
    }
} while ($true)