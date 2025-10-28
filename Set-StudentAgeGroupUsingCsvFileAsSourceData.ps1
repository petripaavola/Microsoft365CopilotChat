<#
.SYNOPSIS
    Asettaa oppilaiden ikäryhmän oletuksena "NotAdult" käyttäen CSV-tiedostoa lähdedatana.
    notAdult (tai Adult) arvo tarvitaan, jotta Microsoft 365 Copilot Chat toimii 13+ oppilailla K12-määritetyissä kouluissa.

.DESCRIPTION
    Skripti lukee CSV-tiedoston, jossa on sarake "UserPrincipalName", ja asettaa kyseisten käyttäjien ikäryhmäksi oletuksena "NotAdult" Microsoft Graphin kautta.
    AgeGroup-arvo voidaan määrittää myös parametrina (null, 'Minor', 'NotAdult', 'Adult'), jos halutaan asettaa jokin muu arvo.

    Tarkoitettu käytettäväksi mm. K12-koulun 13+ -vuotiaiden oppilastileille, joissa sähköpostiosoite vastaa käyttäjän UserPrincipalName-arvoa,
    jotta Microsoft 365 Copilot Chat toimii oppilaiden kanssa.

.PARAMETER CsvFilePath
    Polku CSV-tiedostoon, joka sisältää sarakkeen "UserPrincipalName".
    Esimerkki: C:\Polku\oppilaat.csv

.PARAMETER AgeGroup
    Määrittää asetettavan ikäryhmän. Oletusarvo on "NotAdult", jos parametri jätetään määrittämättä.
    Sallitut arvot: 'null', 'Minor', 'NotAdult', 'Adult'

.PARAMETER WhatIf
    Jos määritetty parametri -WhatIf, niin skripti suorittaa testiajon eli kertoisi mitä se tekisi, mutta ei tee varsinaisia muutoksia käyttäjiin.
    Huom! Tämä testiajo ei tarkista onnistuisiko päivitys oikeasti, vaan simuloi vain toiminnon, että mitä tehtäisiin ja käy mm. läpi kaikki käyttäjät csv-tiedostosta.

.INPUTS
    System.String (CSV-tiedoston polku)

.OUTPUTS
    Luo lokitiedoston skriptihakemistoon muodossa yyyyMMdd-HHmmss-Set-StudentAgeGroupUsingCsvFileAsSourceData.log
    ja kirjaa etenemisen ja mahdolliset virheet myös konsoliin.

    Luo erillisen lokitiedoston epäonnistuneista päivityksistä muodossa yyyyMMdd-HHmmss-Failed-User-Updates.log

.EXAMPLE
    PS> .\Set-StudentAgeGroupUsingCsvFileAsSourceData.ps1 -CsvFilePath "C:\Polku\oppilaat.csv"

    Suorittaa skriptin käyttäen oppilaat.csv-tiedostoa, jossa on "UserPrincipalName"-sarakkeessa oppilaiden UPN:t (eli yleensä sähköpostiosoitteet).
    AgeGroup asetetaan oletusarvoisesti "NotAdult", koska -AgeGroup -parametria ei määritetty.

.EXAMPLE
    PS> .\Set-StudentAgeGroupUsingCsvFileAsSourceData.ps1 -CsvFilePath "C:\Polku\oppilaat.csv" -AgeGroup "NotAdult" -WhatIf

    Suorittaa skriptin käyttäen oppilaat.csv-tiedostoa, jossa on "UserPrincipalName"-sarakkeessa oppilaiden UPN:t. AgeGroup asetetaan "NotAdult".
    Koska -WhatIf -parametri on määritetty, skripti suorittaa testiajon eikä tee varsinaisia muutoksia käyttäjiin.

.EXAMPLE
    PS> .\Set-StudentAgeGroupUsingCsvFileAsSourceData.ps1 -CsvFilePath "C:\Polku\opiskelijat.csv" -AgeGroup "Adult"

    Suorittaa skriptin käyttäen opiskelijat.csv-tiedostoa, jossa on "UserPrincipalName"-sarakkeessa oppilaiden UPN:t. AgeGroup asetetaan "Adult".

.EXAMPLE
    CSV-esimerkki:

    UserPrincipalName
    oppilas1@koulu.fi
    oppilas2@koulu.fi
    oppilas3@koulu.fi

.NOTES
    Vaatimukset:
    - PowerShell-moduulit:
        - Microsoft.Graph.Users
        - Microsoft.Graph.Authentication

    Moduulien asennus (ei tarvita Administrator-oikeuksia):
        Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force


.AUTHOR
    Petri Paavola (Petri.Paavola@yodamiitti.fi)
    Senior Modern Management Principal
    Microsoft MVP - Windows and Intune
    Yodamiitti Oy
    28.10.2025

.LINK
    https://github.com/petripaavola/Microsoft365CopilotChat
#>


[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$csvFilePath,
    [ValidateSet('null','Minor','NotAdult','Adult')]
    [string]$AgeGroup = "NotAdult",
    [switch]$WhatIf
)


# Huge Try-Catch block to capture any unexpected errors in whole script
Try {

    $Version =  '1.0'

    # Get script start time for logging
    $scriptStartTime = Get-Date -Format "yyyyMMdd-HHmmss"

    # Start Transcript for logging
    $logFilePath = "$PSScriptRoot\$($scriptStartTime)-Set-StudentAgeGroupUsingCsvFileAsSourceData.log"
    Start-Transcript -Path $logFilePath -Append

    # Logfile path for failed user updates
    $failedUpdatesLogFilePath = "$PSScriptRoot\$($scriptStartTime)-Failed-User-Updates.log"

    Write-Host "Aloitetaan oppilaiden ikäryhmän asettaminen CSV-tiedoston perusteella.`nSkriptiversio $Version" -ForegroundColor Cyan
    Write-Host ""

    # Test csv-file exists
    if (-Not (Test-Path -Path $csvFilePath)) {
        Write-Host "CSV-tiedostoa ei löydy polusta: $csvFilePath" -ForegroundColor Red
        Stop-Transcript
        exit 1
    }

    # Read CSV-file
    $students = Import-Csv -Path $csvFilePath -Encoding UTF8

    # Check CSV-file read was successful
    if ((-not $students) -or $students.Count -eq 0) {
        Write-Host "CSV-tiedostoa ei voida lukea tai tiedosto on tyhjä." -ForegroundColor Red
        Stop-Transcript
        exit 1
    }

    # Check UserPrincipalName column exists
    if (-not ($students[0].PSObject.Properties.Name -contains "UserPrincipalName")) {
        Write-Host "Saraketta 'UserPrincipalName' ei löydy." -ForegroundColor Red
        Stop-Transcript
        exit 1
    }

    $totalStudents = $students.Count
    Write-Host "Löydettiin $totalStudents oppilasta CSV-tiedostosta" -ForegroundColor White
    Write-Host ""

    Write-Host "Ladataan tarvittavat PowerShell-moduulit..." -ForegroundColor White

    # Import Microsoft.Graph.Users module
    Write-Host "`tMicrosoft.Graph.Users -moduuli: " -NoNewline -ForegroundColor White
    Import-Module Microsoft.Graph.Users
    $Success = $?

    # Check module import was successful
    if ($Success) {
        Write-Host "OK" -ForegroundColor Green
    } else {
        Write-Host "Failed" -ForegroundColor Red
        Write-Host "Microsoft.Graph.Users -moduulia ei voida ladata. Varmista, että moduuli on asennettu." -ForegroundColor Red
        Write-Host "Asenna moduuli komennolla:`nInstall-Module Microsoft.Graph.Users -Scope CurrentUser -Force" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Skriptin suoritus päättyy." -ForegroundColor Cyan
        Stop-Transcript
        exit 1
    }

    # Import Microsoft.Graph.Authentication module
    Write-Host "`tMicrosoft.Graph.Authentication -moduuli: " -NoNewline -ForegroundColor White
    Import-Module Microsoft.Graph.Authentication
    $Success = $?

    # Check module import was successful
    if ($Success) {
        Write-Host "OK" -ForegroundColor Green
    } else {
        Write-Host "Failed" -ForegroundColor Red
        Write-Host "Microsoft.Graph.Authentication -moduulia ei voida ladata. Varmista, että moduuli on asennettu." -ForegroundColor Red
        Write-Host "Asenna moduuli komennolla:`nInstall-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Skriptin suoritus päättyy." -ForegroundColor Cyan
        Stop-Transcript
        exit 1
    }

    Write-Host "PowerShell-moduulit ladattu onnistuneesti.`n" -ForegroundColor Green

    # Connect to Microsoft Graph
    Write-Host "Yhdistetään Microsoft Graphiin..." -ForegroundColor White
    Connect-MgGraph -Scopes "User.ReadWrite.All"
    $Success = $?

    # Test if connection was successful
    if ($Success) {
        # Get MgGraph connection status
        $connectionStatus = Get-MgContext
        if($connectionStatus) {
            Write-Host "Yhteys Microsoft Graphiin muodostettu onnistuneesti." -ForegroundColor Green
            Write-Host "Yhteyden tiedot:" -ForegroundColor White
            Write-Host "  Käyttäjä: $($connectionStatus.Account)" -ForegroundColor White
            Write-Host "  Organisaatio: $($connectionStatus.TenantId)" -ForegroundColor White
            Write-Host ""
        } else {
            Write-Host "Yhteyttä Microsoft Graphiin ei voida muodostaa. Varmista, että sinulla on tarvittavat käyttöoikeudet." -ForegroundColor Red
            Write-Host ""
            Write-Host "Skriptin suoritus päättyy." -ForegroundColor Cyan
            Disconnect-MgGraph -ErrorAction SilentlyContinue

            Stop-Transcript
            exit 1
        }
    } else {
        Write-Host "Yhteyttä Microsoft Graphiin ei voida muodostaa. Varmista, että sinulla on tarvittavat käyttöoikeudet." -ForegroundColor Red
        Write-Host ""
        Write-Host "Skriptin suoritus päättyy." -ForegroundColor Cyan

        Stop-Transcript
        exit 1
    }

    # Confirm action if -WhatIf is not specified
    if (-not $WhatIf) {
        $confirmation = Read-Host "Haluatko varmasti asettaa ikäryhmän '$AgeGroup' kaikille $totalStudents oppilaalle? (K/E)"
        if ($confirmation -ne "K") {
            Write-Host "Toiminto peruutettu käyttäjän toimesta." -ForegroundColor Yellow
            Write-Host ""
            Write-Host "Skriptin suoritus päättyy." -ForegroundColor Cyan
            Disconnect-MgGraph
            Stop-Transcript
            exit 0
        }
    } else {
        Write-Host "-WhatIf-parametri määritetty. Suoritetaan testiajo $totalStudents oppilaalle ilman muutoksia." -ForegroundColor Yellow
        Write-Host "Määritetty AgeGroup: $AgeGroup"
        Write-Host ""
    }


    # Loop through each student and set AgeGroup
    $i = 1
    $successCount = 0
    $failureCount = 0
    foreach ($student in $students) {
        $userId = $student.UserPrincipalName
        if($userId) {
            if($WhatIf) {
                # WhatIf mode specified so just simulate the action what would happen

                Write-Host "[WhatIf] ($i/$totalStudents) Asetettaisiin ikäryhmä '$AgeGroup' käyttäjälle: $userId" -ForegroundColor Cyan
                Update-MgUser -UserId "$userId" -AgeGroup "$AgeGroup" -WhatIf
                $successCount++
            } else {
                # Normal PRODUCTION mode - perform the action

                Write-Host "($i/$totalStudents) Asetetaan ikäryhmä '$AgeGroup' käyttäjälle: $userId" -ForegroundColor Cyan
                Update-MgUser -UserId "$userId" -AgeGroup "$AgeGroup"
                $Success = $?
                if ($Success) {
                    Write-Host "Ikäryhmän '$AgeGroup' asettaminen onnistui käyttäjälle: $userId" -ForegroundColor Green
                    $successCount++
                } else {
                    Write-Host "Ikäryhmän '$AgeGroup' asettaminen epäonnistui käyttäjälle: $userId" -ForegroundColor Red

                    # Log failed update to separate log file
                    $errorMessage = "Ikäryhmän '$AgeGroup' asettaminen epäonnistui käyttäjälle: $userId"
                    $errorMessage | Out-File -FilePath $failedUpdatesLogFilePath -Append -Encoding UTF8

                    $failureCount++
                }
            }

        } else {
            Write-Host "UserPrincipalName puuttuu riviltä/käyttäjältä $($student -join ', '). Ohitetaan..." -ForegroundColor Yellow
            $failureCount++
            continue
        }

        $i++
    } # End Foreach-loop


    Write-Host ""
    Write-Host "Skriptin suoritus valmis." -ForegroundColor Green
    Write-Host ""
    Write-Host "Onnistuneet päivitykset: $successCount" -ForegroundColor Green

    if($WhatIf) {
        # WhatIf mode specified
        Write-Host "Epäonnistuneet päivitykset: 0 (WhatIf-tilassa ei tehty varsinaisia päivityksiä)" -ForegroundColor Green
    } else {
        # Normal PRODUCTION mode
        if($failureCount -gt 0) {
            Write-Host "Epäonnistuneet päivitykset: $failureCount" -ForegroundColor Red
            Write-Host "Katso epäonnistuneiden päivitysten lokitiedosto: $failedUpdatesLogFilePath" -ForegroundColor Yellow
        } else {
            Write-Host "Epäonnistuneet päivitykset: 0" -ForegroundColor Green
        }
    }

    Write-Host ""
    Write-Host "Lokitiedosto tallennettu polkuun: $($logFilePath)" -ForegroundColor Yellow
    Write-Host ""

    # Disconnect from Microsoft Graph
    Disconnect-MgGraph

    # Stop Transcript for logging
    Write-Host "Skriptin suoritus päättyy." -ForegroundColor Cyan
    Stop-Transcript
    Exit 0

} catch {
    # Catch any unexpected errors in whole script
    Write-Host "Tapahtui odottamaton virhe: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Skriptin suoritus keskeytyi."
    Write-Host "Lokitiedosto tallennettu polkuun: $($logFilePath)" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Skriptin suoritus päättyy." -ForegroundColor Cyan

    # Stop Transcript for logging
    Stop-Transcript

    Exit 1
}
