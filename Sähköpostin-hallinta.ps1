try {
    # Tarkista admin käyttöoikeudet
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if (-not $isAdmin) {
        # Nosta käyttöoikeus adminiksi
        Start-Process powershell.exe -Verb RunAs -ArgumentList "-File $($MyInvocation.MyCommand.Path)"
        exit
    }

    # Asenna ExchangeOnlineManagement-moduuli
    Install-Module -Name ExchangeOnlineManagement -Force
    # Lataa ExchangeOnlineManagement-moduuli 
    Import-Module -Name ExchangeOnlineManagement

    # Kysy sähköpostiosoite käyttäjältä ja käytä sitä kirjautumisessa
    $emailAddress = Read-Host -Prompt "Enter the email address to log in to Exchange Online"
    $credential = Get-Credential -UserName $emailAddress
    Connect-ExchangeOnline -Credential $credential

    # Lue sähköpostiosoitteet tiedostosta tai pyydä käyttäjältä
    if (Test-Path -Path .\mail.txt) {
        $emailAddresses = Get-Content -Path .\mail.txt
    } else {
        $emailAddresses = Read-Host -Prompt "Enter the email address"
    }
    $option
    # Luo valikko
    do {
        Write-Host "1. Email cleanup"
        Write-Host "2. Search emails"
        Write-Host "3. Exit"
        $option = Read-Host "Choose an option"

        switch ($input) {
            "1" {
                foreach ($emailAddress in $emailAddresses) {
                    # Siivoa sähköpostit
                    try {
                        $searchResults = Search-Mailbox -Identity $emailAddress -SearchQuery 'From:"sender@example.com"' -DeleteContent -DeleteType HardDelete
                        Write-Host "Permanently deleted $($searchResults.ResultItemsCount) emails from $emailAddress"
                    } catch {
                        Write-Host "Error: $($_.Exception.Message)"
                    }
                }
            }
            "2" {
                $query = Read-Host -Prompt "Enter the search query"
                $saveOption = Read-Host -Prompt "Save results to (1) Text file or (2) Excel spreadsheet"
                $results = @()

                foreach ($emailAddress in $emailAddresses) {
                    # Hae sähköpostit
                    try {
                        $searchResults = Search-Mailbox -Identity $emailAddress -SearchQuery $query -TargetMailbox "discoverysearchmailbox" -TargetFolder "SearchResults" -LogLevel Full

                        foreach ($result in $searchResults) {
                            $email = [PSCustomObject]@{
                                Sender = $result.Sender
                                Content = $result.Content
                            }
                            $results += $email
                        }

                        Write-Host "Found $($searchResults.ResultItemsCount) emails in $emailAddress"
                    } catch {
                        Write-Host "Error: $($_.Exception.Message)"
                    }
                }

                if ($saveOption -eq "1") {
                    $results | Export-Csv -Path .\search_results.csv -NoTypeInformation
                    Write-Host "Search results saved to search_results.csv"
                } elseif ($saveOption -eq "2") {
                    $results | Export-Csv -Path .\search_results.txt -NoTypeInformation -Delimiter "`t"
                    Write-Host "Search results saved to search_results.txt"
                } else {
                    Write-Host "Invalid save option"
                }
            }
            "3" {
                break
            }
        }
    } while ($true)

    # Kirjaudu ulos Exchange Online -ympäristöstä
    Disconnect-ExchangeOnline
} catch {
    Write-Host "Error: $($_.Exception.Message)"
}
