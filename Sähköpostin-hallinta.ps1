# Asenna ExchangeOnlineManagement-moduuli
Install-Module -Name ExchangeOnlineManagement

# Kirjaudu sisään Exchange Onlineen
Connect-ExchangeOnline -UserPrincipalName user@example.com -ShowProgress $true

# Lue sähköpostiosoitteet tiedostosta tai pyydä käyttäjältä
if (Test-Path -Path .\mail.txt) {
    $emailAddresses = Get-Content -Path .\mail.txt
} else {
    $emailAddresses = Read-Host -Prompt "Enter the email address"
}

# Luo valikko
do {
    Write-Host "1. Email cleanup"
    Write-Host "2. Search emails"
    Write-Host "3. Exit"
    $input = Read-Host "Choose an option"

    switch ($input) {
        "1" {
            foreach ($emailAddress in $emailAddresses) {
                # Siivoa sähköpostit
                $searchResults = Search-Mailbox -Identity $emailAddress -SearchQuery 'From:"sender@example.com"' -DeleteContent
                Write-Host "Deleted $($searchResults.ResultItemsCount) emails from $emailAddress"
            }
        }
        "2" {
            $query = Read-Host -Prompt "Enter the search query"
            foreach ($emailAddress in $emailAddresses) {
                # Hae sähköpostit
                $searchResults = Search-Mailbox -Identity $emailAddress -SearchQuery $query -TargetMailbox "discoverysearchmailbox" -TargetFolder "SearchResults" -LogLevel Full
                Write-Host "Found $($searchResults.ResultItemsCount) emails in $emailAddress"
            }
        }
        "3" {
            break
        }
    }
} while ($true)

# Kirjaudu ulos Exchange Online -ympäristöstä
Disconnect-ExchangeOnline