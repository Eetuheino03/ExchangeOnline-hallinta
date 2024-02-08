Exchange Online Email Management Script
Tämä PowerShell-skripti tarjoaa työkaluja sähköpostien hallintaan Exchange Online -ympäristössä. Skripti tarjoaa kaksi pääominaisuutta: sähköpostien siivoaminen ja sähköpostien haku.

Ominaisuudet
Sähköpostien siivoaminen: Skripti lukee sähköpostiosoitteet mail.txt-tiedostosta tai pyytää niitä käyttäjältä, jos tiedostoa ei ole. Se poistaa sähköpostit, jotka täyttävät tietyn ehdon (tässä tapauksessa lähettäjän perusteella).

Sähköpostien haku: Skripti hakee sähköpostit, jotka täyttävät käyttäjän määrittämän ehdon.

Käyttö
Asenna ExchangeOnlineManagement-moduuli PowerShellissa komennolla Install-Module -Name ExchangeOnlineManagement.

Kirjaudu sisään Exchange Onlineen komennolla Connect-ExchangeOnline -UserPrincipalName user@example.com -ShowProgress $true. Korvaa user@example.com omalla käyttäjätunnuksellasi.

Suorita skripti PowerShellissa.

Valitse valikosta toiminto (sähköpostien siivoaminen tai haku) ja seuraa ohjeita.

Huomautukset
Search-Mailbox-komento vaatii järjestelmänvalvojan oikeudet.
Search-Mailbox-komento tallentaa hakutulokset kohdekansioon, joten sinun on varmistettava, että sinulla on oikeus tähän kansioon.
Search-Mailbox-komento -DeleteContent-parametrin kanssa poistaa löydetyt sähköpostit. Käytä varoen.
Tämä skripti on yksinkertainen esimerkki ja se voidaan mukauttaa tarpeidesi mukaan. Esimerkiksi, voit lisätä virheenkäsittelyn puuttuvan mail.txt-tiedoston, virheellisten sähköpostiosoitteiden tai oikeuksien virheiden varalta. Voit myös validoida käyttäjän syötteet varmistaaksesi, että ne täyttävät odotetut muodot ja kriteerit.****
