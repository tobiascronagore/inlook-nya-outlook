# Inlook (moln) - installation i nya Outlook

Tillägget är en statisk webbplats (GitHub Pages) som loggar in med ditt eget
Microsoft-konto och pratar direkt med Microsoft Graph. Ingen lokal server, ingen
COM, ingen IT-inblandning. Det kan finnas parallellt med det lokala tillägget
(annat manifest-id och annat namn: "Inlook (moln)").

Adress: https://tobiascronagore.github.io/inlook-nya-outlook/

Data: anteckningar (`notes.json`) och dagbok
(`dagbok.json`) som JSON i OneDrive-mappen `Inlook`. Ingenting ligger i det här repot.

Behörigheter (provade i Skandglas tenant): User.Read, Files.ReadWrite, Notes.Read/ReadWrite
fungerar utan admin. Microsoft To Do (Tasks.ReadWrite), Mail.* och Calendars.* kräver
administratör och används därför INTE. Kalender i tillägget utgår tills vidare av samma skäl;
mejlet läses bara via Office.js. Inloggningen sker via en Office-dialog (NAA ger bridge-fel här).

## Steg 1 - Lägg till SPA-plattform i appregistreringen (du, i Entra-portalen)

1. Gå till https://entra.microsoft.com och logga in med jobbkontot.
2. Identitet > Program > **Appregistreringar** > fliken **Alla program** (eller
   "Ägda program"). Öppna appen vars klient-ID står i `js/config.js` (samma
   registrering som OneNote-inloggningen i Inlook använder).
3. **Autentisering** > **Lägg till en plattform** > **Ensidesprogram (SPA)**.
4. Lägg till dessa tre omdirigerings-URI:er (en i taget, spara):
   - `brk-multihub://tobiascronagore.github.io`
   - `https://tobiascronagore.github.io/inlook-nya-outlook/`
   - `https://tobiascronagore.github.io/inlook-nya-outlook/dialog.html` (reservinloggning via Office-dialog, används om NAA inte fungerar)
5. Spara. Ändra inget annat. (Befintliga inställningar för OneNote lämnas orörda.)

Kan du inte öppna appen eller inte spara (du är inte ägare)? Då får du skapa en
egen: **Ny registrering**, namn "Inlook moln", kontotyp "endast den här
katalogen", plattform SPA med URI:erna ovan. Kopiera dess "Program-ID (klient)"
till `js/config.js` (clientId) och be mig publicera om. Om "Ny registrering" är
spärrad för dig är det en gräns jag inte ska gå runt; säg till så tar vi det.

## Steg 2 - Sidladda tillägget (du, i nya Outlook)

1. Öppna https://aka.ms/olksideload i webbläsaren (eller i Outlook: Hämta
   tillägg > **Mina tillägg** > **Anpassade tillägg**).
2. **Lägg till ett anpassat tillägg** > **Lägg till från fil** och välj manifest.xml (ladda ner den från
   `https://tobiascronagore.github.io/inlook-nya-outlook/manifest.xml`; "Lägg till från URL" är spärrat i Skandglas tenant).
3. Bekräfta varningen. Öppna ett mejl: knappen **Inlook (moln)** finns i
   menyfliken (under `...` om den är dold). Klicka, fäst panelen med nålen.
4. Om sidladdning är blockerad ("Lägg till anpassat tillägg" saknas eller
   ger fel): skriv ner exakt felmeddelandet och skicka det tillbaka.

## Steg 3 - Förprovet och första inloggningen (du)

1. I panelen: fliken **Inställningar** > **Förprov: behörigheter och inloggning**.
2. Klicka **Kör alla tester**. För varje behörighet kan ett medgivandefönster
   visas. Klicka **Godkänn** där du får; ser du "Behöver administratörsgodkännande"
   är det ett svar vi vill ha, klicka bara Avbryt/Stäng.
3. Kopiera hela rutan "Rapport att kopiera" och skicka tillbaka. Den innehåller
   ingen token, inga mejl och inga anteckningar.

## Steg 4 - Använda tillägget

- **Mejl**: skapa uppgift (Jobb/Privat, förfallodatum valfritt, länk till mejlet) eller spara en
  snabbanteckning.
- **Uppgifter**: egen lista (filtrera Jobb/Privat, kryssa klart, ta bort).
- **Anteckningar**: lista, ta bort, och engångsimport av `notes.json` /
  `dagbok.json` (finns i `desktop\.data` i den gamla Inlook) via **Välj fil**.

## Felsökning

- "Medgivande saknas eller kräver administratör": den behörigheten kan du inte
  godkänna själv i din tenant. Tillägget fungerar då bara för de delar som inte
  behöver den.
- Panelen tom: kontrollera att https://tobiascronagore.github.io/inlook-nya-outlook/taskpane.html
  öppnas i webbläsaren, och att GitHub Pages hunnit publicera (några minuter).
- Ny version: ladda om panelen; stäng och öppna den om det inte hjälper.

## Uppgifter = riktiga Microsoft To Do (Power Automate)

Fliken Uppgifter använder To Do via två flöden i din Power Automate. Bygg dem enligt `FLODEN.md`
(mapparna `inkorg` och `kvitton` samt filen `todo.json` i OneDrive/Inlook).
