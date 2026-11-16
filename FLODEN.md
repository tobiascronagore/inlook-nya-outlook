# Microsoft To Do i Inlook (moln) via Power Automate

Tillägget har ingen behörighet mot To Do (kräver admin). I stället pratar det med två flöden
som du äger, via filer i din OneDrive-mapp `Inlook`:

```
Inlook-panelen --(kommandofil)--> OneDrive/Inlook/inkorg --> Flöde 1 --> Microsoft To Do
                                                               |--> OneDrive/Inlook/kvitton/<id>.json
Microsoft To Do --> Flöde 2 (var 5:e min) --> OneDrive/Inlook/todo.json --> Inlook-panelen
```

Bara standardkopplingar (OneDrive för företag, Microsoft To Do (företag)) och inbyggda åtgärder.
Inget Premium, ingen HTTP. Fördröjning: några minuter (filutlösare och Recurrence pollar).

Importerbart paket (.zip) finns inte: det går inte att göra tillförlitligt utan att kunna
testa import mot din miljö (kopplings-id och åtgärds-id är miljöspecifika). Följ stegen nedan.

## 0. Förberedelser (en gång)

Skapa i OneDrive/Inlook: mapparna `inkorg` och `kvitton`, samt en textfil `todo.json` med
innehållet `{}` (flöde 2 skriver över den; den måste finnas från början).
Ta bort gamla testfiler ur `inkorg`.

## Kontrakt: kommandofil (skrivs av tillägget)

Filnamn `cmd-<id>.json` i `Inlook/inkorg`:

```json
{
  "action": "create",
  "id": "9f1c...32 hextecken, unikt per kommando",
  "listId": "<To Do-listans id från todo.json>",
  "listName": "Jobb",
  "todoId": "<uppgiftens id, bara för complete/reopen/delete>",
  "title": "Rubrik",
  "body": "Öppna mejlet: https://...\n\nFrån mejl: ...",
  "dueDateTime": "2026-11-20T00:00:00Z",
  "importance": "Normal",
  "createdAt": "2026-11-16T10:00:00.000Z"
}
```

`action` är `create`, `complete`, `reopen` eller `delete`. Fält som inte hör till åtgärden är tomma eller saknas.
Kvitto `kvitton/<id>.json`: `{"id":"...","action":"...","ok":true,"at":"..."}`. Tillägget visar
"Väntar på synk" tills kvittot finns OCH `todo.json` skrivits efter kvittot.

`todo.json` (skrivs av flöde 2):

```json
{ "generatedAt": "2026-11-16T10:05:00Z",
  "lists": [ { "id": "...", "name": "Jobb" } ],
  "tasks": [ { "id": "...", "listId": "...", "title": "...", "status": "notStarted|completed",
               "dueDateTime": "2026-11-20T00:00:00Z", "importance": "normal",
               "body": "Öppna mejlet: https://...", "createdAt": "..." } ] }
```

## Flöde 1: "Inlook till To Do" (bygg om det befintliga)

Öppna flödet, Redigera. Ta bort den gamla åtgärden "Lägg till en att göra-uppgift (V3)" med rubriken Test.

1. **Utlösare: När en fil skapas** (OneDrive för företag). Mapp: `/Inlook/inkorg`. Infoga undermappar: Nej.
   Tre punkter på utlösaren, Inställningar, **Samtidighetskontroll: På, Grad 1** (ett kommando i taget).
2. **Villkor** (inbyggd, "Kontroll"): `Filnamn` (dynamiskt innehåll från utlösaren) **slutar med** `.json`.
   Allt nedan läggs under **Om ja**. (Ger dig frihet att lägga andra filer i mappen.)
3. **Hämta filinnehåll** (OneDrive för företag). Fil: dynamiskt innehåll `Fil-identifierare` från utlösaren.
4. **Skapa** (Compose, "Datakörning"), byt namn till `Kommandotext`. Klistra in under fliken Uttryck:

   ```
   if(empty(body('Hämta_filinnehåll')?['$content']), string(body('Hämta_filinnehåll')), base64ToString(body('Hämta_filinnehåll')?['$content']))
   ```

   (Hanterar både att innehållet kommer som JSON-objekt och som base64.)
5. **Parsa JSON** (Datakörning). Innehåll: utdata från `Kommandotext`. Schema, klistra in (eller "Generera från exempel"):

   ```json
   {
     "type": "object",
     "properties": {
       "action": { "type": "string" },
       "id": { "type": "string" },
       "listId": { "type": "string" },
       "listName": { "type": "string" },
       "todoId": { "type": "string" },
       "title": { "type": "string" },
       "body": { "type": "string" },
       "dueDateTime": { "type": "string" },
       "importance": { "type": "string" },
       "createdAt": { "type": "string" }
     }
   }
   ```
6. **Växla** (Kontroll). Aktiv: dynamiskt innehåll `action` från Parsa JSON. Fyra fall:

   - Fall `create`: **Lägg till en att göra-uppgift (V3)**
     - Lista: ange egen post (tre punkter / "Ange anpassat värde"), välj `listId` från Parsa JSON
     - Rubrik: `title`
     - Brödtext: `body`
     - Förfallodatum: uttryck `if(empty(body('Parsa_JSON')?['dueDateTime']), null, body('Parsa_JSON')?['dueDateTime'])`
     - Prioritet: Normal (rullgardin; vill du styra via `importance`, använd uttrycket `body('Parsa_JSON')?['importance']`, värdena är Low/Normal/High)
     - Status: Ej påbörjad
   - Fall `complete`: **Uppdatera att göra-uppgift (V2)**. Lista: `listId`. Uppgift (id): `todoId`. Status: Slutförd (välj i rullgardinen). Lämna övriga fält tomma.
   - Fall `reopen`: **Uppdatera att göra-uppgift (V2)**. Lista: `listId`. Uppgift (id): `todoId`. Status: Ej påbörjad.
   - Fall `delete`: **Ta bort att göra-uppgift (V2)**. Lista: `listId`. Uppgift (id): `todoId`.
   - Standard: tom.
7. Efter Växla (samma nivå): **Skapa fil** (OneDrive för företag) = kvitto.
   Mappsökväg `/Inlook/kvitton`. Filnamn: `@{body('Parsa_JSON')?['id']}.json`. Filinnehåll:

   ```
   {"id":"@{body('Parsa_JSON')?['id']}","action":"@{body('Parsa_JSON')?['action']}","ok":true,"at":"@{utcNow()}"}
   ```
8. Sist: **Ta bort fil** (OneDrive för företag). Fil: `Fil-identifierare` från utlösaren.
   Kommandofilen är då "flyttad ut" och kan inte köras två gånger.

Misslyckas något steg (t.ex. uppgiften finns inte längre) stannar kommandofilen i `inkorg` och kvitto uteblir:
tillägget visar då "Väntar på synk (ovanligt länge, kontrollera flödet)" efter 15 minuter. Se körhistoriken, ta bort filen.

## Flöde 2: "To Do till Inlook"

Nytt flöde, **Schemalagt molnflöde** (utlösaren Recurrence). Namn: "To Do till Inlook".

1. **Recurrence**: Intervall 5, Frekvens Minut.
2. **Initiera variabel** `startad`, typ Sträng, värde uttryck `utcNow()` (tidsstämpel FÖRE läsningen).
3. **Initiera variabel** `alla`, typ Matris, värde tomt. **Initiera variabel** `tmp`, typ Matris, tomt.
4. **Lista alla To-Do-listor (V2)** (Microsoft To Do (företag)).
5. **Välj** (Datakörning), namn `Välj listor`. Från: `value` från steg 4. Mappa (växla till textläge, klistra in):
   `id` = `item()?['id']`, `name` = `coalesce(item()?['displayName'], item()?['name'])`.
6. **För varje** lista (Från: `value` från steg 4). **Inställningar: Samtidighetskontroll På, grad 1.** Inuti:
   1. **Lista att göra-uppgifter** för vald lista (To Do (företag), åtgärdens namn kan vara "Lista att göra-uppgifter efter mapp (V2)" eller "Hämta att göra-uppgifter"). Lista: `id` för aktuell lista.
   2. **Filtrera matris**: Från = värdena från 6.1. Villkor (avancerat läge):
      ```
      @or(not(equals(item()?['status'], 'completed')), greater(ticks(coalesce(item()?['completedDateTime'], '1970-01-01')), ticks(addDays(utcNow(), -7))))
      ```
      (Öppna uppgifter plus de som blev klara senaste 7 dygnen.)
   3. **Välj**, namn `Välj uppgifter`. Från: utdata från 6.2. Mappa:
      `id` = `item()?['id']`; `listId` = `items('För_varje')?['id']`; `title` = `item()?['title']`;
      `status` = `item()?['status']`; `dueDateTime` = `item()?['dueDateTime']`; `importance` = `item()?['importance']`;
      `body` = `if(greater(length(string(item()?['body'])), 600), substring(string(item()?['body']), 0, 600), string(item()?['body']))`;
      `createdAt` = `item()?['createdDateTime']`.
   4. **Ange variabel** `tmp` = `union(variables('alla'), body('Välj_uppgifter'))`.
   5. **Ange variabel** `alla` = `variables('tmp')`.
7. **Hämta filmetadata med sökväg** (OneDrive för företag): `/Inlook/todo.json`.
8. **Uppdatera fil** (OneDrive för företag). Fil: `Id` från steg 7. Filinnehåll, uttryck:
   ```
   concat('{"generatedAt":"', variables('startad'), '","lists":', string(body('Välj_listor')), ',"tasks":', string(variables('alla')), '}')
   ```

**Namnen på fälten** (`status`, `completedDateTime`, `dueDateTime`, `displayName` m.fl.) kan skilja sig lite mellan
kopplingens versioner. Kör flödet en gång, öppna körningen, klicka på steg 6.1 och kontrollera fältnamnen under Utdata;
justera uttrycken i 5, 6.2 och 6.3 till de namn du ser. Öppna sedan `todo.json` i OneDrive och kontrollera att uppgifterna
och listorna finns med. Tilläggets "status" förväntar sig `completed` för klara uppgifter; annat räknas som öppen.

**Valfritt (rekommenderas): extra synk direkt efter ett kommando.** Öppna flöde 2, "Spara som" (kopia), och ersätt
utlösaren Recurrence med **När en fil skapas** i `/Inlook/kvitton`. Kvittot skapas av flöde 1 när kommandot är klart,
så panelen blir uppdaterad inom någon minut i stället för efter upp till fem. Samtidighet på utlösaren: grad 1.

## Prova

1. Kör flöde 2 manuellt (Testa), kontrollera att `Inlook/todo.json` innehåller dina listor.
2. Öppna Inlook (moln), fliken Uppgifter, Ladda om. Listorna ska synas i listvalet.
3. Välj ett mejl, skapa uppgift. Rad med "Väntar på synk" visas. Inom några minuter finns uppgiften i To Do.
4. Bocka av och ta bort i panelen; kontrollera i To Do efter några minuter.

## Felsökning

- Panelen säger "Ingen To Do-data än": `todo.json` saknas/tom, eller flöde 2 har inte körts. Kontrollera körhistoriken.
- Kommando ligger kvar i `inkorg`: flöde 1 har misslyckats (se körhistorik) eller är avstängt.
- "Skapa fil" i flöde 1 säger att filen finns: kommando-id används två gånger (ska inte hända); ta bort kvittot.
- Gamla kvitton (över 3 dygn) städas av panelen automatiskt.
