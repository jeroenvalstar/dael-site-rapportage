# DAEL Site Rapportage

Webapp om op locatie de rapportages voor opdrachtgevers in te vullen:

- **Vodafone** – TSS Report 4.20 (Excel), tabblad *bestaande site* of *nieuwbouw site*
- **KPN** – RCT rapport Nieuwbouw daklocatie (Word)

## Werking

1. Kies voor wie je het rapport invult (en bij Vodafone welk rapport).
2. Beantwoord de vragen per onderdeel. Alles wordt automatisch op het apparaat bewaard (concepten).
3. Bij *Afronden* zie je welke vragen nog leeg zijn en kun je het ingevulde Excel- of Word-bestand mailen of downloaden.
   - Op telefoon/tablet opent **Mail met bijlage** het deelmenu (Mail, Outlook…) met het bestand al als bijlage.
   - Anders downloadt de app het bestand en opent je e-mailprogramma met ontvanger en onderwerp ingevuld.
4. Een (half) ingevuld Excel- of Word-bestand kun je op het startscherm **importeren** om verder in te vullen.
   Het geïmporteerde bestand wordt als basis gebruikt, zodat alles daarin (bijv. extra notities of afbeeldingen) behouden blijft.

De app schrijft de antwoorden rechtstreeks in de originele sjablonen in `templates/`: opmaak, formules,
keuzelijsten en beveiliging blijven intact. Er wordt niets naar een server gestuurd; alles gebeurt in de browser.

## Structuur

| Pad | Inhoud |
| --- | --- |
| `index.html`, `css/app.css` | App en styling |
| `js/app.js` | Schermen, opslaan, importeren, versturen |
| `js/forms.js` | Vragenlijsten en de koppeling van elke vraag naar cel (Excel) of tabelcel/selectievakje (Word) |
| `js/ooxml.js` | Lezen/schrijven van .xlsx en .docx |
| `js/vf-data.js` | Keuzelijsten uit het Vodafone-sjabloon (automatisch gegenereerd) |
| `templates/` | De originele lege formulieren |
| `sw.js` | Offline-cache |

## Nieuwe versie van een formulier

Vervang het bestand in `templates/` (zelfde bestandsnaam) en controleer in `js/forms.js` of de celverwijzingen
(Excel) of tabel/rij/cel-nummers en selectievakje-nummers (Word) nog kloppen. Verhoog daarna `CACHE` in `sw.js`.

## Lokaal draaien

```bash
python3 -m http.server 8080
```

en open http://localhost:8080.
