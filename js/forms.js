/*
 * Vragenlijsten per formulier. Elk veld verwijst naar de plek in het originele bestand:
 *  - Excel: celverwijzing (cell) of een reeks regels voor langere tekst (cells)
 *  - Word:  tabel/rij/cel/alinea (p), alinea-reeks (rest), selectievakje-nummer (cb) of keuze (opts)
 */
(function () {
  'use strict';

  // ---------- hulpfuncties Excel ----------
  const T = (cell, label, o = {}) => ({ id: cell, cell, label, type: 'text', ...o });
  const D = (cell, label, o = {}) => T(cell, label, { type: 'date', ...o });
  const N = (cell, label, o = {}) => T(cell, label, { inputmode: 'decimal', ...o });
  const rows = (col, from, to) => {
    const out = [];
    for (let r = from; r <= to; r++) out.push(col + r);
    return out;
  };
  const AREA = (col, from, to, label, o = {}) => ({
    id: col + from, cells: rows(col, from, to), label, type: 'area', ...o,
  });
  // Raster: kolommen [{col, label}], rijen [{r, label, only?:[cols]}]
  const GRID = (title, cols, rowDefs, o = {}) => ({
    grid: true, title, hint: o.hint,
    cols: cols.map((c) => c.label),
    rows: rowDefs.map((rd) => ({
      label: rd.label,
      fields: cols.map((c) =>
        rd.only && !rd.only.includes(c.col) ? null : T(c.col + rd.r, `${title} – ${rd.label} – ${c.label}`, { short: c.label })),
    })),
  });

  const BANDS_A = [
    [25, 'Laagband · LTE 700'], [26, 'LTE 800'], [27, 'GSM 900'], [28, 'Middenband · LTE 1400'],
    [29, 'Hoogband · LTE 1800'], [30, 'LTE 2100'], [31, 'LTE 2600'], [32, '5G 3500'],
  ];
  const band = (list, offset) => list.map(([r, l]) => ({ r: r + offset, label: l }));
  const BANDS_B = [
    [0, '5G 700'], [1, 'LTE 800'], [2, 'GSM 900'], [3, 'LTE 1400'],
    [4, 'LTE 1800'], [5, 'LTE 2100'], [6, 'LTE 2600'], [7, '5G 3500'],
  ];
  const bandsFrom = (start, labels = BANDS_B) => labels.map(([i, l]) => ({ r: start + i, label: l }));

  const SECTOR_ROWS = (start) => [0, 1, 2].map((i) => ({ r: start + i, label: `S${i + 1}` }));

  const FOTO_INFO = [
    "Volg de checklist voor de fotodocumentatie en zorg dat alle foto's in een mapje staan.",
    "Selecteer het mapje met de foto's en dit TSS-formulier en maak hier één .zip-bestand van.",
    'Upload dit .zip-bestand in het managementsysteem ter goedkeuring voor Ericsson.',
    "Alle foto's dienen rechtgezet en compleet te zijn (± 250 KB per foto, ± 25 MB per 100 foto's).",
  ];

  // ---------- Vodafone – bestaande site ----------
  const vfBestaand = {
    id: 'vf-bestaand',
    klant: 'vodafone',
    title: 'TSS Report – bestaande site',
    subtitle: 'Vodafone · TSS Report 4.20',
    kind: 'xlsx',
    // Het sjabloon bevat nog voorbeeldgegevens van een andere site: nieuw rapport begint leeg
    startEmpty: true,
    template: 'templates/vodafone-tss-4.20.xlsx',
    sheet: 'TSS-bestaande site Report',
    fileName: (v) => `TSS Report 4.20 Vodafone - bestaande site${v.C3 ? ' - ' + v.C3 : ''}${v.C6 ? ' ' + v.C6 : ''}.xlsx`,
    titleField: 'C3',
    sections: [
      { title: '1a. Site informatie algemeen', items: [
        T('C3', 'Sitenummer'), T('C4', 'Adres'), T('C5', 'Postcode'), T('C6', 'Plaats'),
        T('H3', 'Naam RPE'), D('H4', 'Datum TSS'), N('H5', 'X-coördinaat (uit ARM)'), N('H6', 'Y-coördinaat (uit ARM)'),
        T('C7', 'EMF status'), T('C8', 'Parent-Child config'), T('H8', 'Child sitenummer'),
        T('E10', 'Scope – Antenne – BB/RAN Controller'),
      ] },
      { title: '1b. Aanwezig op de TSS', items: [
        GRID('Aanwezig', [{ col: 'A', label: 'Naam' }, { col: 'E', label: 'Functie' }, { col: 'H', label: 'Bedrijf' }], [
          { r: 16, label: 'Persoon 1' }, { r: 17, label: 'Persoon 2' }, { r: 18, label: 'Persoon 3' },
          { r: 19, label: 'ASP' }, { r: 20, label: 'Landlord (LL)' },
        ], { hint: 'Functie en bedrijf worden bij bekende namen automatisch door Excel ingevuld.' }),
      ] },
      { title: '2a. Huidige configuratie', hint: 'In te vullen door de radioplanner.', items: [
        GRID('Huidige configuratie', [
          { col: 'D', label: 'Kast type' }, { col: 'E', label: 'DU type' }, { col: 'F', label: 'Radio sector 1' },
          { col: 'G', label: 'Radio sector 2' }, { col: 'H', label: 'Radio sector 3' }, { col: 'I', label: 'Tx unit' },
        ], BANDS_A.map(([r, l]) => ({ r, label: l, only: r === 25 ? null : ['D', 'E', 'F', 'G', 'H'] }))),
        GRID('Huidige filters', [
          { col: 'D', label: 'Combiner / Triplexer' }, { col: 'E', label: 'Splitter' }, { col: 'F', label: 'TMA' }, { col: 'G', label: 'RIU' },
        ], bandsFrom(35)),
        T('J35', 'Lead batterijen'), T('J36', 'Li-Ion batterijen', {
          hint: 'Batterijen in kast die blijft: geen actie. Lead in kast die weg moet: kast incl. batterijen afvoeren. Li-Ion in kast die weg moet: informeer Battery project team.',
        }),
        GRID('Antennes A1', [{ col: 'B', label: 'Antenne typenummer' }, { col: 'D', label: 'Azimuth' }, { col: 'E', label: 'ABH' }], SECTOR_ROWS(45)),
        GRID('Antennes A2', [{ col: 'G', label: 'Antenne typenummer' }, { col: 'I', label: 'Azimuth' }, { col: 'J', label: 'ABH' }], SECTOR_ROWS(45)),
        GRID('3 sectoren voor L3500', [{ col: 'B', label: 'Antenne typenummer' }, { col: 'D', label: 'Azimuth' }, { col: 'E', label: 'ABH' }], SECTOR_ROWS(51)),
        T('A56', 'PSU 6302'), N('B56', 'PSU 6302 – aantal'),
        T('E56', 'Rectifiers 6150 – aantal', { hint: 'Bij 3.5GHz altijd 6 (4+2), bij 105MHz altijd 5 (4+1), anders 4 (4+0).' }),
      ] },
      { title: '2b. Gewenste configuratie', hint: 'In te vullen door de radioplanner. Bij Parent-Child config voor beide sites een TSS-rapport maken.', items: [
        GRID('Gewenste configuratie', [
          { col: 'D', label: 'Kast type' }, { col: 'E', label: 'DU type' }, { col: 'F', label: 'Radio sector 1' },
          { col: 'G', label: 'Radio sector 2' }, { col: 'H', label: 'Radio sector 3' }, { col: 'I', label: 'Tx unit' },
        ], [...bandsFrom(62).map((b) => ({ ...b, only: b.r === 62 ? null : ['D', 'E', 'F', 'G', 'H'] })),
          { r: 70, label: 'Mechanical tilt 5G 3500', only: ['F', 'G', 'H'] }]),
        GRID('Gewenste filters', [{ col: 'D', label: 'Splitter' }, { col: 'E', label: 'Filter' }], bandsFrom(73, BANDS_B.slice(0, 7))),
        T('J73', 'Li-Ion batterijen', { hint: 'Altijd plaatsen bij een nieuw enclosure.' }),
        GRID('Nieuwe antennes A1', [{ col: 'B', label: 'Antenne typenummer' }, { col: 'D', label: 'Azimuth' }, { col: 'E', label: 'ABH' }], SECTOR_ROWS(84)),
        GRID('Nieuwe antennes A2 (L3500)', [{ col: 'G', label: 'Antenne typenummer' }, { col: 'I', label: 'Azimuth' }, { col: 'J', label: 'ABH' }], SECTOR_ROWS(84),
          { hint: 'Hoogte altijd bovenzijde gelijk of hoger dan de 2L6H antennes.' }),
        T('A89', 'PSU 6302'), N('B89', 'PSU 6302 – aantal'), T('E89', 'Extra rectifiers 6150 – aantal'),
        AREA('A', 93, 96, 'Opmerkingen'),
      ] },
      { title: '3. Locatie', items: [
        T('H98', 'Mast Vodafone'), T('H99', 'Mast OLO'), T('H100', 'Mast derden'),
      ] },
      { title: '4. Werkzaamheden & toegang', items: [
        AREA('A', 106, 120, '4a. Beschrijving van de werkzaamheden', {
          hint: "Inzet kraan + rijplaten, te verwijderen antennes, RRU's, kasten. Nieuw te plaatsen apparatuur, H&S-aanpassingen.", rows: 8 }),
        AREA('A', 124, 126, '4b. Aanpassingen / opmerkingen toegangsprotocol', {
          hint: 'Zijn alle sleutels aanwezig? Klopt het toegangsprotocol? Zo niet: vermelden, SEN laat het aanpassen op de FOP via enn.site.information@ericsson.com.' }),
      ] },
      { title: '5. Inventarisatie transmissie', items: [
        AREA('A', 132, 133, 'Moet er transmissie worden verplaatst tijdens de werkzaamheden?', { hint: 'Geef aan wat er verplaatst moet worden.' }),
        T('E134', 'Verplaatsen dish(es) door Tx team'), N('I134', 'Aantal dishes'),
        T('E135', 'Verplaatsen MPO connector door ASP'), T('E136', 'Verplaatsing apparatuur door TX team'),
      ] },
      { title: '6. Voeding & aarding', items: [
        AREA('A', 140, 146, '6a. Huidige vorm van de hoofdvoeding en eventuele K2', { hint: 'Bijvoorbeeld eigen K1, landlord voeding, aggregaat, etc.' }),
        AREA('A', 150, 153, '6b. Dient de hoofdvoeding aangepast te worden? Zo ja, hoe?', { hint: 'Denk ook aan verlopen keuringen, defecte kastdeuren of open wartels.' }),
        AREA('A', 155, 158, '6c. Dient de secundaire aansluiting (K2) aangepast te worden? Zo ja, hoe?'),
        AREA('A', 160, 164, '6d. Is de aarding in orde?'),
      ] },
      { title: '7. Bereikbaarheid', items: [
        AREA('A', 167, 177, '7a. Bereikbaarheid locatie, kraan/goederen', { hint: 'Hoe kan logistiek leveren, is er een kraan/hoogwerker nodig en waar is de positie?', rows: 6 }),
        AREA('A', 179, 184, '7b. Zijn er rijplaten nodig? Zo ja, hoeveel en waar?'),
        AREA('A', 186, 188, '7c. Is er extra camerabewaking nodig voor de rijplaten?'),
      ] },
      { title: '8. Risico-inventarisatie (RI&E)', hint: 'Er hoeft geen RI&E aangeleverd te worden; gebruik de checklist om de vragen te beantwoorden.', items: [
        AREA('A', 192, 194, '8a. Zijn alle veiligheidsmiddelen gekeurd, en tot wanneer?'),
        AREA('A', 196, 200, '8b. Is de RI&E tijdens de TSS Groen / Geel / Oranje / Rood? Toelichting bij Oranje / Rood'),
        AREA('A', 202, 207, '8c. Moet de site vóór uitvoering van dit project aangepast worden? Toelichting'),
      ] },
    ],
    info: { title: '9. Fotodocumentatie en versturen TSS', lines: FOTO_INFO },
  };

  // ---------- Vodafone – nieuwbouw site ----------
  const DOCS = ['Constructie-/mastberekeningen', 'Constructieberekeningen', 'Tekenpakket beschikbaar', 'Dakplattegrond bestaande situatie',
    'Schoongrondverklaring/-onderzoek', 'Funderingsadvies', 'Vergunning/gedoogbeschikking', 'Aarding bliksem bestaande situatie', 'Kabelberekeningen elektra'];
  const BLIKSEM = ['Bliksembeveiliging aanwezig', 'Aansluiting op bestaand bliksembeveiligingsnet', 'Aansluiten op bestaand daknet',
    'Bliksembeveiliging gekoppeld aan constructie gebouw', 'Aantal koppelingen aan constructie gebouw', 'Meetrapport beschikbaar',
    'Metingen uit te voeren tijdens aanvang bouw', 'Aantal aanwezige aardelektroden', 'Afstand t.o.v. elkaar', 'Nieuwe zakleidingen aanbrengen',
    'Aantal nieuwe zakleidingen (afhankelijk van meting bouw)', 'Voldoet aan LP3'];
  const ENERGIE = [
    ['F203', 'Eigen aansluiting'], ['F204', 'Aftakking'], ['F205', 'Hoofdbeveiliging (Vodafone)'], ['F206', 'Hoofdschakelaar K1 verdeler / aftakking'],
    ['F207', 'Hoofdschakelaar K2 verdeler'], ['F208', 'Aardlekschakelaar (als hoofdschakelaar, indien van toepassing)'], ['F209', 'Type kabel'],
    ['F210', 'Lengte kabel tussen K1 en RBS/enclosure'], ['F211', 'Aantal reserve groepen, 1-3 fase (1)'], ['F212', 'Aantal reserve groepen, 1-3 fase (2)'],
    ['F213', 'Vereffeningsleiding'], ['F214', 'Verdeler uitbreiden'], ['F215', 'Locatie hoofdverdeler K1'], ['F216', 'Toegang K1 verdeler (wat is de route?)'],
  ];

  const vfNieuwbouw = {
    id: 'vf-nieuwbouw',
    klant: 'vodafone',
    title: 'TSS Report – nieuwbouw site',
    subtitle: 'Vodafone · TSS Report 4.20',
    kind: 'xlsx',
    // Het sjabloon bevat nog voorbeeldgegevens van een andere site: nieuw rapport begint leeg
    startEmpty: true,
    template: 'templates/vodafone-tss-4.20.xlsx',
    sheet: 'TSS-nieuwbouw site Report',
    fileName: (v) => `TSS Report 4.20 Vodafone - nieuwbouw site${v.C3 ? ' - ' + v.C3 : ''}${v.C6 ? ' ' + v.C6 : ''}.xlsx`,
    titleField: 'C3',
    sections: [
      { title: '1a. Site informatie algemeen', items: [
        T('C3', 'Sitenummer'), T('C4', 'Adres'), T('C5', 'Postcode'), T('C6', 'Plaats'),
        T('H3', 'Naam RPE'), D('H4', 'Datum TSS'), N('H5', 'X-coördinaat (uit ARM)'), N('H6', 'Y-coördinaat (uit ARM)'),
        T('C7', 'Parent-Child config'), T('H7', 'Child sitenummer'),
        T('E9', 'Scope – Antenne – BB/RAN Controller'),
      ] },
      { title: '1b. Aanwezig op de TSS', items: [
        GRID('Aanwezig', [{ col: 'A', label: 'Naam' }, { col: 'E', label: 'Functie' }, { col: 'H', label: 'Bedrijf' }], [
          { r: 14, label: 'Persoon 1' }, { r: 15, label: 'Persoon 2' }, { r: 16, label: 'Persoon 3' },
          { r: 17, label: 'Persoon 4' }, { r: 18, label: 'ASP' }, { r: 19, label: 'Landlord (LL)' },
        ], { hint: 'Functie en bedrijf worden bij bekende namen automatisch door Excel ingevuld.' }),
      ] },
      { title: '2. Documentatie', items: [
        ...DOCS.map((l, i) => T('F' + (21 + i), l)),
        AREA('A', 32, 35, 'Opmerkingen'),
      ] },
      { title: '3. Gebouw / terrein', hint: 'Aanvinken om wat het gaat.', items: [
        T('C41', 'Monument'), T('C42', 'Kerk'), T('C43', 'Kantoorgebouw'),
        T('G41', 'Fabriek, silo'), T('G42', 'Woongebouw'), T('G43', 'Nieuwe mast'),
        T('J41', 'Mast OLO'), T('J42', 'Mast derden'),
        AREA('A', 46, 49, 'Opmerkingen'),
      ] },
      { title: '3a. Toegang', items: [
        T('E53', 'Positie sleutelkluis bekend'), T('E54', 'Sleutels nodig?'), T('F54', 'Aantal sleutels'),
        AREA('A', 57, 58, 'Opmerkingen'),
      ] },
      { title: '4. Type site', items: [
        T('A62', 'Indoor – type site'), T('A63', 'Outdoor – type site'), T('D63', 'Indien Special: wat voor object?'),
        AREA('A', 66, 69, 'Opmerkingen (geef aan wat voor ruimte)'),
        AREA('A', 74, 96, '4a. Beschrijving van de werkzaamheden', {
          hint: 'Nieuw te plaatsen antennes, nieuw te plaatsen apparatuur, kabelgoten etc. en eventueel te plaatsen H&S.', rows: 8 }),
      ] },
      { title: '5. Gewenste configuratie', hint: 'In te vullen door de radioplanner.', items: [
        GRID('Gewenste configuratie', [
          { col: 'D', label: 'Kast type' }, { col: 'E', label: 'DU type' }, { col: 'F', label: 'Radio sector 1' },
          { col: 'G', label: 'Radio sector 2' }, { col: 'H', label: 'Radio sector 3' }, { col: 'I', label: 'Tx unit' },
        ], [...bandsFrom(106).map((b) => ({ ...b, only: b.r === 106 ? null : ['D', 'E', 'F', 'G', 'H'] })),
          { r: 114, label: 'Mechanical tilt 5G 3500', only: ['F', 'G', 'H'] }]),
        GRID('Filters', [{ col: 'D', label: 'Filter' }], bandsFrom(117, BANDS_B.slice(0, 7))),
        T('J117', 'Li-Ion batterijen', { hint: 'Altijd plaatsen bij een nieuwe enclosure.' }),
        GRID('Antennes A1', [{ col: 'B', label: 'Antenne typenummer' }, { col: 'D', label: 'Azimuth' }, { col: 'E', label: 'ABH' }], SECTOR_ROWS(127)),
        GRID('Antennes A2 (L3500)', [{ col: 'G', label: 'Antenne typenummer' }, { col: 'I', label: 'Azimuth' }, { col: 'J', label: 'ABH' }], SECTOR_ROWS(127),
          { hint: 'Hoogte altijd bovenzijde gelijk of hoger dan de 2L6H antennes.' }),
        T('A132', 'Eltec PSU'), N('B132', 'Eltec PSU – aantal'), T('E132', 'Extra rectifiers 6150 – aantal'),
        AREA('A', 135, 137, 'Opmerkingen'),
      ] },
      { title: '6. Transmissie', items: [
        AREA('A', 143, 144, 'Nieuwe fiberverbinding: omschrijf moeilijkheden in het tracé (bv. koven openen/sluiten). Contact nodig met transmissieteam (SEN TX)?'),
        GRID('6c. Dishes Vodafone', [{ col: 'B', label: 'Diameter (cm)' }, { col: 'C', label: 'A-zijde' }, { col: 'D', label: 'Richting' }, { col: 'E', label: 'Hoogte (m) t.o.v. dak' }],
          [150, 151, 152].map((r, i) => ({ r, label: `Dish ${i + 1}` }))),
      ] },
      { title: '7. Airco (alleen indoor)', items: [
        T('C155', 'Airco plaatsen'), T('B158', 'Fabrikaat'), T('B159', 'Type'), N('B160', 'Aantal'), T('C161', 'Plaats condensor'),
      ] },
      { title: '8. Gedeelde locatie', hint: 'Wordt de locatie gedeeld met één van onderstaande providers? Aankruisen indien van toepassing.', items: [
        T('B165', 'KPN/Althio'), T('B168', 'Odido'), T('E165', 'C2000'), T('E168', 'Other'),
        AREA('A', 173, 174, 'Indien andere partij, hier aangeven (bv. radio-omroep)'),
      ] },
      { title: '9. Bliksembeveiliging', items: [
        ...BLIKSEM.map((l, i) => T('F' + (178 + i), l)),
        AREA('A', 197, 199, 'Opmerkingen'),
      ] },
      { title: '10. Energievoorziening type A', hint: 'Nieuwe situatie.', items: [
        ...ENERGIE.map(([c, l]) => T(c, l)),
        AREA('A', 219, 221, 'Opmerkingen'),
        T('C224', 'K1 straatkast (80 cm RVS)'), T('C225', 'K1 met K2 uitbreiding'), T('E224', 'RAL / kleur straatkast'),
        AREA('A', 228, 229, 'Plaats aftakking landlord kWh-meter'),
        AREA('A', 232, 233, 'Tijdelijke aftakking landlord of OLO'),
        AREA('A', 237, 243, 'Zon-aggregaat Volta Energy: waar kan het aggregaat geplaatst worden?'),
      ] },
      { title: '12. Daken (hoofddak, dakopbouw)', items: [
        T('C246', 'Conditie dakbedekking'), T('C247', 'Is er nog dakgarantie?'),
        T('E249', 'Hoofddak: ligt er nieuwe dakbedekking?'), N('G249', 'Aantal m²'),
        T('D253', '12a. Vloer: prefab'), T('D254', '12a. Vloer: in het werk gestort'), T('D255', '12a. Vloer: houten balklaag'),
        T('D256', '12a. Vloer: staalconstructie met houten liggers'), T('D257', '12a. Vloer: staalconstructie met staalplaten'),
        T('F253', 'Wat is de dikte van de vloer?'),
        AREA('A', 261, 263, 'Opmerkingen vloer'),
        T('C267', '12b. Isolatie: ongeïsoleerd dak'), T('C268', '12b. Isolatie: warm dak'), T('C269', '12b. Isolatie: koud dak'), T('C270', '12b. Isolatie: omgekeerd dak'),
        T('C274', '12c. Dakbedekking: bitumen'), T('C275', '12c. Dakbedekking: kunststof/rubber'),
        AREA('A', 278, 280, '12c. Dakbedekking: other'),
        AREA('A', 283, 285, 'Met welk materiaal is het dak voorzien? (tegels, grind, other)'),
        AREA('A', 288, 289, 'Is het looppad te gebruiken conform richtlijnen van ARBO?'),
        AREA('A', 292, 293, 'Is de huisdakdekker bekend? Zo ja: naam, plaats, telefoonnummer'),
        T('B297', '12d. Wanden: steen'), T('B298', '12d. Wanden: beton'), T('B299', '12d. Wanden: hout'), T('B300', '12d. Wanden: staal'), T('B301', '12d. Wanden: glas'),
        T('C303', 'Kwaliteit wanden'),
        AREA('A', 307, 309, 'Opmerkingen'),
      ] },
      { title: '13. Toegang', items: [
        AREA('A', 313, 317, 'Zijn er specifieke eisen voor toegang? Zo ja, hier vermelden'),
        AREA('A', 320, 322, 'Wat zijn de gegevens van de landlord voor de toegang?'),
        T('D326', 'Personen: veiligheidsmiddelen benodigd'), T('D327', 'Personen: toegang via inpandige trap'), T('D328', 'Personen: toegang via kooiladder'),
        T('D329', 'Personen: toegang via lift'), T('D330', 'Personen: toegang via dakluik met ladder'),
        AREA('A', 333, 335, 'Personen: other'),
        T('E339', 'Goederen: toegang via inpandige trap recht/draai'), T('E340', 'Goederen: toegang via kooiladder'),
        T('E342', 'Goederen: toegang via lift'), T('E343', 'Goederen: toegang via bestaande weg'),
        AREA('A', 347, 353, 'Toegankelijkheid goederen: hoogte lift of aantal rijplaten'),
      ] },
      { title: '14. Gegevens kraan', items: [
        T('D356', 'Mobiele kraan benodigd?'), T('D357', 'Ruimte beschikbaar?'),
        AREA('A', 362, 367, 'Type kraan + locatie, type en m² rijplaten', {
          hint: 'Bv. circa 250 m1 rijplaten dubbel spoor, plateau bij inrijden pad en plateau bij site t.b.v. opstellen materieel.' }),
        T('F370', 'Maximum gewicht (kg)'), T('F371', 'Maximum afmetingen L×B×H (m)'),
        T('F372', 'Verharding toegang opstelplaats (asfalt, tegels, etc.)'), T('F373', 'Obstakels op of naast opstelplaats (bomen, lantaarnpalen etc.)'),
        T('F374', 'Ondergrondse kabels/leidingen'), T('F375', 'Bovengrondse kabels (trein, tram, etc.)'), T('F376', 'Overhangende bomen'),
        AREA('A', 379, 381, '14a. Zijn er rijplaten nodig? Zo ja, hoeveel en waar?'),
        AREA('A', 383, 384, '14b. Is er extra camerabewaking nodig voor de rijplaten?'),
      ] },
      { title: '15. Checklist grond', hint: 'Nog te verrichten werkzaamheden met betrekking tot het gebouw.', items: [
        T('F389', 'Milieurapport noodzakelijk'), T('F390', 'Grondonderzoek en funderingsadvies noodzakelijk'),
        T('F391', 'KLIC-melding noodzakelijk'), T('F392', 'Locatie inmeten met kadastrale gegevens of door kadaster'),
        AREA('A', 397, 399, 'Opmerkingen (bv. wordt dit door de gemeente verzorgd?)'),
      ] },
      { title: '16. Aanwezige faciliteiten', items: [
        T('E403', 'Schaftgelegenheid'), T('E404', 'Parkeerplaats beschikbaar'), T('E405', 'Toilet beschikbaar'), T('E406', 'Brandblusser aanwezig'),
        AREA('A', 409, 412, 'Opmerkingen'),
      ] },
    ],
    info: { title: 'Fotodocumentatie', lines: FOTO_INFO },
  };

  // ---------- hulpfuncties Word ----------
  // P: tekst in alinea, met optioneel vaste tekst ervoor/erna (prefix/suffix)
  const P = (label, t, r, c, p = 0, o = {}) => ({ id: `p${t}_${r}_${c}_${p}`, label, type: 'text', p: [t, r, c, p], ...o });
  // R: meerdere regels tekst vanaf alinea 'from' (t/m 'to') in een cel
  const R = (label, t, r, c, from = 1, o = {}) => ({ id: `a${t}_${r}_${c}_${from}`, label, type: 'area', rest: [t, r, c, from, o.to], ...o });
  // CB: los selectievakje
  const CB = (label, n, o = {}) => ({ id: `cb${n}`, label, type: 'check', cb: n, ...o });
  // CH: keuze uit selectievakjes; multi = meerdere mogelijk
  const CH = (label, opts, o = {}) => ({ id: `ch${opts[0][1]}`, label, type: 'choice', opts, ...o });
  const JN = (label, ja, nee, o = {}) => CH(label, [['Ja', ja], ['Nee', nee]], o);
  const WGRID = (title, cols, rowDefs, o = {}) => ({ grid: true, title, hint: o.hint, cols, rows: rowDefs });

  const VEILIGHEID = [
    [1, 'Valbeveiliging aanwezig'], [2, 'Standaard waarschuwingsborden'], [3, 'Anti-inklimbeveiliging'], [4, 'Söll ladder'],
    [5, 'Ladder met veiligheidslijn'], [6, 'FABA systeem'], [7, 'Latchways klimbeveiliging'], [8, 'Kooiladder 1'], [9, 'Kooiladder 2'],
    [10, 'Handgreep'], [11, 'Losse ladder met inhaakbeugel'], [12, 'Veiligheidslijn'], [14, 'Voetpad 60 cm breed met betontegels 60x40 op dragers'],
    [15, 'Hekwerk'], [17, 'Zwart/geel touw'], [19, 'Additioneel platform'], [20, 'Extra regel 1'], [21, 'Extra regel 2'],
  ];
  let cbn = 37;
  const veiligheidRows = VEILIGHEID.map(([r, label]) => {
    const row = {
      label,
      fields: [
        CB(`${label} – aanwezig`, cbn, { short: 'Aanwezig' }),
        P(`${label} – keuring geldig tot`, 5, r, 2, 0, { short: 'Keuring geldig tot' }),
        CB(`${label} – nog te installeren`, cbn + 1, { short: 'Nog te installeren' }),
        CB(`${label} – n.v.t.`, cbn + 2, { short: 'N.v.t.' }),
      ],
    };
    cbn += 3;
    return row;
  });

  const RAL = ['RBS', 'Antennes', 'Constructies', 'RRUS', 'Feeders', 'RET', 'TMA', 'Stijggoot', 'Elektrakast', 'Bliksembeveiliging', 'Kabels', 'Extra invulveld'];

  const kpn = {
    id: 'kpn-dak',
    klant: 'kpn',
    title: 'RCT rapport – Nieuwbouw daklocatie',
    subtitle: 'KPN · Connect',
    kind: 'docx',
    template: 'templates/kpn-rct-nieuwbouw-daklocatie.docx',
    fileName: (v) => `RCT rapport - Nieuwbouw daklocatie KPN${v.p0_1_1_0 ? ' - ' + v.p0_1_1_0 : ''}${v.p0_4_1_0 ? ' ' + v.p0_4_1_0 : ''}.docx`,
    titleField: 'p0_1_1_0',
    detect: (doc) => /IPO nummer/.test(doc.paraText(doc.cellParas(0, 2, 0)?.paras[0] || doc.body)),
    sections: [
      { title: '1. Algemene gegevens site', items: [
        P('Projectnummer', 0, 0, 1), P('Site nummer', 0, 1, 1), P('IPO nummer', 0, 2, 1), P('N-PROP nummer', 0, 3, 1),
        P('Plaatsnaam project', 0, 4, 1), P('Site adres', 0, 5, 1),
        CH('Site type', [['N2 – Nieuwe op niet-KPN-locatie bestaand gebouw', 0], ['N3 – Nieuwe op KPN-locatie zonder mast', 1]]),
        P('GPS X', 0, 7, 1, 0, { prefix: 'X: ', inputmode: 'decimal' }), P('GPS Y', 0, 7, 2, 0, { prefix: 'Y: ', inputmode: 'decimal' }),
        P('Latitude', 0, 8, 1, 0, { prefix: 'Lat: ', inputmode: 'decimal' }), P('Longitude', 0, 8, 2, 0, { prefix: 'Long: ', inputmode: 'decimal' }),
        JN('Noordpijl gecontroleerd en genoteerd op tekening', 2, 3),
        P('RCT-datum', 0, 10, 1, 0, { type: 'date' }),
      ] },
      { title: '1.1 Aanwezig tijdens RCT', items: [
        WGRID('Aanwezig', ['Naam', 'Telefoon', 'Paraaf'],
          ['Landlord / eigenaar', 'DAEL / Site Engineer', 'Radio Engineer KPN', 'Site Acquisitie Verwerver', 'Tekenaar / Scanner'].map((l, i) => ({
            label: l,
            fields: [1, 2, 3].map((c, k) => P(`${l} – ${['naam', 'telefoon', 'paraaf'][k]}`, 1, i + 1, c, 0, { short: ['Naam', 'Telefoon', 'Paraaf'][k], inputmode: c === 2 ? 'tel' : undefined })),
          }))),
        P('Overige aanwezige – functie', 1, 6, 0),
        P('Overige aanwezige – naam', 1, 6, 1), P('Overige aanwezige – telefoon', 1, 6, 2, 0, { inputmode: 'tel' }), P('Overige aanwezige – paraaf', 1, 6, 3),
      ] },
      { title: '1.2 Landlord informatie', items: [
        P('Naam', 2, 0, 1), P('Functie', 2, 1, 1), P('Adres', 2, 2, 1), P('E-mail', 2, 3, 1, 0, { inputmode: 'email' }),
        P('Telefoonnummer', 2, 4, 1, 0, { inputmode: 'tel' }), P('Bedrijf', 2, 5, 1),
      ] },
      { title: '2.1 Health & Safety (algemeen, voor aanvang RCT)', items: [
        CH('De gevaren en gedragsregels voor deze site zijn vóór aanvang van de RCT bekend of besproken (site-coördinator heeft centrale rol)', [['Akkoord', 4], ['Niet akkoord', 5], ['Nvt', 6]]),
        CH('Iedereen heeft een certificaat veilig werken op hoogte en waar nodig een ESD OVE2-certificaat', [['Akkoord', 7], ['Niet akkoord', 8], ['Nvt', 9]]),
        CH("Iedereen heeft de benodigde PBM's mee en draagt ze waar nodig (veiligheidsschoenen verplicht, ESD-jassen in KPN-centrales)", [['Akkoord', 10], ['Niet akkoord', 11], ['Nvt', 12]]),
        CH("Iedereen is geïnformeerd over eventuele EMV-risico's (vooral antennes van Nozema of broadcast partners)", [['Akkoord', 13], ['Niet akkoord', 14], ['Nvt', 15]]),
        JN('Zijn er personen of verkeer te verwachten rond of op de site tijdens of na de bouw?', 16, 17, { hint: 'Zo ja: extra maatregelen, tijdens werkzaamheden ruim afzetten met lint.' }),
        JN('Is er kans op een onveilige situatie? (vogels, vogelpoep, stoom, asbest, rook, onbeschermde machines)', 18, 19),
        P('Onveilige situatie – specificeer', 3, 5, 2, 0, { prefix: 'Specificeer: ' }),
      ] },
      { title: '2.2 Toegang', items: [
        CH('Toegang tot gebouw (omgeving)', [['Belemmeringen geconstateerd', 20], ['Parkeerplaats aanwezig (PMR nakijken, schets + afmetingen)', 21], ['Parkeerplaats niet aanwezig', 22]], { multi: true }),
        P('Uitleg belemmeringen', 4, 0, 1, 2),
        CH('Transport naar site', [['Vrachtwagen', 23], ['Vrachtwagen met kraan', 24], ['Niet toegankelijk met voertuig', 25]], { multi: true, hint: 'Check de aanrijroute voor grote voertuigen i.v.m. viaducten, bruggen etc.' }),
        CH('Toegang verschaffing tot dak', [['Nieuwe kluis plaatsen', 26], ['Geen kluis nodig of mogelijk (alleen in overleg met landlord)', 27], ['KPN company card', 28], ['Alticom sleutel', 29], ['Lips sleutel', 30], ['Anders', 31]], { multi: true }),
        P('Toegang dak – anders, namelijk', 4, 2, 1, 7),
        CH('Locatie kluis', [['Aan hek', 32], ['In de muur', 33], ['Betonnen paaltje', 34]]),
        CH('Type kluis', [['Vierkant', 35], ['Buis', 36]]),
      ] },
      { title: '2.2 Toegangsmiddelen naar dak / veiligheidsmiddelen', items: [
        WGRID('Veiligheidsmiddelen', ['Aanwezig', 'Keuring geldig tot (foto maken)', 'Nog te installeren', 'N.v.t.'], veiligheidRows),
        P('Omschrijving extra regel 1', 5, 20, 0), P('Omschrijving extra regel 2', 5, 21, 0),
        P('Type veiligheidslijn', 5, 13, 1), P('Hekwerk – hoeveel meters?', 5, 16, 3, 1, { inputmode: 'decimal' }),
        P('Zwart/geel touw – hoeveel meters?', 5, 18, 3, 1, { inputmode: 'decimal' }),
        R('Opmerkingen (daktoegang is veilig en gekeurd)', 5, 22, 0, 1),
        CH('Lift aanwezig?', [['Aanwezig', 91], ['Niet aanwezig', 92], ['N.v.t.', 93]]),
        P('Lift – max. toegestane belasting (kg)', 5, 23, 2, 0, { prefix: 'Max toegestane belasting? ', suffix: ' Kg', inputmode: 'decimal' }),
        P('Lift – afmetingen (L × B × H in m)', 5, 23, 2, 2),
        CH('Dakluik aanwezig?', [['Aanwezig', 94], ['Niet aanwezig', 95], ['NVT', 96]]),
        P('Dakluik in veilige zone?', 5, 25, 2, 0, { prefix: 'In veilige zone?  ', options: ['Ja', 'Nee'] }),
        P('Dakluik – afmetingen (L × B × H in m)', 5, 25, 2, 2),
        JN('Is er slip-, stoot- en/of struikelgevaar op de route?', 97, 98),
        P('Slip-/stoot-/struikelgevaar – waar?', 5, 27, 2, 0, { prefix: 'Waar? ' }),
        R('Opmerkingen (locatie aangeven in schets/tekening)', 5, 29, 0, 1, { to: 2 }),
      ] },
      { title: '3. Dakconstructie', items: [
        CH('Constructie', [['Beton in werk', 99], ['Breedplaat', 100], ['Cellenbeton', 101], ['Staal met plaatmat.', 102], ['Combinatie', 103], ['Staal met hout', 104], ['Hout', 105], ['Kanaalplaat', 106], ['Onbekend / niet duidelijk te zien', 107]], { multi: true, hint: 'Goed kijken of er d.m.v. poeren of losse fundatie geplaatst kan worden.' }),
        CH('Dakisolatie', [['Ja, wel drukvast', 108], ['Ja, niet drukvast', 109], ['Nee', 110]]),
        P('Dikte isolatie (i.v.m. hoogte poeren, mm)', 6, 1, 2, 0, { prefix: 'Dikte (i.v.m. hoogte poeren): ', suffix: ' mm', inputmode: 'decimal' }),
        CH('Plaats dakisolatie', [['Onder dakleer', 111], ['Bovenop (warm dak)', 112]]),
        CH('Bedekkingsmateriaal', [['Bitumen (APP/SBS)', 113], ['PVC', 114], ['Kunststof', 115], ['Rubber', 116], ['Zink/koper', 117], ['Anders', 118]], { multi: true }),
        P('Bedekkingsmateriaal – anders, namelijk', 6, 4, 1, 6),
        CH('Ballast', [['Grind', 119], ['Tegels', 120], ['Looppad(en)', 121], ['Geen', 122]], { multi: true }),
        R('Opmerkingen', 6, 6, 0, 1),
        CH('Uitvoering dakbedekkingswerkzaamheden', [['Huisaannemer', 123], ['DAEL', 124]]),
        P('Huisaannemer dakbedekking – naam contactpersoon', 6, 10, 1), P('Huisaannemer dakbedekking – bedrijf', 6, 11, 1),
        P('Huisaannemer dakbedekking – e-mail', 6, 12, 1, 0, { inputmode: 'email' }), P('Huisaannemer dakbedekking – telefoonnummer', 6, 13, 1, 0, { inputmode: 'tel' }),
        R('Huisaannemer dakbedekking – opmerkingen', 6, 14, 0, 1),
      ] },
      { title: '4. Doorvoeren', items: [
        CH('Dakdoorvoer nodig?', [['Wendyhouse', 125], ['Wandelstok / zwanenhals', 126], ['N.v.t.', 127]]),
        JN('Muurdoorvoer nodig? (Roxtec)', 128, 129),
        JN('Brandvertragend laten afdichten na montage kabels?', 130, 131),
        JN('Moet er geboord worden?', 132, 133),
        P('Boren – waarin?', 7, 3, 2, 0, { prefix: 'Waarin? ' }),
        JN('Is er asbest aanwezig op de locatie van de doorvoer?', 134, 135),
        P('Asbest – waar?', 7, 5, 2, 0, { prefix: 'Waar? ' }),
        R('Opmerkingen doorvoeren', 7, 7, 0, 0),
      ] },
      { title: '5.1 Apparatuur indoor', items: [
        P('Locatie / ruimte (verdieping / ruimtenummer)', 8, 0, 1),
        CH('Toegang naar ruimte (indoor)', [['Nieuwe KPN cilinder', 136], ['Sleutel/pasje via landlord', 137], ['Anders', 138]], { hint: 'Schets maken waar apparatuur komt te staan, met afmetingen.' }),
        P('Toegang ruimte – anders, namelijk', 8, 2, 1, 3),
        P('Afmetingen ruimte – L', 8, 3, 1, 0, { prefix: 'L: ', inputmode: 'decimal' }),
        P('Afmetingen ruimte – B', 8, 3, 2, 0, { prefix: 'B: ', inputmode: 'decimal' }),
        P('Afmetingen ruimte – H', 8, 3, 3, 0, { prefix: 'H: ', inputmode: 'decimal' }),
        P('RBS-types (indoor)', 8, 4, 1),
        CH('Koeling', [['Niet aanwezig', 139], ['Nvt', 140], ['Ventilator', 141], ['Airco', 142]], { multi: true }),
        P('Airco – capaciteit', 8, 6, 2, 0, { prefix: 'Capaciteit: ' }), P('Airco – type', 8, 6, 2, 1, { prefix: 'Type: ' }), P('Airco – merk', 8, 6, 2, 2, { prefix: 'Merk: ' }),
        R('Opmerkingen (denk aan bv. temperatuur)', 8, 7, 0, 1),
      ] },
      { title: '5.2 Apparatuur outdoor', items: [
        P('RBS-types (outdoor)', 9, 0, 1),
        JN('Maaiveld: kan en mag er apparatuur binnen bestaand hekwerk bijgeplaatst worden? (aangeven op schets)', 143, 144),
        P('Uitbreiding opstelplaats RBS – L', 10, 1, 1, 0, { prefix: 'L: ', inputmode: 'decimal' }),
        P('Uitbreiding opstelplaats RBS – B', 10, 1, 2, 0, { prefix: 'B: ', inputmode: 'decimal' }),
        P('Hoeveel m² uitbreiding stelconplaten?', 10, 2, 1, 0, { suffix: ' m2', inputmode: 'decimal' }),
        P('Hoeveel m² uitbreiding bestrating?', 10, 3, 1, 0, { suffix: ' m2', inputmode: 'decimal' }),
        P('Uitbreiding hekwerk – L', 10, 4, 1, 0, { prefix: 'L: ', inputmode: 'decimal' }),
        P('Uitbreiding hekwerk – B', 10, 4, 2, 0, { prefix: 'B: ', inputmode: 'decimal' }),
        P('Uitbreiding hekwerk – H', 10, 4, 3, 0, { prefix: 'H: ', inputmode: 'decimal' }),
        JN('Hekwerkpoort nodig?', 145, 146),
        CB('Aanrijpaaltjes nodig? – Nee', 147, { hint: 'Wel nodig? Vul dan hieronder het aantal in.' }),
        P('Hoeveel aanrijpaaltjes?', 10, 7, 1, 0, { inputmode: 'numeric' }),
        JN('Dak: kan en mag er apparatuur op de draagconstructie geplaatst worden? (aangeven op schets)', 148, 149),
        JN('Dak: kan en mag er apparatuur op bestaande constructie (voormalig OLO) geplaatst worden? (schets + constructie inmeten)', 150, 151),
      ] },
      { title: '5.3 Antennebekabeling / feedertracé', items: [
        R('Indoor: door welke ruimtes loopt het tracé en hoe is de toegang daarnaartoe?', 12, 0, 0, 1),
        JN('Indoor: voldoende ruimte in de bestaande kabeldoorvoer?', 152, 153),
        P('Diameter doorvoer', 12, 2, 1),
        CH('Type doorvoer', [['Beton', 154], ['Hout', 155], ['Staal', 156], ['Steen', 157], ['Anders', 158]], { multi: true }),
        CB('Indoor: voldoende ruimte in de bestaande goten – Ja', 159),
        P('Indoor: hoeveel meter goot aanbrengen?', 12, 5, 1, 0, { suffix: ' m', inputmode: 'decimal' }),
        P('Indoor: type goot', 12, 5, 2, 0, { prefix: 'Type: ' }),
        P('Lengte indoor feedertracé (m)', 12, 6, 1, 0, { suffix: ' m', inputmode: 'decimal', hint: 'Teken de route in op de schets.' }),
        R('Indoor tracé – opmerkingen', 12, 8, 0, 1),
        R('Outdoor: beknopte beschrijving tracé', 13, 0, 0, 1),
        P('Outdoor: moeten er nieuwe voorzieningen geplaatst worden? Zo ja, welke?', 13, 1, 1),
        P('Outdoor: hoeveel meter goot aanbrengen?', 13, 2, 1, 0, { suffix: ' m', inputmode: 'decimal' }),
        P('Outdoor: type goot', 13, 2, 2, 0, { prefix: 'Type: ' }),
        P('Lengte outdoor feedertracé (m)', 13, 3, 1, 0, { suffix: ' m', inputmode: 'decimal', hint: 'Teken de route in op de schets; geef in het bovenaanzicht aan welke kabelstrip geplaatst moet worden.' }),
        R('Outdoor tracé – opmerkingen', 13, 5, 0, 1),
      ] },
      { title: '6. RAL-kleuren', items: [
        WGRID('RAL-kleuren', ['Van toepassing', 'RAL'], RAL.map((l, i) => ({
          label: l,
          fields: [CB(`RAL ${l} – van toepassing`, 160 + i, { short: 'Van toepassing' }), P(`RAL ${l}`, 14, i, 1, 0, { prefix: 'RAL: ', short: 'RAL' })],
        }))),
      ] },
      { title: '7. Kraan / hoogwerker', items: [
        R('Bijzonderheden', 15, 0, 0, 1),
        JN('Kraan nodig?', 172, 173),
        R('Mogelijke opstelplaatsen kraan (indien nodig met landlord bespreken)', 15, 2, 0, 2),
        JN('Hoogwerker nodig?', 174, 175),
        R('Mogelijke opstelplaatsen hoogwerker (indien nodig met landlord bespreken)', 15, 4, 0, 2),
        JN('Survey uitvoeren', 176, 177),
        JN('Rijplaten benodigd?', 178, 179),
        P('Rijplaten – lengte tracé (m)', 15, 6, 1, 1, { prefix: 'Lengte tracé: ', suffix: ' m', inputmode: 'decimal' }),
        JN('Toegang naar site mogelijk met kraan/hoogwerker?', 180, 181),
        JN('Inschatting: vergunning nodig?', 182, 183),
        JN('Verkeersregelaars nodig?', 184, 185),
        JN('Wegafzetting nodig?', 186, 187),
      ] },
      { title: '8. Elektravoorziening', items: [
        CH('Aansluiting', [['Eigen E-aansluiting (KPN wil meestal eigen aansluiting)', 188], ['Via tussenmeter van landlord', 189]]),
        R('Locatie K1 of tussenmeter landlord (aangeven op tekening/plattegrond en/of ruimtenummer)', 16, 1, 0, 1),
        CH('Uitvoering elektrawerkzaamheden', [['Huisaannemer', 190], ['DAEL', 191]]),
        P('Huisaannemer elektra – naam contactpersoon', 18, 1, 1), P('Huisaannemer elektra – bedrijf', 18, 2, 1),
        P('Huisaannemer elektra – e-mail', 18, 3, 1, 0, { inputmode: 'email' }), P('Huisaannemer elektra – telefoonnummer', 18, 4, 1, 0, { inputmode: 'tel' }),
        R('Huisaannemer elektra – opmerkingen', 18, 5, 0, 1),
        CH('8.1 Outdoor: van wie is het terrein?', [['Gemeente', 192], ['Privé', 193], ['Anders', 194]], { hint: 'Denk eraan: energiebedrijf graaft niet over privéterrein.' }),
        P('Terrein – anders, namelijk', 19, 0, 1, 3),
        CH('Ingraven voedingskabel (outdoor)', [['KLIC-melding', 195], ['Vergunning', 196]], { multi: true }),
        P('Lengte voedingskabel (m)', 19, 2, 1, 0, { suffix: ' M', inputmode: 'decimal' }),
        R('8.1 Outdoor – opmerkingen', 19, 3, 0, 1),
        R('8.2 Indoor: door welke ruimtes loopt het tracé en hoe is de toegang daarnaartoe?', 20, 0, 0, 1),
        JN('8.2 Indoor: voldoende ruimte in de bestaande kabeldoorvoer?', 197, 198),
        JN('8.2 Indoor: voldoende ruimte in de bestaande goten?', 199, 200),
        P('8.2 Hoeveel meter goot aanbrengen?', 20, 3, 1, 0, { suffix: ' m', inputmode: 'decimal' }),
        P('8.2 Type goot', 20, 3, 2, 0, { prefix: 'Type: ' }),
        P('8.2 Lengte indoor tracé (m)', 20, 4, 1, 0, { suffix: ' m', inputmode: 'decimal', hint: 'Teken de route in op de schets.' }),
        R('8.2 Indoor – opmerkingen', 20, 6, 0, 1),
        JN('8.3 Voldoende verlichting?', 201, 202),
        P('Zo nee: waar moet de verlichting geplaatst worden?', 21, 1, 2),
        P('Zo nee: waar moet de schakelaar geplaatst worden?', 21, 2, 2, 0, { hint: 'Verlichting intekenen op schets.' }),
      ] },
      { title: '9. Aarding / bliksembeveiliging', items: [
        CH('Type aarding', [['Aluminium', 203], ['Koper', 204], ['Staal', 205]]),
        P('Aantal aardputten (bij apparatuur op maaiveld, aangeven op schets)', 22, 1, 1, 0, { inputmode: 'numeric' }),
        CH('Uitvoering aarding/bliksembeveiligingswerkzaamheden', [['Huisaannemer', 206], ['DAEL', 207]]),
        P('Huisaannemer aarding – naam contactpersoon', 22, 5, 1), P('Huisaannemer aarding – bedrijf', 22, 6, 1),
        P('Huisaannemer aarding – e-mail', 22, 7, 1, 0, { inputmode: 'email' }), P('Huisaannemer aarding – telefoonnummer', 22, 8, 1, 0, { inputmode: 'tel' }),
        R('Huisaannemer aarding – opmerkingen', 22, 9, 0, 1),
      ] },
      { title: '10. Antennes', items: [
        JN('Zijn er vreemde antennes aanwezig?', 208, 209),
        CH('Zo ja, welke antennes of providers?', [['OLO', 210], ['Niet OLO', 211], ['FM-antennes', 212], ['Niet bekend – opvragen', 213]], { multi: true }),
        P('Vreemde antennes – namelijk', 23, 1, 2, 0, { prefix: 'Namelijk:  ' }),
        JN('Kans op te hoge blootstelling aan elektromagnetische velden?', 214, 215),
        JN('Uitgebreide metingen nodig? (indicatie)', 216, 217),
        JN('Is er outage nodig? (indicatie)', 218, 219),
        P('Outage – inschatting hoelang?', 23, 6, 1),
        R('Opmerkingen (geen andere antennes in de omgeving)', 23, 7, 0, 1),
        WGRID('10.2 Nieuw te plaatsen antennes', ['Type', 'Richting (°)', 'Hoogte', 'Bevestiging'], [1, 2, 3].map((s) => ({
          label: `Sector ${s}`,
          fields: [
            P(`Sector ${s} – type`, 24, s, 1, 0, { short: 'Type' }),
            P(`Sector ${s} – richting`, 24, s, 2, 0, { suffix: '°', short: 'Richting (°)', inputmode: 'decimal' }),
            P(`Sector ${s} – hoogte`, 24, s, 3, 0, { short: 'Hoogte', inputmode: 'decimal' }),
            P(`Sector ${s} – bevestiging`, 24, s, 4, 0, { short: 'Bevestiging' }),
          ],
        })), { hint: 'Bevestiging: bv. nieuw te plaatsen constructies op het dak, beugels, wandmontage, etc.' }),
        R('Opmerkingen antennes (LORA: ja/nee)', 25, 0, 0, 1),
        R('10.3 Bestaande OLO-antennes (controleer richting en hoogte a.d.h.v. tekeningen; schotels en andere partijen meenemen in redlines)', 26, 0, 0, 3),
      ] },
      { title: '11–13. Schets, waarnemingen & ten slotte', hint: 'Schets t.b.v. redlines: boven- en vooraanzicht, hoogte gebouw, dakafmetingen en stramienmaten.', items: [
        JN('Persoonlijke waarnemingen genoteerd?', 220, 221),
        R('Zijn er eigenaar-specifieke zaken die meegenomen moeten worden bij de bouw? (vandalisme etc.)', 28, 1, 0, 0),
        R('Andere bijzonderheden om rekening mee te houden', 29, 0, 0, 1, { to: 1 }),
        R('Uitgangspunten', 29, 0, 0, 3),
      ] },
    ],
  };


  // ---------- Odido – V&G-plan Monet (klassieke Word-formuliervelden, genummerd in documentvolgorde) ----------
  const FT = (label, i, o = {}) => ({ id: `ff${i}`, label, type: 'text', ff: i, ...o });
  const FA = (label, i, o = {}) => FT(label, i, { type: 'area', rows: 3, ...o });
  const FC = (label, i, o = {}) => ({ id: `ff${i}`, label, type: 'check', ff: i, ...o });
  const FCH = (label, opts, o = {}) => ({ id: `ffch${opts[0][1]}`, label, type: 'choice', legacy: true, opts, ...o });
  // Risico met vinkje 'van toepassing', maatregel/opmerking en eventueel een eigen omschrijving
  const RISK = (label, cb, text, name) => ({
    group: true,
    title: label,
    fields: [
      FC(`${label} – van toepassing`, cb, { short: 'Van toepassing' }),
      ...(name != null ? [FT(`${label} – omschrijving`, name, { short: 'Omschrijving' })] : []),
      ...(text != null ? [FA(`${label} – maatregel / opmerking`, text, { short: 'Maatregel / opmerking' })] : []),
    ],
  });
  const FGRID = (title, cols, rowDefs, o = {}) => ({
    grid: true, title, hint: o.hint, cols,
    rows: rowDefs.map(([label, ids]) => ({
      label,
      fields: ids.map((n, k) => (n == null ? null : (o.checks && o.checks.includes(k)
        ? FC(`${title} – ${label} – ${cols[k]}`, n, { short: cols[k] })
        : FT(`${title} – ${label} – ${cols[k]}`, n, { short: cols[k], type: o.dates && o.dates.includes(k) ? 'date' : 'text', inputmode: o.modes && o.modes[k] })))),
    })),
  });
  const seq = (start, n, step = 1) => Array.from({ length: n }, (_, k) => start + k * step);
  const CONTACT_ROWS = ['Company name', 'Visiting address', 'Postal address', 'Postcode', 'Town', 'Phone general', 'Contact', 'Phone contact',
    'E-mail contact', 'Supervisor', 'Phone supervisor', 'Direct supervisor', 'Phone direct supervisor', 'H&S official', 'Phone H&S official'];
  const contactGrid = (title, cols, start, hint) => FGRID(title, cols, CONTACT_ROWS.map((l, k) => [l, seq(start + k * 3, 3)]),
    { hint, modes: [] });

  const odido = {
    id: 'odido-vg',
    klant: 'odido',
    title: 'V&G-plan Monet – Rooftop',
    subtitle: 'Odido · Monet 6.2.1 (ENG) · No Subco',
    kind: 'docx',
    template: 'templates/odido-vg-plan-monet-6.2.1-rooftop.docx',
    fileName: (v) => `VG-plan Monet 6.2.1 ODIDO - Rooftop${v.ff0 ? ' - ' + v.ff0 : ''}${v.ff82 ? ' ' + v.ff82 : ''}.docx`,
    titleField: 'ff0',
    defaults: { ff3: 'ODIDO' },
    emptyValues: ['dd-mm-yy'],
    detect: (doc) => doc.legacy.some((f, i) => doc.legacyName(i) === 'PROVIDOR') && doc.legacy.some((f, i) => doc.legacyName(i) === 'SITE_NR'),
    sections: [
      { title: 'Project information', items: [
        FT('Site number', 0), FT('Location number or x/y coordinates', 1), FT('Order number', 2), FT('Provider', 3),
        FT('Version number', 4, { inputmode: 'numeric' }), FT('Date', 5, { type: 'date' }), FT('Status', 6, { dropdown: true }),
      ] },
      { title: 'Version management', items: [
        FGRID('Versions', ['Drafted by', 'Date', 'Changes'], seq(0, 6).map((v) => [`Version ${v}`, seq(7 + v * 3, 3)]), { dates: [1] }),
      ] },
      { title: 'Validation and distribution', hint: 'A = accepted · AV = accepted but subject to remarks · NA = not accepted, resubmit, see remarks.', items: [
        FGRID('Validation', ['Party', 'Name', 'Function', 'A', 'AV', 'NA', 'Initialled', 'Remarks'],
          seq(0, 7).map((k) => [`Regel ${k + 1}`, seq(25 + k * 8, 8)]), { checks: [3, 4, 5] }),
      ] },
      { title: 'General information', items: [
        FT('Address', 81), FT('Town', 82), FT('Company name (landlord)', 83),
        FCH('Type of work / structure', [['New construction', 84], ['Modification', 85], ['Management and maintenance', 86]], { multi: true }),
        FCH('Type location', [['Flat roof', 87], ['Antenna pole', 88], ['Silo', 89], ['Windmill', 90], ['Church tower', 91], ['Chimney', 92],
          ['High Voltage Pylon (HVP)', 93], ['High Voltage Tower', 94], ['High Voltage Station', 95], ['Indoor', 96], ['Billboard', 97], ['Other', 98]], { multi: true }),
        FT('Type location – other, specify', 99),
        FT('Start of work', 100, { type: 'date' }), FT('Scheduled end of work', 101, { type: 'date' }),
        FA('Type of work (describe activities)', 102, { rows: 5, hint: 'For TenneT describe activities day by day.' }),
      ] },
      { title: 'Introduction training', hint: 'Iedereen krijgt vóór de start van het werk elke dag een startwerkinstructie met dit formulier.', items: [
        FT('Date', 105, { type: 'date' }), FT('Trainer', 107),
        FCH('Kopie naar', [['Project supervisor', 103], ['Direct supervisor', 104], ['H&S official', 106], ['Principal (if arranged)', 108]], { multi: true }),
        FGRID('Attendants introduction training', ['Date', 'Name', 'Company'], seq(0, 24).map((k) => [`${k + 1}`, seq(109 + k * 3, 3)]), { dates: [0] }),
      ] },
      { title: '3.1 Discussed topics', hint: 'Alle aangevinkte onderwerpen moeten besproken zijn.', items: [
        RISK('Safe access to the location / site specific risks (see Site RI&E)', 181, 182),
        RISK('Availability Safety Regulations for working at KPN, T-Mobile, Tele2 and Vodafone locations', 183, 184),
        RISK('Safety passports / training ok? (AFAS Pocket app)', 185, 186),
        RISK('Required personal protection equipment', 187, 188),
        RISK('Additional personal protection equipment', 189, 190),
        RISK('On-site emergency services (first aid)', 191, 192),
        RISK('Regulations regarding smoking, alcohol, drugs and medications', 193, 195),
        FA('Smoking, alcohol, drugs – rules', 194),
        RISK('Appointments sanitation (toilet, hygiene)', 196, 197),
        RISK('Order and cleanliness', 198, 199),
        RISK('Slips, trips, falls', 200, 201),
        RISK('Shock hazard to low or high constructions', 202, 203),
        RISK('Safe storage and use of hazardous substances', 204, 205),
        RISK('Separate collection of waste', 206, 207),
        RISK('Prevent environmental pollution and damage to flora and/or fauna', 208, 209),
      ] },
      { title: '3.2 Specific risks landlord', items: [
        RISK('Work on/near a high voltage installation (see 3.3)', 210),
        RISK('Special access procedures', 211, 212),
        RISK('Work permits', 213, 214),
        RISK('Cellnex towers (specific measures applicable)', 215, 216),
        RISK('Work along the road', 217, 218),
        RISK('Work along the track/railroad (training + ProRail access pass)', 219, 220),
        RISK('Extra risico 1', 221, 223, 222),
        RISK('Extra risico 2', 224, 226, 225),
        RISK('Held toolbox meeting(s) (specific risks)', 227, 228),
        FA('Remarks', 229),
      ] },
      { title: '3.3 Working on/nearby high voltage installations', items: [
        FC('Working near high voltage, electrocution hazard (supervision NEN3840)', 230),
        FC('Climbing a high voltage power pole (HVP) or tower', 231),
        FC('Working on a HVP or HV substation and/or installation of mast reinforcements', 232),
        FC('Performing telecom operations near power lines', 233),
        FC('Working with a cherry-picker on or near a HV pylon or power line', 234),
        FC('Working with a crane with workplatform on or near a HV pylon or power line', 235),
        FC('Hoist with a mobile crane on or near a HV pylon or power line', 236),
        FC('Hoist with a hoisting winch on or near a HV pylon or power line', 237),
        FC('Hoist manually in the pole body at or near a HV pylon or power line', 238),
      ] },
      { title: '3.4 Working at heights', items: [
        RISK('1. Within two meters from the eaves', 239, 240),
        RISK('2. Working in the 2–4 meter zone', 241, 242),
        RISK('3. Roof slope in excess of > 15°', 243, 244),
        RISK('4. Initial installation of fall protection equipment', 245, 246),
        RISK('5. Installation of cable ducts', 247, 248),
        RISK('6. Installation of earth leads', 249, 250),
        RISK('7. Installation of antenna pole, antennas, feeders etc.', 251, 252),
        RISK('8. Risk of falling objects', 253, 254),
        RISK('9. Working from a (landing) staircase / ladder', 255, 256),
        RISK('10. Working with (rolling) scaffolding', 257, 258),
        RISK('11. Rope access (working from safety lines)', 259, 260),
        RISK('Other', 261, 263, 262),
        FA('Remarks', 264),
      ] },
      { title: '3.5 Horizontal and vertical transport', items: [
        RISK('1. Hoisting with mobile crane (materials)', 265, 267),
        RISK('2. Working from a crane with workplatform', 268, 269),
        RISK('3. Hoisting with a winch', 270, 271),
        RISK('4. Hoisting manually', 272, 273),
        RISK('5. Working with a cherry picker', 274, 275),
        RISK('6. Attaching loads, using eyebolts and other lifting equipment', 276, 277),
        RISK('Other', 278, 279),
        FA('Remark (work in accordance with Arbo catalogue “Vertical transport”)', 280),
      ] },
      { title: '3.6 Physical factors', items: [
        RISK('1. Working near live electrical components', 281, 282),
        RISK('2. Electromagnetic fields (EMF) exposure', 283, 284),
        RISK('3. Lasers (optical links) > class 1M', 285, 286),
        RISK('4. Noise (church bells, machinery, drilling)', 287, 288),
        RISK('5. Inadequate lighting workstation', 289, 290),
        RISK('6. Injury when using power tools', 291, 292),
        RISK('7. Cold / heat stress', 293, 294),
        RISK('Other', 295, 297, 296),
        FA('Remarks', 298),
      ] },
      { title: '3.7 Hazardous materials', items: [
        RISK('1. Handling and storing hazardous materials', 299, 300),
        RISK('2. Industrial products (near workstation)', 301, 302),
        RISK('3. Flue gasses (near workstation)', 303, 304),
        RISK('4. Exposure to particulate matter (quartz, hardwood)', 305, 306),
        RISK('5. Risk of exposure to asbestos', 307, 308),
        RISK('6. Fibrous materials (glass wool, rock wool)', 309, 310),
        RISK('7. Opening and closing cable conduits', 311, 312),
        RISK('8. Risk of exposure to polluted soil', 313, 314),
        RISK('Other', 315, 317, 316),
        FA('Remarks', 318),
      ] },
      { title: '3.8 Biological agents', items: [
        RISK('1. Legionella', 319, 320),
        RISK('2. Animal droppings (rats, pigeons)', 321, 322),
        RISK('3. Insects (ticks, processionary caterpillar, wasps’ nests)', 323, 324),
        RISK('4. Irritant plants (giant hogweed)', 325, 326),
        RISK('5. Biological matters (near workstation)', 327, 328),
        RISK('Other', 329, 331, 330),
        FA('Remarks', 332),
      ] },
      { title: '3.9 Physical strain', items: [
        RISK('1. Climbing (e.g. double hooks)', 333, 334),
        RISK('2. Manually moving heavy objects', 335, 336),
        RISK('Other', 337, 339, 338),
        FA('Remarks', 340),
      ] },
      { title: '3.10 Environmental hazards', items: [
        RISK('1. Excavation work (CROW 500)', 341, 342),
        RISK('2. Working in confined space', 343, 344),
        RISK('3. Little movement at the workstation', 345, 346),
        RISK('4. Working in the engine room of a lift', 347, 348),
        RISK('5. Risk of fire, explosion (ATEX)', 349, 350),
        RISK('6. Work near or on the water', 351, 352),
        RISK('Other', 353, 355, 354),
        FA('Remarks', 356),
      ] },
      { title: 'Contact information', hint: '* = verplichte velden. Specifieke V&G-regels en risico’s van de landlord opvragen bij de contactpersoon.', items: [
        contactGrid('Principal / contractor / landlord', ['Principal', 'Contractor', 'Landlord / company'], 357),
      ] },
      { title: 'Onderaannemers', hint: 'Alleen invullen als er onderaannemers of nevenaannemers zijn.', items: [
        contactGrid('Onderaannemers 1–3', ['Onderaannemer 1', 'Onderaannemer 2', 'Onderaannemer 3'], 402),
        contactGrid('Onderaannemers 4–5 en nevenaannemer', ['Onderaannemer 4', 'Onderaannemer 5', 'Nevenaannemer'], 447),
      ] },
      { title: 'H&S coordinators main contractor', items: [
        FGRID('H&S coordinators', ['Name', 'Phone', 'E-mail'], [
          ['V&G coördinator ontwerp', seq(492, 3)], ['Vervanger (optie)', seq(495, 3)], ['V&G coördinator uitvoering', seq(498, 3)],
          ['Vervanger (optie)', seq(501, 3)], ['Werkverantwoordelijke (WV)', seq(504, 3)],
        ], { modes: [undefined, 'tel', 'email'] }),
      ] },
      { title: 'Project organisation', hint: 'Extra vakken in het organogram. De meeste namen worden automatisch overgenomen uit de contactgegevens.', items: [
        FT('Secondary contractor(s)', 507),
        FT('Onder onderaannemer 2', 508), FT('Onder onderaannemer 3', 509), FT('Onder onderaannemer 4', 510), FT('Onder onderaannemer 5', 511),
        FT('H&S coordinator (onderaannemer)', 512), FT('Supervisor op locatie (onderaannemer)', 513),
        FT('Onder V&G coördinator uitvoering', 514), FT('Onder leidinggevende op locatie', 515), FT('Mechanics', 516),
      ] },
      { title: 'On-site emergencies', items: [
        FT('On-site emergency services (BHV) alarm number', 517, { inputmode: 'tel' }),
        FGRID('Medische hulp', ['Name', 'Phone (+31)', 'Address'], [['Local hospital', [518, 519, 520]], ['Local doctor', [521, 522, 523]]], { modes: [undefined, 'tel'] }),
      ] },
    ],
  };

  // ---------- Cellnex – TSSR Reworks (zelfde opbouw als KPN: tabel/rij/cel/alinea en selectievakje-nummers) ----------
  const CX_VEILIGHEID = [
    [1, 'Valbeveiliging aanwezig'], [2, 'Standaard waarschuwingsborden'], [3, 'Anti-inklimbeveiliging'], [4, 'Söll ladder'],
    [5, 'Ladder met veiligheidslijn'], [6, 'FABA systeem'], [7, 'Latchways klimbeveiliging'], [8, 'Kooiladder 1'], [9, 'Kooiladder 2'],
    [10, 'Handgreep'], [11, 'Losse ladder met inhaakbeugel'], [12, 'Veiligheidslijn'], [14, 'Voetpad 60 cm breed met betontegels 60x40 op dragers'],
    [15, 'Hekwerk'], [17, 'Hekwerk hoogte en tussenliggers afstand'], [18, 'Rood/wit ketting of zwart/geel touw'], [20, 'Additioneel platform'],
    [21, 'Extra regel 1'], [22, 'Extra regel 2'],
  ];
  let cxn = 30;
  const cxVeiligheidRows = CX_VEILIGHEID.map(([r, label]) => {
    const noInstall = r === 17; // in deze regel zit geen vakje 'nog te installeren'
    const row = {
      label,
      fields: [
        CB(`${label} – aanwezig`, cxn, { short: 'Aanwezig' }),
        P(`${label} – keuring geldig / opmerkingen`, 6, r, 2, 0, { short: 'Keuring geldig / opmerkingen' }),
        noInstall ? null : CB(`${label} – nog te installeren`, cxn + 1, { short: 'Nog te installeren' }),
        CB(`${label} – n.v.t.`, noInstall ? cxn + 1 : cxn + 2, { short: 'N.v.t.' }),
      ],
    };
    cxn += noInstall ? 2 : 3;
    return row;
  });

  const cellnex = {
    id: 'cellnex-reworks',
    klant: 'cellnex',
    title: 'TSSR – Reworks',
    subtitle: 'Cellnex · TSS rapport Reworks',
    kind: 'docx',
    template: 'templates/cellnex-tssr-reworks.docx',
    fileName: (v) => `TSSR Reworks Cellnex${v.p0_1_1_0 ? ' - ' + v.p0_1_1_0 : ''}${v.p0_2_1_0 ? ' ' + v.p0_2_1_0 : ''}.docx`,
    titleField: 'p0_1_1_0',
    detect: (doc) => /Cellnex/i.test(doc.paraText(doc.cellParas(1, 3, 0)?.paras[0] || doc.body)),
    sections: [
      { title: '1. Algemene gegevens site', items: [
        P('Projectnummer', 0, 0, 1), P('Site nummer', 0, 1, 1), P('Plaatsnaam project', 0, 2, 1), P('Site adres', 0, 3, 1),
        CH('Site type', [['Dak', 0], ['Mast', 1], ['Kerk', 2], ['Silo', 3], ['Windmolen', 4], ['Schoorsteen', 5]]),
        P('GPS coördinaten – X / latitude', 0, 5, 1, 0, { inputmode: 'decimal' }), P('GPS coördinaten – Y / longitude', 0, 5, 2, 0, { inputmode: 'decimal' }),
        P('TSS-datum', 0, 6, 1, 0, { type: 'date' }),
      ] },
      { title: '1.1 Aanwezig tijdens TSS', items: [
        WGRID('Aanwezig', ['Naam', 'Telefoon', 'Paraaf'],
          ['Landlord / eigenaar', 'DAEL / Site Engineer', 'Cellnex / Odido', 'Althio / KPN', 'Ericsson / Vodafone', 'Engineering partij', 'Aannemer'].map((l, i) => ({
            label: l,
            fields: [1, 2, 3].map((c, k) => P(`${l} – ${['naam', 'telefoon', 'paraaf'][k]}`, 1, i + 1, c, 0, { short: ['Naam', 'Telefoon', 'Paraaf'][k], inputmode: c === 2 ? 'tel' : undefined })),
          }))),
      ] },
      { title: '1.2 Landlord informatie', items: [
        P('Naam', 2, 0, 1), P('Functie', 2, 1, 1), P('Adres', 2, 2, 1), P('E-mail', 2, 3, 1, 0, { inputmode: 'email' }),
        P('Telefoonnummer', 2, 4, 1, 0, { inputmode: 'tel' }), P('Bedrijf', 2, 5, 1),
      ] },
      { title: '1.3 Aannemer informatie', hint: 'Kies een onderaannemer uit de lijst om de gegevens automatisch in te vullen.', items: [
        P('Bedrijf', 3, 5, 1), P('Naam', 3, 0, 1), P('Functie', 3, 1, 1), P('Adres', 3, 2, 1), P('E-mail', 3, 3, 1, 0, { inputmode: 'email' }),
        P('Telefoonnummer', 3, 4, 1, 0, { inputmode: 'tel' }),
      ] },
      { title: '2.1 Health & Safety (algemeen, voor aanvang TSS)', items: [
        CH('De gevaren en gedragsregels voor deze site zijn vóór aanvang van de TSS bekend of besproken (site-coördinator heeft centrale rol)', [['Akkoord', 6], ['Niet akkoord', 7], ['Nvt', 8]]),
        CH("Iedereen heeft de benodigde PBM's mee en draagt ze waar nodig (veiligheidsschoenen verplicht)", [['Akkoord', 9], ['Niet akkoord', 10], ['Nvt', 11]]),
        CH("Iedereen is geïnformeerd over eventuele EMV-risico's (vooral antennes van Nozema of broadcast partners)", [['Akkoord', 12], ['Niet akkoord', 13], ['Nvt', 14]]),
        JN('Zijn er personen of verkeer te verwachten rond of op de site tijdens of na de bouw?', 15, 16, { hint: 'Zo ja: extra maatregelen, tijdens werkzaamheden ruim afzetten met lint.' }),
        JN('Is er kans op een onveilige situatie? (vogels, vogelpoep, stoom, asbest, rook, onbeschermde machines)', 17, 18),
        P('Onveilige situatie – specificeer', 4, 4, 2, 0, { prefix: 'Specificeer: ' }),
      ] },
      { title: '2.2 Toegang', items: [
        CH('Toegang tot gebouw (omgeving)', [['Belemmeringen geconstateerd', 19], ['Parkeerplaats aanwezig (schets + afmetingen)', 20], ['Parkeerplaats niet aanwezig', 21]], { multi: true }),
        P('Uitleg belemmeringen', 5, 0, 1, 2),
        CH('Transport naar site', [['Vrachtwagen', 22], ['Vrachtwagen met kraan', 23], ['Niet toegankelijk met voertuig', 24]], { multi: true, hint: 'Check de aanrijroute voor grote voertuigen i.v.m. viaducten, bruggen etc.' }),
        CH('Toegang verschaffing tot dak', [['Sleutelkluis', 25], ['Anders', 26]]),
        P('Toegang dak – anders, namelijk', 5, 2, 1, 2),
        CH('Locatie kluis', [['Aan hek', 27], ['In de muur', 28], ['Betonnen paaltje', 29]]),
      ] },
      { title: '2.2 Toegangsmiddelen naar dak / veiligheidsmiddelen', items: [
        WGRID('Veiligheidsmiddelen', ['Aanwezig', 'Keuring geldig / opmerkingen', 'Nog te installeren', 'N.v.t.'], cxVeiligheidRows),
        P('Omschrijving extra regel 1', 6, 21, 0), P('Omschrijving extra regel 2', 6, 22, 0),
        P('Type veiligheidslijn', 6, 13, 1), P('Hekwerk – hoeveel meters?', 6, 16, 3, 1, { inputmode: 'decimal' }),
        P('Ketting/touw – hoeveel meters?', 6, 19, 3, 1, { inputmode: 'decimal' }),
        R('Opmerkingen veiligheidsmiddelen', 6, 23, 0, 1),
        CH('Lift aanwezig?', [['Aanwezig', 86], ['Niet aanwezig', 87], ['N.v.t.', 88]]),
        P('Lift – max. toegestane belasting (kg)', 6, 24, 2, 0, { prefix: 'Max toegestane belasting? ', suffix: ' Kg', inputmode: 'decimal' }),
        P('Lift – afmetingen (L × B × H in m)', 6, 24, 2, 2),
        CH('Dakluik aanwezig?', [['Aanwezig', 89], ['Niet aanwezig', 90], ['NVT', 91]]),
        P('Dakluik in veilige zone?', 6, 26, 2, 0, { prefix: 'In veilige zone?  ', options: ['Ja', 'Nee'] }),
        P('Dakluik – afmetingen (L × B × H in m)', 6, 26, 2, 2),
        JN('Is er slip-, stoot- en/of struikelgevaar op de route?', 92, 93),
        P('Slip-/stoot-/struikelgevaar – waar?', 6, 28, 2, 0, { prefix: 'Waar? ' }),
        R('Opmerkingen toegang', 6, 30, 0, 1),
      ] },
      { title: '3. Dakconstructie', items: [
        CH('Constructie', [['Beton in werk', 94], ['Breedplaat', 95], ['Cellenbeton', 96], ['Staal met plaatmat.', 97], ['Combinatie', 98], ['Staal met hout', 99], ['Hout', 100], ['Kanaalplaat', 101], ['Onbekend / niet duidelijk te zien', 102]], { multi: true, hint: 'Goed kijken of er d.m.v. poeren of losse fundatie geplaatst kan worden.' }),
        CH('Dakisolatie', [['Ja, wel drukvast', 103], ['Ja, niet drukvast', 104], ['Nee', 105]]),
        P('Dikte isolatie (i.v.m. hoogte poeren, mm)', 7, 1, 2, 0, { prefix: 'Dikte (i.v.m. hoogte poeren): ', suffix: ' mm', inputmode: 'decimal' }),
        CH('Plaats dakisolatie', [['Onder dakleer', 106], ['Bovenop (warm dak)', 107]]),
        CH('Bedekkingsmateriaal', [['Bitumen (APP/SBS)', 108], ['PVC', 109], ['Kunststof', 110], ['Rubber', 111], ['Zink/koper', 112], ['Anders', 113]], { multi: true }),
        P('Bedekkingsmateriaal – anders, namelijk', 7, 4, 1, 6),
        CH('Ballast', [['Grind', 114], ['Tegels', 115], ['Looppad(en)', 116], ['Geen', 117]], { multi: true }),
        R('Opmerkingen', 7, 6, 0, 1),
      ] },
      { title: '4. Doorvoeren', items: [
        CH('Dakdoorvoer nodig?', [['Wendyhouse', 118], ['Wandelstok / zwanenhals', 119], ['N.v.t.', 120]]),
        JN('Muurdoorvoer nodig? (Roxtec)', 121, 122),
        JN('Brandvertragend laten afdichten na montage kabels?', 123, 124),
        JN('Moet er geboord worden?', 125, 126),
        P('Boren – waarin?', 8, 3, 2, 0, { prefix: 'Waarin? ' }),
        JN('Is er asbest aanwezig op de locatie van de doorvoer?', 127, 128),
        P('Asbest – waar?', 8, 5, 2, 0, { prefix: 'Waar? ' }),
        R('Opmerkingen doorvoeren', 8, 7, 0, 0),
      ] },
      { title: '5.1 Apparatuur indoor', items: [
        P('Locatie / ruimte (verdieping / ruimtenummer)', 9, 0, 1),
        CH('Toegang naar ruimte (indoor)', [['Sleutel/pasje via landlord', 129], ['Anders', 130]], { hint: 'Schets maken waar apparatuur komt te staan, met afmetingen.' }),
        P('Toegang ruimte – anders, namelijk', 9, 2, 1, 2),
        P('Afmetingen ruimte – L', 9, 3, 1, 0, { prefix: 'L: ', inputmode: 'decimal' }),
        P('Afmetingen ruimte – B', 9, 3, 2, 0, { prefix: 'B: ', inputmode: 'decimal' }),
        P('Afmetingen ruimte – H', 9, 3, 3, 0, { prefix: 'H: ', inputmode: 'decimal' }),
        P('RBS-type (indoor)', 9, 4, 1),
        CH('Koeling', [['Niet aanwezig', 131], ['Nvt', 132], ['Ventilator', 133], ['Airco', 134]], { multi: true }),
        P('Airco – capaciteit', 9, 6, 2, 0, { prefix: 'Capaciteit: ' }), P('Airco – type', 9, 6, 2, 1, { prefix: 'Type: ' }), P('Airco – merk', 9, 6, 2, 2, { prefix: 'Merk: ' }),
        R('Opmerkingen (denk aan bv. temperatuur)', 9, 7, 0, 1),
      ] },
      { title: '5.2 Apparatuur outdoor', items: [
        P('RBS-type (outdoor)', 10, 0, 1),
        JN('Maaiveld: kan en mag er apparatuur binnen bestaand hekwerk bijgeplaatst worden? (aangeven op schets)', 135, 136),
        P('Uitbreiding opstelplaats RBS – L', 11, 1, 1, 0, { prefix: 'L: ', inputmode: 'decimal' }),
        P('Uitbreiding opstelplaats RBS – B', 11, 1, 2, 0, { prefix: 'B: ', inputmode: 'decimal' }),
        P('Hoeveel m² uitbreiding stelconplaten?', 11, 2, 1, 0, { suffix: ' m2', inputmode: 'decimal' }),
        P('Hoeveel m² uitbreiding bestrating?', 11, 3, 1, 0, { suffix: ' m2', inputmode: 'decimal' }),
        P('Uitbreiding hekwerk – L', 11, 4, 1, 0, { prefix: 'L: ', inputmode: 'decimal' }),
        P('Uitbreiding hekwerk – B', 11, 4, 2, 0, { prefix: 'B: ', inputmode: 'decimal' }),
        P('Uitbreiding hekwerk – H', 11, 4, 3, 0, { prefix: 'H: ', inputmode: 'decimal' }),
        JN('Hekwerkpoort nodig?', 137, 138),
        CB('Aanrijpaaltjes nodig? – Nee', 139, { hint: 'Wel nodig? Vul dan hieronder het aantal in.' }),
        P('Hoeveel aanrijpaaltjes?', 11, 7, 1, 0, { inputmode: 'numeric' }),
        JN('Dak: kan en mag er apparatuur op de draagconstructie geplaatst worden? (aangeven op schets)', 140, 141),
        JN('Dak: kan en mag er apparatuur op bestaande constructie (voormalig OLO) geplaatst worden? (schets + constructie inmeten)', 142, 143),
      ] },
      { title: '5.3 Tracé', items: [
        R('Indoor: door welke ruimtes loopt het tracé en hoe is de toegang daarnaartoe?', 13, 0, 0, 1),
        JN('Indoor: voldoende ruimte in de bestaande kabeldoorvoer?', 144, 145),
        P('Diameter doorvoer', 13, 2, 1),
        CH('Type doorvoer', [['Beton', 146], ['Hout', 147], ['Staal', 148], ['Steen', 149], ['Anders', 150]], { multi: true }),
        CB('Indoor: voldoende ruimte in de bestaande goten – Ja', 151),
        P('Indoor: hoeveel meter goot aanbrengen?', 13, 5, 1, 0, { suffix: ' m', inputmode: 'decimal' }),
        P('Indoor: type goot', 13, 5, 2, 0, { prefix: 'Type: ' }),
        P('Lengte indoor feedertracé (m)', 13, 6, 1, 0, { suffix: ' m', inputmode: 'decimal', hint: 'Teken de route in op de schets.' }),
        R('Indoor tracé – opmerkingen', 13, 8, 0, 1),
        R('Outdoor: beknopte beschrijving tracé', 14, 0, 0, 1),
        P('Outdoor: moeten er nieuwe voorzieningen geplaatst worden? Zo ja, welke?', 14, 1, 1),
        P('Outdoor: hoeveel meter goot?', 14, 2, 1, 0, { suffix: ' m', inputmode: 'decimal' }),
        P('Outdoor: type goot', 14, 2, 2, 0, { prefix: 'Type: ' }),
        ...[1, 2, 3, 4].map((sct) => P(`Lengte outdoor tracé DC/fiber – S${sct} (m)`, 14, 3, 1, sct - 1, { prefix: `S${sct} `, suffix: ' M', inputmode: 'decimal' })),
      ] },
      { title: '6. RAL-kleuren', items: [
        WGRID('RAL-kleuren', ['Van toepassing', 'RAL'], RAL.map((l, i) => ({
          label: l,
          fields: [CB(`RAL ${l} – van toepassing`, 152 + i, { short: 'Van toepassing' }), P(`RAL ${l}`, 15, i, 1, 0, { prefix: 'RAL: ', short: 'RAL' })],
        }))),
      ] },
      { title: '7. Kraan / hoogwerker', items: [
        R('Bijzonderheden', 16, 0, 0, 1),
        JN('Kraan nodig?', 164, 165),
        R('Mogelijke opstelplaatsen kraan (indien nodig met landlord bespreken)', 16, 2, 0, 2),
        JN('Hoogwerker nodig?', 166, 167),
        R('Mogelijke opstelplaatsen hoogwerker (indien nodig met landlord bespreken)', 16, 4, 0, 2),
        JN('Survey uitvoeren', 168, 169),
        JN('Rijplaten benodigd?', 170, 171),
        P('Rijplaten – lengte tracé (m)', 16, 6, 1, 1, { prefix: 'Lengte tracé: ', suffix: ' m', inputmode: 'decimal' }),
        JN('Toegang naar site mogelijk met kraan/hoogwerker?', 172, 173),
        JN('Inschatting: vergunning nodig?', 174, 175),
        JN('Verkeersregelaars nodig?', 176, 177),
        JN('Wegafzetting nodig?', 178, 179),
      ] },
      { title: '8. Elektravoorziening', items: [
        CH('Aansluiting', [['Eigen E-aansluiting', 180], ['Via tussenmeter van landlord', 181]]),
        R('Locatie K1 of tussenmeter landlord (aangeven op tekening/plattegrond en/of ruimtenummer)', 17, 1, 0, 1),
        CH('Uitvoering elektrawerkzaamheden', [['Huisaannemer', 182], ['DAEL', 183]]),
        P('Huisaannemer elektra – bedrijf', 19, 2, 1), P('Huisaannemer elektra – naam contactpersoon', 19, 1, 1),
        P('Huisaannemer elektra – e-mail', 19, 3, 1, 0, { inputmode: 'email' }), P('Huisaannemer elektra – telefoonnummer', 19, 4, 1, 0, { inputmode: 'tel' }),
        R('Huisaannemer elektra – opmerkingen', 19, 5, 0, 1),
        CH('8.1 Outdoor: van wie is het terrein?', [['Gemeente', 184], ['Privé', 185], ['Anders', 186]], { hint: 'Denk eraan: energiebedrijf graaft niet over privéterrein.' }),
        P('Terrein – anders, namelijk', 20, 0, 1, 3),
        CH('Ingraven voedingskabel (outdoor)', [['KLIC-melding', 187], ['Vergunning', 188]], { multi: true }),
        P('Lengte voedingskabel (m)', 20, 2, 1, 0, { suffix: ' M', inputmode: 'decimal' }),
        R('8.1 Outdoor – opmerkingen', 20, 3, 0, 1),
        R('8.2 Indoor: door welke ruimtes loopt het tracé en hoe is de toegang daarnaartoe?', 21, 0, 0, 1),
        JN('8.2 Indoor: voldoende ruimte in de bestaande kabeldoorvoer?', 189, 190),
        JN('8.2 Indoor: voldoende ruimte in de bestaande goten?', 191, 192),
        P('8.2 Hoeveel meter goot aanbrengen?', 21, 3, 1, 0, { suffix: ' m', inputmode: 'decimal' }),
        P('8.2 Type goot', 21, 3, 2, 0, { prefix: 'Type: ' }),
        P('8.2 Lengte indoor tracé (m)', 21, 4, 1, 0, { suffix: ' m', inputmode: 'decimal', hint: 'Teken de route in op de schets.' }),
        R('8.2 Indoor – opmerkingen', 21, 6, 0, 1),
        JN('8.3 Voldoende verlichting?', 193, 194),
        P('Zo nee: waar moet de verlichting geplaatst worden?', 22, 1, 2),
        P('Zo nee: waar moet de schakelaar geplaatst worden?', 22, 2, 2, 0, { hint: 'Verlichting intekenen op schets.' }),
      ] },
      { title: '9. Aarding / bliksembeveiliging', items: [
        CH('Type aarding', [['Aluminium', 195], ['Koper', 196], ['Staal', 197]]),
        P('Aantal aardputten (bij apparatuur op maaiveld, aangeven op schets)', 22 + 1, 1, 1, 0, { inputmode: 'numeric' }),
        CH('Uitvoering aarding/bliksembeveiligingswerkzaamheden', [['Huisaannemer', 198], ['DAEL', 199]]),
        P('Zakleidingen – aantal', 23, 3, 1, 0, { prefix: 'Aantal: ', inputmode: 'numeric' }),
        CH('Opmerkingen aarding/bliksembeveiliging – uitvoering', [['Huisaannemer', 200], ['DAEL', 201]]),
        P('Huisaannemer aarding – bedrijf', 23, 7, 1), P('Huisaannemer aarding – naam contactpersoon', 23, 6, 1),
        P('Huisaannemer aarding – e-mail', 23, 8, 1, 0, { inputmode: 'email' }), P('Huisaannemer aarding – telefoonnummer', 23, 9, 1, 0, { inputmode: 'tel' }),
        R('Huisaannemer aarding – opmerkingen', 23, 10, 0, 1),
      ] },
      { title: '10. Antennes', items: [
        JN('Zijn er vreemde antennes aanwezig?', 202, 203),
        CH('Zo ja, welke antennes of providers?', [['OLO', 204], ['Niet OLO', 205], ['FM-antennes', 206], ['Niet bekend – opvragen', 207]], { multi: true }),
        P('Vreemde antennes – namelijk', 24, 1, 2, 0, { prefix: 'Namelijk:  ' }),
        JN('Kans op te hoge blootstelling aan elektromagnetische velden?', 208, 209),
        JN('Uitgebreide metingen nodig? (indicatie)', 210, 211),
        JN('Is er outage nodig? (indicatie)', 212, 213),
        P('Outage – inschatting hoelang?', 24, 6, 1),
        R('Opmerkingen (geen andere antennes in de omgeving)', 24, 7, 0, 1),
      ] },
      { title: '11–12. Health & Safety Reworks & waarnemingen', items: [
        JN('Zijn alle veiligheidsmiddelen gekeurd?', 214, 215),
        P('Veiligheidsmiddelen gekeurd tot wanneer?', 25, 0, 0, 1),
        P('RI&E tijdens de TSS', 25, 1, 1, 0, { options: ['Groen', 'Geel', 'Oranje', 'Rood'] }),
        P('Toelichting RI&E (bij Oranje / Rood)', 25, 1, 2, 0, { prefix: 'Toelichting: ' }),
        JN('Zijn alle veiligheidsmiddelen gekeurd? (tweede vraag in het formulier)', 216, 217),
        JN('Moet de site vóór uitvoering van dit project aangepast worden?', 218, 219),
        P('Toelichting aanpassing site', 25, 3, 0, 1),
        JN('Persoonlijke waarnemingen genoteerd?', 220, 221),
      ] },
      { title: "13. Omschrijving Reworks & foto's", items: [
        R('Omschrijving Reworks', 27, 0, 0, 0, { rows: 6 }),
        { id: 'photos', type: 'photos', label: "Foto's", photos: { tables: [28, 29, 30, 31], photoRows: [0, 2], captionRows: [1, 3], maxW: 4600, maxH: 4200 },
          hint: "Komen in de fotovakken op de laatste pagina, met het bijschrift eronder. Bij meer dan 16 foto's worden extra vakken toegevoegd." },
      ] },
    ],
  };

  // ---------- onderaannemers uit het Monet-sjabloon als keuzelijst ----------
  const SUB_KEYS = ['name', 'visit', 'postal', 'postcode', 'town', 'phone', 'contact', 'contactPhone', 'email', 'supervisor', 'supervisorPhone', 'direct', 'directPhone', 'hs', 'hsPhone'];
  const setSub = (form, id, fill = {}) => {
    const f = form.sections.flatMap((sec) => sec.items.flatMap((it) => (it.grid ? it.rows.flatMap((r) => r.fields) : it.group ? it.fields : [it]))).find((x) => x && x.id === id);
    if (!f) throw new Error('Onbekend veld voor onderaannemer: ' + id);
    f.sub = { fill };
  };
  // Odido: onderaannemers 1–5 en nevenaannemer (kolommen van de contacttabellen)
  [402, 447].forEach((start) => [0, 1, 2].forEach((k) => {
    const fill = {};
    SUB_KEYS.forEach((key, row) => { if (row) fill[`ff${start + row * 3 + k}`] = key; });
    setSub(odido, `ff${start + k}`, fill);
  }));
  seq(0, 7).forEach((k) => setSub(odido, `ff${25 + k * 8}`));
  seq(0, 24).forEach((k) => setSub(odido, `ff${111 + k * 3}`));
  // KPN: huisaannemers dakbedekking, elektra en aarding
  [[6, 10], [18, 1], [22, 5]].forEach(([t, r]) => setSub(kpn, `p${t}_${r + 1}_1_0`, {
    [`p${t}_${r}_1_0`]: 'contact', [`p${t}_${r + 2}_1_0`]: 'email', [`p${t}_${r + 3}_1_0`]: 'contactPhone' }));
  // Cellnex: aannemer en huisaannemers
  setSub(cellnex, 'p3_5_1_0', { p3_0_1_0: 'contact', p3_2_1_0: ['visit', 'postcode', 'town'], p3_3_1_0: 'email', p3_4_1_0: 'contactPhone' });
  [[19, 1], [23, 6]].forEach(([t, r]) => setSub(cellnex, `p${t}_${r + 1}_1_0`, {
    [`p${t}_${r}_1_0`]: 'contact', [`p${t}_${r + 2}_1_0`]: 'email', [`p${t}_${r + 3}_1_0`]: 'contactPhone' }));

  window.FORMS = [vfBestaand, vfNieuwbouw, kpn, odido, cellnex];
  window.CLIENTS = [
    { id: 'vodafone', name: 'Vodafone', desc: 'TSS Report 4.20 (Excel)', forms: ['vf-bestaand', 'vf-nieuwbouw'] },
    { id: 'kpn', name: 'KPN', desc: 'RCT rapport – Nieuwbouw daklocatie (Word)', forms: ['kpn-dak'] },
    { id: 'odido', name: 'Odido', desc: 'V&G-plan Monet 6.2.1 – Rooftop (Word)', forms: ['odido-vg'] },
    { id: 'cellnex', name: 'Cellnex', desc: 'TSSR – Reworks (Word)', forms: ['cellnex-reworks'] },
  ];
})();
