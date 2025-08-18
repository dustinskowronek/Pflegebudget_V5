/** Constants.js — zentrale Konfiguration, NUR HIER definieren! */
if (typeof CFG === 'undefined') { var CFG = {}; }

CFG.SHEETS = {
  KLIENTEN:   'Klienten',
  BUCHUNGEN:  'Buchungen',
  STAMMDATEN: 'Stammdaten',
  K_TEMPLATE: 'K_TEMPLATE' // Vorlage für Klientenblätter
};

CFG.COLUMNS = {
  // Klienten
  CLIENTS_HEADER: [
    'Klient',
    'Pflegegrad',
    'SGB XI 45b - Start',
    'SGB XI 45b - Leistungsmonat',
    'SGB XI 39 - Start',
    'SGB XI 39 - Leistungsmonat',
    'SGB XI 45a',
    'Angelegt Am'
  ],
  // Buchungen
  BOOKINGS_HEADER: [
    'Zeitstempel','Klient','Leistungsmonat','Leistungsart','Betrag'
  ],
  // Stammdaten
  STAMMDATEN_HEADER: [
    'Pflegegrad','SGB XI 45b','SGB XI 39','Sachleistung','SGB XI 45a'
  ],
  // Klientenblatt (Template)
  KLIENT_SHEET_HEADER: [
    'Rechnungsempfänger','Leistungsmonat','Beschreibung','SGB XI 45b','SGB XI 39','SGB XI 45a'
  ]
};

CFG.LEISTUNGSART = {
  UW: 'SGB XI 45b', // Umwidmung/Entlastungsbetrag (monatl.)
  VP: 'SGB XI 39',  // Verhinderungspflege (jährl.)
  EB: 'SGB XI 45a'  // ggf. separat geführt
};

CFG.UI = {
  PANEL_TITLE: 'Pflegebudget'
};