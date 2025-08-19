/** =========================================================================
 *  code.gs  – Menü, Panel, Backend (nur K_<Klient>-Blätter)
 *  Quelle für alle Zahlen: Klientenblatt K_<Name> mit Spalten:
 *  A:Rechnungsempfänger  B:Leistungsmonat  C:Beschreibung  D:SGB XI 45b
 *  E:SGB XI 39           F:SGB XI 45a (Umwidmung / Leistung aus 45a)
 *  G:Rechnungsnummer     H:Limit 45a €/Monat  (Konfigurations-Spalte – NICHT H1)
 *  -------------------------------------------------------------------------
 *  Klassifikation per Beschreibung:
 *    Neuanlage SGB XI 45b     -> Startgutschrift 45b
 *    Neuanlage SGB XI 39      -> Startgutschrift 39
 *    Folgeanlage SGB XI 39    -> Startgutschrift 39 (Folgejahr)
 *    Konfiguration 45a Limit  -> Limit-Zeile für 45a (monatlicher Deckel)
 *    Buchung SGB XI 45b       -> Ausgabe 45b
 *    Buchung SGB XI 39        -> Ausgabe 39
 *    Buchung SGB XI 45a       -> Leistung aus 45a (monatlich gedeckelt)
 *  ========================================================================= */

/***** MENÜ *****/
function onOpen() {
  try {
    if (Repos.ensureMainSheets) Repos.ensureMainSheets();
    ensureClientSheetHeadersAndMigrate_(); // Header setzen + evtl. H1->Konfig-Zeile migrieren
  } catch (e) {
    Logger.log(e);
  }
  SpreadsheetApp.getUi()
    .createMenu('Pflegebudget')
    .addItem('Setup prüfen', 'setupCheck_')
    .addItem('Panel öffnen', 'openPanel_')
    .addItem('Neues Klientenblatt (aus markierter Zeile)', 'createClientSheetFromActiveRow')
    .addToUi();
}

function setupCheck_() {
  if (Repos.ensureMainSheets) Repos.ensureMainSheets();
  ensureClientSheetHeadersAndMigrate_();
  SpreadsheetApp.getActive().toast('Setup geprüft.', 'Pflegebudget', 3);
}

/***** PANEL *****/
function openPanel_() {
  var tpl = HtmlService.createTemplateFromFile('Panel');
  var html = tpl.evaluate()
    .setTitle('Pflegebudget')
    .setWidth(1100)
    .setHeight(720)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  SpreadsheetApp.getUi().showModalDialog(html, 'Pflegebudget');
}
function include(filename){ return HtmlService.createHtmlOutputFromFile(filename).getContent(); }

/** Header sicherstellen + MIGRATION:
 *  - setzt Kopfzeile A1:H1 korrekt
 *  - wenn in H1 fälschlich ein numerischer Wert steht, wird er als
 *    „Konfiguration 45a Limit“-Zeile ans Ende migriert, H1 geleert.
 */
function ensureClientSheetHeadersAndMigrate_() {
  var ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets().filter(sh => /^K_/.test(sh.getName()));
  var wanted = [
    'Rechnungsempfänger','Leistungsmonat','Beschreibung',
    'SGB XI 45b','SGB XI 39','SGB XI 45a','Rechnungsnummer','Limit 45a €/Monat'
  ];

  sheets.forEach(sh => {
    var lastCol = Math.max(8, sh.getLastColumn() || 0);
    if (sh.getLastRow() === 0) {
      sh.getRange(1,1,1,8).setValues([wanted]);
      return;
    }
    var header = sh.getRange(1,1,1,lastCol).getValues()[0];

    // Header korrigieren, falls nötig
    if (header.slice(0,8).join('|') !== wanted.join('|')) {
      sh.getRange(1,1,1,8).setValues([wanted]);
    }

    // MIGRATION: Falls H1 eine Zahl enthält, in Konfig-Zeile verschieben
    var h1 = sh.getRange(1,8).getValue();
    if (typeof h1 === 'number' && !isNaN(h1) && h1 > 0) {
      var clientName = sh.getName().replace(/^K_/, '');
      sh.appendRow([clientName, '', 'Konfiguration 45a Limit', 0, 0, 0, '', h1]);
      sh.getRange(1,8).clearContent();
    }
  });
}

/***** Panel-API *****/

/** Liste der Klienten (für Datalist) */
function listClients() {
  if (Repos.listClients) return Repos.listClients();
  var sh = SpreadsheetApp.getActive().getSheetByName('Klienten');
  if (!sh) return [];
  var rng = sh.getRange(2,1,Math.max(0, sh.getLastRow()-1),1);
  return rng.getValues().flat().filter(v => v && String(v).trim().length);
}

/** Manuelle Buchung / Folgeanlage / 45a-Konfig
 * payload: {
 *   client, month:'01'..'12', year:'2025', amount:number,
 *   art:'SGB XI 45b'|'SGB XI 39'|'SGB XI 45a',
 *   kind:'Buchung'|'Folgeanlage'|'Konfiguration',
 *   invoiceNo?: string
 * }
 */
function postManualBooking(payload) {
  var client = payload && payload.client;
  var month  = payload && payload.month;
  var year   = payload && payload.year;
  var amount = Number(payload && payload.amount || 0);
  var art    = String(payload && payload.art || 'SGB XI 45b');
  var kind   = String(payload && payload.kind || 'Buchung');
  var invoiceNo = String(payload && payload.invoiceNo || '').trim();

  if (!client) throw new Error('Kein Klient gewählt.');
  if (!month || !year) throw new Error('Leistungsmonat/Jahr fehlt.');

  var ym = year + '-' + String(month).padStart(2,'0');

  // Sonderfall: 45a-Konfiguration (Limit setzen/ändern)
  if (art === 'SGB XI 45a' && kind === 'Konfiguration') {
    var shCfg = SpreadsheetApp.getActive().getSheetByName('K_' + client);
    if (!shCfg) throw new Error('Klientenblatt K_' + client + ' nicht gefunden.');
    shCfg.appendRow([client, ym, 'Konfiguration 45a Limit', 0, 0, 0, invoiceNo, amount]);
    return { status:'ok' };
  }

  // Normale Buchungen / Folgeanlage 39
  if (!amount || isNaN(amount)) throw new Error('Betrag ungültig.');

  var s45b=0, s39=0, s45a=0, beschr='';
  if (kind === 'Folgeanlage' && art === 'SGB XI 39') {
    s39 = amount; beschr = 'Folgeanlage SGB XI 39';
  } else if (art === 'SGB XI 45b') {
    s45b = amount; beschr = 'Buchung SGB XI 45b';
  } else if (art === 'SGB XI 39') {
    s39  = amount; beschr = 'Buchung SGB XI 39';
  } else if (art === 'SGB XI 45a') {
    s45a = amount; beschr = 'Buchung SGB XI 45a';
  } else {
    throw new Error('Unbekannter Topf: ' + art);
  }

  var sh = SpreadsheetApp.getActive().getSheetByName('K_' + client);
  if (!sh) throw new Error('Klientenblatt K_' + client + ' nicht gefunden.');
  sh.appendRow([client, ym, beschr, s45b, s39, s45a, invoiceNo, '']); // H nur bei „Konfiguration“
  return { status:'ok' };
}

/** Format-/Helper */
function _toYYYY_MM_(v){
  if (v instanceof Date) return v.getFullYear() + '-' + String(v.getMonth()+1).padStart(2,'0');
  var s = String(v||'').trim();
  var m = s.match(/^(\d{4})[-/.](\d{1,2})$/);
  if (m) return m[1] + '-' + m[2].padStart(2,'0');
  var d = new Date(s); if (!isNaN(d.getTime())) return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0');
  return '';
}
function _ymToDate_(ym){ var y=+ym.slice(0,4), m=+ym.slice(5,7)-1; return new Date(y,m,1); }
function _containsAny_(t, arr){ t=String(t||''); for (var i=0;i<arr.length;i++) if (t.indexOf(arr[i])>-1) return true; return false; }

/** Stammdaten lesen: Pflegegrad → Sachleistung/Monat (für 45a = 40%) */
function _readStammdaten_() {
  var sh = SpreadsheetApp.getActive().getSheetByName('Stammdaten');
  if (!sh) return {};
  var last = sh.getLastRow();
  if (last < 2) return {};
  var vals = sh.getRange(2,1,last-1, sh.getLastColumn()).getValues();
  // Erwartete Spalten: Pflegegrad | SGB XI 45b | SGB XI 39 | Sachleistung | (optional: SGB XI 45a)
  var map = {};
  vals.forEach(function(r){
    var pg = String(r[0]||'').trim();
    var sach = Number(r[3]||0); // „Sachleistung“ an 4. Stelle
    if (pg) map[pg] = { sach: sach, max45a: sach * 0.40 };
  });
  return map;
}

/** Pflegegrad eines Klienten holen (aus Tab „Klienten“) */
function _readPflegegradOfClient_(client) {
  var sh = SpreadsheetApp.getActive().getSheetByName('Klienten');
  if (!sh) return '';
  var last = sh.getLastRow();
  if (last < 2) return '';
  var vals = sh.getRange(2,1,last-1, 2).getValues(); // A:Klient B:Pflegegrad
  for (var i=0;i<vals.length;i++){
    var name = String(vals[i][0]||'').trim();
    if (name === client) return String(vals[i][1]||'').trim();
  }
  return '';
}

/** NEU: 45a-Autolimit bestimmen (wenn kein explizites Limit angegeben) */
function _compute45aLimit_(pflegegrad, limitInput) {
  var lim = Number(limitInput || 0);
  if (lim > 0) return +lim.toFixed(2);
  var sd = _readStammdaten_();
  return (sd[pflegegrad] && sd[pflegegrad].max45a) ? +sd[pflegegrad].max45a.toFixed(2) : 0;
}

/** ViewModel für Budget-Tabelle */
function getBudgetVM(args) {
  args = args || {};
  var client = args.client || '';
  var topf   = args.topf   || 'SGB XI 45b'; // 'SGB XI 39' | 'SGB XI 45b' | 'SGB XI 45a'
  var year   = String(args.year || new Date().getFullYear());

  if (!client) {
    var li = listClients();
    if (!li.length) throw new Error('Kein Klient vorhanden. Bitte zuerst in "Klienten" erfassen.');
    client = li[0];
  }

  var sh = SpreadsheetApp.getActive().getSheetByName('K_' + client);
  if (!sh) return { client: client, topf: topf, year: year, rows: [] };

  var lastRow = sh.getLastRow();
  var vals = lastRow >= 2 ? sh.getRange(1,1,lastRow,Math.max(8, sh.getLastColumn())).getValues() : [];

  // Spalten (1-basiert -> -1)
  var COL_EMP = 1, COL_MONAT = 2, COL_DESC = 3, COL_45B = 4, COL_39 = 5, COL_45A = 6, COL_INVNO = 7, COL_LIM45A = 8;

  // §45b (131€/Monat ab Folgemonat, Verfall Vorjahr 01.07.)
  if (topf === 'SGB XI 45b') return compute45bVM_(vals, year, client);

  // §45a: monatlicher Topf = (jüngstes Konfig-Limit) oder Auto-40%, kein Übertrag
  if (topf === 'SGB XI 45a') {
    var pg = _readPflegegradOfClient_(client);
    var sd = _readStammdaten_();
    var autoLimit = (sd[pg] && sd[pg].max45a) ? sd[pg].max45a : 0;

    // jüngste Konfig-Zeile mit H gesetzt
    var override = null;
    for (var r = 2; r <= (vals.length||0); r++){
      var lim = vals[r-1][COL_LIM45A-1];
      var desc = String(vals[r-1][COL_DESC-1]||'');
      if (lim !== '' && lim != null && desc.indexOf('Konfiguration 45a Limit') > -1) {
        override = Number(lim)||0; // letzter gewinnt
      } else if (lim !== '' && lim != null) {
        // Fallback: allgemein letzter nicht-leerer H-Wert
        override = Number(lim)||0;
      }
    }
    var monthlyCap = (override != null && override > 0) ? override : autoLimit;

    // Buchungen & Empfänger sammeln
    var invoiceByMonth = {}, copayByMonth = {}, payerNamesSet = {};
    for (var r2 = 2; r2 <= (vals.length||0); r2++){
      var row2 = vals[r2-1];
      var ym2  = _toYYYY_MM_(row2[COL_MONAT-1]);
      if (!ym2 || ym2.slice(0,4) !== year) continue;
      var desc2= String(row2[COL_DESC-1]||'');
      var vA   = Number(row2[COL_45A-1]||0);
      var empf = String(row2[COL_EMP-1]||'').trim();
      var inv  = String(row2[COL_INVNO-1]||'').trim();

      if (vA>0 && desc2.indexOf('Buchung SGB XI 45a')>-1) {
        if (inv && empf === client) {
          copayByMonth[ym2] = (copayByMonth[ym2]||0) + vA; // privat, nicht budgetwirksam
        } else {
          invoiceByMonth[ym2] = (invoiceByMonth[ym2]||0) + vA;
          if (empf && empf !== client) payerNamesSet[empf] = true;
        }
      }
    }
    var payerNames = Object.keys(payerNamesSet);
    var error = '', warning = '', payerLabel = '[Kasse]';
    if (payerNames.length > 2)       error   = 'Es sind mehr als 2 verschiedene Rechnungsempfänger erfasst: ' + payerNames.join(', ');
    else if (payerNames.length === 2) warning = 'Es sind 2 Rechnungsempfänger erfasst: ' + payerNames.join(' / ');
    else if (payerNames.length === 1) payerLabel = payerNames[0];

    var rowsA = [];
    for (var m=1;m<=12;m++){
      var ymKey = year + '-' + String(m).padStart(2,'0');
      var start = monthlyCap;
      var invoice = invoiceByMonth[ymKey] || 0;
      var carry = Math.max(0, start - invoice); // privat mindert Budget NICHT
      rowsA.push({
        mon: ymKey,
        start: +start.toFixed(2),
        invoice: +invoice.toFixed(2),
        carry: +carry.toFixed(2),
        copay: +((copayByMonth[ymKey]||0).toFixed(2))
      });
    }
    return { client: client, topf: topf, year: year, payerLabel: payerLabel, warning: warning, error: error, rows: rowsA };
  }

  // §39: Jahresanlage je Jahr, kein Übertrag über 31.12.
  var keyCol = COL_39;
  var startNeedles = ['Neuanlage SGB XI 39', 'Folgeanlage SGB XI 39'];
  var bookNeedles  = ['Buchung SGB XI 39'];

  var creditByMonth = {}, invoiceByMonth = {}, copayByMonth = {}, payerNamesSet = {};

  for (var r = 2; r <= (vals.length||0); r++) {
    var row = vals[r-1];
    var ym  = _toYYYY_MM_(row[COL_MONAT-1]);
    if (!ym || ym.slice(0,4) !== year) continue;

    var desc  = String(row[COL_DESC-1] || '');
    var vTopf = Number(row[keyCol-1] || 0);
    var empf  = String(row[COL_EMP-1]||'').trim();
    var inv   = String(row[COL_INVNO-1]||'').trim();

    if (!vTopf) continue;

    if (_containsAny_(desc, startNeedles)) {
      creditByMonth[ym] = (creditByMonth[ym] || 0) + vTopf;
    } else if (_containsAny_(desc, bookNeedles)) {
      if (inv && empf === client) {
        copayByMonth[ym] = (copayByMonth[ym] || 0) + vTopf;
      } else {
        invoiceByMonth[ym] = (invoiceByMonth[ym] || 0) + vTopf;
        if (empf && empf !== client) payerNamesSet[empf] = true;
      }
    }
  }

  var payerNames = Object.keys(payerNamesSet);
  var error = '', warning = '', payerLabel = '[Kasse]';
  if (payerNames.length > 2)       error   = 'Es sind mehr als 2 verschiedene Rechnungsempfänger erfasst: ' + payerNames.join(', ');
  else if (payerNames.length === 2) warning = 'Es sind 2 Rechnungsempfänger erfasst: ' + payerNames.join(' / ');
  else if (payerNames.length === 1) payerLabel = payerNames[0];

  var carryPrev = 0;
  var rows = [];
  for (var i = 1; i <= 12; i++) {
    var ymKey  = year + '-' + String(i).padStart(2,'0');
    var credit = creditByMonth[ymKey] || 0;
    var start  = carryPrev + credit;
    var invoice= invoiceByMonth[ymKey] || 0;
    var carry  = Math.max(0, start - invoice); // privat mindert Budget NICHT
    rows.push({
      mon: ymKey,
      start: +start.toFixed(2),
      invoice: +invoice.toFixed(2),
      carry: +carry.toFixed(2),
      copay: +((copayByMonth[ymKey]||0).toFixed(2))
    });
    carryPrev = carry;
  }
  return { client: client, topf: topf, year: year, payerLabel: payerLabel, warning: warning, error: error, rows: rows };
}

/* ---------- Helper: §45b-Simulation inkl. Empfänger-/Copay-Logik ---------- */
function compute45bVM_(vals, viewYear, client) {
  var COL_EMP   = 1, COL_MONAT = 2, COL_DESC = 3, COL_45B = 4, COL_INVNO = 7;
  var ENTLA_PRO_MONAT = 131;
  var MAX_YEAR = 2027;

  var startYM = null, startAmt = 0;

  // Sammelcontainer
  var invoiceByMonth = {};   // echte Rechnungen an Dritte (Budget-relevant)
  var copayByMonth   = {};   // private Zuzahlungen (nicht budgetwirksam)
  var payerNamesSet  = {};   // unterschiedliche externe Rechnungsempfänger im Jahr

  // Datendurchlauf
  for (var r = 2; r <= (vals.length||0); r++) {
    var row = vals[r-1];
    var ym  = _toYYYY_MM_(row[COL_MONAT-1]);
    if (!ym) continue;

    var y = ym.slice(0,4);
    var desc = String(row[COL_DESC-1] || '');
    var v45b = Number(row[COL_45B-1] || 0);
    var empf = String(row[COL_EMP-1] || '').trim();
    var inv  = String(row[COL_INVNO-1] || '').trim();

    // Neuanlage finden (früheste)
    if (desc.indexOf('Neuanlage SGB XI 45b') > -1 && v45b > 0) {
      if (!startYM || ym < startYM) { startYM = ym; startAmt = v45b; }
      continue;
    }

    // Buchungen klassifizieren nur fürs Betrachtungsjahr
    if (y === String(viewYear) && v45b > 0 && desc.indexOf('Buchung SGB XI 45b') > -1) {
      if (inv && empf && empf === client) {
        // Privat: zählt nicht gegen Budget
        copayByMonth[ym] = (copayByMonth[ym] || 0) + v45b;
      } else {
        // Rechnung an Dritte (Budget)
        invoiceByMonth[ym] = (invoiceByMonth[ym] || 0) + v45b;
        if (empf && empf !== client) payerNamesSet[empf] = true;
      }
    }
  }

  if (!startYM) {
    // keine Neuanlage → alles 0 zurück
    return {
      client: client, topf: 'SGB XI 45b', year: String(viewYear),
      payerLabel: '[Kasse]', warning: '', error: '',
      rows: Array.from({length:12}, function(_,i){
        var m=String(i+1).padStart(2,'0'); return {mon:viewYear+'-'+m,start:0,invoice:0,carry:0,copay:0};
      })
    };
  }

  // Empfänger-Auswertung
  var payerNames = Object.keys(payerNamesSet);
  var error = '', warning = '', payerLabel = '[Kasse]';
  if (payerNames.length > 2) {
    error = 'Es sind mehr als 2 verschiedene Rechnungsempfänger erfasst: ' + payerNames.join(', ');
  } else if (payerNames.length === 2) {
    warning = 'Es sind 2 Rechnungsempfänger erfasst: ' + payerNames.join(' / ');
  } else if (payerNames.length === 1) {
    payerLabel = payerNames[0];
  }

  // Simulation (wie gehabt)
  var startDate = _ymToDate_(startYM);
  var simStart  = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
  if (Number(viewYear) > MAX_YEAR) viewYear = String(MAX_YEAR);

  var pools = { year: simStart.getFullYear(), prev: 0, curr: 0 };
  pools.curr = startAmt; // Start gehört zum Startjahr
  var accrualStart = new Date(simStart.getFullYear(), simStart.getMonth()+1, 1);

  var startByMonth = {}, carryByMonth = {};
  var cur = new Date(simStart);

  while (cur <= new Date(viewYear, 11, 1)) {
    var curY = cur.getFullYear(), curM = cur.getMonth()+1;
    var ymKey = curY + '-' + String(curM).padStart(2,'0');

    if (curY !== pools.year) { pools.prev = pools.curr; pools.curr = 0; pools.year = curY; }
    if (cur >= accrualStart) pools.curr += ENTLA_PRO_MONAT;
    if (curM === 7) pools.prev = 0; // 01.07.: Vorjahr verfällt

    var startThis = pools.prev + pools.curr;

    // echte Budget-Rechnung im Monat
    var invAmt = Number(invoiceByMonth[ymKey] || 0);

    // zuerst prev, dann curr abbauen
    var rest = invAmt;
    var takePrev = Math.min(pools.prev, rest); pools.prev -= takePrev; rest -= takePrev;
    var takeCurr = Math.min(pools.curr, rest); pools.curr -= takeCurr; rest -= takeCurr;

    var carryThis = pools.prev + pools.curr;

    startByMonth[ymKey] = startThis;
    carryByMonth[ymKey] = carryThis;

    cur = new Date(curY, cur.getMonth()+1, 1);
  }

  // Monatszeilen erzeugen (inkl. copay)
  var rows = [];
  for (var i = 1; i <= 12; i++) {
    var ym = String(viewYear) + '-' + String(i).padStart(2,'0');
    rows.push({
      mon: ym,
      start: +(startByMonth[ym] || 0).toFixed(2),
      invoice: +((invoiceByMonth[ym] || 0)).toFixed(2),
      carry: +(carryByMonth[ym] || 0).toFixed(2),
      copay: +((copayByMonth[ym] || 0)).toFixed(2)
    });
  }

  return {
    client: client,
    topf: 'SGB XI 45b',
    year: String(viewYear),
    payerLabel: payerLabel,
    warning: warning,
    error: error,
    rows: rows
  };
}

/** Klientenblatt aus markierter „Klienten“-Zeile erzeugen/füllen (Neuanlage) */
function createClientSheetFromActiveRow() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Klienten');
  if (!sh) throw new Error('Tabelle "Klienten" nicht gefunden.');
  var r = sh.getActiveCell() ? sh.getActiveCell().getRow() : 2;
  if (r < 2) throw new Error('Bitte eine Datenzeile in "Klienten" markieren.');

  var head = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var vals = sh.getRange(r,1,1,sh.getLastColumn()).getValues()[0];
  var idx = (name)=> head.indexOf(name);

  var client  = String(vals[idx('Klient')] || '').trim();
  var pg      = String(vals[idx('Pflegegrad')] || '').trim();
  var start45b= Number(vals[idx('SGB XI 45b - Start')] || 0);
  var lm45b   = vals[idx('SGB XI 45b - Leistungsmonat')];
  var start39 = Number(vals[idx('SGB XI 39 - Start')] || 0);
  var lm39    = vals[idx('SGB XI 39 - Leistungsmonat')];

  // 45a-Konfig aus „Klienten“
  var umwAktiv = String(vals[idx('45a Umwidmung aktiv')] || '').trim().toLowerCase(); // 'ja'|'nein'|''
  var limit45a = vals[idx('45a Limit €/Monat')]; // Zahl oder leer
  var legacyIndex = idx('SGB XI 45a'); // falls Altspalte noch existiert
  var legacyLimit = legacyIndex > -1 ? vals[legacyIndex] : '';

  if (!client) throw new Error('Klientenname fehlt.');

  var k = Repos.ensureClientSheet(client);
  ensureClientSheetHeadersAndMigrate_(); // sicherheitshalber Header/Migration

  // Wenn Umwidmung aktiv → Limit ermitteln (leer ⇒ Auto 40%) + Konfig-Zeile anhängen
  if (umwAktiv === 'ja' || umwAktiv === 'true' || umwAktiv === '1') {
    var chosen = _compute45aLimit_(pg, (limit45a === '' || limit45a == null) ? legacyLimit : limit45a);
    // ins „Klienten“-Sheet zurückschreiben, falls dort leer war
    if ((limit45a === '' || limit45a == null) && idx('45a Limit €/Monat') > -1) {
      sh.getRange(r, idx('45a Limit €/Monat')+1).setValue(chosen);
    }
    // Konfig-Zeile im K_-Blatt
    var todayYM = _toYYYY_MM_(new Date());
    k.appendRow([client, todayYM, 'Konfiguration 45a Limit', 0, 0, 0, '', chosen]);
  }

  // Startzeilen 45b / 39
  if (start45b) {
    var ym45b = _toYYYY_MM_(lm45b);
    k.appendRow([client, ym45b || '', 'Neuanlage SGB XI 45b', start45b, 0, 0, '', '']);
  }
  if (start39) {
    var ym39 = _toYYYY_MM_(lm39);
    k.appendRow([client, ym39 || '', 'Neuanlage SGB XI 39', 0, start39, 0, '', '']);
  }

  SpreadsheetApp.getActive().toast('K_' + client + ' aktualisiert (inkl. 45a-Konfiguration).', 'Pflegebudget', 3);
}

/** Meldungen für Panel (Topf/Jahr) aggregieren */
function listPanelMessages(args) {
  args = args || {};
  var topf = args.topf || 'SGB XI 39';
  var year = String(args.year || new Date().getFullYear());

  var out = [];
  var clients = listClients() || [];
  clients.forEach(function(c){
    try {
      var vm = getBudgetVM({ client: c, topf: topf, year: year });
      if (vm && vm.error)   out.push({ client: c, level:'error',   text: vm.error });
      if (vm && vm.warning) out.push({ client: c, level:'warning', text: vm.warning });
    } catch (e) {
      out.push({ client: c, level:'error', text: 'Fehler beim Laden: ' + (e.message || e) });
    }
  });
  return out;
}

/** 45a-Konfig lesbar machen (für UI/Debug) */
function get45aConfig(client) {
  var ss = SpreadsheetApp.getActive();
  var k = ss.getSheetByName('K_'+client);
  if (!k) return { active:false, limit:null, mode:'auto' };

  var last = k.getLastRow();
  if (last < 2) return { active:false, limit:null, mode:'auto' };

  var vals = k.getRange(2,1,last-1,8).getValues();
  var limit = null;
  for (var i=0;i<vals.length;i++){
    var desc = String(vals[i][2]||'');
    var lim  = vals[i][7];
    if (desc.indexOf('Konfiguration 45a Limit')>-1 && lim !== '' && lim != null) {
      limit = Number(lim)||0; // letzter gewinnt
    } else if (lim !== '' && lim != null) {
      limit = Number(lim)||0; // Fallback: letzter nicht-leerer
    }
  }
  return { active: limit!=null && limit>0, limit: limit, mode: (limit!=null?'fixed':'auto') };
}