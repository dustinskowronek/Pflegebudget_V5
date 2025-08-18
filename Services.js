/** =========================
 *  Services.js – serverseitige Logik (nur K_<Klient>-Blätter)
 * ========================= */
const Services = {

  /**
   * Legt für alle Einträge aus "Klienten" ein K_<Klient>-Blatt an (falls fehlend)
   * und schreibt die 3 Startzeilen:
   *  1) Neuanlage SGB XI 45b
   *  2) Neuanlage SGB XI 39
   *  3) Neuanlage SGB XI 45a
   *
   * Erwartete Spalten in "Klienten":
   *  'Klient','Pflegegrad',
   *  'SGB XI 45b - Start','SGB XI 45b - Leistungsmonat',
   *  'SGB XI 39 - Start','SGB XI 39 - Leistungsmonat',
   *  'SGB XI 45a','Angelegt Am'
   */
  syncAllClientSheets() {
    const rows = Repos.getClients(); // [{Klient:..., ...}, ...]
    rows.forEach(row => {
      const name     = String(row['Klient'] || '').trim();
      if (!name) return;

      let sh = Repos.getClientSheet(name);
      if (!sh) {
        sh = Repos.ensureClientSheet(name);
        // Kopfzeile sicherstellen
        if (sh.getLastRow() < 1) {
          sh.getRange(1,1,1,6).setValues([[
            'Rechnungsempfänger','Leistungsmonat','Beschreibung','SGB XI 45b','SGB XI 39','SGB XI 45a'
          ]]);
        }

        // Werte aus "Klienten"
        const s45bStart = Number(row['SGB XI 45b - Start'] || 0);
        const s39Start  = Number(row['SGB XI 39 - Start'] || 0);
        const s45aStart = Number(row['SGB XI 45a'] || 0);
        const ym45b     = String(row['SGB XI 45b - Leistungsmonat'] || '').trim(); // yyyy-MM
        const ym39      = String(row['SGB XI 39 - Leistungsmonat'] || '').trim();  // yyyy-MM
        const angelegt  = row['Angelegt Am'] ? new Date(row['Angelegt Am']) : new Date();
        const ym45a     = Utilities.formatDate(angelegt, Session.getScriptTimeZone()||'Europe/Berlin', 'yyyy-MM');

        // 1) Neuanlage 45b
        if (s45bStart > 0 && ym45b) {
          Repos.appendClientBooking({
            client: name,
            ym: ym45b,
            beschr: 'Neuanlage SGB XI 45b',
            s45b: s45bStart, s39: 0, s45a: 0
          });
        }
        // 2) Neuanlage 39
        if (s39Start > 0 && ym39) {
          Repos.appendClientBooking({
            client: name,
            ym: ym39,
            beschr: 'Neuanlage SGB XI 39',
            s45b: 0, s39: s39Start, s45a: 0
          });
        }
        // 3) Neuanlage 45a (Datum aus „Angelegt Am“ → yyyy-MM)
        if (s45aStart > 0) {
          Repos.appendClientBooking({
            client: name,
            ym: ym45a,
            beschr: 'Neuanlage SGB XI 45a',
            s45b: 0, s39: 0, s45a: s45aStart
          });
        }
      }
    });
  },

  /**
   * Manuelle Buchung als bequemer Wrapper.
   * UI nutzt aktuell `postManualBooking` (in code.gs) – diese hier ist kompatibel,
   * falls du sie direkt aufrufen willst.
   *
   * @param {Object} p {client, month, year, amount, art, type}
   *   art:  'SGB XI 39' | 'SGB XI 45b' | 'SGB XI 45a'
   *   type: 'Buchung SGB XI 39' | 'Neuanlage SGB XI 39' | 'Folgeanlage SGB XI 39' | ...
   */
  createManualBooking(p){
    const client = String(p.client||'').trim();
    const month  = String(p.month||'').padStart(2,'0');
    const year   = String(p.year||'');
    const amount = Number(p.amount||0);
    const art    = String(p.art||'').trim();
    const type   = String(p.type||'').trim() || ('Buchung ' + art);

    if (!client) throw new Error('Kein Klient angegeben');
    if (!month || !year) throw new Error('Monat/Jahr fehlen');
    if (!amount || isNaN(amount)) throw new Error('Ungültiger Betrag');
    if (!/^SGB XI (39|45a|45b)$/.test(art)) throw new Error('Unbekannte Leistungsart: '+art);

    const ym = `${year}-${month}`;
    let s45b=0,s39=0,s45a=0;
    if (type.indexOf('SGB XI 45b')>-1) s45b = amount;
    else if (type.indexOf('SGB XI 39')>-1) s39 = amount;
    else if (type.indexOf('SGB XI 45a')>-1) s45a = amount;

    // K_Sheet sicherstellen
    let sh = Repos.getClientSheet(client);
    if (!sh) {
      sh = Repos.ensureClientSheet(client);
      if (sh.getLastRow() < 1) {
        sh.getRange(1,1,1,6).setValues([[
          'Rechnungsempfänger','Leistungsmonat','Beschreibung','SGB XI 45b','SGB XI 39','SGB XI 45a'
        ]]);
      }
    }

    Repos.appendClientBooking({ client, ym, beschr: type, s45b, s39, s45a });
    return { ok:true, client, ym, art, type, amount };
  },

  /**
   * Optionale Routine: Monatsgutschriften §45b (z. B. 131 €) + Jahresstart §39 (Januar)
   * Arbeitet NUR mit K-Blättern. Stammdaten holt sie aus "Stammdaten".
   *
   * Hinweis: Wenn du die 45b-Monatsgutschriften per Simulation in getBudgetVM berechnest,
   * brauchst du diese Funktion nicht regelmäßig auszuführen.
   */
  postMonthlyCreditsWithResets(dateOpt) {
    const today    = dateOpt || new Date();
    const month    = today.getMonth() + 1; // 1..12
    const year     = today.getFullYear();
    const ym       = `${year}-${String(month).padStart(2,'0')}`;

    const clients  = Repos.getClients();
    const sdByPg   = Repos.getStammdatenByPG(); // Map(pg -> {s45b, s39, s45a})

    clients.forEach(row => {
      const name = String(row['Klient']||'').trim();
      if (!name) return;
      const pg = String(row['Pflegegrad']||'').trim();
      const sd = sdByPg.get(pg) || { s45b:0, s39:0, s45a:0 };

      let sh = Repos.getClientSheet(name);
      if (!sh) {
        sh = Repos.ensureClientSheet(name);
        if (sh.getLastRow() < 1) {
          sh.getRange(1,1,1,6).setValues([[
            'Rechnungsempfänger','Leistungsmonat','Beschreibung','SGB XI 45b','SGB XI 39','SGB XI 45a'
          ]]);
        }
      }

      // §45b: monatlich gutschreiben (falls gewünscht)
      if (sd.s45b) {
        Repos.appendClientBooking({
          client: name, ym,
          beschr: 'Monatsgutschrift §45b',
          s45b: sd.s45b, s39: 0, s45a: 0
        });
      }

      // §39: nur im Januar das Jahresbudget als Startgutschrift (falls gepflegt)
      if (sd.s39 && month === 1) {
        Repos.appendClientBooking({
          client: name, ym,
          beschr: 'Jahresstart §39',
          s45b: 0, s39: sd.s39, s45a: 0
        });
      }
    });

    return { ok:true, monthKey: ym };
  },

  /**
   * Aggregation fürs Budget-Panel (Proxy auf deine getBudgetVM aus code.gs)
   * So vermeidest du doppelte Logik.
   */
  getBudgetViewModel({ client, topf, year }) {
    return getBudgetVM({ client, topf, year });
  }, // <— WICHTIG: Komma hier!

  /**
   * Kürzt einen Namen auf max Zeichen (Standard 15) mit „…“.
   * Rein UI-Helfer – Logik bleibt in code.gs.
   */
  shorten(text, maxLen = 15) {
    const s = String(text || '');
    return s.length > maxLen ? s.slice(0, maxLen - 1) + '…' : s;
  },

  /**
   * Holt das Budget-VM aus code.gs und liefert nützliche
   * Label-Infos für die „Rechnung an …“-Spalte zurück.
   * (code.gs liefert payers[] sowie invoiceLabel / invoiceLabelShort mit)
   */
  getBudgetWithPayers({ client, topf, year }) {
    const vm = getBudgetVM({ client, topf, year }); // ruft code.gs
    // Defensive Defaults, falls ältere code.gs-Version noch kein Label schickt:
    const payers = vm.payers || [];
    const label  = vm.invoiceLabel || (payers.length ? `Rechnung an ${payers.join(' / ')}` : 'Rechnung an [Kasse]');
    const short  = vm.invoiceLabelShort || Services.shorten(label, 15);
    return { ...vm, payers, invoiceLabel: label, invoiceLabelShort: short };
  }
};