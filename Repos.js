const Repos = {
  ensureMainSheets() {
    const ss = SpreadsheetApp.getActive();
    const shK = ss.getSheetByName('Klienten') || ss.insertSheet('Klienten');
    const wantK = [
      'Klient','Pflegegrad',
      'SGB XI 45b - Start','SGB XI 45b - Leistungsmonat',
      'SGB XI 39 - Start','SGB XI 39 - Leistungsmonat',
      'SGB XI 45a',        // (ALT: bisher als Topf genutzt – wir interpretieren fortan als Limit, wenn „45a Umwidmung aktiv“ leer ist)
      'Angelegt Am',
      '45a Umwidmung aktiv',          // NEU
      '45a Limit €/Monat'             // NEU (leer = Auto-40%)
    ];
    if (shK.getLastRow() === 0) {
      shK.getRange(1,1,1,wantK.length).setValues([wantK]);
    } else {
      const have = shK.getRange(1,1,1,shK.getLastColumn()).getValues()[0];
      const missing = wantK.filter(h => have.indexOf(h) === -1);
      if (missing.length) {
        shK.insertColumnsAfter(have.length, missing.length);
        shK.getRange(1,have.length+1,1,missing.length).setValues([missing]);
      }
    }

    const tmpl = ss.getSheetByName('K_TEMPLATE') || ss.insertSheet('K_TEMPLATE');
    const header = [
      'Rechnungsempfänger','Leistungsmonat','Beschreibung',
      'SGB XI 45b','SGB XI 39','SGB XI 45a','Rechnungsnummer','Limit 45a €/Monat'
    ];
    tmpl.getRange(1,1,1,header.length).setValues([header]);
  },

  getClients() {
    // … dein bestehender Code …
    const sh = SpreadsheetApp.getActive().getSheetByName('Klienten');
    if (!sh || sh.getLastRow() < 2) return [];
    const data = sh.getRange(2,1,sh.getLastRow()-1, sh.getLastColumn()).getValues();
    const head = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    return data.map(row => {
      const m = {};
      head.forEach((h,i)=> m[h]=row[i]);
      return m;
    });
  },

  getClientSheet(name){ return SpreadsheetApp.getActive().getSheetByName('K_'+name); },
  ensureClientSheet(name){
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName('K_'+name);
    if (!sh) sh = ss.insertSheet('K_'+name);
    // Header sicherstellen (inkl. Rechnungsnummer, Limit 45a)
    const want = ['Rechnungsempfänger','Leistungsmonat','Beschreibung','SGB XI 45b','SGB XI 39','SGB XI 45a','Rechnungsnummer','Limit 45a €/Monat'];
    const have = sh.getRange(1,1,1,Math.max(1, sh.getLastColumn())).getValues()[0];
    if (have.join('|') !== want.join('|')) sh.getRange(1,1,1,want.length).setValues([want]);
    return sh;
  },

  appendClientBooking({client, ym, beschr, s45b, s39, s45a, invoiceNo}) {
    const sh = this.ensureClientSheet(client);
    sh.appendRow([client, ym, beschr, s45b||0, s39||0, s45a||0, invoiceNo||'', '']); // H (Limit) bleibt leer auf Buchungszeilen
  }
};