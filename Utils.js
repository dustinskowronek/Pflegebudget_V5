/** Utils.js — Hilfsfunktionen, NUR HIER definieren! */
if (typeof Utils === 'undefined') { var Utils = {}; }

(function(){
  Utils.tz = function(){ return SpreadsheetApp.getActive().getSpreadsheetTimeZone() || 'Europe/Berlin'; };
  Utils.now = function(){ return new Date(); };
  Utils.nowISO = function(){
    return Utilities.formatDate(Utils.now(), Utils.tz(), "yyyy-MM-dd'T'HH:mm:ss");
  };

  Utils.sheet = function(name){
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(name);
    if (!sh) throw new Error('Sheet nicht gefunden: '+name);
    return sh;
  };

  Utils.ensureSheet = function(name, header){
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      if (header && header.length) sh.appendRow(header);
    } else {
      // Header sicherstellen
      if (header && header.length) {
        var rng = sh.getRange(1,1,1,header.length);
        var current = rng.getValues()[0];
        var changed = false;
        for (var i=0;i<header.length;i++){
          if (current[i] !== header[i]) { current[i] = header[i]; changed = true; }
        }
        if (changed) rng.setValues([current]);
      }
    }
    return sh;
  };

  Utils.ensureSheetFromTemplate = function(templateName, targetName){
    var ss = SpreadsheetApp.getActive();
    var tpl = ss.getSheetByName(templateName);
    var tgt = ss.getSheetByName(targetName);
    if (tgt) return tgt;
    if (tpl) {
      tpl.copyTo(ss).setName(targetName);
      return ss.getSheetByName(targetName);
    }
    // Fallback: blank mit erwarteten Spalten
    var sh = ss.insertSheet(targetName);
    sh.appendRow(CFG.COLUMNS.KLIENT_SHEET_HEADER);
    return sh;
  };

  Utils.readTable = function(sheetName){
    var sh = Utils.sheet(sheetName);
    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    if (lastRow < 2) return [];
    var header = sh.getRange(1,1,1,lastCol).getValues()[0];
    var data = sh.getRange(2,1,lastRow-1,lastCol).getValues();
    return data.map(function(r){
      var o={};
      for (var i=0;i<header.length;i++){ o[header[i]] = r[i]; }
      return o;
    });
  };

  Utils.writeRow = function(sheetName, arr){
    Utils.sheet(sheetName).appendRow(arr);
  };

  Utils.toNumber = function(x){
    var n = Number(String(x||'').replace(',','.'));
    return isNaN(n) ? 0 : n;
  };

  Utils.slug = function(name){
    return String(name||'').trim().toLowerCase().replace(/\s+/g,'-').replace(/[^\w\-äöüß]/g,'');
  };

  /** monthKey: 'YYYY-MM' aus Jahr und Monat '01'..'12' */
  Utils.monthKey = function(year, month2){
    var y = String(year);
    var m = String(month2).padStart(2,'0');
    if (!/^\d{4}$/.test(y)) throw new Error('Ungültiges Jahr: '+y);
    if (!/^\d{2}$/.test(m)) throw new Error('Ungültiger Monat: '+m);
    return y+'-'+m;
  };
})();