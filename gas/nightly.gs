// =============================================================
// WOLFPACK WEATHER — NIGHTLY SERVER-SIDE LOGGER
// =============================================================
// SETUP (one time only):
//   1. Open your Google Sheet → Extensions → Apps Script
//   2. Paste this entire file at the bottom of your existing script
//   3. Select "installNightlyTrigger" in the function dropdown
//   4. Click Run — done. Fires every night at 8 PM ET automatically.
// =============================================================

function installNightlyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'nightlyLog') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('nightlyLog')
    .timeBased()
    .atHour(20)
    .everyDays(1)
    .inTimezone('America/New_York')
    .create();
  Logger.log('Nightly trigger installed — fires daily at 8 PM ET');
}

function nightlyLog() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Accuracy');
  if (!sheet) { Logger.log('ERROR: Accuracy sheet not found'); return; }
  nightlyLogPredictions_(sheet);
  nightlyLogActuals_(sheet);
}

function nightlyLogPredictions_(sheet) {
  var opts = { headers: {'User-Agent': 'WolfpackWeather/2.0'}, muteHttpExceptions: true };

  var pr = UrlFetchApp.fetch('https://api.weather.gov/points/35.686,-78.614', opts);
  if (pr.getResponseCode() !== 200) { Logger.log('NWS points failed: ' + pr.getResponseCode()); return; }
  var forecastUrl = JSON.parse(pr.getContentText()).properties.forecast;
  if (!forecastUrl) { Logger.log('No forecast URL'); return; }

  var fr = UrlFetchApp.fetch(forecastUrl, opts);
  if (fr.getResponseCode() !== 200) { Logger.log('NWS forecast failed: ' + fr.getResponseCode()); return; }
  var periods = JSON.parse(fr.getContentText()).properties.periods;

  var byDate = {};
  periods.forEach(function(p) {
    var ds = p.startTime ? p.startTime.substring(0, 10) : '';
    if (!ds) return;
    if (!byDate[ds]) byDate[ds] = { hi: null, lo: null };
    if ( p.isDaytime && byDate[ds].hi === null) byDate[ds].hi = p.temperature;
    if (!p.isDaytime && byDate[ds].lo === null) byDate[ds].lo = p.temperature;
  });

  var now = new Date();
  for (var lead = 1; lead <= 5; lead++) {
    var td = new Date(now.getFullYear(), now.getMonth(), now.getDate() + lead);
    var ds = Utilities.formatDate(td, 'America/New_York', 'yyyy-MM-dd');
    var day = byDate[ds];
    if (!day || day.hi === null || day.lo === null) continue;
    var hiKey = lead === 1 ? 'pred_high' : 'pred_hi_d' + lead;
    var loKey = lead === 1 ? 'pred_low'  : 'pred_lo_d' + lead;
    nightlyWrite_(sheet, ds, hiKey, day.hi, loKey, day.lo, false);
    Logger.log('Pred D' + lead + ': ' + ds + ' Hi:' + day.hi + ' Lo:' + day.lo);
  }
}

function nightlyLogActuals_(sheet) {
  var opts = { headers: {'User-Agent': 'WolfpackWeather/2.0'}, muteHttpExceptions: true };
  var now = new Date();

  for (var daysBack = 0; daysBack <= 4; daysBack++) {
    var td = new Date(now.getFullYear(), now.getMonth(), now.getDate() - daysBack);
    var ds = Utilities.formatDate(td, 'America/New_York', 'yyyy-MM-dd');
    var obsS = encodeURIComponent(ds + 'T00:00:00-04:00');
    var obsE = encodeURIComponent(ds + 'T23:59:59-04:00');
    var url = 'https://api.weather.gov/stations/KRDU/observations?start=' + obsS + '&end=' + obsE + '&limit=200';

    var resp = UrlFetchApp.fetch(url, opts);
    if (resp.getResponseCode() !== 200) { Logger.log('NWS obs failed ' + ds + ': ' + resp.getResponseCode()); continue; }

    var features = JSON.parse(resp.getContentText()).features || [];
    var temps = features.reduce(function(a, f) {
      var t = f.properties && f.properties.temperature && f.properties.temperature.value;
      if (t != null) a.push(t * 9/5 + 32);
      return a;
    }, []);

    if (!temps.length) { Logger.log('No temps for ' + ds); continue; }
    var hi = Math.round(Math.max.apply(null, temps));
    var lo = Math.round(Math.min.apply(null, temps));
    nightlyWrite_(sheet, ds, 'actual_high', hi, 'actual_low', lo, true);
    Logger.log('Actuals ' + ds + ': H:' + hi + ' L:' + lo);
  }
}

function nightlyWrite_(sheet, dateStr, key1, val1, key2, val2, overwrite) {
  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var dateCol = headers.indexOf('date');
  if (dateCol < 0) dateCol = 0;

  var rowIdx = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][dateCol]).trim() === dateStr) { rowIdx = i; break; }
  }
  if (rowIdx < 0) {
    var newRow = new Array(headers.length).fill('');
    newRow[dateCol] = dateStr;
    sheet.appendRow(newRow);
    data = sheet.getDataRange().getValues();
    rowIdx = data.length - 1;
  }

  var c1 = headers.indexOf(key1);
  var c2 = headers.indexOf(key2);
  if (c1 >= 0 && (overwrite || !data[rowIdx][c1])) sheet.getRange(rowIdx + 1, c1 + 1).setValue(val1);
  if (c2 >= 0 && (overwrite || !data[rowIdx][c2])) sheet.getRange(rowIdx + 1, c2 + 1).setValue(val2);
}
