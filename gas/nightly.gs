// =============================================================
// WOLFPACK WEATHER — NIGHTLY SERVER-SIDE LOGGER
// =============================================================
// SETUP (one time only):
//   1. Open your Google Sheet → Extensions → Apps Script
//   2. Paste this entire file, replacing your existing script
//   3. Select "installNightlyTrigger" in the function dropdown
//   4. Click Run — done. Fires every night at 8 PM ET automatically.
//
// TO TEST MANUALLY:
//   Select "testNightlyLog" and click Run — writes real data now.
// =============================================================

var TZ = 'America/New_York';

function installNightlyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'nightlyLog') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('nightlyLog')
    .timeBased()
    .atHour(20)
    .everyDays(1)
    .inTimezone(TZ)
    .create();
  Logger.log('✓ Nightly trigger installed — fires daily at 8 PM ET');
}

// Called automatically by the trigger every night at 8 PM ET
function nightlyLog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Accuracy');
  if (!sheet) { Logger.log('ERROR: Accuracy sheet not found'); return; }
  Logger.log('=== nightlyLog START ' + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm z') + ' ===');
  nightlyLogPredictions_(sheet);
  nightlyLogActuals_(sheet);
  Logger.log('=== nightlyLog END ===');
}

// Run this manually from the editor anytime to test the full pipeline
function testNightlyLog() {
  Logger.log('--- MANUAL TEST RUN ---');
  nightlyLog();
}

// ─────────────────────────────────────────────────────────────
// PREDICTIONS  (D1–D5 from NWS Raleigh forecast)
// ─────────────────────────────────────────────────────────────
function nightlyLogPredictions_(sheet) {
  var opts = { headers: { 'User-Agent': 'WolfpackWeather/2.0' }, muteHttpExceptions: true };

  var pr = UrlFetchApp.fetch('https://api.weather.gov/points/35.686,-78.614', opts);
  if (pr.getResponseCode() !== 200) { Logger.log('NWS points failed: ' + pr.getResponseCode()); return; }
  var forecastUrl = JSON.parse(pr.getContentText()).properties.forecast;
  if (!forecastUrl) { Logger.log('No forecast URL in NWS response'); return; }

  var fr = UrlFetchApp.fetch(forecastUrl, opts);
  if (fr.getResponseCode() !== 200) { Logger.log('NWS forecast failed: ' + fr.getResponseCode()); return; }
  var periods = JSON.parse(fr.getContentText()).properties.periods;
  if (!periods || !periods.length) { Logger.log('No forecast periods returned'); return; }

  var byDate = {};
  periods.forEach(function(p) {
    var ds = p.startTime ? p.startTime.substring(0, 10) : '';
    if (!ds) return;
    if (!byDate[ds]) byDate[ds] = { hi: null, lo: null };
    if ( p.isDaytime && byDate[ds].hi === null) byDate[ds].hi = p.temperature;
    if (!p.isDaytime && byDate[ds].lo === null) byDate[ds].lo = p.temperature;
  });

  var now = new Date();
  var wrote = 0;
  for (var lead = 1; lead <= 5; lead++) {
    var td  = new Date(now.getFullYear(), now.getMonth(), now.getDate() + lead);
    var ds  = Utilities.formatDate(td, TZ, 'yyyy-MM-dd');
    var day = byDate[ds];
    if (!day || day.hi === null || day.lo === null) {
      Logger.log('Pred D' + lead + ' (' + ds + '): no data in forecast');
      continue;
    }
    var hiKey = lead === 1 ? 'pred_high' : 'pred_hi_d' + lead;
    var loKey = lead === 1 ? 'pred_low'  : 'pred_lo_d' + lead;
    nightlyWrite_(sheet, ds, hiKey, day.hi, loKey, day.lo, false);
    Logger.log('Pred D' + lead + ' → ' + ds + '  Hi:' + day.hi + '  Lo:' + day.lo);
    wrote++;
  }
  Logger.log('Predictions written: ' + wrote + ' of 5 lead days');
}

// ─────────────────────────────────────────────────────────────
// ACTUALS  (NWS KRDU hourly observations, last 5 days)
// ─────────────────────────────────────────────────────────────
function nightlyLogActuals_(sheet) {
  var opts = { headers: { 'User-Agent': 'WolfpackWeather/2.0' }, muteHttpExceptions: true };
  var now  = new Date();

  for (var daysBack = 0; daysBack <= 4; daysBack++) {
    var td = new Date(now.getFullYear(), now.getMonth(), now.getDate() - daysBack);
    var ds = Utilities.formatDate(td, TZ, 'yyyy-MM-dd');

    // FIX: build UTC timestamps from local midnight/11:59 PM so the window
    // is correct in both EST (UTC-5) and EDT (UTC-4).
    // new Date(y, m, d, h, min, sec) in Apps Script uses the SCRIPT timezone,
    // so .toISOString() returns the matching UTC time automatically.
    var startLocal = new Date(td.getFullYear(), td.getMonth(), td.getDate(),  0,  0,  0);
    var endLocal   = new Date(td.getFullYear(), td.getMonth(), td.getDate(), 23, 59, 59);
    var obsS = encodeURIComponent(startLocal.toISOString());
    var obsE = encodeURIComponent(endLocal.toISOString());

    var url  = 'https://api.weather.gov/stations/KRDU/observations?start=' + obsS + '&end=' + obsE + '&limit=200';
    var resp = UrlFetchApp.fetch(url, opts);

    if (resp.getResponseCode() !== 200) {
      Logger.log('NWS obs FAILED ' + ds + ' (HTTP ' + resp.getResponseCode() + ')');
      continue;
    }

    var features = JSON.parse(resp.getContentText()).features || [];
    var temps = features.reduce(function(a, f) {
      var t = f.properties && f.properties.temperature && f.properties.temperature.value;
      if (t != null) a.push(t * 9/5 + 32);
      return a;
    }, []);

    if (!temps.length) {
      Logger.log('Actuals ' + ds + ': 0 temperature readings returned (NWS may not have data yet)');
      continue;
    }

    var hi = Math.round(Math.max.apply(null, temps));
    var lo = Math.round(Math.min.apply(null, temps));
    nightlyWrite_(sheet, ds, 'actual_high', hi, 'actual_low', lo, true);
    Logger.log('Actuals ' + ds + ':  Hi:' + hi + '  Lo:' + lo + '  (' + temps.length + ' obs)');
  }
}

// ─────────────────────────────────────────────────────────────
// SHEET WRITER  (find-or-create row, write two columns)
// ─────────────────────────────────────────────────────────────
function nightlyWrite_(sheet, dateStr, key1, val1, key2, val2, overwrite) {
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var dateCol = headers.indexOf('date');
  if (dateCol < 0) dateCol = 0;

  // Find existing row for this date
  var rowIdx = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][dateCol]).trim() === dateStr) { rowIdx = i; break; }
  }

  // If no row exists, append one
  if (rowIdx < 0) {
    var newRow = new Array(headers.length).fill('');
    newRow[dateCol] = dateStr;
    sheet.appendRow(newRow);
    data   = sheet.getDataRange().getValues();  // re-read after append
    rowIdx = data.length - 1;
    Logger.log('  Created new row for ' + dateStr);
  }

  // Write values (overwrite=true for actuals, false for predictions)
  var c1 = headers.indexOf(key1);
  var c2 = headers.indexOf(key2);
  if (c1 < 0) Logger.log('  WARNING: column "' + key1 + '" not found in sheet headers');
  if (c2 < 0) Logger.log('  WARNING: column "' + key2 + '" not found in sheet headers');
  if (c1 >= 0 && (overwrite || !data[rowIdx][c1])) sheet.getRange(rowIdx + 1, c1 + 1).setValue(val1);
  if (c2 >= 0 && (overwrite || !data[rowIdx][c2])) sheet.getRange(rowIdx + 1, c2 + 1).setValue(val2);
}
