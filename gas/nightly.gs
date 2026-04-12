// =============================================================
// RED WOLF WEATHER — NIGHTLY SERVER-SIDE LOGGER
// =============================================================
// SETUP (one time only):
//   1. Open your Google Sheet → Extensions → Apps Script
//   2. Paste this entire file, replacing your existing script
//   3. Select "installNightlyTrigger" in the function dropdown
//   4. Click Run — done. Fires every night at 8 PM ET automatically.
//
// MICROCLIMATE SETUP (one time):
//   In the Google Sheet, add a tab named "Microclimate" with
//   row 1 headers:  date  rwf_high  rwf_low  rdu_high  rdu_low
//
// TO TEST MANUALLY:
//   Select "testNightlyLog" and click Run — writes real data now.
// =============================================================

var TZ = 'America/New_York';

// ─────────────────────────────────────────────────────────────
// WEB APP ENDPOINTS  (Deploy → Manage Deployments → Web app)
//   Execute as: Me   |   Who has access: Anyone
//
// GET  ?tab=Accuracy           → returns { rows: [...] }
// GET  ?action=write&tab=Accuracy&date=YYYY-MM-DD&col=val...
//                              → writes values, returns { ok:true }
// ─────────────────────────────────────────────────────────────
function doGet(e) {
  var p = e && e.parameter ? e.parameter : {};
  var tabName = p.tab || 'Accuracy';

  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(tabName);
    if (!sheet) return jsonOut({ error: 'Sheet not found: ' + tabName });

    // ── WRITE action ──────────────────────────────────────────
    if (p.action === 'write') {
      var dateStr = p.date || '';
      if (!dateStr) return jsonOut({ error: 'Missing date param' });

      var updates = {};
      Object.keys(p).forEach(function(k) {
        if (k !== 'action' && k !== 'tab' && k !== 'date') updates[k] = p[k];
      });

      var data    = sheet.getDataRange().getValues();
      var headers = data[0].map(function(h) { return String(h).trim(); });
      var dateCol = headers.indexOf('date');
      if (dateCol < 0) dateCol = 0;

      var rowIdx = -1;
      for (var i = 1; i < data.length; i++) {
        if (normDateStr_(data[i][dateCol]) === dateStr) { rowIdx = i; break; }
      }
      if (rowIdx < 0) {
        var newRow = new Array(headers.length).fill('');
        newRow[dateCol] = dateStr;
        sheet.appendRow(newRow);
        data   = sheet.getDataRange().getValues();
        rowIdx = data.length - 1;
      }

      var wrote = [];
      Object.keys(updates).forEach(function(k) {
        var ci = headers.indexOf(k);
        if (ci >= 0) {
          sheet.getRange(rowIdx + 1, ci + 1).setValue(updates[k]);
          wrote.push(k);
        }
      });

      return jsonOut({ ok: true, date: dateStr, wrote: wrote });
    }

    // ── READ (default) ────────────────────────────────────────
    var data    = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim(); });
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var obj = {};
      var hasContent = false;
      for (var c = 0; c < headers.length; c++) {
        var v = row[c];
        if (v instanceof Date) {
          v = Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
        }
        obj[headers[c]] = v;
        if (v !== '' && v !== null && v !== undefined) hasContent = true;
      }
      if (hasContent) rows.push(obj);
    }
    return jsonOut({ rows: rows });

  } catch (err) {
    return jsonOut({ error: err.message });
  }
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

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
  nightlyLogMicroclimate_(ss);
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
    if (p.isDaytime) {
      if (!byDate[ds]) byDate[ds] = { hi: null, lo: null };
      if (byDate[ds].hi === null) byDate[ds].hi = p.temperature;
    } else {
      // Night period leads into next morning — assign lo to D+1
      var d = new Date(ds + 'T12:00:00');
      d.setDate(d.getDate() + 1);
      var nextDs = Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
      if (!byDate[nextDs]) byDate[nextDs] = { hi: null, lo: null };
      if (byDate[nextDs].lo === null) byDate[nextDs].lo = p.temperature;
    }
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

  var hourET = parseInt(Utilities.formatDate(now, TZ, 'H'));
  var startBack = (hourET >= 20) ? 0 : 1;
  for (var daysBack = startBack; daysBack <= 6; daysBack++) {
    var td = new Date(now.getFullYear(), now.getMonth(), now.getDate() - daysBack);
    var ds = Utilities.formatDate(td, TZ, 'yyyy-MM-dd');

    // Build explicit ET-timezone ISO timestamps so the NWS query always covers
    // midnight-to-midnight Eastern Time, regardless of GAS runtime timezone (UTC).
    var tzOff = Utilities.formatDate(td, TZ, 'Z');            // e.g. "-0400" (EDT) or "-0500" (EST)
    var tzIso = tzOff.substring(0, 3) + ':' + tzOff.substring(3); // "-04:00"
    var obsS = encodeURIComponent(ds + 'T00:00:00' + tzIso);
    var obsE = encodeURIComponent(ds + 'T23:59:59' + tzIso);

    var url  = 'https://api.weather.gov/stations/KRDU/observations?start=' + obsS + '&end=' + obsE + '&limit=200';
    var resp = UrlFetchApp.fetch(url, opts);

    if (resp.getResponseCode() !== 200) {
      Logger.log('NWS obs FAILED ' + ds + ' (HTTP ' + resp.getResponseCode() + ')');
      continue;
    }

    var features = JSON.parse(resp.getContentText()).features || [];
    var temps = features.reduce(function(a, f) {
      var props = f.properties || {};
      var t = props.temperature && props.temperature.value;
      if (t != null) a.push(t * 9/5 + 32);
      // Include 6-hour synoptic extremes (reported at 00Z/06Z/12Z/18Z) —
      // these capture true highs/lows even when hourly temp values are null.
      var mn6 = props.minTemperatureLast6Hours && props.minTemperatureLast6Hours.value;
      var mx6 = props.maxTemperatureLast6Hours && props.maxTemperatureLast6Hours.value;
      if (mn6 != null) a.push(mn6 * 9/5 + 32);
      if (mx6 != null) a.push(mx6 * 9/5 + 32);
      return a;
    }, []);

    if (!temps.length) {
      Logger.log('Actuals ' + ds + ': 0 temperature readings returned');
      continue;
    }

    var hi = Math.round(Math.max.apply(null, temps));
    var lo = Math.round(Math.min.apply(null, temps));
    nightlyWrite_(sheet, ds, 'actual_high', hi, 'actual_low', lo, true);
    Logger.log('Actuals ' + ds + ':  Hi:' + hi + '  Lo:' + lo + '  (' + temps.length + ' obs)');
  }
}

// ─────────────────────────────────────────────────────────────
// MICROCLIMATE  (daily High/Low for KNCRALEI761 vs KRDU)
// Writes to "Microclimate" tab: date, rwf_high, rwf_low, rdu_high, rdu_low
// ─────────────────────────────────────────────────────────────
function nightlyLogMicroclimate_(ss) {
  var sheet = ss.getSheetByName('Microclimate');
  if (!sheet) { Logger.log('WARNING: Microclimate sheet not found — skipping microclimate log'); return; }

  var opts   = { headers: { 'User-Agent': 'WolfpackWeather/2.0' }, muteHttpExceptions: true };
  var WU_KEY = '6532d6454b8aa370768e63d6ba5a832e';
  var now    = new Date();

  // Process today (if >= 8 PM) and yesterday — same look-back pattern as nightlyLogActuals_
  var hourET    = parseInt(Utilities.formatDate(now, TZ, 'H'));
  var startBack = (hourET >= 20) ? 0 : 1;

  for (var daysBack = startBack; daysBack <= 1; daysBack++) {
    var td          = new Date(now.getFullYear(), now.getMonth(), now.getDate() - daysBack);
    var ds          = Utilities.formatDate(td, TZ, 'yyyy-MM-dd');
    var dateCompact = Utilities.formatDate(td, TZ, 'yyyyMMdd');

    // ── Red Wolf Farm (KNCRALEI761) via Weather Underground ──────
    var wuUrl  = 'https://api.weather.com/v2/pws/history/daily?stationId=KNCRALEI761&format=json&units=e&date=' + dateCompact + '&apiKey=' + WU_KEY;
    var rwfHigh = null, rwfLow = null;
    try {
      var wuResp = UrlFetchApp.fetch(wuUrl, opts);
      if (wuResp.getResponseCode() === 200) {
        var wuObs = JSON.parse(wuResp.getContentText()).observations;
        if (wuObs && wuObs.length > 0) {
          var imp = wuObs[0].imperial;
          if (imp && imp.tempHigh != null) rwfHigh = Math.round(imp.tempHigh);
          if (imp && imp.tempLow  != null) rwfLow  = Math.round(imp.tempLow);
        }
      } else {
        Logger.log('MC WU fetch failed: HTTP ' + wuResp.getResponseCode());
      }
    } catch (e) {
      Logger.log('MC WU fetch error: ' + e.message);
    }

    // ── RDU Airport (KRDU) via NWS hourly observations ──────────
    var tzOff2 = Utilities.formatDate(td, TZ, 'Z');
    var tzIso2 = tzOff2.substring(0, 3) + ':' + tzOff2.substring(3);
    var obsS = encodeURIComponent(ds + 'T00:00:00' + tzIso2);
    var obsE = encodeURIComponent(ds + 'T23:59:59' + tzIso2);
    var nwsUrl = 'https://api.weather.gov/stations/KRDU/observations?start=' + obsS + '&end=' + obsE + '&limit=200';
    var rduHigh = null, rduLow = null;
    try {
      var nwsResp = UrlFetchApp.fetch(nwsUrl, opts);
      if (nwsResp.getResponseCode() === 200) {
        var features = JSON.parse(nwsResp.getContentText()).features || [];
        var temps = features.reduce(function(a, f) {
          var props = f.properties || {};
          var t = props.temperature && props.temperature.value;
          if (t != null) a.push(t * 9/5 + 32);
          var mn6 = props.minTemperatureLast6Hours && props.minTemperatureLast6Hours.value;
          var mx6 = props.maxTemperatureLast6Hours && props.maxTemperatureLast6Hours.value;
          if (mn6 != null) a.push(mn6 * 9/5 + 32);
          if (mx6 != null) a.push(mx6 * 9/5 + 32);
          return a;
        }, []);
        if (temps.length) {
          rduHigh = Math.round(Math.max.apply(null, temps));
          rduLow  = Math.round(Math.min.apply(null, temps));
        } else {
          Logger.log('MC NWS: 0 temperature readings for ' + ds);
        }
      } else {
        Logger.log('MC NWS fetch failed: HTTP ' + nwsResp.getResponseCode());
      }
    } catch (e) {
      Logger.log('MC NWS fetch error: ' + e.message);
    }

    // ── Write to Microclimate sheet ──────────────────────────────
    if (rwfHigh === null && rwfLow === null && rduHigh === null && rduLow === null) {
      Logger.log('MC: no data retrieved — skipping write for ' + ds);
      continue;
    }

    var data    = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim(); });
    var dateCol = headers.indexOf('date');
    if (dateCol < 0) dateCol = 0;

    var rowIdx = -1;
    for (var i = 1; i < data.length; i++) {
      if (normDateStr_(data[i][dateCol]) === ds) { rowIdx = i; break; }
    }
    if (rowIdx < 0) {
      var newRow = new Array(headers.length).fill('');
      newRow[dateCol] = ds;
      sheet.appendRow(newRow);
      data   = sheet.getDataRange().getValues();
      rowIdx = data.length - 1;
      Logger.log('MC: created new row for ' + ds);
    }

    var cols = { rwf_high: rwfHigh, rwf_low: rwfLow, rdu_high: rduHigh, rdu_low: rduLow };
    Object.keys(cols).forEach(function(k) {
      var ci = headers.indexOf(k);
      if (ci >= 0 && cols[k] !== null) sheet.getRange(rowIdx + 1, ci + 1).setValue(cols[k]);
      else if (ci < 0) Logger.log('MC WARNING: column "' + k + '" not found in Microclimate sheet headers');
    });

    Logger.log('MC ' + ds + ':  RWF Hi:' + rwfHigh + ' Lo:' + rwfLow + '  |  RDU Hi:' + rduHigh + ' Lo:' + rduLow);
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

  var rowIdx = -1;
  for (var i = 1; i < data.length; i++) {
    if (normDateStr_(data[i][dateCol]) === dateStr) { rowIdx = i; break; }
  }

  if (rowIdx < 0) {
    var newRow = new Array(headers.length).fill('');
    newRow[dateCol] = dateStr;
    sheet.appendRow(newRow);
    data   = sheet.getDataRange().getValues();
    rowIdx = data.length - 1;
    Logger.log('  Created new row for ' + dateStr);
  }

  var c1 = headers.indexOf(key1);
  var c2 = headers.indexOf(key2);
  if (c1 < 0) Logger.log('  WARNING: column "' + key1 + '" not found in sheet headers');
  if (c2 < 0) Logger.log('  WARNING: column "' + key2 + '" not found in sheet headers');
  if (c1 >= 0 && (overwrite || !data[rowIdx][c1])) sheet.getRange(rowIdx + 1, c1 + 1).setValue(val1);
  if (c2 >= 0 && (overwrite || !data[rowIdx][c2])) sheet.getRange(rowIdx + 1, c2 + 1).setValue(val2);
}

// ─────────────────────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────────────────────

// Normalize any cell value to "YYYY-MM-DD" string.
// Sheets sometimes stores dates as Date objects — this handles both cases.
function normDateStr_(val) {
  if (!val && val !== 0) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, TZ, 'yyyy-MM-dd');
  }
  return String(val).trim();
}

// ─────────────────────────────────────────────────────────────
// DEDUP + SORT CLEANUP
// Run "recoverAndCleanup" any time the Accuracy sheet gets duplicate rows.
// Merges all rows sharing the same date into one, then sorts by date.
// ─────────────────────────────────────────────────────────────
function recoverAndCleanup() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Accuracy');
  if (!sheet) { Logger.log('ERROR: Accuracy sheet not found'); return; }

  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var dateCol = headers.indexOf('date');
  if (dateCol < 0) { Logger.log('ERROR: no "date" column found'); return; }

  var actualHiCol = headers.indexOf('actual_high');
  var actualLoCol = headers.indexOf('actual_low');

  var merged = {};
  var order  = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var ds  = normDateStr_(row[dateCol]);
    if (!ds) continue;

    if (!merged[ds]) {
      var copy = row.slice();
      copy[dateCol] = ds;
      merged[ds] = copy;
      order.push(ds);
    } else {
      var m = merged[ds];
      for (var c = 0; c < headers.length; c++) {
        var isActual = (c === actualHiCol || c === actualLoCol);
        var val = row[c];
        var hasVal = (val !== '' && val !== null && val !== undefined);
        if (hasVal && (isActual || !m[c] || m[c] === '')) {
          m[c] = val;
        }
      }
    }
  }

  order.sort();

  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();

  order.forEach(function(ds, i) {
    sheet.getRange(i + 2, 1, 1, headers.length).setValues([merged[ds]]);
  });

  Logger.log('✓ Dedup + sort complete. ' + order.length + ' unique dates.');
  Logger.log('Dates: ' + order.join(', '));
}
