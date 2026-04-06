/**
 * SPYHistory.gs — Fetches and stores daily SPY and QQQ closing prices.
 * Version: 2.0 (2026-04-05) — Ported from Schwab. Actual close prices only; no indexed columns.
 *
 * Sheet layout (SHEET_SPY):
 *   Date | SPY Close ($) | QQQ Close ($)
 *
 * Dates are stored as plain text strings ('2026-01-02'), not Date objects.
 *
 * ╔══════════════════════════════════════════════════════════════╗
 * ║  DATA SAFETY GUARANTEE                                       ║
 * ║  Actual close price columns (SPY, QQQ) are NEVER deleted     ║
 * ║  or overwritten. The migration only removes the two          ║
 * ║  "Indexed (Base=100)" columns that are no longer needed.     ║
 * ╚══════════════════════════════════════════════════════════════╝
 */

function getOrCreateSPYSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_SPY);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_SPY);
    const hdr = sheet.getRange(1, 1, 1, 3);
    hdr.setValues([['Date', 'SPY Close ($)', 'QQQ Close ($)']]);
    hdr.setFontWeight('bold').setBackground('#cc0000').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 120);
    [2, 3].forEach(function(c) { sheet.setColumnWidth(c, 140); });
  } else {
    migrateSPYHistoryColumns_(sheet);
  }
  return sheet;
}

function migrateSPYHistoryColumns_(sheet) {
  if (sheet.getLastColumn() < 5) return;
  sheet.deleteColumn(5);
  sheet.deleteColumn(3);
  var hdr = sheet.getRange(1, 1, 1, 3);
  hdr.setValues([['Date', 'SPY Close ($)', 'QQQ Close ($)']]);
  hdr.setFontWeight('bold').setBackground('#cc0000').setFontColor('#ffffff');
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(3, 140);
}

function fetchSPYHistory() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const tz    = Session.getScriptTimeZone();
  const today = fmtDate_(new Date(), tz);

  try {
    ss.toast('Fetching SPY and QQQ price history…', 'Working', -1);

    const sheet   = getOrCreateSPYSheet_();
    const lastRow = sheet.getLastRow();

    var fetchStart, fullRewrite;

    if (lastRow > 1) {
      const allRows = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
      const existingDates = {};
      function toDateStr_(v) {
        if (!v) return null;
        if (typeof v === 'string') return v;
        return fmtDate_(v, tz);
      }
      allRows.forEach(function(r) {
        var s = toDateStr_(r[0]);
        if (s) existingDates[s] = true;
      });

      const lastDate = toDateStr_(allRows[allRows.length - 1][0]);
      const lp       = lastDate.split('-');
      fetchStart     = fmtDate_(new Date(+lp[0], +lp[1] - 1, +lp[2] + 1), tz);
      fullRewrite    = false;

      if (fetchStart > today) {
        ss.toast('SPY/QQQ history is already up to date.', 'No Update Needed', 5);
        return;
      }

      const spyData = fetchCandlesForSymbol_('SPY', fetchStart, today);
      const qqqData = fetchCandlesForSymbol_('QQQ', fetchStart, today);

      if (spyData.length === 0) {
        ss.toast('No new SPY data since ' + fetchStart + '.', 'Already Up To Date', 5);
        return;
      }

      const spyMap = {}, qqqMap = {};
      spyData.forEach(function(c) { spyMap[fmtDate_(new Date(c.datetime), tz)] = c.close; });
      qqqData.forEach(function(c) { qqqMap[fmtDate_(new Date(c.datetime), tz)] = c.close; });

      const datesToWrite = Object.keys(spyMap).sort()
        .filter(function(d) { return !existingDates[d]; });

      if (datesToWrite.length === 0) {
        ss.toast('SPY/QQQ history is already up to date.', 'No Update Needed', 5);
        return;
      }

      const newRows = datesToWrite.map(function(d) {
        return [d, spyMap[d], qqqMap[d] || ''];
      });

      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, newRows.length, 3).setValues(newRows);
      sheet.getRange(startRow, 2, newRows.length, 2).setNumberFormat('"$"#,##0.00');

      backupSPYHistory_();
      ss.toast('✅ Added ' + newRows.length + ' new day(s) — SPY & QQQ', 'Benchmarks Updated', 10);

    } else {
      const spyData = fetchCandlesForSymbol_('SPY', HISTORY_START, today);
      const qqqData = fetchCandlesForSymbol_('QQQ', HISTORY_START, today);

      if (spyData.length === 0) {
        SpreadsheetApp.getUi().alert('No SPY data returned. Check authorization and try again.');
        return;
      }

      const spyMap = {}, qqqMap = {};
      spyData.forEach(function(c) { spyMap[fmtDate_(new Date(c.datetime), tz)] = c.close; });
      qqqData.forEach(function(c) { qqqMap[fmtDate_(new Date(c.datetime), tz)] = c.close; });

      const allDates = Object.keys(spyMap).sort();
      const rows = allDates.map(function(d) {
        return [d, spyMap[d], qqqMap[d] || ''];
      });

      if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
      sheet.getRange(2, 1, rows.length, 3).setValues(rows);
      sheet.getRange(2, 2, rows.length, 2).setNumberFormat('"$"#,##0.00');

      backupSPYHistory_();
      ss.toast(
        '✅ SPY: ' + spyData.length + ' days  |  QQQ: ' + qqqData.length + ' days',
        'Benchmarks Updated', 10
      );
    }

  } catch (e) {
    SpreadsheetApp.getUi().alert('Fetch Error', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    console.error(e);
  }
}

function fetchCandlesForSymbol_(symbol, startDate, endDate) {
  const allCandles = [];
  const start = new Date(startDate);
  const end   = new Date(endDate);

  for (var cur = new Date(start); cur <= end; ) {
    var chunkEnd = new Date(cur);
    chunkEnd.setFullYear(chunkEnd.getFullYear() + 5);
    if (chunkEnd > end) chunkEnd.setTime(end.getTime());

    try {
      var resp = getPriceHistory(symbol, cur.toISOString().slice(0, 10),
                                          chunkEnd.toISOString().slice(0, 10));
      if (resp && Array.isArray(resp.candles)) {
        Array.prototype.push.apply(allCandles, resp.candles);
      }
    } catch (e) {
      console.warn(symbol + ' chunk error: ' + e.message);
    }

    cur = new Date(chunkEnd);
    cur.setDate(cur.getDate() + 1);
  }

  var seen = {};
  return allCandles
    .filter(function(c) {
      if (seen[c.datetime]) return false;
      seen[c.datetime] = true;
      return true;
    })
    .sort(function(a, b) { return a.datetime - b.datetime; });
}

function getBenchmarkCloseMaps_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SPY);
  if (!sheet) return { spy: {}, qqq: {} };

  const tz = Session.getScriptTimeZone();
  const spy = {}, qqq = {};
  sheet.getDataRange().getValues().slice(1).forEach(function(row) {
    if (!row[0] || !row[1]) return;
    var raw = row[0];
    var d = (raw instanceof Date) ? fmtDate_(raw, tz) : String(raw).trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) return;
    spy[d] = parseFloat(row[1]) || 0;
    if (row[2]) qqq[d] = parseFloat(row[2]) || 0;
  });

  return { spy: spy, qqq: qqq };
}

function roundTo2_(n) {
  return Math.round(n * 100) / 100;
}
