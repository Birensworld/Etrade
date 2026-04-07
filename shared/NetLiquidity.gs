/**
 * NetLiquidity.gs — Manages the shared "Net Liquidity" sheet.
 * Version: 1.4 (2026-04-07) — Rename captureNetLiqForAccount → captureNetLiqForAccount to bypass GAS cache
 *
 * Sheet layout:
 *   Date | Net Liq 7806 ($) | Net Liq 8090 ($) | Net Liq 3945 ($) | Total Net Liquidity
 *
 * ╔══════════════════════════════════════════════════════════════╗
 * ║  DATA SAFETY GUARANTEE                                       ║
 * ║  No function in this file ever deletes or clears rows from   ║
 * ║  the Net Liquidity sheet. All writes are append-only or      ║
 * ║  single-cell updates. Manually entered values are preserved. ║
 * ╚══════════════════════════════════════════════════════════════╝
 *
 * Write rules:
 *   • Daily trigger (skipIfExists=true)  → only adds new dates, never overwrites
 *   • Menu "Capture Today"  (skipIfExists=true)  → same; won't overwrite manual entry
 *   • CSV import            (skipIfExists=false) → fills empty cells only if you choose
 */

// ─────────────────────────────────────────────────────────────────
// Sheet bootstrap
// ─────────────────────────────────────────────────────────────────

function getOrCreateNetLiqSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NET_LIQ);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NET_LIQ);

    var headers = ['Date'];
    ACCOUNT_ORDER.forEach(function(s) { headers.push('Net Liquidity ' + s + ' ($)'); });
    headers.push('Total Net Liquidity');

    var hdrRange = sheet.getRange(1, 1, 1, headers.length);
    hdrRange.setValues([headers]);
    hdrRange.setFontWeight('bold').setBackground('#0b5394').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 120);
    ACCOUNT_ORDER.forEach(function(_, i) { sheet.setColumnWidth(i + 2, 185); });
    sheet.setColumnWidth(totalNetLiqCol_(), 160);

    // Pre-format all value columns for the full sheet as $#,##0.00
    var numCols = ACCOUNT_ORDER.length + 1; // account cols + Total
    sheet.getRange(2, 2, sheet.getMaxRows() - 1, numCols)
         .setNumberFormat('"$"#,##0.00');
  }
  return sheet;
}

// ─────────────────────────────────────────────────────────────────
// Public: capture today's value
// ─────────────────────────────────────────────────────────────────

/**
 * Fetches live Net Liquidity from E*Trade for one account and logs it.
 *
 * @param {string}  suffix        e.g. '7806'
 * @param {boolean} skipIfExists  If true, don't overwrite an existing value for today (default: true)
 * @param {boolean} silent        If true, suppress toasts/alerts — caller handles UI (default: false)
 * @returns {{ status: 'captured'|'skipped'|'error', suffix, value, dateStr, message }}
 */
function captureNetLiqForAccount(suffix, skipIfExists, silent) {
  if (skipIfExists === undefined) skipIfExists = true;
  if (silent      === undefined) silent       = false;

  try {
    var value   = fetchNetLiqForSuffix_(suffix);
    var dateStr = todayStr_();
    var written = upsertNetLiqRow_(suffix, dateStr, value, 'API', skipIfExists);

    if (!written) {
      if (!silent) {
        SpreadsheetApp.getActiveSpreadsheet().toast(
          'Account …' + suffix + ': entry for ' + dateStr + ' already exists — skipped.',
          'ℹ️ No Change', 5
        );
      }
      return { status: 'skipped', suffix: suffix, value: value, dateStr: dateStr };
    }

    backupNetLiq_();
    if (!silent) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Account …' + suffix + '  Net Liq: $' +
          value.toLocaleString('en-US', { minimumFractionDigits: 2 }),
        '✅ Captured', 5
      );
    }
    return { status: 'captured', suffix: suffix, value: value, dateStr: dateStr };

  } catch (e) {
    if (!silent) {
      SpreadsheetApp.getUi().alert('Error – Account ' + suffix, e.message,
        SpreadsheetApp.getUi().ButtonSet.OK);
    }
    console.error(e);
    return { status: 'error', suffix: suffix, message: e.message };
  }
}

// ─────────────────────────────────────────────────────────────────
// Public: CSV import
// ─────────────────────────────────────────────────────────────────

function showCsvImportDialog() {
  var html = HtmlService.createHtmlOutputFromFile('CsvImport')
    .setWidth(480).setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(html, 'Import Net Liquidity from CSV');
}

function importNetLiqFromCsv(rows, suffix) {
  suffix = suffix || '7806';
  if (!rows || rows.length === 0) throw new Error('No data received.');
  var tz = Session.getScriptTimeZone();
  rows.forEach(function(row) {
    var d = new Date(row.date);
    if (isNaN(d.getTime())) return;
    upsertNetLiqRow_(suffix, fmtDate_(d, tz), row.value, 'CSV Import', false);
  });
  backupNetLiq_();
  return 'Imported ' + rows.length + ' rows into account …' + suffix + '.';
}

// ─────────────────────────────────────────────────────────────────
// Sheet read / write helpers
// ─────────────────────────────────────────────────────────────────

/**
 * Appends a new date row or fills a single account cell.
 * NEVER deletes or clears any row.
 * @returns {boolean} true if a value was written, false if skipped
 */
function upsertNetLiqRow_(suffix, dateStr, value, source, skipIfExists) {
  var sheet = getOrCreateNetLiqSheet_();
  var tz    = Session.getScriptTimeZone();
  var col   = netLiqCol_(suffix);
  var data  = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var cellVal = data[i][0];
    var cellStr = (cellVal instanceof Date) ? fmtDate_(cellVal, tz) : String(cellVal);
    if (cellStr !== dateStr) continue;

    var existing = data[i][col - 1];
    var hasValue = (existing !== '' && existing !== 0 && existing !== null && existing !== undefined);
    if (skipIfExists && hasValue) return false;

    sheet.getRange(i + 1, col).setValue(value).setNumberFormat('"$"#,##0.00');
    return true;
  }

  var newRow = [dateStr];
  ACCOUNT_ORDER.forEach(function(s) {
    newRow.push(s === suffix ? value : '');
  });
  var lastAcctCol = String.fromCharCode(64 + ACCOUNT_ORDER.length + 1);
  newRow.push('');
  sheet.appendRow(newRow);

  var newRowNum    = sheet.getLastRow();
  var totalFormula = '=SUM(B' + newRowNum + ':' + lastAcctCol + newRowNum + ')';
  sheet.getRange(newRowNum, totalNetLiqCol_()).setFormula(totalFormula);

  // Apply $#,##0.00 format to all value cells in the new row
  var numCols = ACCOUNT_ORDER.length + 1; // account cols + Total
  sheet.getRange(newRowNum, 2, 1, numCols).setNumberFormat('"$"#,##0.00');

  return true;
}

function sortNetLiqSheet_(sheet) {
  var last = sheet.getLastRow();
  if (last > 2) sheet.getRange(2, 1, last - 1, ACCOUNT_ORDER.length + 2).sort(1);
}

/**
 * Returns a date-keyed map of Net Liquidity values for one account.
 * @param {string} suffix  e.g. '7806'
 * @returns {Object}  { 'YYYY-MM-DD': number }
 */
function getNetLiqMap_(suffix) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NET_LIQ);
  if (!sheet) return {};

  var col = netLiqCol_(suffix) - 1;
  var map = {};

  var tz_ = Session.getScriptTimeZone();
  sheet.getDataRange().getValues().slice(1).forEach(function(row) {
    if (!row[0]) return;
    var raw = row[0];
    var dateStr = (raw instanceof Date) ? fmtDate_(raw, tz_) : String(raw).trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return;
    var v = Number(row[col]);
    if (!isFinite(v) || v < 0) return;
    map[dateStr] = v;
  });

  return map;
}

// ─────────────────────────────────────────────────────────────────
// Utilities
// ─────────────────────────────────────────────────────────────────

function fmtDate_(date, tz) {
  return Utilities.formatDate(date, tz || Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function todayStr_() { return fmtDate_(new Date()); }
