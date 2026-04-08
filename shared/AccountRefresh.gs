/**
 * AccountRefresh.gs — Fetches and displays the E*Trade portfolio.
 * Version: 1.3 (2026-04-08) — Deep loss rule: light grey background + red font, evaluated before general loss
 *
 * Writes to sheet "EtradeB":
 *   • One header row per account (label | total value | cash | ALLOC | YTD P/L)
 *   • One positions sub-table per account (Symbol | Qty | Chg% | Avg Price | MV | P/L | P/L% | Weight)
 *   • Totals row at the bottom
 *
 * Depends on:
 *   Code.gs          — ACCOUNTS_URL, ACCOUNT_MAP, SHEET_ETRADE, ACCOUNT_ORDER
 *   Authentication.gs — buildOAuthHeaders_(), renewAccessToken()
 *   EtradeAPI.gs     — fetchEtradeBalance(), mergeHeaders_()
 *   NetLiquidity.gs  — getNetLiqMap_(suffix)
 */

// ─────────────────────────────────────────────────────────────────
// Public
// ─────────────────────────────────────────────────────────────────

function getPortfolio() {
  renewAccessToken();
  var tokens = getAccessTokens_();
  var token  = tokens.token;
  var secret = tokens.secret;

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_ETRADE) || ss.insertSheet(SHEET_ETRADE);
  sheet.activate();
  sheet.clear();
  sheet.setConditionalFormatRules([]);
  setSheetFont();

  var rowIndex    = 1;
  var accountRows = []; // tracks row numbers of account header rows for TOTALS formula
  var accountSummaries = []; // store values for totals YTD calculation

  // ── Fetch account list ────────────────────────────────────────
  var url      = ACCOUNTS_URL + '/list.json';
  var headers  = buildOAuthHeaders_(url, 'GET', token, secret);
  var response = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true, headers: headers });
  var data     = JSON.parse(response.getContentText());
  var accounts = (data.AccountListResponse &&
                  data.AccountListResponse.Accounts &&
                  data.AccountListResponse.Accounts.Account) || [];

  // ── Write each account ────────────────────────────────────────
  accounts.forEach(function(a) {
    var accountId  = a.accountId;
    var accountKey = a.accountIdKey;

    // Resolve display label from ACCOUNT_MAP (fall back to accountType)
    var suffix       = getAccountSuffix_(accountId);
    var displayLabel = (suffix && ACCOUNT_MAP[suffix]) ? ACCOUNT_MAP[suffix].label : a.accountType;
    var instType     = (suffix && ACCOUNT_MAP[suffix]) ? ACCOUNT_MAP[suffix].instType : a.institutionType;

    // ── Balance ───────────────────────────────────────────────
    var cash = 0, totalValue = 0, cashPct = 0, allocPercent = 0, ytdPL = 0;
    try {
      var bal      = fetchEtradeBalance(accountKey, accountId, instType, token, secret);
      cash         = bal.cash;
      totalValue   = bal.totalAccountValue;
      cashPct      = totalValue ? cash / totalValue : 0;
      allocPercent = totalValue ? (totalValue - cash) / totalValue : 0;
      ytdPL        = getEtradeYtdPLPercent_(accountId, totalValue);
    } catch (e) {
      console.warn('Balance fetch failed for ' + accountId + ': ' + e);
    }

    accountSummaries.push({
      accountId: accountId,
      suffix: suffix,
      totalValue: totalValue,
      cash: cash,
      allocPercent: allocPercent,
      ytdPL: ytdPL
    });

    // ── Account summary row ───────────────────────────────────
    sheet.getRange(rowIndex, 1, 1, 8).setValues([[
      displayLabel,
      totalValue,
      cash,
      "",
      "ALLOC:",
      allocPercent,
      "YTD P/L:",
      ytdPL
    ]]);
    accountRows.push(rowIndex);

    sheet.getRange(rowIndex, 1).setBackground('#d9d9d9').setFontSize(12);
    sheet.getRange(rowIndex, 1, 1, 8).setFontWeight('bold');
    sheet.getRange(rowIndex, 2, 1, 8).setBackground('#d0f0c0');
    sheet.getRange(rowIndex, 2, 1, 2).setNumberFormat('#,##0.00');
    sheet.getRange(rowIndex, 6).setNumberFormat('0.00%');
    sheet.getRange(rowIndex, 8).setNumberFormat('0.00%');

    applyYtdConditionalFormatting_(sheet, rowIndex);
    rowIndex++;

    // ── Positions header ──────────────────────────────────────
    var posHeader = ['Symbol', 'Qty', 'Chg %', 'Avg Price', 'MV', 'P/L', 'P/L %', 'Weight', ' ', ' '];
    sheet.getRange(rowIndex, 1, 1, posHeader.length)
      .setValues([posHeader])
      .setFontWeight('bold')
      .setBackground('#cfe2f3');
    rowIndex++;

    // ── Positions data ────────────────────────────────────────
    // Request COMPLETE view so the Quick quote block (changePct) is included
    var posUrl  = ACCOUNTS_URL + '/' + accountKey + '/portfolio.json?view=COMPLETE';
    var posHdrs = buildOAuthHeaders_(posUrl, 'GET', token, secret);
    var posResp = UrlFetchApp.fetch(posUrl, { method: 'get', muteHttpExceptions: true, headers: posHdrs });
    var posData = {};
    try {
      posData = JSON.parse(posResp.getContentText());
    } catch (e) {
      console.warn('Could not parse positions for ' + accountKey);
    }

    var portfolios = (posData.PortfolioResponse && posData.PortfolioResponse.AccountPortfolio) || [];
    if (portfolios.length === 0) {
      sheet.getRange(rowIndex, 1).setValue('No positions for ' + displayLabel);
      rowIndex++;
    } else {
      portfolios.forEach(function(portfolio) {
        var positions = portfolio.Position || [];

        // Sort by market value descending
        positions.sort(function(a, b) {
          return (b.marketValue || 0) - (a.marketValue || 0);
        });

        // ── Resolve change% from portfolio Quick object ───────
        // p.Quick.changePct is already a percentage (e.g. 0.82 = 0.82%)
        // Convert to decimal for consistent sheet formatting.
        // Collect symbols that have no Quick data for a fallback batch quote call.
        var changePctMap  = {};
        var missingSymbols = [];

        positions.forEach(function(p) {
          var sym = (p.Product && p.Product.symbol) || '';
          if (!sym) return;
          if (p.Quick && p.Quick.changePct !== undefined && p.Quick.changePct !== null) {
            changePctMap[sym] = p.Quick.changePct / 100;
          } else {
            missingSymbols.push(sym);
          }
        });

        // Fallback: batch-fetch from E*Trade market quote API for any missing
        if (missingSymbols.length > 0) {
          var fallback = fetchEtradeQuotes_(missingSymbols, token, secret);
          Object.keys(fallback).forEach(function(sym) {
            changePctMap[sym] = fallback[sym];
          });
        }

        var posRows = positions.map(function(p) {
          var symbol    = (p.Product && p.Product.symbol) || '';
          var qty       = p.quantity || 0;
          var avgPrice  = p.pricePaid || 0;
          var mv        = p.marketValue || 0;
          var costBasis = avgPrice * qty;
          var pl        = mv - costBasis;
          var plPct     = costBasis ? (pl / costBasis) : 0;
          var weight    = totalValue ? (mv / totalValue) : 0;
          var chgPct    = changePctMap[symbol] || 0;

          return [symbol, qty, chgPct, avgPrice, mv, pl, plPct, weight, accountId, accountKey];
        });

        if (posRows.length > 0) {
          // Write all rows at once (change% is now a real value, not a formula)
          sheet.getRange(rowIndex, 1, posRows.length, posHeader.length).setValues(posRows);

          // Number formats
          sheet.getRange(rowIndex, 2, posRows.length, 1).setNumberFormat('#,##0');       // Qty
          sheet.getRange(rowIndex, 3, posRows.length, 1).setNumberFormat('0.00%');       // Chg%
          sheet.getRange(rowIndex, 4, posRows.length, 3).setNumberFormat('#,##0.00');    // Avg/MV/P&L
          sheet.getRange(rowIndex, 7, posRows.length, 1).setNumberFormat('0.00%');       // P/L%
          sheet.getRange(rowIndex, 8, posRows.length, 1).setNumberFormat('0.00%');       // Weight
          sheet.getRange(rowIndex, 1, posRows.length, 8).setFontWeight('bold');

          // De-emphasize account key columns (I, J)
          sheet.getRange(rowIndex, 9, posRows.length, 2)
            .setFontColor('#666666')
            .setFontSize(9)
            .setBackground('#f5f5f5');

          // Conditional formatting
          applyPositionConditionalFormatting_(sheet, rowIndex, posRows.length);

          rowIndex += posRows.length;
        }
      });
    }

    rowIndex++; // spacer between accounts
  });

  // ── Totals row ────────────────────────────────────────────────
  if (accountRows.length > 0) {
    var totalValue = accountSummaries.reduce(function(sum, a) {
      return sum + (a.totalValue || 0);
    }, 0);

    var totalCash = accountSummaries.reduce(function(sum, a) {
      return sum + (a.cash || 0);
    }, 0);

    var totalAlloc = totalValue ? ((totalValue - totalCash) / totalValue) : 0;

    var totalStartValue = accountSummaries.reduce(function(sum, a) {
      var startValue = getEtradeStartOfYearNetLiq_(a.suffix);
      return sum + (startValue || 0);
    }, 0);

    var totalYtdPL = totalStartValue > 0
      ? (totalValue - totalStartValue) / totalStartValue
      : 0;

    sheet.getRange(rowIndex, 1, 1, 8).setValues([[
      'TOTALS',
      totalValue,
      totalCash,
      '',
      'ALLOC:',
      totalAlloc,
      'YTD P/L:',
      totalYtdPL
    ]]);

    sheet.getRange(rowIndex, 1).setBackground('#d9d9d9');
    sheet.getRange(rowIndex, 1, 1, 8)
      .setFontWeight('bold')
      .setFontSize(11);
    sheet.getRange(rowIndex, 2, 1, 8).setBackground('#d0f0c0');
    sheet.getRange(rowIndex, 2, 1, 2).setNumberFormat('#,##0.00');
    sheet.getRange(rowIndex, 6).setNumberFormat('0.00%');
    sheet.getRange(rowIndex, 8).setNumberFormat('0.00%');

    applyYtdConditionalFormatting_(sheet, rowIndex);
  }
}

// ─────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────

/** Sets Nunito font on the entire EtradeB sheet. */
function setSheetFont() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ETRADE);
  if (!sheet) return;
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setFontFamily('Nunito');
}

/**
 * Returns the ACCOUNT_ORDER suffix for a given full account ID,
 * or null if not found.
 */
function getAccountSuffix_(accountId) {
  var idStr = String(accountId || '').trim();

  var mappedSuffix = ACCOUNT_ORDER.find(function(suffix) {
    return (
      String(suffix) === idStr.slice(-4) ||
      (ACCOUNT_MAP[suffix] && String(ACCOUNT_MAP[suffix].id || '').trim() === idStr)
    );
  });

  return mappedSuffix || idStr.slice(-4) || null;
}
/** Applies conditional formatting to position columns. */
function applyPositionConditionalFormatting_(sheet, startRow, numRows) {
  var rules = sheet.getConditionalFormatRules();

  // P/L% (G): deep loss < -7% — light grey background + red font.
  // Must be pushed FIRST so it takes priority over the general loss rule below.
  var plPctRange = sheet.getRange(startRow, 7, numRows, 1);
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(-0.07)
      .setBackground('#d9d9d9')
      .setFontColor('red')
      .setBold(true)
      .setRanges([plPctRange])
      .build()
  );

  // Columns: C (Chg%), F (P/L), G (P/L%)
  [3, 6, 7].forEach(function(col) {
    var range = sheet.getRange(startRow, col, numRows, 1);

    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(0)
        .setFontColor('green')
        .setBold(true)
        .setRanges([range])
        .build()
    );

    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(0)
        .setFontColor('red')
        .setBold(true)
        .setRanges([range])
        .build()
    );

    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberEqualTo(0)
        .setFontColor('black')
        .setBold(true)
        .setRanges([range])
        .build()
    );
  });

  sheet.setConditionalFormatRules(rules);
}

/** Applies conditional formatting to YTD P/L cell in column H. */
function applyYtdConditionalFormatting_(sheet, rowNum) {
  var rules = sheet.getConditionalFormatRules();
  var range = sheet.getRange(rowNum, 8);

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setFontColor('#006400')
      .setBold(true)
      .setRanges([range])
      .build()
  );

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setFontColor('red')
      .setBold(true)
      .setRanges([range])
      .build()
  );

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberEqualTo(0)
      .setFontColor('black')
      .setBold(true)
      .setRanges([range])
      .build()
  );

  sheet.setConditionalFormatRules(rules);
}

/**
 * Returns the earliest plain-string current-year Net Liquidity value
 * for the given suffix from the shared Net Liquidity sheet.
 * Uses YYYY-MM-DD string comparison only.
 */
function getEtradeStartOfYearNetLiq_(suffix) {
  if (!suffix) return null;

  var map = getNetLiqMap_(suffix);
  if (!map || Object.keys(map).length === 0) {
    console.log('[ETRADE YTD DEBUG] No NetLiq data for suffix ' + suffix);
    return null;
  }

  var currentYear = String(new Date().getFullYear());
  var earliestDateStr = null;
  var earliestValue = null;

  Object.keys(map).forEach(function(dateStr) {
    dateStr = String(dateStr).trim();

    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return;
    if (!dateStr.startsWith(currentYear + '-')) return;

    var v = Number(map[dateStr]);
    if (!isFinite(v) || v <= 0) return;

    if (!earliestDateStr || dateStr < earliestDateStr) {
      earliestDateStr = dateStr;
      earliestValue = v;
    }
  });

  console.log(
    '[ETRADE YTD DEBUG] Suffix: ' + suffix +
    ' | StartDate: ' + earliestDateStr +
    ' | StartValue: ' + earliestValue
  );

  return earliestValue;
}

/**
 * YTD P/L % = (current value - first current-year net liq) / first current-year net liq
 */
function getEtradeYtdPLPercent_(accountId, currentValue) {
  if (!currentValue) return 0;

  var suffix = getAccountSuffix_(accountId);
  var startValue = getEtradeStartOfYearNetLiq_(suffix);

  var ytd = (!startValue || startValue <= 0)
    ? 0
    : (currentValue - startValue) / startValue;

  console.log(
    '[ETRADE YTD DEBUG] Account: ' + accountId +
    ' | Suffix: ' + suffix +
    ' | StartValue: ' + startValue +
    ' | CurrentValue: ' + currentValue +
    ' | YTD%: ' + ytd
  );
console.log('[ETRADE YTD DEBUG] Account: ' + accountId + ' | Resolved suffix: ' + suffix);
  return ytd;
  
}
