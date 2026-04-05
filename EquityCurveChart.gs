/**
 * EquityCurveChart.gs — Builds a per-account equity curve chart.
 * Version: 2.2 (2026-04-05) — Ported from Schwab. QQQ red; 15th-of-month markers; yearly chart with year prompt.
 *
 * Chart sheet layout (7 columns):
 *   Date | NL% | SPY% | QQQ% | NL Marker | SPY Marker | QQQ Marker
 *
 *   Series 0–2 : solid lines, no individual point markers
 *   Series 3–5 : month-end + 15th-of-month dots (lineWidth=0), value labels, color-matched
 *
 * Colors: NL = black (#000000), SPY = green (#34a853), QQQ = red (#ea4335)
 *
 * IMPORTANT — flat dot-notation setOption() ONLY.
 * Never pass nested objects to setOption(); it silently breaks the chart.
 */

function buildEquityCurveChartForAccount(suffix) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    ss.toast('Building equity curve for account …' + suffix + '…', 'Working', -1);

    var netLiqMap  = getNetLiqMap_(suffix);
    var benchmarks = getBenchmarkCloseMaps_();
    var spyMap     = benchmarks.spy;
    var qqqMap     = benchmarks.qqq;

    if (Object.keys(netLiqMap).length === 0) {
      throw new Error(
        'No Net Liquidity data found for account …' + suffix + '.\n\n' +
        'Enter values in the "' + SHEET_NET_LIQ + '" sheet first.'
      );
    }
    if (Object.keys(spyMap).length === 0) {
      throw new Error('No SPY/QQQ data found. Run "Fetch SPY + QQQ History" first.');
    }

    var commonDates = Object.keys(netLiqMap)
      .filter(function(d) { return spyMap[d]; })
      .sort();

    if (commonDates.length < 2) {
      throw new Error(
        'Not enough overlapping dates between Net Liq …' + suffix + ' and SPY.\n' +
        'Make sure both datasets share common trading days.'
      );
    }

    var allNLDates  = Object.keys(netLiqMap).sort();
    var allSPYDates = Object.keys(spyMap).sort();

    var baseNLDate  = allNLDates.find(function(d)  { return d >= HISTORY_START; }) || allNLDates[0];
    var baseSPYDate = allSPYDates.find(function(d) { return d >= HISTORY_START; }) || allSPYDates[0];

    var baseDate   = baseNLDate;
    var baseNetLiq = netLiqMap[baseNLDate];
    var baseSPY    = spyMap[baseSPYDate];
    var baseQQQ    = qqqMap[baseSPYDate] || null;

    var markers = getMarkerDates_(commonDates);
    var rows    = buildRows_(commonDates, netLiqMap, spyMap, qqqMap,
                             baseNetLiq, baseSPY, baseQQQ, markers);

    var sheetName  = chartSheetName_(suffix);
    var chartSheet = getOrCreateChartSheet_(ss, sheetName);
    writeChartData_(chartSheet, rows, suffix, baseDate);
    SpreadsheetApp.flush();
    insertLineChart_(chartSheet, rows,
      '% Return — Portfolio …' + suffix + ' vs. SPY & QQQ');

    ss.setActiveSheet(chartSheet);
    ss.toast(
      '✅ ' + rows.length + ' days  (' + baseDate + ' → ' + commonDates[commonDates.length - 1] + ')',
      'Equity Curve …' + suffix, 10
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert('Chart Error – Account ' + suffix, e.message,
      SpreadsheetApp.getUi().ButtonSet.OK);
    console.error(e);
  }
}

function promptAndBuildYearlyEquityCurve_(suffix) {
  var ui     = SpreadsheetApp.getUi();
  var result = ui.prompt(
    'Equity Curve – Yearly View – Account …' + suffix,
    'Enter a 4-digit year (e.g. ' + new Date().getFullYear() + '):',
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() !== ui.Button.OK) return;

  var year = result.getResponseText().trim();
  if (!/^\d{4}$/.test(year)) {
    ui.alert('Invalid year', 'Please enter a 4-digit year such as 2026.', ui.ButtonSet.OK);
    return;
  }
  buildEquityCurveChartForYear_(suffix, year);
}

function buildEquityCurveChartForYear_(suffix, year) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    ss.toast('Building ' + year + ' equity curve for account …' + suffix + '…', 'Working', -1);

    var netLiqMap  = getNetLiqMap_(suffix);
    var benchmarks = getBenchmarkCloseMaps_();
    var spyMap     = benchmarks.spy;
    var qqqMap     = benchmarks.qqq;

    var yStart = year + '-01-01';
    var yEnd   = year + '-12-31';

    var commonDates = Object.keys(netLiqMap)
      .filter(function(d) { return spyMap[d] && d >= yStart && d <= yEnd; })
      .sort();

    if (commonDates.length === 0) {
      ss.toast('', '', 1);
      SpreadsheetApp.getUi().alert(
        'No Data for ' + year,
        'No overlapping Net Liquidity and SPY/QQQ data found for ' + year + '.\n\n' +
        'Make sure you have:\n' +
        '  • Net Liquidity entries in "' + SHEET_NET_LIQ + '" for ' + year + '\n' +
        '  • SPY history fetched for ' + year,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    if (commonDates.length < 2) {
      throw new Error('Only 1 data point found for ' + year + ' — need at least 2 to draw a chart.');
    }

    var baseDate   = commonDates[0];
    var baseNetLiq = netLiqMap[baseDate];
    var baseSPY    = spyMap[baseDate];
    var baseQQQ    = qqqMap[baseDate] || null;

    var markers = getMarkerDates_(commonDates);
    var rows    = buildRows_(commonDates, netLiqMap, spyMap, qqqMap,
                             baseNetLiq, baseSPY, baseQQQ, markers);

    var sheetName  = 'Equity Curve ' + suffix + ' (' + year + ')';
    var chartSheet = getOrCreateChartSheet_(ss, sheetName);
    writeChartData_(chartSheet, rows, suffix, baseDate);
    SpreadsheetApp.flush();
    insertLineChart_(chartSheet, rows,
      '% Return ' + year + ' — Portfolio …' + suffix + ' vs. SPY & QQQ');

    ss.setActiveSheet(chartSheet);
    ss.toast(
      '✅ ' + rows.length + ' days  (' + baseDate + ' → ' + commonDates[commonDates.length - 1] + ')',
      'Equity Curve …' + suffix + ' (' + year + ')', 10
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert('Chart Error – Account ' + suffix + ' (' + year + ')',
      e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    console.error(e);
  }
}

function buildRows_(dates, netLiqMap, spyMap, qqqMap,
                    baseNetLiq, baseSPY, baseQQQ, markers) {
  return dates.map(function(d) {
    var nlPct  = roundTo2_((netLiqMap[d] / baseNetLiq - 1) * 100);
    var spyPct = roundTo2_((spyMap[d]    / baseSPY    - 1) * 100);
    var qqqPct = (baseQQQ && qqqMap[d])
      ? roundTo2_((qqqMap[d] / baseQQQ - 1) * 100)
      : '';

    var isMark = markers[d];
    return [
      d,
      nlPct,
      spyPct,
      qqqPct,
      isMark             ? nlPct  : '',
      isMark             ? spyPct : '',
      (isMark && qqqPct !== '') ? qqqPct : '',
    ];
  });
}

function getOrCreateChartSheet_(ss, sheetName) {
  var existing = ss.getSheetByName(sheetName);
  if (existing) ss.deleteSheet(existing);
  return ss.insertSheet(sheetName);
}

function writeChartData_(sheet, rows, suffix, baseDate) {
  var hdr = sheet.getRange(1, 1, 1, 7);
  hdr.setValues([[
    'Date',
    'Portfolio …' + suffix + ' (% Return)',
    'SPY (% Return)',
    'QQQ (% Return)',
    'NL Marker',
    'SPY Marker',
    'QQQ Marker',
  ]]);
  hdr.setFontWeight('bold').setBackground('#0b5394').setFontColor('#ffffff')
     .setHorizontalAlignment('center');

  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 120);
  [2, 3, 4].forEach(function(c) { sheet.setColumnWidth(c, 185); });
  [5, 6, 7].forEach(function(c) { sheet.setColumnWidth(c, 120); });

  sheet.getRange(2, 1, rows.length, 7).setValues(rows);
  sheet.getRange(2, 2, rows.length, 3).setNumberFormat('0.00"%"');
  sheet.getRange(2, 5, rows.length, 3).setNumberFormat('+0.00"%";-0.00"%";"0%"');

  for (var i = 0; i < rows.length; i++) {
    if (i % 2 === 0) sheet.getRange(2 + i, 1, 1, 7).setBackground('#f8f9fa');
  }

  var metaRow = rows.length + 3;
  sheet.getRange(metaRow, 1, 2, 2).setValues([
    ['Account',   '…' + suffix],
    ['Base Date', baseDate],
  ]);
  sheet.getRange(metaRow, 1, 2, 2).setFontColor('#888888').setFontStyle('italic');
}

function insertLineChart_(sheet, rows, title) {
  var n = rows.length + 1;

  var allPcts = [];
  rows.forEach(function(r) {
    [r[1], r[2], r[3]].forEach(function(v) {
      if (typeof v === 'number') allPcts.push(v);
    });
  });
  var dataMin = allPcts.length ? Math.min.apply(null, allPcts) : -5;
  var dataMax = allPcts.length ? Math.max.apply(null, allPcts) : 20;
  var padding = Math.max((dataMax -
