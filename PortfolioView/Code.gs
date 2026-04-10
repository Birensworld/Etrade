// ============================================================
// PORTFOLIO VIEW                                 Version: 1.5
// ============================================================


// ============================================================
// CONFIG
// ============================================================

const sources = [
  { url: "https://docs.google.com/spreadsheets/d/1zRmg5ccftgxrW2_67hiXb2UVoNXSK-d9HU9NbP3c_Q4/edit#gid=2003059889", tab: "Schwab" },
  { url: "https://docs.google.com/spreadsheets/d/1JIfX_MksEehXjUWnMwCHSYCdKBUUL8WkI2mZs822p9E/edit#gid=0",          tab: "EtradeB" },
  { url: "https://docs.google.com/spreadsheets/d/1nGGp8azsFEPra0Awis-mIHV42TmE0eB-Zh98PcweOC0/edit#gid=0",          tab: "EtradeA" }
];

// Only these account names are imported (case-sensitive)
const allowedAccounts = [
  "SW Equity",
  "SW IRA",
  "ETrade B",
  "ETrade B IRA",
  "ETrade A IRA"
];

// Override the sector returned by WISEPRICE for specific tickers
const sectorOverrides = {
  "SPY":   "SPY",
  "VOOG":  "SPY",
  "VOO":   "SPY",
  "UPRO":  "SPY",
  "RSP":   "SPY",
  "TECL":  "Technology",
  "SMH":   "Technology",
  "QQQ":   "Technology",
  "XLK":   "Technology",
  "QQEW":  "Technology",
  "QQQE":  "Technology",
  "IWM":   "Small Caps",
  "EQAL":  "Small Caps",
  "UWM":   "Small Caps",
  "XLE":   "Energy",
  "XLY":   "Consumer Discretionary",
  "BRK.B": "Berkshire",
  "IBIT":  "Bitcoin",
  "BTC":   "Bitcoin"
};

// ============================================================
// MENU
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Portfolio")
    .addItem("Get All Accounts", "mergePortfolioSheets")
    .addSeparator()
    .addItem("Refresh News Feed", "refreshNewsFeed")
    .addItem("Schedule Daily News (9 AM ET)", "createNewsFeedTrigger")
    .addItem("Set Finnhub API Key", "setFinnhubApiKey")
    .addSeparator()
    .addItem("Refresh Earnings Report", "refreshEarningsReport")
    .addItem("Schedule Earnings Report (8 AM & 5 PM ET)", "createEarningsTriggers")
    .addSeparator()
    .addItem("Refresh Top Ratings", "refreshTopRatings")
    .addItem("Schedule Top Ratings (9 AM ET)", "createTopRatingsTrigger")
    .addToUi();
}

// ============================================================
// MAIN
// ============================================================

function mergePortfolioSheets() {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName("Portfolio View") || ss.insertSheet("Portfolio View");
  targetSheet.clear();

  let targetRow  = 1;
  let accountRows = [];

  sources.forEach(src => {
    try {
      const srcSS = SpreadsheetApp.openById(extractSheetId(src.url));
      const sheet = srcSS.getSheetByName(src.tab);
      if (!sheet) return;

      const data = sheet.getDataRange().getValues();
      let rowIndex = 0;

      while (rowIndex < data.length) {
        const row         = data[rowIndex];
        const firstCell   = row[0] ? row[0].toString().trim() : "";
        const accountName = firstCell;

        if (!firstCell || firstCell === "Symbol") { rowIndex++; continue; }
        if (!allowedAccounts.includes(accountName)) { rowIndex++; continue; }

        // --- Green account header row ---
        const accountRowValues = row.slice(0, 10);
        while (accountRowValues.length < 10) accountRowValues.push("");
        targetSheet.getRange(targetRow, 1, 1, 10).setValues([accountRowValues]);
        targetSheet.getRange(targetRow, 1, 1, 10)
          .setBackground("#d9ead3")
          .setFontWeight("bold")
          .setFontSize(12);
        accountRows.push(targetRow);
        targetRow++;

        // --- Blue column-header row ---
        const colHeaders = ["Symbol", "Qty", "Chg %", "Avg Price", "MV", "P/L", "P/L %", "Sector", "Dividend", "Weight"];
        targetSheet.getRange(targetRow, 1, 1, colHeaders.length).setValues([colHeaders]);
        targetSheet.getRange(targetRow, 1, 1, colHeaders.length)
          .setBackground("#cfe2f3")
          .setFontWeight("bold")
          .setHorizontalAlignment("center")
          .setVerticalAlignment("middle");
        targetRow++;

        // Expect the source sheet's own header row next
        const nextRow = data[rowIndex + 1];
        if (!(nextRow && nextRow[0] && nextRow[0].toString().trim() === "Symbol")) {
          rowIndex++;
          continue;
        }

        rowIndex += 2; // skip source sheet header
        const startRow = targetRow;

        // --- Copy position rows ---
        while (
          rowIndex < data.length &&
          data[rowIndex][0] &&
          data[rowIndex][0].toString().trim() !== "Symbol" &&
          !allowedAccounts.includes(data[rowIndex][0].toString().trim())
        ) {
          const posRow = data[rowIndex];
          if (!posRow[0].toString().toLowerCase().includes("no positions")) {
            targetSheet.getRange(targetRow, 1, 1, 7).setValues([posRow.slice(0, 7)]);
            targetRow++;
          }
          rowIndex++;
        }

        const endRow  = targetRow - 1;
        const numRows = endRow - startRow + 1;

        // Wise formulas (spill range: one formula covers the whole block)
        if (numRows > 0) {
          targetSheet.getRange(startRow, 8).setFormula(`=WISEPRICE(A${startRow}:A${endRow},"Sector")`);
          targetSheet.getRange(startRow, 9).setFormula(`=WISE(A${startRow}:A${endRow},"Dividend Yield","ttm")`);
        }

        // Weight column (J)
        if (endRow >= startRow) {
          applyWeightsForBlock_(targetSheet, startRow, endRow, row[1], 10);
        }
      }
    } catch (e) {
      Logger.log("⚠ Error opening " + src.url + " → " + e.message);
    }
  });

  // --- Totals rows ---
  const [labelRow, valueRow] = addTotalsRows(targetSheet, accountRows);

  // --- Formatting passes ---
  applyFormatting(targetSheet, accountRows);
  formatTotalsRows_(targetSheet, labelRow, valueRow);
  enforceSignColorsAndAlerts(targetSheet, accountRows, labelRow, valueRow);
  cleanDividendColumn_(targetSheet, accountRows, labelRow, valueRow);

  // --- Sector overrides + pie chart ---
  applySectorOverrides(ss, targetSheet, valueRow);
}

// ============================================================
// FORMATTING HELPERS
// ============================================================

function applyFormatting(sheet, accountRows) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  sheet.getRange(1, 1, lastRow, 10).setFontWeight("bold").setFontFamily("Nunito");

  sheet.getRange(1, 2,  lastRow).setNumberFormat("#,##0");      // Qty
  sheet.getRange(1, 3,  lastRow).setNumberFormat("0.00%");       // Chg %
  sheet.getRange(1, 4,  lastRow).setNumberFormat("#,##0.00");   // Avg Price
  sheet.getRange(1, 5,  lastRow).setNumberFormat("#,##0");      // MV
  sheet.getRange(1, 6,  lastRow).setNumberFormat("#,##0.00");   // P/L
  sheet.getRange(1, 7,  lastRow).setNumberFormat("0.00%");       // P/L %
  sheet.getRange(1, 9,  lastRow).setNumberFormat("0.00%");       // Dividend
  sheet.getRange(1, 10, lastRow).setNumberFormat("0.00%");       // Weight

  // Account header rows have different formats in some columns
  accountRows.forEach(r => {
    sheet.getRange(r, 2).setNumberFormat("#,##0");   // Total MV
    sheet.getRange(r, 3).setNumberFormat("#,##0");   // Cash
    sheet.getRange(r, 6).setNumberFormat("0%");      // Alloc %
    sheet.getRange(r, 8).setNumberFormat("0.00%");   // YTD P/L %
  });

  // Conditional formatting: Chg % column green/red
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0).setFontColor("green")
      .setRanges([sheet.getRange(1, 3, lastRow)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0).setFontColor("red")
      .setRanges([sheet.getRange(1, 3, lastRow)]).build()
  ]);
}

function formatTotalsRows_(sheet, labelRow, valueRow) {
  [labelRow, valueRow].forEach(r => {
    sheet.getRange(r, 1, 1, 10)
      .setFontWeight("bold")
      .setBackground("#fff2cc")
      .setHorizontalAlignment("center")
      .setFontSize(12)
      .setVerticalAlignment("middle");
  });
  sheet.getRange(valueRow, 2).setNumberFormat("#,##0");   // MVTotal
  sheet.getRange(valueRow, 3).setNumberFormat("#,##0");   // Cash
  sheet.getRange(valueRow, 4).setNumberFormat("0.00%");   // CashPct
  sheet.getRange(valueRow, 5).setNumberFormat("#,##0");   // EquitySum
}

// Colors P/L (col F) and P/L % (col G); grey background when P/L % < -7%
function enforceSignColorsAndAlerts(sheet, accountRows, labelRow, valueRow) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return;

  const skip  = new Set([labelRow, valueRow, ...accountRows]);
  const toNum = v => {
    if (v === null || v === "") return NaN;
    if (typeof v === "number") return v;
    return parseFloat(String(v).replace(/[,%\s]/g, "").replace(/[^0-9.\-]/g, ""));
  };

  // P/L column (F = 6) — green / red font
  const rangeF = sheet.getRange(1, 6, lastRow, 1);
  const valsF  = rangeF.getValues();
  const fontsF = rangeF.getFontColors();
  for (let r = 0; r < valsF.length; r++) {
    if (skip.has(r + 1)) continue;
    const n = toNum(valsF[r][0]);
    if (isNaN(n)) continue;
    fontsF[r][0] = n > 0 ? "green" : n < 0 ? "red" : "black";
  }
  rangeF.setFontColors(fontsF);

  // P/L % column (G = 7) — green/red font + grey bg when < -7%
  // Values are stored as decimals (e.g. -0.08 = -8%), so threshold is -0.07
  const rangeG = sheet.getRange(1, 7, lastRow, 1);
  const valsG  = rangeG.getValues();
  const fontsG = rangeG.getFontColors();
  const bgsG   = rangeG.getBackgrounds();
  for (let r = 0; r < valsG.length; r++) {
    if (skip.has(r + 1)) continue;
    const n = toNum(valsG[r][0]);
    if (isNaN(n)) continue;
    bgsG[r][0]  = n < -0.07 ? "#d3d3d3" : "";
    fontsG[r][0] = n > 0 ? "green" : n < 0 ? "red" : "black";
  }
  rangeG.setBackgrounds(bgsG);
  rangeG.setFontColors(fontsG);
}

// Normalises the Dividend column: blank/N/A/0 → "0%", raw decimals → "X.XX%"
function cleanDividendColumn_(sheet, accountRows, labelRow, valueRow) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const skip    = new Set([labelRow, valueRow, ...accountRows]);
  const range   = sheet.getRange(1, 9, lastRow);
  const vals    = range.getValues();
  let   updated = false;

  for (let r = 0; r < vals.length; r++) {
    if (skip.has(r + 1)) continue;
    const raw = (vals[r][0] || "").toString().trim().toLowerCase();
    if (raw === "dividend") continue; // blue header row

    if (raw === "unavailable" || raw === "n/a" || raw === "" || raw === "0") {
      vals[r][0] = "0%";
      updated = true;
    } else {
      const num = parseFloat(raw);
      if (!isNaN(num) && num > 0 && num < 1) {
        vals[r][0] = (num * 100).toFixed(2) + "%";
        updated = true;
      }
    }
  }

  if (updated) range.setValues(vals);
}

// ============================================================
// DATA HELPERS
// ============================================================

function addTotalsRows(sheet, accountRows) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2 || accountRows.length === 0) return [null, null];

  const labelRow = lastRow + 1;
  const valueRow = lastRow + 2;

  sheet.getRange(labelRow, 1).setValue("TOTALS");
  sheet.getRange(labelRow, 2).setValue("MVTotal");
  sheet.getRange(labelRow, 3).setValue("Cash");
  sheet.getRange(labelRow, 4).setValue("CashPct");
  sheet.getRange(labelRow, 5).setValue("EquitySum");
  sheet.getRange(labelRow, 9).setValue(" ");
  sheet.getRange(valueRow,  9).setValue(" ");

  sheet.getRange(valueRow, 2).setFormula("=" + accountRows.map(r => `B${r}`).join("+"));
  sheet.getRange(valueRow, 3).setFormula("=" + accountRows.map(r => `C${r}`).join("+"));
  sheet.getRange(valueRow, 4).setFormula(`=C${valueRow}/B${valueRow}`);
  sheet.getRange(valueRow, 5).setFormula(`=SUM(E1:E${lastRow})`);

  return [labelRow, valueRow];
}

function applyWeightsForBlock_(sheet, startRow, endRow, accountTotal, weightCol) {
  if (endRow < startRow) return;

  const acctTotal = toNumber_(accountTotal);
  if (!acctTotal || isNaN(acctTotal) || acctTotal === 0) {
    sheet.getRange(startRow, weightCol, endRow - startRow + 1, 1).clearContent();
    return;
  }

  const mvVals  = sheet.getRange(startRow, 5, endRow - startRow + 1, 1).getValues();
  const weights = mvVals.map(([mv]) => {
    const n = toNumber_(mv);
    return isNaN(n) ? [""] : [n / acctTotal];
  });

  const range = sheet.getRange(startRow, weightCol, weights.length, 1);
  range.setValues(weights);
  range.setNumberFormat("0.00%");
}

// ============================================================
// SECTOR OVERRIDES + PIE CHART
// ============================================================

function applySectorOverrides(ss, sheet, valueRow) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;

  const symbols = sheet.getRange(1, 1, lastRow).getValues();
  const sectors = sheet.getRange(1, 8, lastRow).getValues();

  for (let r = 0; r < lastRow; r++) {
    const sym = (symbols[r][0] || "").toString().trim().toUpperCase();
    if (sectorOverrides[sym]) sectors[r][0] = sectorOverrides[sym];
  }
  sheet.getRange(1, 8, lastRow).setValues(sectors);

  createSectorPieChart(ss, sheet, lastRow, valueRow);
}

function createSectorPieChart(ss, targetSheet, lastRow, valueRow) {
  const symbols = targetSheet.getRange(1, 1, lastRow).getValues();
  const sectors = targetSheet.getRange(1, 8, lastRow).getValues();
  const mvs     = targetSheet.getRange(1, 5, lastRow).getValues();

  const sectorTotals = {};
  for (let r = 0; r < lastRow; r++) {
    const sym    = (symbols[r][0] || "").toString().trim();
    const sector = (sectors[r][0] || "").toString().trim();
    const mv     = parseFloat(mvs[r][0]) || 0;
    if (!sym || sym === "Symbol" || sym === "TOTALS" || !sector) continue;
    sectorTotals[sector] = (sectorTotals[sector] || 0) + mv;
  }

  // Add cash from the totals row
  const cashValue = targetSheet.getRange(valueRow, 3).getValue();
  if (cashValue > 0) sectorTotals["Cash"] = cashValue;

  // Write Sector Summary sheet
  let summarySheet = ss.getSheetByName("Sector Summary");
  if (!summarySheet) {
    summarySheet = ss.insertSheet("Sector Summary");
  } else {
    summarySheet.clear();
    summarySheet.getCharts().forEach(c => summarySheet.removeChart(c));
  }

  summarySheet.getRange(1, 1, 1, 2)
    .setValues([["Sector", "Market Value"]])
    .setBackground("#d9ead3")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontFamily("Nunito");

  const rows = Object.entries(sectorTotals).map(([sector, total]) => [sector, total]);
  if (rows.length > 0) {
    summarySheet.getRange(2, 1, rows.length, 2).setValues(rows);
    summarySheet.getRange(2, 2, rows.length, 1).setNumberFormat("#,##0");
    summarySheet.getRange(1, 1, rows.length + 1, 2)
      .setFontWeight("bold")
      .setFontFamily("Nunito");
  }

  const chart = summarySheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(summarySheet.getRange(1, 1, rows.length + 1, 2))
    .setPosition(2, 4, 0, 0)
    .setOption("title", "Portfolio by Sector (Market Value + Cash)")
    .setOption("pieHole", 0.3)
    .build();

  summarySheet.insertChart(chart);
}

// ============================================================
// UTILITIES
// ============================================================

function toNumber_(v) {
  if (v === null || v === "") return NaN;
  if (typeof v === "number") return v;
  return parseFloat(String(v).replace(/,/g, "").replace(/[%\s]/g, ""));
}

function extractSheetId(url) {
  const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (match) return match[1];
  throw new Error("Invalid Google Sheets URL: " + url);
}
