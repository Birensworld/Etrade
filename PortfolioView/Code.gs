// 🔗 CONFIG: paste your sheet URLs + tab names here
const sources = [
  {url: "https://docs.google.com/spreadsheets/d/1zRmg5ccftgxrW2_67hiXb2UVoNXSK-d9HU9NbP3c_Q4/edit#gid=2003059889", tab: "Schwab"},
  {url: "https://docs.google.com/spreadsheets/d/1JIfX_MksEehXjUWnMwCHSYCdKBUUL8WkI2mZs822p9E/edit#gid=0", tab: "EtradeB"},
  {url: "https://docs.google.com/spreadsheets/d/1nGGp8azsFEPra0Awis-mIHV42TmE0eB-Zh98PcweOC0/edit#gid=0", tab: "EtradeA"}
];

// ✅ Only allow these accounts (case-sensitive mapping)
const allowedAccounts = [
  "SW Equity",
  "SW IRA",
  "ETrade B",
  "ETrade B IRA",
  "ETrade A IRA"
];

// ✅ Custom sector overrides
const sectorOverrides = {
  "SPY": "SPY",
  "VOOG": "SPY",
  "VOO": "SPY",
  "TECL": "Technology",
  "UPRO": "SPY",
  "SMH": "Technology",
  "QQQ": "Technology",
  "IWM": "Small Caps",
  "EQAL": "Small Caps",
  "UWM": "Small Caps",
  "XLE": "Energy",
  "XLY": "Consumer Discretionary",
  "XLK": "Technology",
  "BRK.B": "Berkshire",
  "IBIT": "Bitcoin",
  "BTC": "Bitcoin",
  "QQEW": "Technology",
  "QQQE": "Technology",
  "RSP": "SPY"
};

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Portfolio")
    .addItem("Get All Accounts", "mergePortfolioSheets")
    .addToUi();
}

function mergePortfolioSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName("Portfolio View") || ss.insertSheet("Portfolio View");
  targetSheet.clear();

  let targetRow = 1;
  let accountRows = [];

  sources.forEach(src => {
    try {
      const id = extractSheetId(src.url);
      const srcSS = SpreadsheetApp.openById(id);
      const sheet = srcSS.getSheetByName(src.tab);
      if (!sheet) return;

      const data = sheet.getDataRange().getValues();

      let rowIndex = 0;
      while (rowIndex < data.length) {
        const row = data[rowIndex];

        // Detect account row
        if (row[0] && row[0].toString().trim() !== "Symbol") {
          const accountName = row[0].toString().trim();

          if (!allowedAccounts.includes(accountName)) {
            rowIndex++;
            continue;
          }

          // Account row (green)
          const accountRowValues = row.slice(0, 10);
          while (accountRowValues.length < 10) accountRowValues.push("");

          targetSheet.getRange(targetRow, 1, 1, 10).setValues([accountRowValues]);
          targetSheet.getRange(targetRow, 1, 1, 10)
            .setBackground("#d9ead3")
            .setFontWeight("bold")
            .setFontSize(12);
          accountRows.push(targetRow);
          targetRow++;

          // Blue header row
          const posHeader = ["Symbol", "Qty", "Chg %", "Avg Price", "MV", "P/L", "P/L %", "Sector", "Dividend", "Weight"];
          targetSheet.getRange(targetRow, 1, 1, posHeader.length).setValues([posHeader]);
          targetSheet.getRange(targetRow, 1, 1, posHeader.length)
            .setBackground("#cfe2f3")
            .setFontWeight("bold")
            .setHorizontalAlignment("center")
            .setVerticalAlignment("middle");
          targetRow++;

          // Expect header in next row
          const header = data[rowIndex + 1];
          if (header && header[0] && header[0].toString().trim() === "Symbol") {
            rowIndex += 2; // skip sheet header

            const startRow = targetRow; // first position row in Portfolio View

            // Copy positions
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

            const endRow = targetRow - 1;

            // Apply Wise formulas as range
            const numRows = endRow - startRow + 1;
            if (numRows > 0) {
              const sectorFormula = `=WISEPRICE(A${startRow}:A${endRow},"Sector")`;
              const dividendFormula = `=WISE(A${startRow}:A${endRow},"Dividend Yield","ttm")`;
              targetSheet.getRange(startRow, 8).setFormula(sectorFormula);
              targetSheet.getRange(startRow, 9).setFormula(dividendFormula);
            }

            // --- WEIGHTS (Column J = 10) for this account block ---
            const WEIGHT_COL = 10; // J
            const accountTotal = row[1]; // Column B from the green account header row

            if (endRow >= startRow) {
              applyWeightsForBlock_(targetSheet, startRow, endRow, accountTotal, WEIGHT_COL);
            }

          } else {
            rowIndex++;
          }
        } else {
          rowIndex++;
        }
      }
    } catch (e) {
      Logger.log("⚠ Error opening " + src.url + " → " + e.message);
    }
  });

  // Add Totals rows
  const [labelRow, valueRow] = addTotalsRows(targetSheet, accountRows);

  // Base formatting
  applyFormatting(targetSheet, accountRows);

  // Totals formatting
  [labelRow, valueRow].forEach(r => {
    const range = targetSheet.getRange(r, 1, 1, 10);
    range.setFontWeight("bold")
         .setBackground("#fff2cc")
         .setHorizontalAlignment("center")
         .setFontSize(12)
         .setVerticalAlignment("middle");
  });
  targetSheet.getRange(valueRow, 2).setNumberFormat("#,##0");   // MVTotal
  targetSheet.getRange(valueRow, 3).setNumberFormat("#,##0");   // Cash
  targetSheet.getRange(valueRow, 4).setNumberFormat("0.00%");   // CashPct
  targetSheet.getRange(valueRow, 5).setNumberFormat("#,##0");   // EquitySum

  // Explicit IF/ELSE coloring
  enforceSignColorsAndAlerts(targetSheet, accountRows, labelRow, valueRow);

  // 🔄 Clean up Dividend column (I) — single batch, fast
  (function cleanDividendColumn() {
    const lastRow = targetSheet.getLastRow();
    if (lastRow < 2) return;

    const divRange = targetSheet.getRange(1, 9, lastRow);
    const divVals = divRange.getValues();
    let updated = false;

    const skipRows = new Set([labelRow, valueRow, ...accountRows]);

    for (let r = 0; r < divVals.length; r++) {
      const rowNum = r + 1;
      if (skipRows.has(rowNum)) continue;

      let val = (divVals[r][0] || "").toString().trim().toLowerCase();

      // Skip blue header rows
      if (val === "dividend") continue;

      // Normalize invalid values
      if (val === "unavailable" || val === "n/a" || val === "" || val === "0") {
        divVals[r][0] = "0%";
        updated = true;
      } else if (!isNaN(parseFloat(val)) && parseFloat(val) > 0 && parseFloat(val) < 1) {
        // Convert decimals like 0.0123 -> 1.23%
        const pct = (parseFloat(val) * 100).toFixed(2) + "%";
        divVals[r][0] = pct;
        updated = true;
      }
    }

    if (updated) divRange.setValues(divVals);
  })();

  // ==================================================================
  // Sector overrides + build chart with Cash included
  // ==================================================================
  applySectorOverrides(ss, targetSheet, valueRow);
}

// Helper for Weights Column
function applyWeightsForBlock_(sheet, startRow, endRow, accountTotal, weightCol) {
  if (!startRow || !endRow || endRow < startRow) return;

  const acctTotalNum = toNumber_(accountTotal);
  if (!acctTotalNum || isNaN(acctTotalNum) || acctTotalNum === 0) {
    sheet.getRange(startRow, weightCol, endRow - startRow + 1, 1).clearContent();
    return;
  }

  const mvVals = sheet.getRange(startRow, 5, endRow - startRow + 1, 1).getValues();

  const weights = mvVals.map(([mv]) => {
    const mvNum = toNumber_(mv);
    if (!mvNum || isNaN(mvNum)) return [""];
    return [mvNum / acctTotalNum];
  });

  const weightRange = sheet.getRange(startRow, weightCol, weights.length, 1);
  weightRange.setValues(weights);
  weightRange.setNumberFormat("0.00%");
}

// Shared numeric cleaner
function toNumber_(v) {
  if (v === null || v === "") return NaN;
  if (typeof v === "number") return v;
  const s = String(v).replace(/,/g, "").replace(/[%\s]/g, "");
  const n = parseFloat(s);
  return n;
}

// Totals rows
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

  let cashSumFormula = accountRows.map(r => `C${r}`).join("+");
  sheet.getRange(valueRow, 3).setFormula("=" + cashSumFormula);

  let mvHeaderSumFormula = accountRows.map(r => `B${r}`).join("+");
  sheet.getRange(valueRow, 2).setFormula("=" + mvHeaderSumFormula);

  sheet.getRange(valueRow, 4).setFormula(`=C${valueRow}/B${valueRow}`);
  sheet.getRange(valueRow, 5).setFormula(`=SUM(E1:E${lastRow})`);

  sheet.getRange(valueRow, 9).setValue(" ");
  sheet.getRange(labelRow, 9).setValue(" ");

  return [labelRow, valueRow];
}

// Formatting
function applyFormatting(sheet, accountRows) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  sheet.getRange(1, 1, lastRow, 10)
    .setFontWeight("bold")
    .setFontFamily("Nunito");

  sheet.getRange(1, 2, lastRow).setNumberFormat("#,##0");     // Qty / header total
  sheet.getRange(1, 3, lastRow).setNumberFormat("0.00%");      // Chg % / default
  sheet.getRange(1, 4, lastRow).setNumberFormat("#,##0.00");  // Avg Price / default
  sheet.getRange(1, 5, lastRow).setNumberFormat("#,##0");     // MV
  sheet.getRange(1, 6, lastRow).setNumberFormat("#,##0.00");  // P/L / default
  sheet.getRange(1, 7, lastRow).setNumberFormat("0.00%");      // P/L % / default
  sheet.getRange(1, 10, lastRow).setNumberFormat("0.00%");    // Weight

  // Account header row overrides
  accountRows.forEach(r => {
    sheet.getRange(r, 2).setNumberFormat("#,##0");   // Total value in B
    sheet.getRange(r, 3).setNumberFormat("#,##0");   // Cash in C
    sheet.getRange(r, 6).setNumberFormat("0%");      // Alloc % in F
    sheet.getRange(r, 8).setNumberFormat("0.00%");   // YTD P/L % in H
  });

  // Dividend column → percent
  sheet.getRange(1, 9, lastRow).setNumberFormat("0.00%");

  const rules = [];
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0).setFontColor("green")
    .setRanges([sheet.getRange(1, 3, lastRow)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0).setFontColor("red")
    .setRanges([sheet.getRange(1, 3, lastRow)]).build());
  sheet.setConditionalFormatRules(rules);
}

// P/L coloring
function enforceSignColorsAndAlerts(sheet, accountRows, labelRow, valueRow) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return;

  const skip = new Set([labelRow, valueRow, ...accountRows]);
  const toNum = v => {
    if (v === null || v === "") return NaN;
    if (typeof v === "number") return v;
    const s = String(v).replace(/[,%\s]/g, "").replace(/[^0-9.\-]/g, "");
    return parseFloat(s);
  };

  const rangeF = sheet.getRange(1, 6, lastRow, 1);
  const valsF = rangeF.getValues();
  const fontsF = rangeF.getFontColors();
  for (let r = 0; r < valsF.length; r++) {
    const rowNum = r + 1;
    if (skip.has(rowNum)) continue;
    const n = toNum(valsF[r][0]);
    if (isNaN(n)) continue;
    fontsF[r][0] = n > 0 ? "green" : (n < 0 ? "red" : "black");
  }
  rangeF.setFontColors(fontsF);

  const rangeG = sheet.getRange(1, 7, lastRow, 1);
  const valsG = rangeG.getValues();
  const fontsG = rangeG.getFontColors();
  const bgsG = rangeG.getBackgrounds();
  for (let r = 0; r < valsG.length; r++) {
    const rowNum = r + 1;
    if (skip.has(rowNum)) continue;
    const n = toNum(valsG[r][0]);
    if (isNaN(n)) continue;
    bgsG[r][0]  = n < -7 ? "#d3d3d3" : "";
    fontsG[r][0] = n > 0 ? "green" : (n < 0 ? "red" : "black");
  }
  rangeG.setBackgrounds(bgsG);
  rangeG.setFontColors(fontsG);
}

// Apply sector overrides + build chart
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

// ✅ Create Sector Summary + Pie Chart (includes Cash)
function createSectorPieChart(ss, targetSheet, lastRow, valueRow) {
  const symbols = targetSheet.getRange(1, 1, lastRow).getValues();
  const sectors = targetSheet.getRange(1, 8, lastRow).getValues();
  const mvs = targetSheet.getRange(1, 5, lastRow).getValues();

  const sectorTotals = {};
  for (let r = 0; r < lastRow; r++) {
    const sym = (symbols[r][0] || "").toString().trim();
    const sector = (sectors[r][0] || "").toString().trim();
    const mv = parseFloat(mvs[r][0]) || 0;
    if (!sym || sym === "Symbol" || sym === "TOTALS" || !sector) continue;
    if (!sectorTotals[sector]) sectorTotals[sector] = 0;
    sectorTotals[sector] += mv;
  }

  // Include Cash from totals row
  const cashValue = targetSheet.getRange(valueRow, 3).getValue();
  if (cashValue && cashValue > 0) sectorTotals["Cash"] = cashValue;

  let summarySheet = ss.getSheetByName("Sector Summary");
  if (!summarySheet) summarySheet = ss.insertSheet("Sector Summary");
  else {
    summarySheet.clear();
    summarySheet.getCharts().forEach(c => summarySheet.removeChart(c));
  }

  // Header row styling
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
    summarySheet.getRange(1, 1, rows.length + 1, 2).setFontWeight("bold"); // +1 for header
    summarySheet.getRange(1, 1, rows.length + 1, 2).setFontFamily("Nunito");
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

// Helper
function extractSheetId(url) {
  const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (match && match[1]) return match[1];
  throw new Error("Invalid Google Sheets URL: " + url);
}
