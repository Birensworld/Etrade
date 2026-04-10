// ============================================================
// TOP ANALYST RATINGS                            Version: 1.1
// Refreshes daily at 9 AM ET via a time-based trigger.
//
// Shows up to 50 recent upgrade / downgrade / initiation actions
// across the broad market — not limited to your portfolio.
//
// Data sources (tried in order):
//   1. Finnhub  — market-wide stock/upgrade-downgrade (no symbol filter)
//   2. Yahoo Finance — per-ticker for a curated watchlist (fallback)
//
// Shared utilities (finnhubGet_, formatDate_, FINNHUB_KEY_PROP)
// are defined in NewsFeed.gs and available in the same GAS scope.
// ============================================================

const TOP_RATINGS_SHEET = "Top Ratings";

// Curated watchlist used when Finnhub market-wide returns < 10 rows.
// Covers the most-followed US stocks across sectors.
const TOP_RATINGS_WATCHLIST = [
  // Mega-cap tech / AI
  "AAPL","MSFT","NVDA","GOOGL","AMZN","META","TSLA","AVGO","ORCL","AMD",
  // Cloud / SaaS
  "CRM","SNOW","PLTR","NOW","ADBE","INTU","WDAY","MDB","DDOG","ZS","PANW",
  // Semiconductors
  "INTC","QCOM","AMAT","LRCX","KLAC","MRVL","ARM","SMCI","TXN","MU","ON",
  // Finance
  "JPM","BAC","GS","MS","BLK","V","MA","PYPL","AXP","COF","WFC","C",
  // Healthcare / Biotech
  "JNJ","UNH","LLY","PFE","MRK","ABBV","BMY","AMGN","GILD","REGN","MRNA",
  // Consumer
  "WMT","COST","TGT","NKE","MCD","SBUX","HD","LOW","TJX","BABA","SHOP",
  // Energy
  "XOM","CVX","COP","SLB","OXY","PSX","VLO","MPC","HES","EOG",
  // Industrial / Defense / Aerospace
  "GE","CAT","BA","RTX","LMT","NOC","HON","UNP","CSX","DE","EMR",
  // Telecom / Media / Streaming
  "T","VZ","NFLX","DIS","CMCSA","CHTR","SPOT",
  // EV / Clean Energy
  "RIVN","F","GM","ENPH","FSLR",
  // Other high-profile
  "UBER","LYFT","ABNB","COIN","SQ","RBLX","SNAP","PINS","HOOD"
];

// ============================================================
// ENTRY POINT
// ============================================================

function refreshTopRatings() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const apiKey = PropertiesService.getScriptProperties().getProperty(FINNHUB_KEY_PROP);
  if (!apiKey) {
    SpreadsheetApp.getUi().alert(
      "Finnhub API key not set.\n\nRun: setFinnhubApiKey()  (Extensions → Apps Script → Run)"
    );
    return;
  }

  const today   = formatDate_(new Date());
  const from30d = formatDate_(new Date(new Date() - 30 * 864e5));

  // ── Layer 1: Finnhub market-wide (no symbol = all stocks in their DB) ──
  // Note: free tier often returns empty here; Yahoo fallback handles that case.
  let rows = fetchTopRatingsFromFinnhub_(from30d, today, apiKey);
  Logger.log("Finnhub market-wide returned " + rows.length + " directional ratings");

  // ── Layer 2: Yahoo Finance per watchlist ticker ──
  // Always runs alongside Finnhub to supplement coverage.
  // Returns most-recent directional rating per ticker (no date cutoff —
  // the Date column shows when the action happened).
  if (rows.length < 50) {
    Logger.log("Supplementing with Yahoo Finance watchlist fetch …");
    const yahooRows = fetchTopRatingsFromYahoo_(TOP_RATINGS_WATCHLIST);
    const seen = new Set(rows.map(r => r[0]));
    yahooRows.forEach(r => { if (!seen.has(r[0])) { rows.push(r); seen.add(r[0]); } });
    rows.sort((a, b) => (b[5] || "").localeCompare(a[5] || ""));
    rows = rows.slice(0, 50);
  }

  buildTopRatingsSheet_(ss, rows, today);
  Logger.log("Top Ratings refreshed — " + rows.length + " entries written.");
}

// ============================================================
// DATA FETCHING
// ============================================================

// Finnhub: query without a symbol returns all upgrade/downgrade events globally.
// On the free tier this often returns an empty array — Yahoo covers the gap.
// Returns array of [ticker, action, firm, fromGrade, toGrade, date] rows.
function fetchTopRatingsFromFinnhub_(from, to, apiKey) {
  const ACTION_MAP = { up: "Upgrade", down: "Downgrade", init: "Initiates" };
  try {
    const resp = finnhubGet_(`stock/upgrade-downgrade?from=${from}&to=${to}`, apiKey);
    if (!Array.isArray(resp) || resp.length === 0) return [];

    return resp
      .filter(r => ACTION_MAP[(r.action || "").toLowerCase()])
      .sort((a, b) => (b.gradeDate || "").localeCompare(a.gradeDate || ""))
      .slice(0, 50)
      .map(r => [
        (r.symbol    || "").toUpperCase(),
        ACTION_MAP[(r.action || "").toLowerCase()],
        (r.company   || "").trim(),
        ((r.fromGrade || "") || "—").trim(),
        ((r.toGrade   || "") || "—").trim(),
        (r.gradeDate || "").trim()
      ]);
  } catch (e) {
    Logger.log("TopRatings Finnhub: " + e.message);
    return [];
  }
}

// Yahoo Finance: fetch the most-recent directional rating for each watchlist ticker.
// No date cutoff — returns whatever the latest upgrade/downgrade/initiation was,
// with the Date column showing when it happened. Sorted by date desc at the caller.
// Returns the same [ticker, action, firm, fromGrade, toGrade, date] row format.
function fetchTopRatingsFromYahoo_(tickers) {
  const ACTION_MAP = { up: "Upgrade", down: "Downgrade", init: "Initiates" };
  const rows = [];

  tickers.forEach(ticker => {
    try {
      const url = "https://query2.finance.yahoo.com/v10/finance/quoteSummary/" +
                  encodeURIComponent(ticker) +
                  "?modules=upgradeDowngradeHistory&corsDomain=finance.yahoo.com";
      const resp = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        headers: { "User-Agent": "Mozilla/5.0" }
      });
      if (resp.getResponseCode() !== 200) return;

      const history =
        JSON.parse(resp.getContentText())
          ?.quoteSummary?.result?.[0]
          ?.upgradeDowngradeHistory?.history;
      if (!Array.isArray(history)) return;

      // Most-recent directional action — no date cutoff
      const entry = history.find(h =>
        ACTION_MAP[(h.action || "").toLowerCase()] && h.epochGradeDate
      );
      if (!entry) return;

      const date = Utilities.formatDate(
        new Date(entry.epochGradeDate * 1000), "America/New_York", "yyyy-MM-dd"
      );
      rows.push([
        ticker,
        ACTION_MAP[(entry.action || "").toLowerCase()],
        (entry.firm       || "").trim(),
        ((entry.fromGrade || "") || "—").trim(),
        ((entry.toGrade   || "") || "—").trim(),
        date
      ]);
    } catch (e) {
      Logger.log("TopRatings Yahoo " + ticker + ": " + e.message);
    }
    Utilities.sleep(200);
  });

  return rows;
}

// ============================================================
// SHEET BUILDER
// ============================================================

function buildTopRatingsSheet_(ss, rows, dateStr) {
  let sheet = ss.getSheetByName(TOP_RATINGS_SHEET);
  if (!sheet) sheet = ss.insertSheet(TOP_RATINGS_SHEET);
  else sheet.clear();

  sheet.setColumnWidth(1, 80);   // Ticker
  sheet.setColumnWidth(2, 105);  // Action
  sheet.setColumnWidth(3, 220);  // Analyst Firm
  sheet.setColumnWidth(4, 140);  // From Grade
  sheet.setColumnWidth(5, 140);  // To Grade
  sheet.setColumnWidth(6, 110);  // Date

  let r = 1;

  // ── Title banner ──
  sheet.getRange(r, 1, 1, 6).merge()
    .setValue("TOP ANALYST RATINGS  —  most recent per ticker  (as of " + dateStr + ")")
    .setBackground("#434343").setFontColor("#ffffff")
    .setFontWeight("bold").setFontSize(13).setFontFamily("Nunito")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.setRowHeight(r, 42);
  r++;

  // ── Timestamp ──
  sheet.getRange(r, 1, 1, 6).merge()
    .setValue("Refreshed " + Utilities.formatDate(new Date(), "America/New_York", "MMM dd, yyyy  h:mm a") + " ET")
    .setFontStyle("italic").setFontSize(9).setFontFamily("Nunito")
    .setHorizontalAlignment("center").setBackground("#f8f8f8");
  sheet.setRowHeight(r, 20);
  r++;

  r++; // blank spacer

  // ── Column headers ──
  sheet.getRange(r, 1, 1, 6)
    .setValues([["Ticker", "Action", "Analyst Firm", "From Grade", "To Grade", "Date"]])
    .setBackground("#434343").setFontColor("#ffffff")
    .setFontWeight("bold").setFontSize(11).setFontFamily("Nunito")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.setRowHeight(r, 28);
  r++;

  if (rows.length === 0) {
    sheet.getRange(r, 1, 1, 6).merge()
      .setValue("No upgrade / downgrade data found for the last 7 days.")
      .setHorizontalAlignment("center").setFontStyle("italic")
      .setFontFamily("Nunito").setFontSize(10);
    return;
  }

  rows.forEach(([ticker, action, firm, fromGrade, toGrade, date]) => {
    const isUp   = action === "Upgrade"  || action === "Initiates";
    const isDown = action === "Downgrade";
    const rowBg  = r % 2 === 0 ? "#f9f9f9" : "#ffffff";
    const actBg  = isUp   ? "#b6d7a8" : isDown ? "#ea9999" : "#cfe2f3";
    const actFg  = isUp   ? "#274e13" : isDown ? "#660000" : "#1c4587";

    // Apply row background first
    sheet.getRange(r, 1, 1, 6)
      .setBackground(rowBg).setFontFamily("Nunito").setFontSize(10)
      .setVerticalAlignment("middle");

    sheet.getRange(r, 1)
      .setValue(ticker).setFontWeight("bold").setHorizontalAlignment("center");

    // Action cell — colour-coded
    sheet.getRange(r, 2)
      .setValue(action)
      .setBackground(actBg).setFontColor(actFg)
      .setFontWeight("bold").setHorizontalAlignment("center");

    sheet.getRange(r, 3).setValue(firm).setHorizontalAlignment("left");
    sheet.getRange(r, 4).setValue(fromGrade).setHorizontalAlignment("center");
    sheet.getRange(r, 5).setValue(toGrade)
      .setFontWeight(isUp || isDown ? "bold" : "normal")
      .setFontColor(isUp ? "#274e13" : isDown ? "#660000" : "#000000")
      .setHorizontalAlignment("center");
    sheet.getRange(r, 6).setValue(date).setHorizontalAlignment("center");

    sheet.setRowHeight(r, 24);
    r++;
  });
}

// ============================================================
// TRIGGER SETUP
// ============================================================

function createTopRatingsTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "refreshTopRatings")
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("refreshTopRatings")
    .timeBased().atHour(9).everyDays(1)
    .inTimezone("America/New_York").create();

  Logger.log("✅ Daily 9 AM ET trigger created for refreshTopRatings.");
  SpreadsheetApp.getUi().alert("✅ Trigger set: Top Ratings will refresh every day at 9 AM ET.");
}
