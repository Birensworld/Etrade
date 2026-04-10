// ============================================================
// TOP ANALYST RATINGS                            Version: 1.2
// Refreshes daily at 9 AM ET via a time-based trigger.
//
// Shows up to 50 recent upgrade / downgrade / initiation actions
// across the broad market — not limited to your portfolio.
//
// Data sources (tried in order):
//   1. Finnhub  — market-wide stock/upgrade-downgrade (no symbol)
//   2. Yahoo Finance v8 — per-ticker quoteSummary (curated watchlist)
//   3. Finnhub per-ticker — definitive fallback, same call as NewsFeed
//
// Shared utilities (finnhubGet_, formatDate_, FINNHUB_KEY_PROP)
// are defined in NewsFeed.gs and available in the same GAS scope.
// ============================================================

const TOP_RATINGS_SHEET = "Top Ratings";

// Curated watchlist — most-followed US stocks across sectors.
const TOP_RATINGS_WATCHLIST = [
  // Mega-cap tech / AI
  "AAPL","MSFT","NVDA","GOOGL","AMZN","META","TSLA","AVGO","ORCL","AMD",
  // Cloud / SaaS
  "CRM","SNOW","PLTR","NOW","ADBE","INTU","WDAY","MDB","DDOG","ZS","PANW",
  // Semiconductors
  "INTC","QCOM","AMAT","LRCX","KLAC","MRVL","ARM","SMCI","TXN","MU",
  // Finance
  "JPM","BAC","GS","MS","BLK","V","MA","PYPL","AXP","COF","WFC","C",
  // Healthcare / Biotech
  "JNJ","UNH","LLY","PFE","MRK","ABBV","AMGN","GILD","REGN",
  // Consumer
  "WMT","COST","TGT","NKE","MCD","SBUX","HD","TJX","BABA","SHOP",
  // Energy
  "XOM","CVX","COP","SLB","OXY",
  // Industrial / Defense
  "GE","CAT","BA","RTX","LMT","HON",
  // Telecom / Media / Streaming
  "T","VZ","NFLX","DIS","CMCSA","SPOT",
  // Other
  "UBER","ABNB","COIN","RBLX","HOOD"
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
  const from90d = formatDate_(new Date(new Date() - 90 * 864e5));

  const sourceLog = [];  // written to sheet so you can see what worked

  // ── Layer 1: Finnhub market-wide (no symbol) ──
  let rows = [];
  try {
    rows = fetchTopRatingsFromFinnhub_(from90d, today, apiKey);
    sourceLog.push("Finnhub market-wide: " + rows.length + " rows");
  } catch (e) {
    sourceLog.push("Finnhub market-wide: ERROR – " + e.message);
  }
  Logger.log(sourceLog[sourceLog.length - 1]);

  // ── Layer 2: Yahoo Finance per-ticker (v8 endpoint, no optional chaining) ──
  const seen = new Set(rows.map(function(r) { return r[0]; }));
  const yahooRows = fetchTopRatingsFromYahoo_(TOP_RATINGS_WATCHLIST);
  yahooRows.forEach(function(r) { if (!seen.has(r[0])) { rows.push(r); seen.add(r[0]); } });
  sourceLog.push("Yahoo Finance: " + yahooRows.length + " rows");
  Logger.log(sourceLog[sourceLog.length - 1]);

  // ── Layer 3: Finnhub per-ticker (definitive fallback — same call as NewsFeed) ──
  // Only runs if layers 1+2 produced fewer than 10 rows.
  if (rows.length < 10) {
    const fRows = fetchTopRatingsPerTickerFinnhub_(TOP_RATINGS_WATCHLIST, from90d, apiKey);
    fRows.forEach(function(r) { if (!seen.has(r[0])) { rows.push(r); seen.add(r[0]); } });
    sourceLog.push("Finnhub per-ticker: " + fRows.length + " rows");
    Logger.log(sourceLog[sourceLog.length - 1]);
  }

  rows.sort(function(a, b) { return (b[5] || "").localeCompare(a[5] || ""); });
  rows = rows.slice(0, 50);

  buildTopRatingsSheet_(ss, rows, today, sourceLog);
  Logger.log("Top Ratings done — " + rows.length + " entries written.");
}

// ============================================================
// DATA FETCHING
// ============================================================

// Layer 1: Finnhub market-wide (free tier often returns [] — layers 2/3 cover gaps).
function fetchTopRatingsFromFinnhub_(from, to, apiKey) {
  const ACTION_MAP = { up: "Upgrade", down: "Downgrade", init: "Initiates" };
  const resp = finnhubGet_("stock/upgrade-downgrade?from=" + from + "&to=" + to, apiKey);
  if (!Array.isArray(resp) || resp.length === 0) return [];
  return resp
    .filter(function(r) { return ACTION_MAP[(r.action || "").toLowerCase()]; })
    .sort(function(a, b) { return (b.gradeDate || "").localeCompare(a.gradeDate || ""); })
    .slice(0, 50)
    .map(function(r) {
      return [
        (r.symbol    || "").toUpperCase(),
        ACTION_MAP[(r.action || "").toLowerCase()],
        (r.company   || "").trim(),
        (r.fromGrade || "—").trim() || "—",
        (r.toGrade   || "—").trim() || "—",
        (r.gradeDate || "").trim()
      ];
    });
}

// Layer 2: Yahoo Finance per-ticker via v8 quoteSummary.
// Uses explicit null checks (no optional chaining) for broadest GAS compatibility.
function fetchTopRatingsFromYahoo_(tickers) {
  const ACTION_MAP = { up: "Upgrade", down: "Downgrade", init: "Initiates" };
  const rows = [];

  tickers.forEach(function(ticker) {
    try {
      // v8 endpoint — more stable than v10 across Yahoo API changes
      const url = "https://query1.finance.yahoo.com/v8/finance/quoteSummary/" +
                  encodeURIComponent(ticker) +
                  "?modules=upgradeDowngradeHistory";
      const resp = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        headers: { "User-Agent": "Mozilla/5.0 (compatible)" }
      });
      if (resp.getResponseCode() !== 200) {
        Logger.log("Yahoo " + ticker + " HTTP " + resp.getResponseCode());
        return;
      }

      // Explicit null checks — no optional chaining
      const data = JSON.parse(resp.getContentText());
      if (!data || !data.quoteSummary) return;
      const result = data.quoteSummary.result;
      if (!Array.isArray(result) || result.length === 0) return;
      const udh = result[0].upgradeDowngradeHistory;
      if (!udh) return;
      const history = udh.history;
      if (!Array.isArray(history) || history.length === 0) return;

      // Most-recent directional action — no date cutoff
      var entry = null;
      for (var i = 0; i < history.length; i++) {
        var h = history[i];
        if (ACTION_MAP[(h.action || "").toLowerCase()] && h.epochGradeDate) {
          entry = h;
          break;
        }
      }
      if (!entry) return;

      const date = Utilities.formatDate(
        new Date(entry.epochGradeDate * 1000), "America/New_York", "yyyy-MM-dd"
      );
      rows.push([
        ticker,
        ACTION_MAP[(entry.action || "").toLowerCase()],
        (entry.firm       || "").trim(),
        (entry.fromGrade  || "—").trim() || "—",
        (entry.toGrade    || "—").trim() || "—",
        date
      ]);
    } catch (e) {
      Logger.log("Yahoo " + ticker + ": " + e.message);
    }
    Utilities.sleep(200);
  });

  return rows;
}

// Layer 3: Finnhub per-ticker — same endpoint used in NewsFeed, guaranteed to work.
// Sleeps 1.1s between calls to stay under Finnhub free-tier rate limit (60 req/min).
function fetchTopRatingsPerTickerFinnhub_(tickers, from, apiKey) {
  const ACTION_MAP = { up: "Upgrade", down: "Downgrade", init: "Initiates" };
  const rows = [];

  tickers.forEach(function(ticker) {
    try {
      const resp = finnhubGet_(
        "stock/upgrade-downgrade?symbol=" + ticker + "&from=" + from, apiKey
      );
      if (!Array.isArray(resp) || resp.length === 0) return;
      var entry = null;
      for (var i = 0; i < resp.length; i++) {
        if (ACTION_MAP[(resp[i].action || "").toLowerCase()]) { entry = resp[i]; break; }
      }
      if (!entry) return;
      rows.push([
        ticker,
        ACTION_MAP[(entry.action || "").toLowerCase()],
        (entry.company   || "").trim(),
        (entry.fromGrade || "—").trim() || "—",
        (entry.toGrade   || "—").trim() || "—",
        (entry.gradeDate || "").trim()
      ]);
    } catch (e) {
      Logger.log("Finnhub per-ticker " + ticker + ": " + e.message);
    }
    Utilities.sleep(1100); // stay under 60 req/min free-tier limit
  });

  return rows;
}

// ============================================================
// SHEET BUILDER
// ============================================================

function buildTopRatingsSheet_(ss, rows, dateStr, sourceLog) {
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

  // ── Title ──
  sheet.getRange(r, 1, 1, 6).merge()
    .setValue("TOP ANALYST RATINGS  —  most recent per ticker  (as of " + dateStr + ")")
    .setBackground("#434343").setFontColor("#ffffff")
    .setFontWeight("bold").setFontSize(13).setFontFamily("Nunito")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.setRowHeight(r, 42);
  r++;

  // ── Timestamp + source log ──
  sheet.getRange(r, 1, 1, 6).merge()
    .setValue("Refreshed " +
      Utilities.formatDate(new Date(), "America/New_York", "MMM dd, yyyy  h:mm a") +
      " ET  |  Sources: " + sourceLog.join("  •  "))
    .setFontStyle("italic").setFontSize(9).setFontFamily("Nunito")
    .setHorizontalAlignment("center").setBackground("#f8f8f8").setWrap(true);
  sheet.setRowHeight(r, 30);
  r++;

  r++; // spacer

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
      .setValue("No upgrade / downgrade data found. Check the source log in the header for details.")
      .setHorizontalAlignment("center").setFontStyle("italic")
      .setFontFamily("Nunito").setFontSize(10).setFontColor("#cc0000");
    return;
  }

  rows.forEach(function([ticker, action, firm, fromGrade, toGrade, date]) {
    const isUp   = action === "Upgrade" || action === "Initiates";
    const isDown = action === "Downgrade";
    const rowBg  = r % 2 === 0 ? "#f9f9f9" : "#ffffff";
    const actBg  = isUp ? "#b6d7a8" : isDown ? "#ea9999" : "#cfe2f3";
    const actFg  = isUp ? "#274e13" : isDown ? "#660000" : "#1c4587";

    sheet.getRange(r, 1, 1, 6)
      .setBackground(rowBg).setFontFamily("Nunito").setFontSize(10)
      .setVerticalAlignment("middle");

    sheet.getRange(r, 1)
      .setValue(ticker).setFontWeight("bold").setHorizontalAlignment("center");

    sheet.getRange(r, 2)
      .setValue(action).setBackground(actBg).setFontColor(actFg)
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
    .filter(function(t) { return t.getHandlerFunction() === "refreshTopRatings"; })
    .forEach(function(t) { ScriptApp.deleteTrigger(t); });

  ScriptApp.newTrigger("refreshTopRatings")
    .timeBased().atHour(9).everyDays(1)
    .inTimezone("America/New_York").create();

  Logger.log("✅ Daily 9 AM ET trigger created for refreshTopRatings.");
  SpreadsheetApp.getUi().alert("✅ Trigger set: Top Ratings will refresh every day at 9 AM ET.");
}
