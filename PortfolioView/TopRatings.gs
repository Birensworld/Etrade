// ============================================================
// TOP ANALYST RATINGS                            Version: 1.3
// Refreshes daily at 9 AM ET via a time-based trigger.
//
// Shows the most recent upgrade / downgrade / initiation for
// each stock in the watchlist — up to 50 entries, sorted newest first.
//
// Data source: Finnhub stock/upgrade-downgrade per ticker.
// This is the same API call used in the News Feed tab and is
// confirmed to work on the free tier.
//
// Shared utilities (finnhubGet_, formatDate_, FINNHUB_KEY_PROP)
// are defined in NewsFeed.gs and available in the same GAS scope.
// ============================================================

const TOP_RATINGS_SHEET = "Top Ratings";

// ~70 most-followed US stocks. Finnhub is called once per ticker.
// At 1.1s sleep per call this completes in ~80 seconds.
const TOP_RATINGS_WATCHLIST = [
  // Mega-cap tech / AI
  "AAPL","MSFT","NVDA","GOOGL","AMZN","META","TSLA","AVGO","ORCL","AMD",
  // Cloud / SaaS
  "CRM","SNOW","PLTR","NOW","ADBE","INTU","WDAY","MDB","DDOG","ZS","PANW",
  // Semiconductors
  "INTC","QCOM","AMAT","LRCX","MRVL","ARM","SMCI","TXN","MU",
  // Finance
  "JPM","BAC","GS","MS","BLK","V","MA","PYPL","AXP","WFC","C",
  // Healthcare / Biotech
  "JNJ","UNH","LLY","PFE","MRK","ABBV","AMGN","GILD","REGN",
  // Consumer
  "WMT","COST","TGT","NKE","MCD","SBUX","HD","TJX","SHOP",
  // Energy
  "XOM","CVX","COP","SLB","OXY",
  // Industrial / Defense
  "GE","CAT","BA","RTX","LMT","HON",
  // Telecom / Media / Streaming
  "T","VZ","NFLX","DIS","CMCSA","SPOT",
  // Other high-profile
  "UBER","ABNB","COIN","RBLX","HOOD","SQ"
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

  // Also include the user's own portfolio tickers so their stocks show up too
  const portfolioTickers = getPortfolioTickers_(ss);
  const allTickers = [...new Set([...TOP_RATINGS_WATCHLIST, ...portfolioTickers])];

  Logger.log("Fetching ratings for " + allTickers.length + " tickers …");
  const rows = fetchRatingsPerTicker_(allTickers, from90d, apiKey);

  rows.sort(function(a, b) { return (b[5] || "").localeCompare(a[5] || ""); });
  const top50 = rows.slice(0, 50);

  buildTopRatingsSheet_(ss, top50, today, allTickers.length, rows.length);
  Logger.log("Top Ratings done — " + top50.length + " entries displayed (" + rows.length + " found).");
}

// ============================================================
// DATA FETCHING
// ============================================================

// Finnhub per-ticker: same endpoint used in NewsFeed.gs — proven to work.
// Sleeps 1.1s between calls to stay under the 60 req/min free-tier limit.
// Returns [ticker, action, firm, fromGrade, toGrade, date] rows.
function fetchRatingsPerTicker_(tickers, from, apiKey) {
  const ACTION_MAP = { up: "Upgrade", down: "Downgrade", init: "Initiates" };
  const rows = [];

  tickers.forEach(function(ticker) {
    try {
      const resp = finnhubGet_(
        "stock/upgrade-downgrade?symbol=" + encodeURIComponent(ticker) + "&from=" + from,
        apiKey
      );
      if (!Array.isArray(resp) || resp.length === 0) {
        Utilities.sleep(1100);
        return;
      }

      // Most recent directional action (skip maintains / reiterates)
      var entry = null;
      for (var i = 0; i < resp.length; i++) {
        if (ACTION_MAP[(resp[i].action || "").toLowerCase()]) { entry = resp[i]; break; }
      }
      if (!entry) {
        Utilities.sleep(1100);
        return;
      }

      rows.push([
        ticker,
        ACTION_MAP[(entry.action || "").toLowerCase()],
        (entry.company   || "").trim(),
        (entry.fromGrade || "—").trim() || "—",
        (entry.toGrade   || "—").trim() || "—",
        (entry.gradeDate || "").trim()
      ]);
    } catch (e) {
      Logger.log("TopRatings " + ticker + ": " + e.message);
    }
    Utilities.sleep(1100);
  });

  return rows;
}

// ============================================================
// SHEET BUILDER
// ============================================================

function buildTopRatingsSheet_(ss, rows, dateStr, tickerCount, foundCount) {
  var sheet = ss.getSheetByName(TOP_RATINGS_SHEET);
  if (!sheet) sheet = ss.insertSheet(TOP_RATINGS_SHEET);
  else sheet.clear();

  sheet.setColumnWidth(1, 80);   // Ticker
  sheet.setColumnWidth(2, 105);  // Action
  sheet.setColumnWidth(3, 230);  // Analyst Firm
  sheet.setColumnWidth(4, 140);  // From Grade
  sheet.setColumnWidth(5, 140);  // To Grade
  sheet.setColumnWidth(6, 110);  // Date

  var r = 1;

  // ── Title ──
  sheet.getRange(r, 1, 1, 6).merge()
    .setValue("TOP ANALYST RATINGS  —  most recent per ticker  (as of " + dateStr + ")")
    .setBackground("#434343").setFontColor("#ffffff")
    .setFontWeight("bold").setFontSize(13).setFontFamily("Nunito")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.setRowHeight(r, 42);
  r++;

  // ── Timestamp + coverage stats ──
  sheet.getRange(r, 1, 1, 6).merge()
    .setValue("Refreshed " +
      Utilities.formatDate(new Date(), "America/New_York", "MMM dd, yyyy  h:mm a") +
      " ET  |  Scanned " + tickerCount + " tickers — " + foundCount + " had ratings data in last 90 days")
    .setFontStyle("italic").setFontSize(9).setFontFamily("Nunito")
    .setHorizontalAlignment("center").setBackground("#f8f8f8");
  sheet.setRowHeight(r, 22);
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
      .setValue("No rating data returned by Finnhub for the last 90 days. Verify your API key is set correctly.")
      .setHorizontalAlignment("center").setFontStyle("italic").setWrap(true)
      .setFontFamily("Nunito").setFontSize(10).setFontColor("#cc0000");
    return;
  }

  rows.forEach(function(row) {
    var ticker    = row[0];
    var action    = row[1];
    var firm      = row[2];
    var fromGrade = row[3];
    var toGrade   = row[4];
    var date      = row[5];

    var isUp   = action === "Upgrade" || action === "Initiates";
    var isDown = action === "Downgrade";
    var rowBg  = r % 2 === 0 ? "#f9f9f9" : "#ffffff";
    var actBg  = isUp ? "#b6d7a8" : isDown ? "#ea9999" : "#cfe2f3";
    var actFg  = isUp ? "#274e13" : isDown ? "#660000" : "#1c4587";

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
