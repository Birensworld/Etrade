// ============================================================
// TOP ANALYST RATINGS                            Version: 1.4
// Refreshes daily at 9 AM ET via a time-based trigger.
//
// Shows up to 50 recent upgrade / downgrade / initiation actions
// across the broad market — not limited to your portfolio.
//
// Data source: Finnhub company-news (free tier) + headline parsing.
// This is the same underlying data used in the News Feed tab for
// the U/D column, so it is confirmed to work on the free tier.
//
// Shared utilities (finnhubGet_, formatDate_, parseUDFromHeadline_,
// FINNHUB_KEY_PROP) are defined in NewsFeed.gs — same GAS scope.
// ============================================================

const TOP_RATINGS_SHEET = "Top Ratings";

// ~70 most-followed US stocks. One company-news call per ticker.
const TOP_RATINGS_WATCHLIST = [
  // Mega-cap tech / AI
  "AAPL","MSFT","NVDA","GOOGL","AMZN","META","TSLA","AVGO","ORCL","AMD",
  // Cloud / SaaS
  "CRM","SNOW","PLTR","NOW","ADBE","INTU","WDAY","MDB","DDOG","ZS","PANW",
  // Semiconductors
  "INTC","QCOM","AMAT","LRCX","MRVL","ARM","SMCI","TXN","MU",
  // Finance
  "JPM","BAC","GS","MS","BLK","V","MA","PYPL","AXP","WFC","C","SCHW",
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
  const from7d  = formatDate_(new Date(new Date() - 7 * 864e5));

  // Merge watchlist with user's own portfolio tickers so their stocks are included
  const portfolioTickers = getPortfolioTickers_(ss);
  const allTickers = [...new Set([...TOP_RATINGS_WATCHLIST, ...portfolioTickers])];

  Logger.log("Scanning company-news for ratings in " + allTickers.length + " tickers …");

  const rows   = [];
  const seen   = new Set();   // deduplicate by ticker

  allTickers.forEach(function(ticker) {
    if (seen.has(ticker)) return;
    try {
      // Fetch 7-day company news — same free-tier call used by NewsFeed
      const resp = finnhubGet_(
        "company-news?symbol=" + encodeURIComponent(ticker) +
        "&from=" + from7d + "&to=" + today,
        apiKey
      );
      if (!Array.isArray(resp)) return;

      // Scan each headline for upgrade/downgrade/initiation patterns
      for (var i = 0; i < resp.length; i++) {
        const item   = resp[i];
        const parsed = parseUDFromHeadline_(item.headline || "", item.datetime, ticker);
        if (parsed) {
          // parsed = "date  Firm: Action → Grade"
          // Re-split into columns for the sheet
          const dateMatch  = parsed.match(/^(\d{4}-\d{2}-\d{2})\s+(.+?):\s+(.+?)\s+→\s+(.+)$/);
          if (dateMatch) {
            rows.push({
              ticker:    ticker,
              date:      dateMatch[1],
              firm:      dateMatch[2].trim(),
              action:    dateMatch[3].trim(),
              toGrade:   dateMatch[4].trim()
            });
          }
          seen.add(ticker);
          break;  // one rating per ticker
        }
      }
    } catch (e) {
      Logger.log("TopRatings " + ticker + ": " + e.message);
    }
    Utilities.sleep(300);
  });

  // Sort newest first
  rows.sort(function(a, b) { return b.date.localeCompare(a.date); });

  buildTopRatingsSheet_(ss, rows.slice(0, 50), today, allTickers.length);
  Logger.log("Top Ratings done — " + rows.length + " entries found.");
}

// ============================================================
// SHEET BUILDER
// ============================================================

function buildTopRatingsSheet_(ss, rows, dateStr, tickerCount) {
  var sheet = ss.getSheetByName(TOP_RATINGS_SHEET);
  if (!sheet) sheet = ss.insertSheet(TOP_RATINGS_SHEET);
  else sheet.clear();

  sheet.setColumnWidth(1, 80);   // Ticker
  sheet.setColumnWidth(2, 105);  // Action
  sheet.setColumnWidth(3, 230);  // Analyst Firm
  sheet.setColumnWidth(4, 150);  // To Grade
  sheet.setColumnWidth(5, 110);  // Date

  var r = 1;

  // ── Title ──
  sheet.getRange(r, 1, 1, 5).merge()
    .setValue("TOP ANALYST RATINGS  —  last 7 days  (as of " + dateStr + ")")
    .setBackground("#434343").setFontColor("#ffffff")
    .setFontWeight("bold").setFontSize(13).setFontFamily("Nunito")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.setRowHeight(r, 42);
  r++;

  // ── Timestamp ──
  sheet.getRange(r, 1, 1, 5).merge()
    .setValue("Refreshed " +
      Utilities.formatDate(new Date(), "America/New_York", "MMM dd, yyyy  h:mm a") +
      " ET  |  Scanned " + tickerCount + " tickers via Finnhub company news")
    .setFontStyle("italic").setFontSize(9).setFontFamily("Nunito")
    .setHorizontalAlignment("center").setBackground("#f8f8f8");
  sheet.setRowHeight(r, 22);
  r++;

  r++; // spacer

  // ── Column headers ──
  sheet.getRange(r, 1, 1, 5)
    .setValues([["Ticker", "Action", "Analyst Firm", "New Rating", "Date"]])
    .setBackground("#434343").setFontColor("#ffffff")
    .setFontWeight("bold").setFontSize(11).setFontFamily("Nunito")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.setRowHeight(r, 28);
  r++;

  if (rows.length === 0) {
    sheet.getRange(r, 1, 1, 5).merge()
      .setValue("No upgrades or downgrades found in company news for the last 7 days.")
      .setHorizontalAlignment("center").setFontStyle("italic").setWrap(true)
      .setFontFamily("Nunito").setFontSize(10).setFontColor("#cc0000");
    return;
  }

  rows.forEach(function(row) {
    var actionLower = (row.action || "").toLowerCase();
    var isUp   = actionLower.indexOf("upgrade") >= 0 || actionLower.indexOf("initiat") >= 0;
    var isDown = actionLower.indexOf("downgrade") >= 0;
    var rowBg  = r % 2 === 0 ? "#f9f9f9" : "#ffffff";
    var actBg  = isUp ? "#b6d7a8" : isDown ? "#ea9999" : "#cfe2f3";
    var actFg  = isUp ? "#274e13" : isDown ? "#660000" : "#1c4587";

    sheet.getRange(r, 1, 1, 5)
      .setBackground(rowBg).setFontFamily("Nunito").setFontSize(10)
      .setVerticalAlignment("middle");

    sheet.getRange(r, 1)
      .setValue(row.ticker).setFontWeight("bold").setHorizontalAlignment("center");

    sheet.getRange(r, 2)
      .setValue(row.action).setBackground(actBg).setFontColor(actFg)
      .setFontWeight("bold").setHorizontalAlignment("center");

    sheet.getRange(r, 3).setValue(row.firm).setHorizontalAlignment("left");

    sheet.getRange(r, 4).setValue(row.toGrade)
      .setFontWeight("bold")
      .setFontColor(isUp ? "#274e13" : isDown ? "#660000" : "#000000")
      .setHorizontalAlignment("center");

    sheet.getRange(r, 5).setValue(row.date).setHorizontalAlignment("center");

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
