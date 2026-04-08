// ============================================================
// NEWS FEED
// Refreshes daily at 9 AM ET via a time-based trigger.
// Data source: Finnhub (free tier — finnhub.io)
//
// First-time setup:
//   1. Run setFinnhubApiKey()  → paste your free API key
//   2. Run createNewsFeedTrigger() → schedules daily 9 AM ET refresh
// ============================================================

const NEWS_SHEET_NAME   = "News Feed";
const FINNHUB_KEY_PROP  = "FINNHUB_API_KEY";

// Ticker regex: 1–6 uppercase letters / dots (covers BRK.B, etc.)
const TICKER_RE = /^[A-Z][A-Z.]{0,5}$/;

// Keywords that flag a headline as market-moving
const SIGNIFICANT_KEYWORDS = [
  "beat", "miss", "upgrade", "downgrade", "earnings", "guidance",
  "raises", "cuts", "buyback", "merger", "acqui", "lawsuit",
  "fda", "deal", "record", "outlook", "charge", "restatement",
  "investigation", "layoff", "dividend"
];

// ============================================================
// ENTRY POINT
// ============================================================

function refreshNewsFeed() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const apiKey = PropertiesService.getScriptProperties().getProperty(FINNHUB_KEY_PROP);

  if (!apiKey) {
    SpreadsheetApp.getUi().alert(
      "Finnhub API key not set.\n\nRun: setFinnhubApiKey()  (Extensions → Apps Script → Run)"
    );
    return;
  }

  const tickers = getPortfolioTickers_(ss);
  if (tickers.length === 0) {
    Logger.log("No tickers found in Portfolio View sheet.");
    return;
  }

  Logger.log("Fetching news for " + tickers.length + " tickers: " + tickers.join(", "));

  const marketOverview = fetchMarketOverview_(apiKey);

  const rows = [];
  tickers.forEach(ticker => {
    try {
      rows.push(fetchTickerRow_(ticker, apiKey));
    } catch (e) {
      Logger.log("⚠ " + ticker + ": " + e.message);
      rows.push([ticker, "N/A", "—", "—", "Error: " + e.message]);
    }
    Utilities.sleep(250); // stay within 60 req/min free-tier limit
  });

  buildNewsFeedSheet_(ss, marketOverview, rows);
}

// ============================================================
// DATA FETCHING
// ============================================================

function fetchMarketOverview_(apiKey) {
  try {
    const url  = `https://finnhub.io/api/v1/news?category=general&token=${apiKey}`;
    const resp = JSON.parse(UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText());
    if (!Array.isArray(resp) || resp.length === 0) return { direction: "Neutral", headline: "No market data available." };

    // Simple sentiment score on top-10 headlines
    const text = resp.slice(0, 10).map(n => n.headline || "").join(" ").toLowerCase();
    const bullScore = ["rally", "gain", "surge", "rise", "optimism", "record", "beat", "strong", "up"].filter(k => text.includes(k)).length;
    const bearScore = ["fall", "drop", "decline", "recession", "fear", "loss", "weak", "miss", "inflation", "sell-off", "down"].filter(k => text.includes(k)).length;

    const direction = bullScore > bearScore ? "Bullish" : bearScore > bullScore ? "Bearish" : "Neutral";
    return { direction, headline: resp[0].headline || "" };
  } catch (e) {
    return { direction: "N/A", headline: "Could not fetch market overview." };
  }
}

function fetchTickerRow_(ticker, apiKey) {
  const today  = new Date();
  const from3d = formatDate_(new Date(today - 3 * 864e5));   // 3 days back
  const todayS = formatDate_(today);
  const to90d  = formatDate_(new Date(today - -90 * 864e5)); // 90 days ahead

  // ── News ──
  let news = "No significant news";
  try {
    const resp = finnhubGet_(`company-news?symbol=${ticker}&from=${from3d}&to=${todayS}`, apiKey);
    if (Array.isArray(resp) && resp.length > 0) {
      const sig = resp.filter(n => {
        const h = (n.headline || "").toLowerCase();
        return SIGNIFICANT_KEYWORDS.some(k => h.includes(k));
      });
      const src = sig.length > 0 ? sig : resp;
      news = src.slice(0, 2).map(n => n.headline).join("  |  ");
    }
  } catch (e) { Logger.log("News: " + ticker + " – " + e.message); }
  Utilities.sleep(150);

  // ── Analyst consensus → Direction ──
  let direction = "Neutral";
  try {
    const resp = finnhubGet_(`stock/recommendation?symbol=${ticker}`, apiKey);
    if (Array.isArray(resp) && resp.length > 0) {
      const r        = resp[0];
      const bullish  = (r.strongBuy || 0) + (r.buy  || 0);
      const bearish  = (r.strongSell || 0) + (r.sell || 0);
      const neutral  = r.hold || 0;
      if      (bullish > bearish && bullish >= neutral) direction = "Bullish";
      else if (bearish > bullish && bearish >= neutral) direction = "Bearish";
    }
  } catch (e) { Logger.log("Rec: " + ticker + " – " + e.message); }
  Utilities.sleep(150);

  // ── Recent upgrade / downgrade ──
  let upgradeDowngrade = "—";
  try {
    const resp = finnhubGet_(`stock/upgrade-downgrade?symbol=${ticker}&from=${from3d}`, apiKey);
    if (Array.isArray(resp) && resp.length > 0) {
      const r      = resp[0];
      const action = (r.action    || "").replace(/^(up|down)grade$/i, s => s[0].toUpperCase() + s.slice(1).toLowerCase());
      const firm   = r.company  || "";
      const grade  = r.toGrade  || "";
      if (action && grade) upgradeDowngrade = `${firm}: ${action} → ${grade}`;
    }
  } catch (e) { Logger.log("UG: " + ticker + " – " + e.message); }
  Utilities.sleep(150);

  // ── Next earnings date ──
  let earningsDate = "—";
  try {
    const resp = finnhubGet_(`calendar/earnings?from=${todayS}&to=${to90d}&symbol=${ticker}`, apiKey);
    const cal  = resp && resp.earningsCalendar;
    if (Array.isArray(cal) && cal.length > 0) earningsDate = cal[0].date || "—";
  } catch (e) { Logger.log("Earnings: " + ticker + " – " + e.message); }

  return [ticker, direction, upgradeDowngrade, earningsDate, news];
}

// ============================================================
// SHEET BUILDER
// ============================================================

function buildNewsFeedSheet_(ss, marketOverview, rows) {
  let sheet = ss.getSheetByName(NEWS_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(NEWS_SHEET_NAME);
  else sheet.clear();

  // Column widths
  sheet.setColumnWidth(1, 80);   // Ticker
  sheet.setColumnWidth(2, 100);  // Direction
  sheet.setColumnWidth(3, 240);  // Upgrade/Downgrade
  sheet.setColumnWidth(4, 120);  // Earnings Date
  sheet.setColumnWidth(5, 620);  // News

  const dirColor = { Bullish: "#b6d7a8", Bearish: "#ea9999", Neutral: "#ffe599", "N/A": "#f3f3f3" };
  let r = 1;

  // ── Market direction banner ──
  const bannerBg = dirColor[marketOverview.direction] || "#f3f3f3";
  sheet.getRange(r, 1, 1, 5).merge()
    .setValue(`MARKET DIRECTION: ${marketOverview.direction.toUpperCase()}   |   ${marketOverview.headline}`)
    .setBackground(bannerBg)
    .setFontWeight("bold")
    .setFontSize(12)
    .setFontFamily("Nunito")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);
  sheet.setRowHeight(r, 48);
  r++;

  // ── Timestamp ──
  sheet.getRange(r, 1, 1, 5).merge()
    .setValue("Refreshed " + Utilities.formatDate(new Date(), "America/New_York", "MMM dd, yyyy  h:mm a") + " ET")
    .setFontStyle("italic")
    .setFontSize(9)
    .setFontFamily("Nunito")
    .setHorizontalAlignment("center")
    .setBackground("#f8f8f8");
  sheet.setRowHeight(r, 22);
  r++;

  r++; // blank spacer row

  // ── Column headers ──
  sheet.getRange(r, 1, 1, 5)
    .setValues([["Ticker", "Direction", "Upgrade / Downgrade", "Earnings Date", "News"]])
    .setBackground("#cfe2f3")
    .setFontWeight("bold")
    .setFontSize(11)
    .setFontFamily("Nunito")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(r, 30);
  r++;

  // ── Data rows ──
  rows.forEach(([ticker, direction, upgradeDowngrade, earningsDate, news]) => {
    sheet.getRange(r, 1).setValue(ticker).setHorizontalAlignment("center");

    const dirCell = sheet.getRange(r, 2);
    dirCell.setValue(direction).setHorizontalAlignment("center");
    if      (direction === "Bullish") dirCell.setFontColor("#38761d");
    else if (direction === "Bearish") dirCell.setFontColor("#cc0000");
    else                              dirCell.setFontColor("#7d6608");

    sheet.getRange(r, 3).setValue(upgradeDowngrade).setHorizontalAlignment("left");
    sheet.getRange(r, 4).setValue(earningsDate).setHorizontalAlignment("center");
    sheet.getRange(r, 5).setValue(news).setWrap(true).setHorizontalAlignment("left");

    sheet.getRange(r, 1, 1, 5)
      .setBackground(r % 2 === 0 ? "#f9f9f9" : "#ffffff")
      .setFontFamily("Nunito")
      .setFontSize(10);

    sheet.setRowHeight(r, 52);
    r++;
  });

  // Bold ticker column
  sheet.getRange(5, 1, rows.length, 1).setFontWeight("bold");
}

// ============================================================
// SETUP HELPERS
// ============================================================

// Schedules refreshNewsFeed() to run every weekday at 9 AM ET
function createNewsFeedTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "refreshNewsFeed")
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("refreshNewsFeed")
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .inTimezone("America/New_York")
    .create();

  Logger.log("✅ Daily 9 AM ET trigger created for refreshNewsFeed.");
  SpreadsheetApp.getUi().alert("✅ Trigger set: News Feed will refresh every day at 9 AM ET.");
}

// Prompts for and saves the Finnhub API key as a script property
function setFinnhubApiKey() {
  const ui     = SpreadsheetApp.getUi();
  const result = ui.prompt(
    "Finnhub API Key",
    "Paste your free API key from finnhub.io:",
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() !== ui.Button.OK) return;
  const key = result.getResponseText().trim();
  if (!key) { ui.alert("No key entered — nothing saved."); return; }
  PropertiesService.getScriptProperties().setProperty(FINNHUB_KEY_PROP, key);
  ui.alert("✅ API key saved.\n\nNext: run createNewsFeedTrigger() to schedule the daily refresh.");
}

// ============================================================
// PRIVATE UTILITIES
// ============================================================

function finnhubGet_(endpoint, apiKey) {
  const url  = `https://finnhub.io/api/v1/${endpoint}&token=${apiKey}`;
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) throw new Error("HTTP " + resp.getResponseCode());
  return JSON.parse(resp.getContentText());
}

function getPortfolioTickers_(ss) {
  const sheet = ss.getSheetByName("Portfolio View");
  if (!sheet) return [];

  const vals    = sheet.getRange(1, 1, sheet.getLastRow()).getValues();
  const skip    = new Set(["symbol", "totals", ...allowedAccounts.map(a => a.toLowerCase())]);
  const tickers = new Set();

  vals.forEach(([sym]) => {
    const s = (sym || "").toString().trim();
    if (!s || skip.has(s.toLowerCase())) return;
    if (TICKER_RE.test(s)) tickers.add(s);
  });

  return [...tickers].sort();
}

function formatDate_(d) {
  return Utilities.formatDate(d, "America/New_York", "yyyy-MM-dd");
}
