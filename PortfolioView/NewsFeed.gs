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

// A headline must contain at least one of these to be shown.
// These are events that directly move a stock price.
const MARKET_MOVING_KEYWORDS = [
  // Earnings & guidance
  "earnings beat", "earnings miss", "beat estimates", "missed estimates",
  "beats expectations", "misses expectations", "raised guidance", "lowered guidance",
  "cuts forecast", "raises forecast", "profit warning", "revenue warning",
  "quarterly results", "full-year guidance",
  // Analyst actions with price impact
  "price target", "initiates coverage", "upgrades", "downgrades",
  // M&A / corporate events
  "merger", "acquisition", "acquires", "takeover", "buyout",
  "going private", "spinoff", "spin-off", "divests", "sells unit",
  // Capital returns
  "buyback", "share repurchase", "dividend cut", "dividend suspended",
  "special dividend", "raises dividend",
  // Regulatory / legal
  "fda approval", "fda rejects", "fda clears", "sec investigation",
  "doj investigation", "antitrust", "class action", "settlement",
  "subpoena", "indicted",
  // Operational shocks
  "ceo resigns", "ceo fired", "ceo steps down", "cfo resigns",
  "bankruptcy", "chapter 11", "defaults", "debt restructuring",
  "product recall", "data breach", "cyberattack", "plant closure",
  "major layoff", "mass layoff",
  // Contracts & partnerships
  "major contract", "awarded contract", "loses contract",
  "strategic partnership", "joint venture"
];

// Headlines containing these phrases are generic noise — skip even if
// they match a keyword above.
const NOISE_PHRASES = [
  "to present at", "to speak at", "conference call", "webcast",
  "names new vp", "names new director", "promotes", "appoints vp",
  "monthly traffic", "weekly data", "analyst day", "investor day",
  "price target raised by", // lone PT changes without action are minor
  "scheduled to report", "expected to report", "will report earnings on"
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
      rows.push([ticker, "N/A", "Error: " + e.message, "—", "—"]);
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

  // ── News (only truly market-moving headlines; blank otherwise) ──
  let news = "";
  try {
    const resp = finnhubGet_(`company-news?symbol=${ticker}&from=${from3d}&to=${todayS}`, apiKey);
    if (Array.isArray(resp) && resp.length > 0) {
      const sig = resp.filter(n => {
        const h = (n.headline || "").toLowerCase();
        if (NOISE_PHRASES.some(p => h.includes(p))) return false;
        return MARKET_MOVING_KEYWORDS.some(k => h.includes(k));
      });
      if (sig.length > 0) news = sig.slice(0, 2).map(n => n.headline).join("  |  ");
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

  // Columns: Ticker | Direction | News | Earnings Date | Upgrade/Downgrade
  return [ticker, direction, news, earningsDate, upgradeDowngrade];
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
  sheet.setColumnWidth(3, 620);  // News
  sheet.setColumnWidth(4, 120);  // Earnings Date
  sheet.setColumnWidth(5, 260);  // Upgrade/Downgrade

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
    .setValues([["Ticker", "Direction", "News", "Earnings Date", "Upgrade / Downgrade"]])
    .setBackground("#cfe2f3")
    .setFontWeight("bold")
    .setFontSize(11)
    .setFontFamily("Nunito")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(r, 30);
  r++;

  // ── Data rows ──
  // Array order: [ticker, direction, news, earningsDate, upgradeDowngrade]
  rows.forEach(([ticker, direction, news, earningsDate, upgradeDowngrade]) => {
    sheet.getRange(r, 1).setValue(ticker).setHorizontalAlignment("center");

    const dirCell = sheet.getRange(r, 2);
    dirCell.setValue(direction).setHorizontalAlignment("center");
    if      (direction === "Bullish") dirCell.setFontColor("#38761d");
    else if (direction === "Bearish") dirCell.setFontColor("#cc0000");
    else                              dirCell.setFontColor("#7d6608");

    sheet.getRange(r, 3).setValue(news).setWrap(true).setHorizontalAlignment("left");
    sheet.getRange(r, 4).setValue(earningsDate).setHorizontalAlignment("center");
    sheet.getRange(r, 5).setValue(upgradeDowngrade).setHorizontalAlignment("left");

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
