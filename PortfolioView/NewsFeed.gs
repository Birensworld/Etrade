// ============================================================
// NEWS FEED                                      Version: 1.4
// Refreshes daily at 9 AM ET via a time-based trigger.
// Data source: Finnhub (free tier — finnhub.io)
//
// First-time setup:
//   1. Run setFinnhubApiKey()  → paste your free API key
//   2. Run createNewsFeedTrigger() → schedules daily 9 AM ET refresh
// ============================================================

const NEWS_SHEET_NAME  = "News Feed";
const FINNHUB_KEY_PROP = "FINNHUB_API_KEY";

// Ticker regex: 1–6 uppercase letters / dots (covers BRK.B, etc.)
const TICKER_RE = /^[A-Z][A-Z.]{0,5}$/;

// ETFs are excluded — analyst ratings and market-moving news don't
// apply the same way as individual stocks.
const ETF_TICKERS = new Set([
  // Broad market / S&P
  "SPY", "VOO", "VOOG", "UPRO", "RSP", "IVV", "VTI", "SPLG",
  // Tech / Nasdaq
  "QQQ", "TQQQ", "TECL", "SMH", "XLK", "QQEW", "QQQE", "SOXL", "IGV",
  // Small cap
  "IWM", "EQAL", "UWM", "TNA", "VBR",
  // Sector XL-series
  "XLE", "XLY", "XLF", "XLV", "XLI", "XLB", "XLU", "XLP", "XLRE",
  // Bitcoin / Crypto ETFs
  "IBIT", "FBTC", "GBTC", "BITO", "EZBC",
  // Fixed income / commodities / volatility
  "GLD", "SLV", "TLT", "IEF", "HYG", "LQD", "BND", "AGG", "VXX",
  // International
  "EEM", "EFA", "VEA", "ACWI"
]);

// A headline must match at least one of these to appear in News.
// Analyst upgrades/downgrades belong in the U/D column, not here.
const MARKET_MOVING_KEYWORDS = [
  // Earnings & guidance
  "earnings beat", "earnings miss", "beat estimates", "missed estimates",
  "beats expectations", "misses expectations", "raised guidance", "lowered guidance",
  "cuts forecast", "raises forecast", "profit warning", "revenue warning",
  "quarterly results", "full-year guidance",
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
  // Credit / bond
  "credit downgrade", "credit upgrade", "debt downgrade", "rating cut",
  // Contracts & partnerships
  "major contract", "awarded contract", "loses contract",
  "strategic partnership", "joint venture"
];

// Headlines matching any of these are suppressed — they are either
// captured by the U/D column or are generic non-events.
const NOISE_PHRASES = [
  // Analyst rating actions → U/D column handles these
  "upgrades ", "downgrades ", "upgraded to", "downgraded to",
  "reiterates", "maintains rating", "maintains buy", "maintains hold",
  "initiates with", "initiates coverage",
  // Generic non-events
  "to present at", "to speak at", "conference call", "webcast",
  "names new vp", "names new director", "promotes", "appoints vp",
  "monthly traffic", "weekly data", "analyst day", "investor day",
  "price target raised by", "price target lowered by",
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
    Utilities.sleep(300);
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

    const text      = resp.slice(0, 10).map(n => n.headline || "").join(" ").toLowerCase();
    const bullScore = ["rally", "gain", "surge", "rise", "optimism", "record", "beat", "strong"].filter(k => text.includes(k)).length;
    const bearScore = ["fall", "drop", "decline", "recession", "fear", "loss", "weak", "miss", "inflation", "sell-off"].filter(k => text.includes(k)).length;

    const direction = bullScore > bearScore ? "Bullish" : bearScore > bullScore ? "Bearish" : "Neutral";
    return { direction, headline: resp[0].headline || "" };
  } catch (e) {
    return { direction: "N/A", headline: "Could not fetch market overview." };
  }
}

function fetchTickerRow_(ticker, apiKey) {
  const today   = new Date();
  const from7d  = formatDate_(new Date(today - 7  * 864e5));  // 7 days back  (news + U/D fallback)
  const from30d = formatDate_(new Date(today - 30 * 864e5));  // 30 days back (U/D primary)
  const todayS  = formatDate_(today);
  const to90d   = formatDate_(new Date(today - -90 * 864e5)); // 90 days ahead

  // ── Fetch raw news (7-day window) — used for both News and U/D fallback ──
  let rawNews = [];
  try {
    const resp = finnhubGet_(`company-news?symbol=${ticker}&from=${from7d}&to=${todayS}`, apiKey);
    if (Array.isArray(resp)) rawNews = resp;
  } catch (e) { Logger.log("News: " + ticker + " – " + e.message); }
  Utilities.sleep(150);

  // ── News column: market-moving headlines only, blank otherwise ──
  const sig = rawNews.filter(n => {
    const h = (n.headline || "").toLowerCase();
    if (NOISE_PHRASES.some(p => h.includes(p))) return false;
    return MARKET_MOVING_KEYWORDS.some(k => h.includes(k));
  });
  const news = sig.length > 0 ? sig.slice(0, 2).map(n => n.headline).join("  |  ") : "";

  // ── Analyst consensus → Direction ──
  let direction = "Neutral";
  try {
    const resp = finnhubGet_(`stock/recommendation?symbol=${ticker}`, apiKey);
    if (Array.isArray(resp) && resp.length > 0) {
      const r       = resp[0];
      const bullish = (r.strongBuy || 0) + (r.buy  || 0);
      const bearish = (r.strongSell || 0) + (r.sell || 0);
      const neutral = r.hold || 0;
      if      (bullish > bearish && bullish >= neutral) direction = "Bullish";
      else if (bearish > bullish && bearish >= neutral) direction = "Bearish";
    }
  } catch (e) { Logger.log("Rec: " + ticker + " – " + e.message); }
  Utilities.sleep(150);

  // ── Upgrade / Downgrade ──
  // Primary: Finnhub stock/upgrade-downgrade endpoint (30-day lookback)
  // Fallback: parse upgrade/downgrade action directly from news headlines
  let upgradeDowngrade = "—";
  try {
    const resp = finnhubGet_(`stock/upgrade-downgrade?symbol=${ticker}&from=${from30d}`, apiKey);
    if (Array.isArray(resp) && resp.length > 0) {
      const r      = resp[0];
      const firm   = (r.company   || "").trim();
      const grade  = (r.toGrade   || "").trim();
      const date   = (r.gradeDate || "").trim();
      const action = (r.action    || "").replace(/^(up|down)grade$/i,
                       s => s[0].toUpperCase() + s.slice(1).toLowerCase());
      if (firm && grade) upgradeDowngrade = `${date}  ${firm}: ${action} → ${grade}`;
    }
  } catch (e) { Logger.log("UG API: " + ticker + " – " + e.message); }
  Utilities.sleep(150);

  // Fallback: scan news headlines for upgrade/downgrade pattern
  if (upgradeDowngrade === "—") {
    for (const item of rawNews) {
      const parsed = parseUDFromHeadline_(item.headline || "", item.datetime);
      if (parsed) { upgradeDowngrade = parsed; break; }
    }
  }

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

// Parses "[Firm] Upgrades/Downgrades [Company] to [Grade]" from a headline.
// Returns formatted string or null if no match.
function parseUDFromHeadline_(headline, unixtimestamp) {
  // Pattern: anything before action word, then "to <grade>"
  const match = headline.match(
    /^(.+?)\s+(upgrade[sd]?|downgrade[sd]?)\s+.+?\bto\s+([A-Z][A-Za-z\s\-+]+?)(?:\s*[|–—,]|$)/
  );
  if (!match) return null;

  const firm   = match[1].trim();
  const action = match[2].trim();
  const grade  = match[3].trim();

  // Skip if firm looks like a ticker or a generic phrase
  if (firm.length < 3 || /^[A-Z]{1,5}$/.test(firm)) return null;

  const date = unixtimestamp
    ? Utilities.formatDate(new Date(unixtimestamp * 1000), "America/New_York", "yyyy-MM-dd")
    : "";

  return `${date}  ${firm}: ${action[0].toUpperCase() + action.slice(1).toLowerCase()} → ${grade}`;
}

// ============================================================
// SHEET BUILDER
// ============================================================

function buildNewsFeedSheet_(ss, marketOverview, rows) {
  let sheet = ss.getSheetByName(NEWS_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(NEWS_SHEET_NAME);
  else sheet.clear();

  sheet.setColumnWidth(1, 80);   // Ticker
  sheet.setColumnWidth(2, 100);  // Direction
  sheet.setColumnWidth(3, 620);  // News
  sheet.setColumnWidth(4, 120);  // Earnings Date
  sheet.setColumnWidth(5, 280);  // Upgrade/Downgrade

  const dirBg = { Bullish: "#b6d7a8", Bearish: "#ea9999", Neutral: "#ffe599", "N/A": "#f3f3f3" };
  let r = 1;

  // ── Market direction banner ──
  sheet.getRange(r, 1, 1, 5).merge()
    .setValue(`MARKET DIRECTION: ${marketOverview.direction.toUpperCase()}   |   ${marketOverview.headline}`)
    .setBackground(dirBg[marketOverview.direction] || "#f3f3f3")
    .setFontWeight("bold").setFontSize(12).setFontFamily("Nunito")
    .setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  sheet.setRowHeight(r, 48);
  r++;

  // ── Timestamp ──
  sheet.getRange(r, 1, 1, 5).merge()
    .setValue("Refreshed " + Utilities.formatDate(new Date(), "America/New_York", "MMM dd, yyyy  h:mm a") + " ET")
    .setFontStyle("italic").setFontSize(9).setFontFamily("Nunito")
    .setHorizontalAlignment("center").setBackground("#f8f8f8");
  sheet.setRowHeight(r, 22);
  r++;

  r++; // blank spacer

  // ── Column headers ──
  sheet.getRange(r, 1, 1, 5)
    .setValues([["Ticker", "Direction", "News", "Earnings Date", "Upgrade / Downgrade"]])
    .setBackground("#cfe2f3").setFontWeight("bold").setFontSize(11)
    .setFontFamily("Nunito").setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.setRowHeight(r, 30);
  r++;

  const todayStr     = formatDate_(new Date());
  const sevenDaysOut = formatDate_(new Date(new Date() - -7 * 864e5));

  // ── Data rows ──
  rows.forEach(([ticker, direction, news, earningsDate, upgradeDowngrade]) => {
    // Col 1: Ticker
    sheet.getRange(r, 1).setValue(ticker).setHorizontalAlignment("center").setFontWeight("bold");

    // Col 2: Direction — bold + color
    const dirCell = sheet.getRange(r, 2);
    dirCell.setValue(direction).setHorizontalAlignment("center").setFontWeight("bold");
    if      (direction === "Bullish") dirCell.setFontColor("#38761d");
    else if (direction === "Bearish") dirCell.setFontColor("#cc0000");
    else                              dirCell.setFontColor("#7d6608");

    // Col 3: News
    sheet.getRange(r, 3).setValue(news).setWrap(true).setHorizontalAlignment("left");

    // Col 4: Earnings date — red + bold if within 7 days
    const earnCell    = sheet.getRange(r, 4);
    const earningsNear = earningsDate !== "—" && earningsDate >= todayStr && earningsDate <= sevenDaysOut;
    earnCell.setValue(earningsDate).setHorizontalAlignment("center");
    if (earningsNear) earnCell.setFontColor("#cc0000").setFontWeight("bold");

    // Col 5: Upgrade/Downgrade
    sheet.getRange(r, 5).setValue(upgradeDowngrade).setHorizontalAlignment("left");

    sheet.getRange(r, 1, 1, 5)
      .setBackground(r % 2 === 0 ? "#f9f9f9" : "#ffffff")
      .setFontFamily("Nunito").setFontSize(10);

    sheet.setRowHeight(r, 52);
    r++;
  });
}

// ============================================================
// SETUP HELPERS
// ============================================================

function createNewsFeedTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "refreshNewsFeed")
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("refreshNewsFeed")
    .timeBased().atHour(9).everyDays(1)
    .inTimezone("America/New_York").create();

  Logger.log("✅ Daily 9 AM ET trigger created for refreshNewsFeed.");
  SpreadsheetApp.getUi().alert("✅ Trigger set: News Feed will refresh every day at 9 AM ET.");
}

function setFinnhubApiKey() {
  const ui     = SpreadsheetApp.getUi();
  const result = ui.prompt("Finnhub API Key", "Paste your free API key from finnhub.io:", ui.ButtonSet.OK_CANCEL);
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

  const vals  = sheet.getRange(1, 1, sheet.getLastRow()).getValues();
  const skip  = new Set(["symbol", "totals", ...allowedAccounts.map(a => a.toLowerCase())]);
  const tickers = new Set();

  vals.forEach(([sym]) => {
    const s = (sym || "").toString().trim();
    if (!s || skip.has(s.toLowerCase())) return;
    if (ETF_TICKERS.has(s)) return;      // skip ETFs
    if (TICKER_RE.test(s)) tickers.add(s);
  });

  return [...tickers].sort();
}

function formatDate_(d) {
  return Utilities.formatDate(d, "America/New_York", "yyyy-MM-dd");
}
