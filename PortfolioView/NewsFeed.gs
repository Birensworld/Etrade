// ============================================================
// NEWS FEED                                      Version: 1.8
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

// Headlines starting with these words are speculative questions, not events.
// e.g. "Will SCHW Beat Estimates Again?" — not actionable news.
const QUESTION_STARTS = [
  "will ", "can ", "could ", "should ", "is ", "are ", "has ", "have ",
  "does ", "did ", "what ", "when ", "why ", "how ", "which ", "who "
];

// Headlines matching any of these are suppressed — analyst actions go to the
// U/D column; the rest are generic non-events or speculative preview content.
const NOISE_PHRASES = [
  // Analyst rating actions → U/D column handles these
  "upgrades ", "downgrades ", "upgraded to", "downgraded to",
  "reiterates", "maintains rating", "maintains buy", "maintains hold",
  "initiates with", "initiates coverage",
  // Earnings speculation / previews — not actual results
  "beat estimates again", "miss estimates again",
  "earnings preview", "earnings outlook", "earnings expectations",
  "what to expect", "what analysts expect",
  "ahead of earnings", "before earnings", "in its next earnings",
  "next earnings report", "next quarter earnings",
  "preview:", "outlook:", "looking ahead",
  "stock prediction", "price target", "pt to ",
  // Generic non-events
  "to present at", "to speak at", "conference call", "webcast",
  "names new vp", "names new director", "promotes", "appoints vp",
  "monthly traffic", "weekly data", "analyst day", "investor day",
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
  const from7d  = formatDate_(new Date(today - 7  * 864e5));  // 7 days back  (news window)
  const from90d = formatDate_(new Date(today - 90 * 864e5));  // 90 days back (U/D primary lookback)
  const todayS  = formatDate_(today);
  const to90d   = formatDate_(new Date(today - -90 * 864e5)); // 90 days ahead (earnings)

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
    // Drop speculative questions ("Will X beat estimates?", "Can Y recover?", etc.)
    if (QUESTION_STARTS.some(q => h.startsWith(q))) return false;
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
  // Layer 1: Finnhub stock/upgrade-downgrade (90-day lookback)
  // Layer 2: Yahoo Finance quoteSummary upgradeDowngradeHistory (no key needed, great coverage)
  // Layer 3: Parse directly from Finnhub news headlines (regex fallback)
  const UD_ACTION_MAP = { up: "Upgrade", down: "Downgrade", init: "Initiates" };
  let upgradeDowngrade = "—";

  // --- Layer 1: Finnhub ---
  try {
    const resp = finnhubGet_(`stock/upgrade-downgrade?symbol=${ticker}&from=${from90d}`, apiKey);
    if (Array.isArray(resp) && resp.length > 0) {
      const entry = resp.find(r => UD_ACTION_MAP[(r.action || "").toLowerCase()]);
      if (entry) {
        const firm        = (entry.company   || "").trim();
        const grade       = (entry.toGrade   || "").trim();
        const date        = (entry.gradeDate || "").trim();
        const actionLabel = UD_ACTION_MAP[(entry.action || "").toLowerCase()];
        if (firm && grade) upgradeDowngrade = `${date}  ${firm}: ${actionLabel} → ${grade}`;
      }
    }
  } catch (e) { Logger.log("UG Finnhub: " + ticker + " – " + e.message); }
  Utilities.sleep(150);

  // --- Layer 2: Yahoo Finance ---
  if (upgradeDowngrade === "—") {
    const yahoo = fetchUDFromYahoo_(ticker);
    if (yahoo) upgradeDowngrade = yahoo;
  }

  // --- Layer 3: News headline regex ---
  if (upgradeDowngrade === "—") {
    for (const item of rawNews) {
      const parsed = parseUDFromHeadline_(item.headline || "", item.datetime, ticker);
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

// Fetches the most-recent directional analyst rating from Yahoo Finance.
// Uses the public quoteSummary endpoint — no API key required.
// Returns formatted string "date  Firm: Action → Grade" or null.
function fetchUDFromYahoo_(ticker) {
  try {
    const url = "https://query2.finance.yahoo.com/v10/finance/quoteSummary/" +
                encodeURIComponent(ticker) +
                "?modules=upgradeDowngradeHistory&corsDomain=finance.yahoo.com";
    const resp = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: { "User-Agent": "Mozilla/5.0" }
    });
    if (resp.getResponseCode() !== 200) return null;

    const history =
      JSON.parse(resp.getContentText())
        ?.quoteSummary?.result?.[0]
        ?.upgradeDowngradeHistory?.history;
    if (!Array.isArray(history) || history.length === 0) return null;

    const ACTION_MAP = { up: "Upgrade", down: "Downgrade", init: "Initiates" };
    // Most-recent directional action (skip "main" = maintains, "reit" = reiterates)
    const entry = history.find(h => ACTION_MAP[(h.action || "").toLowerCase()]);
    if (!entry) return null;

    const firm        = (entry.firm    || "").trim();
    const grade       = (entry.toGrade || "").trim();
    const actionLabel = ACTION_MAP[(entry.action || "").toLowerCase()];
    if (!firm || !grade) return null;

    const date = entry.epochGradeDate
      ? Utilities.formatDate(new Date(entry.epochGradeDate * 1000), "America/New_York", "yyyy-MM-dd")
      : "";
    return `${date}  ${firm}: ${actionLabel} → ${grade}`;
  } catch (e) {
    Logger.log("UG Yahoo: " + e.message);
    return null;
  }
}

// Parses upgrade/downgrade info from a news headline.
// contextTicker: the stock we're analyzing (used to reject headlines where the
//   "firm" is actually the company being rated, not the analyst firm).
//
// Handles two common formats:
//   Pattern 1: "Firm Upgrades/Downgrades Company to Grade [from OldGrade]"
//              e.g. "UBS Downgrades NOW to Neutral from Buy, Lowers PT..."
//   Pattern 2: "Company Upgraded/Downgraded to Grade at/by Firm"
//              e.g. "NOW downgraded to Neutral at UBS"
// Returns formatted "date  Firm: Action → Grade" string or null.
function parseUDFromHeadline_(headline, unixtimestamp, contextTicker) {
  const h   = headline.trim();
  const ctx = (contextTicker || "").toUpperCase();

  function makeResult_(act, firm, rawGrade, ts) {
    if (!firm || firm.length < 2) return null;
    // Reject if firm is actually the ticker we're analyzing (headline starts with ticker)
    if (ctx && firm.replace(/[^A-Z]/gi, "").toUpperCase() === ctx) return null;
    // Clean grade: strip "from OldGrade" suffix  e.g. "Neutral from Buy" → "Neutral"
    const grade = rawGrade.trim().replace(/\s+from\s+.*/i, "").replace(/[,;].*/, "").trim();
    if (!grade) return null;
    const label = act[0].toUpperCase() + act.slice(1).toLowerCase();
    const date  = ts
      ? Utilities.formatDate(new Date(ts * 1000), "America/New_York", "yyyy-MM-dd")
      : "";
    return `${date}  ${firm}: ${label} → ${grade}`;
  }

  // Pattern 1: headline starts with analyst firm (most common Benzinga/Finnhub format)
  // "UBS Downgrades NOW to Neutral from Buy, Lowers PT to $100 from $170"
  const m1 = h.match(
    /^(.+?)\s+(upgrade[sd]?|downgrade[sd]?|initiates?(?:\s+coverage)?)\s+.+?\bto\s+([A-Za-z][A-Za-z\s\-+]+?)(?:\s*[|–—,]|$)/i
  );
  if (m1) {
    const r = makeResult_(m1[2], m1[1].trim(), m1[3], unixtimestamp);
    if (r) return r;
  }

  // Pattern 2: "Company Upgraded/Downgraded to Grade at/by Firm"
  const m2 = h.match(
    /\b(upgrade[sd]?|downgrade[sd]?|initiates?(?:\s+coverage)?)\s+to\s+([A-Za-z][A-Za-z\s\-+]+?)\s+(?:at|by)\s+(.+?)(?:\s*[|–—,]|$)/i
  );
  if (m2) {
    const r = makeResult_(m2[1], m2[3].trim(), m2[2], unixtimestamp);
    if (r) return r;
  }

  return null;
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
