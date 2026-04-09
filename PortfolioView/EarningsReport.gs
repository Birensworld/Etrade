// ============================================================
// EARNINGS REPORT                                Version: 1.0
// Runs twice daily: 8 AM ET (pre-market) & 5 PM ET (post-close).
// Notable = HIGH_PROFILE company OR YoY EPS increase > 25%.
// Emails birensworld@gmail.com on each refresh.
//
// Shared utilities (finnhubGet_, formatDate_, FINNHUB_KEY_PROP)
// are defined in NewsFeed.gs — same GAS project, same scope.
// ============================================================

const EARNINGS_SHEET_NAME = "Earnings Report";
const EARNINGS_EMAIL      = "birensworld@gmail.com";

// Always included when reporting today, regardless of YoY change.
const HIGH_PROFILE = new Set([
  // Mega-cap tech
  "AAPL","MSFT","GOOGL","GOOG","AMZN","META","NVDA","TSLA","NFLX",
  "AMD","INTC","QCOM","AVGO","TXN","MU","AMAT","ASML","TSM","SMCI",
  // Cloud & SaaS
  "CRM","ORCL","SAP","NOW","SNOW","PLTR","DDOG","CRWD","PANW","ZS",
  "OKTA","HUBS","NET","TWLO","MDB","BILL","GTLB","TEAM","ZM",
  // Consumer tech / platforms
  "UBER","LYFT","ABNB","SHOP","SQ","PYPL","COIN","ROKU","RBLX",
  "SNAP","PINS","SPOT","U",
  // Finance
  "JPM","BAC","GS","MS","WFC","C","BLK","AXP","V","MA","SCHW",
  "BRK.B","BRK.A","HOOD",
  // Healthcare & Pharma
  "JNJ","UNH","PFE","ABBV","MRK","LLY","BMY","AMGN","GILD",
  "CVS","CI","HUM","ISRG","MRNA","BNTX","ILMN",
  // Energy
  "XOM","CVX","COP","SLB","OXY",
  // Consumer & Retail
  "WMT","TGT","COST","HD","LOW","NKE","DIS","SBUX","MCD",
  "KO","PEP","PG","MDLZ","CL",
  // Industrial & Defense
  "BA","CAT","DE","GE","HON","RTX","LMT","NOC","GD","MMM",
  // Telecom / Media
  "T","VZ","CMCSA","WBD","PARA",
  // EV & Auto
  "F","GM","RIVN","LCID","TM","STLA",
  // Semis (extended)
  "LRCX","KLAC","MRVL","MCHP","ON","SWKS"
]);

// Display names for known tickers (avoids extra API calls)
const COMPANY_NAMES = {
  "AAPL":"Apple","MSFT":"Microsoft","GOOGL":"Alphabet","GOOG":"Alphabet",
  "AMZN":"Amazon","META":"Meta","NVDA":"NVIDIA","TSLA":"Tesla","NFLX":"Netflix",
  "AMD":"AMD","INTC":"Intel","QCOM":"Qualcomm","AVGO":"Broadcom","TXN":"Texas Instruments",
  "MU":"Micron","AMAT":"Applied Materials","ASML":"ASML","TSM":"TSMC","SMCI":"Super Micro",
  "CRM":"Salesforce","ORCL":"Oracle","SAP":"SAP","NOW":"ServiceNow","SNOW":"Snowflake",
  "PLTR":"Palantir","DDOG":"Datadog","CRWD":"CrowdStrike","PANW":"Palo Alto Networks",
  "ZS":"Zscaler","OKTA":"Okta","HUBS":"HubSpot","NET":"Cloudflare","TWLO":"Twilio",
  "MDB":"MongoDB","BILL":"Bill.com","GTLB":"GitLab","TEAM":"Atlassian","ZM":"Zoom",
  "UBER":"Uber","LYFT":"Lyft","ABNB":"Airbnb","SHOP":"Shopify","SQ":"Block",
  "PYPL":"PayPal","COIN":"Coinbase","ROKU":"Roku","RBLX":"Roblox",
  "SNAP":"Snap","PINS":"Pinterest","SPOT":"Spotify","U":"Unity",
  "JPM":"JPMorgan","BAC":"Bank of America","GS":"Goldman Sachs","MS":"Morgan Stanley",
  "WFC":"Wells Fargo","C":"Citigroup","BLK":"BlackRock","AXP":"Amex",
  "V":"Visa","MA":"Mastercard","SCHW":"Schwab","BRK.B":"Berkshire","HOOD":"Robinhood",
  "JNJ":"J&J","UNH":"UnitedHealth","PFE":"Pfizer","ABBV":"AbbVie",
  "MRK":"Merck","LLY":"Eli Lilly","BMY":"Bristol-Myers","AMGN":"Amgen",
  "GILD":"Gilead","CVS":"CVS","CI":"Cigna","HUM":"Humana",
  "ISRG":"Intuitive Surgical","MRNA":"Moderna","BNTX":"BioNTech",
  "XOM":"ExxonMobil","CVX":"Chevron","COP":"ConocoPhillips","SLB":"SLB","OXY":"Occidental",
  "WMT":"Walmart","TGT":"Target","COST":"Costco","HD":"Home Depot","LOW":"Lowe's",
  "NKE":"Nike","DIS":"Disney","SBUX":"Starbucks","MCD":"McDonald's",
  "KO":"Coca-Cola","PEP":"PepsiCo","PG":"Procter & Gamble",
  "BA":"Boeing","CAT":"Caterpillar","DE":"Deere","GE":"GE","HON":"Honeywell",
  "RTX":"RTX","LMT":"Lockheed Martin","NOC":"Northrop","GD":"General Dynamics",
  "T":"AT&T","VZ":"Verizon","CMCSA":"Comcast",
  "F":"Ford","GM":"General Motors","RIVN":"Rivian","LCID":"Lucid"
};

// ============================================================
// ENTRY POINT
// ============================================================

function refreshEarningsReport() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const apiKey = PropertiesService.getScriptProperties().getProperty(FINNHUB_KEY_PROP);

  if (!apiKey) {
    SpreadsheetApp.getUi().alert("Finnhub API key not set. Run setFinnhubApiKey() first.");
    return;
  }

  const todayStr = formatDate_(new Date());
  Logger.log("Fetching earnings for " + todayStr);

  const allEarnings = fetchTodayEarnings_(todayStr, apiKey);

  if (!allEarnings.length) {
    buildEarningsSheet_(ss, [], todayStr);
    sendEarningsEmail_([], todayStr);
    return;
  }

  const notableRows = [];
  allEarnings.forEach(e => {
    try {
      const row = buildNotableRow_(e, apiKey);
      if (row) notableRows.push(row);
    } catch (err) {
      Logger.log("⚠ " + e.symbol + ": " + err.message);
    }
    Utilities.sleep(250);
  });

  // Sort: high-profile first, then alpha
  notableRows.sort((a, b) => {
    if (a.isHP && !b.isHP) return -1;
    if (!a.isHP && b.isHP) return 1;
    return a.ticker.localeCompare(b.ticker);
  });

  buildEarningsSheet_(ss, notableRows, todayStr);
  sendEarningsEmail_(notableRows, todayStr);
}

// ============================================================
// DATA FETCHING
// ============================================================

function fetchTodayEarnings_(dateStr, apiKey) {
  try {
    const resp = finnhubGet_(`calendar/earnings?from=${dateStr}&to=${dateStr}`, apiKey);
    return ((resp && resp.earningsCalendar) || []).filter(e => e && e.symbol);
  } catch (e) {
    Logger.log("Earnings calendar: " + e.message);
    return [];
  }
}

// Returns a row object if notable, null otherwise.
function buildNotableRow_(e, apiKey) {
  const ticker    = (e.symbol || "").trim().toUpperCase();
  const isHP      = HIGH_PROFILE.has(ticker);
  const hasActual = e.epsActual !== null && e.epsActual !== undefined;

  // No actual EPS yet → only keep if HIGH_PROFILE (show as "Pending")
  if (!hasActual && !isHP) return null;

  let yoyPct     = null;
  let yoyNotable = false;

  if (hasActual) {
    yoyPct     = fetchYoYChange_(ticker, apiKey);
    yoyNotable = yoyPct !== null && yoyPct > 25;
  }

  if (!isHP && !yoyNotable) return null;

  // ── Format values ──
  const epsActual = hasActual ? e.epsActual : "Pending";
  const epsEst    = (e.epsEstimate !== null && e.epsEstimate !== undefined) ? e.epsEstimate : "—";

  let epsSurprise = "—";
  if (hasActual && typeof epsEst === "number" && epsEst !== 0) {
    const pct = ((e.epsActual - epsEst) / Math.abs(epsEst)) * 100;
    epsSurprise = (pct >= 0 ? "+" : "") + pct.toFixed(1) + "%";
  }

  const fmtRev = v => v ? (v / 1e9).toFixed(2) + "B" : "—";
  const revActual = hasActual ? fmtRev(e.revenueActual) : "Pending";
  const revEst    = fmtRev(e.revenueEstimate);

  let revSurprise = "—";
  if (hasActual && e.revenueActual && e.revenueEstimate && e.revenueEstimate !== 0) {
    const pct = ((e.revenueActual - e.revenueEstimate) / Math.abs(e.revenueEstimate)) * 100;
    revSurprise = (pct >= 0 ? "+" : "") + pct.toFixed(1) + "%";
  }

  const yoyDisplay = yoyPct !== null
    ? (yoyPct >= 0 ? "+" : "") + yoyPct.toFixed(1) + "%"
    : "—";

  const reason = [
    isHP       ? "High Profile"                   : "",
    yoyNotable ? `YoY EPS +${yoyPct.toFixed(1)}%` : ""
  ].filter(Boolean).join("  |  ");

  return {
    ticker,
    company:    COMPANY_NAMES[ticker] || ticker,
    session:    (e.hour || "").toUpperCase() || "—",
    epsActual,  epsEst,      epsSurprise,
    revActual,  revEst,      revSurprise,
    yoyDisplay, reason,
    isHP,       yoyNotable
  };
}

// Returns YoY EPS % change (current quarter vs same quarter last year).
function fetchYoYChange_(ticker, apiKey) {
  try {
    const hist = finnhubGet_(`stock/earnings?symbol=${ticker}&limit=8`, apiKey);
    if (!Array.isArray(hist) || hist.length < 5) return null;

    const curr = hist[0].actual;
    const prev = hist[4].actual; // same quarter, prior year

    if (curr == null || prev == null || prev === 0) return null;
    return ((curr - prev) / Math.abs(prev)) * 100;
  } catch (e) {
    Logger.log("YoY " + ticker + ": " + e.message);
    return null;
  }
}

// ============================================================
// SHEET BUILDER
// ============================================================

function buildEarningsSheet_(ss, rows, dateStr) {
  let sheet = ss.getSheetByName(EARNINGS_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(EARNINGS_SHEET_NAME);
  else sheet.clear();

  sheet.setColumnWidth(1,  80);   // Ticker
  sheet.setColumnWidth(2,  160);  // Company
  sheet.setColumnWidth(3,  75);   // Session
  sheet.setColumnWidth(4,  90);   // EPS Actual
  sheet.setColumnWidth(5,  90);   // EPS Est.
  sheet.setColumnWidth(6,  100);  // EPS Surprise
  sheet.setColumnWidth(7,  110);  // Rev Actual
  sheet.setColumnWidth(8,  110);  // Rev Est.
  sheet.setColumnWidth(9,  110);  // Rev Surprise
  sheet.setColumnWidth(10, 110);  // YoY EPS %
  sheet.setColumnWidth(11, 220);  // Why Notable

  let r = 1;

  // ── Banner ──
  sheet.getRange(r, 1, 1, 11).merge()
    .setValue("EARNINGS REPORT  —  " + dateStr)
    .setBackground("#d9ead3").setFontWeight("bold").setFontSize(13)
    .setFontFamily("Nunito").setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.setRowHeight(r, 40);
  r++;

  // ── Timestamp ──
  sheet.getRange(r, 1, 1, 11).merge()
    .setValue("Refreshed " + Utilities.formatDate(new Date(), "America/New_York", "MMM dd, yyyy  h:mm a") + " ET")
    .setFontStyle("italic").setFontSize(9).setFontFamily("Nunito")
    .setHorizontalAlignment("center").setBackground("#f8f8f8");
  sheet.setRowHeight(r, 20);
  r++;

  r++; // spacer

  if (!rows.length) {
    sheet.getRange(r, 1, 1, 11).merge()
      .setValue("No notable earnings today.")
      .setHorizontalAlignment("center").setFontStyle("italic").setFontFamily("Nunito");
    return;
  }

  // ── Headers ──
  const headers = ["Ticker","Company","Session","EPS Actual","EPS Est.",
                   "EPS Surprise","Rev Actual","Rev Est.","Rev Surprise",
                   "YoY EPS %","Why Notable"];
  sheet.getRange(r, 1, 1, 11).setValues([headers])
    .setBackground("#cfe2f3").setFontWeight("bold").setFontSize(10)
    .setFontFamily("Nunito").setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.setRowHeight(r, 28);
  r++;

  // ── Data rows ──
  rows.forEach(row => {
    const vals = [row.ticker, row.company, row.session,
                  row.epsActual, row.epsEst, row.epsSurprise,
                  row.revActual, row.revEst, row.revSurprise,
                  row.yoyDisplay, row.reason];

    sheet.getRange(r, 1, 1, 11)
      .setValues([vals]).setFontFamily("Nunito").setFontSize(10)
      .setBackground(r % 2 === 0 ? "#f9f9f9" : "#ffffff")
      .setHorizontalAlignment("center");

    sheet.getRange(r, 1).setFontWeight("bold"); // ticker bold
    sheet.getRange(r, 2).setHorizontalAlignment("left");
    sheet.getRange(r, 11).setHorizontalAlignment("left");

    // Color EPS surprise
    colorCell_(sheet.getRange(r, 6), row.epsSurprise);
    // Color Rev surprise
    colorCell_(sheet.getRange(r, 9), row.revSurprise);
    // Color YoY EPS
    colorCell_(sheet.getRange(r, 10), row.yoyDisplay);

    sheet.setRowHeight(r, 26);
    r++;
  });
}

// Colors a cell green/red/black based on a +/- prefixed percent string
function colorCell_(cell, value) {
  const n = parseFloat((value || "").replace(/[+%]/g, ""));
  if (isNaN(n)) return;
  cell.setFontColor(n > 0 ? "#38761d" : n < 0 ? "#cc0000" : "#000000")
      .setFontWeight("bold");
}

// ============================================================
// EMAIL
// ============================================================

function sendEarningsEmail_(rows, dateStr) {
  const ts      = Utilities.formatDate(new Date(), "America/New_York", "MMM dd, yyyy h:mm a") + " ET";
  const subject = `Earnings Report — ${dateStr}  (${rows.length} notable)`;

  if (!rows.length) {
    GmailApp.sendEmail(EARNINGS_EMAIL, subject, "", {
      htmlBody: `<p style="font-family:Arial;color:#555;">No notable earnings today (${dateStr}).</p>`
    });
    return;
  }

  const th = s => `<th style="background:#cfe2f3;font-weight:bold;padding:6px 10px;
                    border:1px solid #aaa;font-family:Arial;font-size:12px;">${s}</th>`;
  const td = (s, extra) => `<td style="padding:5px 10px;border:1px solid #ddd;
                    font-family:Arial;font-size:12px;text-align:center;${extra||''}">${s}</td>`;

  const colorStyle = v => {
    const n = parseFloat((v || "").replace(/[+%]/g, ""));
    if (isNaN(n)) return "";
    return `color:${n > 0 ? "#38761d" : n < 0 ? "#cc0000" : "#000"};font-weight:bold;`;
  };

  const headerRow = ["Ticker","Company","Session","EPS Actual","EPS Est.",
                     "EPS Surprise","Rev Actual","Rev Est.","Rev Surprise",
                     "YoY EPS %","Why Notable"].map(th).join("");

  const dataRows = rows.map((row, i) => {
    const bg = `background:${i % 2 === 0 ? "#fff" : "#f9f9f9"};`;
    return `<tr style="${bg}">
      ${td(row.ticker,       "font-weight:bold;")}
      ${td(row.company,      "text-align:left;")}
      ${td(row.session)}
      ${td(row.epsActual)}
      ${td(row.epsEst)}
      ${td(row.epsSurprise,  colorStyle(row.epsSurprise))}
      ${td(row.revActual)}
      ${td(row.revEst)}
      ${td(row.revSurprise,  colorStyle(row.revSurprise))}
      ${td(row.yoyDisplay,   colorStyle(row.yoyDisplay))}
      ${td(row.reason,       "text-align:left;")}
    </tr>`;
  }).join("");

  const html = `
    <div style="font-family:Arial;max-width:1200px;margin:0 auto;">
      <h2 style="color:#333;border-bottom:3px solid #d9ead3;padding-bottom:8px;">
        Earnings Report — ${dateStr}
      </h2>
      <p style="color:#666;font-size:12px;">${rows.length} notable  |  Refreshed ${ts}</p>
      <table style="border-collapse:collapse;width:100%;">
        <thead><tr>${headerRow}</tr></thead>
        <tbody>${dataRows}</tbody>
      </table>
      <p style="color:#aaa;font-size:10px;margin-top:16px;">
        Notable: High-profile company &nbsp;|&nbsp; YoY EPS increase &gt; 25%
      </p>
    </div>`;

  GmailApp.sendEmail(EARNINGS_EMAIL, subject, "", { htmlBody: html });
  Logger.log("✅ Earnings email sent → " + EARNINGS_EMAIL);
}

// ============================================================
// TRIGGERS
// ============================================================

function createEarningsTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "refreshEarningsReport")
    .forEach(t => ScriptApp.deleteTrigger(t));

  // 8 AM ET  — pre-market / BMO reporters
  ScriptApp.newTrigger("refreshEarningsReport")
    .timeBased().atHour(8).everyDays(1)
    .inTimezone("America/New_York").create();

  // 5 PM ET — post-close / AMC reporters
  ScriptApp.newTrigger("refreshEarningsReport")
    .timeBased().atHour(17).everyDays(1)
    .inTimezone("America/New_York").create();

  Logger.log("✅ Earnings triggers set: 8 AM & 5 PM ET daily.");
  SpreadsheetApp.getUi().alert("✅ Earnings Report will refresh at 8 AM & 5 PM ET daily.\nEmail → " + EARNINGS_EMAIL);
}
