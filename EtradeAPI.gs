/**
 * EtradeAPI.gs — Low-level E*Trade API wrappers.
 * Version: 1.3 (2026-04-06) — Add fetchEtradeQuotes_() for real-time change% from E*Trade market API
 *
 * All functions here deal directly with the E*Trade REST API.
 * Higher-level logic (sheet writes, UI) lives in other files.
 *
 * Depends on:
 *   Authentication.gs — buildOAuthHeaders_(), renewAccessToken()
 *   Code.gs           — BASE_URL, ACCOUNTS_URL, ACCOUNT_MAP
 */

// ─────────────────────────────────────────────────────────────────
// Token helper
// ─────────────────────────────────────────────────────────────────

/**
 * Returns the stored ACCESS_TOKEN and ACCESS_SECRET.
 * Throws a user-friendly error if tokens are missing.
 */
function getAccessTokens_() {
  var props  = PropertiesService.getScriptProperties();
  var token  = props.getProperty('ACCESS_TOKEN');
  var secret = props.getProperty('ACCESS_SECRET');
  if (!token || !secret) {
    throw new Error(
      'No E*Trade access token found.\n\n' +
      'Please complete OAuth authentication:\n' +
      '  E*Trade → 🔐 Authentication → steps 1–4'
    );
  }
  return { token: token, secret: secret };
}

// ─────────────────────────────────────────────────────────────────
// Balance
// ─────────────────────────────────────────────────────────────────

/**
 * Fetches the balance for one account and returns cash + totalAccountValue.
 * @param {string} accountKey   E*Trade accountIdKey
 * @param {string} accountId    Numeric account ID (used for logging only)
 * @param {string} instType     'BROKERAGE', 'BROKERAGE_IRA', etc.
 * @param {string} token        OAuth access token
 * @param {string} secret       OAuth access secret
 * @returns {{ cash: number, totalAccountValue: number }}
 */
function fetchEtradeBalance(accountKey, accountId, instType, token, secret) {
  var queryParams = {
    instType:     instType || 'BROKERAGE',
    realTimeNAV:  'true',
  };
  var queryString = Object.keys(queryParams).sort()
    .map(function(k) { return k + '=' + encodeURIComponent(queryParams[k]); })
    .join('&');

  var url     = ACCOUNTS_URL + '/' + accountKey + '/balance?' + queryString;
  var headers = mergeHeaders_(
    buildOAuthHeaders_(url, 'GET', token, secret),
    { Accept: 'application/json' }
  );

  var resp = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true, headers: headers });
  var code = resp.getResponseCode();
  var body = resp.getContentText();

  if (code >= 200 && code < 300) {
    var data = JSON.parse(body);
    var cb   = (data.BalanceResponse && data.BalanceResponse.Computed) || {};
    var rt   = cb.RealTimeValues || {};

    var cash              = Number(cb.cashAvailableForInvestment || 0);
    var totalAccountValue = Number(rt.totalAccountValue || 0);

    return {
      cash:              isFinite(cash)              ? cash              : 0,
      totalAccountValue: isFinite(totalAccountValue) ? totalAccountValue : 0,
    };
  }

  // Surface the actual API error so the user can diagnose it
  var errDetail = 'HTTP ' + code;
  try {
    var errData = JSON.parse(body);
    var msg = (errData.Error        && errData.Error.message)      ||
              (errData.fault        && errData.fault.faultstring)   ||
              (errData.error        && errData.error.description)   ||
              body.substring(0, 200);
    if (msg) errDetail += ' — ' + msg;
  } catch (_) {
    if (body) errDetail += ' — ' + body.substring(0, 200);
  }
  throw new Error('Balance API error for account …' + String(accountKey).slice(-6) + ': ' + errDetail);
}

// ─────────────────────────────────────────────────────────────────
// Net Liquidity fetch  (used by NetLiquidity.gs)
// ─────────────────────────────────────────────────────────────────

/**
 * Renews the token then fetches totalAccountValue for one account suffix.
 * @param {string} suffix  e.g. '7806'
 * @returns {number}  Total account value (net liquidation value)
 */
function fetchNetLiqForSuffix_(suffix) {
  renewAccessToken();
  var tokens = getAccessTokens_();
  var acct   = ACCOUNT_MAP[suffix];
  if (!acct) throw new Error('Unknown account suffix: ' + suffix);

  var bal = fetchEtradeBalance(acct.key, acct.id, acct.instType, tokens.token, tokens.secret);
  return bal.totalAccountValue;
}

// ─────────────────────────────────────────────────────────────────
// Account discovery  (used for initial Code.gs setup)
// ─────────────────────────────────────────────────────────────────

/**
 * Fetches /v1/accounts/list and displays each account's ID, accountIdKey,
 * and institutionType in an alert dialog so the user can copy them into
 * the ACCOUNT_MAP in Code.gs.
 *
 * Run once after completing OAuth authentication (steps 1–4).
 */
function showAccountKeys() {
  try {
    renewAccessToken();
    var tokens = getAccessTokens_();

    var url     = ACCOUNTS_URL + '/list.json';
    var headers = mergeHeaders_(
      buildOAuthHeaders_(url, 'GET', tokens.token, tokens.secret),
      { Accept: 'application/json' }
    );
    var resp = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true, headers: headers });
    var code = resp.getResponseCode();
    var body = resp.getContentText();

    if (code < 200 || code >= 300) {
      SpreadsheetApp.getUi().alert(
        'Account List Error',
        'HTTP ' + code + '\n\n' + body.substring(0, 500),
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    var data     = JSON.parse(body);
    var accounts = (data.AccountListResponse &&
                    data.AccountListResponse.Accounts &&
                    data.AccountListResponse.Accounts.Account) || [];

    if (accounts.length === 0) {
      SpreadsheetApp.getUi().alert('No accounts found in the response.');
      return;
    }

    var lines = ['Copy these values into ACCOUNT_MAP in Code.gs:\n'];
    accounts.forEach(function(a) {
      var id   = a.accountId        || '(unknown)';
      var key  = a.accountIdKey     || '(unknown)';
      var type = a.institutionType  || a.accountType || '(unknown)';
      var desc = a.accountDesc      || '';
      lines.push(
        '────────────────────────\n' +
        'Description : ' + desc + '\n' +
        'accountId   : ' + id   + '\n' +
        'accountIdKey: ' + key  + '\n' +
        'instType    : ' + type
      );
    });

    SpreadsheetApp.getUi().alert(
      '🔑 E*Trade Account Keys',
      lines.join('\n\n'),
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } catch (e) {
    SpreadsheetApp.getUi().alert('Error', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    console.error(e);
  }
}

// ─────────────────────────────────────────────────────────────────
// Utilities
// ─────────────────────────────────────────────────────────────────

/** Shallow-merges two header objects. */
function mergeHeaders_(h1, h2) {
  var out = {};
  for (var k in h1) out[k] = h1[k];
  for (var k in h2) out[k] = h2[k];
  return out;
}

// ─────────────────────────────────────────────────────────────────
// Market quotes  (used by AccountRefresh.gs)
// ─────────────────────────────────────────────────────────────────

/**
 * Fetches real-time quotes for a list of symbols via E*Trade's market API.
 * Returns a map of { SYMBOL: changePct } where changePct is a decimal (e.g. 0.0082).
 *
 * E*Trade supports up to 25 symbols per request; this function batches as needed.
 *
 * @param {string[]} symbols  e.g. ['AAPL', 'MSFT', 'GOOG']
 * @param {string}   token    OAuth access token
 * @param {string}   secret   OAuth access secret
 * @returns {Object} { 'AAPL': 0.0082, 'MSFT': -0.0031, … }
 */
function fetchEtradeQuotes_(symbols, token, secret) {
  if (!symbols || symbols.length === 0) return {};

  var result = {};
  var BATCH  = 25;

  for (var i = 0; i < symbols.length; i += BATCH) {
    var batch     = symbols.slice(i, i + BATCH);
    var symbolStr = batch.join(',');
    var url       = BASE_URL + '/v1/market/quote/' + symbolStr + '.json';
    var headers   = mergeHeaders_(
      buildOAuthHeaders_(url, 'GET', token, secret),
      { Accept: 'application/json' }
    );

    try {
      var resp = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true, headers: headers });
      var code = resp.getResponseCode();

      if (code < 200 || code >= 300) {
        console.warn('Quote API returned HTTP ' + code + ' for: ' + symbolStr);
        continue;
      }

      var data   = JSON.parse(resp.getContentText());
      var quotes = (data.QuoteResponse && data.QuoteResponse.QuoteData) || [];

      // QuoteData is an object (not array) when only one symbol is requested
      if (!Array.isArray(quotes)) quotes = [quotes];

      quotes.forEach(function(q) {
        var sym = q.Product && q.Product.symbol;
        var pct = q.All    && q.All.changeClosePercentage;
        if (sym && pct !== undefined && pct !== null) {
          result[sym] = pct / 100;   // API returns e.g. 0.82 → store as 0.0082
        }
      });

    } catch (e) {
      console.warn('fetchEtradeQuotes_ batch failed (' + symbolStr + '): ' + e.message);
    }
  }

  return result;
}

// ─────────────────────────────────────────────────────────────────
// Market data  (used by SPYHistory.gs)
// ─────────────────────────────────────────────────────────────────

/**
 * Fetches daily closing prices for a symbol via Yahoo Finance.
 * No authentication required — works independently of E*Trade OAuth.
 *
 * Returns the same shape as Schwab's getPriceHistory() so SPYHistory.gs
 * can be used unchanged:
 *   { candles: [{ datetime: <ms>, close: <number> }], symbol, empty }
 *
 * @param {string} symbol     e.g. 'SPY'
 * @param {string} startDate  'YYYY-MM-DD'
 * @param {string} endDate    'YYYY-MM-DD'
 */
function getPriceHistory(symbol, startDate, endDate) {
  var period1 = Math.floor(new Date(startDate).getTime() / 1000);
  var period2 = Math.floor(new Date(endDate).getTime()   / 1000);

  var url = 'https://query1.finance.yahoo.com/v8/finance/chart/' + symbol +
    '?period1=' + period1 +
    '&period2=' + period2 +
    '&interval=1d' +
    '&events=history';

  try {
    var resp   = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var data   = JSON.parse(resp.getContentText());
    var result = data.chart && data.chart.result && data.chart.result[0];
    if (!result) return { candles: [], symbol: symbol, empty: true };

    var timestamps = result.timestamp || [];
    var closes     = (result.indicators &&
                      result.indicators.quote &&
                      result.indicators.quote[0] &&
                      result.indicators.quote[0].close) || [];

    var candles = [];
    for (var i = 0; i < timestamps.length; i++) {
      if (closes[i] && closes[i] > 0) {
        candles.push({ datetime: timestamps[i] * 1000, close: closes[i] });
      }
    }
    return { candles: candles, symbol: symbol, empty: candles.length === 0 };
  } catch (e) {
    console.error('getPriceHistory failed for ' + symbol + ': ' + e.message);
    return { candles: [], symbol: symbol, empty: true };
  }
}
