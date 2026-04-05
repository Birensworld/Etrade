/**
 * Authentication.gs — E*Trade OAuth 1.0a flow.
 * Version: 1.0 (2026-04-05)
 *
 * Flow:
 *   1. Reset Auth          — clears all stored tokens
 *   2. Get Request Token   — exchanges consumer key/secret for a request token
 *                            and shows the authorization URL (user visits + gets PIN)
 *   3. Set Verifier (PIN)  — stores the PIN the user received from E*Trade
 *   4. Get Access Token    — exchanges request token + PIN for an access token
 *   (Renew Access Token)   — refreshes the access token (call before any API request)
 *
 * Tokens are stored in Script Properties:
 *   REQUEST_TOKEN  / REQUEST_SECRET  — temporary, used during auth only
 *   VERIFIER                         — PIN entered by user
 *   ACCESS_TOKEN   / ACCESS_SECRET   — used for all subsequent API calls
 *
 * Constants BASE_URL, CONSUMER_KEY, CONSUMER_SECRET are defined in Code.gs.
 */

// ─────────────────────────────────────────────────────────────────
// Menu entry points
// ─────────────────────────────────────────────────────────────────

function menuGetRequestToken() {
  getRequestToken();
}

function menuSetVerifier() {
  var ui       = SpreadsheetApp.getUi();
  var response = ui.prompt('Enter the Verifier (PIN) from E*TRADE:');
  if (response.getSelectedButton() === ui.Button.OK) {
    setVerifier(response.getResponseText().trim());
    ui.alert('PIN saved. Now run "Get Access Token".');
  }
}

function menuGetAccessToken() {
  getAccessToken();
}

// ─────────────────────────────────────────────────────────────────
// Step 1 — Request Token
// ─────────────────────────────────────────────────────────────────

function getRequestToken() {
  var url      = BASE_URL + '/oauth/request_token';
  var headers  = buildOAuthHeaders_(url, 'POST');
  var response = UrlFetchApp.fetch(url, {
    method: 'post',
    muteHttpExceptions: true,
    headers: headers,
  });

  var params = parseOAuthParams_(response.getContentText());
  PropertiesService.getScriptProperties().setProperties({
    REQUEST_TOKEN:  params.oauth_token,
    REQUEST_SECRET: params.oauth_token_secret,
  });

  var authUrl = 'https://us.etrade.com/e/t/etws/authorize' +
    '?key='   + CONSUMER_KEY +
    '&token=' + params.oauth_token +
    '&scope=readwrite';

  SpreadsheetApp.getUi().alert(
    'Step 2 — Authorize',
    'Visit this URL in your browser, log in to E*Trade, and note the PIN:\n\n' + authUrl,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ─────────────────────────────────────────────────────────────────
// Step 2 — Store Verifier (PIN)
// ─────────────────────────────────────────────────────────────────

function setVerifier(verifier) {
  PropertiesService.getScriptProperties().setProperty('VERIFIER', verifier);
}

// ─────────────────────────────────────────────────────────────────
// Step 3 — Access Token
// ─────────────────────────────────────────────────────────────────

function getAccessToken() {
  var props    = PropertiesService.getScriptProperties();
  var token    = props.getProperty('REQUEST_TOKEN');
  var secret   = props.getProperty('REQUEST_SECRET');
  var verifier = props.getProperty('VERIFIER');

  if (!token || !verifier) {
    SpreadsheetApp.getUi().alert('Missing request token or verifier. Complete steps 1–3 first.');
    return;
  }

  var url      = BASE_URL + '/oauth/access_token';
  var headers  = buildOAuthHeaders_(url, 'POST', token, secret, verifier);
  var response = UrlFetchApp.fetch(url, {
    method: 'post',
    muteHttpExceptions: true,
    headers: headers,
  });

  var params = parseOAuthParams_(response.getContentText());
  if (params.oauth_token && params.oauth_token_secret) {
    props.setProperties({
      ACCESS_TOKEN:  params.oauth_token,
      ACCESS_SECRET: params.oauth_token_secret,
    });
    SpreadsheetApp.getUi().alert('✅ Access token saved. You can now use the API.');
  } else {
    SpreadsheetApp.getUi().alert('⚠️ No access token in response:\n\n' + response.getContentText());
  }
}

// ─────────────────────────────────────────────────────────────────
// Renew — call before any API request (E*Trade tokens expire daily)
// ─────────────────────────────────────────────────────────────────

function renewAccessToken() {
  var props  = PropertiesService.getScriptProperties();
  var token  = props.getProperty('ACCESS_TOKEN');
  var secret = props.getProperty('ACCESS_SECRET');
  if (!token || !secret) {
    throw new Error('No access token found. Complete OAuth authentication first.');
  }

  var url      = BASE_URL + '/oauth/renew_access_token';
  var headers  = buildOAuthHeaders_(url, 'GET', token, secret);
  var response = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true,
    headers: headers,
  });

  var params = parseOAuthParams_(response.getContentText());
  if (params.oauth_token && params.oauth_token_secret) {
    props.setProperties({
      ACCESS_TOKEN:  params.oauth_token,
      ACCESS_SECRET: params.oauth_token_secret,
    });
  } else {
    console.warn('Token renewal returned no tokens — continuing with existing token.');
  }
}

// ─────────────────────────────────────────────────────────────────
// Reset
// ─────────────────────────────────────────────────────────────────

function resetAuth() {
  PropertiesService.getScriptProperties().deleteAllProperties();
  SpreadsheetApp.getUi().alert('All OAuth credentials cleared.');
}

// ─────────────────────────────────────────────────────────────────
// OAuth 1.0a helpers  (private — used only within this file and EtradeAPI.gs)
// ─────────────────────────────────────────────────────────────────

/**
 * Builds the Authorization header for an OAuth 1.0a signed request.
 * @param {string}  url          Full URL (may include query string)
 * @param {string}  method       HTTP method ('GET' or 'POST')
 * @param {string=} token        OAuth token (omit for request_token step)
 * @param {string=} tokenSecret  OAuth token secret
 * @param {string=} verifier     OAuth verifier / PIN
 */
function buildOAuthHeaders_(url, method, token, tokenSecret, verifier) {
  var oauthParams = {
    oauth_consumer_key:     CONSUMER_KEY,
    oauth_nonce:            Math.random().toString(36).substring(2),
    oauth_signature_method: 'HMAC-SHA1',
    oauth_timestamp:        Math.floor(new Date().getTime() / 1000),
    oauth_version:          '1.0',
  };

  if (url.indexOf('/request_token') !== -1) {
    oauthParams.oauth_callback = 'oob';
  }
  if (token) {
    oauthParams.oauth_token = decodeURIComponent(token);
  }
  if (verifier) {
    oauthParams.oauth_verifier = verifier;
  }

  oauthParams.oauth_signature = signRequest_(
    url, method, oauthParams,
    tokenSecret ? decodeURIComponent(tokenSecret) : null
  );

  var header = 'OAuth ' + Object.keys(oauthParams).sort().map(function(k) {
    return k + '="' + encodeURIComponent(oauthParams[k]) + '"';
  }).join(', ');

  return { Authorization: header };
}

/**
 * Computes the HMAC-SHA1 OAuth signature.
 */
function signRequest_(url, method, params, tokenSecret) {
  var baseUrl     = url;
  var queryParams = {};

  if (url.indexOf('?') !== -1) {
    var parts       = url.split('?');
    baseUrl         = parts[0];
    parts[1].split('&').forEach(function(q) {
      var kv = q.split('=');
      if (kv[0]) queryParams[decodeURIComponent(kv[0])] = decodeURIComponent(kv[1] || '');
    });
  }

  var allParams  = Object.assign({}, params, queryParams);
  var baseString = [
    method.toUpperCase(),
    encodeURIComponent(baseUrl),
    encodeURIComponent(
      Object.keys(allParams).sort()
        .map(function(k) { return k + '=' + encodeURIComponent(allParams[k]); })
        .join('&')
    ),
  ].join('&');

  var key = encodeURIComponent(CONSUMER_SECRET) + '&' +
            (tokenSecret ? encodeURIComponent(tokenSecret) : '');

  return Utilities.base64Encode(
    Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_1, baseString, key)
  );
}

/**
 * Parses a URL-encoded key=value&key=value string into an object.
 */
function parseOAuthParams_(str) {
  var out = {};
  str.split('&').forEach(function(part) {
    var kv = part.split('=');
    if (kv[0]) out[kv[0]] = kv[1] || '';
  });
  return out;
}
