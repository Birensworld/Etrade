/**
 * Liquidation.gs — Closes all positions for a given E*Trade account.
 * Version: 1.0 (2026-04-05)
 *
 * Flow per position:
 *   1. Preview order  (E*Trade requires a preview before placing)
 *   2. Place order    (uses previewId returned in step 1)
 *
 * Results are logged to a "Trade Log" sheet.
 *
 * Depends on:
 *   Code.gs           — ACCOUNTS_URL, ACCOUNT_MAP, SHEET_ETRADE
 *   Authentication.gs — buildOAuthHeaders_()
 *   EtradeAPI.gs      — getAccessTokens_()
 */

// ─────────────────────────────────────────────────────────────────
// Menu entry points (one per account)
// ─────────────────────────────────────────────────────────────────

function close_EtradeB_7806() { closeAllEtradePositions(ACCOUNT_MAP['7806'].key); }
function close_EtradeB_8090() { closeAllEtradePositions(ACCOUNT_MAP['8090'].key); }
function close_EtradeA_3945() { closeAllEtradePositions(ACCOUNT_MAP['3945'].key); }

// ─────────────────────────────────────────────────────────────────
// Core liquidation
// ─────────────────────────────────────────────────────────────────

function closeAllEtradePositions(accountKey) {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var ui     = SpreadsheetApp.getUi();
  var sheet  = ss.getSheetByName(SHEET_ETRADE);
  if (!sheet) throw new Error('Sheet "' + SHEET_ETRADE + '" not found. Run Get Portfolio first.');

  var tokens = getAccessTokens_();
  var token  = tokens.token;
  var secret = tokens.secret;

  // Ensure / create trade log sheet
  var logSheet = ss.getSheetByName('Trade Log') || ss.insertSheet('Trade Log');
  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow(['Timestamp', 'Symbol', 'Qty', 'Action', 'Preview ID', 'Order Status', 'Response']);
  }

  var data = sheet.getDataRange().getValues();
  if (data.length < 3) throw new Error('No data found in "' + SHEET_ETRADE + '" tab.');

  var symbolCol = -1, qtyCol = -1;
  var KEY_COL   = 9;  // Column J holds accountIdKey (0-based index = 9)

  for (var i = 0; i < data.length; i++) {
    var row       = data[i].map(function(v) { return String(v).trim(); });
    var firstCell = row[0].toLowerCase();

    // Detect the positions header row
    if (firstCell.indexOf('symbol') !== -1) {
      symbolCol = row.findIndex(function(c) { return c.toLowerCase().indexOf('symbol') !== -1; });
      qtyCol    = row.findIndex(function(c) { return c.toLowerCase().indexOf('qty')    !== -1; });
      continue;
    }
    if (symbolCol === -1 || qtyCol === -1) continue;

    // Only process rows belonging to this account
    if (row[KEY_COL] !== accountKey) continue;

    var symbol = row[symbolCol];
    var qty    = Number(row[qtyCol].replace(/,/g, ''));
    if (!symbol || qty <= 0) continue;

    var action = 'SELL';   // change to 'BUY' for testing

    // ── Preview ────────────────────────────────────────────────
    var previewUrl  = ACCOUNTS_URL + '/' + accountKey + '/orders/preview.json';
    var prevHeaders = mergeHeaders_(
      buildOAuthHeaders_(previewUrl, 'POST', token, secret),
      { 'Content-Type': 'application/xml', Accept: 'application/json' }
    );
    var xmlPreview  = buildOrderXml_('PreviewOrderRequest', null, symbol, qty, action);
    var prevResp    = UrlFetchApp.fetch(previewUrl, {
      method: 'post', headers: prevHeaders, muteHttpExceptions: true, payload: xmlPreview,
    });

    if (prevResp.getResponseCode() !== 200) {
      logTrade_(logSheet, symbol, qty, action, '', 'Preview Failed', prevResp.getContentText());
      continue;
    }

    var previewId;
    try {
      var pJson = JSON.parse(prevResp.getContentText());
      previewId = pJson.PreviewOrderResponse &&
                  pJson.PreviewOrderResponse.PreviewIds &&
                  pJson.PreviewOrderResponse.PreviewIds[0] &&
                  pJson.PreviewOrderResponse.PreviewIds[0].previewId;
    } catch (e) {
      logTrade_(logSheet, symbol, qty, action, '', 'Preview JSON Error', prevResp.getContentText());
      continue;
    }

    if (!previewId) {
      logTrade_(logSheet, symbol, qty, action, '', 'No Preview ID', prevResp.getContentText());
      continue;
    }

    // ── Place ──────────────────────────────────────────────────
    var placeUrl     = ACCOUNTS_URL + '/' + accountKey + '/orders/place.json';
    var placeHeaders = mergeHeaders_(
      buildOAuthHeaders_(placeUrl, 'POST', token, secret),
      { 'Content-Type': 'application/xml', Accept: 'application/json' }
    );
    var xmlPlace = buildOrderXml_('PlaceOrderRequest', previewId, symbol, qty, action);
    var placeResp = UrlFetchApp.fetch(placeUrl, {
      method: 'post', headers: placeHeaders, muteHttpExceptions: true, payload: xmlPlace,
    });

    var placeCode = placeResp.getResponseCode();
    var status    = (placeCode === 200 || placeCode === 201) ? 'Order Placed' : 'Error ' + placeCode;
    logTrade_(logSheet, symbol, qty, action, previewId, status, placeResp.getContentText());
  }

  var msg = 'Completed transactions for all positions in account (key: …' +
            accountKey.slice(-6) + ').';
  ui.alert(msg);
  ss.toast(msg);
}

// ─────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────

/**
 * Builds the XML payload for a preview or place order request.
 * @param {string}  rootTag    'PreviewOrderRequest' or 'PlaceOrderRequest'
 * @param {string=} previewId  Required for PlaceOrderRequest
 */
function buildOrderXml_(rootTag, previewId, symbol, qty, action) {
  var previewBlock = previewId
    ? '<PreviewIds><previewId>' + previewId + '</previewId></PreviewIds>'
    : '';
  var clientId = 'BM' + Date.now();

  return (
    '<' + rootTag + '>' +
    '<orderType>EQ</orderType>' +
    '<clientOrderId>' + clientId + '</clientOrderId>' +
    previewBlock +
    '<Order>' +
    '<orderTerm>GOOD_FOR_DAY</orderTerm>' +
    '<priceType>MARKET</priceType>' +
    '<marketSession>REGULAR</marketSession>' +
    '<allOrNone>false</allOrNone>' +
    '<Instrument>' +
    '<Product>' +
    '<symbol>' + symbol + '</symbol>' +
    '<securityType>EQ</securityType>' +
    '</Product>' +
    '<symbolDescription>' + symbol + '</symbolDescription>' +
    '<orderAction>' + action + '</orderAction>' +
    '<quantityType>QUANTITY</quantityType>' +
    '<quantity>' + qty + '</quantity>' +
    '</Instrument>' +
    '</Order>' +
    '</' + rootTag + '>'
  );
}

/** Appends one row to the Trade Log sheet. */
function logTrade_(sheet, symbol, qty, action, previewId, status, response) {
  sheet.appendRow([
    new Date(),
    symbol,
    qty,
    action,
    previewId || '',
    status,
    response || '',
  ]);
}
