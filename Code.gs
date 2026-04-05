/**
 * Code.gs — E*Trade Portfolio + Equity Curve — Google Apps Script
 * Version: 1.0 (2026-04-05)
 *
 * Entry point: onOpen() builds all menus.
 * Functionality is split across separate files:
 *   Authentication.gs   — OAuth 1.0a (request token → PIN → access token → renew)
 *   EtradeAPI.gs        — Low-level E*Trade API wrappers
 *   AccountRefresh.gs   — Portfolio refresh (getPortfolio)
 *   Liquidation.gs      — Position liquidation
 *   NetLiquidity.gs     — Net Liquidity sheet management
 *   SPYHistory.gs       — SPY + QQQ price history
 *   EquityCurveChart.gs — % change equity curve with month-end labels
 *   Backup.gs           — Backup to separate spreadsheet
 *   Triggers.gs         — Daily trigger setup / teardown
 */

// ─────────────────────────────────────────────────────────────────
// API + Auth constants  (fill in before deploying)
// ─────────────────────────────────────────────────────────────────
const BASE_URL       = 'https://api.etrade.com';
const ACCOUNTS_URL   = BASE_URL + '/v1/accounts';
const CONSUMER_KEY   = '<your-consumer-key>';
const CONSUMER_SECRET= '<your-consumer-secret>';

// ─────────────────────────────────────────────────────────────────
// Sheet names
// ─────────────────────────────────────────────────────────────────
const SHEET_ETRADE  = 'EtradeB';
const SHEET_NET_LIQ = 'Net Liquidity';
const SHEET_SPY     = 'SPY History';
const HISTORY_START = '2026-01-01';

// ─────────────────────────────────────────────────────────────────
// Account map  (fill in actual IDs and keys before deploying)
//
// Suffix  = last 4 digits of account ID (used as short identifier)
// id      = full numeric account ID
// key     = accountIdKey from E*Trade /v1/accounts/list API
// label   = display name shown in the portfolio sheet
// instType= 'BROKERAGE' or 'BROKERAGE_IRA' etc.
// ─────────────────────────────────────────────────────────────────
const ACCOUNT_MAP = {
  '7806': { id: '<account-id-ending-7806>', key: '<account-key-7806>', label: 'ETrade B',     instType: 'BROKERAGE'     },
  '8090': { id: '<account-id-ending-8090>', key: '<account-key-8090>', label: 'ETrade B IRA', instType: 'BROKERAGE'     },
  '3945': { id: '<account-id-ending-3945>', key: '<account-key-3945>', label: 'ETrade A',     instType: 'BROKERAGE'     },
};

// Column order in the Net Liquidity sheet (add/remove suffixes as needed)
const ACCOUNT_ORDER = ['7806', '8090', '3945'];

// ─────────────────────────────────────────────────────────────────
// Shared helpers used across files
// ─────────────────────────────────────────────────────────────────
function netLiqCol_(suffix)      { return ACCOUNT_ORDER.indexOf(suffix) + 2; }
function totalNetLiqCol_()       { return ACCOUNT_ORDER.length + 2; }
function chartSheetName_(suffix) { return 'Equity Curve ' + suffix; }

// ─────────────────────────────────────────────────────────────────
// Menu
// ─────────────────────────────────────────────────────────────────
function onOpen() {
  var ui = SpreadsheetApp.getUi();

  var authMenu = ui.createMenu('🔐 Authentication')
    .addItem('1 – Reset Auth (clear tokens)',   'resetAuth')
    .addItem('2 – Get Request Token',           'menuGetRequestToken')
    .addItem('3 – Set Verifier (PIN)',           'menuSetVerifier')
    .addItem('4 – Get Access Token',            'menuGetAccessToken')
    .addSeparator()
    .addItem('Renew Access Token',              'renewAccessToken');

  ui.createMenu('E*Trade')
    .addItem('📊 Refresh Portfolio', 'getPortfolio')
    .addSeparator()
    .addSubMenu(authMenu)
    .addToUi();

  ui.createMenu('📈 Equity Curve')
    .addItem('📊 Fetch SPY + QQQ History', 'fetchSPYHistory')
    .addSeparator()
    .addSubMenu(ui.createMenu('💼 Account …7806')
      .addItem("Capture Today's Net Liquidity (skip if exists)", 'fetchTodayNetLiq_7806')
      .addSeparator()
      .addItem('📈 Build / Refresh Equity Curve Chart', 'buildEquityCurveChart_7806')
      .addItem('📅 Build Chart – By Year',              'buildEquityCurveChartYearly_7806'))
    .addSubMenu(ui.createMenu('💼 Account …8090')
      .addItem("Capture Today's Net Liquidity (skip if exists)", 'fetchTodayNetLiq_8090')
      .addSeparator()
      .addItem('📈 Build / Refresh Equity Curve Chart', 'buildEquityCurveChart_8090')
      .addItem('📅 Build Chart – By Year',              'buildEquityCurveChartYearly_8090'))
    .addSubMenu(ui.createMenu('💼 Account …3945')
      .addItem("Capture Today's Net Liquidity (skip if exists)", 'fetchTodayNetLiq_3945')
      .addSeparator()
      .addItem('📈 Build / Refresh Equity Curve Chart', 'buildEquityCurveChart_3945')
      .addItem('📅 Build Chart – By Year',              'buildEquityCurveChartYearly_3945'))
    .addSeparator()
    .addSubMenu(ui.createMenu('⏰ Automation')
      .addItem('Enable Daily Snapshot – All Accounts (4:30 PM ET)', 'setupDailyTrigger')
      .addItem('Disable Daily Snapshot',                            'removeDailyTrigger_menu'))
    .addToUi();
}

// ─────────────────────────────────────────────────────────────────
// Per-account menu wrappers
// ─────────────────────────────────────────────────────────────────
function fetchTodayNetLiq_7806()              { fetchTodayNetLiqForAccount('7806'); }
function fetchTodayNetLiq_8090()              { fetchTodayNetLiqForAccount('8090'); }
function fetchTodayNetLiq_3945()              { fetchTodayNetLiqForAccount('3945'); }

function buildEquityCurveChart_7806()         { buildEquityCurveChartForAccount('7806'); }
function buildEquityCurveChart_8090()         { buildEquityCurveChartForAccount('8090'); }
function buildEquityCurveChart_3945()         { buildEquityCurveChartForAccount('3945'); }

function buildEquityCurveChartYearly_7806()   { promptAndBuildYearlyEquityCurve_('7806'); }
function buildEquityCurveChartYearly_8090()   { promptAndBuildYearlyEquityCurve_('8090'); }
function buildEquityCurveChartYearly_3945()   { promptAndBuildYearlyEquityCurve_('3945'); }

// ─────────────────────────────────────────────────────────────────
// Batch helpers
// ─────────────────────────────────────────────────────────────────

/** Daily trigger target — captures Net Liq for all accounts. */
function fetchTodayNetLiq() {
  ACCOUNT_ORDER.forEach(function(suffix) {
    try { fetchTodayNetLiqForAccount(suffix, true); }
    catch (e) { console.error('Daily snapshot failed for account ' + suffix + ': ' + e.message); }
  });
}

/** Refresh benchmarks + all snapshots + rebuild all charts. */
function refreshAllData() {
  fetchSPYHistory();
  ACCOUNT_ORDER.forEach(function(suffix) {
    fetchTodayNetLiqForAccount(suffix, true);
    buildEquityCurveChartForAccount(suffix);
  });
}
