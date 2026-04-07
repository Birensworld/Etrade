/**
 * Code.gs — EtradeA Sheet — E*Trade Portfolio + Equity Curve
 * Version: 1.2 (2026-04-07) — Enhance debugNetLiqReturn to check live function param count
 *
 * Accounts: …3945 (ETrade A IRA)
 *
 * Deploy to GAS project by copying all files from shared/ + this Code.gs.
 *
 * Shared files (shared/):
 *   Authentication.gs   — OAuth 1.0a flow
 *   EtradeAPI.gs        — Low-level E*Trade API wrappers
 *   AccountRefresh.gs   — Portfolio refresh
 *   Liquidation.gs      — Position liquidation
 *   NetLiquidity.gs     — Net Liquidity sheet management
 *   SPYHistory.gs       — SPY + QQQ price history
 *   EquityCurveChart.gs — % change equity curve chart
 *   Backup.gs           — Backup to separate spreadsheet
 *   Trigger.gs          — Daily trigger setup / teardown
 */

// ─────────────────────────────────────────────────────────────────
// API + Auth constants  (fill in before deploying — do NOT commit real values)
// ─────────────────────────────────────────────────────────────────
const BASE_URL        = 'https://api.etrade.com';
const ACCOUNTS_URL    = BASE_URL + '/v1/accounts';
const CONSUMER_KEY    = '<your-consumer-key>';
const CONSUMER_SECRET = '<your-consumer-secret>';

// ─────────────────────────────────────────────────────────────────
// Sheet names
// ─────────────────────────────────────────────────────────────────
const SHEET_ETRADE  = 'EtradeA';
const SHEET_NET_LIQ = 'NL History';
const SHEET_SPY     = 'SPY History';
const HISTORY_START = '2026-01-01';

// ─────────────────────────────────────────────────────────────────
// Account map
//
// Suffix  = last 4 digits of account ID (short identifier)
// id      = full numeric account ID
// key     = accountIdKey from E*Trade /v1/accounts/list API
// label   = display name in portfolio sheet
// instType= 'BROKERAGE' or 'BROKERAGE_IRA'
// ─────────────────────────────────────────────────────────────────
const ACCOUNT_MAP = {
  '3945': { id: '<account-id-ending-3945>', key: '<account-key-3945>', label: 'ETrade A IRA', instType: 'BROKERAGE' },
};

// Column order in the NL History sheet
const ACCOUNT_ORDER = ['3945'];

// ─────────────────────────────────────────────────────────────────
// Shared helpers
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
    .addItem('1 – Reset Auth (clear tokens)', 'resetAuth')
    .addItem('2 – Get Request Token',         'menuGetRequestToken')
    .addItem('3 – Set Verifier (PIN)',         'menuSetVerifier')
    .addItem('4 – Get Access Token',          'menuGetAccessToken')
    .addSeparator()
    .addItem('Renew Access Token',            'renewAccessToken');

  ui.createMenu('E*Trade Arch')
    .addItem('📊 Refresh Portfolio', 'getPortfolio')
    .addSeparator()
    .addSubMenu(authMenu)
    .addSeparator()
    .addItem('🔑 Show Account Keys (for Code.gs setup)', 'showAccountKeys')
    .addToUi();

  ui.createMenu('📈 Equity Curve')
    .addItem('📊 Fetch SPY + QQQ History',               'fetchSPYHistory')
    .addItem('📸 Capture Today\'s Net Liq – All Accounts', 'captureAllNetLiq')
    .addSeparator()
    .addSubMenu(ui.createMenu('💼 Account …3945')
      .addItem("Capture Today's Net Liquidity (skip if exists)", 'fetchTodayNetLiq_3945')
      .addSeparator()
      .addItem('📈 Build / Refresh Equity Curve Chart', 'buildEquityCurveChart_3945')
      .addItem('📅 Build Chart – By Year',              'buildEquityCurveChartYearly_3945'))
    .addSeparator()
    .addSubMenu(ui.createMenu('⏰ Automation')
      .addItem('Enable Daily Snapshot (4:30 PM ET)', 'setupDailyTrigger')
      .addItem('Disable Daily Snapshot',             'removeDailyTrigger_menu'))
    .addToUi();
}

// ─────────────────────────────────────────────────────────────────
// Per-account menu wrappers
// ─────────────────────────────────────────────────────────────────
function fetchTodayNetLiq_3945()            { fetchTodayNetLiqForAccount('3945'); }
function buildEquityCurveChart_3945()       { buildEquityCurveChartForAccount('3945'); }
function buildEquityCurveChartYearly_3945() { promptAndBuildYearlyEquityCurve_('3945'); }

// ─────────────────────────────────────────────────────────────────
// Batch helpers
// ─────────────────────────────────────────────────────────────────

/**
 * Captures today's Net Liquidity for all accounts in one click.
 * Runs silently per-account then shows a combined summary toast.
 */
function captureAllNetLiq() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Capturing Net Liquidity for all accounts…', 'Working', -1);

  var results = ACCOUNT_ORDER.map(function(suffix) {
    return fetchTodayNetLiqForAccount(suffix, true, true);  // skipIfExists=true, silent=true
  });

  var lines  = [];
  var errors = [];

  results.forEach(function(r, i) {
    if (!r || typeof r.status === 'undefined') {
      var s = ACCOUNT_ORDER[i] || '?';
      lines.push('❌ …' + s + ':  no result — ensure NetLiquidity.gs is up to date');
      errors.push('Account …' + s + ':\nNetLiquidity.gs may not be up to date. Re-deploy both files.');
      return;
    }
    if (r.status === 'captured') {
      lines.push('✅ …' + r.suffix + ':  $' +
        r.value.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }));
    } else if (r.status === 'skipped') {
      lines.push('⏭️ …' + r.suffix + ':  already captured for ' + r.dateStr);
    } else {
      lines.push('❌ …' + r.suffix + ':  error');
      errors.push('Account …' + r.suffix + ':\n' + r.message);
    }
  });

  ss.toast(lines.join('\n'), '📸 Net Liquidity – All Accounts', 15);

  if (errors.length > 0) {
    SpreadsheetApp.getUi().alert(
      'Errors capturing Net Liquidity',
      errors.join('\n\n'),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ─────────────────────────────────────────────────────────────────
// Debug helper — remove after diagnosing
// ─────────────────────────────────────────────────────────────────

/**
 * Run this directly from the GAS editor (Run → debugNetLiqReturn).
 * Shows exactly what fetchTodayNetLiqForAccount returns so we can
 * see why captureAllNetLiq's null-check is firing.
 */
function debugNetLiqReturn() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Step 1: Check how many params the LIVE function has.
  // Old (void) version = 2 params. New (returns object) version = 3 params.
  var paramCount = fetchTodayNetLiqForAccount.length;
  ss.toast(
    'Live param count : ' + paramCount + '  (expect 3)\n' +
    'If 2 → GAS is running cached old version',
    '🔍 Step 1 – Param count', 10
  );
  console.log('fetchTodayNetLiqForAccount.length = ' + paramCount);
  Utilities.sleep(10000);

  // Step 2: Call and capture result
  var result = fetchTodayNetLiqForAccount('3945', true, true);
  ss.toast(
    'typeof result : ' + (typeof result) + '\n' +
    'result        : ' + JSON.stringify(result),
    '🔍 Step 2 – Return value', 30
  );
  console.log('debugNetLiqReturn → ' + JSON.stringify(result));
}

/** Daily trigger target — captures Net Liq for all accounts. */
function fetchTodayNetLiq() {
  ACCOUNT_ORDER.forEach(function(suffix) {
    try { fetchTodayNetLiqForAccount(suffix, true); }
    catch (e) { console.error('Daily snapshot failed for …' + suffix + ': ' + e.message); }
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
