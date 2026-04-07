/**
 * Code.gs — EtradeB Sheet — E*Trade Portfolio + Equity Curve
 * Version: 1.4 (2026-04-07) — Split into EtradeB/EtradeA configs; remove account 3945
 *
 * Accounts: …7806 (ETrade B), …8090 (ETrade B IRA)
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
const SHEET_ETRADE  = 'EtradeB';
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
  '7806': { id: '<account-id-ending-7806>', key: '<account-key-7806>', label: 'ETrade B',     instType: 'BROKERAGE' },
  '8090': { id: '<account-id-ending-8090>', key: '<account-key-8090>', label: 'ETrade B IRA', instType: 'BROKERAGE' },
};

// Column order in the NL History sheet
const ACCOUNT_ORDER = ['7806', '8090'];

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

  ui.createMenu('E*Trade')
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
    .addSeparator()
    .addSubMenu(ui.createMenu('⏰ Automation')
      .addItem('Enable Daily Snapshot – All Accounts (4:30 PM ET)', 'setupDailyTrigger')
      .addItem('Disable Daily Snapshot',                            'removeDailyTrigger_menu'))
    .addToUi();
}

// ─────────────────────────────────────────────────────────────────
// Per-account menu wrappers
// ─────────────────────────────────────────────────────────────────
function fetchTodayNetLiq_7806()            { fetchTodayNetLiqForAccount('7806'); }
function fetchTodayNetLiq_8090()            { fetchTodayNetLiqForAccount('8090'); }

function buildEquityCurveChart_7806()       { buildEquityCurveChartForAccount('7806'); }
function buildEquityCurveChart_8090()       { buildEquityCurveChartForAccount('8090'); }

function buildEquityCurveChartYearly_7806() { promptAndBuildYearlyEquityCurve_('7806'); }
function buildEquityCurveChartYearly_8090() { promptAndBuildYearlyEquityCurve_('8090'); }

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
