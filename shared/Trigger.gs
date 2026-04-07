/**
 * Trigger.gs — Daily snapshot trigger setup and teardown.
 * Version: 1.0 (2026-04-05)
 *
 * setupDailyTrigger()       — Creates a time-driven trigger that calls
 *                             fetchTodayNetLiq() every day at ~4:30 PM in the
 *                             script's timezone (after market close). Removes
 *                             any existing trigger first to prevent duplicates.
 *
 * removeDailyTrigger_menu() — Deletes the daily trigger and confirms via toast.
 *
 * The trigger target is fetchTodayNetLiq() in NetLiquidity.gs, which captures
 * Net Liquidity for all accounts. Weekends/holidays are handled gracefully:
 * E*Trade returns 0 on non-trading days and NetLiquidity.gs ignores zero values.
 *
 * Triggers are identified by handler function name so they can be found and
 * deleted reliably without storing a trigger ID.
 */

var TRIGGER_FUNCTION = 'fetchTodayNetLiq';

// ─────────────────────────────────────────────────────────────────
// Public — called from menu
// ─────────────────────────────────────────────────────────────────

/**
 * Creates (or recreates) a daily ~4:30 PM trigger for fetchTodayNetLiq().
 * Any existing trigger for that function is removed first to avoid duplicates.
 */
function setupDailyTrigger() {
  removeDailyTrigger_();   // remove any existing to prevent duplicates

  ScriptApp.newTrigger(TRIGGER_FUNCTION)
    .timeBased()
    .everyDays(1)
    .atHour(16)            // 4 PM in the script's timezone
    .nearMinute(30)        // Apps Script runs within ~15 min of this — targets ~4:30 PM
    .create();

  var tz = Session.getScriptTimeZone();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '✅ Daily snapshot enabled — runs at ~4:30 PM (' + tz + ') every day.\n' +
    'Net Liquidity will be captured automatically for all accounts.',
    '⏰ Automation Active', 10
  );
  console.log('Daily trigger created for ' + TRIGGER_FUNCTION + ' at ~4:30 PM ' + tz);
}

/**
 * Removes the daily trigger. Called from the "Disable Daily Snapshot" menu item.
 */
function removeDailyTrigger_menu() {
  var removed = removeDailyTrigger_();
  if (removed) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Daily snapshot has been disabled. No further automatic captures will occur.',
      '⏰ Automation Disabled', 8
    );
    console.log('Daily trigger removed.');
  } else {
    SpreadsheetApp.getUi().alert(
      'No Active Trigger',
      'No daily snapshot trigger was found — it may have already been removed.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ─────────────────────────────────────────────────────────────────
// Private helpers
// ─────────────────────────────────────────────────────────────────

/**
 * Deletes all project triggers whose handler matches TRIGGER_FUNCTION.
 * @returns {boolean} true if at least one trigger was deleted
 */
function removeDailyTrigger_() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed  = false;

  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === TRIGGER_FUNCTION) {
      ScriptApp.deleteTrigger(trigger);
      removed = true;
    }
  });

  return removed;
}
