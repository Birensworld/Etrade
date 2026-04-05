/**
 * Backup.gs — Mirrors Net Liquidity and SPY/QQQ data to a backup spreadsheet.
 * Version: 1.0 (2026-04-05) — Ported from Schwab.
 *
 * Backup spreadsheet: "SPY_NetLiq_Backup"
 *   Tab "NetLiq_History"  ← mirrors "Net Liquidity" sheet
 *   Tab "SPY_QQQ_History" ← mirrors "SPY History" sheet
 */

var BACKUP_SS_NAME       = 'SPY_NetLiq_Backup';
var BACKUP_TAB_NETLIQ    = 'NetLiq_History';
var BACKUP_TAB_SPY       = 'SPY_QQQ_History';

function backupNetLiq_() {
  try {
    var src = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NET_LIQ);
    if (!src) return;
    copySheetToBackup_(src, BACKUP_TAB_NETLIQ);
    console.log('Backup: NetLiq_History updated.');
  } catch (e) {
    console.error('Backup (NetLiq) failed — ' + e.message);
  }
}

function backupSPYHistory_() {
  try {
    var src = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SPY);
    if (!src) return;
    copySheetToBackup_(src, BACKUP_TAB_SPY);
    console.log('Backup: SPY_QQQ_History updated.');
  } catch (e) {
    console.error('Backup (SPY/QQQ) failed — ' + e.message);
  }
}

function copySheetToBackup_(srcSheet, backupTabName) {
  var backupSS = getBackupSpreadsheet_();
  var dstSheet = backupSS.getSheetByName(backupTabName);
  if (!dstSheet) {
    throw new Error(
      'Tab "' + backupTabName + '" not found in "' + BACKUP_SS_NAME + '". ' +
      'Please create it manually.'
    );
  }
  var data = srcSheet.getDataRange().getValues();
  if (data.length === 0) return;
  var existingRows = dstSheet.getLastRow();
  if (existingRows > 0) dstSheet.clearContents();
  dstSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
}

function getBackupSpreadsheet_() {
  var files = DriveApp.getFilesByName(BACKUP_SS_NAME);
  if (!files.hasNext()) {
    throw new Error(
      'Backup spreadsheet "' + BACKUP_SS_NAME + '" not found in your Drive. ' +
      'Make sure the name matches exactly (case-sensitive).'
    );
  }
  return SpreadsheetApp.open(files.next());
}
