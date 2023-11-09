function main() {
  if (new Date().getMonth() !== 11) return;

  const original: GoogleAppsScript.Spreadsheet.Sheet | null = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Huidig jaar');
  const backup: GoogleAppsScript.Spreadsheet.Sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(`Backup ${new Date().getFullYear()}`);

  if (!original) {
    throw new Error('Sheet not found');
  }

  const source: GoogleAppsScript.Spreadsheet.Range = original.getRange(1, 1, original.getLastRow(), original.getLastColumn());
  const destination: GoogleAppsScript.Spreadsheet.Range = backup.getRange(1, 1, source.getNumRows(), source.getNumColumns());

  source.copyTo(destination);

  original.deleteRows(2, original.getLastRow());

  formatBackup(backup);
}

function formatBackup(backup: GoogleAppsScript.Spreadsheet.Sheet) {
  backup.deleteColumn(1);
  backup.autoResizeColumns(1, 5);
  backup.setFrozenRows(1);
  backup.setTabColor('#434343');

  const backupProtection: GoogleAppsScript.Spreadsheet.Protection = backup.protect();

  backupProtection.setDescription(`Dit is een backup van ${new Date().getFullYear()}. Deze sheet is alleen-lezen.`);
}