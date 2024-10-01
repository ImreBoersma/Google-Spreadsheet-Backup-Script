function main(): void {
  const currentMonth: number = new Date().getMonth();
  if (currentMonth !== 11) return;

  const spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const original: GoogleAppsScript.Spreadsheet.Sheet | null = spreadsheet.getSheetByName('Huidig jaar');
  
  if (!original) {
    throw new Error('Sheet not found');
  }

  const backup: GoogleAppsScript.Spreadsheet.Sheet = spreadsheet.insertSheet(`Backup ${new Date().getFullYear()}`);
  const sourceRange: GoogleAppsScript.Spreadsheet.Range = original.getDataRange();
  const destinationRange: GoogleAppsScript.Spreadsheet.Range = backup.getRange(1, 1, sourceRange.getNumRows(), sourceRange.getNumColumns());

  sourceRange.copyTo(destinationRange);
  original.deleteRows(2, original.getLastRow() - 1);

  formatBackup(backup);
}

function formatBackup(backup: GoogleAppsScript.Spreadsheet.Sheet): void {
  backup.deleteColumn(1);
  backup.autoResizeColumns(1, 5);
  backup.setFrozenRows(1);
  backup.setTabColor('#434343');

  const backupProtection: GoogleAppsScript.Spreadsheet.Protection = backup.protect();
  backupProtection.setDescription(`Dit is een backup van ${new Date().getFullYear()}. Deze sheet is alleen-lezen.`);
}
