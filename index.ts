function main(): void {
  const now: Date = new Date();
  const month: number = now.getMonth();
  const year: number = now.getFullYear();

  // Only run in December
  if (month !== 11) return;

  const spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet | null = SpreadsheetApp.getActiveSpreadsheet();
  const original: GoogleAppsScript.Spreadsheet.Sheet | null = spreadsheet.getSheetByName('Huidig jaar');

  if (!original) {
    throw new Error('Sheet not found');
  }

  const backup: GoogleAppsScript.Spreadsheet.Sheet = spreadsheet.insertSheet(`Backup ${year}`);
  const sourceRange: GoogleAppsScript.Spreadsheet.Range = original.getDataRange();
  const destinationRange: GoogleAppsScript.Spreadsheet.Range = backup.getRange(1, 1, sourceRange.getNumRows(), sourceRange.getNumColumns());

  sourceRange.copyTo(destinationRange);
  original.deleteRows(2, original.getLastRow() - 1);

  formatBackup(backup, year);
}

function formatBackup(backup: GoogleAppsScript.Spreadsheet.Sheet, year: number): void {
  backup.deleteColumn(1);
  backup.autoResizeColumns(1, 5);
  backup.setFrozenRows(1);
  backup.setTabColor('#434343');

  const backupProtection: GoogleAppsScript.Spreadsheet.Protection = backup.protect();
  backupProtection.setDescription(`Dit is een backup van ${year}. Deze sheet is alleen-lezen.`);
}

export { main, formatBackup };