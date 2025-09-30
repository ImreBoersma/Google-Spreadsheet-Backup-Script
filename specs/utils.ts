export function createMockProtection(
  overrides: Partial<GoogleAppsScript.Spreadsheet.Protection> = {}
): GoogleAppsScript.Spreadsheet.Protection {
  return {
    setDescription: jest.fn(),
    ...overrides,
  } as unknown as GoogleAppsScript.Spreadsheet.Protection;
}

export function createMockRange(
  overrides: Partial<GoogleAppsScript.Spreadsheet.Range> = {}
): GoogleAppsScript.Spreadsheet.Range {
  return {
    copyTo: jest.fn(),
    getNumRows: jest.fn(() => 10),
    getNumColumns: jest.fn(() => 5),
    ...overrides,
  } as unknown as GoogleAppsScript.Spreadsheet.Range;
}

export function createMockSheet(
  overrides: Partial<GoogleAppsScript.Spreadsheet.Sheet> = {},
  protectionOverrides: Partial<GoogleAppsScript.Spreadsheet.Protection> = {}
): GoogleAppsScript.Spreadsheet.Sheet {
  const protection = createMockProtection(protectionOverrides);
  const range = createMockRange();

  return {
    deleteColumn: jest.fn(),
    autoResizeColumns: jest.fn(),
    setFrozenRows: jest.fn(),
    setTabColor: jest.fn(),
    protect: jest.fn(() => protection),
    getDataRange: jest.fn(() => range),
    getLastRow: jest.fn(() => 10),
    getRange: jest.fn(() => createMockRange()),
    deleteRows: jest.fn(),
    ...overrides,
  } as unknown as GoogleAppsScript.Spreadsheet.Sheet;
}

export function createMockSpreadsheet(
  overrides: Partial<GoogleAppsScript.Spreadsheet.Spreadsheet> = {}
): GoogleAppsScript.Spreadsheet.Spreadsheet {
  return {
    getSheetByName: jest.fn(),
    insertSheet: jest.fn(),
    ...overrides,
  } as unknown as GoogleAppsScript.Spreadsheet.Spreadsheet;
}
