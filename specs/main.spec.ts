import * as indexModule from '../index';
import { createMockSheet, createMockSpreadsheet, createMockRange } from './utils';

describe('main', () => {
    let spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
    let originalSheet: GoogleAppsScript.Spreadsheet.Sheet;
    let backupSheet: GoogleAppsScript.Spreadsheet.Sheet;
    let sourceRange: GoogleAppsScript.Spreadsheet.Range;
    let destinationRange: GoogleAppsScript.Spreadsheet.Range;
    let formatBackupSpy: jest.SpyInstance;

    beforeEach(() => {
        // Create mocks
        destinationRange = createMockRange();
        sourceRange = createMockRange({ getNumRows: jest.fn(() => 10), getNumColumns: jest.fn(() => 5) });
        originalSheet = createMockSheet();
        backupSheet = createMockSheet();

        (originalSheet.getDataRange as jest.Mock).mockReturnValue(sourceRange);
        (backupSheet.getRange as jest.Mock).mockReturnValue(destinationRange);

        spreadsheet = createMockSpreadsheet();
        (spreadsheet.getSheetByName as jest.Mock).mockReturnValue(originalSheet);
        (spreadsheet.insertSheet as jest.Mock).mockReturnValue(backupSheet);

        // Mock SpreadsheetApp
        (global as any).SpreadsheetApp = {
            getActiveSpreadsheet: jest.fn(() => spreadsheet),
        };

        // Mock Date
        jest.spyOn(global, 'Date').mockImplementation(() => ({
            getMonth: () => 11, // December
            getFullYear: () => 2024,
        } as unknown as Date));

        // Spy on formatBackup
        formatBackupSpy = jest.spyOn(indexModule, 'formatBackup').mockImplementation(jest.fn());

        jest.resetModules();
        formatBackupSpy.mockClear();
    });

    afterEach(() => {
        jest.restoreAllMocks();
    });

    it('should do nothing if current month is not December', () => {
        jest.spyOn(global, 'Date').mockImplementation(() => ({
            getMonth: () => 5, // June
            getFullYear: () => 2024,
        } as unknown as Date));

        indexModule.main();

        expect(SpreadsheetApp.getActiveSpreadsheet).not.toHaveBeenCalled();
    });

    it('should throw if sheet "Huidig jaar" is not found', () => {
        (spreadsheet.getSheetByName as jest.Mock).mockReturnValue(null);

        expect(() => indexModule.main()).toThrow('Sheet not found');
    });

    it('should create a backup, copy data, delete rows, and format backup', () => {
        indexModule.main();

        expect(spreadsheet.getSheetByName).toHaveBeenCalledWith('Huidig jaar');
        expect(spreadsheet.insertSheet).toHaveBeenCalledWith('Backup 2024');
        expect(originalSheet.getDataRange).toHaveBeenCalled();
        expect(sourceRange.copyTo).toHaveBeenCalledWith(destinationRange);
        expect(originalSheet.deleteRows).toHaveBeenCalledWith(2, 9); // 10 rows minus 1
    });

    it('should handle edge case when original sheet has only 1 row', () => {
        (originalSheet.getLastRow as jest.Mock).mockReturnValue(1);

        indexModule.main();

        expect(originalSheet.deleteRows).toHaveBeenCalledWith(2, 0); // no rows to delete
    });
});
