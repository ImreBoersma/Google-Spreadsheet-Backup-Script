import { formatBackup } from '../index';
import { createMockSheet, createMockProtection } from './utils';

describe('formatBackup', () => {
  let backup: GoogleAppsScript.Spreadsheet.Sheet;

  beforeEach(() => {
    backup = createMockSheet();

    jest.spyOn(global, 'Date').mockImplementation(() => ({
      getFullYear: () => 2024,
    } as unknown as Date));

    jest.resetModules();
  });

  afterEach(() => {
    jest.restoreAllMocks();
  });

  it('should format the backup sheet correctly', () => {
    formatBackup(backup, 2024);

    expect(backup.deleteColumn).toHaveBeenCalledWith(1);
    expect(backup.autoResizeColumns).toHaveBeenCalledWith(1, 5);
    expect(backup.setFrozenRows).toHaveBeenCalledWith(1);
    expect(backup.setTabColor).toHaveBeenCalledWith('#434343');
    expect(backup.protect).toHaveBeenCalled();
    expect((backup.protect() as any).setDescription).toHaveBeenCalledWith(
      'Dit is een backup van 2024. Deze sheet is alleen-lezen.'
    );
  });

  it('should include the current year dynamically in description', () => {
    (Date as unknown as jest.Mock).mockImplementation(() => ({
      getFullYear: () => 2030,
    } as unknown as Date));

    formatBackup(backup, 2030);

    expect((backup.protect() as any).setDescription).toHaveBeenCalledWith(
      'Dit is een backup van 2030. Deze sheet is alleen-lezen.'
    );
  });

  it('should throw if backup.protect throws', () => {
    backup = createMockSheet({
      protect: jest.fn(() => {
        throw new Error('protect failed');
      }),
    });

    expect(() => formatBackup(backup, 2024)).toThrow('protect failed');
  });

  it('should throw if protect returns null', () => {
    backup = createMockSheet({
      protect: jest.fn(() => null as any),
    });

    expect(() => formatBackup(backup, 2024)).toThrow();
  });

  it('should throw if setDescription is missing', () => {
    backup = createMockSheet({}, { setDescription: undefined as any });

    expect(() => formatBackup(backup, 2024)).toThrow();
  });

  it('should call functions in the correct order', () => {
    const calls: string[] = [];
    backup = createMockSheet(
      {
        deleteColumn: jest.fn(function (this: GoogleAppsScript.Spreadsheet.Sheet) {
          calls.push('deleteColumn');
          return this;
        }),
        autoResizeColumns: jest.fn(function (this: GoogleAppsScript.Spreadsheet.Sheet) {
          calls.push('autoResizeColumns');
          return this;
        }),
        setFrozenRows: jest.fn(() => calls.push('setFrozenRows')),
        setTabColor: jest.fn(function (this: GoogleAppsScript.Spreadsheet.Sheet, _color: string) {
          calls.push('setTabColor');
          return this;
        }),
        protect: jest.fn(() => {
          calls.push('protect');
          return createMockProtection({
            setDescription: jest.fn(function (this: any, _desc: string) {
              calls.push('setDescription');
              return this;
            }),
          });
        }),
      }
    );

    formatBackup(backup, 2024);

    expect(calls).toEqual([
      'deleteColumn',
      'autoResizeColumns',
      'setFrozenRows',
      'setTabColor',
      'protect',
      'setDescription',
    ]);
  });

  it('should fail gracefully if a required method is missing', () => {
    backup = createMockSheet({
      deleteColumn: undefined as any,
    });

    expect(() => formatBackup(backup, 2024)).toThrow();
  });
});
