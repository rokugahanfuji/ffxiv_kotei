// Configs
export const Configs = {
  BaseSheet: {
    sheetName: 'Base',
  },
  ScheduleSheet: {
    yearCell: 'E2',
    monthCell: 'G2',
  },
};

// Interfaces
export interface ScheduleSheetInterface {
  sheet: GoogleAppsScript.Spreadsheet.Sheet;
  year: GoogleAppsScript.Spreadsheet.Range;
  month: GoogleAppsScript.Spreadsheet.Range;
}

// Settings
export const Settings = {
  BaseSheet: (() => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Configs.BaseSheet.sheetName);
    const setting: ScheduleSheetInterface = {
      sheet: sheet,
      year: sheet.getRange(Configs.ScheduleSheet.yearCell),
      month: sheet.getRange(Configs.ScheduleSheet.monthCell),
    };
    return setting;
  })(),
};
