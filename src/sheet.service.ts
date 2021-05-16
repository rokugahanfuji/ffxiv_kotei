import Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet;
import { Settings, Configs, ScheduleSheetInterface } from './settings';

export class SheetService {
  static getActiveSpreadSheet(): Spreadsheet {
    return SpreadsheetApp.getActiveSpreadsheet();
  }

  static createScheduleSheet(toDate: Date): GoogleAppsScript.Spreadsheet.Sheet | null {
    const sheet = SheetService.getActiveSpreadSheet();
    // The sheetname format is YYYY-MM
    const toName = toDate.getFullYear() + '-' + ('0' + toDate.getMonth() + 1).slice(-2);
    if (sheet.getSheetByName(toName) != null) {
      console.info(`${toName} is already exists.`);
      return null;
    }
    const from = Settings.BaseSheet.sheet;
    const to = from.copyTo(sheet);
    to.setName(toName);
    return to;
  }

  static getScheduleSheetInterface(
    sheet: GoogleAppsScript.Spreadsheet.Sheet
  ): ScheduleSheetInterface {
    const setting: ScheduleSheetInterface = {
      sheet: sheet,
      year: sheet.getRange(Configs.ScheduleSheet.yearCell),
      month: sheet.getRange(Configs.ScheduleSheet.monthCell),
    };
    return setting;
  }
}
