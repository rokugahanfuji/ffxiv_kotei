import { SheetService } from './sheet.service';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
declare let global: any;

global.createSheetOfThisMonth = (): void => {
  const date = new Date();
  const sheet = SheetService.createScheduleSheet(date);
  if (sheet != null) {
    const accessor = SheetService.getScheduleSheetInterface(sheet);
    accessor.year.setValue(date.getFullYear());
    accessor.month.setValue(date.getMonth() + 1);
    console.info(`${sheet.getName()} created successfully.`);
  }
};

global.createSheetOfNextMonth = (): void => {
  const date = new Date();
  date.setMonth(date.getMonth() + 1);
  const sheet = SheetService.createScheduleSheet(date);
  if (sheet != null) {
    const accessor = SheetService.getScheduleSheetInterface(sheet);
    accessor.year.setValue(date.getFullYear());
    accessor.month.setValue(date.getMonth());
    console.info(`${sheet.getName()} created successfully.`);
  }
};
