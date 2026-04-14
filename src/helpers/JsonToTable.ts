import { showMessage } from "./message-helper";

/* global  Excel,console */

export async function JsonToTable(sheetname: string, tablename: string, data: object[]): Promise<void> {
  try {
    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getItem(sheetname);
      if (!sheet) {
        throw `sheet:${sheetname} not exist...`;
      }
      // let expensesTable = sheet.tables.add("A1:D1", true /*hasHeaders*/); // create table
      // expensesTable.name = "ExpensesTable";
      let expensesTable = sheet.tables.getItem(tablename);

      let newData = data.map((item) => Object.values(item));

      expensesTable.rows.add(null, newData);
      sheet.getUsedRange().format.autofitColumns();
      sheet.getUsedRange().format.autofitRows();

      sheet.activate();

      await context.sync();
    });
  } catch (e) {
    console.log("EXCEPTION: " + e.message + ";" + e.stack);
    showMessage({
      style: "error",
      message: "EXCEPTION: " + e.message + ";",
    });
  }
}
