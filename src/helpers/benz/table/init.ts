/* global console, Excel */

import { changeArgs, FormConfig } from "../type";

export let allTrue = (arr) => arr.every((v) => v === true);
export let allFalse = (arr) => arr.every((v) => v === false);

export async function getConfig(args: changeArgs, config: FormConfig) {
  return await Excel.run(async function (context) {
    try {
      console.log(args);
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      let range = sheet.getRange(args.address);

      range.load(["address", "addressLocal", "rowIndex", "columnIndex"]);
      await context.sync();
      console.log(`on ${args.type}...`);
      console.log(range);

      if (args.type == "TableChanged") {
        return config.table.columns[range.columnIndex];
      } else if (args.type == "WorksheetChanged" || args.type == "WorksheetSingleClicked") {
        // return config.layout.find((ele) => ele.rowIndex === range.rowIndex && ele.index === range.columnIndex);
        return range.rowIndex < config.table.rowIndex
          ? config.layout.find(
              (ele) => ele.cell.address.toLocaleLowerCase() == range.address.split("!")[1].toLocaleLowerCase()
            )
          : config.table.columns[range.columnIndex];
      }
    } catch (e) {
      console.error(e.stack);
    }
  });
}
