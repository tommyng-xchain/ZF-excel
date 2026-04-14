/* global global, console, Excel */

import { dataValidationCheck, FormConfig } from "../../type";

const m = "tabledata_checkvaluesetvalue";

export const compare = {
  "==": (a: any, b: any) => a == b,
  "!=": (a: any, b: any) => a != b,
  "===": (a: any, b: any) => a === b,
  "!==": (a: any, b: any) => a !== b,
};

export async function init(args: Excel.TableChangedEventArgs, config: FormConfig): Promise<boolean> {
  console.log(`on ${m}...`);
  console.log(args);

  if (args.details.valueAfter == "") {
    return true;
  }
  await Excel.run(async function (context) {
    try {
      let range = args.getRange(context);
      range.load(["rowIndex", "columnIndex", "address"]);
      await context.sync();
      console.log(config);
      let rowIndex = range.rowIndex;
      const colconf = config.table.columns[range.columnIndex];
      const dv_conf = colconf.cell[args.type].find((e) => e.type.toLowerCase() === m.toLowerCase());
      if (!dv_conf) {
        return;
      }
      const valueAfters = args.details.valueAfter.toString().split(dv_conf.separator ?? ",");
      global.onChangeDataValidation = { args: args, config: config, columnIndex: range.columnIndex, type: m };
      if (valueAfters) {
        let _continue = dv_conf.continue;
        let parameters1 = _continue.parameters1.replace("{valueAfters}", valueAfters);
        let parameters2 = _continue.parameters2.replace("{valueAfters}", valueAfters);
        if (compare[_continue.operator](parameters1, parameters2)) {
          const target_colconf = config.table.columns.find((e) => e.LogicalName === _continue.targetLogicalName);
          let sheet = context.workbook.worksheets.getActiveWorksheet();
          let target_range = sheet.getRangeByIndexes(rowIndex, target_colconf.index, 1, 1);
          target_range.values = [[_continue.value]];
          await context.sync();
        }
      } else {
        console.log(`on ${m} empty...`);
      }
    } catch (e) {
      console.error(e.stack);
    }
  });
  return true;
}

export async function check(obj: dataValidationCheck) {
  return true;
}
