/* global console, Excel */

import { showMessage } from "../../../../helpers/message-helper";
import { dataValidationCheck, FormConfig } from "../../type";
// import { lockmainObjectValueToCell } from "../../utils";

let allFalse = (arr) => arr.every((v) => v === false);

const m = "tabledata_dependency";

export async function init(args: Excel.TableChangedEventArgs, config: FormConfig): Promise<boolean> {
  console.log(`on ${m}...`);
  console.log(args);
  if (args.details.valueAfter == "") {
    return true;
  }
  return await Excel.run(async function (context) {
    try {

      let range = args.getRange(context);
      range.load(["rowIndex", "columnIndex", "address"]);
      var sheet = context.workbook.worksheets.getActiveWorksheet();
      let expensesTable = sheet.tables.getItem(config.table.name);
      expensesTable.load("columns");
      await context.sync();

      console.log(config);

      const colconf = config.table.columns[range.columnIndex];
      const onchange_dataValidation = colconf.cell[args.type].find((e) => e.type.toLowerCase() === m.toLowerCase());
      console.log(onchange_dataValidation);
      if (!onchange_dataValidation) {
        throw new Error("Can't find onchange_dataValidation");
      }

      const valueAfters = args.details.valueAfter
        .toString()
        .split(onchange_dataValidation.separator ?? ",")
        .map((e) => e.trim());

      const tablefield = onchange_dataValidation.tablefield;

      const taget_colconf = config.table.columns.filter(
        (e) =>
          tablefield.map((e) => e.split(".")[0]).includes(e.EntityLogicalName.toLowerCase()) &&
          tablefield.map((e) => e.split(".")[1]).includes(e.LogicalName.toLowerCase())
      );
      // set columns name
      var o_columns = expensesTable.columns;
      o_columns.load("items");
      console.log(o_columns.toJSON());
      await context.sync();

      var n_columns = expensesTable.columns.items.filter((item) =>
        taget_colconf.map((e) => e.Label).includes(item.name)
      );
      let isVaild = false;
      console.log(n_columns);
      let values = [];
      for (let n_column of n_columns) {
        const columns_index = n_column.index;
        let _range = sheet.getRangeByIndexes(range.rowIndex, columns_index, 1, 1);
        _range.load("values");
        await context.sync();
        let value = _range.values[0][0];
        console.log("values");
        console.log(value.split(onchange_dataValidation.separator ?? ","));
        values = values.concat(value.split(onchange_dataValidation.separator ?? ","));
      }
      console.log("Vailding values");
      console.log(values);
      console.log(valueAfters);

      isVaild = allFalse(values.map((e) => valueAfters.includes(e.trim())));

      return isVaild;
    } catch (e) {
      showMessage({ style: "error", message: e.message });
      console.error(e.stack);
    }
  });
}

export async function check(object: dataValidationCheck) {
  try {
    let { config, onchangeconfig, value, row } = object;
    console.log(`on ${m}...`);
    console.log(onchangeconfig);
    console.log(value);
    let expensesTable = null;
    let o_columns = null;
    await Excel.run(async function (context) {
      var sheet = context.workbook.worksheets.getActiveWorksheet();
      expensesTable = sheet.tables.getItemAt(0);
      expensesTable.load("columns");
      // set columns name
      await context.sync();
      o_columns = expensesTable.columns;
      o_columns.load("items");
      console.log(o_columns.toJSON());
      await context.sync();
    });

    if (!onchangeconfig) {
      throw new Error("Can't find conf");
    }
    if (!value) {
      return true;
    }
    const valueAfters = value.toString().split(onchangeconfig.separator ?? ",");
    const tablefield = onchangeconfig.tablefield;

    const taget_colconf = config.table.columns.filter(
      (e) =>
        tablefield.map((e) => e.split(".")[0]).includes(e.EntityLogicalName.toLowerCase()) &&
        tablefield.map((e) => e.split(".")[1]).includes(e.LogicalName.toLowerCase())
    );

    var n_columns = expensesTable.columns.items.filter((item) => taget_colconf.map((e) => e.Label).includes(item.name));
    let isVaild = false;
    console.log(n_columns);
    let values = [];
    var rowIndex = config.table.rowIndex + 1;
    for (let n_column of n_columns) {
      rowIndex = rowIndex + row;
      const columns_index = n_column.index;
      let _range = null;
      await Excel.run(async function (context) {
        _range = context.workbook.worksheets.getActiveWorksheet().getRangeByIndexes(rowIndex, columns_index, 1, 1);
        _range.load("values");
        await context.sync();
      });

      let value = _range.values[0][0];
      console.log("values");
      console.log(value.split(onchangeconfig.separator ?? ","));
      values = values.concat(value.split(onchangeconfig.separator ?? ","));
    }
    console.log("Vailding values");
    console.log(values);
    console.log(valueAfters);

    isVaild = allFalse(values.map((e) => valueAfters.includes(e.trim())));
    return isVaild;
  } catch (e) {
    console.error(e.stack);
    return false;
  }
}
