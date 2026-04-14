/* global console, Excel */

import { showMessage } from "../../../../helpers/message-helper";
import { dataValidationCheck, FormConfig } from "../../type";
// import { lockmainObjectValueToCell } from "../../utils";

// let allFalse = (arr) => arr.every((v) => v === false);

const m = "tabledata_depenfieldrequired";

export async function init(args: Excel.TableChangedEventArgs, config: FormConfig): Promise<boolean> {
  try {
    console.log(`on ${m}...`);

    console.log(args);
    if (args.details.valueAfter == "") {
      return true;
    }
    let range = null;
    let sheet = null;
    let expensesTable = null;
    let onchange_dataValidation = null;
    let valueAfters = args.details.valueAfter.toString();
    let value = await Excel.run(async function (context) {
      range = args.getRange(context);
      range.load(["rowIndex", "columnIndex", "address"]);
      sheet = context.workbook.worksheets.getActiveWorksheet();
      expensesTable = sheet.tables.getItem(config.table.name);
      expensesTable.load("columns");
      await context.sync();
      console.log(config);

      const colconf = config.table.columns[range.columnIndex];
      onchange_dataValidation = colconf.cell[args.type].find((e) => e.type.toLowerCase() === m.toLowerCase());
      console.log(onchange_dataValidation);
      if (!onchange_dataValidation) {
        throw new Error("Can't find onchange_dataValidation");
      }

      const tablefield = onchange_dataValidation.tablefield[0];

      const n_column = config.table.columns.find((e) => tablefield.toLowerCase() == e.LogicalName.toLowerCase());

      // set columns name
      var o_columns = expensesTable.columns;
      o_columns.load("items");
      console.log(o_columns.toJSON());
      await context.sync();
      const columns_index = n_column.index;
      let _range = sheet.getRangeByIndexes(range.rowIndex, columns_index, 1, 1);
      _range.load("values");
      await context.sync();
      return _range.values[0][0];
    });

    console.log("values");
    console.log(value);
    console.log("Vailding values");
    console.log(valueAfters);
    console.log(
      (value == onchange_dataValidation.targetvalue && valueAfters.toString().length > 0) ||
        !(value == onchange_dataValidation.targetvalue)
    );
    if (
      !(
        (value == onchange_dataValidation.targetvalue && valueAfters.toString().length > 0) ||
        !(value == onchange_dataValidation.targetvalue)
      ) &&
      onchange_dataValidation.errorAlert
    ) {
      showMessage({
        style: onchange_dataValidation.errorAlert.style,
        message: onchange_dataValidation.errorAlert.message,
        showAlert: onchange_dataValidation.errorAlert.showAlert ?? true,
      });
    }
    return (
      (value == onchange_dataValidation.targetvalue && valueAfters.toString().length > 0) ||
      !(value == onchange_dataValidation.targetvalue)
    );
  } catch (e) {
    console.error(e.stack);
  }
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
    console.log(config);

    const tablefield = onchangeconfig.tablefield[0];
    console.log(tablefield);

    const target_colconf = config.table.columns.find((e) => tablefield.toLowerCase() == e.LogicalName.toLowerCase());
    console.log(target_colconf);

    var n_column = expensesTable.columns.items.find((item) => target_colconf.Label.toLocaleLowerCase() == item.name.toLocaleLowerCase());
    const columns_index = n_column.index;
    let _range = null;
    await Excel.run(async function (context) {
      _range = context.workbook.worksheets.getActiveWorksheet().getRangeByIndexes(row, columns_index, 1, 1);
      _range.load("values");
      await context.sync();
    });

    var allColumn = o_columns.toJSON();
    var targetColumn = allColumn.items[columns_index];
    var targetValue = targetColumn.values[row + 1][0];
    console.log("o_columns");
    console.log(o_columns);
    console.log(columns_index);
    console.log(row + 1);
    console.log(allColumn);
    console.log(targetColumn);
    console.log(targetValue[0]);

    let tvalue = _range.values[0][0];
    if (!(targetValue == onchangeconfig.targetvalue 
      && value.toString().length > 0 
      || !(targetValue == onchangeconfig.targetvalue))
      && onchangeconfig.errorAlert) {
        var errorMsg = onchangeconfig.errorAlert.message.replace("{row}", (row + 1).toString());
        throw new Error(errorMsg);
    }
    return (
      (targetValue == onchangeconfig.targetvalue && value.toString().length > 0) || !(targetValue == onchangeconfig.targetvalue)
    );
  } catch (e) {
    global.ErrorMsg = e.message;
    console.error(e.stack);
    return false;
  }
}
