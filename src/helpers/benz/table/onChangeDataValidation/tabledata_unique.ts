// eslint-disable-next-line @typescript-eslint/no-unused-vars
/* global global, console, Excel */

import { showMessage } from "../../../../helpers/message-helper";
import { dataValidationCheck, FormConfig } from "../../type";

const m = "tabledata_unique";

export async function init(args: Excel.TableChangedEventArgs, config: FormConfig): Promise<boolean> {
  console.log(`on ${m}...`);
  console.log(args);
  if (args.details == null) {
    return true;
  }
  if (args.details.valueAfter == "") {
    return true;
  }
  return await Excel.run(async function (context) {
    try {
      var sheet = context.workbook.worksheets.getActiveWorksheet();
      let expensesTable = sheet.tables.getItem(config.table.name);
      expensesTable.load("columns");
      console.log(args);
      let range = args.getRange(context);
      range.load(["columnIndex", "address", "values"]);
      await context.sync();
      // const colconf = config.table.columns[range.columnIndex];
      let column = expensesTable.columns.getItemAt(range.columnIndex);
      column.load(["values"]);
      await context.sync();
      const valueSet = column.values.map((v) => v[0]);
      console.log(`${m} column.values`);
      console.log(column.values);
      console.log(column.values.map((v) => v[0]));
      console.log(findDuplicates(valueSet));
      let isVaild = findDuplicates(valueSet).length == 0;
      !isVaild
        ? showMessage({
            style: "error",
            message: `Duplicates Error: The '${findDuplicates(valueSet)}' is duplicates !`,
          })
        : null;

      return isVaild;
    } catch (e) {
      console.error(e.stack);
      return null;
    }
  });
}

function findDuplicates(array) {
  const duplicates = [];

  for (let i = 0; i < array.length; i++) {
    for (let j = i + 1; j < array.length; j++) {
      if (array[i] === array[j] && !duplicates.includes(array[i])) {
        duplicates.push(array[i]);
      }
    }
  }

  return duplicates;
}

function vaild(array: any[]) {
  var valuesSoFar = Object.create(null);
  for (var i = 0; i < array.length; ++i) {
    var value = array[i];
    if (value in valuesSoFar) {
      return true;
    }
    valuesSoFar[value] = true;
  }
  return false;
}
export async function check(object: dataValidationCheck) {
  let { onchangeconfig, value, fconfig } = object;
  return await Excel.run(async function (context) {
    try {
      console.log(`on ${m}...`);
      console.log(onchangeconfig);
      console.log(value);

      if (!onchangeconfig || !value) {
        throw new Error("Can't find conf");
      }

      console.log(`on ${m}...`);

      var sheet = context.workbook.worksheets.getActiveWorksheet();
      let expensesTable = sheet.tables.getItemAt(0);
      expensesTable.load("columns");
      await context.sync();
      // const colconf = config.table.columns[range.columnIndex];
      let column = expensesTable.columns.getItemAt(fconfig.index);
      column.load(["values"]);
      await context.sync();

      const valueSet = column.values.map((v) => v[0]);
      console.log(`${m} column.values`);
      console.log(column.values);
      console.log(column.values.map((v) => v[0]));
      console.log(findDuplicates(valueSet));
      let isVaild = findDuplicates(valueSet).length == 0;
      !isVaild
        ? showMessage({
            style: "error",
            message: `Duplicates Error: The '${findDuplicates(valueSet)}' is duplicates !`,
          })
        : null;

      return isVaild;
    } catch (e) {
      console.error(e.stack);
      return null;
    }
  });
}
