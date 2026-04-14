// eslint-disable-next-line @typescript-eslint/no-unused-vars
/* global global, console, Excel */

import { showMessage } from "../../../../helpers/message-helper";
import { dataValidationCheck, FormConfig } from "../../type";

const m = "tabledata_commissionnounique";

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
      expensesTable.load(["columns", "rows"]);
      console.log(args);
      let range = args.getRange(context);
      range.load(["rowIndex", "columnIndex", "address", "values"]);
      await context.sync();
      // const colconf = config.table.columns[range.columnIndex];
      const currentRow = range.rowIndex - config.table.rowIndex - 1;
      console.log("currentRow");
      console.log(currentRow);
      let column = expensesTable.columns.getItemAt(range.columnIndex);
      column.load(["values"]);
      await context.sync();
      const valueSet = column.values.map((v) => v[0]);
      console.log(`${m} column.values`);
      // console.log(column.values);
      console.log(column.values.map((v) => v[0]));
      // console.log(findDuplicates(valueSet));
      let isVaild = findDuplicates(currentRow, valueSet).length == 0;
      !isVaild
        ? showMessage({
            style: "error",
            message: `Duplicates Error: The '${findDuplicates(currentRow, valueSet)}' is duplicates !`,
          })
        : null;

      return isVaild;
    } catch (e) {
      console.error(e.stack);
      return null;
    }
  });
}

function findDuplicates(currentRow, array) {
  console.log("CurrentRow");
  console.log(currentRow);
  const duplicates = [];
  const measureNameArr = global.RowMeasureName;
  console.log(array);

  const measure_commission = [];
  for (let i = 1; i < array.length; i++) {
    var measureNum = i - 1;
    measure_commission.push(measureNameArr[measureNum] + array[i]);
  }
  console.log(measure_commission);

  const measure_commission_counts = {};
  measure_commission.forEach((x) => {
    measure_commission_counts[x] = (measure_commission_counts[x] || 0) + 1;
  });
  console.log("Commission Count");
  console.log(measure_commission_counts);

  const current_measure_commission = measureNameArr[currentRow] + array[currentRow + 1];
  if (measure_commission_counts[current_measure_commission] > 1) {
    console.log("commission duplicate", measureNameArr[currentRow], array[currentRow + 1]);
    duplicates.push(array[currentRow + 1]);
  }
  console.log("duplicates result", duplicates);

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
  let { onchangeconfig, value, fconfig, row} = object;
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
      const currentRow = row;
      let column = expensesTable.columns.getItemAt(fconfig.index);
      column.load(["values"]);
      await context.sync();

      const valueSet = column.values.map((v) => v[0]);
      console.log(`${m} column.values`);
      // console.log(column.values);
      console.log(column.values.map((v) => v[0]));
      // console.log(findDuplicates(currentRow, valueSet));
      let isVaild = findDuplicates(currentRow, valueSet).length == 0;
      if (!isVaild) {
        if (onchangeconfig.ErrorMsg != "") {
          var errorMsg = onchangeconfig.ErrorMsg.replace("{measureName}", global.RowMeasureName[currentRow]).replace(
            "{commissionNo}",
            value
          );
          throw new Error(errorMsg);
        }
      }

      return isVaild;
    } catch (e) {
      global.ErrorMsg = e.message;
      console.error(e.stack);
      return null;
    }
  });
}
