/* global console, require, Excel */

import { changeArgs, dataValidationCheck, FormConfig } from "../../type";
import { lockmainObjectValueToCell } from "../../utils";
import { getConfig, allTrue } from "../init";

export async function init(args: changeArgs, config: FormConfig) {
  if (!(args.type == "WorksheetChanged" || args.type == "TableChanged")) {
    return null;
  }
  console.log("on onchanged dataValidation...");

  await RangeEdited(args, config);
}

export async function check(obj: dataValidationCheck) {
  try {
    console.log(obj);
    var dv = require(`./${obj.onchangeconfig.type.toLowerCase()}`);
    console.log(dv);
    return await dv.check(obj);
  } catch (e) {
    console.error(e.stack);
    return true;
  }
}

async function isVailds(args: changeArgs, config: FormConfig) {
  try {
    //,  columnIndex: number,  rowIndex: number
    console.log("start Vaildtion");
    const colconf = await getConfig(args, config);
    console.log(colconf);
    // const colconf =
    //   rowIndex > config.table.rowIndex
    //     ? config.table.columns[columnIndex]
    //     : config.layout.find((ele) => ele.rowIndex === rowIndex && ele.index === columnIndex);
    const onChangedChecks = colconf.cell[args.type];
    const results = [];
    let result = true;
    if (onChangedChecks) {
      for (let onChangedCheck of onChangedChecks) {
        const type = onChangedCheck.type;
        var dv = require(`./${type.toLowerCase()}`);
        result = await dv.init(args, config);
        console.log("result " + type);
        console.log(result);
        if(type == "tabledata_unique" && args['details'] == null){
          result = true;
        }else{
          if(args['details'].valueAfter != ""){
            if (!result) {
              throw new Error("Error: Input text is NOT Vaild!");
            }
          }else{
            result = false;
            if(!colconf.mandatory){
              result = true;
            }
          }
        }

        results.push(result);
      }
    }
    if(results.length == 0){
      if(colconf.mandatory && args['details'].valueAfter == ""){
        result = false;
        results.push(result);
      }
    }
    console.log("end_result");
    console.log(results);
    console.log(result);
    return allTrue(results); //results.length > 1 ? allTrue(results) : results[0];
  } catch (e) {
    console.error(e.stack);
    return false;
  }
}

async function RangeEdited(args: changeArgs, config: FormConfig) {
  let sheet: Excel.Worksheet = null;

  let range: Excel.Range = null;
  let address = "";
  await Excel.run(async function (context) {
    sheet = context.workbook.worksheets.getActiveWorksheet();

    range = sheet.getRange(args.address);

    // let range = args.getRange(context);
    range.load(["address", "rowIndex", "columnIndex", "format/fill/color"]);
    await context.sync();
    address = range.address;
  });

  console.log(`on ${args.type}... address:${address}`);

  const colconf = await getConfig(args, config);
  console.log("args.type: " + args.type);
  console.log(config);
  console.log(colconf);
  // const onChangeds = colconf.cell[args.type];
  // console.log("onChangeds colconf");
  // console.log(onChangeds);
  const onChangedChecks = colconf.cell[args.type];
  try {
    if (onChangedChecks) {
      if (onChangedChecks.length < 1 && !colconf.mandatory) {
        throw new Error("nothing to check data vaildtion...");
      }
      let isVaild = await isVailds(args, config); //, range.columnIndex, range.rowIndex

      await lockmainObjectValueToCell(false);
      await Excel.run(async function (context) {
        try {
          let osheet = context.workbook.worksheets.getActiveWorksheet();

          let orange = osheet.getRange(address);

          if (isVaild) {
            console.log("orange.format.fill.clear");
            orange.format.fill.clear();
          } else {
            orange.format.fill.color = "#FF0000";
          }
          await context.sync();
        } catch (e) {
          console.error(e.stack);
        }
      });
      await lockmainObjectValueToCell(true);
    }
  } catch (e) {
    console.error(e.stack);
  }
}

export async function setEditedRange(args: changeArgs, isVaild: boolean, m?: string) {
  try {
    await lockmainObjectValueToCell(false);

    await Excel.run(async function (context) {
      try {
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        let range = sheet.getRange(args.address);

        console.log(`${m} isVaild: ${isVaild}`);
        if (isVaild) {
          range.format.fill.clear();
        } else {
          range.format.fill.color = "#FF0000";
        }
        context.sync().then(() => {
          lockmainObjectValueToCell(true);
        });
      } catch (e) {
        console.error(args);
        console.error(e.stack);
      }
    });
    // .catch(function (error) {
    //   Office.UI.notify(error);
    //   Office.Utilities.log(error);
    // });
  } catch (e) {
    console.error(args);
    console.error(e.stack);
  }
}
