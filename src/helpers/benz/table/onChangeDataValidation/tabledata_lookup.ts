/* global global, console, Excel */

import { FormConfig } from "../../type";
import { setEditedRange } from "./init";
import { RetrieveMultipleData } from "../../../dataverse-data-helper";
import { set_layout } from "../../utils";
import * as lookup from "../../../systemuserlookupdialog";
const m = "tabledata_lookup";

export async function init(args: Excel.WorksheetSingleClickedEventArgs, config: FormConfig): Promise<boolean> {
  await Excel.run(async function (context) {
    const m = "tabledata_lookup";
    try {
      console.log(`on ${m}...`);
      console.log(args);
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      let range = sheet.getRange(args.address);

      range.load(["rowIndex", "columnIndex", "address"]);
      await context.sync();
      console.log(config);

      const colconf =
        range.rowIndex > config.table.rowIndex
          ? config.table.columns[range.columnIndex]
          : config.layout.find((ele) => ele.rowIndex === range.rowIndex && ele.index === range.columnIndex);
      const dv_conf = colconf.cell[args.type].find((e) => e.type.toLowerCase() === m.toLowerCase());
      if (!dv_conf) {
        return;
      }
      // lookup.showdialog();

      // const valueAfters = args.details.valueAfter.toString().split(dv_conf.separator ?? ",");
      // global.onChangeDataValidation = { args: args, config: config, columnIndex: range.columnIndex, type: m };
      // if (valueAfters) {
      //   getdata(
      //     args,
      //     dv_conf.EntityLogicalName,
      //     "?" + dv_conf.queryString.replace("{valueAfters}", valueAfters[0]) + "&$count=true",
      //     // `$select=${dv_conf.FieldLogicalNames.join(",")}`,
      //     ""
      //   );
      //   return true;
      // } else {
      //   console.log(`on ${m} empty...`);
      //   return true;
      // }
    } catch (e) {
      console.error(e.stack);
    }
  });
  return true;
}

function getdata(
  args: Excel.TableChangedEventArgs,
  entitySet: string,
  queryString: string = "",
  queryOptions: string = ""
) {
  global.Callapiaction = {
    name: "callapiaction",
    action: {
      entitySet: entitySet,
      queryString: queryString,
      queryOptions: queryOptions,
    },
  };
  RetrieveMultipleData(setData);
}

export async function setData(result): Promise<any> {
  await Excel.run(async function (context) {
    let args = global.onChangeDataValidation.args;
    let config = global.onChangeDataValidation.config;
    let range = args.getRange(context);
    let dv_conf = null;
    let sheet = context.workbook.worksheets.getActiveWorksheet();

    try {
      console.log(`${m} setData`);
      console.log(result);
      // =INDIRECT("benz_prototypemodeldesignation")

      range.load(["rowIndex", "columnIndex", "address"]);
      await context.sync();

      // rowindex
      // const colconf = config.table.columns[range.columnIndex];
      let colconf: any =
        range.rowIndex > config.table.rowIndex
          ? config.table.columns[range.columnIndex]
          : config.layout.find((ele) => ele.rowIndex === range.rowIndex && ele.index === range.columnIndex);

      dv_conf = colconf.cell[args.type].find((e) => e.type.toLowerCase() === m.toLowerCase());
      let valueAfters = args.details.valueAfter;
      console.log("valueAfters");
      console.log(valueAfters);
      valueAfters = valueAfters.toString().split(dv_conf.separator ?? ",");
      console.log(valueAfters);
      let values = result.value.map((ele) => ele[colconf.FieldLogicalNames[0]]);
      console.log("values");
      console.log(values);
      if (colconf.cell.dataValidation) {
        colconf.cell.dataValidation.rule.list.source = values.join(",");
      } else {
        colconf.cell.values = " ";
      }
      await set_layout(sheet, [colconf]);

      // let userProfileInfo: string[] = [];
      if (result) {
        console.log("true result");
        console.log(result.value);
        if (result.value.length == 1) {
          setEditedRange(args, true, m);
        } else {
          console.log("false result value length != 1");
          setEditedRange(args, false, m);
        }
      } else {
        console.log("false result");
        throw new Error("No result");
      }
    } catch (exception) {
      console.log("false setData com_conf");
      // setEditedRange(args, true);
      console.error(`${m} setData EXCEPTION: ` + exception.message);
      console.error(result);
    }
  });
}
