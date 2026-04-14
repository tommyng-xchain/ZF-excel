/* global global, console, Excel, require */

import { FormConfig } from "../../type";
import { RetrieveMultipleData } from "../../../dataverse-data-helper";
// import * as lookup from "../../../systemuserlookupdialog";
var jq = require("jquery");
const m = "tabledata_lookupsystemuser";

export async function init(args: Excel.WorksheetSingleClickedEventArgs, config: FormConfig): Promise<boolean> {
  return await Excel.run(async function (context) {
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
          : config.layout.find(
              (ele) => ele.cell.address.toLocaleLowerCase() == range.address.split("!")[1].toLocaleLowerCase()
            );
      const dv_conf = colconf.cell[args.type].find((e) => e.type.toLowerCase() === m.toLowerCase());
      if (!dv_conf) {
        return;
      }

      // lookup.showdialog();

      // const valueAfters = args.details.valueAfter.toString().split(dv_conf.separator ?? ",");
      global.onChangeDataValidation = { args: args, config: config, columnIndex: range.columnIndex, type: m };
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
      // showdialog();
      return true;
    } catch (e) {
      console.error(e.stack);
      return;
    }
  });
  // return true;
}
export function showdialog() {
  jq("input#usersearch").on("change", function () {
    console.log(`input change... `);
    if (jq("usersearch").val.length >= 3) {
      getdata(jq("usersearch").val);
    }
    // getData(jq("search").val);
  });
}

export function getdata(val) {
  global.Callapiaction = {
    name: "callapiaction",
    action: {
      entitySet: "systemusers",
      queryString: `$select=fullname,systemuserid,internalemailaddress&$filter=startswith(internalemailaddress,'${val}')`,
      queryOptions: "",
    },
  };
  RetrieveMultipleData(setData);
}

export async function setData(result): Promise<any> {
  try {
    console.log(result);
    let dv = global.onChangeDataValidation;
    console.log(dv);
    let args = global.onChangeDataValidation.args;
    console.log(args);
    let config = global.onChangeDataValidation.config;
    console.log(config);
    // let userProfileInfo: string[] = [];
    if (result) {
      console.log("true result");
      console.log(result.value);
      let html = "";
      if (result.value) {
        for (let user of result.value) {
          html += `<div class="p-2">
        <input type="radio" class="btn-check" name="options-base" id="${user["systemuserid"]}" autocomplete="off" checked>
        <label class="btn w-100" for="${user["systemuserid"]}">${user["fullname"]}</label>
    </div>`;
        }
        // if (result.value.length == 1) {
        //   setEditedRange(args, true, m);
        // } else {
        //   console.log("false result value length != 1");
        //   setEditedRange(args, false, m);
        // }
        jq("userdataset").append(html);
      } else {
        console.log("false result");
        throw new Error("No result");
      }
    }
  } catch (exception) {
    console.log("false setData com_conf");
    // setEditedRange(args, true);
    console.error(`${m} setData EXCEPTION: ` + exception.message);
    console.error(result);
  }
}

// export async function setDataToRange(result): Promise<any> {
//   await Excel.run(async function (context) {
//     let args = global.onChangeDataValidation.args;
//     let config = global.onChangeDataValidation.config;
//     // let range = args.getRange(context);
//     let dv_conf = null;
//     let sheet = context.workbook.worksheets.getActiveWorksheet();

//     try {
//       console.log(`${m} setData`);
//       console.log(result);
//       // =INDIRECT("benz_prototypemodeldesignation")

//       range.load(["rowIndex", "columnIndex", "address"]);
//       await context.sync();

//       // rowindex
//       // const colconf = config.table.columns[range.columnIndex];
//       let colconf: any =
//         range.rowIndex > config.table.rowIndex
//           ? config.table.columns[range.columnIndex]
//           : config.layout.find((ele) => ele.rowIndex === range.rowIndex && ele.index === range.columnIndex);

//       dv_conf = colconf.cell[args.type].find((e) => e.type.toLowerCase() === m.toLowerCase());
//       let valueAfters = args.details.valueAfter;
//       console.log("valueAfters");
//       console.log(valueAfters);
//       valueAfters = valueAfters.toString().split(dv_conf.separator ?? ",");
//       console.log(valueAfters);
//       let values = result.value.map((ele) => ele[colconf.FieldLogicalNames[0]]);
//       console.log("values");
//       console.log(values);
//       if (colconf.cell.dataValidation) {
//         colconf.cell.dataValidation.rule.list.source = values.join(",");
//       } else {
//         colconf.cell.values = " ";
//       }
//       await set_layout(sheet, [colconf]);

//       // let userProfileInfo: string[] = [];
//       if (result) {
//         console.log("true result");
//         console.log(result.value);
//         if (result.value.length == 1) {
//           setEditedRange(args, true, m);
//         } else {
//           console.log("false result value length != 1");
//           setEditedRange(args, false, m);
//         }
//       } else {
//         console.log("false result");
//         throw new Error("No result");
//       }
//     } catch (exception) {
//       console.log("false setData com_conf");
//       // setEditedRange(args, true);
//       console.error(`${m} setData EXCEPTION: ` + exception.message);
//       console.error(result);
//     }
//   });
// }
