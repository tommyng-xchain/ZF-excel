/* global global, console, Excel, Office, require */

import { RetrieveMultipleData } from "../../dataverse-data-helper";
import { lockmainObjectValueToCell } from "../utils";
import { set_choicesSets } from "./choices";
var jq = require("jquery");
const m = "systemusersearch";
let LogicalName: string = null;
Office.onReady((info) => {
  console.log("onload Office Ready");
  if (info.host === Office.HostType.Excel) {
    init();
  }
});
export function showhide(conditional: boolean, name?: string) {
  if (conditional) {
    jq(`#tab__${m}`).show();
    // jq(`#${m}`).show();
    jq(`#tab__${m}`).click();
    jq(`#${m}`).addClass("show active");
    LogicalName = name;
  } else {
    jq(`#tab__${m}`).hide();
    jq(`#${m}`).hide();
  }
}
export function init() {
  jq(`#${m}_search`).on("change", function () {
    let val = jq(`#${m}_search`).val();
    console.log(`input change... ${val}`);
    if (val.length >= 3) {
      getdata(val);
    }
    // getData(jq("search").val);
  });
  jq(`#${m}_apply`).on("click", async function () {
    let val = jq(`input[name='${m}']:checked`)[0].value;
    console.log(`apply click... ${val}`);
    if (val) {
      await Excel.run(async function (context) {
        try {
          console.log(`on ${m}...`);
          const range = context.workbook.getSelectedRange();
          range.values = [[val]];
          range.load(["rowIndex", "columnIndex", "address"]);
          await lockmainObjectValueToCell(false);
          await context.sync();
          await lockmainObjectValueToCell(true);

          let colconf: any =
            range.rowIndex > global.CurrentConfig.table.rowIndex
              ? global.CurrentConfig.table.columns[range.columnIndex]
              : global.CurrentConfig.layout.find(
                  (ele) => ele.rowIndex === range.rowIndex && ele.index === range.columnIndex
                );

          global.form.object[colconf.LogicalName] = jq(`input[name='${m}']`)[0].id;
          set_choicesSets(LogicalName, global.form.object[m]);
        } catch (e) {
          console.error(`${m} setData EXCEPTION: ` + e.message);
          console.error(e.stack);
        }
      });
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
    if (dv) {
      let args = global.onChangeDataValidation.args;
      console.log(args);
      let config = global.onChangeDataValidation.config;
      console.log(config);
    }
    // let userProfileInfo: string[] = [];
    if (result) {
      console.log("true result");
      console.log(result.value);
      let html = "";
      if (result.value) {
        global.form.object[m] = result.value.map((user) => ({
          id: user["systemuserid"].toString(),
          name: `${user["fullname"]} <${user["internalemailaddress"]}>`,
          value: user["systemuserid"].toString(),
        }));

        for (let user of result.value) {
          html += `<div class="p-2">
        <input type="radio" class="btn-check" name="${m}" id="${user["systemuserid"].toString()}" value="${user[
            "fullname"
          ].toString()}<${user["internalemailaddress"].toString()}>" autocomplete="off">
        <label class="btn w-100 btn-outline-secondary" for="${user["systemuserid"].toString()}">${user[
            "fullname"
          ].toString()}</label>
    </div>`;
        }
        // console.log(html);
        jq(`#${m}_dataset`).empty().append(html);
      } else {
        console.log("false result");
        throw new Error("No result");
      }
    }
  } catch (e) {
    console.error(`${m} setData EXCEPTION: ` + e.message);
    console.error(result);
    console.error(e.stack);
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
