/* global global, console, Excel */

import { changeArgs, dataValidationCheck, FormConfig } from "../../type";
import { setEditedRange } from "./init";
import { RetrieveAndReturnMultipleData} from "../../../dataverse-data-helper";
import { allTrue } from "../init";
import { showMessage } from "../../../message-helper";

const m = "tabledata_getdataisnotexist";

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
      await context.sync();
      console.log(config);

      const colconf = config.table.columns[range.columnIndex];
      const dv_conf = colconf.cell[args.type].find((e) => e.type.toLowerCase() === m.toLowerCase());
      if (!dv_conf) {
        return;
      }
      const valueAfters = args.details.valueAfter.toString().split(dv_conf.separator ?? ",");
      global.onChangeDataValidation = { args: args, config: config, columnIndex: range.columnIndex, type: m };

      var measureID = "";
      console.log(global.RowMeasureName);
      if (global.RowMeasureName.length > 0) {
        var rowIndex = range.rowIndex - config.table.rowIndex - 1;
        measureID = global.RowMeasureName[rowIndex];
      }

      console.log("valueAfters");
      console.log(valueAfters);
      if (valueAfters) {
        let res = [];
        console.log("global.updatingRecord");
        console.log(global.updatingRecord);
        var apiString = dv_conf.queryString;
        // if(global.updatingRecord != undefined){
        if (global.updatingRecord != undefined) {
          apiString += ` and benz_claimname ne '${global.updatingRecord}'`;
        }
        for (let valueAfter of valueAfters) {
          let result = await getdata(
            args,
            dv_conf.EntityLogicalName + "s",
            "?" + apiString.replace("{valueAfters}", valueAfter).replace("{measureID}", measureID) + "&$count=true",
            ""
          );
          console.log("result");
          console.log(result);
          res.push(result["@odata.count"] == 0);

          if (!(result["@odata.count"] == 0)) {
            if (dv_conf.ErrorMsg != "") {
              const errorMsg = dv_conf.ErrorMsg.replace("{currentMeasure}", measureID).replace("{commissionNo}", valueAfter);
              throw new Error(errorMsg);
            }
          }
        }

        console.log("result");
        console.log(res);
        return allTrue(res);
      } else {
        console.log(`on ${m} empty...`);
      }
    } catch (e) {
      global.ErrorMsg = e.message;
      showMessage({ style: "error", message: `Row ${rowIndex + 1}: ${e.message}` });
    }
    return false;
  });
}

async function getdata(
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

  return await RetrieveAndReturnMultipleData(null, {
    entitySet: entitySet,
    queryString: queryString,
    queryOptions: queryOptions,
  });
}

/*
export async function setData(result): Promise<any> {
  await Excel.run(async function (context) {
    let args = global.onChangeDataValidation.args;
    let config = global.onChangeDataValidation.config;
    let range = args.getRange(context);
    let dv_conf = null;
    try {
      console.log(`${m} setData`);
      console.log(result);
      // =INDIRECT("benz_prototypemodeldesignation")

      range.load(["rowIndex", "columnIndex", "address"]);
      await context.sync();

      // rowindex
      const colconf = config.table.columns[range.columnIndex];
      dv_conf = colconf.cell[args.type].find((e) => e.type.toLowerCase() === m.toLowerCase());
      let valueAfters = args.details.valueAfter;
      console.log("valueAfters");
      console.log(valueAfters);
      valueAfters = valueAfters.toString().split(dv_conf.separator ?? ",");
      console.log(valueAfters);

      // let userProfileInfo: string[] = [];
      if (result) {
        console.log("true result");
        console.log(result.value);
        if (result.value.length == 0) {
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
      setEditedRange(args, false, m);
      // setEditedRange(args, true);
      console.error(`${m} setData EXCEPTION: ` + exception.message);
      console.error(result);
    }
  });
}
*/
export async function setData(args: changeArgs, results: any[]): Promise<any> {
  // await Excel.run(async function (context) {
  // let args = global.onChangeDataValidation.args;
  // let config = global.onChangeDataValidation.config;
  // let range = args.getRange(context);
  // let dv_conf = null;
  try {
    //   console.log(`${m} setData`);
    //   console.log(result);
    //   // =INDIRECT("benz_prototypemodeldesignation")

    //   range.load(["rowIndex", "columnIndex", "address"]);
    //   await context.sync();

    //   // rowindex
    //   const colconf = config.table.columns[range.columnIndex];
    //   dv_conf = colconf.cell[args.type].find((e) => e.type.toLowerCase() === m.toLowerCase());
    //   let valueAfters = args.details.valueAfter;
    //   console.log("valueAfters");
    //   console.log(valueAfters);
    //   valueAfters = valueAfters.toString().split(dv_conf.separator ?? ",");
    //   console.log(valueAfters);

    // let userProfileInfo: string[] = [];
    let isVaildresults = results.map((e) => e.value.length == 1);

    // // for (let result of results) {
    //   if (result) {
    //     console.log("true result");
    //     console.log(result.value);
    //   } else {
    //     console.log("false result");
    //     throw new Error("No result");
    //   }
    if (allTrue(isVaildresults)) {
      setEditedRange(args, true, m);
    } else {
      console.log("false result value length != 1");
      setEditedRange(args, false, m);
    }
  } catch (exception) {
    console.log("false setData com_conf");
    // setEditedRange(args, true);
    console.log(results);
    console.error(`${m} setData EXCEPTION: ` + exception.message);
  }
  // });
}

export async function check(object: dataValidationCheck) {
  try {
    let { onchangeconfig, value } = object;
    console.log(`on ${m}...`);
    console.log(onchangeconfig);
    console.log(value);
    if (!onchangeconfig) {
      throw new Error("Can't find conf");
    }
    if (!value) {
      return true;
    }
    const valueAfters = value.toString().split(onchangeconfig.separator ?? ",");

    var measureID = "";
    var measureName = "";
    console.log(global.RowMeasureID);
    if (global.RowMeasureID.length > 0) {
      var rowIndex = object.row;
      console.log(rowIndex);
      measureID = global.RowMeasureID[rowIndex].replace("benz_prototypesalesmeasures(", "").replace(")", "");
      measureName = global.RowMeasureName[rowIndex];
    }
    let res = [];
    var apiString = onchangeconfig.queryStringCheck;
    if (global.updatingRecord != undefined) {
      apiString += ` and benz_claimname ne '${global.updatingRecord}'`;
    }
    for (let valueAfter of valueAfters) {
      let result = await getdata(
        null,
        onchangeconfig.EntityLogicalName + "s",
        "?" + apiString.replace("{valueAfters}", valueAfter).replace("{measureID}", measureID) + "&$count=true",
        ""
      );
      console.log(result);
      res.push(result["@odata.count"] == 0);
      if (!(result["@odata.count"] == 0)) {
        if (onchangeconfig.ErrorMsg != "") {
          var errorMsg = onchangeconfig.ErrorMsg.replace("{currentMeasure}", measureName).replace(
            "{commissionNo}",
            valueAfter
          );
          console.log(errorMsg);

          throw new Error(errorMsg);
        }
      }
    }

    return allTrue(res);
  } catch (e) {
    global.ErrorMsg = e.message;
    showMessage({ style: "error", message: `${e.message}` });
    console.error(e.stack);
    return false;
  }
}
