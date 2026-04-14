/* global global, console, Excel */

import { changeArgs, dataValidationCheck, FormConfig, AppMode } from "../../type";
import { setEditedRange } from "./init";
import { RetrieveAndReturnMultipleData } from "../../../dataverse-data-helper";
import { allTrue } from "../init";

const m = "tabledata_getdataisexist";

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
      const migrationPassArr = ["benz_inconjunctionwith", "benz_notinconjunctionwith"];
      if (global.mode == "migration" && migrationPassArr.includes(colconf.LogicalName)) {
        return true;
      }
      const dv_conf = colconf.cell[args.type].find((e) => e.type.toLowerCase() === m.toLowerCase());
      if (!dv_conf) {
        return;
      }
      const valueAfters = args.details.valueAfter.toString().split(dv_conf.separator ?? ",");
      global.onChangeDataValidation = { args: args, config: config, columnIndex: range.columnIndex, type: m };

      if (valueAfters) {
        let res = [];
        for (let valueAfter of valueAfters) {
          let result = await getdata(
            args,
            dv_conf.EntityLogicalName + "s",
            "?" + dv_conf.queryString.replace("{valueAfters}", valueAfter) + "&$count=true",
            ""
          );
          console.log("result");
          console.log(result);
          res.push(result["@odata.count"] > 0);
        }

        console.log("result");
        console.log(res);
        return allTrue(res);
      } else {
        console.log(`on ${m} empty...`);
      }
    } catch (e) {
      console.error(e.stack);
    }
    return false;
  });
}

export async function check(object: dataValidationCheck) {
  try {
    let { onchangeconfig, value, row, fconfig } = object;
    console.log(`on ${m}...`);
    console.log(onchangeconfig);
    console.log(value);
    if (!onchangeconfig) {
      throw new Error("Can't find conf");
    }
    if (!value) {
      return true;
    }

    const migrationPassArr = ["benz_inconjunctionwith", "benz_notinconjunctionwith"];
    if (global.mode == "migration" && migrationPassArr.includes(fconfig.LogicalName)) {
      return true;
    }

    const valueAfters = value.toString().split(onchangeconfig.separator ?? ",");

    let res = [];
    for (let valueAfter of valueAfters) {
      var queryString = onchangeconfig.queryString.replace("{valueAfters}", valueAfter) + "&$count=true";

      if (onchangeconfig.queryString1 != undefined && onchangeconfig.queryString1 != "") {
        valueAfter = valueAfter.replace(onchangeconfig.EntityLogicalName + "s(", "").replace(")", "");
        queryString = onchangeconfig.queryString1.replace("{" + onchangeconfig.EntityLogicalName + "}", valueAfter) + "&$count=true";
      }
      let result = await getdata(null, onchangeconfig.EntityLogicalName + "s", "?" + queryString, "");
      console.log(result);
      res.push(result["@odata.count"] > 0);
    }

    return allTrue(res);
  } catch (e) {
    console.error(e.stack);
    return false;
  }
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
