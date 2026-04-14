/* global global, console, Excel */

import { changeArgs, dataValidationCheck, FormConfig } from "../../type";
import { allTrue } from "../init";
import { setEditedRange } from "./init";
import { RetrieveAndReturnMultipleData } from "../../../dataverse-data-helper";
import { get_table_data_arrary } from "../../utils";
import { showMessage } from "../../../message-helper";

const m = "tabledata_checkisnotinconjunctionwith";

// export async function getTableRes(
//   dataset,
//   value,
//   config,
//   parameters,
//   rowIndex: number,
//   columnIndex: number,
//   valueafter
// ) {
//   let objs = [];
//   console.log("getTableRes");
//   console.log("rowIndex");
//   console.log(dataset);
//   console.log(rowIndex - config.table.rowIndex - 1);
//   console.log(parameters);
//   rowIndex = rowIndex - config.table.rowIndex - 1;
//   let col_config = config.table.columns[columnIndex];
//   let dv_conf = col_config.cell["TableChanged"].find((e) => e.type.toLowerCase() === m.toLowerCase());
//   console.log(col_config);

//   if (col_config) {
//     console.log(dataset);
//     for (const row of dataset.items) {
//       let obj = {};
//       if(rowIndex == dataset.indexOf(row)){
//         for (const parameter of dv_conf.FieldLogicalNames) {
//           col_config = config.table.columns.find((e) => e.LogicalName == parameter.replace("{", "").replace("}", ""));
//           console.log(col_config);
//           console.log("row...");
//           console.log(row);
//           console.log(row["values"][0]);
//           console.log(row["values"][0][col_config.index]);

//           obj[parameter] = row["values"][0][col_config.index];
//           console.log(obj);
//         }
//       }
//       console.log(obj);

//       if (obj[dv_conf.findKey] == value[`{${dv_conf.findKey}}`]) {
//         objs.push(obj);
//       }
//     }
//   }
//   console.log(objs);

//   return objs;
// }
export async function getTableRes_inconjunctionwith(
  dataset,
  config,
  parameters,
  rowIndex: number,
  columnIndex: number,
  value,
  valueAfter
) {
  try {
    let objs = [];
    console.log("getTableRes");
    console.log(rowIndex - config.table.rowIndex - 1);
    console.log(parameters);
    console.log(value);
    rowIndex = rowIndex - config.table.rowIndex - 1;

    const colconf = config.table.columns[columnIndex];
    const dv_conf = colconf.cell["TableChanged"].find((e) => e.type.toLowerCase() === m.toLowerCase());
    if (!parameters) {
      console.log("null parameters", parameters);
      return [];
    }

    for (const row of dataset.items) {
      let obj = {};
      if (rowIndex != dataset.items.indexOf(row)) {
        let inc = true;
        obj[`{valueAfters}`] = valueAfter;
        for (const parameter of parameters) {
          console.log("parameter", parameter);

          const col_config = config.table.columns.find(
            (e) => e.LogicalName == parameter.replace("{", "").replace("}", "")
          );
          console.log(col_config);

          if (col_config) {
            var currentParamterValue = row["values"][0][col_config.index];
            console.log(row);
            console.log(row["values"][0]);
            console.log(currentParamterValue);
            console.log("col_config.LogicalName == parameter", col_config.LogicalName == parameter);
            console.log(
              `value[dv_conf.findKey] != row["values"][0][col_config.index]`,
              value[dv_conf.findKey] != currentParamterValue
            );

            obj[parameter] = row["values"][0][col_config.index];
            console.log("value obj");
            console.log(value);
            console.log(obj);
            if (value[`{${dv_conf.findKey}}`] != obj[`{${dv_conf.findKey}}`]) {
              inc = false;
            }
          }
        }
        if (inc) {
          objs.push(obj);
        }
      }
    }
    console.log("getTableRes", objs);
    let res = [];

    for (const obj of objs) {
      var isNoCommissionNo = false;
      let queryString1 = dv_conf.queryString1;
      var currentLineItemMeasure = "";
      var currentCommissionNo = "";
      var otherMeasureName = "";
      for (const key of Object.keys(obj)) {
        console.log(key);
        console.log(global.RowMeasureName[rowIndex]);
        if (key == "{valueAfters}") {
          currentLineItemMeasure = global.RowMeasureName[rowIndex];
          queryString1 = queryString1.replaceAll(key, currentLineItemMeasure);
        } else {
          if (key == "{benz_commissionnumber}") {
            currentCommissionNo = obj[key];
          }
          if (key == "{benz_prototypesalesmeasure}") {
            otherMeasureName = obj[key];
          }
          queryString1 = queryString1.replaceAll(key, obj[key]);
        }
        if (key == "{" + dv_conf.EntityLogicalName + "}" && obj[key] == "") {
          isNoCommissionNo = true;
        }
      }
      console.log("queryString1 updated");
      console.log(queryString1);
      console.log("isNoCommissionNo", isNoCommissionNo);
      if (isNoCommissionNo) {
        res.push(true);
      } else {
        var getDataResult = await getdata(null, dv_conf.EntityLogicalName1 + "s", queryString1, "");
        res.push(getDataResult.value.length == 0);

        if (!(getDataResult.value.length == 0)) {
          if (dv_conf.ErrorMsg != "") {
            var errorMsg = dv_conf.ErrorMsg.replace("{currentMeasure}", currentLineItemMeasure)
              .replace("{commissionNo}", currentCommissionNo)
              .replace("{otherMeasure}", otherMeasureName);
            throw new Error(errorMsg);
          }
        }
      }
    }

    console.log("getTableRes end res", res);
    return allTrue(res);
  } catch (e) {
    global.ErrorMsg = e.message;
    showMessage({ style: "error", message: `Row ${rowIndex + 1}: ${e.message}` });
    console.error(e.message);
  }
}
export async function getTableRes_inconjunctionwith_oncheck(
  dataset,
  config,
  parameters,
  rowIndex: number,
  columnIndex: number,
  value,
  valueAfter
) {
  try {
    let objs = [];
    console.log("getTableRes");
    console.log(rowIndex);
    console.log(parameters);
    console.log(value);

    const colconf = config.table.columns[columnIndex];
    const dv_conf = colconf.cell["TableChanged"].find((e) => e.type.toLowerCase() === m.toLowerCase());
    console.log("colconf");
    console.log(colconf);
    if (!parameters) {
      console.log("null parameters", parameters);
      return [];
    }
    console.log(dataset);
    for (var i = 0; i < dataset.length; i++) {
      var row = i;
      var rowValue = dataset[i];
      let obj = {};
      console.log("getTableRes_inconjunctionwith_oncheck");
      console.log(rowValue);
      console.log(rowValue.indexOf(rowValue));

      if (rowIndex != rowValue.indexOf(row)) {
        let inc = true;
        obj[`{valueAfters}`] = valueAfter;
        for (const parameter of parameters) {
          console.log("parameter", parameter);

          const col_config = config.table.columns.find(
            (e) => e.LogicalName == parameter.replace("{", "").replace("}", "")
          );

          if (col_config) {
            console.log(row);
            console.log("col_config.LogicalName == parameter", col_config.LogicalName == parameter);
            console.log(
              `value[dv_conf.findKey] != row[col_config.index]`,
              value[dv_conf.findKey] != rowValue[col_config.index]
            );
            console.log(value);
            console.log(dv_conf.findKey);
            console.log(rowValue);
            console.log(col_config.index);

            obj[parameter] = rowValue[col_config.index];
            console.log("value obj");
            console.log(value);
            console.log(obj);
            if (value[`{${dv_conf.findKey}}`] != obj[`{${dv_conf.findKey}}`]) {
              inc = false;
            }
          }
        }
        if (inc) {
          objs.push(obj);
        }
      }
    }
    console.log("getTableRes", objs);
    let res = [];
    var currentLineItemMeasure = "";
    var currentCommissionNo = "";
    var otherMeasureName = "";
    for (const obj of objs) {
      let queryString1 = dv_conf.queryString1;
      console.log(dv_conf);
      for (const key of Object.keys(obj)) {
        if (key == "{valueAfters}") {
          currentLineItemMeasure = global.RowMeasureName[rowIndex];
          queryString1 = queryString1.replaceAll(key, currentLineItemMeasure);
        } else {
          if (key == "{benz_commissionnumber}") {
            currentCommissionNo = obj[key];
          }
          if (key == "{benz_prototypesalesmeasure}") {
            otherMeasureName = obj[key];
          }
          queryString1 = queryString1.replaceAll(key, obj[key]);
        }
      }
      var getDataResult = await getdata(null, dv_conf.EntityLogicalName1 + "s", queryString1, "");
      res.push(getDataResult.value.length == 0);

      if (!(getDataResult.value.length == 0)) {
        if (dv_conf.ErrorMsg != "") {
          var errorMsg = dv_conf.ErrorMsg.replace("{currentMeasure}", currentLineItemMeasure)
            .replace("{commissionNo}", currentCommissionNo)
            .replace("{otherMeasure}", otherMeasureName);
          throw new Error(errorMsg);
        }
      }
    }
    console.log("getTableRes end res", res);
    return allTrue(res);
  } catch (e) {
    global.ErrorMsg = e.message;
    showMessage({ style: "error", message: `Row ${rowIndex + 1}: ${e.message}` });
    console.error(e.message);
  }
}

export async function getvalue(dataset, config, parameters, rowIndex: number, columnIndex: number, valueafter) {
  let obj = {};
  console.log("rowIndex");
  console.log(rowIndex - config.table.rowIndex - 1);
  console.log(parameters);
  rowIndex = rowIndex - config.table.rowIndex - 1;
  for (const parameter of parameters) {
    if (parameter == "{valueAfters}") {
      obj[parameter] = valueafter;
    }
    const col_config = config.table.columns.find((e) => e.LogicalName == parameter.replace("{", "").replace("}", ""));
    console.log(col_config);

    if (col_config) {
      console.log(dataset);
      console.log(dataset.items[rowIndex]);
      console.log(dataset.items[rowIndex]["values"][0]);
      console.log(dataset.items[rowIndex]["values"][0][col_config.index]);
      obj[parameter] = dataset.items[rowIndex]["values"][0][col_config.index];
      if (col_config.index == 2) {
        var rowNumber = rowIndex;
        global.RowMeasureName[rowNumber] = dataset.items[rowIndex]["values"][0][col_config.index];
      }
    }
  }
  return obj;
}

export async function init(args: Excel.TableChangedEventArgs, oconfig: FormConfig): Promise<boolean> {
  console.log(`on ${m}...`);
  console.log(args);
  if (args.details.valueAfter == "") {
    return true;
  }
  return await Excel.run(async function (context) {
    try {
      const config = oconfig;
      console.log(context);
      let range = args.getRange(context);
      range.load(["rowIndex", "columnIndex", "address"]);
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      let table = sheet.tables.getItemAt(0);
      let obj = table.toJSON();
      console.log("table obj");
      console.log(obj);
      table.load(["columns", "rows"]);
      // set columns name
      await context.sync();
      // let data = table.columns.toJSON();
      let data = table.rows.toJSON();
      console.log("table obj: ", [data]);
      console.log(config);
      const colconf = config.table.columns[range.columnIndex];
      const dv_conf = colconf.cell[args.type].find((e) => e.type.toLowerCase() === m.toLowerCase());
      if (!dv_conf) {
        return;
      }
      const valueAfter = args.details.valueAfter.toString();
      global.onChangeDataValidation = { args: args, config: config, columnIndex: range.columnIndex, type: m };
      if (valueAfter) {
        let res = [];
        // for (let valueAfter of valueAfters) {
        const parameters = dv_conf.queryString.match(/\{[^}]+\}/g);

        console.log("dv_conf.queryString");
        console.log(dv_conf.queryString);
        console.log("parameters");
        console.log(parameters);
        const values = parameters
          ? await getvalue(data, config, parameters, range.rowIndex, range.columnIndex, valueAfter)
          : {};
        console.log("values");
        console.log(values);
        let queryString = dv_conf.queryString;
        var currentMeasure = "";
        for (const key of Object.keys(values)) {
          queryString = queryString.replace(key, values[key]);
          if (key == "{benz_prototypesalesmeasure}") {
            currentMeasure = values[key];
          }
        }

        res.push(
          await getdata(
            args,
            dv_conf.EntityLogicalName + "s",
            queryString,
            // `$select=${dv_conf.FieldLogicalNames.join(",")}`,
            ""
          )
        );

        console.log(valueAfter);
        console.log(res);
        let v = [];
        for (const re of res) {
          for (let r of re.value) {
            console.log("v");
            console.log(v);
            for (const p of r.benz_prototypeclaimitem_Commissionnumber_benz_) {
              if (p.benz_PrototypeSalesMeasure != null) {
                var notInconjunctionWithRecordArr = p.benz_PrototypeSalesMeasure.benz_notinconjunctionwith.split(",");
                for (var i = 0; i < notInconjunctionWithRecordArr.length; i++) {
                  v.push(currentMeasure != notInconjunctionWithRecordArr[i]);
                }
              } else {
                v.push(p.benz_PrototypeSalesMeasure == null);
              }
            }
          }
        }

        //check current measure not conjunctionwith measureID
        var currentRow = range.rowIndex - config.table.rowIndex;
        var checkCurrentMeasureNotConjunction = await currentMeasureNotConjunction(
          values,
          dv_conf,
          currentMeasure,
          currentRow
        );
        v = v.concat(checkCurrentMeasureNotConjunction);

        const TableRes = await getTableRes_inconjunctionwith(
          data,
          config,
          parameters,
          range.rowIndex,
          range.columnIndex,
          values,
          valueAfter
        );
        v = v.concat(TableRes);
        console.log("end res");
        console.log(v);
        console.log(allTrue(v));
        // setData(args, res);
        return allTrue(v);
      } else {
        console.log(`on ${m} empty...`);
        return false;
      }
    } catch (e) {
      console.error(e.stack);
    }
  });
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
    const parameters = onchangeconfig.queryString.match(/\{[^}]+\}/g);
    const oconfig = object.config;
    const currentRow = object.row;
    const currentColumn = object.fconfig.index;
    const colconf = oconfig.table.columns[currentColumn];
    const dv_conf = colconf.cell["TableChanged"].find((e) => e.type.toLowerCase() === m.toLowerCase());
    var values = {};
    console.log("currentRow" + currentRow);
    for (const parameter of parameters) {
      if (parameter == "{valueAfters}") {
        values[parameter] = valueAfters[0];
      }
      const col_config = oconfig.table.columns.find((e) => e.LogicalName == parameter.replace("{", "").replace("}", ""));

      if (col_config) {
        if (parameter.includes("benz_prototypesalesmeasure")) {
          values[parameter] = global.RowMeasureName[currentRow];
        } else {
          values[parameter] = valueAfters[0];
        }
      }
    }
    let currentDatas = await get_table_data_arrary(global.CurrentConfig);

    let res = [];

    console.log("values");
    console.log(values);
    let queryString = dv_conf.queryString;
    var currentMeasure = "";
    for (const key of Object.keys(values)) {
      queryString = queryString.replace(key, values[key]);
      if (key == "{benz_prototypesalesmeasure}") {
        currentMeasure = values[key];
      }
    }
    console.log("queryString");
    console.log(dv_conf.queryString);
    res.push(
      await getdata(
        null,
        dv_conf.EntityLogicalName + "s",
        queryString,
        // `$select=${dv_conf.FieldLogicalNames.join(",")}`,
        ""
      )
    );

    // let result = await getdata(
    //   null,
    //   onchangeconfig.EntityLogicalName + "s",
    //   "?" + onchangeconfig.queryString.replace("{valueAfters}", valueAfter) + "&$count=true",
    //   ""
    // );

    // if (result != null && result != "") {
    //   res.push(result["@odata.count"] > 0);
    // } else {
    //   res.push(true);
    // }

    console.log(res);
    let v = [];
    for (const re of res) {
      if (re != null) {
        for (let r of re.value) {
          console.log("v");
          console.log(v);
          for (const p of r.benz_prototypeclaimitem_Commissionnumber_benz_) {
            if (p.benz_PrototypeSalesMeasure != null) {
              var notInconjunctionWithRecordArr = p.benz_PrototypeSalesMeasure.benz_notinconjunctionwith.split(",");
              for (var i = 0; i < notInconjunctionWithRecordArr.length; i++) {
                v.push(currentMeasure != notInconjunctionWithRecordArr[i]);
              }
            } else {
              v.push(p.benz_PrototypeSalesMeasure == null);
            }
          }
        }
      } else {
        v.push(true); //true for unfound measure in claim measure checking
      }
    }

    //check current measure not conjunctionwith measureID
    var checkCurrentMeasureNotConjunction = await currentMeasureNotConjunction(
      values,
      dv_conf,
      currentMeasure,
      currentRow
    );
    v = v.concat(checkCurrentMeasureNotConjunction);

    const TableRes = await getTableRes_inconjunctionwith_oncheck(
      currentDatas,
      oconfig,
      parameters,
      currentRow,
      currentColumn,
      values,
      valueAfters[0]
    );
    v = v.concat(TableRes);
    console.log("end res");
    console.log(v);
    console.log(allTrue(v));
    // setData(args, res);
    return allTrue(v);
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

export async function currentMeasureNotConjunction(values, dv_conf, currentMeasure, currentRow) {
  try {
    var res = [];

    var getNotConjunctionRecord =
      "$select=benz_notinconjunctionwith&$filter=benz_name eq '{benz_prototypesalesmeasure}' ";
    for (const key of Object.keys(values)) {
      getNotConjunctionRecord = getNotConjunctionRecord.replace(key, values[key]);
    }
    var currentMeasureNotConjunction = await getdata(null, "benz_prototypesalesmeasures", getNotConjunctionRecord, "");
    console.log("currentMeasureNotConjunction");
    console.log(currentMeasureNotConjunction);
    console.log(currentMeasureNotConjunction["value"]);
    if (currentMeasureNotConjunction["value"].length > 0) {
      if (currentMeasureNotConjunction["value"][0]["benz_notinconjunctionwith"] != null) {
        var currentMeasureNotConjunctionArr =
          currentMeasureNotConjunction["value"][0]["benz_notinconjunctionwith"].split(",");
        for (var i = 0; i < currentMeasureNotConjunctionArr.length; i++) {
          var measureID = currentMeasureNotConjunctionArr[i];
          var commissionNo = "";
          let newQueryString =
            "$filter=benz_name eq '{benz_commissionnumber}' and benz_savedmeasureidname eq '{benz_prototypesalesmeasure}' and statecode eq 0";
          for (const key of Object.keys(values)) {
            if (key == "{benz_prototypesalesmeasure}") {
              newQueryString = newQueryString.replace(key, measureID);
            } else {
              if (key == "{benz_commissionnumber}") {
                commissionNo = values[key];
              }
              newQueryString = newQueryString.replace(key, values[key]);
            }
          }

          if (global.updatingRecord != undefined) {
            newQueryString += ` and benz_claimname eq '${global.updatingRecord}'`;
          }

          console.log("newQueryString");
          console.log(newQueryString);

          var apiResult = await getdata(
            null,
            dv_conf.EntityLogicalName + "s",
            newQueryString + "&$count=true",
            // `$select=${dv_conf.FieldLogicalNames.join(",")}`,
            ""
          );
          console.log(apiResult);
          console.log(apiResult["@odata.count"]);
          res.push(apiResult["@odata.count"] == 0);
          if (!(apiResult["@odata.count"] == 0)) {
            throw new Error(
              dv_conf.ErrorMsg.replace("{currentMeasure}", currentMeasure)
                .replace("{commissionNo}", commissionNo)
                .replace("{otherMeasure}", measureID)
            );
          }
        }
      }
    } else {
      res.push(true);
    }
  } catch (e) {
    global.ErrorMsg = e.message;
    showMessage({ style: "error", message: `Row ${currentRow}: ${e.message}` });
  }

  return res;
}
