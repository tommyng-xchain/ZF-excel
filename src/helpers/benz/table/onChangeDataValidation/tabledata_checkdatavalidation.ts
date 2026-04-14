/* global global, console, Excel */

import { dataValidationCheck, FormConfig, onchange_dataValidation, setFieldData } from "../../type";
import { RetrieveAndReturnMultipleData } from "../../../dataverse-data-helper";
import { lockmainObjectValueToCell, set_layout } from "../../utils";
import { showMessage } from "../../../message-helper";
import { callRetrieveMultipleData } from "../../../middle-tier-calls";
import { allTrue } from "../init";

const m = "tabledata_checkdatavalidation";

function formatDate(date) {
  var d = new Date(date),
    month = "" + (d.getMonth() + 1),
    day = "" + d.getDate(),
    year = d.getFullYear();

  if (month.length < 2) month = "0" + month;
  if (day.length < 2) day = "0" + day;

  return [year, month, day].join("-");
}

export async function init(args: Excel.TableChangedEventArgs, config: FormConfig): Promise<boolean> {
  console.log(`on ${m}...`);
  console.log(args);
  if (args.details.valueAfter == "") {
    return true;
  }
  return await Excel.run(async function (context) {
    return true;
    // try {
    //   let range = args.getRange(context);
    //   range.load(["rowIndex", "columnIndex", "address"]);
    //   await context.sync();
    //   console.log(config);

    //   const colconf = config.table.columns[range.columnIndex];
    //   const dv_conf = colconf.cell[args.type].find((e) => e.type.toLowerCase() === m.toLowerCase());
    //   if (!dv_conf) {
    //     return;
    //   }
    //   const recordColumn = range.rowIndex - config.table.rowIndex - 1;
    //   const valueAfters = args.details.valueAfter.toString().split(dv_conf.separator ?? ",");
    //   // global.onChangeDataValidation = { args: args, config: config, columnIndex: range.columnIndex, type: m };
    //   if (valueAfters) {
    //     return await getdata(config, dv_conf, valueAfters[0], recordColumn);
    //   } else {
    //     console.log(`on ${m} empty...`);
    //     throw new Error(`Can't find '${valueAfters}'`);
    //   }
    // } catch (e) {
    //   console.error(e.stack);
    //   showMessage({ style: "error", message: e.message });
    //   return false;
    // }
  });
}

export async function check(object: dataValidationCheck) {
  try {
    let { config, onchangeconfig, value, row } = object;
    console.log(`on ${m}...`);
    console.log(onchangeconfig);
    console.log(value);
    if (value != undefined) {
      const valueAfters = value.toString().split(onchangeconfig.separator ?? ",");
      const getvalueregex = "([\\w\\-]+)";
      const regex = RegExp(getvalueregex, "g");
      const checkHasGUID = regex.exec(value)[0];
      // if (checkHasGUID != "") {
      //   return true;
      // } else {
      if (valueAfters) {
        return await getdata(config, onchangeconfig, valueAfters[0], row);
      } else {
        console.log(`on ${m} empty...`);
        throw new Error(`Can't find '${valueAfters}'`);
      }
      // }
    } else {
      return true;
    }
  } catch (e) {
    console.error(e.stack);
    showMessage({ style: "error", message: e.message });
    return false;
  }
}

export async function getdata(config: FormConfig, dv_conf: onchange_dataValidation, valueAfters: any, row: number) {
  console.log(row);
  var measureID = global.RowMeasureName[row];
  console.log(global.RowMeasureName);
  var fromLogicalNames = dv_conf.setFieldData[0].FromLogicalNames.join(",");
  global.Callapiaction = {
    name: "callapiaction",
    action: {
      entitySet: dv_conf.EntityLogicalName + "s",
      queryString:
        "?" + dv_conf.queryString.replace("{valueAfters}", measureID).replace("{fromLogicalNames}", fromLogicalNames),
      queryOptions: dv_conf.queryOptions,
    },
  };

  // RetrieveMultipleData(setData);
  let res = await RetrieveAndReturnMultipleData(null, {
    entitySet: dv_conf.EntityLogicalName + "s",
    queryString:
      "?" + dv_conf.queryString.replace("{valueAfters}", measureID).replace("{fromLogicalNames}", fromLogicalNames),
    queryOptions: dv_conf.queryOptions,
  });
  return await setData(res, config, dv_conf, valueAfters, row);
}

export async function setData(
  result,
  config: FormConfig,
  dv_conf: onchange_dataValidation,
  valueAfters: any,
  row: number
): Promise<any> {
  let isVaild = false;
  try {
    // let sheet = null;
    // await Excel.run(async function (context) {
    //   sheet = context.workbook.worksheets.getActiveWorksheet();
    //   await context.sync();
    // });
    console.log(`${m} setData`);
    console.log(result);
    // rowindex
    console.log("valueAfters");
    console.log(valueAfters);
    valueAfters = valueAfters.split(dv_conf.separator ?? ",")[0];
    console.log(valueAfters);

    // let userProfileInfo: string[] = [];
    if (result) {
      console.log("true result");
      console.log(result.value);
      if (result.value.length > 0) {
        isVaild = true;
        let res = result.value[0];
        let sets: setFieldData[] = dv_conf.setFieldData;
        let validationResults = [];
        for (let setFieldConf of sets) {
          // for (let { index, fieldname } of dv_conf.tablefield.map((fieldname, index) => ({ index, fieldname }))) {
          for (let ToLogicalName of setFieldConf.ToLogicalName) {
            let com_conf = config.table.columns.find((ele) => ele.LogicalName == ToLogicalName);
            let validationResult = false;
            console.log("setData com_conf");
            console.log(com_conf);
            if (setFieldConf.settype == "list") {
              let LogicalName = setFieldConf.FromLogicalNames[0];

              if (LogicalName == "benz_FinanceType.benz_name") {
                var measureName = res["benz_name"];
                var fsProductArr = await getRecordByUrl(global.ApiAccessToken, measureName);
                console.log("fsProductArr");
                console.log(fsProductArr);
                if (fsProductArr["value"][0]["benz_FinanceType"] != null) {
                  var fsProductNameArr =
                    fsProductArr["value"][0]["benz_FinanceType"]["benz_fsproduct_FinanceProduct_benz_financeproduct"];
                  console.log("fsProductNameArr");
                  console.log(fsProductNameArr);
                  var nameArr = [];
                  fsProductNameArr.map((v) => nameArr.push(v["benz_name"]));
                  console.log(nameArr);
                  var valueAfter = await getFinanceProduct(global.ApiAccessToken, valueAfters);
                  console.log(valueAfter);

                  if (nameArr.includes(valueAfter["benz_name"])) {
                    validationResult = true;
                  }
                }
              } else {
                console.log("setListData " + LogicalName);
                let keys = LogicalName.toString().split(".");
                console.log(keys);
                console.log(res);
                let arr = res[keys[0]];
                if (arr) {
                  console.log(arr);
                  console.log(valueAfters);
                  if (!Array.isArray(arr)) {
                    console.log("arr is not array");
                    if (arr == "ALL") {
                      validationResult = true;
                    } else {
                      var newArr = arr.split(",");
                      if (newArr.includes(valueAfters)) {
                        validationResult = true;
                      }
                    }
                  } else {
                    console.log("arr is array");
                    if (arr.includes(valueAfters)) {
                      validationResult = true;
                    }
                  }
                }
              }
            } else if (setFieldConf.settype == "date") {
              if (setFieldConf.FromLogicalNames.length == 2) {
                console.log("setDate " + ToLogicalName);
                console.log(setFieldConf);
                let formulaDate1 = res[setFieldConf.FromLogicalNames[0]];
                let formulaDate2 = res[setFieldConf.FromLogicalNames[1]];
                if (formulaDate1 != null) {
                  console.log("formulaDate1");
                  console.log(formulaDate1);
                  console.log("formulaDate2");
                  console.log(formulaDate2);
                  if (setFieldConf.operator == "Between") {
                    console.log(formatDate(formulaDate1));
                    console.log(formatDate(formulaDate2));
                    if (
                      formatDate(valueAfters) >= formatDate(formulaDate1) &&
                      formatDate(valueAfters) <= formatDate(formulaDate2)
                    ) {
                      validationResult = true;
                    }
                  }
                } else {
                  if (setFieldConf.FromLogicalNames[0] == "benz_finalcontractdatefrom" || setFieldConf.FromLogicalNames[0] == "benz_finalapplicationdatefrom") {
                    console.log("LessThanOrEqualTo formulaDate2 >= valueAfters");
                    console.log(formatDate(formulaDate2) >= formatDate(valueAfters));
                    if (formulaDate2 != null) {
                      if (formatDate(formulaDate2) >= formatDate(valueAfters)) {
                        validationResult = true;
                      }
                    } else {
                      validationResult = true;
                    }
                  }
                }
              } else if (setFieldConf.FromLogicalNames.length == 1) {
                console.log("setDate " + ToLogicalName);
                console.log(setFieldConf);
                let formulaDate1 = res[setFieldConf.FromLogicalNames[0]];
                // let formulaDate2 = new Date();
                console.log(formulaDate1);
                if (formulaDate1 != null) {
                  if (setFieldConf.operator == "LessThanOrEqualTo") {
                    console.log("LessThanOrEqualTo formulaDate1 > valueAfters");
                    console.log(formatDate(formulaDate1) >= formatDate(valueAfters));
                    if (formatDate(formulaDate1) >= formatDate(valueAfters)) {
                      validationResult = true;
                    }
                  } else if (setFieldConf.operator == "GreaterThanOrEqualTo") {
                    console.log("GreaterThanOrEqualTo formulaDate1 < valueAfters");
                    console.log(formatDate(formulaDate1) <= formatDate(valueAfters));
                    if (formatDate(formulaDate1) <= formatDate(valueAfters)) {
                      validationResult = true;
                    }
                  }
                }
                // if (setFieldConf.canPass) {
                //   validationResult = true;
                // }
              } else {
                console.error("set date is empty");
              }
            } else if (setFieldConf.settype == "value") {
              let LogicalName = setFieldConf.FromLogicalNames[0];
              console.log("setData " + LogicalName);
              console.log(res[LogicalName]);
            } else if (setFieldConf.settype == "decimal") {
              let LogicalName = setFieldConf.FromLogicalNames[0];
              console.log("setData " + LogicalName);
              console.log(res[LogicalName]);

              const canExceed = res[setFieldConf.ExceedCheck];
              var formula1Set = res[setFieldConf.FromLogicalNames[0]] == null ? 0 : res[setFieldConf.FromLogicalNames[0]];
              let operatorSet = setFieldConf.operator;

              if (canExceed == 1) {
                formula1Set = 0;
                if (valueAfters >= formula1Set) {
                  validationResult = true;
                }
              } else {
                if (operatorSet == "LessThanOrEqualTo" && valueAfters <= formula1Set) {
                  validationResult = true;
                }
                if (operatorSet == "GreaterThanOrEqualTo" && valueAfters >= formula1Set) {
                  validationResult = true;
                }
              }
            }
            validationResults.push(validationResult);
          }
        }

        return allTrue(validationResults);
        // setEditedRange(args, true);
      } else {
        console.log("false result value");
        return false;
        // setEditedRange(args, false);
        // throw new Error("No result");
      }
    } else {
      console.log("false result");
      return false;
      // setEditedRange(args, false);
      // throw new Error("No result");
    }
  } catch (e) {
    // setEditedRange(args, false);
    console.error(e.stack);
    console.log("false setData com_conf");
    console.error(`${m} setData EXCEPTION: ` + e.message);
    console.error(result);
  }

  return isVaild;
}

function getRecordByUrl(token,measureName){
  var args = {
    entitySet: "benz_prototypesalesmeasures",
    queryString: "$filter=benz_name eq '" + measureName +"'&$expand=benz_FinanceType($select=benz_name;$expand=benz_fsproduct_FinanceProduct_benz_financeproduct($select=benz_name))",
    queryOptions: ""
  };

  return callRetrieveMultipleData(token, args);
}

function getFinanceProduct(token, valueAfters) {
  var args = {
    entitySet: valueAfters,
    queryString: "?$select=benz_name",
    queryOptions: ""
  };

  return callRetrieveMultipleData(token, args);
}