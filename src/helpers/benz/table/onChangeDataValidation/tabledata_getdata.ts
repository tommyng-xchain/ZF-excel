/* global global, console, Excel */

import { dataValidationCheck, FormConfig, onchange_dataValidation, setFieldData } from "../../type";
import { RetrieveAndReturnMultipleData } from "../../../dataverse-data-helper";
import { lockmainObjectValueToCell, set_layout, get_table_data_arrary } from "../../utils";
import { showMessage } from "../../../../helpers/message-helper";
import { callRetrieveMultipleData } from "../../../../helpers/middle-tier-calls";

const m = "tabledata_getdata";

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
    try {
      let range = args.getRange(context);
      range.load(["rowIndex", "columnIndex", "address"]);
      await context.sync();
      console.log(config);
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
      await getvalue(data, config, range.rowIndex, range.columnIndex, args.details.valueAfter);

      const colconf = config.table.columns[range.columnIndex];
      const dv_conf = colconf.cell[args.type].find((e) => e.type.toLowerCase() === m.toLowerCase());
      if (!dv_conf) {
        return;
      }
      const valueAfters = args.details.valueAfter.toString().split(dv_conf.separator ?? ",");
      // global.onChangeDataValidation = { args: args, config: config, columnIndex: range.columnIndex, type: m };
      if (valueAfters) {
        return await getdata(config, dv_conf, valueAfters[0], range.rowIndex);
      } else {
        console.log(`on ${m} empty...`);
        throw new Error(`Can't find '${valueAfters}'`);
      }
    } catch (e) {
      console.error(e.stack);
      showMessage({ style: "error", message: e.message });
      return false;
    }
  });
}

export async function check(object: dataValidationCheck) {
  try {
    let { config, onchangeconfig, value, row } = object;
    console.log(`on ${m}...`);
    console.log(onchangeconfig);
    console.log(value);
    const valueAfters = value.toString().split(onchangeconfig.separator ?? ",");
    const getvalueregex = "([\\w\\-]+)";
    const regex = RegExp(getvalueregex, "g");
    const checkHasGUID = regex.exec(value)[0];
    if (checkHasGUID != "") {
      return true;
    } else {
      if (valueAfters) {
        return await getdata(config, onchangeconfig, valueAfters[0], row);
      } else {
        console.log(`on ${m} empty...`);
        throw new Error(`Can't find '${valueAfters}'`);
      }
    }
  } catch (e) {
    console.error(e.stack);
    showMessage({ style: "error", message: e.message });
    return false;
  }
}

export async function getdata(config: FormConfig, dv_conf: onchange_dataValidation, valueAfters: any, row: number) {
  global.Callapiaction = {
    name: "callapiaction",
    action: {
      entitySet: dv_conf.EntityLogicalName + "s",
      queryString: "?" + dv_conf.queryString.replace("{valueAfters}", valueAfters),
      queryOptions: dv_conf.queryOptions,
    },
  };

  // RetrieveMultipleData(setData);
  let res = await RetrieveAndReturnMultipleData(null, {
    entitySet: dv_conf.EntityLogicalName + "s",
    queryString: "?" + dv_conf.queryString.replace("{valueAfters}", valueAfters),
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
        let com_confs = [];
        let specialCaseSkip = ["benz_productionyear","benz_applicablecontractdate","benz_licenseddate","benz_mbhksupportamount"];
        let currentRow = row - config.table.rowIndex - 1;
        for (let setFieldConf of sets) {
          // for (let { index, fieldname } of dv_conf.tablefield.map((fieldname, index) => ({ index, fieldname }))) {
          for (let ToLogicalName of setFieldConf.ToLogicalName) {
            let com_conf = config.table.columns.find((ele) => ele.LogicalName == ToLogicalName);
            if (specialCaseSkip.includes(ToLogicalName) && global.RowSpecialCase[currentRow] == "Yes") {
              console.log("special case skip validation");
              com_conf.cell.dataValidation.rule = {};
              com_conf.cell.values = "";

              com_conf.rowIndex = row;
              com_confs.push(com_conf);
              console.log(com_conf);

              continue;
            }
            console.log("setData com_conf");
            console.log(com_conf);
            com_conf.cell.values = "";
            if (setFieldConf.settype == "list") {
              let LogicalName = setFieldConf.FromLogicalNames[0];

              if(LogicalName == "benz_FinanceType.benz_name"){
                var measureName = res["benz_name"];
                var fsProductArr = await getRecordByUrl(global.ApiAccessToken,measureName);
                console.log("fsProductArr");
                console.log(fsProductArr);
                if (fsProductArr['value'][0]['benz_FinanceType'] != null) {
                  var fsProductNameArr = fsProductArr['value'][0]['benz_FinanceType']['benz_fsproduct_FinanceProduct_benz_financeproduct'];
                  console.log(fsProductNameArr);
                  console.log(fsProductNameArr.map((e) => e['benz_name']).join(","));
                  com_conf.cell.dataValidation.rule = {
                    "list": {
                      "inCellDropDown": true,
                      "source": fsProductNameArr.map((e) => e['benz_name']).join(",")
                    }
                  };
                }
              }else{
                console.log("setListData " + LogicalName);
                let keys = LogicalName.toString().split(".");
                console.log(keys);
                console.log(res);
                let arr = res[keys[0]];
                if (arr) {
                  console.log(arr);
                  if (!Array.isArray(arr)) {
                    if(arr == "ALL"){
                      com_conf.cell.dataValidation.rule = {};
                    }else{
                      let date = arr.split(",");
                      com_conf.cell.dataValidation.rule = {
                        "list": {
                          "inCellDropDown": true,
                          "source": arr
                        }
                      };
                      // com_conf.cell.values = date[0];
                    }
                  } else {
                    console.log(arr.map((e) => e[keys[1]]).join(","));
                    com_conf.cell.dataValidation.rule = {
                      "list": {
                        "inCellDropDown": true,
                        "source": arr.map((e) => e[keys[1]]).join(",")
                      }
                    };
                    // com_conf.cell.values = "";
                  }
                  console.log(com_conf);
                }
              }
            } else if (setFieldConf.settype == "date") {
              if (setFieldConf.FromLogicalNames.length == 2) {
                console.log("setDate " + ToLogicalName);
                console.log(setFieldConf);
                let formulaDate1 = res[setFieldConf.FromLogicalNames[0]];
                let formulaDate2 = res[setFieldConf.FromLogicalNames[1]];
                if(formulaDate1 != null){
                  com_conf.cell.dataValidation.rule = {
                    date: {
                      formula1: formatDate(formulaDate1),
                      formula2: formatDate(formulaDate2),
                      operator: setFieldConf.operator,
                    },
                  };
                } else {
                  if (setFieldConf.FromLogicalNames[0] == "benz_finalcontractdatefrom" || setFieldConf.FromLogicalNames[0] == "benz_finalapplicationdatefrom") {
                    if (formulaDate2 != null) {
                      com_conf.cell.dataValidation.rule = {
                        date: {
                          formula1: formatDate(formulaDate2),
                          operator: "LessThanOrEqualTo",
                        },
                      };
                    } else {
                      com_conf.cell.dataValidation.rule = {
                        date: {
                                operator: "GreaterThan",
                                formula1: "1/1/2000"
                        }
                      };
                    }
                  }
                }
              } else if (setFieldConf.FromLogicalNames.length == 1) {
                console.log("setDate " + ToLogicalName);
                console.log(setFieldConf);
                let formulaDate1 = res[setFieldConf.FromLogicalNames[0]];
                // let formulaDate2 = new Date();
                console.log(formulaDate1);
                if(formulaDate1 != null){
                  if (setFieldConf.operator == "LessThanOrEqualTo") {
                    com_conf.cell.dataValidation.rule = {
                      date: {
                        formula1: formatDate(formulaDate1),
                        operator: setFieldConf.operator,
                      },
                    };
                    // com_conf.cell.values = "";
                  }
                }
              } else {
                console.error("set date is empty");
              }
            } else if (setFieldConf.settype == "value") {
              let LogicalName = setFieldConf.FromLogicalNames[0];
              console.log("setData " + LogicalName);
              console.log(res[LogicalName]);
              com_conf.cell.values = res[LogicalName];
            } else if (setFieldConf.settype == "decimal") {
              let LogicalName = setFieldConf.FromLogicalNames[0];
              console.log("setData " + LogicalName);
              console.log(res[LogicalName]);

              const canExceed = res[setFieldConf.ExceedCheck];
              var formula1Set =
                res[setFieldConf.FromLogicalNames[0]] == null ? 0 : res[setFieldConf.FromLogicalNames[0]];
              let operatorSet = setFieldConf.operator;

              if (canExceed == 1) {
                formula1Set = 0;
                operatorSet = "GreaterThanOrEqualTo";
              }
              com_conf.cell.dataValidation.rule = {
                decimal: {
                  formula1: formula1Set,
                  operator: operatorSet,
                },
              };

              // console.log("global.CurrentConfig");
              // console.log(global.CurrentConfig);
              // let currentDatas = await get_table_data_arrary(global.CurrentConfig);
              // if (currentDatas.length > 0) {
              //   let currentRowIndex = com_conf.rowIndex - row;
              //   console.log(com_conf.rowIndex);
              //   console.log(row);
              //   console.log(currentRowIndex);
              //   let currentIndex = com_conf.index;
              //   console.log(currentIndex);
              //   let dataValue = currentDatas[currentRowIndex][currentIndex];
              //   console.log(dataValue);
              //   if (dataValue != "") {
              //     com_conf.cell.values = dataValue;
              //   }
              //   // com_conf.cell.values = res[LogicalName];
              // }
            }
            com_conf.rowIndex = row;
            com_confs.push(com_conf);
            console.log(com_conf);
          }
        }

        const fa = await lockmainObjectValueToCell(false);
        console.log(com_confs);
        await set_layout(null, com_confs);
        const fb = await lockmainObjectValueToCell(true);
        console.log("fa vs fb");
        console.log(fa);
        console.log(fb);
        // setEditedRange(args, true);
      } else {
        console.log("false result value");
        // setEditedRange(args, false);
        // throw new Error("No result");
      }
    } else {
      console.log("false result");
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

  return callRetrieveMultipleData(token,args);
}

export async function getvalue(dataset, config, rowIndex: number, columnIndex: number, valueafter) {
  let obj = {};
  console.log("rowIndex");
  console.log(rowIndex - config.table.rowIndex - 1);
  rowIndex = rowIndex - config.table.rowIndex - 1;
  console.log(dataset.items);
  // for (const parameter of parameters) {
  //   if (parameter == "{valueAfters}") {
  //     obj[parameter] = valueafter;
  //   }
  //   const col_config = config.table.columns.find((e) => e.LogicalName == parameter.replace("{", "").replace("}", ""));
  //   console.log(col_config);

    global.RowSpecialCase[rowIndex] = dataset.items[rowIndex]["values"][0][19];
  // }
}
