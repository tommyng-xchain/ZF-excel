/* global global, console, Excel, require */

import { getSelectedModel } from "../../components/carModel";
import { dataValidationCheck, DropdownChoices, FormConfig, TableColumns } from "../../type";
var jq = require("jquery");

// let allTrue = (arr) => arr.every((v) => v === true);

const m = "tabledata_model";

export async function init(args: Excel.TableChangedEventArgs, config: FormConfig) {
  showhide(true);
  let range: Excel.Range = null;
  console.log(`on ${m}...`);

  console.log(args);

  if (args.details.valueAfter == "") {
    return true;
  }
  try {
    await Excel.run(async function (context) {
      range = args.getRange(context);
      range.load(["rowIndex", "columnIndex", "address"]);
      await context.sync();
    });

    console.log(config);

    let colconf = config.table.columns[range.columnIndex];
    colconf.rowIndex = range.rowIndex;
    colconf.address = range.address;
    let onchange_conf = colconf.cell[args.type].find((e) => e.type.toLowerCase() === m.toLowerCase());
    if (!onchange_conf) {
      throw new Error("Can't find onchange_conf");
    }
    // const valueAfters = args.details.valueAfter.toString();
    // let ress: DropdownChoices[] = [];
    let pres: DropdownChoices[] = [];
    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    let obj = {
      address: range.address,
      row: range.rowIndex,
      model: pres,
      Inclusion: [],
      InclusionGroup: [],
      Exclusion: [],
      ExclusionGroup: [],
    };
    // await setEditedRange(args, await check(config, colconf, args));
    return await onchangecheck(config, colconf, args);
  } catch (e) {
    console.error(e.stack);
  }
}
export function showhide(conditional: boolean) {
  if (conditional) {
    jq(`#tab__model`).show();
    jq(`#tab__modelGrouph`).show();
    // jq("#modelgroups").show();
    // jq("#models").show();
    // jq("#tab__Table_id__model").trigger("click");
    jq("#models").addClass("show");
    jq("#models").addClass("active");
    // jq(`#applyModel`).show();
  } else {
    jq(`#tab__model`).hide();
    jq(`#tab__modelGrouph`).hide();
    jq("#models").removeClass("show");
    jq("#models").removeClass("active");
    jq("#modelgroups").removeClass("show");
    jq("#modelgroups").removeClass("active");
  }
}

export async function onchangecheck(config: FormConfig, colconf: TableColumns, args: Excel.TableChangedEventArgs) {
  try {
    console.log(`on ${m}...`);

    if (!config || !colconf) {
      throw new Error("Can't find conf");
    }
    // const valueAfters = args.details.valueAfter.toString();
    // let ress: DropdownChoices[] = [];
    let pres: DropdownChoices[] = [];
    let obj = await getSelectedModel(config, colconf, args, m);

    let oc = global.modulesOChange.find((e) => e.row == obj.row);
    oc ? (oc = obj) : global.modulesOChange.push(obj);
    console.log("global.modulesOChange");
    console.log(global.modulesOChange);
    console.log(pres);
    var valueArr = obj.all_ress.map((item) => {
      return item.name;
    });
    var isDuplicate = valueArr.every((item, idx) => valueArr.indexOf(item) != idx);
    if(valueArr.length == 0){
      isDuplicate = false;
    }
    
    console.log("end res...isDuplicate");
    console.log(valueArr);
    console.log(isDuplicate);

    return !isDuplicate; // && !visDuplicate;
  } catch (e) {
    console.error(e.stack);
  }
}

export async function check(object: dataValidationCheck) {
  try {
    let { fconfig, config } = object;
    console.log(`on ${m}...`);

    if (!config) {
      throw new Error("Can't find conf");
    }
    // const valueAfters = args.details.valueAfter.toString();
    // let ress: DropdownChoices[] = [];
    let pres: DropdownChoices[] = [];
    let obj = await getSelectedModel(config, fconfig, null, m);

    let oc = global.modulesOChange.find((e) => e.row == obj.row);
    oc ? (oc = obj) : global.modulesOChange.push(obj);
    console.log("global.modulesOChange");
    console.log(global.modulesOChange);
    console.log(pres);
    console.log(obj);
    var valueArr = obj.all_ress.map((item) => {
      return item.name;
    });
    var isDuplicate = valueArr.some((item, idx) => valueArr.indexOf(item) != idx);

    console.log("end res...isDuplicate");
    console.log(valueArr);
    console.log(isDuplicate);

    return !isDuplicate; // && !visDuplicate;
  } catch (e) {
    console.error(e.stack);
  }
}

// export async function check(config: FormConfig, colconf: TableColumns, args: Excel.TableChangedEventArgs) {
//   return await Excel.run(async function (context) {
//     try {
//       console.log(`on ${m}...`);
//       const sheet = context.workbook.worksheets.getActiveWorksheet();

//       // console.log(args);
//       // const sheet = context.workbook.worksheets.getActiveWorksheet();
//       // let range = args.getRange(context);
//       // range.load(["rowIndex", "columnIndex", "address"]);
//       // await context.sync();
//       // console.log(config);

//       // const colconf = config.table.columns[range.columnIndex];
//       const onchange_conf = colconf.cell[args.type].find((e) => e.type.toLowerCase() === m.toLowerCase());
//       if (!onchange_conf) {
//         new Error("Can't find onchange_conf");
//       }
//       // const valueAfters = args.details.valueAfter.toString();
//       let ress: DropdownChoices[] = [];
//       let vess: string[] = [];
//       let pres: DropdownChoices[] = [];
//       let obj = {
//         address: colconf.address,
//         row: colconf.rowIndex,
//         model: pres,
//         Inclusion: [],
//         InclusionGroup: [],
//         Exclusion: [],
//         ExclusionGroup: [],
//       };
//       //$expand=benz_ModelDesignation_ModelGroupM2M($select=benz_name)
//       for (let queryConf of onchange_conf.querys) {
//         console.log("queryConf");
//         console.log(queryConf);
//         const tcolconf = config.table.columns.find((e) => e.LogicalName === queryConf.LogicalName);
//         console.log("tcolconf");
//         console.log(tcolconf);
//         if (tcolconf) {
//           let trange = sheet.getRangeByIndexes(colconf.rowIndex, tcolconf.index, 1, 1);
//           trange.load(["rowIndex", "columnIndex", "address", "values"]);
//           await context.sync();

//           // fomet the string to arrary string 'xxx','xxx'
//           let valuestring = "";
//           const valueAfters = trange.values.toString();
//           if (valueAfters) {
//             valuestring = "'" + valueAfters.split(",").join(`','`) + "'";
//           }
//           console.log("valueAfters");
//           console.log(valueAfters);
//           console.log(valuestring);

//           if (valueAfters) {
//             vess = vess.concat(valueAfters.split(","));

//             let res: DropdownChoices[] = null;
//             if (queryConf.type == "webapi") {
//               res = await getdata(
//                 queryConf.EntityLogicalName,
//                 "?" +
//                   queryConf.queryString.replace("{valueAfters}", valuestring) +
//                   (queryConf.expand
//                     ? `&$expand=${queryConf.expand.EntityLogicalName}(${queryConf.expand.queryString})`
//                     : ""),
//                 queryConf.queryOptions
//               );
//               res = processRes(res, queryConf);
//             } else {
//               res = get_choicesSets(queryConf.EntityLogicalName).filter((e) => valueAfters.split(",").includes(e.name));
//             }
//             console.log("res");
//             console.log(queryConf.id);
//             console.log(res);
//             ress = ress.concat(res);
//             console.log(ress);
//             obj[queryConf.id] = res;
//             if (queryConf.add) {
//               pres = pres.concat(res);
//             } else {
//               let map_res = res.map((e) => e.id);
//               pres = pres.filter((e) => {
//                 return !map_res.includes(e.id);
//               });
//             }
//           } else {
//             console.log(`on ${m} empty...`);
//             // setData(null);
//           }
//         }
//       }

//       obj.model = pres;
//       let oc = global.modulesOChange.find((e) => e.row == obj.row);
//       oc ? (oc = obj) : global.modulesOChange.push(obj);
//       console.log("global.modulesOChange");
//       console.log(global.modulesOChange);
//       console.log(pres);
//       var valueArr = ress.map((item) => {
//         return item.name;
//       });
//       var isDuplicate = valueArr.some((item, idx) => valueArr.indexOf(item) != idx);
//       // eslint-disable-next-line @typescript-eslint/no-unused-vars
//       var visDuplicate = vess.some((item, idx) => vess.indexOf(item) != idx);

//       console.log("end res...isDuplicate");
//       console.log(valueArr);
//       console.log(isDuplicate);
//       // console.log(vess);
//       // console.log(visDuplicate);

//       return !isDuplicate; // && !visDuplicate;
//       // const tablefield = onchange_conf.tablefield;
//       // const isVaild = allTrue(
//       //   valueAfters.map((v) =>
//       //     tablefield.map((ele) => global.choicesSets[ele.split(".")[0]].map((v) => v.name).includes(v)).includes(true)
//       //   )
//       // );
//       // return isVaild;
//       // return true;
//     } catch (e) {
//       console.error(e.stack);
//     }
//   });
// }

// export async function setData(result): Promise<any> {
//   await Excel.run(async function (context) {
//     let args = global.onChangeDataValidation.args;
//     let config = global.onChangeDataValidation.config;
//     let range = args.getRange(context);
//     let dv_conf = null;
//     try {
//       console.log(`${m} setData`);
//       console.log(result);
//       // =INDIRECT("benz_prototypemodeldesignation")

//       range.load(["rowIndex", "columnIndex", "address"]);
//       await context.sync();

//       // rowindex
//       const colconf = config.table.columns[range.columnIndex];
//       dv_conf = colconf.cell[args.type].find((e) => e.type.toLowerCase() === m.toLowerCase());
//       let valueAfters = args.details.valueAfter;
//       console.log("valueAfters");
//       console.log(valueAfters);
//       valueAfters = valueAfters.toString().split(dv_conf.separator ?? ",");
//       console.log(valueAfters);

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
