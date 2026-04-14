import * as config from "../config/table/model";
import "datatables.net-select-bs5";
import { get_choicesSets } from "./choices";
import { DropdownChoices, FormConfig, query, TableColumns } from "../type";
import * as datahelp from "../../dataverse-data-helper";
/* global global, console , Excel*/

global.tableConfigs["benz_prototypemodeldesignation"] = config.config;

export async function getSelectedModel(
  config: FormConfig,
  colconf: TableColumns,
  args: Excel.TableChangedEventArgs,
  m: string
) {
  return await Excel.run(async function (context) {
    try {
      console.log(`on ${m}...`);
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const onchange_conf = colconf.cell[args?.type ?? "TableChanged"].find(
        (e) => e.type.toLowerCase() === m.toLowerCase()
      );
      if (!onchange_conf) {
        new Error("Can't find onchange_conf");
      }
      // const valueAfters = args.details.valueAfter.toString();
      //   let all_ress: DropdownChoices[] = [];
      let vess: string[] = [];
      let pres: DropdownChoices[] = [];
      let obj = {
        address: colconf.address,
        row: colconf.rowIndex,
        model: pres,
        Inclusion: [],
        InclusionGroup: [],
        Exclusion: [],
        ExclusionGroup: [],
        all_ress: [],
      };
      //$expand=benz_ModelDesignation_ModelGroupM2M($select=benz_name)
      for (let queryConf of onchange_conf.querys) {
        console.log("queryConf");
        console.log(queryConf);
        console.log(config);
        const tcolconf = config.table.columns.find((e) => e.LogicalName === queryConf.LogicalName);
        console.log("tcolconf");
        console.log(tcolconf);
        if (tcolconf) {
          let trange = sheet.getRangeByIndexes(colconf.rowIndex, tcolconf.index, 1, 1);
          trange.load(["rowIndex", "columnIndex", "address", "values"]);
          await context.sync();

          // fomet the string to arrary string 'xxx','xxx'
          let valuestring = "";
          const valueAfters = trange.values.toString();
          if (valueAfters) {
            valuestring = "'" + valueAfters.split(",").join(`','`) + "'";
          }
          console.log("valueAfters");
          console.log(valueAfters);
          console.log(valuestring);

          if (valueAfters) {
            vess = vess.concat(valueAfters.split(","));

            let res: DropdownChoices[] = null;
            if (queryConf.type == "webapi") {
              res = await getdata(
                queryConf.EntityLogicalName + "s",
                "?" +
                  queryConf.queryString.replace("{valueAfters}", valuestring) +
                  (queryConf.expand
                    ? `&$expand=${queryConf.expand.EntityLogicalName}(${queryConf.expand.queryString})`
                    : ""),
                queryConf.queryOptions
              );
              res = processRes(res, queryConf);
            } else {
              res = get_choicesSets(queryConf.EntityLogicalName)?.filter((e) => valueAfters.split(",").includes(e.name));
            }
            console.log("res");
            console.log(queryConf.id);
            console.log(res);
            if(res){
              obj.all_ress = obj.all_ress.concat(res);
              console.log(obj.all_ress);
              obj[queryConf.id] = res;
              if (queryConf.add) {
                pres = pres.concat(res);
              } else {
                let map_res = res.map((e) => e.id);
                pres = pres.filter((e) => {
                  return !map_res.includes(e.id);
                });
              }
            }else{
              console.error(`Can't find ${valuestring} ...`);
            }
          } else {
            console.log(`on ${m} empty...`);
            // setData(null);
          }
        }
      }

      obj.model = pres;
      return obj;
    } catch (e) {
      console.error(e.stack);
    }
  });
}
export async function getSelectedModelbyValue(querys: query[], item) {
  try {
    console.log("getSelectedModelbyValue");
    console.log(item);
    if (!item) {
      return;
    }
    let vess: string[] = [];
    let pres: DropdownChoices[] = [];
    let obj = {
      model: pres,
      // Inclusion: [],
      // InclusionGroup: [],
      // Exclusion: [],
      // ExclusionGroup: [],
      all_ress: [],
    };
    //$expand=benz_ModelDesignation_ModelGroupM2M($select=benz_name)
    for (let queryConf of querys) {
      console.log("queryConf");
      console.log(queryConf);
      if (item[queryConf.followFieldLogicalName]) {
        // fomet the string to arrary string 'xxx','xxx'
        const valueAfters = item[queryConf.followFieldLogicalName].toString();
        if (!valueAfters) {
          break;
        }
        console.log("valueAfters", valueAfters, encodeURI(valueAfters));
        obj[queryConf.followFieldLogicalName] = valueAfters;
        if (valueAfters) {
          vess = vess.concat(valueAfters.split(","));
          let val = `'${valueAfters.split(",").join("','")}'`;
          let model: DropdownChoices[] = null;
          let group: DropdownChoices[] = null;
          let res = null;
          if (queryConf.type == "webapi") {
            let queryString =
              "?" +
              queryConf.queryString.replace("{valueAfters}", val) +
              (queryConf.expand
                ? `&$expand=${queryConf.expand.EntityLogicalName}(${queryConf.expand.queryString})`
                : "");
            res = await getdata(queryConf.EntityLogicalName + "s", queryString, queryConf.queryOptions);
            model = processRes(res, queryConf);
            group = processModelGroup(res, queryConf);
          } else if (queryConf.type == "local") {
            model = get_choicesSets(queryConf.EntityLogicalName).filter((e) => valueAfters.split(",").includes(e.name));
          } else {
            res = get_choicesSets(queryConf.EntityLogicalName).filter((e) => valueAfters.split(",").includes(e.name));
          }

          console.log("queryConf.id");
          console.log(queryConf.id);
          console.log("res");
          console.log(res);
          console.log(model);
          console.log(group);

          if (group) {
            obj[queryConf.id] = (obj[queryConf.id] ?? []).concat(group);
          }
          if (model) {
            obj[queryConf.id] = (obj[queryConf.id] ?? []).concat(model);
          }

          obj.all_ress = obj.all_ress.concat(model);
          console.log(obj.all_ress);
          // obj[queryConf.id] = res;
          if (queryConf.add) {
            pres = pres.concat(model);
          } else {
            let map_res = model.map((e) => e.id);
            pres = pres.filter((e) => {
              return !map_res.includes(e.id);
            });
          }
          console.log(obj);
        } else {
          // console.log(`on ${m} empty...`);
          // setData(null);
        }
      }
    }

    obj.model = pres;
    console.log("last obj");
    console.log(obj);
    return obj;
  } catch (e) {
    console.error(e.stack);
  }
}
function processRes(res, queryConf): DropdownChoices[] {
  let choi: DropdownChoices[] = [];
  if (res) {
    if (res.value) {
      if (queryConf.expand) {
        for (let val of res.value) {
          for (let sval of val[queryConf.expand.EntityLogicalName]) {
            choi.push({
              id: sval["benz_prototypemodeldesignationid"],
              name: sval["benz_name"],
              value: sval["benz_prototypemodeldesignationid"],
            });
          }
        }
      } else {
        for (let val of res.value) {
          choi.push({
            id: val["benz_prototypemodeldesignationid"],
            name: val["benz_name"],
            value: val["benz_prototypemodeldesignationid"],
          });
        }
      }
    }
  }
  return choi;
}

function processModelGroup(res, queryConf): DropdownChoices[] {
  let choi: DropdownChoices[] = [];
  if (res) {
    if (res.value) {
      if (queryConf.expand) {
        for (let val of res.value) {
          choi.push({
            id: val["benz_prototypemodelgroupid"],
            name: val["benz_name"],
            value: val["benz_prototypemodelgroupid"],
          });
        }
      }
    }
  }
  return choi;
}
async function getdata(entitySet: string, queryString: string = "", queryOptions: string = "") {
  global.Callapiaction = {
    name: "callapiaction",
    action: {
      entitySet: entitySet,
      queryString: queryString,
      queryOptions: queryOptions,
    },
  };

  return await datahelp.RetrieveAndReturnMultipleData(null, {
    entitySet: entitySet,
    queryString: queryString,
    queryOptions: queryOptions,
  });
}
