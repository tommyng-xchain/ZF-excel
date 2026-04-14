/* global global, console, Excel */

import { dataValidationCheck, FormConfig } from "../../type";
import { RetrieveAndReturnMultipleData} from "../../../dataverse-data-helper";

let allTrue = (arr) => arr.every((v) => v === true);

const m = "tabledata_choicesinselection";

export async function init(args: Excel.TableChangedEventArgs, config: FormConfig): Promise<boolean> {
  console.log(`on ${m}...`);
  console.log(args);
  console.log(args.details.valueAfter == "");
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
      const conf = colconf.cell[args.type].find((e) => e.type.toLowerCase() === m.toLowerCase());
      if (!conf) {
        throw new Error("Can't find conf");
      }
      const valueAfters = args.details.valueAfter.toString();
      var currentRow = range.rowIndex - config.table.rowIndex - 1;
      var measureID = global.RowMeasureName[currentRow];
      console.log(currentRow);
      console.log(global.RowMeasureName[currentRow]);
      console.log(conf);
      let result = await getdata(
        args,
        conf.EntityLogicalName + "s",
        "?" + conf.queryString.replace("{measureID}", measureID) + "&$count=true",
        ""
      );
      console.log("result");
      console.log(result);
      let keys = conf.EntityLogicalName1.toString().split(".");
      let res = result.value[0];
      let arr = res[keys[0]];
      console.log(arr);

      const fieldLogicalNameArr = conf.FieldLogicalNames;
      var isVaild = false;
      for(let i = 0; i < arr.length; i++){
        console.log(arr[i]);
        if(arr[i]['benz_name'] == valueAfters){
            isVaild = true;
            break;
        }
      }
      //model[ele.split(".")[0]].includes(v)).includes(true)
      return isVaild;
    } catch (e) {
      console.error(e.stack);
    }
  });
}
export async function check(object: dataValidationCheck) {
  try {
    let { onchangeconfig, value, fconfig, row } = object;
    console.log(`on ${m}...`);
    console.log(onchangeconfig);
    console.log(value);
    console.log(fconfig);
    if (!onchangeconfig) {
      throw new Error("Can't find conf");
    }
    const valueAfters = value.toString();
    var measureID = global.RowMeasureName[row];
    // console.log(currentRow);
    console.log(global.RowMeasureName[row]);
    console.log(onchangeconfig);
    console.log("valueAfters");
    console.log(valueAfters);
    let result = await getdata(
      null,
      onchangeconfig.EntityLogicalName + "s",
      "?" + onchangeconfig.queryString.replace("{measureID}", measureID) + "&$count=true",
      ""
    );
    console.log("result");
    console.log(result);
    let keys = onchangeconfig.EntityLogicalName1.toString().split(".");
    let res = result.value[0];
    let arr = res[keys[0]];
    console.log(arr);

    const fieldLogicalNameArr = onchangeconfig.FieldLogicalNames;
    const valueAfter = valueAfters.replace(onchangeconfig.tablefield[0] + "(", "").replace(")", "");
    var isVaild = false;
    console.log(valueAfter);
    for(let i = 0; i < arr.length; i++){
        for(const logicalName of fieldLogicalNameArr){
            if(arr[i][logicalName] == valueAfter){
                isVaild = true;
                break;
            }
        }
        if(isVaild){
            break;
        }
    }
    //model[ele.split(".")[0]].includes(v)).includes(true)
    return isVaild;
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