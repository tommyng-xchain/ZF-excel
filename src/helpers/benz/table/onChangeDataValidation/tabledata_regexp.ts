// eslint-disable-next-line @typescript-eslint/no-unused-vars
/* global global, console, Excel */

import { dataValidationCheck, FormConfig } from "../../type";
import { allTrue } from "../init";

const m = "tabledata_regexp";

export async function init(args: Excel.TableChangedEventArgs, config: FormConfig): Promise<boolean> {
  return await Excel.run(async function (context) {
    try {
      console.log(`on ${m}...`);

      console.log(args);
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
      const re = new RegExp(dv_conf.regexp, "g");
      var vaild = re.test(valueAfters);
      console.log("regexp");
      console.log(vaild);
      if(vaild){
        if(valueAfters.length > 1){
          console.log("Check duplicate record");
          var stringArr = [];
          for(var i = 0; i < valueAfters.length; i++){
            if(!stringArr.includes(valueAfters[i])){
              stringArr.push(valueAfters[i]);
            }else{
              vaild = false;
              break;
            }
          }
        }
        console.log("Check duplicate result: " + vaild);
      }
      return vaild;
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
    var isVaild = allTrue(valueAfters.map((v) => new RegExp(onchangeconfig.regexp, "g").test(v)));
    return isVaild;
  } catch (e) {
    console.error(e.stack);
    return false;
  }
}
