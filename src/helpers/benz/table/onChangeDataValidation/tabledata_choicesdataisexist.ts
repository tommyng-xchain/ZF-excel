/* global global, console, Excel */

import { dataValidationCheck, FormConfig } from "../../type";

let allTrue = (arr) => arr.every((v) => v === true);

const m = "tabledata_choicesdataisexist";

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
      range.load(["columnIndex", "address"]);
      await context.sync();
      console.log(config);

      const colconf = config.table.columns[range.columnIndex];
      const conf = colconf.cell[args.type].find((e) => e.type.toLowerCase() === m.toLowerCase());
      if (!conf) {
        throw new Error("Can't find conf");
      }
      const valueAfters = args.details.valueAfter.toString().split(conf.separator ?? ",");
      
      const tablefield = conf.tablefield;
      const isVaild = allTrue(
        valueAfters.map((v) =>
          tablefield.map((ele) => global.choicesSets[ele.split(".")[0]].map((v) => v.name).includes(v)).includes(true)
        )
      );
      return isVaild;
    } catch (e) {
      console.error(e.stack);
    }
  });
}
export async function check(object: dataValidationCheck) {
  try {
    let { onchangeconfig, value, fconfig } = object;
    console.log(`on ${m}...`);
    console.log(onchangeconfig);
    console.log(value);
    console.log(fconfig);
    if (!onchangeconfig) {
      throw new Error("Can't find conf");
    }
    const valueAfters = value.toString().split(onchangeconfig.separator ?? ",");

    const tablefield = onchangeconfig.tablefield;
    const isVaild = allTrue(
      valueAfters.map((v) =>
        tablefield.map((ele) => global.choicesSets[ele.split(".")[0]].map((v) => v.name).includes(v)).includes(true)
      )
    );
    return isVaild;
  } catch (e) {
    console.error(e.stack);
    return false;
  }
}
