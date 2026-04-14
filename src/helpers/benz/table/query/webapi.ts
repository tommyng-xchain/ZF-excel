/* global global, console */

import { RetrieveAndReturnMultipleData } from "../../../dataverse-data-helper";
import { get_choicesSets } from "../../components/choices";
import { query } from "../../type";

const m = "webapi";

export async function init(args: query, obj) {
  const str = args.queryString;
  const regexp = /\{([A-za-b0-9]+)\}/g;
  let strarrays = [...str.matchAll(regexp)];
  console.log(strarrays);
  for (let strarray of strarrays) {
    console.log(args.queryString);
    console.log(strarray[0]);
    console.log(strarray[1]);
    let val = null;
    if (strarray[1] == "benz_memotype") {
      val = get_choicesSets(strarray[1].toLowerCase())?.find((e) => e.name === obj[strarray[1]].toString())?.id;
    } else {
      val = obj[strarray[1]];
    }

    args.queryString = args.queryString.replace(strarray[0], val);
  }
  console.log("args.queryString");
  console.log(args.queryString);
  global.Callapiaction = {
    name: "callapiaction",
    action: {
      entitySet: args.entitySet,
      queryString: args.queryString,
      queryOptions: "",
    },
  };
  let res = await RetrieveAndReturnMultipleData(null, {
    entitySet: args.entitySet,
    queryString: args.queryString,
    queryOptions: "",
  });
  console.log("res");
  console.log(res);
  obj[args.settoLogicalName] = res[args.getApiValue];
  console.log("args.LogicalName");
  console.log(res[args.getApiValue]);
  console.log(obj);
  return obj;
}
