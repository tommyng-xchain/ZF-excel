import { query } from "../../type";

const m = "setvalue";

export async function init(args: query, obj) {
  let valuse = args.valueset.map((key) => obj[key]);
  obj[args.settoLogicalName] = valuse.join("");
  return obj;
}
