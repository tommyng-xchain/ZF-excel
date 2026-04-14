/* global console, require */

import { query } from "../../type";
export async function init(args: query, obj) {
  console.log("on query...");
  var q = require(`./${args.type.toLowerCase()}`);
  return await q.init(args, obj);
}
