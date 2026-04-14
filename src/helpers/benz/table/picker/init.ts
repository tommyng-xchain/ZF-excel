/* global console, require */

import { changeArgs, FormConfig } from "../../type";
import { getConfig } from "../init";
var jq = require("jquery");

export async function init(args: changeArgs, config: FormConfig) {
  try {
    console.log("on query...");
    if (args.type != "WorksheetSingleClicked") {
      return;
    }
    jq("#main li.nav-item").hide();
    jq(".tab-pane").removeClass("show");
    jq(".tab-pane").removeClass("active");
    jq("#main li.nav-item").addClass("hide");

    console.log("start Vaildtion");
    const colconf = await getConfig(args, config);
    console.log(colconf);
    const cs = colconf.cell[args.type];
    if (cs) {
      for (let c of cs) {
        const type = c.type;
        var dv = require(`./${type.toLowerCase()}`);
        await dv.init(args, config);
      }
    }
  } catch (e) {
    console.error(e.stack);
  }
}
