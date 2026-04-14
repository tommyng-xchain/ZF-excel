/* global Excel, console, require */

import { FormConfig } from "../../type";

var jq = require("jquery");

// let allTrue = (arr) => arr.every((v) => v === true);

const m = "modelpicker";

export async function init(args: Excel.TableChangedEventArgs, config: FormConfig) {
  console.log(`on ${m}...`, args, config);
  showhide(true);
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
