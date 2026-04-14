/* global require */

var jq = require("jquery");
export function loadingspinner(c) {
  c ? jq("#loadingspinner").show() : jq("#loadingspinner").hide();
}
