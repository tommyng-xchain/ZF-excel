/* global require */

var jq = require("jquery");

export const id = "#main-nav";

export function show() {
  jq(id).show();
}

export function hide() {
  jq(id).hide();
}
