/* global global, require */
import { FormConfig } from "../../type";
import * as Form from "../factory";

export const config: FormConfig = require("../../config/claim/mbfs.json");

// Office.onReady((info) => {
//   console.log(`onload Office Ready - claim-mbfs`);
//   if (info.host === Office.HostType.Excel) {
//     // init_SM();
//     global.modules[config.module] = config;
//     new CLAIM(config);
//   }
// });

export const Module = config.module;

export function init() {
  global.modules[config.module] = { config: config, object: new Form.CLAIM(config) };
}
