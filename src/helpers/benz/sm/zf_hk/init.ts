/* global global, require */
import { FormConfig } from "../../type";
import * as Form from "../factory";

export const config: FormConfig = require("../../config/sm/zf_hk.json");

export const Module = config.module;
export function init() {
  global.modules[config.module] = { config: config, object: new Form.SM(config) };
}
