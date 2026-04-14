import * as BENZ from "../type";
import { Action } from "./factory";

export const action = async (config: BENZ.FormConfig) => new Action(config).create();
