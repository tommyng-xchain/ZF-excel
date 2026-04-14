// import { FormConfig } from "../type";
// import { Action } from "./factory";
import * as BENZ from "../type";
import { Action } from "./factory";

// export const action = async (config: FormConfig) => new Action(config).update();
export const action = async (config: BENZ.FormConfig) => new Action(config).update();
