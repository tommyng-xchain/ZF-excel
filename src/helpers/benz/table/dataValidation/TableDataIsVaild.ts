/* global global, console, require, Excel, document */

import { FormConfig } from "../../type";

export async function TableDataIsVaild(args: Excel.TableChangedEventArgs, config: FormConfig) {
    console.log("on TableDataIsVaild...");
    console.log(args);
    console.log(config);
  
  }