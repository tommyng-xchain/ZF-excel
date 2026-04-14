/* global console, Excel  */

import { FormConfig } from "../../type";
import { showhide } from "../../components/systemusersearch";
import { getConfig } from "../init";
// import * as lookup from "../../../systemuserlookupdialog";
const m = "peoplepicker";

export async function init(args: Excel.WorksheetSingleClickedEventArgs, config: FormConfig): Promise<void> {
  try {
    console.log(`on ${m}...`, args);
    showhide(true, (await getConfig(args, config)).LogicalName);
  } catch (e) {
    console.error(e.stack);
  }
}
