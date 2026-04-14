/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, require, console, Office */
import * as Benz from "../../dataverse-data-helper";
import DataTable from "datatables.net-searchpanes-bs5";
import "datatables.net-select-bs5";
import "./systemusersearch";
// import { set_dataChoices } from "./dataValidation";
import { DropdownChoices, initConfig } from "../type";
import * as Choices from "./choices";
import { loadingspinner } from "../../../helpers/loadingspinner-helper";
var jq = require("jquery");

export const config: initConfig = require("../config/init.json");
// import * as mode from "../../modeDialog";
import * as BenzType from "../../benz/type";

Office.onReady(async (info) => {
  console.log("onload Office Ready");
  if (info.host === Office.HostType.Excel) {
    // setApplyModel(false);
    console.log("get access token");

    // Benz.get_Data__benz_supporttypes(writeToOfficeDocument__benz_supporttypes);
    // loadingspinner(true);
    console.log("environment_name", global.environment_name);
    console.log("clientId", global.clientId);
    console.log("authority", global.authority);
    // await mode.showPopup(await getready_need_data);
    var config = require(`../../mode/${BenzType.AppMode.PRODUCTION}/app.json`);
    global.environment_name = config.environment_name;
    global.clientId = config.clientId;
    global.authority = config.authority;
    await getready_need_data();
    // loadingspinner(false);
  }
});

export async function getready_need_data() {
  console.log("environment_name", global.environment_name);
  console.log("clientId", global.clientId);
  console.log("authority", global.authority);

  await getready_choices().then((e) => {
    // if (global.mode) {
    jq("#main-nav").show();
    loadingspinner(false);
    // } else {
    //   mode.showPopup();
    // }
    config.forms.forEach((e) => {
      var form = require(`../${e.module}/${e.company}/init`);
      form.init();
    });
  });
  // get config to global.readyCallapiaction
  // await get_choices();
}

export async function getready_choices() {
  for (let ele of config.init_object) {
    const args = {
      name: "callapiaction",
      id: ele.key.toLocaleLowerCase(),
      type: ele.Type.toLocaleLowerCase(),
      oType: ele.oType.toLocaleLowerCase(),
      fields: ele.fields,
      entitySet: ele.entitySet,
      queryString: ele.queryString,
      queryOptions: ele.queryOptions,
    };
    const result = await Benz.RetrieveAndReturnMultipleData(getready_need_data, args);
    console.log("getready_choices", result);
    if (!result) {
      throw new Error("can't Retrieve choices Multiple Data");
    }
    await set_result(result, args);

    // global.readyCallapiaction.push({
    //   name: "callapiaction",
    //   id: ele.key.toLocaleLowerCase(),
    //   type: ele.Type.toLocaleLowerCase(),
    //   oType: ele.oType.toLocaleLowerCase(),
    //   fields: ele.fields,
    //   action: {
    //     entitySet: ele.entitySet,
    //     queryString: ele.queryString,
    //     queryOptions: ele.queryOptions,
    //   },
    // });
  }
}

export async function get_choices() {
  global.Callapiaction = global.readyCallapiaction[0];
  await Benz.RetrieveMultipleData(get_choices_callback);
}

export async function get_choices_callback(result: Object) {
  try {
    if (global.Callapiaction.type == "choices") {
      set_choices(global.Callapiaction, result);
    } else if (global.Callapiaction.type == "table") {
      set_table(global.Callapiaction, result);
    }
    global.readyCallapiaction = global.readyCallapiaction.slice(1);
    if (global.readyCallapiaction.length > 0) {
      get_choices();
    } else {
      // await getUserRoles();
      jq("#main-nav").show();
      loadingspinner(false);
    }
  } catch (e) {
    console.log(e.stack);
  }
}
export async function set_result(result: Object, args?) {
  try {
    const setting = args ?? global.Callapiaction;
    if (setting.type == "choices") {
      set_choices(setting, result);
    } else if (setting.type == "table") {
      set_table(setting, result);
    } else if (setting.type == "environmentvariablevalues") {
      set_environmentvariablevalues(setting, result);
    }
    if (!args) {
      global.readyCallapiaction = global.readyCallapiaction.slice(1);
      if (global.readyCallapiaction.length > 0) {
        get_choices();
      } else {
        // await getUserRoles();
        jq("#main-nav").show();
        loadingspinner(false);
      }
    }
  } catch (e) {
    console.log(e.stack);
  }
}
export async function set_environmentvariablevalues(Callapiaction, result) {
  try {
    console.log(Callapiaction.id + " set environmentvariablevalues...");
    // console.log(result);
    let data;
    if (result) {
      if (Callapiaction.oType == "environmentvariablevalues") {
        if (result.value) {
          for (let ele of result.value) {
            console.log(ele);
            // [envtype, vartype, form, field] = ele.schemaname.split("__");
            var path = ele.schemaname.split("__");
            path = path.slice(1);
            console.log(path);
            var last = path.pop();
            console.log(last);

            path.reduce(function (o, k) {
              o[k] = o[k] || {};

              return o[k];
            }, global.EnvironmentVariable)[last] = ele.value;
            // global.EnvironmentVariable[vartype][form][field] = ele.value;
          }
        }
      }
    }

    console.log("global.EnvironmentVariable", global.EnvironmentVariable);
  } catch (e) {
    console.log(e.stack);
  }
}

export async function set_choices(Callapiaction, result) {
  try {
    console.log(Callapiaction.id + " set_choices...");
    // console.log(result);
    let data: DropdownChoices[] = [];
    if (result) {
      if (Callapiaction.oType == "choices") {
        if (result.OptionSet) {
          for (let ele of result.OptionSet.Options) {
            data.push({
              id: ele.Label.UserLocalizedLabel.MetadataId,
              name: ele.Label.UserLocalizedLabel.Label,
              value: ele.Value,
            });
          }
        }
      } else if (Callapiaction.oType == "table") {
        result.value;
        if (result.value) {
          for (let ele of result.value) {
            data.push({
              id: ele[Callapiaction.fields[0]],
              name: ele[Callapiaction.fields[1]],
              value: ele[Callapiaction.fields[1]],
            });
            // console.log("value");
            // console.log(ele);
          }
        }
      }
    }
    Choices.set_choicesSets(Callapiaction.id, data);
    // console.log(Callapiaction.id + " data");
    // console.log(data.map((ele) => ele.name));
    Choices.set_dataChoices(
      Callapiaction.id,
      data.map((ele) => ele.name)
    );
    // console.log(global.choicesSets);
    // console.log(global.dataChoices);
  } catch (e) {
    console.log(e.stack);
  }
}

export async function set_table(Callapiaction, result) {
  try {
    console.log(Callapiaction.id + " set_table...");
    console.log(result);
    let data: DropdownChoices[] = [];
    let datatable_datas = [];
    const obj = config.init_object.find((i) => i.key.toLowerCase() === Callapiaction.id.toLowerCase());
    if (result) {
      result.value;
      if (result.value) {
        for (let ele of result.value) {
          // global.Callapiaction.id
          data.push({
            id: ele[Callapiaction.fields[0]],
            name: ele[Callapiaction.fields[1]],
            value: ele[Callapiaction.fields[1]],
          });
          datatable_datas.push(obj.columns.map((k) => ele[k]));
        }
      }
    }
    Choices.set_choicesSets(Callapiaction.id, data);
    console.log(global.choicesSets);
    console.log(datatable_datas);
    create_pageTable(Callapiaction.id.toLowerCase(), datatable_datas, obj.searchPanes);
  } catch (e) {
    console.log(e.stack);
  }
}

export async function create_pageTable(key: string, data, searchPanes: boolean): Promise<any> {
  try {
    console.log("create_pageTable " + key);
    console.log(global.tableConfigs);
    console.log(global.tableConfigs[key]);

    if (global.tableConfigs[key]) {
      global.tableConfigs[key].data = data;

      global.pagetable[key] = new DataTable(`#${key}`, global.tableConfigs[key]);
      if (searchPanes) {
        global.pagetable[key].searchPanes.container().prependTo(global.pagetable[key].table().container());
        global.pagetable[key].searchPanes.resizePanes();
      }
      console.log(`create_pageTable ${key} Done...`);
    }
  } catch (e) {
    console.log("create_pageTable EXCEPTION: " + e.message);
    console.log(e.stack);
  }
}
