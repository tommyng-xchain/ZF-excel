/* global global, require, console, Excel */
import * as utils from "./utils";
import * as dataHelp from "../dataverse-data-helper";
import { action, FormConfig } from "./type";
import * as alertdialog from "../alertdialog";
import { loadingspinner } from "../loadingspinner-helper";
var jq = require("jquery");

export class FORM {
  public config: FormConfig;
  public module: string;
  public constructor(config: FormConfig) {
    this.config = config;
    this.module = this.config.module;
    this.init();
  }

  init() {
    try {
      global.form = { main: "", items: [], relationship_items: {}, object: {} };
      global.form.items = [];
      global.form.relationship_items = {};
      global.form.object = {};
      for (let ele of this.config.relationship) {
        global.form.relationship_items[ele.id] = [];
      }

      this.config["readonly"] = false;

      this.init_button();
      // init_forrmLayout();
    } catch (e) {
      console.log("Start_init_button_onclick error");
      console.log(e.message);
      console.error(e.stack);
    }
  }
  public async getconfig() {
    try {
      let sheetname = (await utils.getActiveWorksheet()).split("-");
      console.log(`${sheetname[0].toLocaleLowerCase()}`);
      return require(`./${sheetname[0].toLocaleLowerCase()}/${sheetname[1].toLocaleLowerCase()}/init`).config;
      // return this.merge_env_var_to_Conf(
      //   require(`./${sheetname[0].toLocaleLowerCase()}/${sheetname[1].toLocaleLowerCase()}/init`).config
      // );
    } catch (e) {
      console.error("Can't get config");
      console.error(e.stack);
      return null;
    }
  }
  public merge_env_var_to_Conf(config: FormConfig) {
    // global.EnvironmentVariable[vartype][form][field]

    try {
      const env_vars = global.EnvironmentVariable["formfieldvalue"][config.Table_Item_LogicalName];
      Object.typedKeys(env_vars).forEach((key) => {
        config.table.columns.forEach((field) => {
          if (field.LogicalName == key) {
            console.log(key.toString() + "is" + field.LogicalName);
            field.values = env_vars[key];
            field.defalutValues = env_vars[key];
          }
        });
      });
      return config;
    } catch (e) {
      console.error("Can't get config");
      console.error(e.stack);
      return null;
    }
  }
  public init_button() {
    console.log("init_button... " + this.module);
    try {
      utils.create_pageElement(
        "button",
        "module-nav",
        this.module,
        {
          id: this.module + "_start_button",
          className: `mod-${this.module} module btn btn-primary start m-1 col-12 col-lg-5`,
          textContent: `Start ${this.module}`,
        },
        this.Start_init_button_onclick
      );
      utils.create_pageElement(
        "button",
        "module-nav",
        this.module,
        {
          id: this.module + "_create_button",
          className: `mod-${this.module} module btn btn-primary m-1 col-5 col-lg-2`,
          textContent: "Create",
        },
        this.Create_init_button_onclick
      );

      utils.create_pageElement("input", "module-nav", this.module, {
        id: this.module + "_MomeID_input",
        type: "text",
        className: `mod-${this.module} input form-control m-1 col-5 col-lg-2`,
        textContent: "Memo id",
        placeholder: "Memo id",
      });

      utils.create_pageElement(
        "button",
        "module-nav",
        this.module,
        {
          id: this.module + "_Update_button",
          className: `mod-${this.module} module btn btn-primary m-1 col-5 col-lg-2`,
          textContent: "Update Form",
        },
        this.Update_init_button_onclick
      );
      utils.create_pageElement(
        "button",
        "module-nav",
        this.module,
        {
          id: this.module + "_GetData_button",
          className: `mod-${this.module} module btn btn-primary m-1 col-5 col-lg-2`,
          textContent: "Get Memo",
        },
        this.GetData_init_button_onclick
      );
      utils.create_pageElement(
        "button",
        "action-nav",
        null,
        {
          id: this.module + "_ImportData_button",
          className: `mod-${this.module} module action btn btn-success m-1 col-5 col-lg-2`,
          textContent: "Import",
        },
        this.post_Data
      );
      jq(`.mod-${this.module}`).hide();
      jq(`.start`).show();
    } catch (e) {
      console.log("Start_init_button_onclick error");
      console.log(e.message);
      console.error(e.stack);
    }
  }

  public Start_init_button_onclick(module) {
    try {
      console.log("Start_init_button_onclick " + module);
      global.CurrentConfig = global.modules[module].config;
      jq(`.mod-${module}`).show();
      jq(`.start`).hide();
      jq(`#action-nav`).hide();
    } catch (e) {
      console.log("Start_init_button_onclick error");
      console.log(e.message);
      console.error(e.stack);
    }
  }

  public async buildForm(module) {
    try {
      console.log("buildForm " + module);
      loadingspinner(true);

      await this.init_forrmLayout();
      await this.init_set_layout(this.config.layout.filter((ele) => ele.type !== "table_dataValidation"));

      // // check sheet has not data
      let currentDatas = await utils.get_table_data_arrary(this.config);
      console.log("currentDatas");
      console.log(currentDatas);

      // set table
      await this.init_set_table(this.config);
      if (currentDatas) {
        for (let i = 0; i < currentDatas.length; i++) {
          console.log("i");
          console.log(i);
          console.log(this.config.table.rowIndex + 1);
          console.log(currentDatas[i]);
          let cols = this.config.table.columns;
          let columnss = cols
            .filter((ele) => ele.enable)
            .sort((a, b) => a.index - b.index)
            .map((ele) => {
              ele.rowIndex = this.config.table.rowIndex + 1 + i;
              ele.cell.values = currentDatas[i][ele.index];
              // ele.index = ele.index + 1;
              return ele;
            });
          await this.init_set_layout(columnss);
        }
      } else {
        await this.init_set_layout(
          this.config.table.columns
            .filter((ele) => ele.enable)
            .sort((a, b) => a.index - b.index)
            .map((ele) => {
              ele.rowIndex = this.config.table.rowIndex + 1;
              // ele.index = ele.index + 1;
              return ele;
            })
        );
      }

      utils.addActionOnSelect(this.config);
      utils.addTableDataChange(this.config);

      jq(`#module-nav .mod-${module}`).hide();
      jq(`#action-nav .mod-${module}`).show();
      jq(`#action-nav`).show();
      // jq("#applyModel").show();
      jq("#main").show();
      jq("#createItem").show();
      jq("#addrow").show();
      jq("#add10row").show();

      // global.action_back = init_forrmLayout_action_back();
      // jq("#createItem").on("click", post_Data);
    } catch (e) {
      console.log("buildForm error");
      console.log(e.message);
      console.error(e.stack);
    }
    loadingspinner(false);
  }

  public async Create_init_button_onclick(module) {
    try {
      console.log("Create_init_button_onclick " + module);
      global.CurrentConfig = global.modules[module].config;
      // this.config["action"] = "create";
      // this.config["readonly"] = false;

      // this.config.module = module;
      let obj = new FORM(global.modules[module].config);
      obj.config["action"] = action.create;
      obj.config["readonly"] = false;
      await obj.buildForm(module);
    } catch (e) {
      console.log("Create_init_button_onclick error");
      console.log(e.message);
      console.error(e.stack);
    }
  }

  // public async Update_init_button_onclick(module) {
  //   try {
  //     console.log("Create_init_button_onclick " + module);
  //     global.CurrentConfig = global.modules[module].config;
  //     this.config["action"] = "create";
  //     this.config["readonly"] = false;

  //     this.config.module = module;
  //     await BENZ.Update_init_button_onclick(module);
  //   } catch (e) {
  //     console.log("Create_init_button_onclick error");
  //     console.log(e.message);
  //     console.error(e.stack);
  //   }
  // }

  public GetData_init_button_onclick(module) {
    try {
      console.log("GetData_init_button_onclick " + module);
      global.CurrentConfig = global.modules[module].config;

      console.log(jq(`#${module + "_MomeID_input"}`).val());
      console.log(!(jq(`#${module + "_MomeID_input"}`).val().length > 0));

      if (!(jq(`#${module + "_MomeID_input"}`).val().length > 0)) {
        return;
      } else {
        console.log("GetData_init_button_onclick " + module + ` :${jq(`#${module + "_MomeID_input"}`).val}`);
        this.config["readonly"] = true;
        this.Update_init_button_onclick(module);
      }
    } catch (e) {
      console.log("GetData_init_button_onclick error");
      console.log(e.message);
      console.error(e.stack);
    }
  }

  public async Update_init_button_onclick(module) {
    try {
      console.log("Update_init_button_onclick " + module);

      global.CurrentConfig = global.modules[module].config;

      console.log(jq(`#${module + "_MomeID_input"}`).val());
      console.log(!(jq(`#${module + "_MomeID_input"}`).val().length > 0));

      if (!(jq(`#${module + "_MomeID_input"}`).val().length > 0)) {
        return;
      } else {
        console.log("Update_init_button_onclick " + module + ` :${jq(`#${module + "_MomeID_input"}`).val()}`);
        // this.config["action"] = "update";

        // this.config["readonly"] = false;

        // this.config.module = module;
        let obj = new FORM(global.modules[module].config);
        obj.config["action"] = action.update;
        obj.config["readonly"] = false;
        await obj.Get_main_Data(jq(`#${module + "_MomeID_input"}`).val());
      }
    } catch (e) {
      console.log("Update_init_button_onclick error");
      console.log(e.message);
      console.error(e.stack);
    }
  }

  public async init_forrmLayout_action_back() {
    jq("#module-nav").show();
    jq(".module.btn").show();
    // jq("#applyModel").hide();
    jq("#main").hide();
    jq("#createItem").hide();
  }
  public async post_Data() {
    console.log(`on ${this.config.module} post data...`);
    try {
      console.log(await this.tableIsVliad());
      if (await this.tableIsVliad()) {
        let sheetname = (await utils.getActiveWorksheet()).split("-");
        console.log(`${sheetname[0].toLocaleLowerCase()}`);
        return require(`./${sheetname[0].toLocaleLowerCase()}/${this.config["action"]}`).action(await this.getconfig());
      }
    } catch (e) {
      console.error("Can't get post data");
      console.error(e.stack);
    }
  }

  public async Get_main_Data(id) {
    console.log(`on ${this.config.module} get main data...`);
    try {
      global.Callapiaction = {
        name: "callapiaction",
        action: {
          entitySet: this.config["Table_Main_LogicalName"] + "s",
          queryString: `$filter=contains(${this.config["Table_Main_id_LogicalName"]},'${id}'),accounts?$select=email&$expand=benz_prototypesalesmeasuremainform($select=benz_nameofpreparer;filter=contains(contains(${this.config["Table_Main_id_LogicalName"]},'${id}'))`,
          //queryString: `$filter=contains(${this.config["Table_Main_id_LogicalName"]},'${id}'),accounts?$select=email&$expand=benz_prototypesalesmeasuremainform($select=benz_nameofpreparer;filter=contains(contains(${this.config["Table_Main_id_LogicalName"]},'${id}'))`,
          queryOptions: "",
        },
      };
      //jq(`#${Module + "_MomeID_input"}`)
      const object = await dataHelp.RetrieveAndReturnMultipleData(null);
      if (object.value.length < 1) {
        throw new Error(`Can't not find Memo ID "${id}"`);
      }
      await utils.lockmainObjectValueToCell(false);
      await this.buildForm(this.config.module);

      await this.set__mainData(object);
      await utils.lockmainObjectValueToCell(true);

      console.log("object");
      console.log(object);
    } catch (e) {
      console.error(e.message);
      console.error(e.stack);
      // const id = "123456789"; // $("#notificationId").val().toString();
      // const details: Office.NotificationMessageDetails = {
      //   type: "errorMessage",
      //   message: e.message + " - " + id,
      // };
      // Office.context.mailbox.item.notificationMessages.addAsync(id, details, null);
      alertdialog.showdialog("Error", e.message);
    }
  }
  public async Get_items_Data(config, mainid) {
    console.log(`on ${config.module} get items data...`);
    try {
      global.Callapiaction = {
        name: "callapiaction",
        action: {
          entitySet: config["Table_Item_LogicalName"] + "s",
          queryString: `$filter=_${config["Table_Main_LogicalName"]}_value eq '${mainid}'`,
          queryOptions: "",
        },
      };
      console.log(global.Callapiaction);

      let obj = await dataHelp.RetrieveAndReturnMultipleData(null);
      await this.set__tableData(obj);

      console.log("object");
      console.log(obj);
    } catch (e) {
      console.log(e.message);
      console.error(e.stack);
    }
  }

  public async set__mainData(response) {
    console.log("set__mainData");
    console.log(response);
    const conf = this.config.layout.filter(
      (ele) => ele.type == "main-form-field" || ele.type == "main-form-field-gen" // && ele.action.includes("update")
    );
    const obj = response.value[0];
    if (obj) {
      await utils.mainObjectValueToCell(conf, obj);
      await utils.remove_tablerows(this.config);
      await this.Get_items_Data(this.config, obj[this.config["Table_Main_LogicalName"] + "id"]);
    }
  }

  public async set__tableData(response) {
    console.log("set__tableData");
    console.log(response);
    const conf = this.config.table.columns.filter((item) => item.enable);
    const arr = utils.ObjectToTableArray(conf, response.value);
    global.lastnumber = Math.max.apply(
      Math,
      response.value.map((o) => o[this.config.Table_linenumber_LogicalName])
    );

    console.log(arr);
    await utils.setDataToTable(this.config, arr);
    await utils.addTableDataChange(this.config);
  }

  public async tableIsVliad() {
    console.log("tableIsVliad");
    try {
      //   await fetch("./mbhk_config.json").then((response) => console.log(response.json()));
      //   let rawData = fs.readFileSync("./mbhk_config.json", "utf8");
      //   mbhk_config = JSON.parse(rawData);
      return await Excel.run(async function (context) {
        try {
          console.log("create sheet... ");
          let tableIsVliad = true;
          var sheet = context.workbook.worksheets.getActiveWorksheet();
          let table = sheet.tables.getItemAt(0);
          let range = table.getDataBodyRange();
          range.load(["values", "format/fill/color", "rowIndex"]);
          await context.sync();
          let rowIndex = range.rowIndex;
          let valuess = range.values;
          let ranges: Excel.Range[] = [];
          // for(let val of values){
          // }
          for (let y = 0; y < valuess.length; y++) {
            for (let x = 0; x < valuess[y].length; x++) {
              range = sheet.getRangeByIndexes(rowIndex + y, x, 1, 1);
              range.load(["values", "format/fill/color", "rowIndex"]);
              ranges.push(range);
            }
          }
          await context.sync();
          for (let nrange of ranges) {
            tableIsVliad = tableIsVliad && !(nrange.format.fill.color == "#FF0000");
          }
          return tableIsVliad;
        } catch (e) {
          console.log("tableIsVliad", e.message);
          console.log(e.stack);
        }
        console.log("end tableIsVliad");
      });
    } catch (e) {
      console.log("tableIsVliad", e.message);
      console.log(e.stack);
    }
  }
  public async init_forrmLayout() {
    console.log("init_forrmLayout");
    try {
      //   await fetch("./mbhk_config.json").then((response) => console.log(response.json()));

      //   let rawData = fs.readFileSync("./mbhk_config.json", "utf8");
      //   mbhk_config = JSON.parse(rawData);
      return await Excel.run(async function (context) {
        let sheet = null;

        try {
          console.log("create sheet... ");
          let sheets = context.workbook.worksheets;

          sheet = sheets.add(this.config.module);
          sheet.activate();
          return await context.sync().then(() => {
            return true;
          });
          // .then(() => {
          //   utils.set_layout(sheet, this.config.layout);
          // })
          // .then(() => {
          //   utils.set_table(sheet, this.config.table);
          // })
          // .then(() => {
          //   // utils.set_tableHeaders(sheet, this.config.table);
          // })
          // .then(() => {
          //   utils.set_tablebody(sheet, this.config.table);
          // });
        } catch (e) {
          console.log("get sheet");
          console.error(e.stack);
          sheet = context.workbook.worksheets.getItem(this.config.module);
          sheet.activate();
          await context.sync();
          // .then(() => {
          //   utils.set_layout(sheet, this.config.layout);
          // })
          // .then(() => {
          //   utils.set_table(sheet, this.config.table);
          // });
        }
        console.log("end get sheet... ");

        console.log("end init_forrmLayout");
      });
    } catch (e) {
      console.log("init_forrmLayout", e.message);
      console.log(e.stack);
    }
  }
  public async init_set_layout(elements) {
    console.log("init_set_layout");
    try {
      //   await fetch("./mbhk_config.json").then((response) => console.log(response.json()));

      //   let rawData = fs.readFileSync("./mbhk_config.json", "utf8");
      //   mbhk_config = JSON.parse(rawData);
      return await Excel.run(async function (context) {
        let sheet = null;

        try {
          let sheet = context.workbook.worksheets.getItem(this.config.module);

          return utils.set_layout(sheet, elements);
          // .then(() => {
          // })
          // .then(() => {
          //   utils.set_table(sheet, this.config.table);
          // })
          // .then(() => {
          //   // utils.set_tableHeaders(sheet, this.config.table);
          // })
          // .then(() => {
          //   utils.set_tablebody(sheet, this.config.table);
          // });
        } catch (e) {
          console.log("Error get sheet");
          console.error(e.stack);
          sheet = context.workbook.worksheets.getItem(this.config.module);
          sheet.activate();
          await context.sync();
          // .then(() => {
          //   utils.set_layout(sheet, this.config.layout);
          // })
          // .then(() => {
          //   utils.set_table(sheet, this.config.table);
          // });
        }
        console.log("end get sheet... ");

        console.log("end init_set_layout");
      });
    } catch (e) {
      console.log("init_set_layout", e.message);
      console.log(e.stack);
    }
  }

  public async init_set_table(config) {
    console.log("init_set_table");
    try {
      //   await fetch("./mbhk_config.json").then((response) => console.log(response.json()));

      //   let rawData = fs.readFileSync("./mbhk_config.json", "utf8");
      //   mbhk_config = JSON.parse(rawData);
      return await Excel.run(async function (context) {
        let sheet = null;

        try {
          let sheet = context.workbook.worksheets.getItem(config.module);

          return utils.set_table(sheet, config);
          // .then(() => {
          //   utils.set_layout(sheet, this.config.layout);
          // })
          // .then(() => {
          //   utils.set_table(sheet, this.config.table);
          // })
          // .then(() => {
          //   // utils.set_tableHeaders(sheet, this.config.table);
          // })
          // .then(() => {
          //   utils.set_tablebody(sheet, this.config.table);
          // });
        } catch (e) {
          console.log("Error get sheet");
          console.error(e.stack);
          sheet = context.workbook.worksheets.getItem(this.config.module);
          sheet.activate();
          await context.sync();
          //     .then(() => {
          //       utils.set_layout(sheet, this.config.layout);
          //     })
          //     .then(() => {
          //       utils.set_table(sheet, this.config.table);
          //     });
        }
        console.log("end get sheet... ");

        console.log("end init_set_table");
      });
    } catch (e) {
      console.log("init_set_table", e.message);
      console.log(e.stack);
    }
  }

  public async checkSheetTableHasData(config) {
    console.log("checkSheetTableHasData");
    try {
      //   await fetch("./mbhk_config.json").then((response) => console.log(response.json()));

      //   let rawData = fs.readFileSync("./mbhk_config.json", "utf8");
      //   mbhk_config = JSON.parse(rawData);
      return await Excel.run(async function (context) {
        let sheet: Excel.Worksheet = null;
        sheet = context.workbook.worksheets.getItem(config.module);

        let table = sheet.tables.getItemAt(0);
        let obj = table.toJSON();
        console.log(obj);
        return obj;
      });
    } catch (e) {
      console.error("checkSheetTableHasData", e.message);
      console.error(e.stack);
    }
  }
}
