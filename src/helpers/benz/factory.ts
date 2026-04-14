/* global global, require, console, Excel,  */ //setTimeout
import * as utils from "./utils";
import * as dataHelp from "../dataverse-data-helper";
import { action, dataValidationCheck, FormConfig, FormLayout, TableColumns } from "./type";
import { get_choicesSets } from "./components/choices";
import * as tabledata from "./table/init";
import * as onchange_dataValidation from "./table/onChangeDataValidation/init";
import { showMessage } from "../message-helper";
import { getSelectedModelbyValue } from "./components/carModel";
import { loadingspinner } from "../loadingspinner-helper";
import { hideUnhideColumnAsync } from "./utils";
import * as tabledata_getdata from "./table/onChangeDataValidation/tabledata_getdata";
var jq = require("jquery");
type itemObject = {
  rowindex: number;
  object: any;
};
export class BENZ {
  Config: FormConfig;

  Module: string;

  permission: any[];

  start_button_permission: boolean;
  create_button_permission: boolean;
  Update_button_permission: boolean;
  GetData_button_permission: boolean;
  ImportData_button_permission: boolean;

  constructor(config) {
    this.Config = config;
    this.Module = this.Config.module;

    // const awaitTimeout = (delay) => new Promise((resolve) => setTimeout(resolve, delay));
    // awaitTimeout(5000).then(() => {
    //   if (global.PowerAccount?.roles.length > 0) {
    //     console.log("***awaitTimeout");
    //     console.log(this.start_button_permission);
    //     console.log(this.create_button_permission);
    //     console.log(this.Update_button_permission);
    //     console.log(this.GetData_button_permission);
    //     console.log(this.ImportData_button_permission);
    //     // this.setupPermissions();
    //   }
    // });
    // this.start_button_permission = true;
    // this.create_button_permission = true;
    // this.Update_button_permission = true;
    // this.GetData_button_permission = true;
    // this.ImportData_button_permission = true;
    this.setupPermissions();
    this.init();
  }

  setupPermissions() {
    console.log(global.PowerAccount.roles);
    console.log(this.Config.permissions?.start_button);
    this.start_button_permission = this.Config.permissions?.start_button
      .map((e) => e.toLocaleLowerCase())
      .some((item) => global.PowerAccount.roles.map((e) => e.toLocaleLowerCase()).includes(item));

    this.create_button_permission = this.Config.permissions?.create_button
      .map((e) => e.toLocaleLowerCase())
      .some((item) => global.PowerAccount.roles.map((e) => e.toLocaleLowerCase()).includes(item));

    this.Update_button_permission = this.Config.permissions?.Update_button.map((e) => e.toLocaleLowerCase()).some(
      (item) => global.PowerAccount.roles.map((e) => e.toLocaleLowerCase()).includes(item)
    );

    this.GetData_button_permission = this.Config.permissions?.GetData_button.map((e) => e.toLocaleLowerCase()).some(
      (item) => global.PowerAccount.roles.map((e) => e.toLocaleLowerCase()).includes(item)
    );

    this.ImportData_button_permission = this.Config.permissions?.ImportData_button.map((e) =>
      e.toLocaleLowerCase()
    ).some((item) => global.PowerAccount.roles.map((e) => e.toLocaleLowerCase()).includes(item));
  }

  init() {
    try {
      global.form = { main: "", items: [], relationship_items: {}, object: {} };
      global.form.items = [];
      global.form.relationship_items = {};
      global.form.object = {};
      for (let ele of this.Config.relationship) {
        global.form.relationship_items[ele.id] = [];
      }

      this.Config["readonly"] = false;

      this.init_button();
      // init_forrmLayout();
    } catch (e) {
      console.log("Start_init_button_onclick error");
      console.log(e.message);
      console.error(e.stack);
    }
  }
  static async getconfig() {
    try {
      let sheetname = (await utils.getActiveWorksheet()).split("-");
      console.log(`${sheetname[0].toLocaleLowerCase()}`);
      return require(`./${sheetname[0].toLocaleLowerCase()}/${sheetname[1].toLocaleLowerCase()}/init`).config;
    } catch (e) {
      console.error("Can't get config");
      console.error(e.stack);
      return null;
    }
  }

  init_button() {
    console.log("init_button... " + this.Module);
    global.isZFRecord = false;
    try {
      if (this.start_button_permission) {
        utils.create_pageElement(
          "button",
          "module-nav",
          this.Module,
          {
            id: this.Module + "_start_button",
            className: `mod-${this.Module} module btn btn-primary start m-1 col-12 col-lg-5`,
            textContent: `Start ${this.Module}`,
          },
          this.Start_init_button_onclick
        );
      }
      if (this.create_button_permission) {
        utils.create_pageElement(
          "button",
          "module-nav",
          this.Module,
          {
            id: this.Module + "_create_button",
            className: `mod-${this.Module} module btn btn-primary m-1 col-5 col-lg-2`,
            textContent: "Create",
          },
          this.Create_init_button_onclick
        );
      }
      if (this.Update_button_permission || this.GetData_button_permission) {
        var modelIdName = "";
        if (!this.Module.includes("CLAIM")) {
          modelIdName = "Memo id";
        } else {
          modelIdName = "Claim id";
        }
        utils.create_pageElement("input", "module-nav", this.Module, {
          id: this.Module + "_MomeID_input",
          type: "text",
          className: `mod-${this.Module} input form-control m-1 col-5 col-lg-2`,
          textContent: "Memo id",
          placeholder: modelIdName,
        });
      }
      if (this.Update_button_permission) {
        utils.create_pageElement(
          "button",
          "module-nav",
          this.Module,
          {
            id: this.Module + "_Update_button",
            className: `mod-${this.Module} module btn btn-primary m-1 col-5 col-lg-2`,
            textContent: "Update Form",
          },
          this.Update_init_button_onclick
        );
      }
      // if (this.GetData_button_permission) {
      //   utils.create_pageElement(
      //     "button",
      //     "module-nav",
      //     this.Module,
      //     {
      //       id: this.Module + "_GetData_button",
      //       className: `mod-${this.Module} module btn btn-primary m-1 col-5 col-lg-2`,
      //       textContent: "Get Memo",
      //     },
      //     this.GetData_init_button_onclick
      //   );
      // }
      if (this.ImportData_button_permission) {
        utils.create_pageElement(
          "button",
          "action-nav",
          null,
          {
            id: this.Module + "_ImportData_button",
            className: `mod-${this.Module} module action btn btn-success m-1 col-5 col-lg-2`,
            textContent: "Import",
          },
          this.post_Data
        );
      }
      jq(`.mod-${this.Module}`).hide();
      jq(`.start`).show();
    } catch (e) {
      console.log("Start_init_button_onclick error");
      console.log(e.message);
      console.error(e.stack);
    }
  }

  Start_init_button_onclick(module) {
    try {
      console.log("Start_init_button_onclick " + module);
      global.CurrentConfig = global.modules[module].config;
      // console.log("security check", global.PowerAccount.roles.map((e) => e.toLocaleLowerCase()).includes("benz_zfhk"));
      // if (global.PowerAccount.roles.map((e) => e.toLocaleLowerCase()).includes("benz_zfhk") && module == "CLAIM-MBFS") {
      //   global.isZFRecord = true;
      // }
      // console.log("global.isZFRecord");
      // console.log(global.isZFRecord);
      jq(`.mod-${module}`).show();
      jq(`.start`).hide();
      jq(`#action-nav`).hide();
    } catch (e) {
      console.log("Start_init_button_onclick error");
      console.log(e.message);
      console.error(e.stack);
    }
  }

  static async buildForm(module) {
    try {
      console.log("Create_init_button_onclick " + module);
      loadingspinner(true);

      await init_forrmLayout();
      await init_set_layout(global.CurrentConfig.layout.filter((ele) => ele.type !== "table_dataValidation"));

      // // check sheet has not data
      let currentDatas = await utils.get_table_data_arrary(global.CurrentConfig);
      console.log("currentDatas");
      console.log(currentDatas);

      if (currentDatas != null) {
        // let currentDataObj = await utils.get_table_data_obj(global.CurrentConfig);
        // let keyValues = await utils.ConvertToKey(currentDataObj, global.CurrentConfig);
        // only use measure id record to get data validation
        const gettableChange = global.CurrentConfig.table.columns[2].cell.TableChanged;
        console.log(gettableChange);
        const getdata_tableChanged = gettableChange[1];
        console.log(getdata_tableChanged);

        if (module.includes("CLAIM")) {
          await Excel.run(async function (context) {
            var sheet = context.workbook.worksheets.getActiveWorksheet();
            let table = sheet.tables.getItemAt(0);
            let range = table.getDataBodyRange();
            range.load(["rowIndex", "columnIndex"]);
            await context.sync();
            let rowIndex = range.rowIndex;
            let rowCount = 0;

            for (const data of currentDatas) {
              console.log(rowIndex);
              console.log(data);
              if (rowCount > 0) {
                rowIndex = rowIndex + 1;
              } else {
                rowIndex = rowIndex + rowCount;
              }

              console.log("new rowIndex");
              console.log(rowIndex);
              var measureName = data[2];
              var specialCase = data[19];
              global.RowMeasureName[rowCount] = measureName;
              global.RowSpecialCase[0] = specialCase;
              await tabledata_getdata.getdata(global.CurrentConfig, getdata_tableChanged, measureName, rowIndex);
              rowCount++;
            }
          });
        }
      }

      if (currentDatas != null) {
        for (let i = 0; i < global.CurrentConfig.table.columns.length; i++) {
          global.CurrentConfig.table.columns[i].cell.values = "";
        }
      }

      // set table
      await init_set_table(global.CurrentConfig);
      if (currentDatas) {
        for (let i = 0; i < currentDatas.length; i++) {
          console.log("i");
          console.log(i);
          console.log(global.CurrentConfig.table.rowIndex + 1);
          console.log(currentDatas[i]);
          let cols = global.CurrentConfig.table.columns;
          let columnss = cols
            .filter((ele) => ele.enable)
            .sort((a, b) => a.index - b.index)
            .map((ele) => {
              ele.rowIndex = global.CurrentConfig.table.rowIndex + 1 + i;
              ele.cell.values = currentDatas[i][ele.index];
              // ele.index = ele.index + 1;
              return ele;
            });
          await init_set_layout(columnss);
        }
      } else {
        await init_set_layout(
          global.CurrentConfig.table.columns
            .filter((ele) => ele.enable)
            .sort((a, b) => a.index - b.index)
            .map((ele) => {
              ele.rowIndex = global.CurrentConfig.table.rowIndex + 1;
              // ele.index = ele.index + 1;
              return ele;
            })
        );
      }

      utils.addActionOnSelect(global.CurrentConfig);
      utils.addTableDataChange(global.CurrentConfig);
      await hideUnhideColumnAsync(true);

      jq(`#module-nav .mod-${module}`).hide();
      jq(`#action-nav .mod-${module}`).show();
      jq(`#action-nav`).show();
      // jq("#applyModel").show();
      jq("#main").show();
      jq("#createItem").show();
      jq("#addrow").show();
      jq("#add10row").show();

      loadingspinner(false);

      // global.action_back = init_forrmLayout_action_back();
      // jq("#createItem").on("click", post_Data);
    } catch (e) {
      console.log("Create_init_button_onclick error");
      console.log(e.message);
      console.error(e.stack);
    }
    loadingspinner(false);
  }
  async Create_init_button_onclick(module) {
    try {
      console.log("Create_init_button_onclick " + module);
      global.CurrentConfig = global.modules[module].config;
      global.CurrentConfig.action = action.create;
      global.CurrentConfig["readonly"] = false;

      global.CurrentConfig.module = module;
      await utils.lockmainObjectValueToCell(false);
      await BENZ.buildForm(module);
      await utils.lockmainObjectValueToCell(true);
    } catch (e) {
      console.log("Create_init_button_onclick error");
      console.log(e.message);
      console.error(e.stack);
    }
  }
  async Update_init_button_onclick(module) {
    try {
      console.log("Create_init_button_onclick " + module);
      global.CurrentConfig = global.modules[module].config;
      global.CurrentConfig.action = action.update;
      global.CurrentConfig["readonly"] = false;

      global.CurrentConfig.module = module;
      await BENZ.Update_init_button_onclick(module);
    } catch (e) {
      console.log("Create_init_button_onclick error");
      console.log(e.message);
      console.error(e.stack);
    }
  }

  GetData_init_button_onclick(module) {
    try {
      console.log("GetData_init_button_onclick " + module);
      global.CurrentConfig = global.modules[module].config;

      console.log(jq(`#${module + "_MomeID_input"}`).val());
      console.log(!(jq(`#${module + "_MomeID_input"}`).val().length > 0));

      if (!(jq(`#${module + "_MomeID_input"}`).val().length > 0)) {
        return;
      } else {
        console.log("GetData_init_button_onclick " + module + ` :${jq(`#${module + "_MomeID_input"}`).val}`);
        global.CurrentConfig["readonly"] = true;
        BENZ.Update_init_button_onclick(module);
      }
    } catch (e) {
      console.log("GetData_init_button_onclick error");
      console.log(e.message);
      console.error(e.stack);
    }
  }

  static async Update_init_button_onclick(module) {
    try {
      console.log("Update_init_button_onclick " + module);

      global.CurrentConfig = global.modules[module].config;

      console.log(jq(`#${module + "_MomeID_input"}`).val());
      console.log(!(jq(`#${module + "_MomeID_input"}`).val().length > 0));

      if (!(jq(`#${module + "_MomeID_input"}`).val().length > 0)) {
        return;
      } else {
        console.log("Update_init_button_onclick " + module + ` :${jq(`#${module + "_MomeID_input"}`).val()}`);
        global.CurrentConfig.action = action.update;

        global.CurrentConfig["readonly"] = false;

        global.CurrentConfig.module = module;

        await Get_main_Data(jq(`#${module + "_MomeID_input"}`).val());
      }
    } catch (e) {
      console.log("Update_init_button_onclick error");
      console.log(e.message);
      console.error(e.stack);
    }
  }

  async init_forrmLayout_action_back() {
    jq("#module-nav").show();
    jq(".module.btn").show();
    // jq("#applyModel").hide();
    jq("#main").hide();
    jq("#createItem").hide();
  }
  async post_Data() {
    console.log(`on ${global.CurrentConfig.module} post data...`);
    try {
      console.log(await tableIsVliad());
      if (await tableIsVliad()) {
        // await post_main_Data();
        let sheetname = (await utils.getActiveWorksheet()).split("-");
        console.log(`${sheetname[0].toLocaleLowerCase()}`);
        return require(`./${sheetname[0].toLocaleLowerCase()}/${global.CurrentConfig["action"]}`).action(
          await BENZ.getconfig()
        );
      }
    } catch (e) {
      console.error("Can't get post data");
      console.error(e.stack);
    }
  }
}

export abstract class FormAction {
  name: string;
  config: FormConfig;
  mainObj: any;
  itemObjs: any[];
  mainObjOtherActions: any[];
  itemObjsOtherActions: any[];
  dataisValid: boolean;
  posted_mainData: any;
  posted_itemData: itemObject[];
  mainObjGen?: any;

  constructor(config: FormConfig) {
    this.config = config;
    this.mainObjOtherActions = [];
  }
  // abstract create(): void;

  // abstract update(): void;

  abstract beforePostData(): void;

  abstract afterPostData(): void;

  abstract _beforePostData(): void;

  async _afterPostData() {
    await utils.lockmainObjectValueToCell(false);
    if (this.posted_mainData) {
      await this.setidOnSheet(
        this.config,
        `${this.config.Table_Main_LogicalName}id`,
        this.posted_mainData[`${this.config.Table_Main_LogicalName}id`]
      );
    }

    if (this.posted_itemData) {
      for (let { rowindex, object } of this.posted_itemData) {
        await this.setitemidOnSheet(
          this.config,
          `${this.config.Table_Item_LogicalName}id`,
          object[`${this.config.Table_Item_LogicalName}id`],
          rowindex
        );
      }
    }
    const dateTime = new Date();
    const date = dateTime.toISOString().split('T')[0];
    const time = dateTime.toTimeString().split(' ')[0];
    const modifiedon = date.toString() + " " + time.toString();
    await this.setidOnSheet(this.config, "modifiedon", modifiedon);
    await hideUnhideColumnAsync(true);

    await utils.lockmainObjectValueToCell(true);
  }

  abstract newCommissionPostData(obj, row): void;

  async update() {
    console.log(`on ${this.config.module} update main data...`);
    loadingspinner(true);
    global.updating = true;
    try {
      this.mainObj = "";
      this.mainObjOtherActions = [];
      // get main data
      this.mainObj = await utils.get_main_data_by_object(
        this.config,
        this.config.layout.filter((ele) => ele.type == "main-form-field"),
        this.mainObjOtherActions
      );
      if (!this.mainObj) {
        throw new Error("Can't get form main data !");
      }
      // get table data
      this.itemObjs = await utils.getready_table_data_by_object(this.config, this.mainObj);

      if (!this.itemObjs) {
        throw new Error("Can't get form item data !");
      }

      console.log("mainObj 2");
      console.log(this.mainObj);
      console.log("itemObjs");
      console.log(this.itemObjs);

      console.log(global.updatingRecord);
      console.log(this.mainObj["benz_name"] == undefined);
      console.log(this.mainObj["benz_name"]);

      // if (global.isZFRecord) {
      //   this.mainObj[global.CurrentConfig["main_hardcode"][0]["LogicalName"]] = global.CurrentConfig["main_hardcode"][0]["ZFCompany"];
      // }

      if(this.mainObj["benz_name"] == undefined){
        if(global.updatingRecord == undefined || global.updatingRecord  == ""){
          this.mainObjGen = await utils.get_main_data_by_object(
            this.config,
            this.config.layout.filter((ele) => ele.type == "main-form-field-gen"),
            []
          );
          global.updatingRecord = this.mainObjGen["benz_name"];
        }
        this.mainObj["benz_name"] = global.updatingRecord;
      }
      console.log(this.mainObj);
      console.log(this.mainObj["benz_name"]);
      // if(global.updatingRecord){
      //   this.itemObjs.map((itemObj) => (itemObj["benz_name"] = global.updatingRecord + "-" + itemObj["benz_id"]));
      //   console.log("new this.itemObjs");
      //   console.log(this.itemObjs);
      // }

      if (this.mainObj["benz_name"] != "") {
        if (this.config.module.includes("SM")) {
          var args = {
            entitySet: "benz_prototypesalesmeasuremainforms",
            queryString: "?$select=benz_formstatus&$filter=benz_prototypesalesmeasuremainformid eq " + this.mainObj["benz_prototypesalesmeasuremainformid"],
            queryOptions: ""
          };
          
          var memoResult = await dataHelp.RetrieveAndReturnMultipleData(null, args);
          console.log("memoResult");
          console.log(memoResult);
          if (memoResult.value.length > 0) {
            var memoDetail = memoResult.value[0];
            const currentMainStatus = memoDetail["benz_formstatus"];
            if (currentMainStatus != 650230000 && currentMainStatus != 650230002 && currentMainStatus != null) {
              throw new Error("Memo Form \"" + this.mainObj["benz_name"] + "\" cannot be updated by current status - " 
                + memoDetail["benz_formstatus@OData.Community.Display.V1.FormattedValue"]);
            }
          }
        } else {
          console.log("item update checking");
          for (var itemDetail of this.itemObjs) {
            var currentClaimID = itemDetail["benz_prototypeclaimitemid"];

            var args = {
              entitySet: "benz_prototypeclaimitems",
              queryString: "?$select=benz_claimitemstatus,benz_claimapprovalstatus&$filter=benz_prototypeclaimitemid eq " + currentClaimID,
              queryOptions: ""
            };

            var claimResult = await dataHelp.RetrieveAndReturnMultipleData(null, args);
            console.log("claimResult");
            console.log(claimResult);
            if (claimResult.value.length > 0) {
              var claimDetail = claimResult.value[0];
              if (!(claimDetail["benz_claimapprovalstatus"] == 650230008 || claimDetail["benz_claimitemstatus"] == 650230000)) {
                throw new Error("Some claim item cannot update by items' current approval status, please use update form to get current form to update items");
              }
            }
          }
        }
      }

      let isValid = await this.dataValidation(
        this.mainObj,
        this.config,
        this.config.layout.filter((ele) => ele.type == "main-form-field"),
        -1
      );

      for (let itemObj of this.itemObjs) {
        let visv = await this.dataValidation(
          itemObj,
          this.config,
          this.config.table.columns,
          this.itemObjs.indexOf(itemObj)
        );
        // isValid = isValid && visv;
        // if (!isValid) {
        //   break;
        // }
      }

      this.posted_mainData = await utils.updateDataReturnData(
        this.config["Table_Main_LogicalName"] + "s",
        this.mainObj[this.config.Table_Main_LogicalName + "id"],
        this.mainObj
      );
      // if (!this.posted_mainData) {
      // global.form.main = this.posted_mainData;
      // global.form.relationship_items["main"].push(this.posted_mainData);
      // if (!this.posted_mainData) {
      //   throw new Error("Can't Main update Data");
      // }
      // if (!this.posted_mainData["benz_name"]) {
      //   throw new Error("Can't Main benz_name");
      // }

      // console.log("posted_mainData");
      // console.log(this.posted_mainData);

      await this._beforePostData();

      this.posted_itemData = [];
      var currentRow = 0;
      for (let itemObj of this.itemObjs) {
        var itemUID = itemObj[this.config.Table_Item_LogicalName + "id"];
        var companyLogicalName = global.CurrentConfig["main_hardcode"][0]["LogicalName"];
        if (itemUID !== null && itemUID !== undefined) {
          // if (global.isZFRecord) {
          //   itemObj[companyLogicalName] = global.CurrentConfig["main_hardcode"][0]["ZFCompany"];
          // }
          delete itemObj["benz_commissionnumber"];
          if (this.config.module.includes("CLAIM")) {
            delete itemObj["benz_claimtype"];
          }

          await utils.updateDataReturnData(
            this.config["Table_Item_LogicalName"] + "s",
            itemObj[this.config.Table_Item_LogicalName + "id"],
            itemObj
          );

          if (this.config.module.includes("SM")) {
            await this.post_remove_current_model(itemObj);
            for (let conf of this.config.relationship) {
              if (conf.id == "model") {
                const RelatedDataConf = await this.ready_model_related_Data(
                  conf,
                  itemObj
                );
                const res = await this.post_related_Data(RelatedDataConf);
                console.log("map related_Data res");
                console.log(res);
              }
            }
          }
        } else {
          console.log(itemObj);
          if (global.CurrentConfig["main_hardcode"]) {
            itemObj[companyLogicalName] = global.CurrentConfig["main_hardcode"][0]["value"];
            // if (global.isZFRecord) {
            //   itemObj[companyLogicalName] = global.CurrentConfig["main_hardcode"][0]["ZFCompany"];
            // }
          }
          if (this.config.module.includes("CLAIM")) {
            var newObj = await this.newCommissionPostData(itemObj, currentRow);
            console.log(newObj);
            itemObj = newObj;
          }

          for (const objName in itemObj) {
            if (itemObj[objName] == null) {
              delete itemObj[objName];
            }
          }
          this.posted_itemData.push({
            rowindex: this.itemObjs.indexOf(itemObj) + 1,
            object: await utils.postDataReturnData(this.config["Table_Item_LogicalName"] + "s", itemObj),
          });
        }
        currentRow++;
      }

      // console.log("posted_itemData");
      // console.log(this.posted_itemData);
      // global.form.items = this.posted_itemData;
      const RelatedDataConf = await this.readyRelatedData(
        this.config,
        this.mainObj,
        this.posted_itemData.map((e) => e.object)
      );
      const res = await this.post_related_Data(RelatedDataConf);
      console.log("map related_Data res");
      console.log(res);
      await this._afterPostData();
      // await this.afterPostData();
      showMessage({ style: "success", message: `${this.config.moduleTypeFullName} import update successfully !` });
    } catch (e) {
      console.error(e.message);
      console.error(e.stack);
      showMessage({ style: "error", message: "Error: " + e.message });
    }
    loadingspinner(false);
  }

  async create() {
    console.log(`on ${this.config.module} post main data...`);
    loadingspinner(true);
    global.updating = false;
    try {
      // get main data
      this.mainObj = await utils.get_main_data_by_object(
        this.config,
        this.config.layout.filter((ele) => ele.type == "main-form-field"),
        this.mainObjOtherActions
      );
      // console.log("global.isZFRecord");
      // console.log(global.isZFRecord);
      // if (global.isZFRecord) {
      //   this.mainObj[global.CurrentConfig["main_hardcode"][0]["LogicalName"]] = global.CurrentConfig["main_hardcode"][0]["ZFCompany"];
      // }
      if (!this.mainObj) {
        throw new Error("Can't get form data !");
      }
      // ***check is not has id will call update function to update recound
      if(this.mainObj[`${this.config.Table_Main_LogicalName}id`]){
        global.CurrentConfig.action = action.update;
        this.config.action = action.update;
        this.update();
      }
      // get table data
      this.itemObjs = await utils.getready_table_data_by_object(this.config, this.mainObj);

      if (!this.itemObjs) {
        throw new Error("Can't get form data !");
      }

      console.log("mainObj 2");
      console.log(this.mainObj);
      console.log("itemObjs");
      console.log(this.itemObjs);

      let isValid = await this.dataValidation(
        this.mainObj,
        this.config,
        this.config.layout.filter((ele) => ele.type == "main-form-field"),
        -1
      );
      for (let itemObj of this.itemObjs) {
        let visv = await this.dataValidation(
          itemObj,
          this.config,
          this.config.table.columns,
          this.itemObjs.indexOf(itemObj)
        );
        // isValid = isValid && visv;
        // if (!isValid) {
        //   break;
        // }
      }

      this.posted_mainData = await utils.postDataReturnData(this.config["Table_Main_LogicalName"] + "s", this.mainObj);
      global.form.main = this.posted_mainData;
      global.form.relationship_items["main"].push(this.posted_mainData);
      if (!this.posted_mainData) {
        throw new Error("Can't Main post Data");
      }
      if (!this.posted_mainData["benz_name"]) {
        throw new Error("Can't Main benz_name");
      }

      console.log("posted_mainData");
      console.log(this.posted_mainData);

      await this.beforePostData();

      this.posted_itemData = [];
      for (let itemObj of this.itemObjs) {
        if (global.CurrentConfig["main_hardcode"]) {
          var companyLogicalName = global.CurrentConfig["main_hardcode"][0]["LogicalName"];
          itemObj[companyLogicalName] = global.CurrentConfig["main_hardcode"][0]["value"];
          if (global.isZFRecord) {
            itemObj[companyLogicalName] = global.CurrentConfig["main_hardcode"][0]["ZFCompany"];
          }
        }
        console.log(itemObj);
        this.posted_itemData.push({
          rowindex: this.itemObjs.indexOf(itemObj) + 1,
          object: await utils.postDataReturnData(this.config["Table_Item_LogicalName"] + "s", itemObj),
        });
      }

      console.log("posted_itemData");
      console.log(this.posted_itemData);
      global.form.items = this.posted_itemData.map((e) => e.object);

      console.log(global.form);
      const RelatedDataConf = await this.readyRelatedData(
        this.config,
        this.posted_mainData,
        this.posted_itemData.map((e) => e.object)
      );
      const res = await this.post_related_Data(RelatedDataConf);
      console.log("map related_Data res");
      console.log(res);

      await this.afterPostData();
      await this._afterPostData();
      global.CurrentConfig.action = action.update;

      await showMessage({ style: "success", message: `${this.config.moduleTypeFullName} import successfully !` });
    } catch (e) {
      console.error(e.message);
      console.error(e.stack);
      showMessage({ style: "error", message: "Error: " + e.message });
    }
    loadingspinner(false);
  }

  async readyData() {
    console.log(`on ${this.config.module} post main data...`);
    try {
      // get main data
      this.mainObj = await utils.get_main_data_by_object(
        this.config,
        this.config.layout.filter((ele) => ele.type == "main-form-field"),
        this.mainObjOtherActions
      );
      if (!this.mainObj) {
        throw new Error("Can't get form data !");
      }
      // get table data
      this.itemObjs = await utils.getready_table_data_by_object(this.config, this.mainObj);

      if (!this.itemObjs) {
        throw new Error("Can't get form data !");
      }
      // mainObj = await getMemoNo(config, mainObj);

      console.log("mainObj 2");
      console.log(this.mainObj);
      console.log("itemObjs");
      console.log(this.itemObjs);
    } catch (e) {
      console.error(e.message);
      console.error(e.stack);
      showMessage({ style: "error", message: "Error: " + e.message });
    }
  }

  async isValid() {
    console.log(`on ${this.config.module} check data is Valid...`);
    try {
      this.dataisValid = await this.dataValidation(
        this.mainObj,
        this.config,
        this.config.layout.filter((ele) => ele.type == "main-form-field")
      );
      for (let itemObj of this.itemObjs) {
        this.dataisValid =
          this.dataisValid && (await this.dataValidation(itemObj, this.config, this.config.table.columns));
      }
      console.log("isValid");
      console.log(this.dataisValid);
      if (!this.dataisValid) {
        throw new Error("data not Validation ");
      }
    } catch (e) {
      console.error(e.message);
      console.error(e.stack);
      showMessage({ style: "error", message: "Error: " + e.message });
    }
  }

  // async post() {
  //   console.log(`on ${this.config.module} post main data...`);
  //   try {
  //     this.posted_mainData = await utils.postDataReturnData(this.config["Table_Main_LogicalName"] + "s", this.mainObj);
  //     global.form.relationship_items["main"].push(this.posted_mainData);
  //     if (!this.posted_mainData) {
  //       throw new Error("Can't Main post Data");
  //     }
  //     if (!this.posted_mainData["benz_name"]) {
  //       throw new Error("Can't Main benz_name");
  //     }

  //     console.log("posted_mainData");
  //     console.log(this.posted_mainData);

  //     this.beforeSetData();

  //     //post data
  //     this.posted_itemData = [];
  //     for (let itemObj of this.itemObjs) {
  //       this.posted_itemData.push(await utils.postDataReturnData(this.config["Table_Item_LogicalName"] + "s", itemObj));
  //     }

  //     console.log("posted_itemData");
  //     console.log(this.posted_itemData);
  //     global.form.items = this.posted_itemData;

  //     //Related Data
  //     console.log(global.form);
  //     const RelatedDataConf = await this.readyRelatedData(this.config, this.posted_mainData, this.posted_itemData);
  //     const res = this.post_related_Data(RelatedDataConf);
  //     console.log("map related_Data res");
  //     console.log(res);

  //     await this.afterSetData();
  //   } catch (e) {
  //     console.error(e.message);
  //     console.error(e.stack);
  //           showMessage({ style: "error", message: e.message });
  //   }
  // }

  async setidOnSheet(config: FormConfig, LogicalName: string, value: string) {
    try {
      let conf = config.layout.find((e) => e.LogicalName == LogicalName);
      console.log(conf);
      conf.cell.values = value;
      await utils.set_layout(null, [conf]);
    } catch (e) {
      console.error(e.stack);
    }
  }

  async setitemidOnSheet(config: FormConfig, LogicalName: string, value: string, rowIndex: number) {
    try {
      let conf = config.table.columns.find((e) => e.LogicalName == LogicalName);
      console.log(conf);
      conf.rowIndex = config.table.rowIndex + rowIndex;
      conf.cell.values = value;
      await utils.set_layout(null, [conf]);
    } catch (e) {
      console.error(e.stack);
    }
  }

  async dataValidation(obj, config, configs: FormLayout[] | TableColumns[], row?: number) {
    let ObjectIsVaild = {};
    try {
      global.CurrentLineNumber = 0;
      for (let target of configs) {
        console.log(target.LogicalName);
        console.log(row);
        console.log(obj);
        console.log(obj[target.LogicalName]);
        let value = obj[target.LogicalName] ?? "";

        if (row >= 0 && target.LogicalName == "benz_id") {
          global.CurrentLineNumber = obj["benz_id"];
        }

        // get value for lookup field
        if (target.AttributeType.toString().toLowerCase() == "lookup" && target.LookupAllowUsingExist) {
          console.log("lookup @odata.bind");
          console.log(obj);
          console.log(`${target.SchemaName}@odata.bind`);

          value = obj[`${target.SchemaName}@odata.bind`];
        }

        let specialcase = false;

        if (target.type != "main-form-field" && config?.specialcase) {
          const specialcase_field = config.table.columns.find((e) => e.LogicalName == config?.specialcase?.target);
          console.log("specialcase");
          console.log("specialcase_field", specialcase_field);
          specialcase = config.specialcase.available;
        }

        let isMandatoryVaild = false;
        const mandatory = target.LogicalName == config?.specialcase?.target ? false : target.mandatory;
        if (target.type != "main-form-field") {
          specialcase == true && target.specialcase == true && target.LogicalName == config.specialcase.target
            ? false
            : target.mandatory;
        }
        
        //.length !== 0
        console.log("mandatory", mandatory);
        console.log("value", value);
        // if (mandatory && (target.AttributeType == "Integer" || target.AttributeType == "Number") && value === "") {
        //   value = 0;
        // }

        if (mandatory && (value === undefined || value === "")) {
          isMandatoryVaild = false;
          if (row == -1 || !(target.AttributeType.toString().toLowerCase() == "lookup")) {
            throw new Error(`"${target.Label ? target.Label : target.id}" field is not Valid!`);
          }
        } else {
          isMandatoryVaild = true;
        }

        if (value || value != "") {
          if (target.cell.dataValidation?.target && target.AttributeType.toString().toLowerCase() == "choice") {
            console.log(get_choicesSets(target.cell.dataValidation.target.toLowerCase()));
            console.log(get_choicesSets(target.cell.dataValidation.target.toLowerCase()).find((e) => e.value == value));
            ObjectIsVaild[target.LogicalName] =
              isMandatoryVaild &&
              get_choicesSets(target.cell.dataValidation.target.toLowerCase()).find((e) => e.value == value)
                ? true
                : false;
          } else if (
            (target.LookupInoutType ?? "").toString().toLowerCase() == "choice" &&
            target.AttributeType.toString().toLowerCase() == "lookup"
          ) {
            // value = obj[`${target.SchemaName}@odata.bind`]
            //   .toString()
            //   ?.replace(
            //     `${config.relationship.find((e) => e.from_FieldLogicalName === target.LogicalName).to_LogicalName}s(`,
            //     ""
            //   )
            //   .replace(")", "");
            const getvalueregex = "([\\w\\-]+)";
            const regex = RegExp(getvalueregex, "g");
            const schemaName = regex.exec(obj[`${target.SchemaName}@odata.bind`])[0];

            console.log(get_choicesSets(target.cell.dataValidation.target.toLowerCase()));
            console.log(
              get_choicesSets(target.cell.dataValidation.target.toLowerCase())?.find((e) => e.id === schemaName.toString())
            );
            ObjectIsVaild[target.LogicalName] =
              isMandatoryVaild &&
              get_choicesSets(target.cell.dataValidation.target.toLowerCase())?.find((e) => e.id === schemaName.toString())
                ? true
                : false;
          } else if (
            (target.LookupInoutType ?? "").toString().toLowerCase() == "string" &&
            target.AttributeType.toString().toLowerCase() == "lookup"
          ) {
            console.log("Lookup...");

            // let id = null;
            // try {
            //   ObjectIsVaild[target.LogicalName] =
            //     ObjectIsVaild[target.LogicalName] &&
            //     get_choicesSets(target.LogicalName)?.find((e) => e.value === value.toString())
            //       ? true
            //       : false;
            // } catch (e) {
            //   console.error("Error lookup");
            // }

            // if (target.outOfTarget) {
            //   ObjectIsVaild[target.outOfTarget.targetLogicalName] = obj[target.outOfTarget.targetLogicalName]
            //     ? true
            //     : false;
            // }

            // value = obj[`${target.SchemaName}@odata.bind`]
            //   .replace(
            //     `${config.relationship.find((e) => e.from_FieldLogicalName === target.LogicalName).to_LogicalName}s(`,
            //     ""
            //   )
            //   .replace(")", "");
            const getvalueregex = "([\\w\\-]+)";
            const regex = RegExp(getvalueregex, "g");
            const schemaName = regex.exec(obj[`${target.SchemaName}@odata.bind`])[0];

            console.log(get_choicesSets(target.cell.dataValidation.target.toLowerCase()));
            console.log(
              get_choicesSets(target.cell.dataValidation.target.toLowerCase())?.find((e) => e.id === schemaName.toString())
            );
            ObjectIsVaild[target.LogicalName] =
              isMandatoryVaild &&
              get_choicesSets(target.cell.dataValidation.target.toLowerCase())?.find((e) => e.id === schemaName.toString())
                ? true
                : false;
          } // if (target.AttributeType.toString().toLowerCase() == "Choice") {
          // } else
          else if (target.AttributeType.toString().toLowerCase() == "string") {
            ObjectIsVaild[target.LogicalName] = isMandatoryVaild && value ? true : false;
          } else if (target.AttributeType.toString().toLowerCase() == "boolean") {
            console.log(target.cell.valuesMap);
            console.log(Object.values(target.cell.valuesMap));
            ObjectIsVaild[target.LogicalName] =
              isMandatoryVaild && Object.values(target.cell.valuesMap).includes(value) ? true : false;
            // nitem[target.LogicalName] = value.toString().toLowerCase() == "yes" ? true : false;
          } else if (target.AttributeType.toString().toLowerCase() == "number") {
            ObjectIsVaild[target.LogicalName] = isMandatoryVaild && (Number(value ?? 0) ? true : false);
          } else if (
            target.AttributeType.toString().toLowerCase() == "datetime" ||
            target.AttributeType.toString().toLowerCase() == "date"
          ) {
            ObjectIsVaild[target.LogicalName] = isMandatoryVaild && (new Date(value) ? true : false);
          } else {
            ObjectIsVaild[target.LogicalName] = isMandatoryVaild && value ? true : false;
          }

          if (mandatory && (value === undefined || value === "")) {
            throw new Error(`"${target.Label}" field is not Valid!`);
          }
        }
        console.log(`${target.LogicalName} : ${ObjectIsVaild[target.LogicalName]}`);

        if(target.needFixedValue && value.toString() != target.fixedValue){
          throw new Error(`"${target.Label}" field is not Valid!`);
        }

        if ((target.cell.TableChanged && value) || target.cell.fouceCheckChange) {
          for (let TableChanged of target.cell.TableChanged) {
            console.log("special case and pass check");
            console.log(obj[config.specialcase.target]);
            console.log(target.canPass);
            if (!(obj[config.specialcase.target] && target.canPass)) {
              let tocheck: dataValidationCheck = {
                config: config,
                onchangeconfig: TableChanged,
                value: value,
                fconfig: target,
                row,
              };
              ObjectIsVaild[target.LogicalName] = isMandatoryVaild && (await onchange_dataValidation.check(tocheck));
              if (!ObjectIsVaild[target.LogicalName]) {
                // throw new Error((row ? "Row " + row.toString() + " " : "") + `"${target.Label}" field is not Valid!`);
                var errorMsg = `"${target.Label}" field is not Valid!`;
                if (global.ErrorMsg != "") {
                  errorMsg = global.ErrorMsg;
                }
                throw new Error(errorMsg);
              }
            }
          }
        }

        console.log(ObjectIsVaild[target.LogicalName]);
        console.log(value);
        console.log(target.AttributeType.toString().toLowerCase());
        if (!ObjectIsVaild[target.LogicalName] && mandatory && row != -1 && (value == undefined || value == "" || value == null) 
          && target.AttributeType.toString().toLowerCase() == "lookup" ) {
          throw new Error(`"${target.Label}" field is not Valid!`);
        }

        console.log(`${target.LogicalName} : ${ObjectIsVaild[target.LogicalName]}`);
        if (target.LogicalName == "benz_prototypesalesmeasure") {
          global.RowMeasureID[row] = value;
        }

        if (target.cell.checkSubmitTo) {
          console.log("checkSubmitTo");
          console.log(this.mainObj.benz_submitto);
          console.log(value);
          if (this.mainObj.benz_submitto == 650230001) {
            if (value == undefined) {
              throw new Error(`"${target.Label}" field is not Valid!`);
            }
          }
        }
      }
      console.log("ObjectIsVaild");
      console.log(ObjectIsVaild);
      console.log(Object.values(ObjectIsVaild));
      console.log(tabledata.allTrue(Object.values(ObjectIsVaild)));
      return tabledata.allTrue(Object.values(ObjectIsVaild));
    } catch (e) {
      global.ErrorMsg = "";
      console.error(e.stack);
      throw new Error((row == -1 ? "" : "Row " + (row + 1).toString() + " ") + e.message);
      // return false;
    }
  }

  async post_related_Data(conf) {
    let res = [];
    if (conf.length > 0) {
      for (let c of conf) {
        res.push(await dataHelp.post_MapAssociate(c));
      }
    } else {
      return null;
    }
    return res;
  }
  MapAssociate_callback(response) {
    console.log(`${global.CurrentConfig.module} MapAssociate callback`);
    console.log(response);
  }
  async readyRelatedData(config, main, items) {
    try {
      console.log(`on ${config.module} post related_Data...`);
      console.log(main);
      console.log(items);
      let perpostrelationship_items = [];
      for (let item of items) {
          console.log(item);
        if(!item){
          console.error(`the item is null`);
        }else{
        for (let conf of config.relationship) {
          // if (conf.id === "systemuser") {
          //   await getready_systemuser_related_api_action(conf.filter((e) => e.to_LogicalName === "systemuser"));
          // } else
          if (conf.id === "main") {
            perpostrelationship_items.push(await this.ready_main_related_api_action(conf, main, item));
          } else if (conf.id === "model") {
            // await getready_model_related_Data(conf);
            let r = await this.ready_model_related_Data(conf, item);
            for (let x of r) {
              perpostrelationship_items.push(x);
            }
            // perpostrelationship_items.concat(r);
            // await this.ready_model_related_Data(conf, item);
          }
          //  else if (conf.id === "exmodelgroup") {
          //   await getready_modelGroup_related_Data(conf);
          // }
        }
      }
      }
      console.log("perpostrelationship_items");
      console.log(perpostrelationship_items);
      return perpostrelationship_items;
    } catch (e) {
      console.error(e.stack);
      console.log(e.message);
      console.error(e.stack);
    }
  }

  async ready_main_related_api_action(config, mainobj, item) {
    console.log("ready_main_related_api_action");
    console.log(mainobj);
    console.log(item);
    if (!item || !mainobj) {
      return;
    }
    return {
      Callapiaction: {
        name: "callapiaction",
        action: {
          entitySet: config.from_LogicalName + "s",
          id: item[config.from_LogicalName + "id"],
          relationship: config.relationship_LogicalName,
          relatedEntitySet: config.to_LogicalName + "s",
          relatedEntityId: mainobj[config.to_LogicalName + "id"],
          queryOptions: "",
        },
      },
    };
  }

  async ready_model_related_Data(conf, item) {
    try {
      console.log(`on ready model related_Data...`);
      console.log(item);
      console.log(conf);
      if(!item){
        console.error("Can't get item data!");
        return [];
      }
      const car = await getSelectedModelbyValue(conf.querys, item);
      console.log("ready model related_Data car");
      console.log(car);
      let p = [];
      if(car){

        for (const obj of car.model) {
          try {
            // const obj = global.CarModels.find((el) => el.name.toLowerCase().trim() === modelname.toLowerCase().trim());
            if (obj.id) {
              let c = {
                Callapiaction: {
                  name: "callapiaction",
                  action: {
                    entitySet: conf.from_LogicalName + "s",
                    id: item[conf.from_LogicalName + "id"],
                    relationship: conf.relationship_LogicalName,
                    relatedEntitySet: conf.to_LogicalName + "s",
                    relatedEntityId: obj.id,
                    queryOptions: "",
                  },
                },
              };
              console.log(c);
              p.push(c);
            }
          } catch (e) {
            console.log(e.message);
            console.error(e.stack);
          }
        }
        for (const key of Object.keys(car)) {
          const qconf = conf.querys.find((e) => e.id == key);
          if (qconf) {
            for (const obj of car[key]) {
              try {
                // const obj = global.CarModels.find((el) => el.name.toLowerCase().trim() === modelname.toLowerCase().trim());
                if (obj.id) {
                  let c = {
                    Callapiaction: {
                      name: "callapiaction",
                      action: {
                        entitySet: qconf.from_LogicalName + "s",
                        id: item[qconf.from_LogicalName + "id"],
                        relationship: qconf.relationship_LogicalName,
                        relatedEntitySet: qconf.to_LogicalName + "s",
                        relatedEntityId: obj.id,
                        queryOptions: "",
                      },
                    },
                  };
                  console.log(c);
                  p.push(c);
                }
              } catch (e) {
                console.log(e.message);
                console.error(e.stack);
              }
            }
          }
        }
      }
      return p;
    } catch (e) {
      console.log(e.message);
      console.error(e.stack);
    }
  }

  async post_remove_current_model(item) {
    console.log(item);
    var args = {
      entitySet: "benz_prototypesalesmeasures",
      queryString: "?$filter=benz_prototypesalesmeasureid eq '" + item["benz_prototypesalesmeasureid"] + "'&$expand=benz_m2m_SalesMeasure_CarModel_result",
      queryOptions: ""
    };
  
    var measureResult = await dataHelp.RetrieveAndReturnMultipleData(null, args);
    console.log(measureResult);
    var currentModel = measureResult.value[0]["benz_m2m_SalesMeasure_CarModel_result"];

    console.log("current model result");
    console.log(currentModel);
    for (const model of currentModel) {
      console.log(model["benz_name"]);
      const modelID = model["benz_prototypemodeldesignationid"];
      global.Callapiaction = {
        name: "callapiaction",
        action: {
          entitySet: "benz_prototypesalesmeasures",
          id: item["benz_prototypesalesmeasureid"],
          property: "benz_m2m_SalesMeasure_CarModel_result",
          relatedEntityId: modelID
        },
      };
  
      let c = {
        Callapiaction: {
          name: "callapiaction",
          action: {
            entitySet: "benz_prototypesalesmeasures",
            id: item["benz_prototypesalesmeasureid"],
            property: "benz_m2m_SalesMeasure_CarModel_result",
            relatedEntityId: modelID
          },
        },
      };
      await dataHelp.post_Disassociate(c);
    }
  }
}

export async function tableIsVliad() {
  console.log("tableIsVliad");
  try {
    //   await fetch("./mbhk_config.json").then((response) => console.log(response.json()));

    //   let rawData = fs.readFileSync("./mbhk_config.json", "utf8");
    //   mbhk_config = JSON.parse(rawData);
    return await Excel.run(async function (context) {
      try {
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
            range.load(["values", "format/fill/color", "rowIndex", "columnIndex"]);
            ranges.push(range);
          }
        }
        await context.sync();
        for (let nrange of ranges) {
          nrange.format.fill.color == "#FF0000"
            ? showMessage({
                style: "error",
                message: `The field is not vaild!<br> Row: ${nrange.rowIndex + 1}; column: ${nrange.columnIndex};`,
              })
            : null;
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

export async function init_forrmLayout() {
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

        sheet = sheets.add(global.CurrentConfig.module);
        sheet.activate();
        return await context.sync().then(() => {
          return true;
        });
        // .then(() => {
        //   utils.set_layout(sheet, global.CurrentConfig.layout);
        // })
        // .then(() => {
        //   utils.set_table(sheet, global.CurrentConfig.table);
        // })
        // .then(() => {
        //   // utils.set_tableHeaders(sheet, global.CurrentConfig.table);
        // })
        // .then(() => {
        //   utils.set_tablebody(sheet, global.CurrentConfig.table);
        // });
      } catch (e) {
        console.log("get sheet");
        console.error(e.stack);
        sheet = context.workbook.worksheets.getItem(global.CurrentConfig.module);
        sheet.activate();
        await context.sync();
        // .then(() => {
        //   utils.set_layout(sheet, global.CurrentConfig.layout);
        // })
        // .then(() => {
        //   utils.set_table(sheet, global.CurrentConfig.table);
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

export async function init_set_layout(elements) {
  console.log("init_set_layout");
  try {
    //   await fetch("./mbhk_config.json").then((response) => console.log(response.json()));

    //   let rawData = fs.readFileSync("./mbhk_config.json", "utf8");
    //   mbhk_config = JSON.parse(rawData);
    return await Excel.run(async function (context) {
      let sheet = null;

      try {
        let sheet = context.workbook.worksheets.getItem(global.CurrentConfig.module);

        return utils.set_layout(sheet, elements);
        // .then(() => {
        // })
        // .then(() => {
        //   utils.set_table(sheet, global.CurrentConfig.table);
        // })
        // .then(() => {
        //   // utils.set_tableHeaders(sheet, global.CurrentConfig.table);
        // })
        // .then(() => {
        //   utils.set_tablebody(sheet, global.CurrentConfig.table);
        // });
      } catch (e) {
        console.log("Error get sheet");
        console.error(e.stack);
        sheet = context.workbook.worksheets.getItem(global.CurrentConfig.module);
        sheet.activate();
        await context.sync();
        // .then(() => {
        //   utils.set_layout(sheet, global.CurrentConfig.layout);
        // })
        // .then(() => {
        //   utils.set_table(sheet, global.CurrentConfig.table);
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

export async function init_set_table(config) {
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
        //   utils.set_layout(sheet, global.CurrentConfig.layout);
        // })
        // .then(() => {
        //   utils.set_table(sheet, global.CurrentConfig.table);
        // })
        // .then(() => {
        //   // utils.set_tableHeaders(sheet, global.CurrentConfig.table);
        // })
        // .then(() => {
        //   utils.set_tablebody(sheet, global.CurrentConfig.table);
        // });
      } catch (e) {
        console.log("Error get sheet");
        console.error(e.stack);
        sheet = context.workbook.worksheets.getItem(global.CurrentConfig.module);
        sheet.activate();
        await context.sync();
        //     .then(() => {
        //       utils.set_layout(sheet, global.CurrentConfig.layout);
        //     })
        //     .then(() => {
        //       utils.set_table(sheet, global.CurrentConfig.table);
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

export async function checkSheetTableHasData(config) {
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

export async function getSheetTableData(config) {
  console.log("getSheetTableData");
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

export async function Get_main_Data(id) {
  console.log(`on ${global.CurrentConfig.module} get main data...`);
  var expand = "";
  var modelIdName = "";
  if (!global.CurrentConfig.module.includes("CLAIM")) {
    expand = "&$expand=benz_NameofPreparer($select=internalemailaddress)";
    modelIdName = "Memo ID";
  } else {
    modelIdName = "Claim ID";
  }
  var statusChecking = " and (benz_formstatus eq 650230000 or benz_formstatus eq 650230001 or benz_formstatus eq 650230002 or benz_formstatus eq null)";
  try {
    global.Callapiaction = {
      name: "callapiaction",
      action: {
        entitySet: global.CurrentConfig["Table_Main_LogicalName"] + "s",
        queryString: `$filter=contains(${global.CurrentConfig["Table_Main_id_LogicalName"]},'${id}')${statusChecking}${expand}`,
        queryOptions: "",
      },
    };
    //jq(`#${Module + "_MomeID_input"}`)
    const object = await dataHelp.RetrieveAndReturnMultipleData(null, {
      entitySet: global.CurrentConfig["Table_Main_LogicalName"] + "s",
      queryString: `$filter=contains(${global.CurrentConfig["Table_Main_id_LogicalName"]},'${id}')${statusChecking}${expand}`,
      queryOptions: "",
    });
    if (object.value.length < 1) {
      throw new Error(`Can't not find "${modelIdName}" "${id}"`);
    }

    global.updatingRecord = id;

    await utils.lockmainObjectValueToCell(false);

    await BENZ.buildForm(global.CurrentConfig.module);

    await set__mainData(object);
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

    showMessage({ style: "error", message: "Error: " + e.message });
  }
}

export async function Get_items_Data(config, mainid) {
  console.log(`on ${config.module} get items data...`);
  var statusChecking = "";
  if (global.CurrentConfig.module.includes("CLAIM")) {
    statusChecking = " and (benz_claimapprovalstatus eq 650230008 or benz_claimapprovalstatus eq null)";
  }
  try {
    global.Callapiaction = {
      name: "callapiaction",
      action: {
        entitySet: config["Table_Item_LogicalName"] + "s",
        queryString: `$filter=_${config["Table_Main_LogicalName"]}_value eq '${mainid}' ${statusChecking} and statecode eq 0 &$orderby=benz_id asc`,
        queryOptions: "",
      },
    };
    console.log(global.Callapiaction);

    let obj = await dataHelp.RetrieveAndReturnMultipleData(null, {
      entitySet: config["Table_Item_LogicalName"] + "s",
      queryString: `$filter=_${config["Table_Main_LogicalName"]}_value eq '${mainid}' ${statusChecking} and statecode eq 0 &$orderby=benz_id asc`,
      queryOptions: "",
    });
    console.log("Line Data");
    console.log(obj);
    await set__tableData(obj);

    console.log("object");
    console.log(obj);
  } catch (e) {
    console.log(e.message);
    console.error(e.stack);
  }
}

export async function set__mainData(response) {
  console.log("set__mainData");
  console.log(response);
  const conf = global.CurrentConfig.layout.filter(
    (ele) => ele.type == "main-form-field" || ele.type == "main-form-field-gen" // && ele.action.includes("update")
  );
  const obj = response.value[0];
  if (obj) {
    await utils.mainObjectValueToCell(conf, obj);
    await utils.remove_tablerows(global.CurrentConfig);
    await Get_items_Data(global.CurrentConfig, obj[global.CurrentConfig["Table_Main_LogicalName"] + "id"]);
  }
}

export async function set__tableData(response) {
  console.log("set__tableData");
  console.log(response);
  const conf = global.CurrentConfig.table.columns.filter((item) => item.enable);
  const arr = utils.ObjectToTableArray(conf, response.value);
  global.lastnumber = Math.max.apply(
    Math,
    response.value.map((o) => o[global.CurrentConfig.Table_linenumber_LogicalName])
  );

  console.log(arr);
  await utils.setDataToTable(global.CurrentConfig, arr);
  await utils.addTableDataChange(global.CurrentConfig);
}
