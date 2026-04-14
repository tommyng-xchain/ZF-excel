/* global global, console */

import { BENZ, FormAction } from "../factory";
import * as BenzType from "../type";
import * as utils from "../utils";
import * as datahelper from "../../dataverse-data-helper";
import { get_choicesSets } from "../components/choices";
import { showMessage } from "../../message-helper";
import { callRetrieveMultipleData } from "../../middle-tier-calls";

export class CLAIM extends BENZ {}

export class Action extends FormAction {
  constructor(config: BenzType.FormConfig) {
    super(config); // must call super()
  }
  // async create() {
  //   console.log(`on ${this.config.module} post main data...`);
  //   try {
  //     // get main data
  //     let mainObj = await utils.get_main_data_by_object(
  //       this.config,
  //       this.config.layout.filter((ele) => ele.type == "main-form-field"),
  //       []
  //     );
  //     if (!mainObj) {
  //       throw new Error("Can't get form data !");
  //     }
  //     // get table data
  //     let itemObjs = await utils.getready_table_data_by_object(this.config);

  //     if (!itemObjs) {
  //       throw new Error("Can't get form data !");
  //     }
  //     // mainObj = await getMemoNo(config, mainObj);

  //     console.log("mainObj 2");
  //     console.log(mainObj);
  //     console.log("itemObjs");
  //     console.log(itemObjs);

  //     let isValid = await this.dataValidation(
  //       mainObj,
  //       this.config,
  //       this.config.layout.filter((ele) => ele.type == "main-form-field")
  //     );
  //     for (let itemObj of itemObjs) {
  //       isValid = isValid && (await this.dataValidation(itemObj, this.config, this.config.table.columns));
  //     }

  //     console.log("isValid");
  //     console.log(isValid);
  //     if (!isValid) {
  //       throw new Error("data not Validation ");
  //     }

  //     const posted_mainData = await utils.postDataReturnData(this.config["Table_Main_LogicalName"] + "s", mainObj);
  //     global.form.relationship_items["main"].push(posted_mainData);
  //     if (!posted_mainData) {
  //       throw new Error("Can't Main post Data");
  //     }
  //     if (!posted_mainData["benz_name"]) {
  //       throw new Error("Can't Main benz_name");
  //     }

  //     console.log("posted_mainData");
  //     console.log(posted_mainData);

  //     const posted_itemData: any[] = [];
  //     for (let itemObj of itemObjs) {
  //       posted_itemData.push(await utils.postDataReturnData(this.config["Table_Item_LogicalName"] + "s", itemObj));
  //     }
  //     console.log("posted_itemData");
  //     console.log(posted_itemData);
  //     global.form.items = posted_itemData;

  //     // set id
  //   } catch (e) {
  //     console.error(e.message);
  //     console.error(e.stack);
  //         showMessage({ style: "error", message: "Error: " + e.message });
  //   }
  // }

  // async update() {
  //   try {
  //     throw new Error("update not working.");
  //   } catch (e) {
  //     showMessage({ style: "error", message: "Error: " + e.message });
  //   }
  // }
  async beforePostData(): Promise<void> {
    // setup Measure ID
    let o = [];
    var countRow = 0;
    for (let obj of this.itemObjs) {
      console.log("itemObjs");
      console.log(obj);
      obj["benz_name"] = `${this.posted_mainData["benz_name"]}-${obj["benz_id"]}`;
      let cof: BenzType.queryUpdateObject = { config: this.config, object: obj };
      if (!global.updatingRecord) {
        obj = await this.newcommissionno(cof, countRow);
      } else {
        delete obj["benz_commissionnumber"];
      }
      o.push(obj);
      countRow++;
    }

    this.itemObjs = o;
  }
  async afterPostData(): Promise<void> {
    await utils.lockmainObjectValueToCell(false);
    await this.setidOnSheet(this.config, "benz_name", this.posted_mainData["benz_name"]);
    await utils.lockmainObjectValueToCell(true);
  }

  async newCommissionPostData(itemObj, row): Promise<void> {
    let cof: BenzType.queryUpdateObject = { config: this.config, object: itemObj };
    itemObj = await this.newcommissionno(cof, row);

    return itemObj;
  }

  async _beforePostData(): Promise<void> {
    let o = [];
    var countRow = 0;
    for (let obj of this.itemObjs) {
      console.log("itemObjs");
      console.log(obj);
      obj["benz_name"] = `${this.mainObj["benz_name"]}-${obj["benz_id"]}`;

      let cof: BenzType.queryUpdateObject = { config: this.config, object: obj };
      if (obj["benz_prototypeclaimitemid"] != null) {
        obj = await this.updatecommissionno(cof, countRow);
      }

      o.push(obj);
      countRow++;
    }
    this.itemObjs = o;
  }

  async getMemoNo(q: BenzType.queryUpdateObject) {
    console.log("getMemoNo");
    console.log(q);
    // get ready db CountByModth
    let memotype_Value = q.object[`${q.config["memotype_LogicalName"]}@odata.bind`];
    memotype_Value = memotype_Value.substring(memotype_Value.indexOf("(") + 1, memotype_Value.lastIndexOf(")"));

    global.Callapiaction = {
      name: "callapiaction",
      action: {
        entitySet: "benz_prototypesalesmeasuremainforms",
        queryString: `?$select=benz_prototypesalesmeasuremainformid&$filter=contains(benz_memonoyear,'${q.object["benz_memonoyear"]}') and contains(benz_memonomonth,'${q.object["benz_memonomonth"]}') and _benz_memotype_value eq ${memotype_Value}&$count=true`,
        queryOptions: "",
      },
    };

    let res = await datahelper.getCountByModth();
    if (res) {
      console.log("after get_memo_no res");
      console.log(res);
      console.log(res["@odata.count"]);
      console.log(typeof res["@odata.count"]);
      const memo_no = res["@odata.count"] + 1;
      console.log("after get_memo_no");
      q.object["benz_memonoserialno"] = memo_no.toString();
      const memoid_set = q.config["memoid_set"]
        .map((key) => {
          console.log("get_choicesSets " + key);
          console.log(q.object[key]);
          if (
            key.toString().trim().toLocaleLowerCase() ==
            q.config["memotype_LogicalName"].toString().trim().toLocaleLowerCase()
          ) {
            console.log(
              get_choicesSets(key).find(
                (e) =>
                  e.id.toString().trim().toLocaleLowerCase() == memotype_Value.toString().trim().toLocaleLowerCase()
              ).name
            );
            return (
              get_choicesSets(key).find(
                (e) =>
                  e.id.toString().trim().toLocaleLowerCase() == memotype_Value.toString().trim().toLocaleLowerCase()
              ).name ?? q.object[key].toString()
            );
            //  get_choicesSets(key).find((e) => e.value === mainFormObject[key].toString())?.name;
          }
          if (q.object[key] == undefined) {
            return key.toString();
          }
          return q.object[key].toString();
        })
        .join("");
      console.log("memoid_set");
      console.log(memoid_set);
      q.object["benz_memonoserialno"] = memo_no.toString();
      q.object["benz_name"] = memoid_set;
      console.log("getMemoNo done");
      console.log(q.object);
      // return await afterAction(entity, mainFormObject, config, callback);
    } else {
      new Error("Memo no Can't no generate !");
    }
    return q.object;
  }

  async newcommissionno(q: BenzType.queryUpdateObject, row) {
    console.log("newcommissionno");
    console.log(q);
    let LogicalName = "benz_commissionnumber";
    let SchemaName = "benz_Commissionnumber";

    let measureLogicalName = "benz_PrototypeSalesMeasure@odata.bind";
    let ItemNumberLogicalName = "benz_id";
    // get ready db CountByModth
    let v = q.object[LogicalName];
    let measureIDLookUp = q.object[measureLogicalName];
    let currentRow = row;
    let measureID = global.RowMeasureName[currentRow];
    console.log("New Commission Number");
    console.log(currentRow);
    console.log(global.RowMeasureName);
    console.log(measureID);
    let claimName = q.object["benz_name"].split("-")[0];

    global.Callapiaction = {
      name: "callapiaction",
      action: {
        entitySet: `${LogicalName}s`,
        queryString: { benz_name: v.toString(), "benz_MeasureID@odata.bind": measureIDLookUp, benz_savedmeasureidname: measureID, benz_claimname: claimName},
        queryOptions: "",
      },
    };

    let res = await datahelper.post_Data_ReturnData();
    console.log(res);
    if (res) {
      q.object[`${SchemaName}@odata.bind`] = `${LogicalName}s(${res[`${LogicalName}id`]})`;
      delete q.object[LogicalName];
    } else {
      throw new Error("Memo no Can't no generate !");
    }
    return q.object;
  }

  async updatecommissionno(q: BenzType.queryUpdateObject, row) {
    console.log("updatecommissionno");
    console.log(q);
    let LogicalName = "benz_commissionnumber";
    let SchemaName = "benz_Commissionnumber";
  
    let measureLogicalName = "benz_PrototypeSalesMeasure@odata.bind";
    let ItemNumberLogicalName = "benz_id";
    // get ready db CountByModth
    let v = q.object[LogicalName];
    let measureIDLookUp = q.object[measureLogicalName];
    let currentRow = row;
    let measureID = global.RowMeasureName[currentRow];
    console.log("Update Commission Number");
    console.log(currentRow);
    console.log(global.RowMeasureName);
    console.log(measureID);
    let claimName = q.object["benz_name"].split("-")[0];
    const currentClaimItemGUID = q.object["benz_prototypeclaimitemid"];

    global.Callapiaction = {
      name: "callapiaction",
      action: {
        entitySet: `${LogicalName}s`,
        queryString: `$select=benz_name,benz_savedmeasureidname,_benz_measureid_value&$filter=benz_savedmeasureidname eq '${measureID}' and benz_claimname eq '${claimName}'`,
        queryOptions: "",
      },
    };

    try {
      const existRecords = await callRetrieveMultipleData(global.ApiAccessToken);
      console.log(existRecords);
      console.log(existRecords["value"].length);
      if (existRecords["value"].length > 0) {
        global.Callapiaction = {
          name: "callapiaction",
          action: {
            entitySet: `benz_prototypeclaimitems`,
            queryString: `$select=_benz_commissionnumber_value&$filter=benz_prototypeclaimitemid eq '${currentClaimItemGUID}'`,
            queryOptions: "",
          },
        };
        const claimLineItemReocrd = await callRetrieveMultipleData(global.ApiAccessToken);
        const currentCommissionID = claimLineItemReocrd["value"][0]["_benz_commissionnumber_value"];
        for (var i = 0; i < existRecords["value"].length; i++) {
          if (
            existRecords["value"][i]["benz_name"] != v &&
            existRecords["value"][i]["benz_commissionnumberid"] == currentCommissionID
          ) {
            console.log("Update record:", global.RowMeasureID[currentRow]);
            var obj = { benz_name: v };
            await utils.updateDataReturnData(LogicalName + "s", existRecords["value"][i][LogicalName + "id"], obj);
            break;
          }
        }
      }
      delete q.object[LogicalName];

      return q.object;
    } catch (error) {
      throw new Error("Commission No cannot update");
    }

    return q.object;
  }
}
