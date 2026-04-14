/* global global, console */

import { BENZ, FormAction } from "../factory";
import * as BenzType from "../type";
import * as utils from "../utils";
import * as datahelper from "../../dataverse-data-helper";
import { get_choicesSets } from "../components/choices";
import * as Benz_type from "../type";

export class SM extends BENZ {}

export class Action extends FormAction {
  constructor(config: BenzType.FormConfig) {
    super(config); // must call super()
    // let memoIdLength = this.config.memoid_set.length;
    console.log("action checkj");
    console.log(global.CurrentConfig["action"]);
    if (global.CurrentConfig["action"] !== "update") {
      this.mainObjOtherActions.push(this.getMemoNo);
    }
  }
  // async create() {
  //   console.log(`on ${this.config.module} post main data...`);
  //   try {
  //     // get main data
  //     let mainObj = await utils.get_main_data_by_object(
  //       this.config,
  //       this.config.layout.filter((ele) => ele.type == "main-form-field"),
  //       [this.getMemoNo]
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

  //     // setup Measure ID
  //     itemObjs = itemObjs.map((obj) => {
  //       console.log("itemObjs");
  //       console.log(obj);
  //       obj["benz_name"] = `${posted_mainData["benz_name"]}-${obj["benz_id"]}`;
  //       return obj;
  //     });

  //     const posted_itemData: any[] = [];
  //     for (let itemObj of itemObjs) {
  //       posted_itemData.push(await utils.postDataReturnData(this.config["Table_Item_LogicalName"] + "s", itemObj));
  //     }
  //     // = await Promise.all(
  //     //   itemObjs.map(
  //     //     async (itemObj): Promise<any> => await utils.postDataReturnData(config["Table_Item_LogicalName"] + "s", itemObj)
  //     //   )
  //     // );
  //     // const posted_itemData = [];
  //     // for (const itemObj of itemObjs) {
  //     //   const nitemObj = ;
  //     //   posted_itemData.push(nitemObj);

  //     // }
  //     console.log("posted_itemData");
  //     console.log(posted_itemData);
  //     global.form.items = posted_itemData;

  //     console.log(global.form);
  //     const RelatedDataConf = await this.readyRelatedData(this.config, posted_mainData, posted_itemData);
  //     const res = this.post_related_Data(RelatedDataConf);
  //     console.log("map related_Data res");
  //     console.log(res);

  //     // set id

  //     await utils.lockmainObjectValueToCell(false);
  //     await this.setidOnSheet(this.config, "benz_memonoserialno", posted_mainData["benz_memonoserialno"]);
  //     await this.setidOnSheet(this.config, "benz_name", posted_mainData["benz_name"]);
  //     await utils.lockmainObjectValueToCell(true);
  //   } catch (e) {
  //     console.error(e.message);
  //     console.error(e.stack);
  //      showMessage({ style: "error", message: "Error: " + e.message });
  //   }
  // }

  // async update() {
  //   throw new Error("Method not implemented.");
  // }

  async beforePostData(): Promise<void> {
    // setup Measure ID
    this.itemObjs = this.itemObjs.map((obj) => {
      console.log("itemObjs");
      console.log(obj);
      obj["benz_name"] = `${this.posted_mainData["benz_name"]}-${obj["benz_id"]}`;
      return obj;
    });
  }

  async afterPostData(): Promise<void> {
    await utils.lockmainObjectValueToCell(false);
    await this.setidOnSheet(this.config, "benz_memonoserialno", this.posted_mainData["benz_memonoserialno"]);
    await this.setidOnSheet(this.config, "benz_name", this.posted_mainData["benz_name"]);
    const dateTime = new Date(this.posted_mainData["modifiedon"]);
    const date = dateTime.toISOString().split('T')[0];
    const time = dateTime.toTimeString().split(' ')[0];
    const modifiedon = date.toString() + " " + time.toString();
    await this.setidOnSheet(this.config, "modifiedon", modifiedon);
    // await this.setidOnSheet(
    //   this.config,
    //   "benz_prototypesalesmeasuremainformid",
    //   this.posted_mainData["benz_prototypesalesmeasuremainformid"]
    // );
    // await this.setitemidOnSheet(
    //   this.config,
    //   "benz_prototypesalesmeasureid",
    //   this.posted_itemData["benz_prototypesalesmeasureid"]
    // );
    await utils.lockmainObjectValueToCell(true);
  }

  async newCommissionPostData(itemObj, currentRow) {}
  
  async _beforePostData() {
    // setup new row Measure ID for update
    this.itemObjs = this.itemObjs.map((obj) => {
      console.log("itemObjs");
      console.log(obj);
      obj["benz_name"] = `${this.mainObj["benz_name"]}-${obj["benz_id"]}`;
      return obj;
    });
  }

  async getMemoNo(q: BenzType.queryUpdateObject) {
    console.log("getMemoNo");
    console.log(q);
    // get ready db CountByModth
    let memotype_Value = q.object[`${q.config["memotype_LogicalName"]}@odata.bind`];
    if (!memotype_Value) {
      throw new Error("Memotype can not empty");
    }
    memotype_Value = memotype_Value.substring(memotype_Value.indexOf("(") + 1, memotype_Value.lastIndexOf(")"));

    global.Callapiaction = {
      name: "callapiaction",
      action: {
        entitySet: "benz_prototypesalesmeasuremainforms",
        queryString: `?$select=benz_prototypesalesmeasuremainformid,benz_memonoserialno&$filter=contains(benz_memonoyear,'${q.object["benz_memonoyear"]}') and contains(benz_memonomonth,'${q.object["benz_memonomonth"]}') and _benz_memotype_value eq ${memotype_Value}&$count=true`,
        queryOptions: "",
      },
    };

    let res = await datahelper.getCountByModth();
    console.log(res);
    if (res) {
      var new_number =
        res["@odata.count"] > 0
          ? res.value?.length == 0
            ? 1
            : Math.max(...res.value.map((e) => e["benz_memonoserialno"])) + 1
          : 1;
      // const memo_no = res["@odata.count"] + 1;

      const memo_no = new_number >= 10 ? new_number : "0" + new_number;

      console.log("mode check");
      console.log(global.mode);
      if (global.mode == Benz_type.AppMode.MIGRATION) {
        q.config["memoid_set"] = ["benz_name"];
      }

      console.log("after get_memo_no");
      q.object["benz_memonoserialno"] = memo_no;
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
            return key;
          }
          return q.object[key];
        })
        .join("");
      console.log("memoid_set");
      console.log(memoid_set);
      q.object["benz_memonoserialno"] = memo_no;
      q.object["benz_name"] = memoid_set;
      console.log("getMemoNo done");
      console.log(q.object);
      // return await afterAction(entity, mainFormObject, config, callback);
    } else {
      new Error("Memo no Can't no generate !");
    }
    return q.object;
  }
}
