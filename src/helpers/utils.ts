/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, require, console, document, Excel */

// import { showMessage, clearMessage } from "../helpers/message-helper";
import { getUserData } from "../helpers/sso-helper";
import {
  get_Data__benz_supporttypes,
  get_Data__benz_typeClasses,
  get_Data__benz_carModel,
} from "../helpers/dataverse-data-helper";
// import { Entity } from "../helpers/dataverse-webapi/lib/models";
// import { retrieve, retrieveMultiple, WebApiConfig } from "../helpers/dataverse-webapi/lib/node";
// import { JsonToTable } from "../helpers/JsonToTable";
// import { executionAsyncId } from "async_hooks";
// import DataTable from 'datatables.net-bs5';
// import 'datatables.net-searchpanes'
import DataTable from "datatables.net-searchpanes-bs5";
import "datatables.net-select-bs5";
var jq = require("jquery");

var typeclasses = [];
var supporttypes = [];
var currentRangeText = "";
var canApplyModel = false;
var applyModelButtonId = "applyModel";
var Table_id__model = "Table_id__model";
var Table__model = null;

export let applyValidation = function () {
  return true;
};

export var setApplyModel = function (condition) {
  canApplyModel = condition;
  // if (!condition) {
  //   console.log(jq("#applyModel"));
  //   jq(`#${"applyModelButtonId"}`).removeAttr("disabled");
  // } else {
  //   jq(`#${"applyModelButtonId"}`).attr("disabled", "disabled");
  // }
};

var afterSupporttypes = function () {
  get_Data__benz_typeClasses(writeToOfficeDocument__benz_typeClasses);
};
var afterTypeClasses = function () {
  get_Data__benz_carModel(writeToOfficeDocument__benz_carmodels);
};
// var afterSupporttypes = ;

export async function getProfileButton() {
  getUserData(writeDataToOfficeDocument);
}

export async function _get_Data__benz_supporttypes() {
  // await get_Data__benz_supporttypes(writeToOfficeDocument__benz_supporttypes);
  get_Data__benz_supporttypes(writeToOfficeDocument__benz_supporttypes);
  // await get_Data__benz_typeClasses(writeToOfficeDocument__benz_typeClasses);
  // await get_Data__benz_carModel(writeToOfficeDocument__benz_carmodels);
}

function findIndex(stringArr, keyString) {
  // Initialising result array to -1
  // in case keyString is not found
  let result = [-1, -1];

  // Iteration over all the elements
  // of the 2-D array

  // Rows
  for (let i = 0; i < stringArr.length; i++) {
    // Columns
    for (let j = 0; j < stringArr[i].length; j++) {
      // If keyString is found
      if (stringArr[i][j] == keyString) {
        result[0] = i;
        result[1] = j;
        return { id: stringArr[i][j], value: stringArr[i][j + 1] };
      }
    }
  }

  // If keyString is not found
  // then -1 is returned
  return result;
}
// Function to customize the selected cells
function customizeSelectedCells() {
  Excel.run(function (context) {
    // Get the selected range
    var range = context.workbook.getSelectedRange();

    // Load the values and format of the range
    range.load(["values", "format"]);

    return context
      .sync()
      .then(function () {
        // Iterate through each cell in the range
        for (var row = 0; row < range.values.length; row++) {
          for (var col = 0; col < range.values[row].length; col++) {
            var cellValue = range.values[row][col];

            // Customize the cell based on the cell value
            if (cellValue > 0) {
              range.getCell(row, col).format.fill.color = "green";
            } else if (cellValue < 0) {
              range.getCell(row, col).format.fill.color = "red";
            } else {
              range.getCell(row, col).format.fill.color = "yellow";
            }
          }
        }
      })
      .then(context.sync);
  }).catch(function (error) {
    console.log(error);
  });
}
export async function addActionOnSelect(): Promise<any> {
  return await Excel.run(async function (context) {
    // document.getElementById("message-area").innerHTML = "<br>writeDataverseDataToOfficeDocument: " ;
    console.log(`addActionOnSelect`);

    try {
      var sheet = context.workbook.worksheets.getActiveWorksheet();
      var range = sheet.getRange();

      // Bind the range to monitor for changes
      sheet.onSelectionChanged.add(async function (eventArgs) {
        global.currentSelectRangeAddress = eventArgs.address;
        console.log(global.currentSelectRangeAddress);
        // Load the values and format of the selected range
        // var range = sheet.getRange(global.currentSelectRangeAddress);
        await showHideDetaleButton(context, sheet, global.currentSelectRangeAddress);
        await onSelectionChanged__model(context, sheet, global.currentSelectRangeAddress);

        return await context.sync();
      });

      // return context.sync();

      // console.log('addActionOnSelect')
      // let range = context.workbook.getSelectedRange();
      // range.load("address");

      // await context.sync();

      // console.log(`The address of the selected range is "${range.address}"`);
    } catch (exception) {
      console.log("addActionOnSelect EXCEPTION: " + exception.message);
    }
    return context.sync();
  });
}

export async function showHideDetaleButton(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  selectedRange: string
) {
  try {
    console.log(`showHideDetaleButton...`);
    var sheetname = "";
    var values = [];
    var columnIndex = -1;
    var rowIndex = -1;
    sheet.load(["name"]);
    var range = sheet.getRange(selectedRange);
    range.load(["text", "columnIndex", "rowIndex"]);
    await context.sync();

    sheetname = sheet.name;
    console.log(`sheet.name:"${sheetname}"`);

    values = range.text;
    columnIndex = range.columnIndex;
    rowIndex = range.rowIndex;
    console.log(
      `sheet.name:"${sheetname}",rowIndex:"${JSON.stringify(rowIndex)}",columnIndex:"${JSON.stringify(columnIndex)}" `
    );
    if (columnIndex == 0 && rowIndex > 19) {
      jq("#clearRow").show();
      jq("#deleterow").show();
    } else {
      jq("#clearRow").hide();
      jq("#deleterow").hide();
    }
  } catch (e) {
    console.log(`Error onSelectionChanged__model;${e.message}`);
  }
}
export async function onSelectionChanged__model(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  selectedRange: string
) {
  //SM ZF (Hong Kong)

  try {
    console.log(`onSelectionChanged__model...`);
    var sheetname = "";
    var values = [];
    var columnIndex = -1;
    var rowIndex = -1;
    sheet.load(["name"]);
    await context.sync();

    sheetname = sheet.name;
    console.log(`sheet.name:"${sheetname}"`);

    var range = sheet.getRange(selectedRange);
    range.load(["text", "columnIndex", "rowIndex"]);
    await context.sync();

    values = range.text;
    columnIndex = range.columnIndex;
    rowIndex = range.rowIndex;
    console.log(
      `sheet.name:"${sheetname}",rowIndex:"${JSON.stringify(rowIndex)}",columnIndex:"${JSON.stringify(columnIndex)}" `
    );
    // sheet.name === "SM ZF (Hong Kong)" && columnIndex === 2 && rowIndex > 12
    // console.log(`is "SM ZF (Hong Kong)" selecting "model"`);
    setApplyModel(true);
  } catch (e) {
    console.log(`Error onSelectionChanged__model;${e.message}`);
  }
  return context.sync();
}

export function writeDataToOfficeDocument(result: Object): Promise<any> {
  return Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    let data = [];
    let userProfileInfo: string[] = [];
    userProfileInfo.push(result["displayName"]);
    userProfileInfo.push(result["jobTitle"]);
    userProfileInfo.push(result["mail"]);
    userProfileInfo.push(result["mobilePhone"]);
    userProfileInfo.push(result["officeLocation"]);

    for (let i = 0; i < userProfileInfo.length; i++) {
      if (userProfileInfo[i] !== null) {
        let innerArray = [];
        innerArray.push(userProfileInfo[i]);
        data.push(innerArray);
      }
    }
    const rangeAddress = `B5:B${5 + (data.length - 1)}`;
    const range = sheet.getRange(rangeAddress);
    range.values = data;
    range.format.autofitColumns();

    return context.sync();
  });
}

export async function writeToOfficeDocument__benz_supporttypes(result): Promise<any> {
  // =INDIRECT("benz_supporttypes")
  return await Excel.run(async function (context) {
    // document.getElementById("message-area").innerHTML = "<br>writeDataverseDataToOfficeDocument: " ;

    try {
      console.log("writeToOfficeDocument__benz_supporttypes");
      // const sheet = context.workbook.worksheets.getActiveWorksheet();
      let data = [];
      // let userProfileInfo: string[] = [];
      if (result) {
        // result.value;
        if (result.value) {
          data = result.value.map((element) => [element.benz_name]);
        }
      }
      if (!data) {
        throw new Error("data error");
      }
      let sheet = context.workbook.worksheets.getItem("Dropdown list");
      if (!sheet) {
        throw new Error('can\'t find "Dropdown list" sheet');
      }
      let expensesTable = sheet.tables.getItem("benz_supporttypes");
      if (!expensesTable) {
        throw new Error('can\'t find "Dropdown list" benz_supporttypes');
      }
      // console.log( "data: " + JSON.stringify(data));
      expensesTable.rows.load("items");
      await context.sync();
      expensesTable.rows.deleteRows(expensesTable.rows.items);
      console.log("benz_supporttypes table is deleted");
      expensesTable.rows.add(
        null, // index, Adds rows to the end of the table.
        data,
        true // alwaysInsert, Specifies that the new rows be inserted into the table.
      );

      sheet.getUsedRange().format.autofitColumns();
      sheet.getUsedRange().format.autofitRows();
      await context.sync();

      // after run function
      console.log("after run function...");
      await afterSupporttypes();
    } catch (exception) {
      console.log("writeToOfficeDocument__benz_supporttypes EXCEPTION: " + exception.message);
    }
    return context.sync();
  });
}

export async function writeToOfficeDocument__benz_typeClasses(result): Promise<any> {
  // =INDIRECT("benz_prototypemodeldesignationtypeclasses")
  return await Excel.run(async function (context) {
    // document.getElementById("message-area").innerHTML = "<br>writeDataverseDataToOfficeDocument: " ;

    try {
      console.log("writeToOfficeDocument__benz_typeClasses");
      // const sheet = context.workbook.worksheets.getActiveWorksheet();
      let data = [];
      // let userProfileInfo: string[] = [];
      if (result) {
        result.value;
        if (result.value) {
          data = result.value.map((element) => [
            element.benz_prototypemodeldesignationtypeclassid,
            element.benz_typeclass,
          ]);
        }
      }
      typeclasses = data;

      // after run function
      await afterTypeClasses();
      // await get_Data__benz_carModel(writeToOfficeDocument__benz_carmodels);
    } catch (exception) {
      console.log("writeToOfficeDocument__benz_typeClasses EXCEPTION: " + exception.message);
    }
    return context.sync();
  });
}

export async function writeToOfficeDocument__benz_carmodels(result): Promise<any> {
  // =INDIRECT("benz_prototypemodeldesignation")

  try {
    console.log("writeToOfficeDocument__benz_car_models");
    // const sheet = context.workbook.worksheets.getActiveWorksheet();
    let data = [];
    // let userProfileInfo: string[] = [];
    if (result) {
      result.value;
      if (result.value) {
        data = result.value.map(
          (element) => [
            element.benz_name,
            element._benz_prototypemodeldesignationtypeclass_value,
            element.benz_icebevhybrid,
            element.benz_amgmaybachna,
          ]
          // console.log(element.benz_name)
        );
      }
    }
    if (!data) {
      throw new Error("data error");
    }
    // // let sheet = context.workbook.worksheets.getItem("Dropdown list");
    // if(!sheet){
    //   throw new Error('can\'t find "Dropdown list" sheet');
    // }
    // // let expensesTable = sheet.tables.getItem("benz_prototypemodeldesignation");
    // if(!expensesTable){
    //   throw new Error('can\'t find "Dropdown list" benz_prototypemodeldesignation');
    // }
    // console.log( "data: " + (data));
    global.Table__model = new DataTable(`#Table_id__model`, {
      searchPanes: true,
      select: {
        style: "multi",
      },

      deferRender: true,
      data: data,
      columns: [
        { title: "Model Name" },
        { title: "Type Class" },
        { title: "ICE / BEV" },
        { title: "AMG / Non-AMG" },
      ],
      //     columnDefs: [
      //     {
      //       orderable: false,
      //       className: 'select-checkbox',
      //       targets:   0,
      //       width: '30'
      //     }
      //  ],
      columnDefs: [
        {
          targets: [1],
          searchPanes: {
            show: true,
            orthogonal: "searchpanes",
          },
          render: function (data, type) {
            return findIndex(typeclasses, data)["value"] ?? "NULL";
          },
        },
        {
          targets: [2],
          searchPanes: {
            show: true,
            orthogonal: "searchpanes",
          },
        },
        {
          targets: [3],
          searchPanes: {
            show: true,
            orthogonal: "searchpanes",
          },
        },
      ],

      dom: "Pfrtip",
    });
    global.Table__model.searchPanes.container().prependTo(global.Table__model.table().container());
    global.Table__model.searchPanes.resizePanes();
    console.log("writeToOfficeDocument__benz_car_models Done...");
    document.getElementById(`${applyModelButtonId}`).onclick = applyNewModel;
  } catch (exception) {
    console.log("writeToOfficeDocument__benz_car_models EXCEPTION: " + exception.message);
  }
}

// export async function applyNewModel() {
//   return await Excel.run(async function (context) {
//     console.log("applyNewModel...");
//     try {
//       let modelNames = [];
//       if (Table__model) {
//         let data = Table__model.rows({ selected: true }).data();
//         let ndata = [];
//         var newarray = [];
//         for (var i = 0; i < data.length; i++) {
//           ndata.push(data[i][0]);
//           console.log("Name: " + data[i][0] + " Address: " + data[i][1] + " Office: " + data[i][2]);
//         }

//         var sData = ndata.join(",");
//         console.log(`rows_selected: ${sData}`);

//         var sheetname = "";
//         var sheet = context.workbook.worksheets.getActiveWorksheet();

//         sheet.load(["name"]);
//         await context.sync();
//         sheetname = sheet.name;
//         console.log(`sheet.name:"${sheetname}"`);

//         var range = sheet.getRange(global.currentSelectRangeAddress);
//         range.load(["text", "columnIndex", "rowIndex"]);
//         await context.sync();

//         var columnIndex = range.columnIndex;
//         var rowIndex = range.rowIndex;
//         console.log(
//           `sheet.name:"${sheetname}",rowIndex:"${JSON.stringify(rowIndex)}",columnIndex:"${JSON.stringify(
//             columnIndex
//           )}" `
//         );
//         // sheet.name === "SM ZF (Hong Kong)" && columnIndex === 2 && rowIndex > 12

//         range.values = [[sData]];
//         // range.format.autofitColumns();
//         await context.sync();
//         //   if(applyValidation()){
//         // }else{
//         //   console.log(`Error "SM ZF (Hong Kong)" no selecting "model"`);

//         // }
//       } else {
//         console.log("Error applyNewModel: table null");
//       }
//     } catch (e) {
//       console.log("Error applyNewModel: " + e.message);
//     }
//     return context.sync();
//   });
// }
