//------ benz/utils.js ------

/* global global, console, require, Excel, document */

import {
  Cell,
  FormConfig,
  FormLayout,
  initConfig,
  mainformobjectKey,
  queryUpdateObject,
  Table,
  TableColumns,
} from "./type";

// eslint-disable-next-line @typescript-eslint/no-unused-vars
import { get_choices, get_choices_list, get_choicesSets } from "./components/choices";
import * as datahelper from "../dataverse-data-helper";
import * as onchange_dataValidation from "./table/onChangeDataValidation/init";
import * as picker from "./table/picker/init";

import { BENZ } from "../benz/factory";
import * as Benz_type from "../benz/type";

var jq = require("jquery");

export async function set_layout(sheet: Excel.Worksheet, config: FormLayout[] | TableColumns[], c_range?: Excel.Range) {
  try {
    return await Excel.run(async function (context) {
      try {
        console.log("set_layout");
        sheet = context.workbook.worksheets.getActiveWorksheet();
        var useSupportUnitFormula = false;
        for (let element of config) {
          console.log(element);

          try {
            // const isFormLayout = (x: any): x is FormLayout => element.includes(x);
            // const isTableColumns = (x: any): x is TableColumns => element.includes(x);

            // if (!element.cell.address && !element.index) {
            // }
            let range: Excel.Range = null;
            if (!element.cell) {
              throw new Error("Error: can't find cell json config");
            }
            if ("address" in element.cell) {
              if (element.cell.address) {
                range = sheet.getRange(element.cell.address);
              }
            } else if ("index" in element) {
              range = sheet.getCell(element.rowIndex, element.index);
            } else if (c_range) {
              range = c_range;
            } else {
              throw new Error("Error: can't find range");
            }
            // set_layout_range(element.cell.address, element);
            // console.log(global.CurrentConfig["action"]);
            // if (!element.action.includes(global.CurrentConfig["action"])) {
            //   throw new Error("This element config no eq action");
            // }
            // if (element.cell.columnHidden) {
            //   range.columnHidden = element.cell.columnHidden;
            // }

            if (element.cell.values) {
              console.log("element.Label");
              console.log(element.Label);
              console.log(element.cell.values);
              if (element.LogicalName == "benz_formbhkmbfsonly") {
                range.load(["values"]);
                await context.sync();
                const currentValue = range.values[0];
                if (currentValue[0] == "") {
                  range.values = [[element.cell.values ?? ""]];
                }
              } else {
                range.values = [[element.cell.values ?? ""]];
              }

              if (element.setToString) {
                range.values = [[element.cell.values.toString() ?? ""]];
              }

              if (element.LogicalName == "benz_prototypesupporttype") {
                if (element.cell.values.startsWith("1) ")) {
                  useSupportUnitFormula = true;
                }
              }
            }
            if (element.isFormula) {
              console.log("environment_name check");
              // console.log(element.useEnvironmentVaule);
              if (element.useEnvironmentVaule) {
                var environmentValue =
                  global.EnvironmentVariable[element.useEnvironmentVaule][element.EntityLogicalName][
                    element.LogicalName
                  ];
                //benz_excel__formfieldvalue__benz_prototypesalesmeasure__benz_supportperunit
                if (useSupportUnitFormula) {
                  range.values = [[environmentValue]];
                }
              } else {
                range.values = [[element.cell.defalutValues]];
              }
            }
            if (element.cell.numberFormat) {
              range.numberFormat = [[element.cell.numberFormat ?? ""]];
            }

            if (element.cell.formulas) {
              console.log("element.cell.formulas");
              console.log(element.cell.formulas);
              range.formulas = [[element.cell.formulas]];
            }

            // if (element.cell.numberFormat) {
            //   console.log("element.cell.numberFormat");
            //   console.log(element.cell.numberFormat);
            //   range.numberFormat = element.cell.numberFormat;
            // }
            // if (element.type == "label" || element.type == "main-form-field" || element.type == "highlight") {
            if (element.cell.format) {
              if (element.cell.format.fill) {
                if (element.cell.format.fill.color) {
                  range.format.fill.color = element.cell.format.fill.color ?? "white";
                }
              }
              if (element.cell.format.font) {
                if (element.cell.format.font.color) {
                  range.format.font.color = element.cell.format.font.color ?? "white";
                }
                if (element.cell.format.font.size) {
                  range.format.font.size = element.cell.format.font.size ?? 11;
                }
                if (element.cell.format.font.name) {
                  range.format.font.name = element.cell.format.font.name ?? "Arial";
                }
                if (element.cell.format.horizontalAlignment) {
                  range.format.horizontalAlignment =
                    Excel.HorizontalAlignment[element.cell.format.horizontalAlignment] ??
                    Excel.HorizontalAlignment["General"];
                }

                if (element.cell.format.verticalAlignment) {
                  range.format.verticalAlignment =
                    Excel.VerticalAlignment[element.cell.format.verticalAlignment] ?? Excel.VerticalAlignment["Top"];
                }
              }
              range.format.columnWidth = element.cell.format.columnWidth ?? 200;

              // if (element.cell.format.columnWidth) {
              //   range.format.columnWidth = element.cell.format.columnWidth ?? 200;
              // } else {
              //   range.format.autofitColumns();
              // }
            }
            // }

            if (element.cell.dataValidation) {
              let orule = null;
              // console.log("dataValidation.target");
              // console.log(element.cell.dataValidation.target);

              let options = null;
              if (element.cell.dataValidation.target) {
                if (element.cell.dataValidation.errorAlert) {
                  range.dataValidation.errorAlert = element.cell.dataValidation.errorAlert;
                }

                if (element.cell.dataValidation.prompt) {
                  range.dataValidation.prompt = element.cell.dataValidation.prompt;
                }
                options = get_choices_list(element.cell.dataValidation.target);

                if (options) {
                  //.sort()
                  orule = options.map((option) => option.name ?? option).join(",");
                }
                if (orule) {
                  range.dataValidation.rule = {
                    list: {
                      inCellDropDown: true,
                      source: orule,
                    },
                  };
                } else {
                  range.dataValidation.rule = element.cell.dataValidation.rule;
                }
                console.log(orule);
              } else if (element.cell.dataValidation) {
                console.log("element.cell.dataValidation.rule");
                console.log(element.cell.dataValidation.rule);
                range.dataValidation.ignoreBlanks = element.cell.dataValidation.ignoreBlanks ?? true;
                range.dataValidation.rule = element.cell.dataValidation.rule;
              }
            }
            if (element.type == "clear_dataValidation") {
              range.dataValidation.clear();
            }

            if (element.type == "table_formulas" && element.cell.formulas) {
              range.formulas = element.cell.formulas;
            }
            // set wrapText
            if (element.cell.address != "B1") {
              range.format.wrapText = true;
            } else {
              range.format.wrapText = false;
            }
            // if (element.cell.conditionalFormats) {
            //   console.log("conditionalFormats");
            //   try {
            //     let type: string = element.cell.conditionalFormats.ConditionalFormatType;
            //     const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType[type]);

            //     // Set the font of negative numbers to red.
            //     if (element.cell.conditionalFormats.format.fill) {
            //       conditionalFormat.custom.format.fill.color = element.cell.conditionalFormats.format.fill.color;
            //     }
            //     if (element.cell.conditionalFormats.rule) {
            //       conditionalFormat.custom.rule.formula = element.cell.conditionalFormats.rule.formula;
            //     }
            //   } catch (e) {
            //     console.log("conditionalFormats:" + element.id);
            //     console.log(e.message);
            //     console.log(e.stack);
            //   }
            // }
            console.log("mode check");
            console.log(global.mode);
            if (global.mode == Benz_type.AppMode.MIGRATION && element.id == "MemoID") {
              element["readonly"] = false;
              range.format.protection.locked = false;
            } else {
              range.format.protection.locked = element.cell?.format?.protection?.locked ?? true;
            }
            console.log(element.LogicalName);
            if (
              element.type == "main-form-field" &&
              (element["action"].includes("create") || element["action"].includes("update")) &&
              !element["readonly"]
            ) {
              if (!element.MemoNameGen) {
                range.format.protection.locked = false;
                console.log("main-form-field");
              }
            }
          } catch (e) {
            console.error("Error: set_layout; ");
            console.error(element);
            console.error(e.message);
            console.error(e.stack);
          }
          await context.sync();
        }
      } catch (e) {
        console.error(e.stack);
      }
      // await context.sync();

      console.log("end set layout");

      return true;
    });
  } catch (e) {
    console.error(e.stack);
  }
}

export async function set_layout_range(address: any, element: any) {
  Excel.run(async function (context) {
    console.log("set_layout_range");
    try {
      let sheet = context.workbook.worksheets.getActiveWorksheet();

      console.log(address);
      console.log(element);
      let range: Excel.Range = sheet.getRange(address);
      try {
        if (!address) {
          throw new Error("can't find range");
        }
        let cell: Cell = element.cell;
        // console.log(global.CurrentConfig["action"]);
        // if (!element.action.includes(global.CurrentConfig["action"])) {
        //   throw new Error("This element config no eq action");
        // }
        if (cell.values) {
          range.values = [[cell.values]];
        }

        if (cell.format) {
          if ("fill" in cell.format) {
            if (cell.format.fill.color) {
              range.format.fill.color = cell.format.fill.color ?? "white";
            }
          }
          if ("font" in cell.format) {
            if (cell.format.font.color) {
              range.format.font.color = cell.format.font.color ?? "white";
            }
            if (cell.format.font.size) {
              range.format.font.size = cell.format.font.size ?? 11;
            }
            if (cell.format.font.name) {
              range.format.font.name = cell.format.font.name ?? "Arial";
            }
          }
        }

        if (cell.dataValidation) {
          let orule = null;
          // console.log("dataValidation.target");
          // console.log(cell.dataValidation.target);

          let options = null;
          if (cell.dataValidation.target) {
            options = get_choices_list(cell.dataValidation.target);

            if (options) {
              //.sort()
              orule = options.map((option) => option.name ?? option).join(",");
            }
            if (orule) {
              range.dataValidation.rule = {
                list: {
                  inCellDropDown: true,
                  source: orule,
                },
              };
            } else {
              range.dataValidation.rule = cell.dataValidation.rule;
            }
            console.log(orule);
          } else if (cell.dataValidation) {
            range.dataValidation.rule = cell.dataValidation.rule;
          }
        }
        if (element.type == "clear_dataValidation") {
          range.dataValidation.clear();
        }

        range.formulas = cell.formulas ?? [[""]];

        // if (cell.conditionalFormats) {
        //   console.log("conditionalFormats");
        //   try {
        //     let type: string = cell.conditionalFormats.ConditionalFormatType;
        //     const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType[type]);

        //     // Set the font of negative numbers to red.
        //     if ("format" in cell.conditionalFormats) {
        //       if (cell.conditionalFormats.format) {
        //         conditionalFormat.custom.format.fill.color = cell.conditionalFormats.format.fill.color ?? "";
        //       }
        //     }
        //     if ("rule" in cell.conditionalFormats) {
        //       conditionalFormat.custom.rule.formula = cell.conditionalFormats.rule.formula ?? [[""]];
        //     }
        //   } catch (e) {
        //     console.log("conditionalFormats:" + element.id);
        //     console.log(e.message);
        //     console.log(e.stack);
        //   }
        // }
        // if (element.type == "column") {
        //   range.format.columnWidth = cell.format.columnWidth;
        // }

        // if (cell.format?.protection) {
        //   range.format.protection.locked = cell.format.protection.locked;
        // }
        // if (
        //   element.type == "main-form-field" &&
        //   (global.CurrentConfig["action"] == "create" || global.CurrentConfig["action"] == "update") &&
        //   !global.CurrentConfig["readonly"]
        // ) {
        //   range.format.protection.locked = false;
        //   console.log("main-form-field");
        // }
      } catch (e) {
        console.error("Error set_layout_range:" + e.stack);
        console.error(e.stack);
        // console.log(e.stack);
      }
      console.log("end set layout_range");
      console.log(range);

      return await context.sync().then(() => {
        console.log("end set layout_range");
        console.log(range);
      });
    } catch (e) {
      console.error(e.stack);
    }
  });
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function search(key, nameKey, myArray) {
  for (let i = 0; i < myArray.length; i++) {
    if (myArray[i][key] === nameKey) {
      return myArray[i];
    }
  }
}
export async function set_table(sheet: Excel.Worksheet, formConfig: FormConfig) {
  try {
    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    var cols_n: Excel.TableColumn[] = [];
    const a = await Excel.run(async function (context) {
      console.log("set table");
      console.log(formConfig);
      const config = formConfig.table;
      console.log(config);
      try {
        if (config.columns) {
          let header = config.columns.filter((ele) => ele.enable).map((ele) => ele.Label);
          console.log(header);
          console.log(header.length);
          var sheet = context.workbook.worksheets.getActiveWorksheet();

          let range = sheet.getRangeByIndexes(config.rowIndex, config.columnIndex, 2, header.length);
          let expensesTable: Excel.Table = null;
          try {
            expensesTable = sheet.tables.add(range, config.hasHeaders /*hasHeaders*/);
            // await lockmainObjectValueToCell(false);

            await context.sync();
            // await lockmainObjectValueToCell(true);
          } catch (e) {
            expensesTable = sheet.tables.getItemAt(0);
          }
          expensesTable.name = config.name;

          let formats = header.map(() => "text");
          console.log(formats);
          expensesTable.getHeaderRowRange().numberFormat = [formats];
          expensesTable.getHeaderRowRange().values = [header];
          if (["create", "update"].includes(global.CurrentConfig["action"]) && expensesTable.getDataBodyRange()) {
            expensesTable.getDataBodyRange().format.protection.locked = false;
            console.log("set table protection false");
          }
          await context.sync();
        }
        return await context.sync().then(async () => {
          console.log("end set layout_range");
          return true;
        });
      } catch (e) {
        console.error("Error: set_table");
        console.error(e.stack);
      }
    });
    return a;
  } catch (e) {
    console.error("set_table error");
    console.error(e.stack);
  }
}
export async function hideUnhideColumnAsync(hidden) {
  console.log("hideUnhideColumnAsync...", hidden);
  Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange("A:A").setColumnProperties([{ columnHidden: true }]);
    sheet.getRange("A:A").columnHidden = true;
    sheet.getRange("A:A").format.columnWidth = 0;

    await context.sync();
  });
}
export async function set_tablebody(sheet: Excel.Worksheet, config: Table) {
  try {
    Excel.run(async function (context) {
      console.log("set table body");
      // console.log(config);
      try {
        let expensesTable = sheet.tables.getItemAt(0);
        expensesTable.name = config.name;
        let body = config.columns.filter((ele) => ele.enable).map((ele) => ele.values);
        console.log(body);
        console.log(body.length);

        expensesTable.getDataBodyRange().values = [body];
        if (["create", "update"].includes(global.CurrentConfig["action"])) {
          expensesTable.getDataBodyRange().format.protection.locked = false;
          console.log("set table protection false");
        }
        return await context.sync().then(() => {
          console.log("end set table body");
        });
      } catch (e) {
        console.log("set_table body error");
        console.error(e.stack);
      }
    });
  } catch (e) {
    console.log("set_table body error");
    console.error(e.stack);
  }
}

export function set_columeRule(sheet: Excel.Worksheet, config: Table) {
  try {
    Excel.run(async function (context) {
      console.log("set columeRule");
      // console.log(config);
      try {
        if (!config.columns) {
          throw new Error("can't get columns");
        }
        let expensesTable = sheet.tables.getItem(config.name);
        expensesTable.columns.load("items");

        await context.sync();

        expensesTable.columns.items[0].name = "Purchase date";

        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();

        await context.sync();

        //   config.columns.forEach(async (element: TableColumns, index: number) => {
        //     if (element.formulas) {
        //         let columns = expensesTable.columns;
        //         console.log(columns);
        //     //   .items[index].getDataBodyRange();
        //     //   rag.load("address");
        //     //   await context.sync().then(() => {
        //     //     console.log(rag.address);
        //     //   });

        //       //   rag.formulas = [[element.formulas]];
        //       //   rag.format.autofitColumns();
        //     }
        //     // await context.sync().then(async () => {
        //     //     rag.load("formulas");
        //     //     await context.sync();
        //     //     console.log(rag.formulas);
        //     //   });
        //   });
        // getDataBodyRange
      } catch (e) {
        console.error(e.stack);
      }
    });
  } catch (e) {
    console.error(e.stack);
  }
}

export function getActiveWorksheet() {
  return Excel.run((context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");
    return context.sync().then(() => {
      console.log("The active worksheet is " + sheet.name);
      return sheet.name;
    });
  }).catch(function (error) {
    console.log(error.debugInfo);
    return "";
  });
}

export function get_mainformfield(sheet: Excel.Worksheet, config: FormLayout[]) {
  try {
    Excel.run(async function (context) {
      console.log("set_info");
      //   try {
      //     var rang = sheet.getUsedRange();
      //     rang.format.protection.locked = true;

      //     // rang.load("address");
      //     //   rang.load(["address", "format/protection/locked"]);
      //     await context.sync().then(() => {
      //     //   console.log(`The range address is: ` + rang.address);
      //     });
      //   } catch (e) {
      //     console.error(e.stack);
      //   }

      // console.log(config);'
      try {
        for (let element of config) {
          try {
            //   console.log(element.title_address);
            // console.log(element.cell.format.text);
            // console.log(element.cell.format.address);
            // console.log(element.cell.format.fill_color);
            //   let sheet = context.workbook.worksheets.getItem(config.module);

            let range: Excel.Range = null;
            // console.log(element.cell.address);
            if (!element.cell.address) {
              throw new Error("can't find range");
            }
            range = sheet.getRange(element.cell.address);
            // if (element.type == "column" ) {
            //     range.format.columnWidth = element.columnWidth;
            // }
            // console.log(element.id);
            // console.log(element.cell.value);

            if (element.cell.values) {
              range.values = [[element.cell.values]];
            }
            if (element.cell.format) {
              if (element.cell.format.fill) {
                if (element.cell.format.fill.color) {
                  range.format.fill.color = element.cell.format.fill.color ?? "white";
                }
              }
              if (element.cell.format.font) {
                if (element.cell.format.font.color) {
                  range.format.font.color = element.cell.format.font.color ?? "white";
                }
                if (element.cell.format.font.size) {
                  range.format.font.size = element.cell.format.font.size ?? 11;
                }
                if (element.cell.format.font.name) {
                  range.format.font.name = element.cell.format.font.name ?? "Arial";
                }
              }
            }

            if (element.cell.dataValidation) {
              let orule = null;
              console.log("dataValidation.target");
              console.log(element.cell.dataValidation.target);

              let options = null;
              if (element.cell.dataValidation.target) {
                options = get_choices_list(element.cell.dataValidation.target);
                if (options) {
                  //.sort()
                  orule = options.map((option) => option.name ?? option).join(",");
                }
                if (orule) {
                  range.dataValidation.rule = {
                    list: {
                      inCellDropDown: true,
                      source: orule,
                    },
                  };
                } else {
                  range.dataValidation.rule = element.cell.dataValidation.rule;
                }
                console.log(orule);
              }
            }
            if (element.type == "clear_dataValidation") {
              range.dataValidation.clear();
            }

            if (element.type == "table_formulas" && element.cell.formulas) {
              range.formulas = element.cell.formulas;
            }
            if (element.type == "column") {
              range.format.columnWidth = element.cell.format.columnWidth;
            }
            if (element.cell.format?.protection) {
              range.format.protection.locked = element.cell.format.protection.locked;
            }
            if (element.type == "main-form-field" && global.CurrentConfig["action"] == "create") {
              range.format.protection.locked = false;
              console.log("main-form-field");
            }

            await context.sync();
          } catch (e) {
            console.log("set_layout:" + element.id);
            console.log(e.message);
            console.error(e.stack);
          }
        }
      } catch (e) {
        console.error(e.stack);
      }
    });
  } catch (e) {
    console.error(e.stack);
  }
}

// Function to customize the selected cells
// function customizeSelectedCells() {
//   Excel.run(function (context) {
//     // Get the selected range
//     var range = context.workbook.getSelectedRange();

//     // Load the values and format of the range
//     range.load(["values", "format"]);

//     return context
//       .sync()
//       .then(function () {
//         // Iterate through each cell in the range
//         for (var row = 0; row < range.values.length; row++) {
//           for (var col = 0; col < range.values[row].length; col++) {
//             var cellValue = range.values[row][col];

//             // Customize the cell based on the cell value
//             if (cellValue > 0) {
//               range.getCell(row, col).format.fill.color = "green";
//             } else if (cellValue < 0) {
//               range.getCell(row, col).format.fill.color = "red";
//             } else {
//               range.getCell(row, col).format.fill.color = "yellow";
//             }
//           }
//         }
//       })
//       .then(context.sync);
//   }).catch(function (error) {
//     console.log(error);
//   });
// }
function pad(num, size) {
  num = num.toString();
  while (num.length < size) num = "0" + num;
  return num;
}

async function get_memo_no(res) {
  try {
    console.log("get_memo_no...");
    console.log(res);
    return res["@odata.count"] + 1;
  } catch (e) {
    console.error("get_memo_no error");
  }
}
// eslint-disable-next-line @typescript-eslint/no-unused-vars
async function set_memo_no(num) {
  try {
    console.log("get_memo_no...");
    console.log(num);
    global.MemoNo = num + 1;
    Excel.run(async (context) => {
      for (let field of global.CurrentConfig.layout.filter((ele) => ele.LogicalName == "benz_memonoserialno")) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        let range = sheet.getRange(field.cell.address);
        range.values = [[global.MemoNo.toString()]];
      }
      await context.sync();
    });
  } catch (e) {
    console.error("set_memo_no error");
  }
}
// async function set_memoid(res) {
//   try {
//     console.log("get_memo_no...");
//     console.log(res);
//     console.log(res["@odata.count"]);
//     global.MemoNo = res["@odata.count"] + 1;
//     Excel.run(async (context) => {
//       global.CurrentConfig.layout
//         .filter((ele) => ele.LogicalName == "benz_memonoserialno")
//         .forEach(async (field) => {
//           var sheet = context.workbook.worksheets.getActiveWorksheet();
//           let range = sheet.getRange(field.cell.address);
//           range.values = [[global.MemoNo.toString()]];
//         });
//       await context.sync();
//     });
//   } catch (e) {
//     console.error("get_memo_no error");
//   }
// }
export async function getready_main_data_by_object(config: FormConfig, entity: string, afterAction, callback) {
  try {
    var mainFormObject = {};
    var chObject = {};
    console.log("getready_main_data_by_object");
    console.log(config);
    if (!config) {
      new Error("can't get main data object, can't find config");
    }
    await Excel.run(async (context) => {
      console.log(config.layout.filter((ele) => ele.type == "main-form-field")); // && ele.enable
      for (let field of config.layout.filter((ele) => ele.type == "main-form-field")) {
        // && ele.enable
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        let range = sheet.getRange(field.cell.address);
        range.load("values");
        mainFormObject[field.LogicalName] = { Field: field, Range: range };
        console.log(mainFormObject);
      }
      await context.sync();
    });
    console.log("mainFormObject making...");
    console.log(mainFormObject);
    for (let key of Object.keys(mainFormObject)) {
      const cell = mainFormObject[key].Field.cell;
      console.log(key);
      console.log(mainFormObject[key].Field);
      console.log(mainFormObject[key].Range);
      let val = mainFormObject[key].Range.values[0][0];

      console.log("mode check");
      console.log(global.mode);
      if (global.mode == Benz_type.AppMode.MIGRATION) {
        config["memoid_set"] = ["benz_name"];
      }

      if (mainFormObject[key].Field.cell.pad) {
        val = pad(mainFormObject[key].Range.values[0][0], mainFormObject[key].Field.cell.pad);
      }
      if (config["memoid_set"].includes(mainFormObject[key].Field.LogicalName)) {
        chObject[mainFormObject[key].Field.LogicalName] = val.toString();
      }
      console.log(
        mainFormObject[key].Field.cell?.WorksheetSingleClicked?.find((e) => e.type === "tabledata_lookupsystemuser")
      );
      if (mainFormObject[key].Field.AttributeType.toString().toLowerCase() != "string") {
        if (mainFormObject[key].Field.cell.dataValidation?.target ?? null) {
          console.log("dataValidation target");
          console.log(mainFormObject[key].Field.cell.dataValidation.target);
          console.log(get_choicesSets(mainFormObject[key].Field.cell.dataValidation.target));
          mainFormObject[key] = get_choicesSets(
            mainFormObject[key].Field.cell.dataValidation.target.toLowerCase()
          )?.find((e) => e.name === val.toString())?.value;
          console.log(mainFormObject[key]);
        } else if (
          mainFormObject[key].Field.AttributeType.toString().toLowerCase() == "lookup"
          // mainFormObject[key].Field.cell?.WorksheetSingleClicked?.find(
          //   (e) => e.type === "tabledata_lookupsystemuser"
          // ) ??
          // null
        ) {
          console.log("lookup");
          console.log(mainFormObject[key].Field.ValueLogicalName);
          console.log(
            config.relationship.find((e) => e.from_FieldLogicalName === mainFormObject[key].Field.LogicalName)
          );
          let id = null;
          try {
            id = get_choicesSets(mainFormObject[key].Field.LogicalName)?.find((e) => e.name === val.toString())?.value;
          } catch (e) {
            console.log("Error lookup");
          }
          console.log("id");
          console.log(id);

          if (!id) {
            id = get_choicesSets(mainFormObject[key].Field.LogicalName)?.find(
              (e) => e.name === mainFormObject[key].Field.outOfTarget.value
            )?.value;
            mainFormObject[key] = val.toString();
          }
          mainFormObject[`${mainFormObject[key].Field.SchemaName}@odata.bind`] = `${
            config.relationship.find((e) => e.from_FieldLogicalName === mainFormObject[key].Field.LogicalName)
              .to_LogicalName
          }s(${id})`;
          delete mainFormObject[key];
          console.log(mainFormObject);
        } else if (mainFormObject[key].Field.AttributeType.toString().toLowerCase() == "number") {
          mainFormObject[key] = Number(val.toString());
        } else if (mainFormObject[key].Field.AttributeType.toString().toLowerCase() == "boolean") {
          mainFormObject[key] = mainFormObject[key].Field.cell.valuesMap[val.toString()];
          // mainFormObject[key] = val.toString().toLowerCase() == "yes" ? true : false;
        }
      } else {
        mainFormObject[key] = val;
      }
      // convert datatype
      console.log(mainFormObject[key]);
      if (cell.inputtype == "text" && mainFormObject[key]) {
        mainFormObject[key] = mainFormObject[key].toString();
      }
    }
    if (config["main_hardcode"]) {
      mainFormObject = add_hardcode(mainFormObject, config["main_hardcode"]);
    }
    console.log("get_main_data_by_object");
    console.log(mainFormObject);

    global.Callapiaction = {
      name: "callapiaction",
      action: {
        entitySet: "benz_prototypesalesmeasuremainforms",
        queryString: `?$select=benz_prototypesalesmeasuremainformid&$filter=contains(benz_memonoyear,'${
          mainFormObject["benz_memonoyear"]
        }') and contains(benz_memonomonth,'${mainFormObject["benz_memonomonth"]}') and ${
          config["memotype_LogicalName"]
        } eq ${mainFormObject[config["memotype_LogicalName"]]}&$count=true`,
        queryOptions: "",
      },
    };

    return await datahelper.get_Count__byModth(get_memo_no).then(async () => {
      console.log("after get_memo_no");
      if (global.MemoNo) {
        mainFormObject["benz_memonoserialno"] = global.MemoNo.toString();
        chObject["benz_memonoserialno"] = global.MemoNo.toString();
        const memoid_set = config["memoid_set"].map((key) => (chObject[key] ?? key).toString()).join("");
        console.log(memoid_set);
        mainFormObject["benz_name"] = memoid_set;
        return await afterAction(entity, mainFormObject, config, callback);
      } else {
        new Error("Memo no is null");
      }
    });
  } catch (error) {
    console.log(error);
  }
}

export async function get_cell_value(address: string) {
  return await Excel.run(async (context) => {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    let range = sheet.getRange(address);
    range.load("values");
    await context.sync();
    return range.values[0][0];
  });
}

export async function get_main_data_by_object(config: FormConfig, fields: FormLayout[], actions: any[]) {
  try {
    var mainFormObject: mainformobjectKey = {};
    var mainFormObjectValue = {};
    var mainObject = {};
    console.log("getready_main_data_by_object");
    console.log(config);
    if (!config) {
      new Error("can't get main data object, can't find config");
    }
    await Excel.run(async (context) => {
      console.log(fields); //config.layout.filter((ele) => ele.type == "main-form-field")); // && ele.enable
      for (let field of fields) {
        //config.layout.filter((ele) => ele.type == "main-form-field")) {
        // && ele.enable
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        let range = sheet.getRange(field.cell.address);
        range.load("values");
        mainFormObject[field.LogicalName] = { Field: field, Range: range };
        mainFormObjectValue[field.LogicalName] = range;
        console.log(mainFormObject);
      }
      await context.sync();
    });
    for (const key of Object.keys(mainFormObjectValue)) {
      // let field = fields.find((e) => e.LogicalName == key);
      // if (field.cell.querys?.length > 1) {
      //   for (let query of field.cell.querys) {
      //     mainFormObjectValue = await querydata.init(query, mainFormObjectValue);
      //   }
      // } else {
      // }
      mainFormObjectValue[key] = mainFormObjectValue[key].values[0][0];

      console.log("querydata mainFormObjectValue");
      console.log(mainFormObjectValue);
    }
    console.log("mainFormObject making...");
    console.log(mainFormObjectValue);

    mainObject = await ConvertToKey([mainFormObjectValue], config);

    console.log("get_main_data_by_object");
    console.log(mainObject);
    mainObject = mainObject[0];
    if (config["main_hardcode"]) {
      mainObject = add_hardcode(mainObject, config["main_hardcode"]);
    }

    console.log("get_main_data_by_object 2");
    console.log(mainObject);

    for (let action of actions) {
      let obj: queryUpdateObject = { config: config, object: mainObject };
      mainObject = await action(obj);
    }
    console.log("get_main_data_by_object 3");
    console.log(mainObject);

    return mainObject;
  } catch (e) {
    console.error(e.stack);
    throw new Error(e.message);
  }
}
// export async function dataValidation_main_data_by_object(config, entity: string, afterAction, callback) {
//   try {
//     var mainFormObject = {};
//     var chObject = {};
//     await Excel.run(async (context) => {
//       config.layout
//         .filter((ele) => ele.type == "main-form-field" && !ele.disable)
//         .forEach(async (field) => {
//           var sheet = context.workbook.worksheets.getActiveWorksheet();
//           let range = sheet.getRange(field.cell.address);
//           range.load("values");
//           mainFormObject[field.LogicalName] = { Field: field, Range: range };
//         });
//       await context.sync();
//     });
//     console.log("mainFormObject making...");
//     Object.keys(mainFormObject).forEach((key) => {
//       const cell = mainFormObject[key].Field.cell;
//       console.log(key);
//       console.log(mainFormObject[key].Range.values);
//       console.log(mainFormObject[key].Range.values[0][0]);
//       let val = mainFormObject[key].Range.values[0][0];

//       if (mainFormObject[key].Field.cell.pad) {
//         val = pad(mainFormObject[key].Range.values[0][0], mainFormObject[key].Field.cell.pad);
//       }
//       if (config["memoid_set"].includes(mainFormObject[key].Field.LogicalName)) {
//         chObject[mainFormObject[key].Field.LogicalName] = val.toString();
//       }
//       if (mainFormObject[key].Field.cell.dataValidation) {
//         mainFormObject[key] = get_choices(mainFormObject[key].Field.cell.dataValidation.target, val.toString());
//       } else {
//         mainFormObject[key] = val;
//       }
//       // convert datatype
//       if (cell.inputtype == "text") {
//         console.log(mainFormObject[key]);
//         mainFormObject[key] = mainFormObject[key].toString();
//       }
//     });
//     if (global.CurrentConfig["main_hardcode"]) {
//       mainFormObject = add_hardcode(mainFormObject, global.CurrentConfig["main_hardcode"]);
//     }
//     console.log("get_main_data_by_object");
//     console.log(mainFormObject);

//     global.Callapiaction = {
//       name: "callapiaction",
//       action: {
//         entitySet: "benz_prototypesalesmeasuremainforms",
//         queryString: `?$select=benz_prototypesalesmeasuremainformid&$filter=contains(benz_memonoyear,'${
//           mainFormObject["benz_memonoyear"]
//         }') and contains(benz_memonomonth,'${mainFormObject["benz_memonomonth"]}') and ${
//           config["memotype_LogicalName"]
//         } eq ${mainFormObject[config["memotype_LogicalName"]]}&$count=true`,
//         queryOptions: "",
//       },
//     };

//     return await datahelper.get_Count__byModth(get_memo_no).then(async () => {
//       console.log("after get_memo_no");
//       mainFormObject["benz_memonoserialno"] = global.MemoNo.toString();
//       chObject["benz_memonoserialno"] = global.MemoNo.toString();
//       const memoid_set = config["memoid_set"].map((key) => chObject[key].toString()).join("");
//       console.log(memoid_set);
//       mainFormObject["benz_name"] = memoid_set;
//       return await afterAction(entity, mainFormObject, config, callback);
//     });
//   } catch (error) {
//     console.log(error);
//   }
// }
export async function get_table_data_arrary(config) {
  console.log("get_table_data_arrary...");
  var items = [];
  try {
    let _range: Excel.Range = null;
    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getItem(config.module);
      let table = sheet.tables.getItemAt(0);

      let range = table.getDataBodyRange();
      range.load("values");
      _range = range;
      await context.sync();
    });
    console.log("get_table_data_arrary");
    console.log(items);
    // return ConvertToKey(items, config);
    return _range.values;
    // afterAction(entity, items, config, callback);
  } catch (error) {
    console.log(error);
    return null;
  }
}

export async function get_table_data_obj(config) {
  console.log("get_table_data_obj...");
  var items = [];
  try {
    let _range = null;
    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getItem(config.module);
      let table = sheet.tables.getItemAt(0);

      let range = table.getRange();
      range.load("values");
      _range = range;
      await context.sync();
    });
    var tableHeader = _range.values[0];
    // console.log("tableHeader");
    // console.log(tableHeader);
    var tableDataLine = _range.values;

    // Convert table data to key-value pairs
    for (var i = 1; i < tableDataLine.length; i++) {
      let keyValuePairs = {};
      for (var j = 0; j < tableHeader.length; j++) {
        var key = tableHeader[j];
        var value = tableDataLine[i][j];
        keyValuePairs[key] = value;
      }
      items.push(keyValuePairs);
    }
    console.log("get_table_data_obj");
    console.log(items);
    // return ConvertToKey(items, config);
    return items;
    // afterAction(entity, items, config, callback);
  } catch (error) {
    console.log(error);
  }
}

export async function getready_table_data_by_object(config, mainobj) {
  console.log("getready_table_data_by_object...");
  var items = [];
  try {
    let _range = null;
    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      // let table = sheet.tables.getItem(config.table.name);
      let table = sheet.tables.getItemAt(0);

      let range = table.getRange();
      range.load("values");
      _range = range;
      await context.sync();
    });
    var tableHeader = _range.values[0];
    // console.log("tableHeader");
    // console.log(tableHeader);
    var tableDataLine = _range.values;

    // Convert table data to key-value pairs
    for (var i = 1; i < tableDataLine.length; i++) {
      let keyValuePairs = {};
      for (var j = 0; j < tableHeader.length; j++) {
        var key = tableHeader[j];
        var value = tableDataLine[i][j];
        keyValuePairs[key] = value;
      }
      items.push(keyValuePairs);
    }
    console.log("get_table_data_by_object");
    console.log(items);
    let nitems = await ConvertToKey(items, config);
    nitems[`${config.Table_Main_SchemaName}@odata.bind`] = `${config.Table_Main_LogicalName}s(${
      mainobj[config.Table_Main_LogicalName + "id"]
    })`;
    console.log(nitems);
    // return ConvertToKey(items, config);
    return nitems;
    // afterAction(entity, items, config, callback);
  } catch (error) {
    console.log(error);
  }
}

export async function post_table_data_by_object(config, entity: string, items, afterAction, callback) {
  console.log("post_table_data_by_object...");
  try {
    afterAction(entity, items, config, callback);
  } catch (e) {
    console.log(e);
  }
}

export function post_Data(entity: string, items: any, config, callback) {
  console.log("post_Data");
  console.log("items", items);
  global.Callapiaction = {
    name: "callapiaction",
    action: {
      entitySet: entity,
      queryString: items,
      queryOptions: "",
    },
  };
  return datahelper.post_Data(callback);
}

export function postDataReturnData(entity: string, items: any) {
  console.log("post_Data");
  console.log("items", items);
  global.Callapiaction = {
    name: "callapiaction",
    action: {
      entitySet: entity,
      queryString: items,
      queryOptions: "",
    },
  };
  return datahelper.post_Data_ReturnData();
}

export function updateDataReturnData(entity: string, id: string, items: any) {
  console.log("update_Data");
  console.log("items", items);
  global.Callapiaction = {
    name: "callapiaction",
    action: {
      entitySet: entity,
      id: id,
      queryString: items,
      queryOptions: "",
    },
  };
  return datahelper.update_Data_ReturnData();
}
export async function ConvertToKey(oitems: any[], config: FormConfig): Promise<object[]> {
  let nitems = [];
  console.log("ConvertToKey...");
  console.log(oitems);
  var countRow = 0;
  for (let element of oitems) {
    let nitem = {};
    console.log(nitem);

    for (const key of Object.keys(element)) {
      // console.log(`${key}:${value}`);
      const value = element[key];
      const target =
        config.layout.find((element) => element.LogicalName == key) ??
        config.table.columns.find((element) => element.Label == key);
      console.log(`${key}: ${value}`);
      console.log(target.LogicalName);
      console.log(target.LogicalName.length);
      console.log(target.LogicalName.length === 0);
      console.log(value);
      console.log(nitem);
      console.log(target);
      try {
        if (target.LogicalName.length !== 0 && value !== "") {
          //.length !== 0
          if (target.cell.dataValidation?.target && target.AttributeType.toString().toLowerCase() == "choice") {
            console.log("dataValidation target");
            console.log(target.cell.dataValidation.target);
            console.log(get_choicesSets(target.cell.dataValidation.target));
            nitem[target.LogicalName] = get_choicesSets(target.cell.dataValidation.target.toLowerCase())?.find(
              (e) => e.name === value.toString()
            )?.value;
          } else if (
            (target.LookupInoutType ?? "").toString().toLowerCase() == "choice" &&
            target.AttributeType.toString().toLowerCase() == "lookup"
          ) {
            console.log("dataValidation target");
            console.log(target.cell.dataValidation.target);
            // console.log(get_choicesSets(target.cell.dataValidation.target));
            let id = get_choicesSets(target.cell.dataValidation.target.toLowerCase())?.find(
              (e) => e.name === value.toString()
            )?.id;
            console.log(id);

            nitem[`${target.SchemaName}@odata.bind`] = `${
              config.relationship.find((e) => e.from_FieldLogicalName === target.LogicalName).to_LogicalName
            }s(${id})`;
          } else if (
            (target.LookupInoutType ?? "").toString().toLowerCase() == "string" &&
            target.AttributeType.toString().toLowerCase() == "lookup"
          ) {
            console.log("Lookup...");

            let id = null;
            try {
              id = get_choicesSets(nitem[key].Field.LogicalName)?.find((e) => e.name === value.toString())?.value;
            } catch (e) {
              console.error("Error lookup");
            }
            console.log("id");
            console.log(id);

            if (!id) {
              id = get_choicesSets(nitem[key].Field.LogicalName)?.find(
                (e) => e.name === nitem[key].Field.outOfTarget.value
              )?.value;
              nitem[key] = value.toString();
            }

            nitem[`${target.SchemaName}@odata.bind`] = `${
              config.relationship.find((e) => e.from_FieldLogicalName === target.LogicalName).to_LogicalName
            }s(${id})`;
            // delete nitem[key];
            console.log(nitem);
          } else if (
            target.AttributeType.toString().toLowerCase() == "lookup" &&
            config.relationship.find((e) => e.from_FieldLogicalName === target.LogicalName).querys?.length > 0
          ) {
            console.log("lookup");
            console.log(target.ValueLogicalName);
            console.log(config.relationship.find((e) => e.from_FieldLogicalName === target.LogicalName));
            try {
              const relationship = config.relationship.find((e) => e.from_FieldLogicalName === target.LogicalName);
              console.log(value.toString().length);

              if (!value) {
                throw new Error("lookup value is null");
              }
              if (value.toString().length < 1) {
                throw new Error("lookup value is empty");
              }
              let lookupid = get_choicesSets(target.LogicalName)?.find((e) => e.name === value.toString())?.value;

              if (target.LogicalName == "benz_prototypesalesmeasure") {
                console.log("Key Value");
                console.log(value);
                console.log(countRow);
                global.RowMeasureName[countRow] = value;
              }

              let val = value.toString();
              if (relationship.querys[0].getvalueregex) {
                const regex = RegExp(relationship.querys[0].getvalueregex, "g");
                val = regex.exec(value.toString())[0];
              }

              if (!lookupid && value && target.LookupAllowUsingExist) {
                global.Callapiaction = {
                  name: "callapiaction",
                  action: {
                    entitySet: relationship.querys[0].EntityLogicalName,
                    queryString: relationship.querys[0].queryString.replace("{valueAfters}", val),
                    queryOptions: "",
                  },
                };
                const res = await datahelper.RetrieveAndReturnMultipleData(null, {
                  entitySet: relationship.querys[0].EntityLogicalName,
                  queryString: relationship.querys[0].queryString.replace("{valueAfters}", val),
                  queryOptions: "",
                });
                console.log(res);
                if (res) {
                  if (res.value.length > 0) {
                    lookupid = res.value[0][`${relationship.to_LogicalName}id`];
                  } else {
                    if (target.LogicalName == "benz_modeldesignation") {
                      console.log("global.choicesSets");
                      console.log(global.choicesSets);
                      console.log(global.choicesSets[target.LookupEntityLogicalName]);
                      var modelSets = global.choicesSets[target.LookupEntityLogicalName];
                      var currentModelSet = [];
                      for (let modelSet of modelSets) {
                        if (modelSet.name == val) {
                          currentModelSet = modelSet;
                          break;
                        }
                      }
                      console.log(currentModelSet);
                      lookupid = currentModelSet["id"];
                    }
                  }
                }
              }

              // if (target.LogicalName == "benz_commissionnumber" && global.updatingRecord != "") {
              //   var targetQuertString = relationship.querys[0].queryString.replace("{valueAfters}", val);
              //   if (relationship.querys[0].EntityLogicalName == "benz_commissionnumbers") {
              //     // var claimLineItem = nitem["benz_id"];
              //     var measureName = nitem["benz_prototypesalesmeasure"];
              //     console.log(measureName);
              //     targetQuertString = targetQuertString.replace("{measureID}", measureName);
              //   }
              //   global.Callapiaction = {
              //     name: "callapiaction",
              //     action: {
              //       entitySet: relationship.querys[0].EntityLogicalName,
              //       queryString: targetQuertString,
              //       queryOptions: "",
              //     },
              //   };
              //   const res = await datahelper.RetrieveAndReturnMultipleData(null, {
              //     entitySet: relationship.querys[0].EntityLogicalName,
              //     queryString: targetQuertString,
              //     queryOptions: "",
              //   });
              //   console.log(res);
              //   if (res) {
              //     lookupid = res.value[0][`${relationship.to_LogicalName}id`];
              //   }
              //   console.log("lookupid commission number");
              //   console.log(lookupid);
              // }

              if (lookupid) {
                nitem[`${target.SchemaName}@odata.bind`] = `${relationship.to_LogicalName}s(${lookupid})`;
                delete nitem[key];
                console.log(nitem);
              } else {
                nitem[target.LogicalName] = value;
              }
              console.log("lookup res");
              console.log(nitem);
            } catch (e) {
              console.error(e.stack);
            }
          } else if (target.AttributeType.toString().toLowerCase() == "string") {
            nitem[target.LogicalName] = value.toString();
          } else if (target.AttributeType.toString().toLowerCase() == "boolean") {
            console.log(value.toString());
            console.log(target.cell.valuesMap);
            console.log(target.cell.valuesMap[value.toString()]);
            nitem[target.LogicalName] = target.cell.valuesMap[value.toString()];

            // nitem[target.LogicalName] = value.toString().toLowerCase() == "yes" ? true : false;
          } else if (target.AttributeType.toString().toLowerCase() == "number") {
            nitem[target.LogicalName] = Number(value.toString());
          } else if (target.AttributeType.toString().toLowerCase() == "percentage") {
            nitem[target.LogicalName] = Number(value * 100);
          } else if (
            target.AttributeType.toString().toLowerCase() == "datetime" ||
            target.AttributeType.toString().toLowerCase() == "date"
          ) {
            let val = value as number;
            let javaScriptDate = new Date(Math.round((val - 25569) * 86400 * 1000));
            nitem[target.LogicalName] = new Date(javaScriptDate).toISOString();
          } else {
            nitem[target.LogicalName] = value;
          }
          // console.log(key + ":" + target.LogicalName + ":" + nitem[target.LogicalName]);
        }

        if (target.LogicalName.length !== 0 && value == "" && global.updating) {
          // if (target.AttributeType.toString().toLowerCase() == "string") {
          //   nitem[target.LogicalName] = null;
          // }
          // if (target.AttributeType.toString().toLowerCase() == "number") {
          //   nitem[target.LogicalName] = Number(0);
          // }
          // if (target.AttributeType.toString().toLowerCase() == "datetime" ||
          //   target.AttributeType.toString().toLowerCase() == "date") {
          //   nitem[target.LogicalName] = null;
          // }
          // if (target.AttributeType.toString().toLowerCase() == "percentage") {
          //   nitem[target.LogicalName] = Number(0);
          // }
          if (
            target.AttributeType.toString().toLowerCase() == "percentage" ||
            target.AttributeType.toString().toLowerCase() == "number"
          ) {
            nitem[target.LogicalName] = Number(null);
          } else {
            nitem[target.LogicalName] = null;
          }
        }
      } catch (e) {
        console.error(e.stack);
      }
    }
    nitem;
    nitems.push(nitem);
    countRow++;
  }
  return nitems;
}

export function mainObjectValueToCell(config, obj) {
  console.log("mainObjectValueToCell");
  try {
    Excel.run(async function (context) {
      console.log("set_info");
      try {
        for (let element of config) {
          try {
            let range = null;
            // console.log(element.cell.address);
            // console.log(obj[element.LogicalName]);
            if (!element.cell.address) {
              throw new Error("can't find range");
            }
            var sheet = context.workbook.worksheets.getActiveWorksheet();

            range = sheet.getRange(element.cell.address);
            range.values = [[getObjectValueByDataverseEntity(element, obj, false)]];
          } catch (e) {
            console.log("set_value:" + element.id);
            console.log(e.message);
            console.error(e.stack);
          }
        }
      } catch (e) {
        console.error(e.stack);
      }
      await context.sync();
    });
  } catch (e) {
    console.error(e.stack);
  }
}

export async function lockmainObjectValueToCell(condition: boolean) {
  try {
    return Excel.run(async function (context) {
      console.log(`lockmainObjectValueToCell ${condition}`);
      let workbook = context.workbook;
      var sheet = workbook.worksheets.getActiveWorksheet();
      if (condition) {
        sheet.protection.protect(null, "1234");
        // workbook.protection.protect();
      } else {
        sheet.protection.unprotect("1234");
        // workbook.protection.unprotect();
      }
      try {
        await context.sync();
        sheet.load(["protection/isPaused"]);
        await context.sync();
        return sheet.protection.isPaused;
      } catch (e) {
        console.error(e.stack);
      }
    });
  } catch (e) {
    console.error(e.stack);
  }
}
export function getObjectValueByDataverseEntity(ele, obj, checkSupportType) {
  console.log(obj);
  console.log(ele.LogicalName);
  console.log(obj[ele.LogicalName]);

  let val = null;
  if (ele.relationship) {
    const key = ele.LogicalName + "@OData.Community.Display.V1.FormattedValue";
    val = obj[key] ?? "";
  } else if (obj[ele.LogicalName] == null) {
    val = "";
  } else if (ele.AttributeType.toString().toLowerCase() == "choice") {
    val = obj[ele.LogicalName + "@OData.Community.Display.V1.FormattedValue"] ?? "";
    if (val == "") {
      val = obj[ele.LogicalName] ?? "";
    }
  } else if (
    (ele.LookupInoutType ?? "").toString().toLowerCase() == "choice" &&
    ele.AttributeType.toString().toLowerCase() == "lookup"
  ) {
    val = obj[ele.LogicalName + "@OData.Community.Display.V1.FormattedValue"] ?? "";
  } else if (
    (ele.LookupInoutType ?? "").toString().toLowerCase() == "string" &&
    ele.AttributeType.toString().toLowerCase() == "lookup"
  ) {
    val = obj[ele.LogicalName + "@OData.Community.Display.V1.FormattedValue"] ?? "";
  } else if (ele.AttributeType.toString().toLowerCase() == "string") {
    val = obj[ele.LogicalName].toString() ?? "";
  } else if (ele.AttributeType.toString().toLowerCase() == "boolean") {
    for (const [key, conditional] of Object.entries(ele.cell.valuesMap)) {
      if (conditional == obj[ele.LogicalName]) {
        val = key;
        break;
      }
    }
    val = val ?? "";
  } else if (ele.AttributeType.toString().toLowerCase() == "number") {
    val = obj[ele.LogicalName] ?? "";
  } else if (ele.AttributeType.toString().toLowerCase() == "datetime") {
    val = new Date(obj[ele.LogicalName]).toLocaleString() ?? "";
  } else if (ele.AttributeType.toString().toLowerCase() == "date") {
    // const year = date.getFullYear();
    const dateString = obj[ele.LogicalName];
    const formattedDate = dateString.slice(0, 10);
    const dateArr = formattedDate.split("-");

    // const date = new Date(obj[ele.LogicalName]);
    const date = new Date(formattedDate + " 15:00:00 GMT+0800");
    console.log("change date");
    console.log(date);
    console.log(obj[ele.LogicalName]);
    val =
      date.toLocaleDateString("en", {
        day: "numeric",
        month: "short",
        year: "numeric",
      }) ?? "";
    // val = `${day}/${month}/${year}`;
    // val = date.toLocaleDateString();
    console.log(val);
  } else {
    val = obj[ele.LogicalName] ?? "";
  }

  if (ele.AttributeType == "percentage") {
    val = val / 100;
  }

  if (ele.AttributeType.toString().toLowerCase() == "lookup" && val == "") {
    if (ele.LogicalName == "benz_nameofpreparer") {
      let response: any;
      const userName = obj["_" + ele.LogicalName + "_value@OData.Community.Display.V1.FormattedValue"] ?? "";
      const userEmail = obj["benz_NameofPreparer"]["internalemailaddress"];
      val = userName + "<" + userEmail + ">";
    } else {
      val = obj["_" + ele.LogicalName + "_value@OData.Community.Display.V1.FormattedValue"] ?? "";
    }
  }

  if (ele.isFormula) {
    console.log("environment_name check");
    console.log(val);
    // console.log(ele.useEnvironmentVaule);
    if (ele.useEnvironmentVaule) {
      var environmentName = global.EnvironmentVariable[ele.useEnvironmentVaule][ele.EntityLogicalName][ele.LogicalName];
      //benz_excel__formfieldvalue__benz_prototypesalesmeasure__benz_supportperunit
      if (checkSupportType) {
        val = environmentName;
      }
    } else {
      val = ele.cell.defalutValues;
    }
  }

  return val ?? "";
}

export function ObjectToTableArray(config, objects) {
  try {
    let datas = [];
    for (let obj of objects) {
      let data = [];
      var checkSupportType = false;
      for (let ele of config) {
        try {
          if (ele.LogicalName == "benz_prototypesupporttype") {
            var supportTypeVal = getObjectValueByDataverseEntity(ele, obj, false);
            console.log("supportTypeVal", supportTypeVal);
            if (supportTypeVal.startsWith("1) ")) {
              checkSupportType = true;
            }
          }
          data.push(getObjectValueByDataverseEntity(ele, obj, checkSupportType));
        } catch (e) {
          console.log("set_value:" + ele.id);
          console.log(e.message);
          console.error(e.stack);
        }
      }
      datas.push(data);
    }
    return datas;
  } catch (e) {
    console.error(e.stack);
  }
}

export async function setDataToTable(config, arr_data) {
  console.log("setDataToTable");
  console.log(arr_data);
  if (arr_data) {
    await Excel.run(async (context) => {
      try {
        let sheet = context.workbook.worksheets.getActiveWorksheet();
        let expensesTable = sheet.tables.getItem(config.table.name);

        if (expensesTable) {
          expensesTable.rows.add(0 /*add rows to the end of the table*/, arr_data);
        }

        sheet.activate();

        await context.sync();
      } catch (e) {
        console.error(e.stack);
      }
    });
  }
}

export async function addTableRow() {
  console.log("addRowToTable");
  let module = global.CurrentConfig.module.split("-")[0].toLocaleLowerCase();
  let fileName = global.CurrentConfig.module.split("-")[1].toLocaleLowerCase();
  var c: FormConfig = require(`./config/${module}/${fileName}.json`);

  await Excel.run(async (context) => {
    try {
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      let expensesTable = sheet.tables.getItemAt(0); //.getItem(config.table.name);
      // let expensesTable = sheet.tables.getItem(global.CurrentConfig.table.name);
      console.log("global.CurrentConfig.table");
      console.log(global.CurrentConfig.module);

      let lastRowRecord = await getLastRowRecord();
      var newRowRecord = [];

      // let companyName = sheet.tables.getItem(config.table.name);
      // console.log(companyName);
      console.log("config.table");
      console.log(c.table);

      await lockmainObjectValueToCell(false);
      if (lastRowRecord) {
        console.log(lastRowRecord);
        // console.log(expensesTable.columns);
        // var currentFormColumn = global.CurrentConfig.table.columns;
        for (var i = 0; i < lastRowRecord.length; i++) {
          if (c.table.columns[i]["LogicalName"] == "benz_id") {
            global.lastnumber = lastRowRecord[i] + 1;
            newRowRecord.push(global.lastnumber);
          } else {
            // console.log(c.table.columns[i].cell['defalutValues']);
            newRowRecord.push(c.table.columns[i].cell["defalutValues"]);
          }
        }
        console.log("newRowRecord", newRowRecord);
        expensesTable.rows.add(null, [newRowRecord], true);
      }

      await context.sync();
      await lockmainObjectValueToCell(true);
    } catch (e) {
      console.error(e.stack);
    }
  });
}
export async function addRows(context: Excel.RequestContext, formConfig: FormConfig, num: number) {
  try {
    console.log("addRows", global.CurrentConfig.module);

    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let expensesTable = sheet.tables.getItemAt(0); //.getItem(config.table.name);

    let newRow = [];
    expensesTable.load(["columns", "rows"]);
    // set columns name
    await context.sync();
    let obj = expensesTable.rows.toJSON();
    console.log("expensesTable", obj);

    if (formConfig.table.columns) {
      for (var y = 0; y < num; y++) {
        let newRowRecord = [];

        for (var x = 0; x < formConfig.table.columns.length; x++) {
          if (formConfig.table.columns[x]["LogicalName"] == "benz_id") {
            newRowRecord.push(obj.count + y + 1);
          } else {
            newRowRecord.push(formConfig.table.columns[x].cell["defalutValues"]);
          }
        }
        newRow.push(newRowRecord);
      }
      console.log("newRows", newRow);
      expensesTable.rows.add(null, newRow, true);
    }

    await context.sync();
  } catch (e) {
    console.error(e.stack);
  }
}

export async function addTableRows(args: any) {
  console.log("addTableRows");
  let module = global.CurrentConfig.module.split("-")[0].toLocaleLowerCase();
  let fileName = global.CurrentConfig.module.split("-")[1].toLocaleLowerCase();
  var formConfig: FormConfig = require(`./config/${module}/${fileName}.json`);
  await Excel.run(async (context) => {
    try {
      await lockmainObjectValueToCell(false);
      await addRows(context, formConfig, args.data.num);
      await lockmainObjectValueToCell(true);
    } catch (e) {
      console.error(e.stack);
    }
  });
}

export async function clearSelectTableRow() {
  console.log("clearSelectTableRow");
  await Excel.run(async (context) => {
    try {
      await lockmainObjectValueToCell(false);
      let module = global.CurrentConfig.module.split("-")[0].toLocaleLowerCase();
      let fileName = global.CurrentConfig.module.split("-")[1].toLocaleLowerCase();
      let formConfig: FormConfig = require(`./config/${module}/${fileName}.json`);

      let range = context.workbook.getSelectedRange();
      range.load(["address", "rowCount", "rowIndex", "values", "valuesAsJson"]);
      await context.sync();
      let rangeText = range.valuesAsJson;
      let rangeRowCount = range.rowCount;
      let rangeRowIndex = range.rowIndex;
      console.log(rangeText);
      let newRows = [];
      for (var y = 0; y < rangeRowCount; y++) {
        let newRowRecord = [];
        for (var x = 0; x < formConfig.table.columns.length; x++) {
          if (formConfig.table.columns[x]["LogicalName"] == "benz_id" || formConfig.table.columns[x]["title"] == "id") {
            newRowRecord.push(rangeText[y][x].basicValue ?? "");
          } else {
            newRowRecord.push(formConfig.table.columns[x].cell["defalutValues"] ?? "");
          }
        }
        newRows.push(newRowRecord);
      }
      console.log(newRows);
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      range = sheet.getRangeByIndexes(rangeRowIndex, 0, rangeRowCount, formConfig.table.columns.length);
      range.values = newRows;

      await context.sync();
      await lockmainObjectValueToCell(true);
    } catch (e) {
      console.error(e.stack);
    }
  });
}

export async function removeSelectTableRow() {
  console.log("removeSelectTableRow");
  await Excel.run(async (context) => {
    try {
      await lockmainObjectValueToCell(false);

      let range = context.workbook.getSelectedRange();
      range.delete(Excel.DeleteShiftDirection.up);

      await context.sync();
      await lockmainObjectValueToCell(true);
    } catch (e) {
      console.error(e.stack);
    }
  });
}

export async function remove_tablerows(config) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = sheet.tables.getItem(config.table.name);
    const tableDataRange = expensesTable.getDataBodyRange();
    tableDataRange.load("address");
    await context.sync();

    console.log(tableDataRange.address);

    tableDataRange.delete(Excel.DeleteShiftDirection.up);

    await context.sync();
  });
}

export async function addTableDataChange(config) {
  await Excel.run(async function (context) {
    // document.getElementById("message-area").innerHTML = "<br>writeDataverseDataToOfficeDocument: " ;
    console.log(`addTableDataChange`);
    global.TableDataChangeAction = [];
    try {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const table = sheet.tables.getItem(config.table.name);
      console.log(table);
      table.onChanged.add(async function (eventArgs) {
        console.log(`table onChanged`);
        global.TableDataChangeAction.push(eventArgs);
        console.log(eventArgs);
        if (eventArgs.details !== undefined) {
          // var passConfig: Benz_type.passConfig = require("../benz/config/migration.json");
          // console.log(passConfig.passVaildation);
          // var logicalArr = passConfig.passVaildation.LogicalName;
          // passConfig.passVaildation.LogicalName
          // if (global.mode != Benz_type.AppMode.MIGRATION) {
          onchange_dataValidation.init(eventArgs, config);
          // }
        }
      });
      table.onSelectionChanged.add(async function (eventArgs) {
        console.log(`table onSelectionChanged`);
        // global.TableDataChangeAction.push(eventArgs);
        console.log(eventArgs);
      });

      await context.sync();
    } catch (exception) {
      console.log("addTableDataChange EXCEPTION: " + exception.message);
    }
  });
}

export async function showHideDetaleButton(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  selectedRange: string
) {
  try {
    console.log(`showHideDetaleButton...`);
    let sheetname = "";
    let columnIndex = -1;
    let rowIndex = -1;
    sheet.load(["name"]);
    let range = sheet.getRange(selectedRange);
    let tablerange = sheet.tables.getItemAt(0).getHeaderRowRange();
    tablerange.load(["rowIndex"]);
    range.load(["text", "columnIndex", "rowIndex", "isEntireRow"]);
    await context.sync();

    sheetname = sheet.name;
    console.log(`sheet.name:"${sheetname}"`);
    let tablerowindex = tablerange.rowIndex;

    columnIndex = range.columnIndex;
    rowIndex = range.rowIndex;
    console.log(
      `sheet.name:"${sheetname}",rowIndex:"${JSON.stringify(rowIndex)}",columnIndex:"${JSON.stringify(columnIndex)}" `
    );
    if (rowIndex > tablerowindex && range.isEntireRow) {
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
export async function addActionOnSelect(config: FormConfig): Promise<any> {
  return await Excel.run(async function (context) {
    // document.getElementById("message-area").innerHTML = "<br>writeDataverseDataToOfficeDocument: " ;
    console.log(`addActionOnSelect`);

    try {
      var sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.onSingleClicked.add(async function (eventArgs) {
        console.log("onSelectionChanged");
        global.currentSelectRangeAddress = eventArgs.address;
        console.log(global.currentSelectRangeAddress);
        await picker.init(eventArgs, config);
        return await context.sync();
      });

      // Bind the range to monitor for changes
      sheet.onSelectionChanged.add(async function (eventArgs) {
        console.log("onSelectionChanged");
        global.currentSelectRangeAddress = eventArgs.address;
        console.log(global.currentSelectRangeAddress);
        // Load the values and format of the selected range
        // var range = sheet.getRange(global.currentSelectRangeAddress);
        await showHideDetaleButton(context, sheet, global.currentSelectRangeAddress);
        await onSelectionChanged__model(context, sheet, global.currentSelectRangeAddress);
        console.log(`onSelectionChanged`);
        // global.TableDataChangeAction.push(eventArgs);
        console.log(eventArgs);
        // await onchange_dataValidation(eventArgs, config);
        return await context.sync();
      });
      sheet.onChanged.add(async function (eventArgs) {
        console.log("onChanged");
        global.currentSelectRangeAddress = eventArgs.address;
        console.log(eventArgs);
        console.log(global.currentSelectRangeAddress);
        // onChangedCheckDataValidation(eventArgs);
        // return await context.sync();
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

export async function onChangedCheckDataValidation(eventArgs: Excel.TableChangedEventArgs) {
  return await Excel.run(async function (context) {
    try {
      console.log(`on Changed Check Data Validation...`);
      console.log(eventArgs);
      var sheetname = "";
      var value = "";
      var values = [];
      var columnIndex = -1;
      var rowIndex = -1;
      // var sheet = context.workbook.worksheets.getActiveWorksheet();

      if (eventArgs.changeType == "RangeEdited") {
        let table = context.workbook.tables.getItem(eventArgs.tableId);
        let sheet = context.workbook.worksheets.getItem(eventArgs.worksheetId);
        let range = eventArgs.getRange(context);
        range.load(["address", "text", "columnIndex", "rowIndex", "format"]);

        sheet.load(["name"]);
        table.load(["name"]);
        await context.sync();
        let config = global.WorksheetTableConfig[table.name];

        console.log(`sheet.name:"${sheet.name}"`);
        value = range.text[0][0];
        columnIndex = range.columnIndex;
        rowIndex = range.rowIndex;

        let col = table.columns.getItem(columnIndex + 1);
        col.load(["name"]);
        await context.sync();

        console.log(`sheet.column.Index:"${columnIndex}"`);
        console.log(col);
        console.log(config);
        const confCol = config.columns.find((ele: TableColumns) => ele.Label === col.name);
        console.log(confCol);

        const Coldatavalidation = confCol.datavalidation;
        if ("separate" in Coldatavalidation) {
          values = value.split(Coldatavalidation.separate);
        }
        let vaild = false;

        if (values) {
          for (let val of values) {
            if ("regexp" in Coldatavalidation) {
              console.log("Coldatavalidation.regexp...");
              console.log(Coldatavalidation.regexp);
              const re = new RegExp(Coldatavalidation.regexp);
              vaild = re.test(val);
              console.log(vaild);
            }
          }
          try {
            let confCell = null;
            if (vaild) {
              confCell = Coldatavalidation.TrueCell;
            } else {
              confCell = Coldatavalidation.FalseCell;
            }
            await lockmainObjectValueToCell(false);
            console.log(confCell.format.fill.color);
            range.format.fill.color = confCell?.format?.fill?.color ?? "white";

            await context.sync();
            await lockmainObjectValueToCell(true);

            // range.format.fill.color = confcell.format.fill.color ?? "white";
            // console.log(confcell.format.fill.color);
            // setCell(eventArgs, confcell);
            // if (confcell.format.fill.clear) {
            //   range.format.fill.clear();
            // }
          } catch (e) {
            console.log(`Error Data Validation; ${e.message}`);
            console.error(e.stack);
          }
          // await context.sync();
        }
        if (columnIndex == 0) {
          throw new Error("column Index = 0");
        }
        console.log(
          `sheet.name:"${sheetname}",rowIndex:"${JSON.stringify(rowIndex)}",columnIndex:"${JSON.stringify(
            columnIndex
          )}" `
        );
      }
    } catch (e) {
      console.log(`Error onChangedCheckDataValidation;${e.message}`);
      console.log(e.stack);
    }
    return context.sync();
  });
}

// export async function setCell(eventArgs: Excel.TableChangedEventArgs, confCell: any) {
//   return await Excel.run(async function (context) {
//     try {
//       console.log(`on setCell...`);
//       console.log(eventArgs);
//       if (eventArgs.changeType == "RangeEdited") {
//         // let table = context.workbook.tables.getItem(eventArgs.tableId);
//         let sheet = context.workbook.worksheets.getItem(eventArgs.worksheetId);
//         let range = sheet.getRange(eventArgs.address);
//         try {
//           lockmainObjectValueToCell(false).then(() => {
//             console.log(confCell.format.fill.color);
//             range.format.fill.color = confCell?.format?.fill?.color ?? "white";

//             context.sync().then(() => {
//               lockmainObjectValueToCell(true);
//             });
//           });
//         } catch (e) {
//           console.log(`Error setCell in Data Validation; ${e.message}`);
//           console.error(e.stack);
//         }
//       }
//     } catch (e) {
//       console.log(`Error setCell;${e.message}`);
//       console.log(e.stack);
//     }
//     // return context.sync();
//   });
// }

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

    values.push(range.text);
    columnIndex = range.columnIndex;
    rowIndex = range.rowIndex;
    if (columnIndex == 0) {
      throw new Error("column Index = 0");
    }
    console.log(
      `sheet.name:"${sheetname}",rowIndex:"${JSON.stringify(rowIndex)}",columnIndex:"${JSON.stringify(columnIndex)}" `
    );
    //mainform_MBHK
    //SM MBHK
    // if(sheet.name === "SM MBHK" && columnIndex === 2 && rowIndex > 17){
    //   console.log(`is "SM ZF (Hong Kong)" selecting "model"`);
    //   setApplyModel(true);
    // }
    // if(sheet.name === "SM ZF (Hong Kong)" && columnIndex === 2 && rowIndex > 12){
    //   console.log(`is "SM ZF (Hong Kong)" selecting "model"`);
    //   setApplyModel(true);
    // }
    setApplyModel(true);
  } catch (e) {
    console.log(`Error onSelectionChanged__model;${e.message}`);
  }
  return context.sync();
}
export var setApplyModel = function (condition) {
  console.log("#applyModel " + condition);
  console.log(jq("#applyModel"));
  // if (condition) {
  //   jq(`#${"applyModelButtonId"}`)[0].prop("disabled", "disabled");
  // } else {
  //   jq(`#${"applyModelButtonId"}`)[0].removeAttr("disabled");
  // }
};

export function create_pageElement(type: string, addtoId: string, module_name, arge: object, callback?: any) {
  const ele = document.createElement(type);
  for (let key of Object.keys(arge)) {
    ele[key] = arge[key];
  }
  if (module_name) {
    ele.setAttribute("data-module", module_name);
  }
  if (callback) {
    ele.addEventListener(
      "click",
      function () {
        const m = this.getAttribute("data-module");
        console.log("onClick....");
        console.log(m);
        if (m) {
          callback(m);
        } else {
          callback();
        }
      },
      false
    );
  }
  document.getElementById(addtoId).appendChild(ele);
  // if (callback) {
  //   ele.onclick = (e) => {
  //     console.log(e.target); // Get ID of Clicked Element
  //     console.log(e.target.getAttribute("data-module", )); // Get ID of Clicked Element
  //     console.log(e.target["id"]); // Get ID of Clicked Element
  //     // callback(Module);
  //   };
  // }
}

export function add_hardcode(obj: object, arr: object[]) {
  try {
    for (let ele of arr) {
      try {
        obj[ele["LogicalName"]] = ele["value"];
      } catch (e) {
        console.error(e.stack);
      }
    }
    return obj;
  } catch (e) {
    console.log(e.stack);
  }
}

export function findIndex(stringArr, keyString) {
  let result = [-1, -1];
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
  return result;
}

export function getDataTabledata(TableNames: string[]): string[] {
  let ndata = [];
  for (let key of TableNames) {
    if (global.pagetable[key]) {
      let data = global.pagetable[key].rows({ selected: true }).data();
      // var newarray = [];
      for (var i = 0; i < data.length; i++) {
        console.log(data[i]);
        ndata.push(data[i][0]);
        console.log("Name: " + data[i][0] + " Address: " + data[i][1] + " Office: " + data[i][2]);
      }
      //   if(applyValidation()){
      // }else{
      //   console.log(`Error "SM ZF (Hong Kong)" no selecting "model"`);

      // }
    } else {
      console.log("Error applyNewModel: table null");
    }
  }
  return ndata;
}

export function clearDataTableSelectedData(TableNames: string[]) {
  for (let key of TableNames) {
    if (global.pagetable[key]) {
      global.pagetable[key].rows({ selected: true }).deselect();
      // global.pagetable[key].rows(".selected").nodes().to$().removeClass("selected");
    } else {
      console.log("Error applyNewModel: table null");
    }
  }
}

export async function applyNewModel() {
  return await Excel.run(async function (context) {
    console.log("applyNewModel...");
    const init_conf: initConfig = require("./config/init.json");
    let ndata = [];
    try {
      // let modelNames = [];
      ndata = getDataTabledata(Object.keys(init_conf.applyString));

      var sData = ndata.join(",");
      console.log(`rows_selected: ${sData}`);

      var sheetname = "";
      var sheet = context.workbook.worksheets.getActiveWorksheet();

      sheet.load(["name"]);
      await context.sync();
      sheetname = sheet.name;
      console.log(`sheet.name:"${sheetname}"`);

      var range = sheet.getRange(global.currentSelectRangeAddress);
      range.load(["text", "columnIndex", "rowIndex"]);
      await context.sync();

      var columnIndex = range.columnIndex;
      var rowIndex = range.rowIndex;
      console.log(
        `sheet.name:"${sheetname}",rowIndex:"${JSON.stringify(rowIndex)}",columnIndex:"${JSON.stringify(columnIndex)}" `
      );
      // sheet.name === "SM ZF (Hong Kong)" && columnIndex === 2 && rowIndex > 12

      range.values = [[sData]];
      // range.format.autofitColumns();
      await context.sync().then(() => {
        clearDataTableSelectedData(Object.keys(init_conf.applyString));
      });
    } catch (e) {
      console.log("Error applyNewModel: " + e.message);
    }
    return context.sync();
  });
}
export async function post_Data_ReturnData(entity: string, items: any) {
  console.log("post_Data");
  console.log("items", items);
  global.Callapiaction = {
    name: "callapiaction",
    action: {
      entitySet: entity,
      queryString: items,
      queryOptions: "",
    },
  };
  return await datahelper.post_Data_ReturnData();
}

export async function getLastRowRecord() {
  let currentDatas = await get_table_data_arrary(global.CurrentConfig);

  let lastRowRecord = [];

  if (currentDatas) {
    for (let i = currentDatas.length - 1; i < currentDatas.length; i++) {
      lastRowRecord.push(currentDatas[i]);
    }
  }

  return lastRowRecord[0];
}
