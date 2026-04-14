/* global global, require, console, location, Office */
import { LogLevel, PublicClientApplication } from "@azure/msal-browser";
import * as Choices from "./components/choices";
import * as utils from "./utils";
// import * as mode from "../modeDialog";
var config = require("./config/init.json");
var jq = require("jquery");
global.dataChoices = {};

global.WorksheetTableConfig = {};

global.modules = {};

global.pagetable = [];

global.modulesOChange = [];

global.readyCallapiaction = [];

global.tableConfigs = {};

global.choicesSets = {};

global.AccountID = null;

global.RowMeasureID = [];

global.RowMeasureName = [];

global.RowSpecialCase = [];

global.EnvironmentVariable = {};

// const accessScope = `api://${window.location.host}/${clientId}/user_impersonation`;
global.environment_name = null; //"org8390b622";

global.clientId = null; // "589a390c-39e0-4726-ad2f-c8a3bfc0e676"; //This is your client ID

global.msalConfig = {
  auth: {
    clientId: global.clientId,
    authority: "https://login.microsoftonline.com/common",
    // authority: "https://login.microsoftonline.com/c807653f-1965-49f9-a1fb-851f41a414e7",

    redirectUri:
      location.protocol +
      "//" +
      location.hostname +
      (location.port ? ":" + location.port : "") +
      "/fallbackauthdialog.html", // Update config script to enable `https://${window.location.host}/fallbackauthdialog.html`,
    navigateToLoginRequestUrl: false,
  },
  cache: {
    cacheLocation: "memoryStorage", // Needed to avoid "User login is required" error.sessionStorage  localStorage	memoryStorage
    storeAuthStateInCookie: true, // Recommended to avoid certain IE/Edge issues.
  },
  system: {
    loggerOptions: {
      loggerCallback: (level, message, containsPii) => {
        if (containsPii) {
          return;
        }
        switch (level) {
          case LogLevel.Error:
            console.error(message);
            return;
          case LogLevel.Info:
            console.info(message);
            return;
          case LogLevel.Verbose:
            console.debug(message);
            return;
          case LogLevel.Warning:
            console.warn(message);
            return;
        }
      },
    },
  },
};
global.PublicClientApp = new PublicClientApplication(global.msalConfig);

let month: any[] = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"];
month = month.map((val) => val.toString());
global.Month = month.map((val) => ({ id: val.toString(), name: val.toString(), value: val.toString() }));
Choices.set_choicesSets("month", global.Month);
Choices.set_dataChoices("month", month);
console.log(global.choicesSets);
Office.onReady((info) => {
  console.log("onload Office Ready - benz/init");
  if (info.host === Office.HostType.Excel) {
    jq("#addrow").on("click", utils.addTableRow);
    jq("#add10row").click({ num: 10 }, utils.addTableRows);
    jq("#clearRow").on("click", utils.clearSelectTableRow);
    jq("#deleterow").on("click", utils.removeSelectTableRow);
    jq(".applyModel").on("click", utils.applyNewModel);

    // jq("#action_back").on("click", global.action_back);
  }
});
