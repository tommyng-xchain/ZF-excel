/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file shows how to use MSAL.js to get an access token to your server and pass it to the task pane.
 */
/* global global, require, console, location, Office */

import { Configuration, LogLevel, PublicClientApplication, RedirectRequest } from "@azure/msal-browser";
import { callGetUserData } from "./middle-tier-calls";
import { setAccount } from "./dataverse-data-helper";
import { showMessage } from "./message-helper";
import { loadingspinner } from "../helpers/loadingspinner-helper";
import * as nav from "../helpers/benz/components/main-nav";
import * as Benz from "./benz/type";

/* global localStorage */

let Dialog: Office.Dialog = null;
let DialogMessageReceivedEvent = null;
let callbackFunction = null;
let action = null;
let actionArgs = null;

Office.onReady(() => {
  if (Office.context.ui.messageParent) {
    // publicClientApp
    //   .handleRedirectPromise()
    //   .then(handleResponse)
    //   .catch((error) => {
    //     console.error(error);
    //     Office.context.ui.messageParent(JSON.stringify({ status: "failure", result: error }));
    //   });

  }
});

function handleResponse(response) {
  if (response.tokenType === "id_token") {
    localStorage.setItem("loggedIn", "yes");
  } else {
    global.ApiAccessToken = response.accessToken;
    global.Account = response.account;
    global.AccountID = response.account.homeAccountId;
    Office.context.ui.messageParent(
      JSON.stringify({ status: "success", result: response.accessToken, accountId: response.account.homeAccountId })
    );
  }
}


// This handler responds to the success or failure message that the pop-up dialog receives from the identity provider
// and access token provider.
async function processMessage(arg) {
  console.info("modeDialog processMessage ...");

  let msg = JSON.parse(arg.message);
  console.info("modeDialog processMessage ...");
  console.info(msg);
  global.mode = msg.mode;
  var config = require(`./mode/${global.mode == Benz.AppMode.UAT ? Benz.AppMode.UAT.toLowerCase() : Benz.AppMode.PRODUCTION.toLowerCase()}/app.json`);

  global.environment_name = config.environment_name;
  global.clientId =config.clientId;
  global.authority =config.authority;
  // nav.show();
  // loadingspinner(false);
  await Dialog.close()
  
  setTimeout(() => { 

  DialogMessageReceivedEvent();
}, 2000);

}

// Use the Office dialog API to open a pop-up and display the sign-in page for the identity provider.
export function showPopup(DialogMessageReceivedEventHandler) {
  console.log(`showPopup...`);
  DialogMessageReceivedEvent = DialogMessageReceivedEventHandler;
  var fullUrl = location.protocol + "//" + location.hostname + (location.port ? ":" + location.port : "") + "/modeDialog.html";
  console.log(`showPopup url:${fullUrl}`);

  // height and width are percentages of the size of the parent Office application, e.g., PowerPoint, Excel, Word, etc.
  Office.context.ui.displayDialogAsync(fullUrl, { height: 20, width: 10 }, function (result) {
    console.log("Dialog has initialized. Wiring up events");
    Dialog = result.value;
    Dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
  });
}

export function retrieve_Data(value) {
  Office.context.ui.messageParent(value);
}
