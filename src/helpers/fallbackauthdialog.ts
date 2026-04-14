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
import * as Benz from "./benz/type";

var config = require(`./mode/${Benz.AppMode.PRODUCTION}/app.json`);

/* global localStorage */
export var token: string = null;
const clientId = config.clientId;
const authority = config.authority;
// const accessScope = `api://${window.location.host}/${clientId}/user_impersonation`;
var environment_name: string = config.environment_name;

const msalConfig: Configuration = {
  auth: {
    clientId: clientId,
    authority: authority,
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

const publicClientApp: PublicClientApplication = new PublicClientApplication(msalConfig);

let loginDialog: Office.Dialog = null;
let callbackFunction = null;
let action = null;
let actionArgs = null;

Office.onReady(() => {
  if (Office.context.ui.messageParent) {
    publicClientApp
      .handleRedirectPromise()
      .then(handleResponse)
      .catch((error) => {
        console.error(error);
        Office.context.ui.messageParent(JSON.stringify({ status: "failure", result: error }));
      });

    const loginRequest: RedirectRequest = {
      scopes: [`https://${environment_name}.api.crm5.dynamics.com/user_impersonation`],
      // extraScopesToConsent: ["user.read"],
    };

    // The very first time the add-in runs on a developer's computer, msal.js hasn't yet
    // stored login data in localStorage. So a direct call of acquireTokenRedirect
    // causes the error "User login is required". Once the user is logged in successfully
    // the first time, msal data in localStorage will prevent this error from ever hap-
    // pening again; but the error must be blocked here, so that the user can login
    // successfully the first time. To do that, call loginRedirect first instead of
    // acquireTokenRedirect.
    if (localStorage.getItem("loggedIn") === "yes") {
      publicClientApp.acquireTokenRedirect(loginRequest);
    } else {
      // This will login the user and then the (response.tokenType === "id_token")
      // path in authCallback below will run, which sets localStorage.loggedIn to "yes"
      // and then the dialog is redirected back to this script, so the
      // acquireTokenRedirect above runs.
      publicClientApp.loginRedirect(loginRequest);
    }
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

export async function dialogFallback(callback, callGetData?: any, args?) {
  // const environment_name: string = "org8390b622";

  const loginRequest: RedirectRequest = {
    scopes: [`https://${environment_name}.api.crm5.dynamics.com/user_impersonation`],
    // extraScopesToConsent: ["user.read"],
  };

  action = callGetData;
  actionArgs = args;

  // Attempt to acquire token silently if user is already signed in.
  if (global.AccountID !== null) {
    const result = await publicClientApp.acquireTokenSilent(loginRequest);
    if (result !== null && result.accessToken !== null) {
      let response = null;
      // const response = await callGetUserData(result.accessToken);
      token = result.accessToken;
      global.ApiAccessToken = result.accessToken;
      if (typeof callGetData === "function") {
        response = await callGetData(result.accessToken);
      } else {
        response = await callGetUserData(result.accessToken);
      }
      callbackFunction(response);
    }
  } else {
    callbackFunction = callback;
    // We fall back to Dialog API for any error.
    const url = "/fallbackauthdialog.html";
    showLoginPopup(url);
  }
}

// This handler responds to the success or failure message that the pop-up dialog receives from the identity provider
// and access token provider.
async function processMessage(arg) {
  let messageFromDialog = JSON.parse(arg.message);

  if (messageFromDialog.status === "success") {
    console.info("loginDialog success...");
    // We now have a valid access token.
    loginDialog.close();

    // Configure MSAL to use the signed-in account as the active account for future requests.
    const homeAccount = publicClientApp.getAccountByHomeId(messageFromDialog.accountId);
    if (homeAccount) {
      global.AccountID = messageFromDialog.accountId;
      publicClientApp.setActiveAccount(homeAccount);
    }
    let response: any;
    global.ApiAccessToken = messageFromDialog.result;
    await setAccount(messageFromDialog.result);
    if (!action) {
      response = await callGetUserData(messageFromDialog.result, actionArgs);
    } else {
      response = await action(messageFromDialog.result);
    }
    showMessage({ style: "success", message: "Login success" });
    callbackFunction(response);
  } else if (messageFromDialog.error === undefined && messageFromDialog.result.errorCode === undefined) {
    console.error("no loginDialog error...");
    showMessage({ style: "error", message: "Login error" });

    // Need to pick the user to use to auth
  } else {
    console.error("loginDialog error...");
    showMessage({ style: "error", message: "Login error" });
    // Something went wrong with authentication or the authorization of the web application.
    loginDialog.close();
  }
}

// Use the Office dialog API to open a pop-up and display the sign-in page for the identity provider.
function showLoginPopup(url) {
  console.log(`showLoginPopup...`);

  var fullUrl = location.protocol + "//" + location.hostname + (location.port ? ":" + location.port : "") + url;
  console.log(`showLoginPopup url:${fullUrl}`);

  // height and width are percentages of the size of the parent Office application, e.g., PowerPoint, Excel, Word, etc.
  Office.context.ui.displayDialogAsync(fullUrl, { height: 60, width: 30 }, function (result) {
    console.log("Dialog has initialized. Wiring up events");
    loginDialog = result.value;
    loginDialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
  });
}
