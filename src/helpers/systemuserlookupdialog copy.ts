/* global global, require, console, localStorage, location, sessionStorage, Office, document, window */
import { Configuration, LogLevel, PublicClientApplication, RedirectRequest } from "@azure/msal-browser";
import { callGetUserData } from "./middle-tier-calls";
// import "bootstrap-select";
// Import our custom CSS
// import "bootstrap/scss/bootstrap.scss";

// Import all of Bootstrap's JS
import "bootstrap";
import { getGlobalVariable, setGlobalVariable } from "../commands/commands";
var jq = require("jquery");

const m = "systemuserlookupdialog";
export const url = "/systemuserlookupdialog.html";

export var token: string = null;
const clientId = "589a390c-39e0-4726-ad2f-c8a3bfc0e676"; //This is your client ID
// const accessScope = `api://${window.location.host}/${clientId}/user_impersonation`;

const msalConfig: Configuration = {
  auth: {
    clientId: clientId,
    authority: "https://login.microsoftonline.com/common",
    // authority: "https://login.microsoftonline.com/c807653f-1965-49f9-a1fb-851f41a414e7",

    redirectUri: location.protocol + "//" + location.hostname + (location.port ? ":" + location.port : "") +"/fallbackauthdialog.html", // Update config script to enable `https://${window.location.host}/fallbackauthdialog.html`,
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

let Dialog: Office.Dialog = null;
let callbackFunction = null;
let action = null;

export function onRegisterMessageComplete(asyncResult) {
  if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
    console.error("Emailpopup error: ", asyncResult.error.message);
  } else {
    console.log(asyncResult);
    console.log("Emailpopup registering complete ");
  }
}

export function onMessageFromParent(arg: Office.DialogParentMessageReceivedEventArgs) {
  try {
    console.log("Emailcontent: ");
    console.log(arg);
  } catch (e) {
    console.error("Error selectpicker...");
    console.error(e.stack);
  }
  // document.getElementById('email').innerHTML = arg.message
}
Office.onReady((info) => {
  try {
    console.log(Office.context.ui);
    console.log(info);
    Office.context.ui.addHandlerAsync(
      Office.EventType.DialogParentMessageReceived,
      onMessageFromParent,
      onRegisterMessageComplete
    );
    var apiAccessToken = getGlobalVariable("accessToken");
    // // var selectpicker = jq("#selectpicker").selectpicker();
    // console.log(selectpicker);
    // let input = jq("#selectpicker").parent().find("input");
    // jq(document).ready(function () {
    //   // The document is ready.
    //     console.log("document ready...");
    //     console.log(input);
    jq("input#user").on("change", function () {
      console.log(`input change... ${jq("search")}`);

      getData(jq("search").val);
    });
  } catch (e) {
    console.error("Error selectpicker...");
    console.error(e.stack);
  }
  // document.getElementById("name").value = global.ApiAccessToken;
  // document.getElementById("submit").onclick = () => tryCatch(sendStringToParentPage);
});
async function getData(val: string) {
  return ["test1", "test2", "1234test3", "1234test4", "1234test5"].filter((v) => v.startsWith(val));
}
function sendStringToParentPage() {
  // const userName = document.getElementById("name").value;
  Office.context.ui.messageParent("abc");
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}
// Office.onReady(() => {
//   if (Office.context.ui.messageParent) {
//     publicClientApp
//       .handleRedirectPromise()
//       .then(handleResponse)
//       .catch((error) => {
//         console.log(error);
//         Office.context.ui.messageParent(JSON.stringify({ status: "failure", result: error }));
//       });

//     const loginRequest: RedirectRequest = {
//       scopes: [`https://${global.environment_name}.api.crm5.dynamics.com/user_impersonation`],
//       // extraScopesToConsent: ["user.read"],
//     };

//     if (localStorage.getItem("loggedIn") === "yes") {
//       publicClientApp.acquireTokenRedirect(loginRequest);
//     } else {
//       // This will login the user and then the (response.tokenType === "id_token")
//       // path in authCallback below will run, which sets localStorage.loggedIn to "yes"
//       // and then the dialog is redirected back to this script, so the
//       // acquireTokenRedirect above runs.
//       publicClientApp.loginRedirect(loginRequest);
//     }
//   }
// });

function handleResponse(response) {
  // console.log(`handleResponse... ;${JSON.stringify(response)}`);
  if (response.tokenType === "id_token") {
    console.log("LoggedIn");
    localStorage.setItem("loggedIn", "yes");
  } else {
    console.log("token type is:" + response.tokenType);
    global.ApiAccessToken = response.accessToken;
    console.log("token:" + global.ApiAccessToken);
    global.Account = response.account;
    global.AccountID = response.account.homeAccountId;
    console.log("response.account");
    console.log(response.account);

    Office.context.ui.messageParent(
      JSON.stringify({ status: "success", result: response.accessToken, accountId: response.account.homeAccountId })
    );
  }
}

export async function dialogFallback(callback, callGetData?: any) {

  const loginRequest: RedirectRequest = {
    scopes: [`https://${global.environment_name}.api.crm5.dynamics.com/user_impersonation`],
    // extraScopesToConsent: ["user.read"],
  };
  console.log("dialogFallback...");

  action = callGetData;

  // Attempt to acquire token silently if user is already signed in.
  if (global.AccountID !== null) {
    const result = await publicClientApp.acquireTokenSilent(loginRequest);
    if (result !== null && result.accessToken !== null) {
      let response = null;
      // const response = await callGetUserData(result.accessToken);
      token = result.accessToken;
      global.ApiAccessToken = result.accessToken;
      console.log("token: " + global.ApiAccessToken);
      if (typeof callGetData === "function") {
        console.log("callGetData...");
        console.log("result...");
        console.log(result);
        response = await callGetData(result.accessToken);
      } else {
        console.log("callGetUserData...");
        response = await callGetUserData(result.accessToken);
      }
      callbackFunction(response);
    }
  } else {
    console.log("dialogFallback...global.homeAccountId = null");

    callbackFunction = callback;

    // We fall back to Dialog API for any error.
    showLoginPopup();
  }
}

// This handler responds to the success or failure message that the pop-up dialog receives from the identity provider
// and access token provider.
async function processMessage(arg) {
  console.log("processMessage...");
  // Uncomment to view message content in debugger, but don't deploy this way since it will expose the token.
  console.log("Message received in processMessage: " + JSON.stringify(arg));

  let messageFromDialog = JSON.parse(arg.message);
  // // console.log(messageFromDialog);

  // if (messageFromDialog.status === "success") {
  //   console.log("Dialog success...");
  //   // We now have a valid access token.
  //   Dialog.close();

  //   // Configure MSAL to use the signed-in account as the active account for future requests.
  //   const homeAccount = publicClientApp.getAccountByHomeId(messageFromDialog.accountId);
  //   if (homeAccount) {
  //     global.AccountID = messageFromDialog.accountId;
  //     publicClientApp.setActiveAccount(homeAccount);
  //   }
  //   let response: any;
  //   // const response = await callGetUserData(messageFromDialog.result);
  //   // console.log("messageFromDialog.result:"+messageFromDialog.result);
  //   // console.log("action:" + action);
  //   global.ApiAccessToken = messageFromDialog.result;
  //   if (!action) {
  //     response = await callGetUserData(messageFromDialog.result);
  //   } else {
  //     response = await action(messageFromDialog.result);
  //   }
  //   console.log("callbackFunction...");

  //   // if(response){

  //   //   console.log("has response:"+response);
  //   // }

  //   callbackFunction(response);
  // } else if (messageFromDialog.error === undefined && messageFromDialog.result.errorCode === undefined) {
  //   console.log("no Dialog error...");

  //   // Need to pick the user to use to auth
  // } else {
  //   console.log("Dialog error...");
  //   // Something went wrong with authentication or the authorization of the web application.
  //   Dialog.close();
  // }
  Dialog.close();
}

function setMessageChild() {
  const messageToDialog = JSON.stringify({
    name: "My Sheet",
    position: 2,
  });

  Dialog.messageChild(messageToDialog, { targetOrigin: "*" });
}
// Use the Office dialog API to open a pop-up and display the sign-in page for the identity provider.
export function showLoginPopup() {
  console.log(`${m} showLoginPopup...`);
  // localStorage.setItem("ApiAccessToken", global.ApiAccessToken);
  sessionStorage.setItem("key", "value");

  setGlobalVariable("apiaccesstoken", global.ApiAccessToken);
  setGlobalVariable("getdata", JSON.stringify({ startwith: 3 }));
  var fullUrl = location.protocol + "//" + location.hostname + (location.port ? ":" + location.port : "") + url;
  console.log(`showLoginPopup url:${fullUrl}`);
  // height and width are percentages of the size of the parent Office application, e.g., PowerPoint, Excel, Word, etc.
  Office.context.ui.displayDialogAsync(
    `${window.location.origin}/${url}`,
    { height: 60, width: 30 },
    function (result) {
      console.log("Dialog has initialized. Wiring up events");
      console.log(result.value);
      Dialog = result.value;
      Dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
      Dialog.messageChild(JSON.stringify({ token: "sdfdsfad" }));
    }
  );
}
