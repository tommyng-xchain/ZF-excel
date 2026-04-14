/* global global, console, localStorage, location, Office, window, require, document */

import "bootstrap";
var jq = require("jquery");

export var m = "alertdialog";
export var url = "alertdialog.html";

let dialog: Office.Dialog = null;
var messageBanner;

Office.initialize = function (reason) {
  jq(document).ready(function () {
    // Initialize the FabricUI notification mechanism and hide it
    // var element = document.querySelector(".ms-MessageBanner");
    // messageBanner = new app.notification.MessageBanner(element);
    // messageBanner.hideBanner();
  });
};

function errorHandler(error) {
  showNotification(error);
}

// Display notifications in message banner at the top of the task pane.
function showNotification(content) {
  jq("#msg").text(content);
  // messageBanner.showBanner();
  // messageBanner.toggleExpansion();
}

function dialogCallback(asyncResult) {
  if (asyncResult.status == "failed") {
    // In addition to general system errors, there are 3 specific errors for
    // displayDialogAsync that you can handle individually.
    switch (asyncResult.error.code) {
      case 12004:
        showNotification("Domain is not trusted");
        break;
      case 12005:
        showNotification("HTTPS is required");
        break;
      case 12007:
        showNotification("A dialog is already opened.");
        break;
      default:
        showNotification(asyncResult.error.message);
        break;
    }
  } else {
    dialog = asyncResult.value;
    /*Messages are sent by developers programatically from the dialog using office.context.ui.messageParent(...)*/
    dialog.addEventHandler(Office.EventType.DialogMessageReceived, messageHandler);

    /*Events are sent by the platform in response to user actions or errors. For example, the dialog is closed via the 'x' button*/
    dialog.addEventHandler(Office.EventType.DialogEventReceived, eventHandler);
  }
}

function messageHandler(arg) {
  dialog.close();
  showNotification("HTTPS is required.");
  showNotification(arg.message);
}

function eventHandler(arg) {
  // In addition to general system errors, there are 2 specific errors
  // and one event that you can handle individually.
  showNotification("Cannot load URL, no such page or bad URL syntax.");
  switch (arg.error) {
    case 12002:
      showNotification("Cannot load URL, no such page or bad URL syntax.");
      break;
    case 12003:
      showNotification("HTTPS is required.");
      break;
    case 12006:
      // The dialog was closed, typically because the user the pressed X button.
      showNotification("Dialog closed by user");
      break;
    default:
      showNotification("Undefined error in dialog window");
      break;
  }
}

export function showdialog(title, msg) {
  console.log(`show ${m} dialog...`);
  console.log(`showdialog url:${window.location.origin}/${url}?title=${title}&msg=${msg}`);
  // Office.context.ui.messageParent(JSON.stringify({ data: "This is some data from the dialog" }));
  document.cookie = JSON.stringify({ title: title, msg: msg });
  // height and width are percentages of the size of the parent Office application, e.g., PowerPoint, Excel, Word, etc.
  Office.context.ui.displayDialogAsync(`${window.location.origin}/${url}`, { height: 20, width: 30 }, dialogCallback);
}
