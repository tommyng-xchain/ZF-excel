/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, Office, localStorage */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});
export function setGlobalVariable(key, value) {
  try {
    localStorage.setItem(key, value);
  } catch (e) {
    console.error(`Get global variable Error: {key:${key}, value:${value}}`);
    throw e;
  }
}
export function getGlobalVariable(key) {
  try {
    return localStorage.getItem(key);
  } catch (e) {
    console.error(`Get global variable Error: {key:${key}}`);
    throw e;
  }
}
/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
export function action(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message.
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

// Register the function with Office.
Office.actions.associate("action", action);
