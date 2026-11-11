// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE in the project root.
// Modernized 2026 by Djordje Radulovic: Office.onReady, Graph sendMail via NAA, Office.actions.associate.
"use strict";

let sendDialog = null;
let pendingEvent = null;

Office.onReady(function () {
  Office.actions.associate("sendActionableMessage", sendActionableMessage);
});

function notify(key, type, message) {
  const details = { type: type, message: message.substring(0, 150) };
  if (type === Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage) {
    details.icon = "icon-16";
    details.persistent = false;
  }
  Office.context.mailbox.item.notificationMessages.replaceAsync(key, details);
}

function finish() {
  if (pendingEvent) {
    pendingEvent.completed();
    pendingEvent = null;
  }
}

function sendActionableMessage(event) {
  pendingEvent = event;
  const url = new URL("../dialog/dialog.html", window.location.href).toString();

  Office.context.ui.displayDialogAsync(url, { width: 50, height: 70, displayInIframe: true }, function (result) {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      notify("send-msg-error", Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        "Could not open the dialog: " + result.error.message);
      finish();
      return;
    }
    sendDialog = result.value;
    sendDialog.addEventHandler(Office.EventType.DialogMessageReceived, onDialogMessage);
    sendDialog.addEventHandler(Office.EventType.DialogEventReceived, onDialogClosed);
  });
}

async function onDialogMessage(arg) {
  sendDialog.close();
  sendDialog = null;

  let context;
  try {
    context = JSON.parse(arg.message);
  } catch {
    notify("send-msg-error", Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      "The dialog returned invalid JSON.");
    finish();
    return;
  }

  try {
    const recipient = Office.context.mailbox.userProfile.emailAddress;
    await AddinAuth.sendMail(AddinCard.buildMessage(context, recipient));
    notify("send-msg-success", Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      "Actionable message sent. Check your inbox.");
  } catch (err) {
    notify("send-msg-error", Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      "Could not send the message: " + err.message);
  }
  finish();
}

function onDialogClosed() {
  sendDialog = null;
  finish();
}
