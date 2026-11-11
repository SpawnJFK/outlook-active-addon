// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE in the project root.
// Modernized 2026 by Djordje Radulovic: no jQuery/URI.js/moment; the dialog only edits the context
// and hands it to the function file, which sends the mail through Microsoft Graph.
"use strict";

Office.onReady(function () {
  const contextBox = document.getElementById("init-context");
  contextBox.value = JSON.stringify(sampleContext(), null, 2);
  document.getElementById("send-msg").addEventListener("click", submit);
});

function sampleContext() {
  return {
    dateStamp: new Date().toISOString(),
    myStringProperty: "Hello world",
    myBooleanProperty: true
  };
}

function submit() {
  hideError();
  const raw = document.getElementById("init-context").value;
  let context;
  try {
    context = JSON.parse(raw);
  } catch {
    showError("Could not parse the initialization context. Make sure it is valid JSON and try again.");
    return;
  }
  if (context === null || typeof context !== "object" || Array.isArray(context)) {
    showError("The initialization context must be a JSON object, for example { \"key\": \"value\" }.");
    return;
  }

  const button = document.getElementById("send-msg");
  button.disabled = true;
  button.textContent = "Sending...";
  Office.context.ui.messageParent(JSON.stringify(context));
}

function showError(message) {
  document.getElementById("error-msg").textContent = message;
  document.getElementById("error").hidden = false;
}

function hideError() {
  document.getElementById("error").hidden = true;
  document.getElementById("error-msg").textContent = "";
}
