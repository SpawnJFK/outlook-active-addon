// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE in the project root.
// Modernized 2026 by Djordje Radulovic: Office.onReady, no jQuery, InitializationContextChanged
// is registered once and actually renders the new context.
"use strict";

// Error code Outlook returns when there is no initialization context.
const NO_CONTEXT_ERROR = 9020;

Office.onReady(function () {
  // Pinned task pane: the user selects another message.
  Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadMessage);
  loadMessage();
});

function loadMessage() {
  resetPage();
  const item = Office.context.mailbox.item;
  if (!item) {
    // ItemChanged fires with no item when the selection is cleared.
    return;
  }
  document.getElementById("subject").textContent = item.subject || "(no subject)";

  item.getInitializationContextAsync(function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      renderContext(result.value);
    } else if (result.error && result.error.code === NO_CONTEXT_ERROR) {
      show("no-context");
    } else {
      showError(JSON.stringify(result.error, null, 2));
    }
  });

  // The pane is already open when the user clicks the card button again.
  item.addHandlerAsync(Office.EventType.InitializationContextChanged, function (args) {
    resetPage();
    document.getElementById("subject").textContent = item.subject || "(no subject)";
    renderContext(args.initializationContext);
  });
}

function renderContext(raw) {
  if (!raw) {
    show("no-context");
    return;
  }
  try {
    const context = JSON.parse(raw);
    document.getElementById("init-context").textContent = JSON.stringify(context, null, 2);
    show("has-context");
  } catch {
    showError("The initialization context is not valid JSON: " + raw);
  }
}

function resetPage() {
  ["error", "no-context", "has-context"].forEach(function (id) {
    document.getElementById(id).hidden = true;
  });
  document.getElementById("error-msg").textContent = "";
  document.getElementById("init-context").textContent = "";
}

function show(id) {
  document.getElementById(id).hidden = false;
}

function showError(message) {
  document.getElementById("error-msg").textContent = message;
  show("error");
}
