// Copyright (c) Djordje Radulovic. Licensed under the MIT license. See LICENSE in the project root.
// Builds the actionable message (Adaptive Card) that invokes the add-in.
"use strict";

(function () {
  function toFacts(context) {
    return Object.keys(context).map(function (key) {
      const value = context[key];
      return {
        title: key,
        value: typeof value === "object" ? JSON.stringify(value) : String(value)
      };
    });
  }

  function escapeForScriptTag(json) {
    // Keep "</script>" inside a value from closing the embedding tag early.
    return json.replace(/</g, "\\u003c");
  }

  function buildCard(context) {
    const config = window.ADDIN_CONFIG;
    return {
      type: "AdaptiveCard",
      version: "1.0",
      originator: config.originator,
      hideOriginalBody: true,
      body: [
        {
          type: "TextBlock",
          text: "Activate the Actionable Message Activation add-in",
          size: "large",
          weight: "bolder",
          wrap: true
        },
        {
          type: "TextBlock",
          text: "The first button opens the add-in task pane with this initialization context:",
          wrap: true
        },
        { type: "FactSet", facts: toFacts(context) },
        {
          type: "TextBlock",
          text: "The second button opens Message Header Analyzer and shows on-demand install of a store add-in.",
          wrap: true,
          spacing: "large"
        }
      ],
      actions: [
        {
          type: "Action.InvokeAddInCommand",
          title: "Invoke \"View Initialization Context\"",
          addInId: config.addInId,
          desktopCommandId: "showInitContext",
          initializationContext: context
        },
        {
          type: "Action.InvokeAddInCommand",
          title: "Invoke Message Header Analyzer",
          addInId: config.storeAddInId,
          desktopCommandId: "mhaOpenPaneButton"
        }
      ]
    };
  }

  function buildMessage(context, recipient) {
    const card = escapeForScriptTag(JSON.stringify(buildCard(context)));
    const html =
      "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">" +
      "<script type=\"application/adaptivecard+json\">" + card + "</script></head>" +
      "<body><p>If you don't see a card with buttons above, your mail client doesn't support " +
      "actionable messages. Open this message in Outlook on the web or new Outlook for Windows.</p></body></html>";

    return {
      subject: "ACTIVATE YOUR ADD-IN",
      body: { contentType: "HTML", content: html },
      toRecipients: [{ emailAddress: { address: recipient } }]
    };
  }

  window.AddinCard = { buildMessage: buildMessage };
})();
