// Copyright (c) Djordje Radulovic. Licensed under the MIT license. See LICENSE in the project root.
// Microsoft Graph access through Nested App Authentication (NAA).
// Replaces the retired Outlook REST v2.0 endpoint and getCallbackTokenAsync({ isRest: true }).
"use strict";

(function () {
  const GRAPH_SCOPES = ["Mail.Send"];
  let msalInstance = null;

  async function getMsal() {
    if (msalInstance) {
      return msalInstance;
    }
    if (!Office.context.requirements.isSetSupported("NestedAppAuth", "1.1")) {
      throw new Error(
        "This version of Outlook does not support nested app authentication. " +
        "Use new Outlook for Windows, Outlook on the web or a current Outlook for Mac."
      );
    }
    const clientId = window.ADDIN_CONFIG.clientId;
    if (!clientId || clientId.startsWith("YOUR-")) {
      throw new Error("Set clientId in config.js to your Entra app registration ID.");
    }
    msalInstance = await msal.createNestablePublicClientApplication({
      auth: {
        clientId: clientId,
        authority: "https://login.microsoftonline.com/common"
      },
      cache: { cacheLocation: "localStorage" }
    });
    return msalInstance;
  }

  async function getGraphToken() {
    const app = await getMsal();
    const request = { scopes: GRAPH_SCOPES };
    try {
      const result = await app.acquireTokenSilent(request);
      return result.accessToken;
    } catch (silentError) {
      if (silentError instanceof msal.InteractionRequiredAuthError) {
        const result = await app.acquireTokenPopup(request);
        return result.accessToken;
      }
      throw silentError;
    }
  }

  async function sendMail(message) {
    const token = await getGraphToken();
    const response = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
      method: "POST",
      headers: {
        Authorization: "Bearer " + token,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ message: message, saveToSentItems: true })
    });
    // Graph answers 202 Accepted with an empty body on success.
    if (response.status !== 202) {
      const detail = await response.text();
      throw new Error("Graph sendMail failed (" + response.status + "): " + detail);
    }
  }

  window.AddinAuth = { sendMail: sendMail };
})();
