// Copyright (c) Djordje Radulovic. Licensed under the MIT license. See LICENSE in the project root.
// Settings you change before running the add-in. See README.md, "Setup".
"use strict";

window.ADDIN_CONFIG = {
  // Must match <Id> in actionable-add-in.xml.
  addInId: "5e61056d-d67d-46e5-b001-781e266da729",

  // Application (client) ID of your Microsoft Entra app registration.
  // The registration needs the SPA redirect URI brk-multihub://localhost:3000
  // and the delegated Microsoft Graph permission Mail.Send.
  clientId: "YOUR-ENTRA-APP-CLIENT-ID",

  // Provider ID from the Actionable Email Developer Dashboard (https://aka.ms/publishoam).
  // A provider with the "Test Users" scope is enough to send cards to your own mailbox.
  originator: "YOUR-ACTIONABLE-MESSAGE-PROVIDER-ID",

  // ID of the Message Header Analyzer store add-in, used to demo on-demand install.
  storeAddInId: "62916641-fc48-44ae-a2a3-163811f1c945"
};
