# Outlook add-in: actionable message activation

An Outlook add-in that sends you an actionable message (an Adaptive Card in an email). One of the card's buttons opens the add-in's task pane and hands it an initialization context, so the add-in starts with the data from that specific email.

The add-in adds two ribbon buttons when you read a message:

- **Send Add-in Activation** opens a dialog where you edit a JSON context. It then sends you an email with a card containing two buttons:
  - **Invoke "View Initialization Context"** opens this add-in's task pane with the context.
  - **Invoke Message Header Analyzer** opens a store add-in. If you don't have it, Outlook offers to install it on demand.
- **View Initialization Context** opens the task pane by hand. With no context it tells you to open it from the card instead.

## What changed in 1.1 (2026)

Version 1.0 (2018) was built on APIs Microsoft has since retired. This version brings it up to date:

| Area | 1.0 (2018) | 1.1 (2026) |
| --- | --- | --- |
| Sending mail | Outlook REST v2.0 with `getCallbackTokenAsync({ isRest: true })`. The endpoint has been retired. | Microsoft Graph `/me/sendMail`, with tokens from nested app authentication (MSAL) |
| Card format | Legacy `MessageCard` | Adaptive Card 1.0 with `Action.InvokeAddInCommand` |
| Office.js startup | `Office.initialize` | `Office.onReady`, plus `Office.actions.associate` for the ribbon function |
| Libraries | jQuery, URI.js, moment.js, Office UI Fabric CDN | No runtime dependencies except Office.js and msal-browser. Plain DOM code and its own small stylesheet, with dark mode. |
| Task pane | The `InitializationContextChanged` handler only logged its arguments | The handler renders the new context when you click a card button with the pane already open |
| Tooling | None | Local HTTPS dev server, ESLint |

## Setup

You need Node.js 20 or later, a Microsoft 365 account, and Outlook on the web, new Outlook for Windows or a current Outlook for Mac. Nested app authentication isn't available in older Outlook builds.

1. **Register an app in Microsoft Entra ID** (Azure portal, then App registrations, then New registration):
   - Supported account types: any organizational directory and personal Microsoft accounts.
   - Redirect URI: platform **Single-page application**, value `brk-multihub://localhost:3000`.
   - Under API permissions, add the delegated Microsoft Graph permission **Mail.Send**.
   - Copy the Application (client) ID into `clientId` in [`config.js`](config.js).
2. **Register an actionable message provider** in the [Actionable Email Developer Dashboard](https://aka.ms/publishoam). The **Test Users** scope is enough for cards sent to your own mailbox. Copy the provider ID into `originator` in `config.js`.
3. Install and trust the local dev certificate, then start the server:

   ```bash
   npm install
   npm run certs
   npm start
   ```

   The add-in is now served at `https://localhost:3000`.
4. [Sideload](https://learn.microsoft.com/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing) `actionable-add-in.xml` in Outlook.

To host it somewhere other than localhost, replace every `https://localhost:3000` in the manifest with your HTTPS origin and add `brk-multihub://<your-domain>` as a redirect URI.

## Project layout

```
actionable-add-in.xml   add-in manifest (ribbon buttons, task pane, function file)
config.js               client ID, provider ID and add-in ID
shared/auth.js          MSAL nested app auth + Graph sendMail
shared/card.js          builds the Adaptive Card email
functionfile/           ribbon button handler (opens the dialog, sends the mail)
dialog/                 JSON context editor
taskpane/               shows the initialization context
scripts/serve.mjs       HTTPS dev server on port 3000
```

## Scripts

| Command | What it does |
| --- | --- |
| `npm start` | Serves the add-in over HTTPS on port 3000 |
| `npm run certs` | Installs the trusted Office dev certificate |
| `npm run lint` | Runs ESLint |
| `npm run validate` | Validates the manifest with Microsoft's validation service |

## Credits

Based on Microsoft's 2018 Outlook sample for [invoking an add-in from an actionable message](https://learn.microsoft.com/outlook/actionable-messages/invoke-add-in), released under the MIT license. The 2026 update is by Djordje Radulovic. See [LICENSE](LICENSE).
