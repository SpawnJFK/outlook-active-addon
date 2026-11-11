import js from "@eslint/js";
import globals from "globals";

const officeGlobals = {
  Office: "readonly",
  msal: "readonly",
  AddinAuth: "readonly",
  AddinCard: "readonly"
};

export default [
  { ignores: ["node_modules/**"] },
  js.configs.recommended,
  {
    files: ["config.js", "shared/**/*.js", "dialog/**/*.js", "taskpane/**/*.js", "functionfile/**/*.js"],
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: "script",
      globals: { ...globals.browser, ...officeGlobals }
    }
  },
  {
    files: ["scripts/**/*.mjs", "eslint.config.mjs"],
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: "module",
      globals: globals.node
    }
  }
];
