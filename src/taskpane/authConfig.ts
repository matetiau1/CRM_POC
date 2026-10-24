// src/taskpane/authConfig.ts
import { Configuration } from "@azure/msal-browser";

export const msalConfig: Configuration = {
  auth: {
    clientId: "7277cae3-d05a-4a6f-8afe-2edb6b9b6365",
    authority: "https://login.microsoftonline.com/4ad84be0-f918-4dfa-8006-a408554ec666",
        redirectUri: `${window.location.origin}${window.location.pathname}`,
  },
  cache: {
    cacheLocation: "localStorage",
  },
};

export const graphScopes = ["User.Read", "Sites.ReadWrite.All"];

// SharePoint
export const SP_HOST = "mescocompl.sharepoint.com";
export const SP_SITE_PATH = "/sites/CRMSystemPOC";
export const SP_LIST_NAME = "Support";