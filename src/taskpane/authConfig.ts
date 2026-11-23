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

// Skrzynka wspoldzielona, na ktorej dziala przeplyw zgloszen
export const SUPPORT_MAILBOX = "support@mesco.com.pl";

// Pelna aplikacja CRM w Power Apps - przycisk "CRM" w dolnym pasku
export const POWERAPPS_URL =
  "https://apps.powerapps.com/play/e/default-4ad84be0-f918-4dfa-8006-a408554ec666" +
  "/a/47ff301e-7b3e-4df7-8988-e916ea18451b?tenantId=4ad84be0-f918-4dfa-8006-a408554ec666";
/** Zakres tokenu do Power Automate. Podwojny ukosnik jest celowy:
 *  identyfikator zasobu konczy sie ukosnikiem, pojedynczy daje blad uprawnien. */
export const flowScopes = ["https://service.flow.microsoft.com//User"];

/** Adres wyzwalacza HTTP przeplywu Support_TicketCRM_forward_HTTP (wymaga tokenu konta z tenanta). */
export const FLOW_URL =
  "https://default4ad84be0f9184dfa8006a408554ec6.66.environment.api.powerplatform.com:443" +
  "/powerautomate/automations/direct/cu/16/workflows/443eadb1013444258afc8523c7ccf7e8" +
  "/triggers/manual/paths/invoke?api-version=1";

/**
 * Wyzwalacze HTTP generatorow ofert ("GeneratorOfert... http", opis w http.md), klucz = TypBiznesu szansy
 * albo "TECS". Adres i tak trafia do przegladarki w taskpane.js - chroni go wymagany token konta z tenanta.
 */
export const OFERTA_URL: Record<string, string> = {
  "Paid-up": "https://default4ad84be0f9184dfa8006a408554ec6.66.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/09/workflows/bd6404ba0dca4d84a36b1360a8a2dcfc/triggers/manual/paths/invoke?api-version=1",
  Lease: "https://default4ad84be0f9184dfa8006a408554ec6.66.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/13/workflows/da7d375630334bdaba41209dbe93eb5f/triggers/manual/paths/invoke?api-version=1",
  AEC: "https://default4ad84be0f9184dfa8006a408554ec6.66.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/21/workflows/39067dbd491e46ab99c49da7cf995899/triggers/manual/paths/invoke?api-version=1",
  TECS: "https://default4ad84be0f9184dfa8006a408554ec6.66.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/28/workflows/888e90ea54074e0cbeef3d4dace5fded/triggers/manual/paths/invoke?api-version=1",
};
