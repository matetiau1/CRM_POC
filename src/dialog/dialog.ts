// src/dialog/dialog.ts
// Strona otwierana przez Office Dialog API. Loguje uzytkownika i odsyla token do panelu.
/* global Office, document, window */
import { PublicClientApplication } from "@azure/msal-browser";
import { msalConfig, graphScopes, flowScopes, SP_HOST } from "../taskpane/authConfig";

const dialogRedirect = `${window.location.origin}/dialog.html`;

// Jedyne zakresy, o jakie panel kiedykolwiek prosi (acquire() w graphService.ts) -
// parametr z URL nie moze wymusic logowania z dowolnym, obcym zakresem.
const ZNANE_ZAKRESY = new Set([...graphScopes, ...flowScopes, "Contacts.ReadWrite", `https://${SP_HOST}/AllSites.Write`]);

// panel moze poprosic o dodatkowy zakres (np. Contacts.ReadWrite) - ?scopes=a b
const zParametru = new URLSearchParams(window.location.search).get("scopes");
const zadane = zParametru ? zParametru.split(" ").filter(Boolean) : [];
const scopes = zadane.length && zadane.every((s) => ZNANE_ZAKRESY.has(s)) ? zadane : graphScopes;

Office.onReady(async () => {
  try {
    const app = new PublicClientApplication({
      ...msalConfig,
      auth: { ...msalConfig.auth, redirectUri: dialogRedirect },
    });
    await app.initialize();

    // Powrot z logowania? Tu zlapiemy wynik.
    const redirectResult = await app.handleRedirectPromise();

    if (redirectResult?.accessToken) {
      send({ ok: true, token: redirectResult.accessToken });
      return;
    }

    // Jest juz konto - probujemy po cichu (refresh token, bez ciasteczek).
    const account = app.getAllAccounts()[0];
    if (account) {
      try {
        const silent = await app.acquireTokenSilent({ scopes, account });
        send({ ok: true, token: silent.accessToken });
        return;
      } catch {
        // brak waznego refresh tokenu -> logowanie interaktywne ponizej
      }
    }

    // Pelne logowanie: przekierowanie w obrebie okna dialogu.
    await app.loginRedirect({ scopes, redirectUri: dialogRedirect });
  } catch (e: any) {
    show(`Blad logowania: ${e?.message ?? e}`);
    send({ ok: false, error: String(e?.message ?? e) });
  }
});

function send(payload: { ok: boolean; token?: string; error?: string }) {
  if (Office?.context?.ui?.messageParent) {
    Office.context.ui.messageParent(JSON.stringify(payload));
  } else {
    // poza Outlookiem nie ma do kogo wyslac - pokazujemy w oknie
    show(payload.ok ? "Token uzyskany (test poza Outlookiem)" : `Blad: ${payload.error}`);
  }
}

function show(msg: string) {
  const box = document.getElementById("box");
  if (!box) return;
  box.innerHTML = "";
  const p = document.createElement("p");
  p.className = "err";
  p.textContent = msg;
  box.appendChild(p);
}
