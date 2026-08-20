// src/taskpane/graphService.ts
import { PublicClientApplication, AccountInfo, InteractionRequiredAuthError } from "@azure/msal-browser";
import { msalConfig, graphScopes, SP_HOST, SP_SITE_PATH, SP_LIST_NAME } from "./authConfig";

const GRAPH = "https://graph.microsoft.com/v1.0";

export const USE_MOCK = false; // true = dane wymyslone, bez wywolan do Graph

let msal: PublicClientApplication | null = null;
let initPromise: Promise<PublicClientApplication> | null = null;

/** Inicjalizacja tylko raz, nawet przy kilku rownoleglych wywolaniach. */
function getMsal(): Promise<PublicClientApplication> {
  if (!initPromise) {
    initPromise = (async () => {
      const app = new PublicClientApplication(msalConfig);
      await app.initialize();
      // KLUCZOWE: czysci flage "interaction_in_progress" po przerwanym logowaniu
      await app.handleRedirectPromise();
      msal = app;
      return app;
    })();
  }
  return initPromise;
}

let tokenPromise: Promise<string> | null = null;

/**
 * Token do Graph. Rownolegle wywolania dostaja te sama obietnice,
 * wiec popup logowania otwiera sie najwyzej raz.
 */
export function getToken(): Promise<string> {
  if (!tokenPromise) {
    tokenPromise = acquire().finally(() => {
      tokenPromise = null;
    });
  }
  return tokenPromise;
}

function tokenViaOfficeDialog(): Promise<string> {
  return new Promise((resolve, reject) => {
    const url = `${window.location.origin}/dialog.html`;

    Office.context.ui.displayDialogAsync(
      url,
      { height: 60, width: 30, promptBeforeOpen: false },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error(`Nie udalo sie otworzyc okna logowania: ${result.error?.message}`));
          return;
        }

        const dialog = result.value;

        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg: any) => {
          dialog.close();
          try {
            const payload = JSON.parse(arg.message);
            if (payload.ok && payload.token) resolve(payload.token);
            else reject(new Error(payload.error || "Logowanie nie zwrocilo tokenu"));
          } catch {
            reject(new Error("Nieczytelna odpowiedz z okna logowania"));
          }
        });

        dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg: any) => {
          reject(new Error(arg.error === 12006 ? "Logowanie anulowane" : `Blad okna: ${arg.error}`));
        });
      }
    );
  });
}

async function acquire(): Promise<string> {
  const app = await getMsal();
  const account = app.getAllAccounts()[0] ?? null;

  if (account) {
    try {
      const res = await app.acquireTokenSilent({ scopes: graphScopes, account });
      return res.accessToken;
    } catch {
      // token wygasl -> logowanie przez dialog
    }
  }

  return tokenViaOfficeDialog();
}

/** Awaryjne czyszczenie stanu MSAL - do podpiecia pod przycisk, gdy logowanie sie zablokuje. */
export async function resetAuth(): Promise<void> {
  const app = await getMsal();
  app.getAllAccounts().forEach((a) => app.setActiveAccount(null) as void);
  Object.keys(localStorage)
    .filter((k) => k.startsWith("msal."))
    .forEach((k) => localStorage.removeItem(k));
  Object.keys(sessionStorage)
    .filter((k) => k.startsWith("msal."))
    .forEach((k) => sessionStorage.removeItem(k));
}

async function graphGet<T>(path: string, token: string, extraHeaders: Record<string, string> = {}): Promise<T> {
  const r = await fetch(`${GRAPH}${path}`, {
    headers: { Authorization: `Bearer ${token}`, ...extraHeaders },
  });
  if (!r.ok) {
    const body = await r.text();
    throw new Error(`Graph ${r.status}: ${body.slice(0, 300)}`);
  }
  return (await r.json()) as T;
}

export type ListItem = { id: string; fields: Record<string, unknown> };

const MOCK: ListItem[] = [
  { id: "1", fields: { Title: "Firma A Sp. z o.o.", Email: "kontakt@firma-a.pl", Opiekun: "JK" } },
  { id: "2", fields: { Title: "Firma B Consulting", Email: "biuro@firma-b.pl", Opiekun: "MT" } },
];

/** site-id jest potrzebny, bo Graph nie przyjmuje sciezki witryny. Cache, bo sie nie zmienia. */
let siteIdCache: string | null = null;
async function getSiteId(token: string): Promise<string> {
  if (!siteIdCache) {
    const site = await graphGet<{ id: string }>(`/sites/${SP_HOST}:${SP_SITE_PATH}`, token);
    siteIdCache = site.id;
  }
  return siteIdCache;
}

const listIdCache = new Map<string, string>();
async function getListId(siteId: string, name: string, token: string): Promise<string> {
  const hit = listIdCache.get(name);
  if (hit) return hit;
  const res = await graphGet<{ value: Array<{ id: string; displayName: string }> }>(
    `/sites/${siteId}/lists?$filter=displayName eq '${name}'`,
    token
  );
  if (!res.value.length) throw new Error(`Nie znaleziono listy "${name}"`);
  listIdCache.set(name, res.value[0].id);
  return res.value[0].id;
}

export async function getListItems(top = 20, listName = SP_LIST_NAME, query = ""): Promise<ListItem[]> {
  if (USE_MOCK) return MOCK;
  const token = await getToken();
  const siteId = await getSiteId(token);
  const listId = await getListId(siteId, listName, token);
  const res = await graphGet<{ value: ListItem[] }>(
    `/sites/${siteId}/lists/${listId}/items?expand=fields&$top=${top}${query}`,
    token
  );
  return res.value;
}
export type ColumnDef = {
  name: string;
  displayName: string;
  hidden?: boolean;
  readOnly?: boolean;
  [k: string]: unknown;
};

/** Pelna definicja kolumn listy - takze tych pustych w rekordach. */
export async function getListColumns(listName = SP_LIST_NAME): Promise<ColumnDef[]> {
  const token = await getToken();
  const siteId = await getSiteId(token);
  const listId = await getListId(siteId, listName, token);
  const res = await graphGet<{ value: ColumnDef[] }>(
    `/sites/${siteId}/lists/${listId}/columns`,
    token
  );
  return res.value;
}
/**
 * Zgloszenia filtrowane po kolumnie. Naglowek Prefer pozwala pytac
 * po kolumnach nieindeksowanych - inaczej Graph odrzuca zapytanie.
 */
export async function getTicketsByField(
  field: string,
  value: string,
  top = 50,
  listName = "Support"
): Promise<ListItem[]> {
  const token = await getToken();
  const siteId = await getSiteId(token);
  const listId = await getListId(siteId, listName, token);
  const safe = value.replace(/'/g, "''");
  const res = await graphGet<{ value: ListItem[] }>(
    `/sites/${siteId}/lists/${listId}/items?expand=fields&$top=${top}&$filter=fields/${field} eq '${safe}'&$orderby=fields/TicketID desc`,
    token,
    { Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" }
  );
  return res.value;
}
// ===== Dopisz na koncu src/taskpane/graphService.ts =====

/** Zapis pol elementu listy. Zwraca zaktualizowane pola. */
export async function updateTicketFields(
  itemId: string,
  fields: Record<string, unknown>,
  etag?: string,
  listName = "Support"
): Promise<Record<string, unknown>> {
  const token = await getToken();
  const siteId = await getSiteId(token);
  const listId = await getListId(siteId, listName, token);

  const headers: Record<string, string> = {
    Authorization: `Bearer ${token}`,
    "Content-Type": "application/json",
  };
  // Optymistyczna blokada: jesli ktos zmienil rekord w miedzyczasie, dostaniemy 412
  if (etag) headers["If-Match"] = etag;

  const r = await fetch(`${GRAPH}/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, {
    method: "PATCH",
    headers,
    body: JSON.stringify(fields),
  });

  if (r.status === 412) {
    throw new Error("Ktoś zmienił to zgłoszenie w międzyczasie. Odśwież i spróbuj ponownie.");
  }
  if (!r.ok) {
    const body = await r.text();
    throw new Error(`Graph ${r.status}: ${body.slice(0, 300)}`);
  }
  return (await r.json()) as Record<string, unknown>;
}
// ===== Dopisz na koncu src/taskpane/graphService.ts =====

/**
 * Firmy. Graph bywa kapryśny przy startswith na kolumnach SharePointa,
 * więc próbujemy po stronie serwera, a w razie odmowy filtrujemy lokalnie.
 */
export async function getCompanies(search = "", top = 200): Promise<ListItem[]> {
  const token = await getToken();
  const siteId = await getSiteId(token);
  const listId = await getListId(siteId, "Companies", token);
  const base = `/sites/${siteId}/lists/${listId}/items?expand=fields&$top=${top}`;
  const prefer = { Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" };

  if (search.trim()) {
    const safe = search.trim().replace(/'/g, "''");
    try {
      const res = await graphGet<{ value: ListItem[] }>(
        `${base}&$filter=startswith(fields/Nazwa,'${safe}')`,
        token,
        prefer
      );
      if (res.value.length) return res.value;
    } catch {
      // serwer odmówił - lecimy dalej i przefiltrujemy lokalnie
    }
  }

  const res = await graphGet<{ value: ListItem[] }>(
    `${base}&$orderby=fields/TECSh desc`,
    token,
    prefer
  );

  if (!search.trim()) return res.value;

  const q = search.trim().toLowerCase();
  return res.value.filter((i) =>
    String((i.fields as any).Nazwa ?? "").toLowerCase().startsWith(q)
  );
}
// ===== Dopisz na koncu src/taskpane/graphService.ts =====

/** Handlowcy z listy MEScoStaff, kolumna CRM_outlook. Format zgodny z Companies.Creator. */
let opiekunowieCache: string[] | null = null;

export async function getOpiekunowie(): Promise<string[]> {
  if (opiekunowieCache) return opiekunowieCache;

  const token = await getToken();
  const siteId = await getSiteId(token);
  const listId = await getListId(siteId, "MEScoStaff", token);

  const res = await graphGet<{ value: ListItem[] }>(
    `/sites/${siteId}/lists/${listId}/items?expand=fields&$top=200`,
    token,
    { Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" }
  );

  const zbior = new Set<string>();
  res.value.forEach((i) => {
    const f = i.fields as any;
    const v = String(f.CRM_Outlook ?? f.CRM_outlook ?? "").trim();
    if (v) zbior.add(v);
  });

  opiekunowieCache = [...zbior].sort((a, b) => a.localeCompare(b, "pl"));
  return opiekunowieCache;
}

/**
 * Firmy z opcjonalnym filtrem po nazwie i opiekunie.
 * Oba filtry idą na serwer, więc wynik obejmuje całą listę, nie tylko pobraną porcję.
 */
// ===== W src/taskpane/graphService.ts =====
// Zamień CAŁĄ funkcję getCompanies2 (razem z sortowaniem, które dodawaliśmy)
// na poniższe. getOpiekunowie zostaw bez zmian.

/** Wszystkie firmy, pobrane raz na sesję. Graph oddaje po 200 na stronę. */
let firmyCache: ListItem[] | null = null;

export async function getAllCompanies(force = false): Promise<ListItem[]> {
  if (firmyCache && !force) return firmyCache;

  const token = await getToken();
  const siteId = await getSiteId(token);
  const listId = await getListId(siteId, "Companies", token);

  const zebrane: ListItem[] = [];
  let url: string | null =
    `${GRAPH}/sites/${siteId}/lists/${listId}/items?expand=fields&$top=200`;

  // zabezpieczenie przed pętlą, gdyby serwer zwracał wciąż ten sam link
  let strony = 0;

  while (url && strony < 50) {
    const r: Response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly",
      },
    });
    if (!r.ok) {
      const body = await r.text();
      throw new Error(`Graph ${r.status}: ${body.slice(0, 300)}`);
    }
    const dane: { value: ListItem[]; "@odata.nextLink"?: string } = await r.json();
    zebrane.push(...dane.value);
    url = dane["@odata.nextLink"] ?? null;
    strony++;
  }

  firmyCache = zebrane;
  return zebrane;
}

/** Filtrowanie i sortowanie lokalne - działa na pełnym zbiorze, bez limitu 200. */
export async function getCompanies2(
  search = "",
  opiekun = "",
  _top = 0
): Promise<ListItem[]> {
  const wszystkie = await getAllCompanies();

  const q = search.trim().toLowerCase();
  const o = opiekun.trim().toLowerCase();

  return wszystkie
    .filter((i) => {
      const f = i.fields as any;
      const nazwa = String(f.Nazwa ?? "").toLowerCase();
      const creator = String(f.Creator ?? "").toLowerCase();
      return (!q || nazwa.startsWith(q)) && (!o || creator === o);
    })
    .sort(
      (a, b) =>
        Number((b.fields as any).TECSh ?? 0) - Number((a.fields as any).TECSh ?? 0)
    );
}

/** Do przycisku odświeżania - wymusza ponowne pobranie firm. */
export function wyczyscCacheFirm(): void {
  firmyCache = null;
}
// ===== Dopisz na koncu src/taskpane/graphService.ts =====

/** Kontakt po adresie e-mail. Zwraca null, gdy nadawcy nie ma w bazie. */
export async function getContactByEmail(email: string): Promise<ListItem | null> {
  const token = await getToken();
  const siteId = await getSiteId(token);
  const listId = await getListId(siteId, "Contacts", token);
  const safe = email.trim().replace(/'/g, "''");

  const res = await graphGet<{ value: ListItem[] }>(
    `/sites/${siteId}/lists/${listId}/items?expand=fields&$top=1&$filter=fields/Email eq '${safe}'`,
    token,
    { Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" }
  );
  return res.value[0] ?? null;
}

/** Nowy wiersz w Support. TicketID nadaje przepływ jako ID + 10000. */
export async function createTicket(
  fields: Record<string, unknown>,
  listName = "Support"
): Promise<{ id: string; fields: Record<string, unknown> }> {
  const token = await getToken();
  const siteId = await getSiteId(token);
  const listId = await getListId(siteId, listName, token);

  const r = await fetch(`${GRAPH}/sites/${siteId}/lists/${listId}/items`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ fields }),
  });

  if (!r.ok) {
    const body = await r.text();
    throw new Error(`Graph ${r.status}: ${body.slice(0, 300)}`);
  }
  return (await r.json()) as { id: string; fields: Record<string, unknown> };
}
// ===== Dopisz na koncu src/taskpane/graphService.ts =====

export type Serwisant = {
  nazwa: string;      // SupportName - DisplayName
  mail: string;       // Support - adres e-mail
  siteId: number | null; // SupportID - Id uzytkownika witryny (do pola Osoba)
};

let serwisanciCache: Serwisant[] | null = null;

/** Serwisanci z listy MEScoStaff. Pobierane raz na sesje. */
export async function getSerwisanci(): Promise<Serwisant[]> {
  if (serwisanciCache) return serwisanciCache;

  const token = await getToken();
  const siteId = await getSiteId(token);
  const listId = await getListId(siteId, "MEScoStaff", token);

  const res = await graphGet<{ value: ListItem[] }>(
    `/sites/${siteId}/lists/${listId}/items?expand=fields&$top=200`,
    token,
    { Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" }
  );

  const lista: Serwisant[] = [];
  res.value.forEach((i) => {
    const f = i.fields as any;
    const nazwa = String(f.SupportName ?? "").trim();
    const mail = String(f.Support ?? "").trim().toLowerCase();
    if (!nazwa || !mail) return;

    const raw = String(f.SupportID ?? "").trim();
    const num = Number(raw);
    // Pole Osoba w SharePoint wymaga Id uzytkownika WITRYNY (liczba, np. 50),
    // nie GUID z Entra. Jesli w kolumnie jest cos innego - pomijamy.
    const siteUserId = raw && Number.isInteger(num) && num > 0 ? num : null;

    lista.push({ nazwa, mail, siteId: siteUserId });
  });

  serwisanciCache = lista.sort((a, b) => a.nazwa.localeCompare(b.nazwa, "pl"));
  return serwisanciCache;
}
// ===== Dopisz na koncu src/taskpane/graphService.ts =====

/** Kontakty przypisane do firmy po IdentyfikatorCompanies. */
export async function getContactsByCompany(companyGuid: string, top = 100): Promise<ListItem[]> {
  const token = await getToken();
  const siteId = await getSiteId(token);
  const listId = await getListId(siteId, "Contacts", token);
  const safe = companyGuid.trim().replace(/'/g, "''");

  const res = await graphGet<{ value: ListItem[] }>(
    `/sites/${siteId}/lists/${listId}/items?expand=fields&$top=${top}` +
      `&$filter=fields/IdentyfikatorCompanies eq '${safe}'`,
    token,
    { Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" }
  );
  return res.value;
}
// ===== Dopisz na koncu src/taskpane/graphService.ts =====

/**
 * Zgloszenia z laczonymi filtrami. Puste wartosci pomijamy,
 * wiec brak filtrow = najnowsze zgloszenia z calej listy.
 */
export async function getTicketsFiltered(
  opts: { osobaMail?: string; mailCreator?: string; companyGuid?: string; top?: number } = {}
): Promise<ListItem[]> {
  const token = await getToken();
  const siteId = await getSiteId(token);
  const listId = await getListId(siteId, "Support", token);

  const esc = (v: string) => v.trim().replace(/'/g, "''");
  const warunki: string[] = [];
  if (opts.osobaMail?.trim()) warunki.push(`fields/OsobaMail eq '${esc(opts.osobaMail)}'`);
  if (opts.mailCreator?.trim()) warunki.push(`fields/MailCreator eq '${esc(opts.mailCreator)}'`);
  if (opts.companyGuid?.trim()) warunki.push(`fields/IdentyfikatorCompanies eq '${esc(opts.companyGuid)}'`);

  let url = `/sites/${siteId}/lists/${listId}/items?expand=fields&$top=${opts.top ?? 50}`;
  if (warunki.length) url += `&$filter=${warunki.join(" and ")}`;
  url += `&$orderby=fields/TicketID desc`;

  const res = await graphGet<{ value: ListItem[] }>(url, token, {
    Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly",
  });
  return res.value;
}