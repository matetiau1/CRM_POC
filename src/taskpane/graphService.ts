// src/taskpane/graphService.ts
import {
  PublicClientApplication,
  createNestablePublicClientApplication,
  InteractionRequiredAuthError,
} from "@azure/msal-browser";
import type { IPublicClientApplication } from "@azure/msal-browser";
import { msalConfig, graphScopes, SP_HOST, SP_SITE_PATH, SP_LIST_NAME } from "./authConfig";

const GRAPH = "https://graph.microsoft.com/v1.0";

export const USE_MOCK = false; // true = dane wymyslone, bez wywolan do Graph

/** NAA dziala w nowym Outlooku i Outlooku na webie; classic Outlook na Macu go nie ma. */
function naaDostepne(): boolean {
  try {
    return Office.context.requirements.isSetSupported("NestedAppAuth", "1.1");
  } catch {
    return false;
  }
}

let app: IPublicClientApplication | null = null;
let initPromise: Promise<IPublicClientApplication> | null = null;

function getMsal(): Promise<IPublicClientApplication> {
  if (!initPromise) {
    initPromise = (async () => {
      if (naaDostepne()) {
        // NAA: host (Outlook) brokeruje token, zadne okno nie jest otwierane.
        // Wymaga redirect URI typu SPA: brk-multihub://<origin>
        app = await createNestablePublicClientApplication(msalConfig);
      } else {
        const zwykly = new PublicClientApplication(msalConfig);
        await zwykly.initialize();
        // handleRedirectPromise NIE jest wspierane pod NAA - tylko w trybie zwyklym
        await zwykly.handleRedirectPromise();
        app = zwykly;
      }
      return app;
    })();
  }
  return initPromise;
}

/** Okno logowania przez Outlooka - sciezka awaryjna dla classic Outlook (macOS). */
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
  const instancja = await getMsal();
  const konto = instancja.getAllAccounts()[0] ?? null;

  // 1. Po cichu - dziala i pod NAA, i w trybie zwyklym
  try {
    const res = await instancja.acquireTokenSilent({
      scopes: graphScopes,
      ...(konto ? { account: konto } : {}),
    });
    return res.accessToken;
  } catch (e) {
    if (!(e instanceof InteractionRequiredAuthError) && konto) throw e;
    // brak konta albo wymagana interakcja -> nizej
  }

  // 2. Interaktywnie
  if (naaDostepne()) {
    // pod NAA popup jest wspierany i obsluguje go host
    const res = await instancja.acquireTokenPopup({ scopes: graphScopes });
    return res.accessToken;
  }

  // classic Outlook: popup nie ma kanalu powrotnego, wiec okno Office
  return tokenViaOfficeDialog();
}

let tokenPromise: Promise<string> | null = null;

/**
 * Token do Graph. Rownolegle wywolania dostaja te sama obietnice,
 * wiec logowanie uruchamia sie najwyzej raz.
 */
export function getToken(): Promise<string> {
  if (!tokenPromise) {
    tokenPromise = acquire().finally(() => {
      tokenPromise = null;
    });
  }
  return tokenPromise;
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
/** Zalogowany uzytkownik z Graph. userProfile zwraca adres otwartej skrzynki,
 *  wiec na skrzynce wspoldzielonej dawalby support@mesco.com.pl zamiast osoby. */
let mojMailCache: string | null = null;

export async function getMojMail(): Promise<string> {
  if (mojMailCache) return mojMailCache;

  const token = await getToken();
  const me = await graphGet<{ mail?: string; userPrincipalName?: string }>(
    "/me?$select=mail,userPrincipalName",
    token
  );
  mojMailCache = String(me.mail || me.userPrincipalName || "").toLowerCase();
  return mojMailCache;
}
/**
 * Elementy dowolnej listy powiazane z firma.
 * Wiekszosc list uzywa IdentyfikatorCompanies, ale UczestnicySzkolen ma GUIDFirmy.
 */
export async function getByCompany(
  listName: string,
  guid: string,
  opts: { pole?: string; top?: number; orderby?: string } = {}
): Promise<ListItem[]> {
  const token = await getToken();
  const siteId = await getSiteId(token);
  const listId = await getListId(siteId, listName, token);
  const pole = opts.pole ?? "IdentyfikatorCompanies";
  const safe = guid.trim().replace(/'/g, "''");
 
  let url =
    `/sites/${siteId}/lists/${listId}/items?expand=fields&$top=${opts.top ?? 100}` +
    `&$filter=fields/${pole} eq '${safe}'`;
  if (opts.orderby) url += `&$orderby=${opts.orderby}`;
 
  const res = await graphGet<{ value: ListItem[] }>(url, token, {
    Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly",
  });
  return res.value;
}
 
/** Pelny rekord firmy po CompaniesGUID. */
export async function getCompanyByGuid(guid: string): Promise<ListItem | null> {
  const items = await getByCompany("Companies", guid, { pole: "CompaniesGUID", top: 1 });
  return items[0] ?? null;
}
/** Nowy element dowolnej listy. */
export async function createListItem(
  listName: string,
  fields: Record<string, unknown>
): Promise<ListItem> {
  const token = await getToken();
  const siteId = await getSiteId(token);
  const listId = await getListId(siteId, listName, token);

  const res = await fetch(`${GRAPH}/sites/${siteId}/lists/${listId}/items`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ fields }),
  });

  if (!res.ok) {
    const tresc = await res.text();
    throw new Error(`Nie zapisano (${res.status}): ${tresc.slice(0, 300)}`);
  }
  return res.json();
}

/** Pozycje listy slownikowej do pola typu lookup. Pobierane raz na sesje. */
const cacheSlownikow = new Map<string, { id: string; etykieta: string }[]>();

export async function getLookupOptions(
  listName: string,
  labelField: string
): Promise<{ id: string; etykieta: string }[]> {
  const klucz = `${listName}/${labelField}`;
  const z = cacheSlownikow.get(klucz);
  if (z) return z;

  const token = await getToken();
  const siteId = await getSiteId(token);
  const listId = await getListId(siteId, listName, token);

  const res = await graphGet<{ value: ListItem[] }>(
    `/sites/${siteId}/lists/${listId}/items?expand=fields&$top=100`,
    token
  );

  const lista = res.value
    .map((it) => ({
      id: String((it.fields as any).id ?? ""),
      etykieta: String((it.fields as any)[labelField] ?? ""),
    }))
    .filter((o) => o.id && o.etykieta);

  cacheSlownikow.set(klucz, lista);
  return lista;
}

/** Imie i nazwisko zalogowanego - do pola Creator w notatkach. */
let mojaNazwaCache: string | null = null;

export async function getMojaNazwa(): Promise<string> {
  if (mojaNazwaCache) return mojaNazwaCache;
  const token = await getToken();
  const me = await graphGet<{ displayName?: string }>("/me?$select=displayName", token);
  mojaNazwaCache = String(me.displayName ?? "");
  return mojaNazwaCache;
}
/** Usuniecie elementu z dowolnej listy. */
export async function deleteListItem(listName: string, itemId: string): Promise<void> {
  const token = await getToken();
  const siteId = await getSiteId(token);
  const listId = await getListId(siteId, listName, token);

  const res = await fetch(`${GRAPH}/sites/${siteId}/lists/${listId}/items/${itemId}`, {
    method: "DELETE",
    headers: { Authorization: `Bearer ${token}` },
  });

  if (!res.ok && res.status !== 204) {
    const tresc = await res.text();
    throw new Error(`Nie usunięto (${res.status}): ${tresc.slice(0, 300)}`);
  }
}
/** Zgloszenie po numerze z tematu wiadomosci (TicketID, bez prefiksu). */
export async function getTicketByNumber(nr: number): Promise<ListItem | null> {
  const token = await getToken();
  const siteId = await getSiteId(token);
  const listId = await getListId(siteId, "Support", token);

  const res = await graphGet<{ value: ListItem[] }>(
    `/sites/${siteId}/lists/${listId}/items?expand=fields&$top=1` +
      `&$filter=fields/TicketID eq ${Math.trunc(nr)}`,
    token,
    { Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" }
  );
  return res.value[0] ?? null;
}