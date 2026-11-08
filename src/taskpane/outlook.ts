// src/taskpane/outlook.ts
/* global document, Office, console, setTimeout, window, navigator, clearTimeout */
import {
  getListItems, getTicketsByField, updateTicketFields, getCompanies2, getOpiekunowie,
  wyczyscCacheFirm, getContactByEmail, createTicket, getSerwisanci, getContactsByCompany,
  getTicketsFiltered, getMojMail, getByCompany, getCompanyByGuid,createListItem, getLookupOptions, getMojaNazwa, deleteListItem,
  getTicketByNumber, getAllCompanies, getChoiceOptions, getAllContacts, wyczyscCacheKontaktow,
  getLiczbaSupportowRok, getSiteUserId, getCennik, getOpcjeLookupu,
} from "./graphService";
import type { ListItem, PozycjaCennika } from "./graphService";
import {
  toTicket, toCompany, formatData, dniOdZmiany, KOLOR_STATUSU, ticketIdZTematu, Ticket, Company,
  toContact, Contact, toCompanyFull, CompanyFull, toNota, toTecs, toSzansa,
  toLicencjaTecs, toLicencjaBez, Licencja, toSzkolenie, dniDo, kolorTerminu,
  ETAP_KOLORY, ETAP_SKROTY, rozbijNazwe, plecZImienia, nowyGuid, TecsRow,
} from "./crm";
import { SUPPORT_MAILBOX, POWERAPPS_URL } from "./authConfig";
import { pobierzPoNip, pelnaNazwa, nipPoprawny } from "./rejestr";

// ---------------------------------------------------------------- start

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) return;

  document.getElementById("sideload-msg")!.style.display = "none";
  document.getElementById("app-body")!.style.display = "block";

  document.getElementById("run")!.onclick = () => void odswiez();
  document.getElementById("nav-tickets")!.onclick = () => void pokazZgloszenia();
  document.getElementById("nav-companies")!.onclick = () => void pokazFirmy();
  document.getElementById("nav-contacts")!.onclick = () => void pokazKontakty();
  document.getElementById("nav-crm")!.onclick = () => window.open(POWERAPPS_URL, "_blank");

  Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => {
    if (widok === "zgloszenia") void pokazZgloszenia();
  });

  void pokazZgloszenia();
});

// ---------------------------------------------------------------- stan

type Widok = "zgloszenia" | "firmy" | "firma" | "kontakty";
let widok: Widok = "zgloszenia";
let ostatniaFirma: Company | null = null;

let zakresZgloszen: "mail" | "firma" | "wszystkie" = "mail";
let opiekunZgloszen = "";
let frazaKontaktow = "";

type Zakladka =
  | "Kontakty" | "Supporty" | "Notatki" | "TECS"
  | "Szanse" | "Posiadane" | "Szkolenia" | "Uwagi";

let zakladkaFirmy: Zakladka = "Kontakty";
let tecsTryb: "Bieżące" | "Archiwalne" | "Demo" = "Bieżące";
let otworzKodyDemo = false; // kafelek "Nowe kody demo" -> TECS/Demo z otwartym formularzem
let szanseZamkniete = false;
let posTryb: "Aktywne TECS" | "Wygasłe" | "Bez TECS" = "Aktywne TECS";

/** Pelne dane firmy - pobierane raz przy wejsciu w widok. */
let firmaPelna: CompanyFull | null = null;
let tecsAlert: { ile: number; dni: number } | null = null;
/** Kontakt nadawcy - pamietany, bo pyta o niego takze panelNowego. */
const kontaktCache = new Map<string, any>();

async function kontaktDlaNadawcy(email: string): Promise<any | null> {
  const klucz = email.toLowerCase();
  if (kontaktCache.has(klucz)) return kontaktCache.get(klucz);
  const k = await getContactByEmail(klucz);
  kontaktCache.set(klucz, k);
  return k;
}

// ---------------------------------------------------------------- pomocnicze

const root = () => document.getElementById("item-subject")!;

function el(tag: string, cls?: string, text?: string): HTMLElement {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text) e.textContent = text;
  return e;
}

type NavId = "nav-tickets" | "nav-companies" | "nav-contacts";

function ustawNaglowek(tytul: string, aktywny: NavId) {
  document.getElementById("nav-title")!.textContent = tytul;
  (["nav-tickets", "nav-companies", "nav-contacts"] as NavId[]).forEach((id) =>
    document.getElementById(id)!.classList.toggle("crm-dol-akt", id === aktywny)
  );
}

function status(tekst: string): HTMLElement {
  const s = el("div", "crm-status", tekst);
  root().appendChild(s);
  return s;
}

function parseGodziny(v: string): number {
  const s = v.trim().replace(",", ".");
  if (!s) return 0;
  const n = Number(s);
  return Number.isFinite(n) && n >= 0 ? n : NaN;
}

const dataPL = (iso: string) =>
  iso ? new Date(iso).toLocaleDateString("pl-PL", { day: "2-digit", month: "2-digit", year: "numeric" }) : "—";

const euro = (v: number) => v.toLocaleString("pl-PL", { maximumFractionDigits: 0 }) + " €";

const godziny = (v: number) =>
  v.toLocaleString("pl-PL", { maximumFractionDigits: v % 1 === 0 ? 0 : 1 }) + " h";

/** Schowek z zapasowa sciezka - navigator.clipboard bywa blokowany w panelu. */
async function kopiuj(tekst: string): Promise<boolean> {
  try {
    if (navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(tekst);
      return true;
    }
  } catch {
    // lecimy dalej
  }
  try {
    const ta = document.createElement("textarea");
    ta.value = tekst;
    ta.style.position = "fixed";
    ta.style.opacity = "0";
    document.body.appendChild(ta);
    ta.select();
    const ok = document.execCommand("copy");
    document.body.removeChild(ta);
    return ok;
  } catch {
    return false;
  }
}

/** Krotki komunikat przy elemencie, po kliknieciu "skopiowano". */
function blysk(obok: HTMLElement, tekst: string) {
  const b = el("span", "crm-blysk", tekst);
  obok.appendChild(b);
  setTimeout(() => b.remove(), 1400);
}

// ikony Lucide - ten sam zestaw co w dolnym pasku
const IKONY = {
  mail: '<rect width="20" height="16" x="2" y="4" rx="2"/><path d="m22 7-8.97 5.7a1.94 1.94 0 0 1-2.06 0L2 7"/>',
  firma: '<path d="M6 22V4a2 2 0 0 1 2-2h8a2 2 0 0 1 2 2v18Z"/><path d="M6 12H4a2 2 0 0 0-2 2v6a2 2 0 0 0 2 2h2"/><path d="M18 9h2a2 2 0 0 1 2 2v9a2 2 0 0 1-2 2h-2"/><path d="M10 6h4"/><path d="M10 10h4"/><path d="M10 14h4"/><path d="M10 18h4"/>',
  osoba: '<path d="M19 21v-2a4 4 0 0 0-4-4H9a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/>',
  sluchawki: '<path d="M3 11h3a2 2 0 0 1 2 2v3a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-5Zm0 0a9 9 0 1 1 18 0m0 0v5a2 2 0 0 1-2 2h-1a2 2 0 0 1-2-2v-3a2 2 0 0 1 2-2h3Z"/><path d="M21 16v2a4 4 0 0 1-4 4h-5"/>',
  zegar: '<circle cx="12" cy="12" r="10"/><path d="M12 6v6l4 2"/>',
  kalendarz: '<rect width="18" height="18" x="3" y="4" rx="2"/><path d="M16 2v4"/><path d="M8 2v4"/><path d="M3 10h18"/>',
  klodka: '<rect width="18" height="11" x="3" y="11" rx="2" ry="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/>',
  klodkaOtwarta: '<rect width="18" height="11" x="3" y="11" rx="2" ry="2"/><path d="M7 11V7a5 5 0 0 1 9.9-1"/>',
  klucz: '<path d="M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z"/>',
  czapka: '<path d="M21.42 10.922a1 1 0 0 0-.019-1.838L12.83 5.18a2 2 0 0 0-1.66 0L2.6 9.08a1 1 0 0 0 0 1.832l8.57 3.908a2 2 0 0 0 1.66 0z"/><path d="M22 10v6"/><path d="M6 12.5V16a6 3 0 0 0 12 0v-3.5"/>',
  pinezka: '<path d="M20 10c0 4.993-5.539 10.193-7.399 11.799a1 1 0 0 1-1.202 0C9.539 20.193 4 14.993 4 10a8 8 0 0 1 16 0"/><circle cx="12" cy="10" r="3"/>',
  opiekun: '<path d="M2 21a8 8 0 0 1 13.292-6"/><circle cx="10" cy="8" r="5"/><path d="m16 19 2 2 4-4"/>',
  tarcza: '<path d="M20 13c0 5-3.5 7.5-7.66 8.95a1 1 0 0 1-.67-.01C7.5 20.5 4 18 4 13V6a1 1 0 0 1 1-1c2 0 4.5-1.2 6.24-2.72a1.17 1.17 0 0 1 1.52 0C14.51 3.81 17 5 19 5a1 1 0 0 1 1 1z"/>',
  alert: '<path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3"/><path d="M12 9v4"/><path d="M12 17h.01"/>',
  kopiuj: '<rect width="14" height="14" x="8" y="8" rx="2" ry="2"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/>',
  folder: '<path d="M20 20a2 2 0 0 0 2-2V8a2 2 0 0 0-2-2h-7.9a2 2 0 0 1-1.69-.9L9.6 3.9A2 2 0 0 0 7.93 3H4a2 2 0 0 0-2 2v13a2 2 0 0 0 2 2Z"/>',
  wykres: '<path d="M22 7 13.5 15.5 8.5 10.5 2 17"/><path d="M16 7h6v6"/>',
  etykieta: '<path d="M12.586 2.586A2 2 0 0 0 11.172 2H4a2 2 0 0 0-2 2v7.172a2 2 0 0 0 .586 1.414l8.704 8.704a2.426 2.426 0 0 0 3.42 0l6.58-6.58a2.426 2.426 0 0 0 0-3.42z"/><circle cx="7.5" cy="7.5" r=".5" fill="currentColor"/>',
  paczka: '<path d="M11 21.73a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73z"/><path d="M12 22V12"/><path d="m3.3 7 7.977 4.947a2 2 0 0 0 2.046 0L20.7 7"/>',
  licencja: '<path d="M2.586 17.414A2 2 0 0 0 2 18.828V21a1 1 0 0 0 1 1h3a1 1 0 0 0 1-1v-1a1 1 0 0 1 1-1h1a1 1 0 0 0 1-1v-1a1 1 0 0 1 1-1h.172a2 2 0 0 0 1.414-.586l.814-.814a6.5 6.5 0 1 0-4-4z"/><circle cx="16.5" cy="7.5" r="1"/>',
  dokument: '<path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/><path d="M14 2v4a2 2 0 0 0 2 2h4"/><path d="M10 9H8"/><path d="M16 13H8"/><path d="M16 17H8"/>',
  pioro: '<path d="M12 20h9"/><path d="M16.376 3.622a1 1 0 0 1 3.002 3.002L7.368 18.635a2 2 0 0 1-.855.506l-2.872.838a.5.5 0 0 1-.62-.62l.838-2.872a2 2 0 0 1 .506-.854z"/>',
  dymek: '<path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>',
  dzwonek: '<path d="M10.268 21a2 2 0 0 0 3.464 0"/><path d="M3.262 15.326A1 1 0 0 0 4 17h16a1 1 0 0 0 .74-1.673C19.41 13.956 18 12.499 18 8A6 6 0 0 0 6 8c0 4.499-1.411 5.956-2.738 7.326"/>',
  cel: '<circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="6"/><circle cx="12" cy="12" r="2"/>',
  euro: '<path d="M4 10h12"/><path d="M4 14h9"/><path d="M19 6a7.7 7.7 0 0 0-5.2-2A7.9 7.9 0 0 0 6 12c0 4.4 3.5 8 7.8 8 2 0 3.8-.8 5.2-2"/>',
  klepsydra: '<path d="M5 22h14"/><path d="M5 2h14"/><path d="M17 22v-4.172a2 2 0 0 0-.586-1.414L12 12l-4.414 4.414A2 2 0 0 0 7 17.828V22"/><path d="M7 2v4.172a2 2 0 0 0 .586 1.414L12 12l4.414-4.414A2 2 0 0 0 17 6.172V2"/>',
  flaga: '<path d="M4 15s1-1 4-1 5 2 8 2 4-1 4-1V3s-1 1-4 1-5-2-8-2-4 1-4 1z"/><path d="M4 22v-7"/>',
  teczka: '<path d="M16 20V4a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16"/><rect width="20" height="14" x="2" y="6" rx="2"/>',
  kosz: '<path d="M3 6h18"/><path d="M19 6v14c0 1-1 2-2 2H7c-1 0-2-1-2-2V6"/><path d="M8 6V4c0-1 1-2 2-2h4c1 0 2 1 2 2"/>',
  telefon: '<path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72 12.84 12.84 0 0 0 .7 2.81 2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45 12.84 12.84 0 0 0 2.81.7A2 2 0 0 1 22 16.92z"/>',
};

function ikona(nazwa: keyof typeof IKONY): SVGSVGElement {
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("viewBox", "0 0 24 24");
  svg.setAttribute("aria-hidden", "true");
  svg.classList.add("crm-ikona");
  svg.innerHTML = IKONY[nazwa];
  return svg;
}

/** Zamienia element w przycisk kopiujacy podana wartosc. */
function zrobKopiowalny(elem: HTMLElement, wartosc: string, etykieta = "skopiowano") {
  elem.classList.add("crm-kopiuj");
  elem.title = "Kliknij, aby skopiować";
  elem.onclick = async (e: MouseEvent) => {
    e.stopPropagation(); // zeby nie rozwinac karty
    const ok = await kopiuj(wartosc);
    blysk(elem, ok ? etykieta : "nie udało się");
  };
}

/** Mala pastylka z kolorem tekstu i tla, opcjonalnie z ikona. */
function pastylka(tekst: string, kolor: string, tlo: string, ik?: keyof typeof IKONY): HTMLElement {
  const p = el("span", "crm-pill", ik ? undefined : tekst);
  if (ik) p.append(ikona(ik), el("span", undefined, tekst));
  p.style.color = kolor;
  p.style.background = tlo;
  return p;
}

/** Pole karty: <ikona> tekst. */
function poleIk(klasa: string, ik: keyof typeof IKONY, tekst: string, tytul?: string): HTMLElement {
  const w = el("div", `crm-ik ${klasa}`.trim());
  if (tytul) w.title = tytul;
  w.append(ikona(ik), el("span", "crm-ikona-txt", tekst));
  return w;
}

/** Przycisk z ikona zamiast emoji. */
function przyciskIk(klasa: string, ik: keyof typeof IKONY, tekst: string): HTMLButtonElement {
  const b = el("button", klasa) as HTMLButtonElement;
  b.append(ikona(ik), el("span", undefined, tekst));
  return b;
}

function pusto(box: HTMLElement, tekst: string) {
  box.appendChild(el("div", "crm-pusto", tekst));
}

/** Pasek wyboru trybu wewnatrz zakladki. */
function segment(box: HTMLElement, opcje: string[], aktywna: string, wybor: (v: string) => void) {
  const pasek = el("div", "crm-segment");
  opcje.forEach((o) => {
    const b = el("button", "crm-segment-btn", o) as HTMLButtonElement;
    if (o === aktywna) b.classList.add("crm-segment-akt");
    b.onclick = () => wybor(o);
    pasek.appendChild(b);
  });

  box.appendChild(pasek);
}

/** Pasek postepu do konca terminu - pelny rok = 100%. */
function pasekTerminu(dni: number | null): HTMLElement {
  const tor = el("div", "crm-tor");
  const bar = el("div", "crm-bar");
  const proc = dni === null ? 0 : Math.max(0, Math.min(1, dni / 365));
  bar.style.width = `${Math.max(3, proc * 100)}%`;
  bar.style.background = kolorTerminu(dni);
  tor.appendChild(bar);
  return tor;
}

/** Identyfikator wiadomosci w formacie REST - Office.js domyslnie zwraca EWS. */
function mailIdRest(item: Office.MessageRead): string {
  try {
    return Office.context.mailbox.convertToRestId(
      item.itemId,
      Office.MailboxEnums.RestVersion.v2_0
    );
  } catch {
    return item.itemId ?? "";
  }
}

/**
 * Czy otwarta wiadomosc lezy w skrzynce supportowej.
 * Skrzynka dodana jako osobne konto: userProfile zwraca jej adres.
 * Skrzynka otwarta jako folder wspoldzielony: adres daje getSharedPropertiesAsync.
 */
async function czySkrzynkaSupportu(item: Office.MessageRead): Promise<boolean> {
  const cel = SUPPORT_MAILBOX.toLowerCase();
  if ((Office.context.mailbox.userProfile.emailAddress ?? "").toLowerCase() === cel) return true;

  const it = item as any;
  if (typeof it.getSharedPropertiesAsync !== "function") return false;
  return new Promise((resolve) => {
    it.getSharedPropertiesAsync((r: Office.AsyncResult<Office.SharedProperties>) => {
      resolve(
        r.status === Office.AsyncResultStatus.Succeeded &&
          (r.value?.owner ?? "").toLowerCase() === cel
      );
    });
  });
}

// ---------------------------------------------------------------- zgloszenia

function karta(t: Ticket): HTMLElement {
  const kolor = KOLOR_STATUSU[t.status] ?? "#3860B2";

  const card = el("div", "crm-card");
  card.style.borderLeftColor = kolor;

  const head = el("div", "crm-card-head");
  head.tabIndex = 0;
  head.setAttribute("role", "button");

  // wiersz z ikona: <ikona> tekst
  const zIkona = (klasa: string, nazwa: keyof typeof IKONY, tekst: string, tytul: string) => {
    const w = el("span", klasa);
    w.title = tytul;
    const txt = el("span", "crm-ikona-txt", tekst);
    w.append(ikona(nazwa), txt);
    return { w, txt };
  };

  const top = el("div", "crm-card-top");
  top.appendChild(zIkona("crm-firma", "firma", t.firma || "(brak firmy)", "Firma").w);
  const idEl = el("span", "crm-id", `MESco-ID-${t.ticketId}`);
  zrobKopiowalny(idEl, `[MESco-ID-${t.ticketId}]`, "skopiowano numer");
  top.appendChild(idEl);
  head.appendChild(top);

  const r1 = el("div", "crm-row");
  r1.appendChild(zIkona("crm-osoba", "osoba", t.osoba || "—", "Zgłaszający").w);
  const czas = zIkona("crm-czas", "zegar", t.czas ? `${t.czas} h` : "—", "Czas obsługi");
  const czasEtykieta = czas.txt;
  r1.appendChild(czas.w);
  head.appendChild(r1);

  const r2 = el("div", "crm-row");
  r2.appendChild(zIkona("crm-opiekun", "sluchawki", t.opiekun || "—", "Obsługujący").w);
  head.appendChild(r2);

  // skrót do pełnego widoku firmy
  if (t.companyId) {
    const doFirmy = el("button", "crm-skrot", "↗ firma") as HTMLButtonElement;
    doFirmy.title = `Otwórz ${t.firma}`;
    doFirmy.onclick = (e: MouseEvent) => {
      e.stopPropagation(); // nie rozwijaj karty
      zakladkaFirmy = "Kontakty";
      void pokazFirma({ guid: t.companyId, nazwa: t.firma } as Company);
    };
    top.appendChild(doFirmy);
  }
  const foot = el("div", "crm-foot");
  foot.appendChild(zIkona("crm-data", "kalendarz", formatData(t.data), "Data zgłoszenia").w);
  const lock = el("span", "crm-lock");
  const ustawLock = () => {
    lock.replaceChildren(
      ikona(t.zamkniete ? "klodka" : "klodkaOtwarta"),
      el("span", undefined, t.zamkniete ? "zamknięte" : "otwarte")
    );
    lock.classList.toggle("crm-lock-zamk", t.zamkniete);
  };
  ustawLock();
  foot.appendChild(lock);
  if (t.status) {
    const chip = el("span", "crm-chip", t.status);
    chip.style.background = kolor;
    foot.appendChild(chip);
  }
  head.appendChild(foot);

  const dni = dniOdZmiany(t.modified);
  if (!t.zamkniete && dni > 7) head.appendChild(el("div", "crm-stale", `bez zmian od ${dni} dni`));

  card.appendChild(head);

  const odswiezNaglowek = () => {
    czasEtykieta.textContent = t.czas ? `${t.czas} h` : "—";
    ustawLock();
  };

  const detale = szczegoly(t, odswiezNaglowek);
  detale.style.display = "none";
  card.appendChild(detale);

  const przelacz = () => {
    const otwarte = detale.style.display !== "none";
    detale.style.display = otwarte ? "none" : "block";
    card.classList.toggle("crm-open", !otwarte);
  };
  head.onclick = przelacz;
  head.onkeydown = (e: KeyboardEvent) => {
    if (e.key === "Enter" || e.key === " ") { e.preventDefault(); przelacz(); }
  };

  return card;
}

function szczegoly(t: Ticket, poZapisie: () => void): HTMLElement {
  const box = el("div", "crm-detale");

  const info = el("div", "crm-meta");
  if (t.osobaMail) info.appendChild(el("div", undefined, `Kontakt: ${t.osobaMail}`));
  if (t.opiekunMail) info.appendChild(el("div", undefined, `Opiekun: ${t.opiekunMail}`));
  info.appendChild(el("div", undefined, `Ostatnia zmiana: ${formatData(t.modified)}`));
  box.appendChild(info);

  box.appendChild(el("label", "crm-label", "Czas (godziny)"));
  const czas = document.createElement("input");
  czas.type = "text";
  czas.inputMode = "decimal";
  czas.className = "crm-input";
  czas.placeholder = "np. 1,5";
  czas.value = t.czas ? String(t.czas) : "";
  box.appendChild(czas);

  box.appendChild(el("label", "crm-label", "Opis"));
  const ta = document.createElement("textarea");
  ta.className = "crm-textarea";
  ta.value = t.opis;
  ta.rows = 5;
  box.appendChild(ta);

  const zamkniete = document.createElement("input");
  zamkniete.type = "checkbox";
  zamkniete.checked = t.zamkniete;
  zamkniete.id = `zam-${t.id}`;
  const lbl = el("label", "crm-check") as HTMLLabelElement;
  lbl.htmlFor = zamkniete.id;
  lbl.appendChild(zamkniete);
  lbl.appendChild(el("span", undefined, "Zgłoszenie zamknięte"));
  box.appendChild(lbl);

  const podpowiedz = el("div", "crm-podpowiedz");
  box.appendChild(podpowiedz);

  const sprawdzWarunki = () => {
    const g = parseGodziny(czas.value);
    const brak: string[] = [];
    if (!(g > 0)) brak.push("czas");
    if (!ta.value.trim()) brak.push("opis");
    const ok = brak.length === 0;

    if (!ok && !zamkniete.checked) {
      zamkniete.disabled = true;
      podpowiedz.textContent = `Aby zamknąć zgłoszenie, uzupełnij: ${brak.join(" i ")}.`;
      podpowiedz.style.display = "block";
    } else {
      zamkniete.disabled = false;
      podpowiedz.style.display = "none";
    }
    return ok;
  };
  czas.oninput = sprawdzWarunki;
  ta.oninput = sprawdzWarunki;
  sprawdzWarunki();

  const pasek = el("div", "crm-akcje");
  const zapisz = el("button", "crm-btn", "Zapisz") as HTMLButtonElement;
  const komunikat = el("span", "crm-komunikat");
  pasek.appendChild(zapisz);
  pasek.appendChild(komunikat);
  box.appendChild(pasek);

  const ostrzezenie = el("div", "crm-ostrzezenie");
  ostrzezenie.style.display = "none";
  box.appendChild(ostrzezenie);

  const wykonajZapis = async () => {
    const g = parseGodziny(czas.value);
    if (Number.isNaN(g)) {
      komunikat.textContent = "Czas musi być liczbą (np. 1,5)";
      komunikat.className = "crm-komunikat crm-blad";
      return;
    }
    if (zamkniete.checked && !(g > 0 && ta.value.trim())) {
      komunikat.textContent = "Zamknięcie wymaga czasu i opisu";
      komunikat.className = "crm-komunikat crm-blad";
      return;
    }

    zapisz.disabled = true;
    komunikat.textContent = "Zapisywanie...";
    komunikat.className = "crm-komunikat";
    try {
      const pola: Record<string, unknown> = {
        Opis: ta.value,
        Czas: g,
        Status: zamkniete.checked,
      };

      // data zamkniecia: ustawiamy przy zamykaniu, czyscimy przy ponownym otwarciu
      if (zamkniete.checked && !t.zamkniete) {
        pola["Datazako_x0144_czenia"] = new Date().toISOString();
      } else if (!zamkniete.checked && t.zamkniete) {
        pola["Datazako_x0144_czenia"] = null;
      }

      const wynik = await updateTicketFields(t.id, pola, t.etag);
      t.opis = ta.value;
      t.czas = g;
      t.zamkniete = zamkniete.checked;
      t.etag = String((wynik as any)["@odata.etag"] ?? t.etag);
      poZapisie();
      komunikat.textContent = "Zapisano";
      komunikat.className = "crm-komunikat crm-ok";
    } catch (e: any) {
      komunikat.textContent = e?.message ?? String(e);
      komunikat.className = "crm-komunikat crm-blad";
    } finally {
      zapisz.disabled = false;
      ostrzezenie.style.display = "none";
    }
  };

  zapisz.onclick = async () => {
    const uwagi: string[] = [];

    // cudze zgloszenie - zapis zmienia czas i opis komus innemu
    let moje = true;
    try {
      const ja = (await getMojMail()).toLowerCase();
      moje = !t.opiekunMail || t.opiekunMail === ja;
    } catch {
      moje = true; // nie wiemy - nie strasz bez powodu
    }
    if (!moje) {
      uwagi.push(`To nie jest Twoje zgłoszenie — opiekunem jest ${t.opiekun || t.opiekunMail}.`);
    }

    const stary = t.opis.trim();
    const nowy = ta.value.trim();
    if (stary.length > 0 && !nowy.includes(stary)) {
      uwagi.push(nowy
        ? "Zmieniasz istniejący opis — poprzedni zniknie bez śladu."
        : "Kasujesz cały opis.");
    }

    if (uwagi.length) {
      ostrzezenie.innerHTML = "";
      uwagi.forEach((u) => ostrzezenie.appendChild(el("div", undefined, u)));
      const potwierdz = el("button", "crm-btn crm-btn-ostrz", "Zapisz mimo to") as HTMLButtonElement;
      const anuluj = el("button", "crm-btn crm-btn-cichy", "Anuluj") as HTMLButtonElement;
      potwierdz.onclick = () => void wykonajZapis();
      anuluj.onclick = () => { ostrzezenie.style.display = "none"; };
      const rzad = el("div", "crm-akcje");
      rzad.appendChild(potwierdz);
      rzad.appendChild(anuluj);
      ostrzezenie.appendChild(rzad);
      ostrzezenie.style.display = "block";
      return;
    }
    void wykonajZapis();
  };

  return box;
}

async function panelNowego() {
  const item = Office.context.mailbox.item as Office.MessageRead;
  if (!item) return;

  const temat = item.subject ?? "";
  if (ticketIdZTematu(temat) !== null) return; // watek istniejacego zgloszenia

  const nadawca = (item.from?.emailAddress ?? "").toLowerCase();
  if (!nadawca) return;

  const box = el("div", "crm-nowe");
  root().insertBefore(box, root().children[1] ?? null);

  let kontakt: any = null;
  try {
    kontakt = await kontaktDlaNadawcy(nadawca);
  } catch (e: any) {
    box.appendChild(el("div", "crm-nowe-info", `Nie sprawdzono kontaktu: ${e?.message ?? e}`));
    return;
  }

  if (!kontakt) {
    box.className = "crm-nowe crm-nowe-brak";
    await panelNowegoKontaktu(box, nadawca, item.from?.displayName ?? "");
    return;
  }

  const f = kontakt.fields as Record<string, string>;
  const osoba = f["Imi_x0119__x0020_i_x0020_nazwisk"]
    || [f["Imi_x0119_"], f["Nazwisko"]].filter(Boolean).join(" ");

  box.appendChild(el("div", "crm-nowe-info",
    `${osoba || nadawca} · ${f.FirmaID || "(brak firmy)"}`));

  // --- wybor opiekuna supportu ---
  box.appendChild(el("label", "crm-label", "Opiekun supportu"));
  const wybor = document.createElement("select");
  wybor.className = "crm-select";
  box.appendChild(wybor);

  let zalogowany = "";
  try {
    zalogowany = await getMojMail();
  } catch {
    zalogowany = (Office.context.mailbox.userProfile.emailAddress ?? "").toLowerCase();
  }

  let serwisanci: Awaited<ReturnType<typeof getSerwisanci>> = [];
  try {
    serwisanci = await getSerwisanci();
  } catch (e: any) {
    box.appendChild(el("div", "crm-nowe-info", `Nie pobrano listy serwisantów: ${e?.message ?? e}`));
  }

  serwisanci.forEach((s, idx) => {
    const opt = document.createElement("option");
    opt.value = String(idx);
    opt.textContent = s.nazwa;
    if (s.mail === zalogowany) opt.selected = true;
    wybor.appendChild(opt);
  });

  // zalogowanego nie ma w MEScoStaff - wymuszamy swiadomy wybor
  if (!serwisanci.some((s) => s.mail === zalogowany)) {
    const pusta = document.createElement("option");
    pusta.value = "";
    pusta.textContent = "— wybierz opiekuna —";
    pusta.selected = true;
    wybor.insertBefore(pusta, wybor.firstChild);
  }

  const btn = el("button", "crm-btn crm-btn-nowe", "+ Nowe zgłoszenie") as HTMLButtonElement;
  const komunikat = el("span", "crm-komunikat");
  const rzad = el("div", "crm-akcje");
  rzad.appendChild(btn);
  rzad.appendChild(komunikat);
  box.appendChild(rzad);

  // zgloszenie z wlasnej skrzynki omija przeplyw skrzynki supportowej
  let zSupportu = true;
  try {
    zSupportu = await czySkrzynkaSupportu(item);
  } catch {
    zSupportu = true; // nie wiemy - nie strasz bez powodu
  }

  const ostrzezenie = el("div", "crm-ostrzezenie");
  ostrzezenie.style.display = "none";
  box.appendChild(ostrzezenie);

  /** reczne = bez MailID i MailSubject, jak zgloszenie z kontaktu firmy */
  const zaloz = async (reczne = false) => {
    const wybrany = serwisanci[Number(wybor.value)];
    if (!wybrany) {
      komunikat.textContent = "Wybierz opiekuna supportu";
      komunikat.className = "crm-komunikat crm-blad";
      return;
    }

    ostrzezenie.style.display = "none";
    btn.disabled = true;
    wybor.disabled = true;
    komunikat.textContent = "Zakładanie...";
    komunikat.className = "crm-komunikat";

    const dzis = new Date();
    const dataISO = new Date(Date.UTC(
      dzis.getFullYear(), dzis.getMonth(), dzis.getDate(), 7, 0, 0
    )).toISOString();

    const pola: Record<string, unknown> = {
      Firma: f.FirmaID ?? "",
      Osoba: osoba,
      OsobaMail: nadawca,
      Data: dataISO,
      Status: false,
      IdentyfikatorContact: f.ContactGUID ?? "",
      IdentyfikatorCompanies: f.IdentyfikatorCompanies ?? "",
      MailCreator: wybrany.mail,
      Dodatkowy_opiekun: wybrany.nazwa,
    };
    // przeplyw rozpoznaje zgloszenie z maila po MailID - reczne idzie sciezka bez maili
    if (!reczne) {
      pola["MailSubject"] = temat;
      pola["MailID"] = mailIdRest(item);
    }

    if (wybrany.siteId !== null) {
      pola["Opiekun_supportuLookupId"] = String(wybrany.siteId);
    }

    try {
      await createTicket(pola);

      box.className = "crm-nowe crm-nowe-ok";
      box.innerHTML = "";
      box.appendChild(el("div", "crm-nowe-info",
        `${reczne ? "Zgłoszenie ręczne założone" : "Zgłoszenie założone"} dla ${wybrany.nazwa}. ` +
        "Numer nadawany przez przepływ — pojawi się w ciągu kilku minut."));
      const odsw = el("button", "crm-btn crm-btn-cichy", "Odśwież listę") as HTMLButtonElement;
      odsw.onclick = () => void pokazZgloszenia();
      box.appendChild(odsw);
    } catch (e: any) {
      komunikat.textContent = e?.message ?? String(e);
      komunikat.className = "crm-komunikat crm-blad";
      btn.disabled = false;
      wybor.disabled = false;
    }
  };

  btn.onclick = () => {
    if (zSupportu) return void zaloz();

    ostrzezenie.innerHTML = "";
    ostrzezenie.appendChild(el("div", undefined,
      `Dodajesz support ze swojej skrzynki dla ${osoba || nadawca}, czy jesteś pewny? Odbiorca nie dostanie ` +
      "powiadomienia o zgłoszeniu, a wątek będzie poza skrzynką supportową."));
    ostrzezenie.appendChild(el("div", undefined,
      "Możesz założyć zgłoszenie ręczne — tak jak z poziomu firmy, bez powiązania z tym mailem."));
    const potwierdz = el("button", "crm-btn crm-btn-ostrz", "Utwórz support") as HTMLButtonElement;
    const anuluj = el("button", "crm-btn crm-btn-cichy", "Anuluj") as HTMLButtonElement;
    potwierdz.onclick = () => void zaloz(true);
    anuluj.onclick = () => { ostrzezenie.style.display = "none"; };
    const rzadO = el("div", "crm-akcje");
    rzadO.appendChild(potwierdz);
    rzadO.appendChild(anuluj);
    ostrzezenie.appendChild(rzadO);
    ostrzezenie.style.display = "block";
  };
}
/** Nadawcy nie ma w Contacts - jesli jego firma jest w bazie, pozwalamy go dodac. */
async function panelNowegoKontaktu(box: HTMLElement, nadawca: string, nazwaNadawcy: string) {
  const domena = nadawca.split("@")[1] ?? "";
  box.appendChild(el("div", "crm-nowe-info", `${nadawca} nie figuruje w kontaktach.`));

  let firmy: CompanyFull[] = [];
  try {
    firmy = (await getAllCompanies())
      .map(toCompanyFull)
      .filter((f) => domena && f.domena === domena && f.guid);
  } catch (e: any) {
    box.appendChild(el("div", "crm-nowe-info", `Nie sprawdzono firm: ${e?.message ?? e}`));
    return;
  }

  if (!firmy.length) {
    box.appendChild(el("div", "crm-nowe-info",
      `Brak firmy z domeną ${domena}. Kontakt dodasz z widoku firmy: Firmy → firma → Kontakty.`));
    return;
  }

  const btn = el("button", "crm-btn crm-btn-nowe", "+ Dodaj kontakt") as HTMLButtonElement;
  box.appendChild(btn);

  btn.onclick = () => {
    btn.remove();
    const start = { ...rozbijNazwe(nazwaNadawcy), email: nadawca };
    const formularz = el("div");

    const rysuj = (f: CompanyFull) => {
      formularz.innerHTML = "";
      void formularzKontaktu(formularz, f, start, (nowy) => {
        kontaktCache.set(nadawca, nowy);
        void pokazZgloszenia();
      });
    };

    // kilka firm z ta sama domena (np. oddzialy) - swiadomy wybor
    if (firmy.length > 1) {
      box.appendChild(el("label", "crm-label", "Firma"));
      const wybor = document.createElement("select");
      wybor.className = "crm-select";
      firmy.forEach((f, idx) => {
        const opt = document.createElement("option");
        opt.value = String(idx);
        opt.textContent = f.skrocona || f.pelna;
        wybor.appendChild(opt);
      });
      wybor.onchange = () => rysuj(firmy[Number(wybor.value)]);
      box.appendChild(wybor);
    } else {
      box.appendChild(el("div", "crm-nowe-info", `Firma: ${firmy[0].skrocona || firmy[0].pelna}`));
    }

    box.appendChild(formularz);
    rysuj(firmy[0]);
  };
}

/** Nowy kontakt w firmie - odpowiednik formularza NewContact z Power Apps. */
async function formularzKontaktu(
  box: HTMLElement,
  firma: CompanyFull,
  start: { imie: string; nazwisko: string; email: string },
  poZapisie: (nowy: ListItem) => void,
  anuluj?: () => void
) {
  const pole = (etykieta: string, wartosc = "", typ = "text") => {
    box.appendChild(el("label", "crm-label", etykieta));
    const i = document.createElement("input");
    i.type = typ;
    i.className = "crm-input";
    i.value = wartosc;
    box.appendChild(i);
    return i;
  };

  const imie = pole("Imię", start.imie);
  const nazwisko = pole("Nazwisko", start.nazwisko);
  const email = pole("E-mail", start.email, "email");
  const telefon = pole("Telefon", "", "tel");
  const stanowisko = pole("Stanowisko");

  // --- plec (wymagana) i rola - opcje z definicji kolumn ---
  let plcie: string[] = [];
  let role: string[] = [];
  try {
    [plcie, role] = await Promise.all([
      getChoiceOptions("Contacts", "P_x0142_e_x0107_"),
      getChoiceOptions("Contacts", "RolawMEDDIC"),
    ]);
  } catch (e: any) {
    box.appendChild(el("div", "crm-nowe-info", `Nie pobrano opcji: ${e?.message ?? e}`));
  }

  let plec = "";
  let plecRecznie = false;
  const plecBox = el("div", "crm-chipy");
  const zaznaczPlec = (v: string) => {
    plec = v;
    plecBox.querySelectorAll(".crm-chip-wybor").forEach((x) =>
      x.classList.toggle("crm-chip-akt", x.textContent === v)
    );
  };
  if (plcie.length) {
    box.appendChild(el("label", "crm-label", "Płeć"));
    plcie.forEach((o) => {
      const b = el("button", "crm-chip-wybor", o) as HTMLButtonElement;
      b.onclick = () => { plecRecznie = true; zaznaczPlec(o); sprawdz(); };
      plecBox.appendChild(b);
    });
    box.appendChild(plecBox);
  }
  const podpowiedzPlci = () => {
    const p = plecZImienia(imie.value);
    if (!plecRecznie && plcie.includes(p)) zaznaczPlec(p);
  };
  podpowiedzPlci();

  const wybraneRole = new Set<string>();
  if (role.length) {
    box.appendChild(el("label", "crm-label", "Rola (MEDDIC)"));
    const rolaBox = el("div", "crm-chipy");
    role.forEach((o) => {
      const b = el("button", "crm-chip-wybor", o) as HTMLButtonElement;
      b.onclick = () => {
        if (wybraneRole.has(o)) wybraneRole.delete(o);
        else wybraneRole.add(o);
        b.classList.toggle("crm-chip-akt", wybraneRole.has(o));
        sprawdz();
      };
      rolaBox.appendChild(b);
    });
    box.appendChild(rolaBox);
  }

  // --- wysylka Messengera - domyslnie obie wylaczone ---
  const przelacznik = (etykieta: string) => {
    const c = document.createElement("input");
    c.type = "checkbox";
    c.id = `kt-${nowyGuid()}`;
    const lbl = el("label", "crm-check") as HTMLLabelElement;
    lbl.htmlFor = c.id;
    lbl.appendChild(c);
    lbl.appendChild(el("span", undefined, etykieta));
    box.appendChild(lbl);
    return c;
  };
  const stala = przelacznik("Stała wysyłka Messengera");
  const tymczasowa = przelacznik("Tymczasowa wysyłka Messengera");

  box.appendChild(el("label", "crm-label", "Uwagi"));
  const uwagi = document.createElement("textarea");
  uwagi.className = "crm-textarea";
  uwagi.rows = 2;
  box.appendChild(uwagi);

  const brakujace = el("div", "crm-podpowiedz");
  box.appendChild(brakujace);

  const zapisz = el("button", "crm-btn", "Zapisz kontakt") as HTMLButtonElement;
  const komunikat = el("span", "crm-komunikat");
  const rzad = el("div", "crm-akcje");
  rzad.appendChild(zapisz);
  if (anuluj) {
    const a = el("button", "crm-btn crm-btn-cichy", "Anuluj") as HTMLButtonElement;
    a.onclick = anuluj;
    rzad.appendChild(a);
  }
  rzad.appendChild(komunikat);
  box.appendChild(rzad);

  function sprawdz(): boolean {
    const brak: string[] = [];
    if (!imie.value.trim() && !nazwisko.value.trim()) brak.push("imię lub nazwisko");
    if (plcie.length && !plec) brak.push("płeć");
    if (role.length && !wybraneRole.size) brak.push("rola");
    const ok = brak.length === 0;
    zapisz.disabled = !ok;
    brakujace.textContent = ok ? "" : `Uzupełnij: ${brak.join(", ")}.`;
    brakujace.style.display = ok ? "none" : "block";
    return ok;
  }
  imie.oninput = () => { podpowiedzPlci(); sprawdz(); };
  nazwisko.oninput = () => sprawdz();
  sprawdz();

  zapisz.onclick = async () => {
    if (!sprawdz()) return;
    const mail = email.value.trim().toLowerCase();

    zapisz.disabled = true;
    komunikat.textContent = "Zapisywanie...";
    komunikat.className = "crm-komunikat";

    try {
      // ten sam adres w bazie = duplikat, a dopasowanie nadawcy bierze pierwszy
      if (mail && (await getContactByEmail(mail))) {
        throw new Error(`Kontakt z adresem ${mail} już jest w bazie.`);
      }

      const pola: Record<string, unknown> = {
        ContactGUID: nowyGuid(),
        IdentyfikatorCompanies: firma.guid,
        Imi_x0119_: imie.value.trim(),
        Nazwisko: nazwisko.value.trim(),
        Email: mail,
        Phone: telefon.value.trim(),
        Stanowisko: stanowisko.value.trim(),
        UWAGI: uwagi.value.trim(),
        FirmaID: firma.skrocona,
        nazwafirmyKontakt: firma.pelna || firma.skrocona,
        ulicakontakt: firma.ulica,
        miastoKontakt: firma.miasto,
        kodpocztowykontakt: firma.kod,
        Sta_x0142_awysy_x0142_ka: stala.checked,
        Tymczasowawysy_x0142_ka: tymczasowa.checked,
      };
      if (plec) pola["P_x0142_e_x0107_"] = plec;
      if (wybraneRole.size) {
        pola["RolawMEDDIC@odata.type"] = "Collection(Edm.String)";
        pola["RolawMEDDIC"] = [...wybraneRole];
      }

      const nowy = await createListItem("Contacts", pola);
      komunikat.textContent = "Kontakt zapisany";
      komunikat.className = "crm-komunikat crm-ok";
      poZapisie(nowy);
    } catch (e: any) {
      komunikat.textContent = e?.message ?? String(e);
      komunikat.className = "crm-komunikat crm-blad";
      zapisz.disabled = false;
    }
  };
}

async function panelZgloszenia() {
  const item = Office.context.mailbox.item as Office.MessageRead;
  if (!item) return;
 
  const nr = ticketIdZTematu(item.subject ?? "");
  if (nr === null) return;
 
  const box = el("div", "crm-watek");
  root().insertBefore(box, root().children[1] ?? null);
  box.appendChild(el("div", "crm-watek-info", `Wczytywanie MESco-ID-${nr}...`));
 
  let t: Ticket;
  try {
    const rekord = await getTicketByNumber(nr);
    if (!rekord) {
      box.innerHTML = "";
      box.className = "crm-watek crm-watek-brak";
      box.appendChild(el("div", "crm-watek-info",
        `Temat wskazuje na MESco-ID-${nr}, ale takiego zgłoszenia nie ma w bazie.`));
      return;
    }
    t = toTicket(rekord);
  } catch (e: any) {
    box.innerHTML = "";
    box.appendChild(el("div", "crm-watek-info", `Nie wczytano zgłoszenia: ${e?.message ?? e}`));
    return;
  }
 
  box.innerHTML = "";
 
  const top = el("div", "crm-watek-top");
  top.appendChild(el("span", "crm-watek-nr", `MESco-ID-${t.ticketId}`));
  top.appendChild(el("span", "crm-watek-firma", t.firma || ""));
  box.appendChild(top);
 
  // --- czyje to zgloszenie ---
  let moje = true;
  try {
    const ja = (await getMojMail()).toLowerCase();
    moje = !t.opiekunMail || t.opiekunMail.toLowerCase() === ja;
  } catch {
    moje = true; // nie wiemy - nie strasz bez powodu
  }
 
  if (!moje) {
    box.appendChild(el("div", "crm-watek-cudze",
      `Opiekunem jest ${t.opiekun || t.opiekunMail}. Odpowiadasz na cudze zgłoszenie.`));
  }
 
  if (t.zamkniete) {
    box.appendChild(el("div", "crm-watek-info", "Zgłoszenie jest już zamknięte."));
    return;
  }
 
  // --- szybkie zamkniecie ---
  box.appendChild(el("label", "crm-label", "Czas (godziny)"));
  const czas = document.createElement("input");
  czas.type = "text";
  czas.inputMode = "decimal";
  czas.className = "crm-input";
  czas.placeholder = "np. 1,5";
  czas.value = t.czas ? String(t.czas) : "";
  box.appendChild(czas);
 
  box.appendChild(el("label", "crm-label", "Opis"));
  const ta = document.createElement("textarea");
  ta.className = "crm-textarea";
  ta.rows = 3;
  ta.value = t.opis;
  box.appendChild(ta);
 
  const podpowiedz = el("div", "crm-podpowiedz");
  box.appendChild(podpowiedz);
 
  const zamknij = el("button", "crm-btn", "Zamknij zgłoszenie") as HTMLButtonElement;
  const komunikat = el("span", "crm-komunikat");
  const rzad = el("div", "crm-akcje");
  rzad.appendChild(zamknij);
  rzad.appendChild(komunikat);
  box.appendChild(rzad);
 
  const sprawdz = () => {
    const g = parseGodziny(czas.value);
    const brak: string[] = [];
    if (!(g > 0)) brak.push("czas");
    if (!ta.value.trim()) brak.push("opis");
    const ok = brak.length === 0;
    zamknij.disabled = !ok;
    podpowiedz.textContent = ok ? "" : `Do zamknięcia brakuje: ${brak.join(" i ")}.`;
    podpowiedz.style.display = ok ? "none" : "block";
    return ok;
  };
  czas.oninput = sprawdz;
  ta.oninput = sprawdz;
  sprawdz();
 
  zamknij.onclick = async () => {
    if (!sprawdz()) return;
    const g = parseGodziny(czas.value);
    if (Number.isNaN(g)) {
      komunikat.textContent = "Czas musi być liczbą (np. 1,5)";
      komunikat.className = "crm-komunikat crm-blad";
      return;
    }
 
    zamknij.disabled = true;
    komunikat.textContent = "Zamykanie...";
    komunikat.className = "crm-komunikat";
 
    try {
      await updateTicketFields(
        t.id,
        {
          Opis: ta.value,
          Czas: g,
          Status: true,
          Datazako_x0144_czenia: new Date().toISOString(),
        },
        t.etag
      );
      box.className = "crm-watek crm-watek-ok";
      box.innerHTML = "";
      box.appendChild(el("div", "crm-watek-info",
        `MESco-ID-${t.ticketId} zamknięte · ${g} h`));
      const odsw = el("button", "crm-btn crm-btn-cichy", "Odśwież listę") as HTMLButtonElement;
      odsw.onclick = () => void pokazZgloszenia();
      box.appendChild(odsw);
    } catch (e: any) {
      komunikat.textContent = e?.message ?? String(e);
      komunikat.className = "crm-komunikat crm-blad";
      zamknij.disabled = false;
    }
  };
}
async function pokazZgloszenia() {
  widok = "zgloszenia";
  ustawNaglowek("Support", "nav-tickets");
  root().innerHTML = "";

  const item = Office.context.mailbox.item as Office.MessageRead;
  const nadawca = (item?.from?.emailAddress ?? "").toLowerCase();

  // --- pasek filtrow ---
  const filtry = el("div", "crm-filtry");

  // numer zgloszenia (poczatek numeru) - szuka w calej liscie, pomija pozostale filtry
  const numer = document.createElement("input");
  numer.type = "search";
  numer.className = "crm-szukaj";
  numer.placeholder = "nr zgłoszenia, np. 1234 lub MESco-ID-1234";
  numer.inputMode = "numeric";
  filtry.appendChild(numer);

  const zakres = document.createElement("select");
  zakres.className = "crm-select";
  [
    { v: "mail", t: nadawca ? `Osoba: ${nadawca}` : "Z tej wiadomości" },
    { v: "firma", t: "Firma z tej wiadomości" },
    { v: "wszystkie", t: "Wszystkie zgłoszenia" },
  ].forEach((o) => {
    const opt = document.createElement("option");
    opt.value = o.v;
    opt.textContent = o.t;
    if (o.v === zakresZgloszen) opt.selected = true;
    zakres.appendChild(opt);
  });
  if (!nadawca) {
    zakres.value = "wszystkie";
    zakresZgloszen = "wszystkie";
    (zakres.querySelector('option[value="mail"]') as HTMLOptionElement).disabled = true;
    (zakres.querySelector('option[value="firma"]') as HTMLOptionElement).disabled = true;
  }
  filtry.appendChild(zakres);

  const wybor = document.createElement("select");
  wybor.className = "crm-select";
  const pusta = document.createElement("option");
  pusta.value = "";
  pusta.textContent = "wszyscy opiekunowie";
  wybor.appendChild(pusta);
  filtry.appendChild(wybor);

  root().appendChild(filtry);

  void panelNowego();
  void panelZgloszenia();

  const s = status("Pobieranie zgłoszeń...");
  const lista = el("div");
  root().appendChild(lista);

  void (async () => {
    try {
      const osoby = await getSerwisanci();
      osoby.forEach((o) => {
        const opt = document.createElement("option");
        opt.value = o.mail;
        opt.textContent = o.nazwa;
        if (o.mail === opiekunZgloszen) opt.selected = true;
        wybor.appendChild(opt);
      });
    } catch (e: any) {
      pusta.textContent = `opiekunowie: ${e?.message ?? e}`;
    }
  })();

  let pokolenie = 0; // starsza odpowiedz nie nadpisze nowszej przy szybkim pisaniu
  const wczytaj = async () => {
    const moje = ++pokolenie;
    lista.innerHTML = "";
    s.textContent = "Pobieranie zgłoszeń...";
    s.classList.remove("crm-error");

    try {
      const nr = numer.value.replace(/\D/g, "").replace(/^0+/, "");
      if (nr) {
        const items = await getTicketsFiltered({ ticketPrefix: nr, top: 50 });
        if (moje !== pokolenie) return;
        s.textContent = `Numer zaczyna się od ${nr} — ${items.length}`;
        if (!items.length) lista.appendChild(el("div", "crm-pusto", `Brak zgłoszeń z numerem od ${nr}.`));
        items.map(toTicket).forEach((t) => lista.appendChild(karta(t)));
        return;
      }

      const opis: string[] = [];
      let osobaMail = "";
      let companyGuid = "";

      if (zakresZgloszen === "mail" && nadawca) {
        osobaMail = nadawca;
        opis.push(nadawca);
      }

      if (zakresZgloszen === "firma" && nadawca) {
        const kontakt = await kontaktDlaNadawcy(nadawca);
        const f = kontakt?.fields as Record<string, string> | undefined;
        companyGuid = f?.IdentyfikatorCompanies ?? "";

        if (!companyGuid) {
          s.textContent = `${nadawca} nie ma przypisanej firmy w kontaktach.`;
          s.classList.add("crm-error");
          return;
        }
        opis.push(`firma: ${f?.FirmaID || companyGuid}`);
      }

      if (opiekunZgloszen) {
        const nazwa = wybor.options[wybor.selectedIndex]?.textContent ?? opiekunZgloszen;
        opis.push(`opiekun: ${nazwa}`);
      }

      const items = await getTicketsFiltered({
        osobaMail,
        companyGuid,
        mailCreator: opiekunZgloszen,
        top: 50,
      });
      if (moje !== pokolenie) return;

      s.textContent = opis.length
        ? `${opis.join(" · ")} — ${items.length}`
        : `Najnowsze zgłoszenia (${items.length})`;

      if (!items.length) lista.appendChild(el("div", "crm-pusto", "Brak zgłoszeń dla tych kryteriów."));
      items.map(toTicket).forEach((t) => lista.appendChild(karta(t)));
    } catch (e: any) {
      console.error(e);
      s.textContent = `Błąd: ${e?.message ?? e}`;
      s.classList.add("crm-error");
    }
  };

  let timer: number | undefined;
  numer.oninput = () => {
    clearTimeout(timer);
    timer = window.setTimeout(() => void wczytaj(), 300);
  };

  zakres.onchange = () => {
    zakresZgloszen = zakres.value as "mail" | "firma" | "wszystkie";
    void wczytaj();
  };
  wybor.onchange = () => {
    opiekunZgloszen = wybor.value;
    void wczytaj();
  };

  await wczytaj();
}


// ---------------------------------------------------------------- galeria firm

/** 1 support, 2-4 supporty, 5+ supportow (12-14 tez supportow) */
function supportyOdmiana(n: number): string {
  if (n === 1) return "1 support";
  const d = n % 10, s = n % 100;
  return `${n} ${d >= 2 && d <= 4 && (s < 12 || s > 14) ? "supporty" : "supportów"}`;
}

function kartaFirmy(c: Company, liczby?: Promise<Map<string, number>>): HTMLElement {
  const card = el("div", "crm-firma-card");
  card.tabIndex = 0;
  card.setAttribute("role", "button");

  const linia = (klasa: string, nazwa: keyof typeof IKONY, tekst: string, tytul: string) => {
    const w = el("div", klasa);
    w.title = tytul;
    w.append(ikona(nazwa), el("span", "crm-ikona-txt", tekst));
    return w;
  };

  const top = el("div", "crm-firma-top");
  const nazwa = el("span", "crm-firma-nazwa");
  nazwa.append(ikona("firma"), el("span", "crm-ikona-txt", c.nazwa || "(bez nazwy)"));
  top.appendChild(nazwa);
  top.appendChild(el("span", "crm-tecs", `TECS: ${c.tecs}`));
  card.appendChild(top);

  const adres = [c.ulica, [c.kod, c.miasto].filter(Boolean).join(" ")].filter(Boolean);
  if (adres.length) card.appendChild(linia("crm-adres", "pinezka", adres.join("\n"), "Adres"));

  if (c.opiekun) card.appendChild(linia("crm-firma-opiekun", "opiekun", c.opiekun, "Opiekun"));

  // liczba supportow z ostatnich 12 miesiecy - dociagana po narysowaniu karty
  if (liczby) {
    const stat = linia("crm-firma-stat", "klucz", "…", "Supporty z ostatnich 12 miesięcy");
    const txt = stat.querySelector(".crm-ikona-txt")!;
    card.appendChild(stat);
    liczby.then(
      (m) => {
        const n = m.get(c.guid) ?? 0;
        txt.textContent = n ? `${supportyOdmiana(n)} w ostatnim roku` : "brak supportów w ostatnim roku";
        stat.classList.toggle("crm-firma-stat-zero", !n);
      },
      () => stat.remove()
    );
  }

  const otworz = () => {
    zakladkaFirmy = "Kontakty";
    void pokazFirma(c);
  };
  card.onclick = otworz;
  card.onkeydown = (e: KeyboardEvent) => {
    if (e.key === "Enter" || e.key === " ") { e.preventDefault(); otworz(); }
  };

  return card;
}

async function pokazFirmy(fraza = "", opiekunWybrany = "") {
  widok = "firmy";
  ustawNaglowek("Firmy", "nav-companies");
  root().innerHTML = "";

  const filtry = el("div", "crm-filtry");

  const szukaj = document.createElement("input");
  szukaj.type = "search";
  szukaj.className = "crm-szukaj";
  szukaj.placeholder = "szukaj po nazwie firmy";
  szukaj.value = fraza;
  filtry.appendChild(szukaj);

  const wybor = document.createElement("select");
  wybor.className = "crm-select";
  const pusta = document.createElement("option");
  pusta.value = "";
  pusta.textContent = "wszyscy opiekunowie";
  wybor.appendChild(pusta);
  filtry.appendChild(wybor);

  const nowa = przyciskIk("crm-btn crm-nf-otworz", "firma", "Nowa firma");
  nowa.onclick = () => {
    root().innerHTML = "";
    ustawNaglowek("Nowa firma", "nav-companies");
    void formularzNowejFirmy(root(), () => void pokazFirmy(szukaj.value, wybor.value));
  };
  filtry.appendChild(nowa);

  root().appendChild(filtry);

  const lista = el("div");
  const s = status("Pobieranie firm...");
  root().appendChild(lista);

  void (async () => {
    try {
      const osoby = await getOpiekunowie();
      osoby.forEach((o) => {
        const opt = document.createElement("option");
        opt.value = o;
        opt.textContent = o;
        if (o === opiekunWybrany) opt.selected = true;
        wybor.appendChild(opt);
      });
    } catch (e: any) {
      pusta.textContent = `opiekunowie: ${e?.message ?? e}`;
    }
  })();

  const wczytaj = async (q: string, o: string) => {
    lista.innerHTML = "";
    s.textContent = "Pobieranie firm...";
    s.classList.remove("crm-error");
    try {
      const firmy = (await getCompanies2(q, o)).map(toCompany);

      const opis: string[] = [];
      if (q) opis.push(`nazwa: ${q}`);
      if (o) opis.push(`opiekun: ${o}`);
      s.textContent = opis.length
        ? `${opis.join(" · ")} — ${firmy.length}`
        : `Największy TECS (${firmy.length})`;

      if (!firmy.length) lista.appendChild(el("div", "crm-pusto", "Brak firm dla tych kryteriów."));
      const liczby = getLiczbaSupportowRok();
      liczby.catch((e) => console.error("liczba supportow:", e));
      firmy.forEach((c) => lista.appendChild(kartaFirmy(c, liczby)));
    } catch (e: any) {
      console.error(e);
      s.textContent = `Błąd: ${e?.message ?? e}`;
      s.classList.add("crm-error");
    }
  };

  let timer: any;
  szukaj.oninput = () => {
    clearTimeout(timer);
    timer = setTimeout(() => void wczytaj(szukaj.value, wybor.value), 350);
  };
  wybor.onchange = () => void wczytaj(szukaj.value, wybor.value);

  await wczytaj(fraza, opiekunWybrany);
}

// ---------------------------------------------------------------- nowa firma

const RODZAJE_PROSPECTU = ["Prospect", "Cold", "Warm", "Hot", "Marketing", "Zamknięty"];

// domyslne Źródło jak w Power Apps - po zalogowanym handlowcu
const ZRODLO_HANDLOWCA: Record<string, string> = {
  "Mateusz Nowakowski": "MNow",
  "Przemysław Draga": "PDra",
  "Grzegorz Woźniak": "GWoz",
  "Jacek Maj": "JMaj",
  "Łukasz Borgul": "ŁBor",
};

/** Nowa firma w Companies - dane po NIP z rejestru VAT, reszta recznie. Odpowiednik conNowaFirma z Power Apps. */
async function formularzNowejFirmy(miejsce: HTMLElement, anuluj: () => void) {
  const box = el("div", "crm-nowe crm-nf");
  miejsce.appendChild(box);
  box.appendChild(poleIk("crm-demo-tytul", "firma", "Nowa firma"));
  box.appendChild(el("div", "crm-nowe-info", "Wpisz NIP i pobierz dane z rejestru VAT – resztę pól uzupełnisz ręcznie."));

  const sekcja = (t: string) => box.appendChild(el("div", "crm-nf-sekcja", t));
  const select = (opcje: string[], start = "") => {
    const s = document.createElement("select");
    s.className = "crm-select";
    opcje.forEach((o) => {
      const opt = document.createElement("option");
      opt.value = o;
      opt.textContent = o || "wybierz…";
      opt.selected = o === start;
      s.appendChild(opt);
    });
    box.appendChild(s);
    return s;
  };

  // --- identyfikacja ---
  sekcja("Identyfikacja");
  const oddzial = document.createElement("input");
  oddzial.type = "checkbox";
  oddzial.id = `nf-oddzial-${nowyGuid()}`;
  const oddzialLbl = el("label", "crm-check") as HTMLLabelElement;
  oddzialLbl.htmlFor = oddzial.id;
  oddzialLbl.append(oddzial, el("span", undefined, "Oddział innej firmy"));
  box.appendChild(oddzialLbl);

  const nipLbl = el("label", "crm-label", "NIP");
  box.appendChild(nipLbl);
  const nipRzad = el("div", "crm-nf-nip");
  const nip = document.createElement("input");
  nip.type = "text";
  nip.inputMode = "numeric";
  nip.className = "crm-input";
  nip.placeholder = "10 cyfr, bez kresek";
  const pobierz = el("button", "crm-btn crm-btn-cichy", "Pobierz dane") as HTMLButtonElement;
  pobierz.title = "Nazwa i adres z Wykazu podatników VAT (Ministerstwo Finansów)";
  nipRzad.append(nip, pobierz);
  box.appendChild(nipRzad);
  const nipInfo = el("div", "crm-komunikat crm-nf-info");
  box.appendChild(nipInfo);

  const domena = poleFormularza(box, "Domena lub e-mail *");
  domena.placeholder = "np. firma.pl";
  const domenaInfo = el("div", "crm-komunikat crm-nf-info");
  box.appendChild(domenaInfo);
  const pelna = poleFormularza(box, "Pełna nazwa");
  pelna.placeholder = "z rejestru lub wpisz ręcznie";
  const skrocona = poleFormularza(box, "Skrócona nazwa");
  skrocona.placeholder = "np. SECO";

  // --- adres ---
  sekcja("Adres");
  const ulica = poleFormularza(box, "Ulica i numer");
  const kodMiasto = el("div", "crm-nf-kod");
  const kolumna = (etykieta: string, ph = "") => {
    const k = el("div");
    k.appendChild(el("label", "crm-label", etykieta));
    const i = document.createElement("input");
    i.type = "text";
    i.className = "crm-input";
    i.placeholder = ph;
    k.appendChild(i);
    kodMiasto.appendChild(k);
    return i;
  };
  const kod = kolumna("Kod pocztowy", "00-000");
  const miasto = kolumna("Miasto");
  box.appendChild(kodMiasto);
  const nda = poleFormularza(box, "NDA ważne do", "date");

  // --- klasyfikacja - opcje z definicji kolumn Companies ---
  sekcja("Klasyfikacja");
  let statusy: string[] = [];
  let branze: string[] = [];
  let zrodla: Awaited<ReturnType<typeof getOpcjeLookupu>> = null;
  let opiekunowie: string[] = [];
  let ja = "";
  let firmy: ListItem[] = [];
  try {
    [statusy, branze, zrodla, opiekunowie, ja, firmy] = await Promise.all([
      getChoiceOptions("Companies", "Status"),
      getChoiceOptions("Companies", "Bran_x017c_a"),
      getOpcjeLookupu("Companies", "Źródło"),
      getOpiekunowie().catch(() => [] as string[]),
      getMojaNazwa().catch(() => ""),
      getAllCompanies(),
    ]);
  } catch (e: any) {
    box.appendChild(el("div", "crm-nowe-info crm-blad", `Nie pobrano opcji: ${e?.message ?? e}`));
  }

  box.appendChild(el("label", "crm-label", "Status *"));
  const status = select(["", ...statusy], statusy.includes("Prospect") ? "Prospect" : "");
  const rodzajBox = el("div");
  box.appendChild(rodzajBox);
  rodzajBox.appendChild(el("label", "crm-label", "Rodzaj prospectu"));
  const rodzaj = select(RODZAJE_PROSPECTU, "Prospect");
  rodzajBox.appendChild(rodzaj);
  const pokazRodzaj = () => (rodzajBox.style.display = status.value === "Prospect" ? "" : "none");
  pokazRodzaj();

  box.appendChild(el("label", "crm-label", "Branża *"));
  const wybraneBranze = new Set<string>();
  const branzaBox = el("div", "crm-chipy");
  branze.forEach((o) => {
    const b = el("button", "crm-chip-wybor", o) as HTMLButtonElement;
    b.onclick = () => {
      if (wybraneBranze.has(o)) wybraneBranze.delete(o);
      else wybraneBranze.add(o);
      b.classList.toggle("crm-chip-akt", wybraneBranze.has(o));
      sprawdz();
    };
    branzaBox.appendChild(b);
  });
  box.appendChild(branzaBox);

  box.appendChild(el("label", "crm-label", "Opiekun"));
  if (ja && !opiekunowie.includes(ja)) opiekunowie = [ja, ...opiekunowie];
  const opiekun = select(["", ...opiekunowie], ja);

  box.appendChild(el("label", "crm-label", "Źródło"));
  // Źródło to lookup - pokazujemy etykiety, zapisujemy Id pozycji
  const etykietyZrodel = (zrodla?.opcje ?? []).map((o) => o.etykieta);
  const zrodloStart = ZRODLO_HANDLOWCA[ja] ?? "TNV";
  const zrodlo = select(["", ...etykietyZrodel], etykietyZrodel.includes(zrodloStart) ? zrodloStart : "");
  if (!zrodla) zrodlo.title = "Nie znaleziono kolumny Źródło w Companies";

  // klucz pola w Graph: kolumny zaczynajace sie od "_x" bywaja z prefiksem OData_ - bierzemy ten z danych
  const kluczZrodla = (() => {
    if (!zrodla) return null;
    const kandydaci = [`${zrodla.kolumna}LookupId`, `OData_${zrodla.kolumna}LookupId`];
    for (const i of firmy) {
      const hit = kandydaci.find((k) => k in (i.fields as any));
      if (hit) return hit;
    }
    return kandydaci[0];
  })();

  box.appendChild(el("label", "crm-label", "Uwagi"));
  const uwagi = document.createElement("textarea");
  uwagi.className = "crm-textarea";
  uwagi.rows = 3;
  uwagi.placeholder = "Dodatkowe informacje o firmie";
  box.appendChild(uwagi);

  // --- duplikaty: NIP blokuje zapis, domena tylko ostrzega ---
  const poNipie = new Map<string, string>();
  const poDomenie = new Map<string, string>();
  firmy.forEach((i) => {
    const f = i.fields as any;
    const nazwa = String(f.Nazwa || f.Pe_x0142_nanazwa || "");
    const v = String(f.VAT ?? "").replace(/\D/g, "");
    if (v) poNipie.set(v, nazwa);
    const d = String(f.DomenA ?? "").trim().toLowerCase();
    if (d) poDomenie.set(d, nazwa);
  });
  const samaDomena = () => (domena.value.trim().toLowerCase().split("@").pop() ?? "").trim();

  // --- przyciski ---
  const brakujace = el("div", "crm-podpowiedz");
  box.appendChild(brakujace);
  const zapisz = el("button", "crm-btn", "Zapisz firmę") as HTMLButtonElement;
  const wstecz = el("button", "crm-btn crm-btn-cichy", "Anuluj") as HTMLButtonElement;
  const komunikat = el("span", "crm-komunikat");
  const rzad = el("div", "crm-akcje");
  rzad.append(zapisz, wstecz, komunikat);
  box.appendChild(rzad);
  wstecz.onclick = anuluj;

  function sprawdz(): boolean {
    const n = nip.value;
    const pelny = /^\d{10}$/.test(n);
    const dubel = !oddzial.checked && pelny ? poNipie.get(n) : undefined;

    nipLbl.className = "crm-label";
    if (oddzial.checked) {
      nipLbl.textContent = "NIP jednostki nadrzędnej";
    } else if (!n) {
      nipLbl.textContent = "NIP";
    } else if (!pelny) {
      nipLbl.textContent = "NIP · niepełny numer";
      nipLbl.classList.add("crm-nf-uwaga");
    } else if (dubel !== undefined) {
      nipLbl.textContent = `NIP · firma już jest w bazie: ${dubel || "(bez nazwy)"}`;
      nipLbl.classList.add("crm-blad");
    } else if (!nipPoprawny(n)) {
      nipLbl.textContent = "NIP · błędna suma kontrolna, sprawdź numer";
      nipLbl.classList.add("crm-nf-uwaga");
    } else {
      nipLbl.textContent = "NIP · poprawny";
      nipLbl.classList.add("crm-ok");
    }
    nipRzad.classList.toggle("crm-nf-oddzial", oddzial.checked);
    pobierz.disabled = !pelny;

    const d = samaDomena();
    const zajeta = d ? poDomenie.get(d) : undefined;
    domenaInfo.textContent = zajeta !== undefined ? `Ta domena jest już przypisana do: ${zajeta || "(bez nazwy)"}` : "";
    domenaInfo.className = zajeta !== undefined ? "crm-komunikat crm-nf-info crm-nf-uwaga" : "crm-komunikat crm-nf-info";

    const brak: string[] = [];
    if (!pelny || dubel !== undefined) brak.push("NIP");
    if (!status.value) brak.push("status");
    if (branze.length && !wybraneBranze.size) brak.push("branża");
    if (!domena.value.trim()) brak.push("domena");
    const ok = brak.length === 0;
    zapisz.disabled = !ok;
    brakujace.textContent = ok ? "" : `Uzupełnij: ${brak.join(", ")}.`;
    brakujace.style.display = ok ? "none" : "block";
    return ok;
  }

  // wklejony "PL 123-456-78-90" tez przejdzie
  nip.oninput = () => {
    nip.value = nip.value.replace(/\D/g, "").slice(0, 10);
    sprawdz();
  };
  oddzial.onchange = () => {
    nipInfo.textContent = "";
    sprawdz();
  };
  domena.oninput = () => sprawdz();
  status.onchange = () => { pokazRodzaj(); sprawdz(); };
  sprawdz();

  pobierz.onclick = async () => {
    pobierz.disabled = true;
    nipInfo.textContent = "Pobieranie z rejestru VAT…";
    nipInfo.className = "crm-komunikat crm-nf-info";
    try {
      const d = await pobierzPoNip(nip.value);
      if (!d) {
        nipInfo.textContent = `Brak NIP ${nip.value} w wykazie podatników VAT – uzupełnij dane ręcznie.`;
        nipInfo.classList.add("crm-nf-uwaga");
      } else {
        pelna.value = pelnaNazwa(d.nazwa);
        ulica.value = d.ulica;
        kod.value = d.kod;
        miasto.value = d.miasto;
        const vat = d.statusVat && d.statusVat !== "Czynny" ? ` Status VAT: ${d.statusVat}.` : "";
        nipInfo.textContent = `Pobrano dane z rejestru VAT.${vat}`;
        nipInfo.classList.add(vat ? "crm-nf-uwaga" : "crm-ok");
      }
    } catch (e: any) {
      nipInfo.textContent = e?.message ?? String(e);
      nipInfo.classList.add("crm-blad");
    }
    sprawdz();
  };

  zapisz.onclick = async () => {
    if (!sprawdz()) return;
    zapisz.disabled = true;
    komunikat.textContent = "Zapisywanie...";
    komunikat.className = "crm-komunikat";

    try {
      // pole Osoba wymaga Id uzytkownika witryny - szukamy maila wybranego opiekuna
      let opiekunId: number | null = null;
      if (opiekun.value) {
        const mail = opiekun.value === ja
          ? await getMojMail()
          : (await getSerwisanci().catch(() => [])).find((s) => s.nazwa === opiekun.value)?.mail ?? "";
        if (mail) opiekunId = await getSiteUserId(mail);
      }

      const guid = nowyGuid();
      const pola: Record<string, unknown> = {
        CompaniesGUID: guid,
        Pe_x0142_nanazwa: pelna.value.trim(),
        Nazwa: skrocona.value.trim(),
        Oddzia_x0142_: oddzial.checked,
        VAT: oddzial.checked ? "" : nip.value,
        Ulica: ulica.value.trim(),
        Kodpocztowy: kod.value.trim(),
        Miasto: miasto.value.trim(),
        Status: status.value,
        Status1: status.value,
        Uwagi: uwagi.value.trim(),
        Domen_x0040_: domena.value.trim(),
        DomenA: samaDomena(),
        Tw_x00f3_rcaFirmy: ja,
        Creator: opiekun.value,
      };
      if (oddzial.checked) pola["NIPjednostkiNadrzednej"] = Number(nip.value);
      if (status.value === "Prospect") pola["ProspectStatus"] = rodzaj.value;
      if (nda.value) pola["DataNDA"] = dzienSP(nda.value);
      if (wybraneBranze.size) {
        pola["Bran_x017c_a@odata.type"] = "Collection(Edm.String)";
        pola["Bran_x017c_a"] = [...wybraneBranze];
      }
      const idZrodla = zrodla?.opcje.find((o) => o.etykieta === zrodlo.value)?.id;
      if (kluczZrodla && idZrodla) pola[kluczZrodla] = idZrodla;
      if (opiekunId !== null) pola["Utworzoneprzez_x002f_OpiekunLookupId"] = String(opiekunId);

      const nowy = await createListItem("Companies", pola);
      wyczyscCacheFirm();

      const c: Company = {
        id: String(nowy.id ?? (nowy.fields as any)?.id ?? ""),
        guid,
        nazwa: skrocona.value.trim() || pelna.value.trim(),
        ulica: ulica.value.trim(),
        kod: kod.value.trim(),
        miasto: miasto.value.trim(),
        tecs: 0,
        opiekun: opiekun.value,
      };
      zakladkaFirmy = "Kontakty";
      await pokazFirma(c);
      if (opiekun.value && opiekunId === null) {
        console.warn("Nowa firma bez pola Opiekun (osoba) - nie znaleziono użytkownika witryny:", opiekun.value);
      }
    } catch (e: any) {
      komunikat.textContent = e?.message ?? String(e);
      komunikat.className = "crm-komunikat crm-blad";
      zapisz.disabled = false;
    }
  };
}

// ---------------------------------------------------------------- kontakty firmy

function kartaKontaktu(k: Contact, c: Company): HTMLElement {
  const card = el("div", "crm-kontakt-card");

  const top = el("div", "crm-firma-top");
  top.appendChild(el("span", "crm-firma-nazwa", k.imieNazwisko || "(bez nazwiska)"));
  if (!k.aktywny) top.appendChild(el("span", "crm-nieaktywny", "nie pracuje"));
  card.appendChild(top);

  if (k.stanowisko) card.appendChild(el("div", "crm-adres", k.stanowisko));

  if (k.email) {
    const mail = el("div", "crm-kontakt-linia");
    mail.append(ikona("mail"), el("span", undefined, k.email));
    zrobKopiowalny(mail, k.email, "skopiowano adres");
    card.appendChild(mail);
  }

  if (k.telefon) {
    const tel = document.createElement("a");
    tel.className = "crm-kontakt-linia crm-tel";
    tel.href = `tel:${k.telefon.replace(/[^\d+]/g, "")}`;
    tel.append(ikona("telefon"), el("span", undefined, k.telefon));
    tel.title = "Kliknij, aby zadzwonić";
    card.appendChild(tel);
  }

    const btn = el("button", "crm-btn crm-btn-cichy", "+ zgłoszenie") as HTMLButtonElement;
  const btnNota = el("button", "crm-btn crm-btn-cichy", "+ notatka") as HTMLButtonElement;
  const komunikat = el("span", "crm-komunikat");
  const rzad = el("div", "crm-akcje");
  rzad.appendChild(btn);
  rzad.appendChild(btnNota);
  rzad.appendChild(komunikat);
  card.appendChild(rzad);

  btn.onclick = () => void formularzReczny(card, btn, komunikat, k, c);
  btnNota.onclick = () => void formularzNotatki(card, btnNota, komunikat, k, c);

  return card;
}

/** Reczne zgloszenie - bez MailID i MailSubject, bo nie wychodzi z wiadomosci. */
async function formularzReczny(
  card: HTMLElement,
  btn: HTMLButtonElement,
  komunikat: HTMLElement,
  k: Contact,
  c: Company
) {
  btn.disabled = true;

  const box = el("div", "crm-nowe");
  card.appendChild(box);

  box.appendChild(el("label", "crm-label", "Opiekun supportu"));
  const wybor = document.createElement("select");
  wybor.className = "crm-select";
  box.appendChild(wybor);

  let zalogowany = "";
  try {
    zalogowany = await getMojMail();
  } catch {
    zalogowany = (Office.context.mailbox.userProfile.emailAddress ?? "").toLowerCase();
  }

  let serwisanci: Awaited<ReturnType<typeof getSerwisanci>> = [];
  try {
    serwisanci = await getSerwisanci();
  } catch (e: any) {
    box.appendChild(el("div", "crm-nowe-info", `Nie pobrano serwisantów: ${e?.message ?? e}`));
  }

  serwisanci.forEach((s, idx) => {
    const opt = document.createElement("option");
    opt.value = String(idx);
    opt.textContent = s.nazwa;
    if (s.mail === zalogowany) opt.selected = true;
    wybor.appendChild(opt);
  });
  if (!serwisanci.some((s) => s.mail === zalogowany)) {
    const pusta = document.createElement("option");
    pusta.value = "";
    pusta.textContent = "— wybierz opiekuna —";
    pusta.selected = true;
    wybor.insertBefore(pusta, wybor.firstChild);
  }

  const zapisz = el("button", "crm-btn", "Załóż zgłoszenie") as HTMLButtonElement;
  const anuluj = el("button", "crm-btn crm-btn-cichy", "Anuluj") as HTMLButtonElement;
  const rzad = el("div", "crm-akcje");
  rzad.appendChild(zapisz);
  rzad.appendChild(anuluj);
  box.appendChild(rzad);

  anuluj.onclick = () => {
    box.remove();
    btn.disabled = false;
  };

  zapisz.onclick = async () => {
    const wybrany = serwisanci[Number(wybor.value)];
    if (!wybrany) {
      komunikat.textContent = "Wybierz opiekuna";
      komunikat.className = "crm-komunikat crm-blad";
      return;
    }

    zapisz.disabled = true;
    wybor.disabled = true;
    komunikat.textContent = "Zakładanie...";
    komunikat.className = "crm-komunikat";

    const dzis = new Date();
    const dataISO = new Date(Date.UTC(
      dzis.getFullYear(), dzis.getMonth(), dzis.getDate(), 7, 0, 0
    )).toISOString();

    const pola: Record<string, unknown> = {
      Firma: k.firmaId || c.nazwa,
      Osoba: k.imieNazwisko,
      OsobaMail: k.email,
      Data: dataISO,
      Status: false,
      IdentyfikatorContact: k.guid,
      IdentyfikatorCompanies: k.companyGuid || c.guid,
      MailCreator: wybrany.mail,
      Dodatkowy_opiekun: wybrany.nazwa,
    };
    if (wybrany.siteId !== null) pola["Opiekun_supportuLookupId"] = String(wybrany.siteId);

    try {
      await createTicket(pola);
      box.remove();
      komunikat.textContent = `Założono dla ${wybrany.nazwa}`;
      komunikat.className = "crm-komunikat crm-ok";
    } catch (e: any) {
      komunikat.textContent = e?.message ?? String(e);
      komunikat.className = "crm-komunikat crm-blad";
      zapisz.disabled = false;
      wybor.disabled = false;
      btn.disabled = false;
    }
  };
}
/** Notatka przypieta do osoby i firmy - lista Notes. */
async function formularzNotatki(
  card: HTMLElement,
  btn: HTMLButtonElement,
  komunikat: HTMLElement,
  k: Contact,
  c: Company
) {
  btn.disabled = true;

  const box = el("div", "crm-nowe");
  card.appendChild(box);

  box.appendChild(el("div", "crm-nowe-info", `Notatka: ${k.imieNazwisko}`));

  // --- rodzaj interakcji (lookup do listy activities) ---
  box.appendChild(el("label", "crm-label", "Rodzaj"));
  const rodzaje = el("div", "crm-chipy crm-chipy-rowne");
  box.appendChild(rodzaje);

  let rodzaj = "";
  let rodzajId = "";
  try {
    const opcje = await getLookupOptions("activities", "Notes_types");
    if (!opcje.length) throw new Error("lista rodzajów jest pusta");
    opcje.forEach((o) => {
      const b = el("button", "crm-chip-wybor", o.etykieta) as HTMLButtonElement;
      b.title = o.etykieta;
      b.onclick = () => {
        rodzaj = o.etykieta;
        rodzajId = o.id;
        rodzaje.querySelectorAll(".crm-chip-wybor").forEach((x) =>
          x.classList.remove("crm-chip-akt")
        );
        b.classList.add("crm-chip-akt");
        sprawdz();
      };
      rodzaje.appendChild(b);
    });
  } catch (e: any) {
    box.appendChild(el("div", "crm-nowe-info", `Nie pobrano rodzajów: ${e?.message ?? e}`));
  }

  // --- data ---
  box.appendChild(el("label", "crm-label", "Data"));
  const data = document.createElement("input");
  data.type = "date";
  data.className = "crm-input";
  data.value = new Date().toISOString().slice(0, 10);
  box.appendChild(data);

  // --- tresc ---
  box.appendChild(el("label", "crm-label", "Treść"));
  const ta = document.createElement("textarea");
  ta.className = "crm-textarea";
  ta.rows = 4;
  ta.placeholder = "Co ustaliliście?";
  ta.oninput = () => sprawdz();
  box.appendChild(ta);

  // --- follow-up ---
  const fu = document.createElement("input");
  fu.type = "checkbox";
  fu.id = `fu-${k.id}`;
  const fuLbl = el("label", "crm-check") as HTMLLabelElement;
  fuLbl.htmlFor = fu.id;
  fuLbl.appendChild(fu);
  fuLbl.appendChild(el("span", undefined, "Przypomnienie (follow-up)"));
  box.appendChild(fuLbl);

  const fuData = document.createElement("input");
  fuData.type = "date";
  fuData.className = "crm-input";
  fuData.style.display = "none";
  const za7 = new Date();
  za7.setDate(za7.getDate() + 7);
  fuData.value = za7.toISOString().slice(0, 10);
  box.appendChild(fuData);

  const szybkie = el("div", "crm-chipy");
  szybkie.style.display = "none";
  [7, 14, 28, 56].forEach((d) => {
    const b = el("button", "crm-chip-wybor", `+${d} dni`) as HTMLButtonElement;
    b.onclick = () => {
      const t = new Date(data.value || new Date());
      t.setDate(t.getDate() + d);
      fuData.value = t.toISOString().slice(0, 10);
      szybkie.querySelectorAll(".crm-chip-wybor").forEach((x) => x.classList.remove("crm-chip-akt"));
      b.classList.add("crm-chip-akt");
    };
    szybkie.appendChild(b);
  });
  box.appendChild(szybkie);

  fu.onchange = () => {
    fuData.style.display = fu.checked ? "block" : "none";
    szybkie.style.display = fu.checked ? "flex" : "none";
  };

  // --- przyciski ---
  const brakujace = el("div", "crm-podpowiedz");
  box.appendChild(brakujace);

  const zapisz = el("button", "crm-btn", "Zapisz notatkę") as HTMLButtonElement;
  const anuluj = el("button", "crm-btn crm-btn-cichy", "Anuluj") as HTMLButtonElement;
  const rzadBtn = el("div", "crm-akcje");
  rzadBtn.appendChild(zapisz);
  rzadBtn.appendChild(anuluj);
  box.appendChild(rzadBtn);

  function sprawdz(): boolean {
    const brak: string[] = [];
    if (!rodzajId) brak.push("rodzaj");
    if (!ta.value.trim()) brak.push("treść");
    const ok = brak.length === 0;
    zapisz.disabled = !ok;
    brakujace.textContent = ok ? "" : `Uzupełnij: ${brak.join(", ")}.`;
    brakujace.style.display = ok ? "none" : "block";
    return ok;
  }
  sprawdz();

  anuluj.onclick = () => {
    box.remove();
    btn.disabled = false;
  };

  zapisz.onclick = async () => {
    if (!sprawdz()) return;

    zapisz.disabled = true;
    komunikat.textContent = "Zapisywanie...";
    komunikat.className = "crm-komunikat";

    // data bez strefy - SharePoint zapisuje jako dzien
    const dzien = (v: string) => (v ? new Date(v + "T07:00:00Z").toISOString() : null);

    let autor = "";
    try {
      autor = await getMojaNazwa();
    } catch {
      autor = Office.context.mailbox.userProfile.displayName ?? "";
    }

    const pola: Record<string, unknown> = {
      Data: dzien(data.value),
      Osoba: k.imieNazwisko,
      Telefon: k.telefon,
      TypeLookupId: rodzajId,
      Rodzaj1: rodzaj,
      Totatka: ta.value,            // uwaga: literowka w nazwie kolumny
      Firma: firmaPelna?.skrocona || c.nazwa,
      Creator: autor,
      IdentyfikatorCompanies: k.companyGuid || c.guid,
    };
    if (fu.checked && fuData.value) pola["Follow_up"] = dzien(fuData.value);

    try {
      await createListItem("Notes", pola);
      box.remove();
      btn.disabled = false;
      komunikat.textContent = "Notatka zapisana";
      komunikat.className = "crm-komunikat crm-ok";
    } catch (e: any) {
      komunikat.textContent = e?.message ?? String(e);
      komunikat.className = "crm-komunikat crm-blad";
      zapisz.disabled = false;
    }
  };
}

// ---------------------------------------------------------------- widok firmy

function naglowekFirmy(c: Company): HTMLElement {
  const box = el("div", "crm-firma-head");

  const powrot = el("button", "crm-powrot", "← wróć do firm") as HTMLButtonElement;
  powrot.onclick = () => void pokazFirmy();
  box.appendChild(powrot);

  const f = firmaPelna;
  // dane firmy w karcie - ten sam styl co karty na liscie
  const karta = el("div", "crm-firma-karta");
  box.appendChild(karta);
  karta.appendChild(poleIk("crm-firma-tytul", "firma", f?.skrocona || c.nazwa));
  if (f?.pelna && f.pelna !== f.skrocona) karta.appendChild(el("div", "crm-firma-pelna", f.pelna));
  if (f?.ulica || f?.miasto) {
    karta.appendChild(
      poleIk("crm-firma-adres", "pinezka", [f.ulica, [f.kod, f.miasto].filter(Boolean).join(" ")].filter(Boolean).join(", "), "Adres")
    );
  }

  // --- pastylki ---
  const pills = el("div", "crm-pills");

  const statusy: Record<string, [string, string]> = {
    Klient: ["#16A34A", "#E6F6EE"],
    Prospekt: ["#D97706", "#FEF3E2"],
    Prospect: ["#D97706", "#FEF3E2"],
    GSS: ["#2F6FD6", "#E8F0FD"],
    Nieaktywny: ["#6B7280", "#F3F4F6"],
  };
  const st = f?.status || "";
  const [sc, sb] = statusy[st] ?? ["#3860B2", "#EEF0FC"];
  pills.appendChild(pastylka(st || "brak statusu", sc, sb, "etykieta"));

  const dniNDA = dniDo(f?.dataNDA ?? "");
  const ndaKolor: [string, string] =
    !f?.dataNDA ? ["#6B7280", "#F3F4F6"]
    : dniNDA !== null && dniNDA < 0 ? ["#DC2626", "#FEE2E2"]
    : dniNDA !== null && dniNDA <= 30 ? ["#D97706", "#FEF3E2"]
    : ["#16A34A", "#E6F6EE"];
  pills.appendChild(
    pastylka("NDA " + (f?.dataNDA ? "do " + dataPL(f.dataNDA) : "brak"), ndaKolor[0], ndaKolor[1], "tarcza")
  );

  if (f?.opiekun) pills.appendChild(pastylka(f.opiekun, "#3860B2", "#EEF0FC", "opiekun"));
  if (tecsAlert) {
    const p = pastylka(
      `TECS kończy się za ${tecsAlert.dni} dni` +
        (tecsAlert.ile > 1 ? ` (${tecsAlert.ile} poz.)` : ""),
      "#DC2626",
      "#FEE2E2",
      "alert"
    );
    p.style.cursor = "pointer";
    p.title = "Pokaż licencje";
    p.onclick = () => {
      zakladkaFirmy = "Posiadane";
      posTryb = "Aktywne TECS";
      void pokazFirma(c);
    };
    pills.appendChild(p);
  }
  karta.appendChild(pills);

  // --- przyciski akcji ---
  const akcje = el("div", "crm-firma-akcje");

  if (f?.ulica || f?.miasto) {
    const mapa = przyciskIk("crm-akcja", "pinezka", "Nawiguj");
    mapa.title = `${f.ulica}, ${f.kod} ${f.miasto}`;
    mapa.onclick = () => {
      const adres = encodeURIComponent(`${f.ulica}, ${f.kod} ${f.miasto}, Polska`);
      window.open(`https://www.google.com/maps/dir/?api=1&travelmode=driving&destination=${adres}`, "_blank");
    };
    akcje.appendChild(mapa);
  }

  if (f?.nip) {
    const czysty = f.nip.toUpperCase().replace(/[\s-]/g, "").replace(/^PL/, "");
    const widoczny =
      czysty.length === 10
        ? `${czysty.slice(0, 3)}-${czysty.slice(3, 5)}-${czysty.slice(5, 7)}-${czysty.slice(7)}`
        : czysty;
    const nipBtn = przyciskIk("crm-akcja", "kopiuj", widoczny);
    nipBtn.title = "Kopiuj NIP";
    nipBtn.onclick = async () => {
      const ok = await kopiuj(czysty);
      blysk(nipBtn, ok ? "skopiowano" : "nie udało się");
    };
    akcje.appendChild(nipBtn);
  }

  const dok = przyciskIk("crm-akcja", "folder", "Dokumenty");
  dok.onclick = () =>
    window.open(
      "https://mescocompl.sharepoint.com/sites/CRMSystemPOC/Shared%20Documents/DokumentyKlienci/" +
        encodeURIComponent(f?.skrocona || c.nazwa),
      "_blank"
    );
  akcje.appendChild(dok);
  karta.appendChild(akcje);

  // --- liczniki godzin ---
  const liczniki = el("div", "crm-liczniki");
  ([
    ["TECS", f?.tecsH ?? 0, "#2F6FD6", "tarcza"],
    ["Support Day", f?.supportDay ?? 0, "#D97706", "klucz"],
    ["PreSales", f?.preSalesH ?? 0, "#3860B2", "wykres"],
  ] as [string, number, string, keyof typeof IKONY][]).forEach(([lbl, val, kol, ik]) => {
    const kafel = el("div", "crm-licznik");
    kafel.appendChild(poleIk("crm-licznik-lbl", ik, lbl));
    const v = el("div", "crm-licznik-val", godziny(val));
    v.style.color = val === 0 ? "#9CA3AF" : kol;
    kafel.appendChild(v);
    liczniki.appendChild(kafel);
  });
  karta.appendChild(liczniki);

  // --- zakladki (przewijane w poziomie) ---
  const pasek = el("div", "crm-zakladki");
    const nazwy: Zakladka[] = [
    "Kontakty", "Supporty", "Posiadane", "Notatki", "TECS", "Szanse", "Szkolenia", "Uwagi",
  ];
  nazwy.forEach((z) => {
    const b = el("button", "crm-zakladka", z) as HTMLButtonElement;
    if (z === zakladkaFirmy) b.classList.add("crm-zakladka-akt");
    b.onclick = () => {
      zakladkaFirmy = z;
      void pokazFirma(c);
    };
    pasek.appendChild(b);
  });
  // szeroki kafelek na cala siatke - kody demo laduja w TECS / Demo
  const demo = el("button", "crm-zakladka crm-zakladka-szeroka") as HTMLButtonElement;
  demo.append(ikona("licencja"), el("span", undefined, "Nowe kody demo"));
  demo.onclick = () => {
    zakladkaFirmy = "TECS";
    tecsTryb = "Demo";
    otworzKodyDemo = true;
    void pokazFirma(c);
  };
  pasek.appendChild(demo);
  box.appendChild(pasek);

  return box;
}

async function pokazFirma(c: Company) {
  widok = "firma";
  ostatniaFirma = c;
  ustawNaglowek(c.nazwa || "Firma", "nav-companies");
  root().innerHTML = "";

  // pelne dane firmy - tylko gdy zmienila sie firma
 if (!firmaPelna || firmaPelna.guid !== c.guid) {
    firmaPelna = null;
    tecsAlert = null;
    try {
      const item = await getCompanyByGuid(c.guid);
      if (item) firmaPelna = toCompanyFull(item);
    } catch (e) {
      console.error("Nie pobrano danych firmy", e);
    }
    // licencje konczace sie w ciagu 30 dni - zeby bylo widac przy czytaniu maila,
    // a nie dopiero po wejsciu w zakladke
    try {
      const lic = (await getByCompany("Oferty_TECS", c.guid, { top: 200 })).map(toLicencjaTecs);
      const bliskie = lic
        .map((l) => ({ l, d: dniDo(l.koniec) }))
        .filter((x) => !x.l.zamkniete && x.d !== null && x.d >= 0 && x.d <= 30);
      if (bliskie.length) {
        tecsAlert = {
          ile: bliskie.length,
          dni: Math.min(...bliskie.map((x) => x.d as number)),
        };
      }
    } catch (e) {
      console.error("Nie sprawdzono terminów TECS", e);
    }
  }

  root().appendChild(naglowekFirmy(c));

  const tresc = el("div");
  root().appendChild(tresc);
  const s = status("Pobieranie...");

  if (!c.guid) {
    s.textContent = "Firma nie ma CompaniesGUID — brak powiązań.";
    s.classList.add("crm-error");
    return;
  }

  try {
    switch (zakladkaFirmy) {
      case "Kontakty":  await zakladkaKontakty(tresc, s, c); break;
      case "Supporty":  await zakladkaSupporty(tresc, s, c); break;
      case "Notatki":   await zakladkaNotatki(tresc, s, c); break;
      case "TECS":      await zakladkaTecs(tresc, s, c); break;
      case "Szanse":    await zakladkaSzanse(tresc, s, c); break;
      case "Posiadane": await zakladkaPosiadane(tresc, s, c); break;
      case "Szkolenia": await zakladkaSzkolenia(tresc, s, c); break;
      case "Uwagi":     zakladkaUwagi(tresc, s); break;
    }
  } catch (e: any) {
    console.error(e);
    s.textContent = `Błąd: ${e?.message ?? e}`;
    s.classList.add("crm-error");
  }
}

// ---------------------------------------------------------------- zakladki

async function zakladkaKontakty(box: HTMLElement, s: HTMLElement, c: Company) {
  const items = await getContactsByCompany(c.guid);
  const wszyscy = items
    .map(toContact)
    .sort((a, b) => a.imieNazwisko.localeCompare(b.imieNazwisko, "pl"));

  // --- nowy kontakt w tej firmie ---
  const dodaj = el("button", "crm-btn crm-btn-cichy", "+ kontakt") as HTMLButtonElement;
  const nowyBox = el("div", "crm-nowe");
  nowyBox.style.display = "none";
  // szukajka (jesli sa kontakty) i przycisk w jednym rzedzie
  const pasekGorny = el("div", "crm-kontakt-pasek");
  pasekGorny.appendChild(dodaj);
  box.appendChild(pasekGorny);
  box.appendChild(nowyBox);
  dodaj.onclick = () => {
    dodaj.disabled = true;
    nowyBox.innerHTML = "";
    nowyBox.style.display = "block";
    const f: CompanyFull = firmaPelna ?? {
      id: "", guid: c.guid, skrocona: c.nazwa, pelna: "", nip: "", ulica: c.ulica,
      miasto: c.miasto, kod: c.kod, status: "", dataNDA: "", opiekun: c.opiekun,
      tecsH: 0, supportDay: 0, preSalesH: 0, uwagi: "", domena: "",
    };
    // dane nadawcy otwartego maila - chyba ze to ktos z MESco
    const od = (Office.context.mailbox.item as Office.MessageRead | undefined)?.from;
    const mail = (od?.emailAddress ?? "").toLowerCase();
    const wewnetrzny = mail.endsWith("@" + SUPPORT_MAILBOX.split("@")[1]);
    const start = mail && !wewnetrzny
      ? { ...rozbijNazwe(od?.displayName ?? ""), email: mail }
      : { imie: "", nazwisko: "", email: "" };

    void formularzKontaktu(
      nowyBox,
      f,
      start,
      () => void pokazFirma(c),
      () => {
        nowyBox.style.display = "none";
        dodaj.disabled = false;
      }
    );
  };

  if (!wszyscy.length) {
    s.textContent = "Kontakty — 0";
    return pusto(box, "Brak kontaktów dla tej firmy.");
  }

  // --- szukanie ---
  const szukaj = document.createElement("input");
  szukaj.type = "search";
  szukaj.className = "crm-szukaj";
  szukaj.placeholder = "szukaj osoby, stanowiska, maila";
  pasekGorny.insertBefore(szukaj, dodaj);

  // --- przelacznik: wszyscy / tylko pracujacy ---
  let tylkoAktywni = true;
  const maNieaktywnych = wszyscy.some((k) => !k.aktywny);
  if (maNieaktywnych) {
    segment(box, ["Pracujący", "Wszyscy"], tylkoAktywni ? "Pracujący" : "Wszyscy", (v) => {
      tylkoAktywni = v === "Pracujący";
      box.querySelectorAll(".crm-segment-btn").forEach((b) =>
        b.classList.toggle("crm-segment-akt", b.textContent === v)
      );
      rysuj();
    });
  }

  const lista = el("div");
  box.appendChild(lista);

  const rysuj = () => {
    const q = szukaj.value.trim().toLowerCase();
    const widoczni = wszyscy.filter((k) => {
      if (tylkoAktywni && !k.aktywny) return false;
      if (!q) return true;
      return [k.imieNazwisko, k.stanowisko, k.email, k.telefon]
        .filter(Boolean)
        .some((p) => p.toLowerCase().includes(q));
    });

    lista.innerHTML = "";
    s.textContent = q || !tylkoAktywni
      ? `Kontakty — ${widoczni.length} z ${wszyscy.length}`
      : `Kontakty — ${widoczni.length}`;

    if (!widoczni.length) {
      pusto(lista, q ? `Nic nie pasuje do „${szukaj.value.trim()}”.` : "Brak kontaktów.");
      return;
    }
    widoczni.forEach((k) => lista.appendChild(kartaKontaktu(k, c)));
  };

  let timer: any;
  szukaj.oninput = () => {
    clearTimeout(timer);
    timer = setTimeout(rysuj, 150);
  };

  rysuj();
}

async function zakladkaSupporty(box: HTMLElement, s: HTMLElement, c: Company) {
  const items = await getTicketsFiltered({ companyGuid: c.guid, top: 50 });
  s.textContent = `Zgłoszenia — ${items.length}`;
  if (!items.length) return pusto(box, "Brak zgłoszeń.");
  items.map(toTicket).forEach((t) => box.appendChild(karta(t)));
}

async function zakladkaNotatki(box: HTMLElement, s: HTMLElement, c: Company) {
  const items = await getByCompany("Notes", c.guid, { top: 50 });
    // kosz pokazujemy tylko przy wlasnych notatkach
  let ja = "";
  try {
    ja = (await getMojaNazwa()).trim().toLowerCase();
  } catch {
    ja = "";
  }
  const lista = items.map(toNota).sort((a, b) => (b.data || "").localeCompare(a.data || ""));
  s.textContent = `Notatki — ${lista.length}`;
  if (!lista.length) return pusto(box, "Brak notatek dla tej firmy.");

  lista.forEach((nt) => {
    const k = el("div", "crm-nota");
    const top = el("div", "crm-nota-top");
    top.appendChild(poleIk("crm-nota-data", "kalendarz", dataPL(nt.data), "Data"));

    // typ notatki w prawym rogu
    const prawo = el("div", "crm-nota-prawo");
    if (nt.rodzaj) prawo.appendChild(el("span", "crm-nota-typ", nt.rodzaj));
    const moja = !!ja && nt.autor.trim().toLowerCase() === ja;
    const kosz = el("button", "crm-usun") as HTMLButtonElement;
    kosz.appendChild(ikona("kosz"));
    kosz.title = "Usuń notatkę";
    if (moja) prawo.appendChild(kosz);
    top.appendChild(prawo);
    k.appendChild(top);

    if (nt.osoba) k.appendChild(poleIk("crm-nota-osoba", "osoba", nt.osoba, "Dla kogo"));
    if (nt.followUp) {
      const po = (dniDo(nt.followUp) ?? 0) < 0 && !nt.followDone;
      const fu = pastylka(
        (nt.followDone ? "✓ " : "") + "Follow-up " + dataPL(nt.followUp),
        po ? "#DC2626" : "#D97706",
        po ? "#FEE2E2" : "#FEF3E2",
        "dzwonek"
      );
      fu.classList.add("crm-nota-fu");
      k.appendChild(fu);
    }
    if (nt.tresc) {
      const tresc = el("div", "crm-nota-tresc");
      tresc.append(ikona("dymek"), el("div", "crm-nota-tresc-txt", nt.tresc));
      k.appendChild(tresc);
    }
    if (nt.autor) k.appendChild(poleIk("crm-nota-autor", "pioro", nt.autor, "Autor"));

    // potwierdzenie - notatka to zapis historii, nie kasujemy jednym kliknieciem
    const pytanie = el("div", "crm-nota-usun");
    pytanie.style.display = "none";
    pytanie.appendChild(el("div", undefined, "Usunąć tę notatkę? Tego nie można cofnąć."));

    const tak = el("button", "crm-btn crm-btn-ostrz", "Usuń") as HTMLButtonElement;
    const nie = el("button", "crm-btn crm-btn-cichy", "Anuluj") as HTMLButtonElement;
    const rzadU = el("div", "crm-akcje");
    rzadU.appendChild(tak);
    rzadU.appendChild(nie);
    pytanie.appendChild(rzadU);
    k.appendChild(pytanie);

    kosz.onclick = () => {
      pytanie.style.display = "block";
      kosz.style.visibility = "hidden";
    };
    nie.onclick = () => {
      pytanie.style.display = "none";
      kosz.style.visibility = "visible";
    };
    tak.onclick = async () => {
      tak.disabled = true;
      nie.disabled = true;
      tak.textContent = "Usuwanie...";
      try {
        await deleteListItem("Notes", nt.id);
        k.remove();
      } catch (e: any) {
        tak.textContent = "Usuń";
        tak.disabled = false;
        nie.disabled = false;
        pytanie.appendChild(el("div", "crm-komunikat crm-blad", e?.message ?? String(e)));
      }
    };

    box.appendChild(k);
  });
}

async function zakladkaTecs(box: HTMLElement, s: HTMLElement, c: Company) {
  segment(box, ["Bieżące", "Archiwalne", "Demo"], tecsTryb, (v) => {
    tecsTryb = v as typeof tecsTryb;
    void pokazFirma(c);
  });

  // nowe kody demo - formularz nad lista, otwierany kafelkiem pod zakladkami
  if (otworzKodyDemo) {
    otworzKodyDemo = false;
    const miejsce = el("div");
    box.appendChild(miejsce);
    void formularzKodowDemo(miejsce, c, () => void pokazFirma(c), () => miejsce.remove());
  }

  const items = await getByCompany("Opps_TECS", c.guid, { top: 100 });
  const wszystkie = items.map(toTecs);
  const lista =
    tecsTryb === "Bieżące"
      ? wszystkie.filter((t) => t.etap === "Przygotowanie oferty" || t.etap === "Oferta wysłana")
      : tecsTryb === "Archiwalne"
      ? wszystkie.filter((t) => ["Oferta wysłana", "Zamówione", "Rezygnacja z TECS"].includes(t.etap))
      : wszystkie.filter((t) => t.etap === "Demo");

  s.textContent = `TECS ${tecsTryb.toLowerCase()} — ${lista.length}`;
  if (!lista.length) return pusto(box, `Brak pozycji: ${tecsTryb}.`);

  lista.forEach((t) => {
    const k = el("div", "crm-kafel");
    const [kol, tlo] = ETAP_KOLORY[t.etap] ?? ["#6B7280", "#F3F4F6"];
    k.style.borderLeftColor = kol;

    const gora = el("div", "crm-kafel-rzad");
    gora.appendChild(poleIk("crm-kafel-nazwa", "cel", t.nazwa, "Nazwa szansy"));
    const etap = pastylka(ETAP_SKROTY[t.etap] ?? t.etap, kol, tlo, "flaga");
    etap.title = `Etap: ${t.etap}`;
    gora.appendChild(etap);
    k.appendChild(gora);

    const wart =
      t.demo ? (t.klient ? "Klient " + t.klient : "Kody demo") : t.tp ? "TechPool" : euro(t.wartosc);
    const ikWart = t.demo ? (t.klient ? "osoba" : "licencja") : t.tp ? "paczka" : "euro";
    k.appendChild(poleIk("crm-kafel-wartosc", ikWart, wart, "Wartość"));

    const d = dniDo(t.koniec);
    const meta = poleIk(
      "crm-kafel-meta",
      "klepsydra",
      (t.demo
        ? `${dataPL(t.poczatek)} → ${dataPL(t.koniec)}`
        : `${t.tp ? "TechPool" : "TECS"} do ${dataPL(t.koniec)}`) +
        (d === null ? "" : ` · ${d < 0 ? "po terminie" : d + (d === 1 ? " dzień" : " dni")}`),
      "Do kiedy"
    );
    meta.style.color = kolorTerminu(d);
    k.appendChild(meta);
    k.appendChild(pasekTerminu(d));

    if (t.demo) {
      if (t.odbiorca) k.appendChild(poleIk("crm-kafel-opps", "osoba", t.odbiorca, "Kody demo dla"));
      const rzadL = el("div", "crm-akcje");
      const lic = przyciskIk("crm-btn crm-btn-cichy", "paczka", "Dodaj licencje");
      rzadL.appendChild(lic);
      k.appendChild(rzadL);
      lic.onclick = () => {
        lic.disabled = true;
        const miejsce = el("div");
        k.appendChild(miejsce);
        void formularzLicencjiDemo(miejsce, t, c, () => {
          miejsce.remove();
          lic.disabled = false;
          blysk(lic, "licencje zapisane");
        }, () => {
          miejsce.remove();
          lic.disabled = false;
        });
      };
    }
    box.appendChild(k);
  });
}

// ---------------------------------------------------------------- kody demo

/** Administracja kodow demo - szybki wybor jak w Power Apps (domyslnie Monika). */
const ADMINI_DEMO = [
  { krotko: "Monika", nazwa: "Monika Szafarz", mail: "mszafarz@mesco.com.pl" },
  { krotko: "Ula", nazwa: "Urszula Horzela", mail: "uhorzela@mesco.com.pl" },
  { krotko: "Magda", nazwa: "Magdalena Turczyńska", mail: "mturczynska@mesco.com.pl" },
];

const dzisISO = () => {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
};
const plusDni = (iso: string, dni: number) => {
  const d = new Date((iso || dzisISO()) + "T12:00:00");
  d.setDate(d.getDate() + dni);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
};
// data bez strefy - SharePoint zapisuje jako dzien (jak w notatkach)
const dzienSP = (v: string) => (v ? new Date(v + "T07:00:00Z").toISOString() : null);

function poleFormularza(box: HTMLElement, etykieta: string, typ = "text"): HTMLInputElement {
  box.appendChild(el("label", "crm-label", etykieta));
  const i = document.createElement("input");
  i.type = typ;
  i.className = "crm-input";
  box.appendChild(i);
  return i;
}

function chipyWyboru<T>(
  box: HTMLElement,
  opcje: { tekst: string; wartosc: T }[],
  start: T | null,
  wybor: (v: T) => void
): HTMLElement {
  const rzad = el("div", "crm-chipy");
  opcje.forEach((o) => {
    const b = el("button", "crm-chip-wybor", o.tekst) as HTMLButtonElement;
    if (o.wartosc === start) b.classList.add("crm-chip-akt");
    b.onclick = () => {
      rzad.querySelectorAll(".crm-chip-wybor").forEach((x) => x.classList.remove("crm-chip-akt"));
      b.classList.add("crm-chip-akt");
      wybor(o.wartosc);
    };
    rzad.appendChild(b);
  });
  box.appendChild(rzad);
  return rzad;
}

/** Nowy wpis kodow demo w Opps_TECS (Demo = true, Etap1 = "Demo"). */
async function formularzKodowDemo(
  miejsce: HTMLElement,
  c: Company,
  poZapisie: () => void,
  anuluj: () => void
) {
  const f = firmaPelna;
  const skrocona = f?.skrocona || c.nazwa;

  const box = el("div", "crm-nowe crm-demo");
  miejsce.appendChild(box);
  box.appendChild(poleIk("crm-demo-tytul", "licencja", "Nowe kody demo"));
  box.appendChild(el("div", "crm-nowe-info", f?.pelna || skrocona));

  // --- dla kogo: kontakt z firmy albo wpisane recznie ---
  box.appendChild(el("label", "crm-label", "Kody demo dla"));
  const osobaSel = document.createElement("select");
  osobaSel.className = "crm-select";
  const brakOsoby = document.createElement("option");
  brakOsoby.value = "";
  brakOsoby.textContent = "wybierz osobę z firmy…";
  osobaSel.appendChild(brakOsoby);
  box.appendChild(osobaSel);
  const odbiorca = document.createElement("input");
  odbiorca.type = "text";
  odbiorca.className = "crm-input";
  odbiorca.placeholder = "imię i nazwisko";
  box.appendChild(odbiorca);

  let kontakty: Contact[] = [];
  let wybranaOsoba: Contact | null = null;
  try {
    kontakty = (await getContactsByCompany(c.guid)).map(toContact)
      .filter((k) => k.aktywny)
      .sort((a, b) => a.imieNazwisko.localeCompare(b.imieNazwisko, "pl"));
    kontakty.forEach((k, i) => {
      const o = document.createElement("option");
      o.value = String(i);
      o.textContent = k.imieNazwisko + (k.stanowisko ? ` · ${k.stanowisko}` : "");
      osobaSel.appendChild(o);
    });
  } catch (e: any) {
    brakOsoby.textContent = `kontakty: ${e?.message ?? e}`;
  }
  osobaSel.onchange = () => {
    wybranaOsoba = osobaSel.value ? kontakty[Number(osobaSel.value)] : null;
    if (wybranaOsoba) odbiorca.value = wybranaOsoba.imieNazwisko;
    sprawdz();
  };
  odbiorca.oninput = () => sprawdz();

  const numer = poleFormularza(box, "Numer klienta Ansys");

  // --- okres ---
  box.appendChild(el("label", "crm-label", "Długość demo"));
  let dlugosc = 14;
  const od = document.createElement("input");
  const doD = document.createElement("input");
  chipyWyboru(box, [7, 14, 21, 28].map((d) => ({ tekst: `${d} dni`, wartosc: d })), dlugosc, (d) => {
    dlugosc = d;
    doD.value = plusDni(od.value, dlugosc);
  });

  const okres = el("div", "crm-demo-okres");
  [["Od", od], ["Do", doD]].forEach(([lbl, inp]) => {
    const kol = el("div");
    kol.appendChild(el("label", "crm-label", lbl as string));
    (inp as HTMLInputElement).type = "date";
    (inp as HTMLInputElement).className = "crm-input";
    kol.appendChild(inp as HTMLInputElement);
    okres.appendChild(kol);
  });
  box.appendChild(okres);
  od.value = dzisISO();
  doD.value = plusDni(od.value, dlugosc);
  od.onchange = () => {
    doD.value = plusDni(od.value, dlugosc);
    fuData.value = plusDni(od.value, 7);
    sprawdz();
  };
  doD.onchange = () => sprawdz();

  // --- administracja ---
  box.appendChild(el("label", "crm-label", "Administracja"));
  let admin = ADMINI_DEMO[0];
  chipyWyboru(box, ADMINI_DEMO.map((a) => ({ tekst: a.krotko, wartosc: a })), admin, (a) => (admin = a));

  // --- spotkanie / follow-up ---
  const spotkanie = document.createElement("input");
  spotkanie.type = "checkbox";
  spotkanie.id = `demo-spot-${c.guid}`;
  const spLbl = el("label", "crm-check") as HTMLLabelElement;
  spLbl.htmlFor = spotkanie.id;
  spLbl.append(spotkanie, el("span", undefined, "Spotkanie z klientem"));
  box.appendChild(spLbl);

  // data spotkania tylko po zaznaczeniu - inaczej nic nie idzie do Data_Follow_Up
  const fuBox = el("div");
  fuBox.style.display = "none";
  box.appendChild(fuBox);
  fuBox.appendChild(el("label", "crm-label", "Follow-up z klientem"));
  const fuRzad = el("div", "crm-demo-okres");
  const fuData = document.createElement("input");
  fuData.type = "date";
  fuData.className = "crm-input";
  fuData.value = plusDni(od.value, 7);
  const fuGodz = document.createElement("input");
  fuGodz.type = "time";
  fuGodz.className = "crm-input";
  fuGodz.value = "09:00";
  fuRzad.append(fuData, fuGodz);
  fuBox.appendChild(fuRzad);
  spotkanie.onchange = () => (fuBox.style.display = spotkanie.checked ? "" : "none");

  // --- przyciski ---
  const brakujace = el("div", "crm-podpowiedz");
  box.appendChild(brakujace);
  const zapisz = el("button", "crm-btn", "Utwórz kody demo") as HTMLButtonElement;
  const wstecz = el("button", "crm-btn crm-btn-cichy", "Anuluj") as HTMLButtonElement;
  const komunikat = el("span", "crm-komunikat");
  const rzad = el("div", "crm-akcje");
  rzad.append(zapisz, wstecz, komunikat);
  box.appendChild(rzad);

  function sprawdz(): boolean {
    const brak: string[] = [];
    if (!odbiorca.value.trim()) brak.push("dla kogo");
    if (!od.value || !doD.value) brak.push("okres");
    else if (doD.value < od.value) brak.push("data „do” przed „od”");
    const ok = brak.length === 0;
    zapisz.disabled = !ok;
    brakujace.textContent = ok ? "" : `Uzupełnij: ${brak.join(", ")}.`;
    brakujace.style.display = ok ? "none" : "block";
    return ok;
  }
  sprawdz();

  wstecz.onclick = anuluj;

  zapisz.onclick = async () => {
    if (!sprawdz()) return;
    zapisz.disabled = true;
    komunikat.textContent = "Zapisywanie...";
    komunikat.className = "crm-komunikat";

    try {
      // pole Osoba wymaga Id uzytkownika witryny - opiekun to zalogowany
      const ja = await getMojMail();
      const opiekunId = await getSiteUserId(ja);
      if (opiekunId === null) {
        throw new Error(`nie znaleziono ${ja} na witrynie CRM - pole Opiekun jest wymagane`);
      }
      const adminId = await getSiteUserId(admin.mail);

      const [h, m] = (fuGodz.value || "09:00").split(":").map(Number);
      const [fy, fm, fd] = fuData.value.split("-").map(Number);
      const followUp =
        spotkanie.checked && fuData.value ? new Date(fy, fm - 1, fd, h, m).toISOString() : null;

      const guid = nowyGuid();
      const nazwa = `Kody demo_${skrocona}`;
      const pola: Record<string, unknown> = {
        Nazwaszansysprzeda_x017c_y_TECS: nazwa,
        Demo: true,
        Etap1: "Demo",
        TECSGUID: guid,
        IdentyfikatorCompanies: c.guid,
        FirmaID: skrocona,
        Numer_klienta: numer.value.trim(),
        Pocz_x0105_tek_TECS: dzienSP(od.value),
        Koniec_TECS: dzienSP(doD.value),
        Odbiorca_oferty: odbiorca.value.trim(),
        MailKontaktu: wybranaOsoba?.email ?? "",
        TelefonKontaktu: wybranaOsoba?.telefon ?? "",
        FollowUpKlient: spotkanie.checked,
        OpiekunLookupId: String(opiekunId),
      };
      if (followUp) pola["Data_Follow_Up"] = followUp;
      if (adminId !== null) pola["DemoAdmLookupId"] = String(adminId);

      const nowy = await createListItem("Opps_TECS", pola);

      // krok 2 - od razu licencje, jak w Power Apps po zapisie formularza
      box.remove();
      const kody: TecsRow = {
        id: String(nowy.id ?? (nowy.fields as any)?.id ?? ""), nazwa, wartosc: 0, etap: "Demo",
        poczatek: dzienSP(od.value) ?? "", koniec: dzienSP(doD.value) ?? "",
        tp: false, demo: true, klient: numer.value.trim(),
        guid, firmaId: skrocona, odbiorca: odbiorca.value.trim(),
      };
      const uwaga = adminId === null
        ? ` (bez administracji - nie znaleziono ${admin.mail} wśród użytkowników witryny CRM, szczegóły w konsoli)`
        : "";
      const ok = el("div", "crm-nowe-info crm-ok", `Kody demo utworzone${uwaga}. Dodaj licencje:`);
      miejsce.appendChild(ok);
      const krok2 = el("div");
      miejsce.appendChild(krok2);
      void formularzLicencjiDemo(krok2, kody, c, poZapisie, poZapisie);
    } catch (e: any) {
      komunikat.textContent = e?.message ?? String(e);
      komunikat.className = "crm-komunikat crm-blad";
      zapisz.disabled = false;
    }
  };
}

/** Licencje do kodow demo - wpis w Oferty_TECS powiazany przez IdentyfikatorTECS. */
async function formularzLicencjiDemo(
  miejsce: HTMLElement,
  kody: TecsRow,
  c: Company,
  poZapisie: () => void,
  anuluj: () => void
) {
  const box = el("div", "crm-nowe crm-demo");
  miejsce.appendChild(box);
  box.appendChild(poleIk("crm-demo-tytul", "paczka", "Dodaj licencje do kodów demo"));
  box.appendChild(
    el("div", "crm-nowe-info", `${kody.firmaId || c.nazwa} · ${dataPL(kody.poczatek)} – ${dataPL(kody.koniec)}`)
  );

  if (!kody.guid) {
    box.appendChild(el("div", "crm-komunikat crm-blad", "Ten wpis nie ma TECSGUID - licencji nie da się powiązać."));
    const zamknij = el("button", "crm-btn crm-btn-cichy", "Zamknij") as HTMLButtonElement;
    zamknij.onclick = anuluj;
    box.appendChild(zamknij);
    return;
  }

  box.appendChild(el("label", "crm-label", "Wybierz licencje z cennika"));
  const szukaj = document.createElement("input");
  szukaj.type = "search";
  szukaj.className = "crm-szukaj";
  szukaj.placeholder = "wyszukaj licencję";
  box.appendChild(szukaj);

  const wyniki = el("div", "crm-demo-wyniki");
  box.appendChild(wyniki);

  // wybrana z cennika pozycja czeka na liczbe licencji i "Dodaj"
  const doDodania = el("div", "crm-demo-dodaj");
  doDodania.style.display = "none";
  box.appendChild(doDodania);

  box.appendChild(el("label", "crm-label", "Licencje do wygenerowania"));
  const wybrane = el("div", "crm-demo-lista");
  box.appendChild(wybrane);

  const zapisz = el("button", "crm-btn", "Wyślij do generowania") as HTMLButtonElement;
  const wstecz = el("button", "crm-btn crm-btn-cichy", "Anuluj") as HTMLButtonElement;
  const komunikat = el("span", "crm-komunikat");
  const rzad = el("div", "crm-akcje");
  rzad.append(zapisz, wstecz, komunikat);
  box.appendChild(rzad);

  let cennik: PozycjaCennika[] = [];
  const wybor: { p: PozycjaCennika; ile: number }[] = [];
  let zapisanych = 0; // gdy padnie samo DemoOffice, przycisk zostaje do ponowienia

  const rysujWybrane = () => {
    wybrane.innerHTML = "";
    if (!wybor.length) wybrane.appendChild(el("div", "crm-podpowiedz", "Wyszukaj licencję i dodaj ją z liczbą sztuk."));
    wybor.forEach((w, i) => {
      const wiersz = el("div", "crm-demo-poz");
      wiersz.append(ikona("paczka"), el("span", "crm-ikona-txt", w.p.tytul), el("span", "crm-demo-ile", `×${w.ile}`));
      const x = el("button", "crm-demo-x", "×") as HTMLButtonElement;
      x.title = "Usuń";
      x.onclick = () => {
        wybor.splice(i, 1);
        rysujWybrane();
      };
      wiersz.appendChild(x);
      wybrane.appendChild(wiersz);
    });
    // wysylka dopiero gdy jest co wyslac
    zapisz.style.display = wybor.length || zapisanych ? "" : "none";
  };

  const pokazDodawanie = (p: PozycjaCennika) => {
    doDodania.innerHTML = "";
    doDodania.style.display = "";
    doDodania.appendChild(poleIk("crm-demo-dodaj-nazwa", "paczka", p.tytul));
    const rzadD = el("div", "crm-demo-dodaj-rzad");
    const ile = document.createElement("input");
    ile.type = "number";
    ile.min = "1";
    ile.value = "1";
    ile.className = "crm-input";
    ile.title = "Liczba licencji";
    const dodaj = el("button", "crm-btn", "Dodaj") as HTMLButtonElement;
    const nie = el("button", "crm-btn crm-btn-cichy", "Anuluj") as HTMLButtonElement;
    rzadD.append(el("span", "crm-label", "Liczba licencji"), ile, dodaj, nie);
    doDodania.appendChild(rzadD);
    ile.oninput = () => (dodaj.disabled = !(Number(ile.value) > 0));
    ile.onkeydown = (e: KeyboardEvent) => { if (e.key === "Enter" && !dodaj.disabled) dodaj.click(); };
    dodaj.onclick = () => {
      const juz = wybor.find((w) => w.p === p);
      if (juz) juz.ile += Number(ile.value);
      else wybor.push({ p, ile: Number(ile.value) });
      doDodania.style.display = "none";
      rysujWybrane();
      szukaj.focus();
    };
    nie.onclick = () => (doDodania.style.display = "none");
    ile.focus();
    ile.select();
  };

  const rysujWyniki = () => {
    wyniki.innerHTML = "";
    const q = szukaj.value.trim().toLowerCase();
    if (q.length < 2) return;
    cennik
      .filter((p) => p.tytul.toLowerCase().includes(q))
      .slice(0, 8)
      .forEach((p) => {
        const b = el("button", "crm-demo-wynik", p.tytul) as HTMLButtonElement;
        b.onclick = () => {
          szukaj.value = "";
          wyniki.innerHTML = "";
          pokazDodawanie(p);
        };
        wyniki.appendChild(b);
      });
    if (!wyniki.children.length) wyniki.appendChild(el("div", "crm-podpowiedz", "Brak w cenniku."));
  };
  szukaj.oninput = rysujWyniki;
  rysujWybrane();

  try {
    szukaj.placeholder = "wczytywanie cennika…";
    cennik = await getCennik();
    szukaj.placeholder = `wyszukaj licencję (${cennik.length} w cenniku)`;
    szukaj.focus();
  } catch (e: any) {
    szukaj.placeholder = "";
    box.insertBefore(el("div", "crm-komunikat crm-blad", `Cennik: ${e?.message ?? e}`), szukaj);
  }

  wstecz.onclick = anuluj;

  zapisz.onclick = async () => {
    if (!wybor.length && !zapisanych) return;
    zapisz.disabled = true;
    wstecz.disabled = true;
    komunikat.className = "crm-komunikat";
    try {
      // jeden wpis w Oferty_TECS na licencje - kazda ze swoja liczba
      while (wybor.length) {
        const w = wybor[0];
        komunikat.textContent = `Zapisywanie: ${w.p.tytul}...`;
        await createListItem("Oferty_TECS", {
          ProducID: w.p.nazwa,
          Liczba_licencji: w.ile,
          Pocz_x0105_tek_TECS: kody.poczatek || null,
          Koniec_TECS: kody.koniec || null,
          FirmaIDofertaTECS: kody.firmaId || c.nazwa,
          TECSOppsIDoferta: kody.nazwa,
          IdentyfikatorTECS: kody.guid,
          IdentyfikatorCompanies: c.guid,
          Demo: true,
        });
        wybor.shift(); // zapisana - przy bledzie ponowna proba nie zdubluje
        zapisanych++;
        rysujWybrane();
      }

      // jak w Power Apps: sa licencje -> DemoOffice = true, czyli zlecenie generowania
      komunikat.textContent = "Wysyłanie do generowania...";
      if (!kody.id) throw new Error("licencje zapisane, ale brak id wpisu kodów - zaznacz DemoOffice ręcznie");
      await updateTicketFields(kody.id, { DemoOffice: true }, undefined, "Opps_TECS");
      poZapisie();
    } catch (e: any) {
      wstecz.disabled = false;
      komunikat.textContent = e?.message ?? String(e);
      komunikat.className = "crm-komunikat crm-blad";
      zapisz.disabled = false;
    }
  };
}

async function zakladkaSzanse(box: HTMLElement, s: HTMLElement, c: Company) {
  segment(box, ["Aktywne", "Zamknięte"], szanseZamkniete ? "Zamknięte" : "Aktywne", (v) => {
    szanseZamkniete = v === "Zamknięte";
    void pokazFirma(c);
  });

  const items = await getByCompany("Opps", c.guid, { top: 100 });
  const lista = items
    .map(toSzansa)
    .filter((o) => (szanseZamkniete ? o.status.startsWith("Zamknięty") : o.status === "Aktywny"));

  s.textContent = `Szanse — ${lista.length}`;
  if (!lista.length) return pusto(box, szanseZamkniete ? "Brak zamkniętych szans." : "Brak aktywnych szans.");

  lista.forEach((o) => {
    const k = el("div", "crm-kafel");
    const [kol, tlo]: [string, string] =
      o.status === "Zamknięty wygrany" ? ["#16A34A", "#E6F6EE"]
      : o.status === "Zamknięty przegrany" ? ["#DC2626", "#FEE2E2"]
      : ["#3860B2", "#EEF0FC"];
    k.style.borderLeftColor = kol;

    const gora = el("div", "crm-kafel-rzad");
    gora.appendChild(poleIk("crm-kafel-nazwa", "cel", o.nazwa, "Nazwa szansy"));
    const etap = pastylka(o.status || "—", kol, tlo, "flaga");
    etap.title = "Etap";
    gora.appendChild(etap);
    k.appendChild(gora);

    const rzad = el("div", "crm-kafel-rzad");
    rzad.appendChild(poleIk("crm-kafel-wartosc", "euro", euro(o.wartosc), "Wartość"));
    if (o.typ) rzad.appendChild(pastylka(o.typ, "#2F6FD6", "#E8F0FD", "teczka"));
    k.appendChild(rzad);

    k.appendChild(poleIk("crm-kafel-meta", "kalendarz", `Otwarta ${dataPL(o.otwarcie)}`, "Data otwarcia"));
    box.appendChild(k);
  });
}

async function zakladkaPosiadane(box: HTMLElement, s: HTMLElement, c: Company) {
  segment(box, ["Aktywne TECS", "Wygasłe", "Bez TECS"], posTryb, (v) => {
    posTryb = v as typeof posTryb;
    void pokazFirma(c);
  });

  let lista: Licencja[];
  if (posTryb === "Bez TECS") {
    const items = await getByCompany("StanPosiadania", c.guid, { top: 100 });
    lista = items.map(toLicencjaBez).sort((a, b) => a.nazwa.localeCompare(b.nazwa, "pl"));
  } else {
    const items = await getByCompany("Oferty_TECS", c.guid, { top: 200 });
    const wszystkie = items.map(toLicencjaTecs);
    const aktywna = (l: Licencja) => !l.zamkniete && (dniDo(l.koniec) ?? -1) >= 0;
    lista =
      posTryb === "Aktywne TECS"
        ? wszystkie.filter(aktywna).sort((a, b) => (a.koniec || "").localeCompare(b.koniec || ""))
        : wszystkie.filter((l) => !aktywna(l)).sort((a, b) => (b.koniec || "").localeCompare(a.koniec || ""));
  }

  s.textContent = `${posTryb} — ${lista.length}`;
  if (!lista.length) return pusto(box, `Brak licencji: ${posTryb}.`);

  lista.forEach((l) => {
    const k = el("div", "crm-kafel");
    const d = dniDo(l.koniec);
    k.style.borderLeftColor = !l.koniec || (d !== null && d < 0) ? "#CFCDCB" : kolorTerminu(d);

    const top = el("div", "crm-kafel-rzad");
    top.appendChild(poleIk("crm-kafel-nazwa", "paczka", l.nazwa, "Produkt"));
    const ile = pastylka(`${l.ile} lic.`, "#0D9488", "#E6F6F4", "licencja");
    ile.title = "Liczba licencji";
    top.appendChild(ile);
    k.appendChild(top);

    if (l.opps) k.appendChild(poleIk("crm-kafel-opps", "dokument", l.opps, "Oferta"));

    const meta = poleIk(
      "crm-kafel-meta",
      "klepsydra",
      !l.koniec
        ? "bez TECS"
        : `TECS do ${dataPL(l.koniec)}` + (d === null ? "" : ` · ${d < 0 ? "wygasł" : d + " dni"}`),
      "Ważność TECS"
    );
    meta.style.color = d !== null && d < 0 ? "#6B7280" : kolorTerminu(d);
    k.appendChild(meta);

    if (l.tp || l.demo) k.appendChild(pastylka(l.tp ? "TechPool" : "Demo", "#7C3AED", "#F1EAFE", "etykieta"));
    if (d !== null && d >= 0) k.appendChild(pasekTerminu(d));
    box.appendChild(k);
  });
}

async function zakladkaSzkolenia(box: HTMLElement, s: HTMLElement, c: Company) {
  // uwaga: ta lista uzywa GUIDFirmy, nie IdentyfikatorCompanies
  const items = await getByCompany("UczestnicySzkolen", c.guid, { pole: "GUIDFirmy", top: 100 });
  const lista = items.map(toSzkolenie);
  if (!lista.length) {
    s.textContent = "Szkolenia";
    return pusto(box, "Brak szkoleń dla tej firmy.");
  }

  // lista ma wiersz na uczestnika - grupujemy w jedno szkolenie z lista osob
  const czas = (d: string) => {
    const t = Date.parse(d);
    return isNaN(t) ? 0 : t;
  };
  const grupy = new Map<string, { nazwa: string; data: string; osoby: string[] }>();
  lista.forEach((z) => {
    const klucz = `${z.nazwa}|${z.data}`;
    const g = grupy.get(klucz) ?? { nazwa: z.nazwa, data: z.data, osoby: [] };
    if (z.osoba) g.osoby.push(z.osoba);
    grupy.set(klucz, g);
  });
  const szkolenia = [...grupy.values()].sort((a, b) => czas(b.data) - czas(a.data)); // najnowsze u gory

  s.textContent = `Szkolenia — ${szkolenia.length} · uczestników: ${lista.length}`;

  const dzis = new Date();
  dzis.setHours(0, 0, 0, 0);

  szkolenia.forEach((g) => {
    const k = el("div", "crm-szkolenie");

    const top = el("div", "crm-szkolenie-top");
    const tytul = el("div", "crm-szkolenie-nazwa");
    tytul.title = "Szkolenie";
    tytul.append(ikona("czapka"), el("span", "crm-ikona-txt", g.nazwa || "(bez nazwy)"));
    top.appendChild(tytul);
    if (czas(g.data) >= dzis.getTime()) top.appendChild(el("span", "crm-szkolenie-nowe", "nadchodzące"));
    k.appendChild(top);

    if (g.data) {
      const data = el("div", "crm-szkolenie-data");
      data.title = "Data szkolenia";
      data.append(ikona("kalendarz"), el("span", undefined, czas(g.data) ? formatData(g.data) : g.data));
      k.appendChild(data);
    }

    if (g.osoby.length) {
      const osoby = el("div", "crm-szkolenie-osoby");
      osoby.title = "Uczestnicy";
      g.osoby.sort((a, b) => a.localeCompare(b, "pl")).forEach((o) => {
        const chip = el("span", "crm-szkolenie-osoba");
        chip.append(ikona("osoba"), el("span", undefined, o));
        osoby.appendChild(chip);
      });
      k.appendChild(osoby);
    }

    box.appendChild(k);
  });
}

function zakladkaUwagi(box: HTMLElement, s: HTMLElement) {
  s.textContent = "Uwagi do firmy";
  const t = (firmaPelna?.uwagi ?? "").trim();
  box.appendChild(el("div", t ? "crm-uwagi" : "crm-pusto", t || "Brak uwag dla tej firmy."));
}

// ---------------------------------------------------------------- kontakty (cala baza)

/** Wyszukiwarka po wszystkich kontaktach - z przejsciem do firmy. */
async function pokazKontakty() {
  widok = "kontakty";
  ustawNaglowek("Kontakty", "nav-contacts");
  root().innerHTML = "";
  // firmaPelna dotyczy ostatnio otwartej firmy - formularz notatki nie moze jej uzyc
  firmaPelna = null;
  tecsAlert = null;

  const szukaj = document.createElement("input");
  szukaj.type = "search";
  szukaj.className = "crm-szukaj";
  szukaj.placeholder = "szukaj: nazwisko, mail, telefon, firma";
  szukaj.value = frazaKontaktow;
  root().appendChild(szukaj);

  const s = status("Pobieranie kontaktów...");
  const lista = el("div");
  root().appendChild(lista);

  let wszyscy: Contact[];
  try {
    wszyscy = (await getAllContacts())
      .map(toContact)
      .sort((a, b) =>
        Number(b.aktywny) - Number(a.aktywny) || a.imieNazwisko.localeCompare(b.imieNazwisko, "pl")
      );
  } catch (e: any) {
    console.error(e);
    s.textContent = `Błąd: ${e?.message ?? e}`;
    s.classList.add("crm-error");
    return;
  }

  const LIMIT = 50;

  const rysuj = () => {
    const q = szukaj.value.trim().toLowerCase();
    frazaKontaktow = szukaj.value;
    lista.innerHTML = "";

    if (q.length < 2) {
      s.textContent = `Kontakty w bazie — ${wszyscy.length}`;
      pusto(lista, "Wpisz co najmniej 2 znaki.");
      return;
    }

    const trafienia = wszyscy.filter((k) =>
      [k.imieNazwisko, k.email, k.telefon, k.stanowisko, k.firmaId]
        .filter(Boolean)
        .some((p) => p.toLowerCase().includes(q))
    );

    s.textContent = trafienia.length > LIMIT
      ? `Znaleziono ${trafienia.length} — pokazuję ${LIMIT}, zawęź wyszukiwanie`
      : `Znaleziono — ${trafienia.length}`;

    if (!trafienia.length) {
      pusto(lista, `Nic nie pasuje do „${szukaj.value.trim()}”.`);
      return;
    }

    trafienia.slice(0, LIMIT).forEach((k) => {
      const c = { guid: k.companyGuid, nazwa: k.firmaId } as Company;
      const card = kartaKontaktu(k, c);

      // firma kontaktu - klik prowadzi do pelnego widoku firmy
      if (k.companyGuid) {
        const doFirmy = el("button", "crm-kontakt-firma", `↗ ${k.firmaId || "firma"}`) as HTMLButtonElement;
        doFirmy.title = "Otwórz firmę";
        doFirmy.onclick = () => {
          zakladkaFirmy = "Kontakty";
          void pokazFirma(c);
        };
        card.insertBefore(doFirmy, card.children[1] ?? null);
      }
      lista.appendChild(card);
    });
  };

  let timer: any;
  szukaj.oninput = () => {
    clearTimeout(timer);
    timer = setTimeout(rysuj, 150);
  };

  rysuj();
  szukaj.focus();
}

// ---------------------------------------------------------------- odswiezanie

async function odswiez() {
  if (widok === "firmy") {
    wyczyscCacheFirm();
    void getLiczbaSupportowRok(true);
    return pokazFirmy();
  }
  if (widok === "kontakty") {
    wyczyscCacheKontaktow();
    return pokazKontakty();
  }
  if (widok === "firma" && ostatniaFirma) {
    firmaPelna = null; // wymus ponowne pobranie danych firmy
    return pokazFirma(ostatniaFirma);
  }
  return pokazZgloszenia();
}
