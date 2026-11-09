// src/taskpane/outlook.ts
/* global document, Office, console, setTimeout, window, navigator, clearTimeout */
import {
  getListItems, getTicketsByField, updateTicketFields, getCompanies2, getOpiekunowie,
  wyczyscCacheFirm, getContactByEmail, createTicket, getSerwisanci, getContactsByCompany,
  getTicketsFiltered, getMojMail, getByCompany, getCompanyByGuid,createListItem, getLookupOptions, getMojaNazwa, deleteListItem,
  getTicketByNumber, getAllCompanies, getChoiceOptions, getAllContacts, wyczyscCacheKontaktow,
  getLiczbaSupportowRok, getSiteUserId, getCennik, getOpcjeLookupu, czyPracownikMESco, getZgloszeniaObslugujacego, nazwaWewnetrzna, getCenyTecs,
} from "./graphService";
import type { ListItem, PozycjaCennika, CenaTecs } from "./graphService";
import {
  toTicket, toCompany, formatData, dniOdZmiany, KOLOR_STATUSU, ticketIdZTematu, Ticket, Company,
  toContact, Contact, toCompanyFull, CompanyFull, toNota, Nota, toTecs, toSzansa, Szansa,
  toLicencjaTecs, toLicencjaBez, Licencja, toSzkolenie, dniDo, kolorTerminu,
  ETAP_KOLORY, ETAP_SKROTY, rozbijNazwe, plecZImienia, nowyGuid, TecsRow, toProduktTecs, ProduktTecs,
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
  document.getElementById("nav-szanse")!.onclick = () => void pokazSzanse();

  Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => {
    if (widok === "zgloszenia") void pokazZgloszenia();
    if (widok === "szanse") void pokazSzanse();
  });

  void start();
});

/** Handlowiec (MEScoStaff.Email) we wlasnej skrzynce startuje od szans firmy nadawcy. */
async function start() {
  root().appendChild(el("div", "crm-status", "Wczytywanie..."));
  let handlowiec = false;
  try {
    handlowiec = await czyPracownikMESco();
  } catch (e) {
    console.error("MEScoStaff:", e);
  }
  document.getElementById("nav-szanse")!.hidden = !handlowiec;
  document.querySelector(".crm-dol")!.classList.toggle("crm-dol-5", handlowiec);

  const supportowa = await czySkrzynkaSupportu(Office.context.mailbox.item as Office.MessageRead);
  if (handlowiec && !supportowa) return pokazSzanse();
  return pokazZgloszenia();
}

// ---------------------------------------------------------------- stan

type Widok = "zgloszenia" | "firmy" | "firma" | "kontakty" | "szanse";
let widok: Widok = "zgloszenia";
let ostatniaFirma: Company | null = null;

type ZakresZgl = "moje" | "mail" | "firma" | "wszystkie";
// osobno dla skrzynki supportu i wlasnej - we wlasnej domyslnie "Moje zgloszenia"
const zakresy: Record<"support" | "wlasna", ZakresZgl> = { support: "mail", wlasna: "moje" };
let pokazZglLicznik = 0;
let opiekunZgloszen = "";
let frazaKontaktow = "";

type Zakladka =
  | "Kontakty" | "Supporty" | "Notatki" | "TECS"
  | "Szanse" | "Posiadane" | "Szkolenia" | "Uwagi";

let zakladkaFirmy: Zakladka = "Kontakty";
let tecsTryb: "Bieżące" | "Archiwalne" | "Demo" = "Bieżące";
let otworzKodyDemo = false; // kafelek "Nowe kody demo" -> TECS/Demo z otwartym formularzem
let otworzNowyTecs = false; // "Zapisz i utworz TECS" przy wygranej -> TECS/Biezace z otwartym formularzem
let nowyTecsZSzansy: Szansa | null = null; // wygrana szansa - jej produkty podpowiadamy do TECS-u
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

type NavId = "nav-szanse" | "nav-tickets" | "nav-companies" | "nav-contacts";

function ustawNaglowek(tytul: string, aktywny: NavId) {
  document.getElementById("nav-title")!.textContent = tytul;
  (["nav-szanse", "nav-tickets", "nav-companies", "nav-contacts"] as NavId[]).forEach((id) =>
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
  uscisk: '<path d="m11 17 2 2a1 1 0 1 0 3-3"/><path d="m14 14 2.5 2.5a1 1 0 1 0 3-3l-3.88-3.88a3 3 0 0 0-4.24 0l-.88.88a1 1 0 1 1-3-3l2.81-2.81a5.79 5.79 0 0 1 7.06-.87l.47.28a2 2 0 0 0 1.42.25L21 4"/><path d="m21 3 1 11h-2"/><path d="M3 3 2 14l6.5 6.5a1 1 0 1 0 3-3"/><path d="M3 4h8"/>',
  wideo: '<path d="m16 13 5.223 3.482a.5.5 0 0 0 .777-.416V7.87a.5.5 0 0 0-.752-.432L16 10.5"/><rect x="2" y="6" width="14" height="12" rx="2"/>',
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

/**
 * Rodzaje notatek (activities.Notes_types) maja w nazwie emoji, np. "📞 Telefon".
 * W SharePoint zostaja bez zmian, w dodatku pokazujemy ikone + tekst bez emoji.
 */
const EMOJI = /[\u2190-\u2BFF\u2600-\u27BF\uFE0F\u200D\u20E3]|[\uD83C-\uD83E][\uDC00-\uDFFF]/g;

const IKONA_EMOJI: [string, keyof typeof IKONY][] = [
  ["📞☎📱📲", "telefon"],
  ["📧✉📨📩📤📥", "mail"],
  ["🤝", "uscisk"],
  ["💻🖥🎥📹📺", "wideo"],
  ["👥👤🧑👨👩", "osoba"],
  ["📝📄📃🗒📋✍", "dokument"],
  ["💰💶💵💸💳🧾", "euro"],
  ["🎓📚", "czapka"],
  ["📅🗓📆", "kalendarz"],
  ["🚗🚙🏢📍", "pinezka"],
  ["🔧🛠⚙", "klucz"],
  ["🎯", "cel"],
  ["🔔⏰", "dzwonek"],
];

const IKONA_SLOWA: [RegExp, keyof typeof IKONY][] = [
  [/telefon|rozmow|dzwon/, "telefon"],
  [/mail|wiadomo/, "mail"],
  [/teams|online|webinar|wideo|video|zdaln/, "wideo"],
  [/spotkan|wizyt/, "uscisk"],
  [/ofert|faktur|zam[oó]wien/, "euro"],
  [/szkolen|warsztat/, "czapka"],
  [/support|serwis/, "klucz"],
  [/demo|licencj/, "licencja"],
  [/notatk|dokument/, "dokument"],
];

function rodzajNotatki(etykieta: string): { tekst: string; ik: keyof typeof IKONY } {
  const tekst = etykieta.replace(EMOJI, "").replace(/\s+/g, " ").trim() || etykieta;
  const emoji = etykieta.match(EMOJI)?.join("") ?? "";
  const zEmoji = emoji && IKONA_EMOJI.find(([e]) => [...emoji].some((z) => z.trim() && e.includes(z)));
  if (zEmoji) return { tekst, ik: zEmoji[1] };
  const zeSlowa = IKONA_SLOWA.find(([re]) => re.test(tekst.toLowerCase()));
  return { tekst, ik: zeSlowa ? zeSlowa[1] : "dymek" };
}

/** Kafelek wyboru rodzaju notatki: ikona + nazwa bez emoji. */
function chipRodzaju(etykieta: string): HTMLButtonElement {
  const { tekst, ik } = rodzajNotatki(etykieta);
  const b = el("button", "crm-chip-wybor crm-chip-rodzaj") as HTMLButtonElement;
  b.title = tekst;
  b.append(ikona(ik), el("span", "crm-chip-txt", tekst));
  return b;
}

const MAX_MAIL_BODY = 60000; // zapas pod limit kolumny wielowierszowej SharePoint

/**
 * Przy rodzaju "mail" - opcja dolaczenia otwartej wiadomosci do notatki.
 * Tresc maila idzie do MailBody, nie do pola notatki (Totatka).
 */
function opcjaMaila(box: HTMLElement, zmiana: () => void) {
  const item = Office.context.mailbox.item as Office.MessageRead | undefined;
  const w = el("div", "crm-mail-opcja");
  w.style.display = "none";
  const cb = document.createElement("input");
  cb.type = "checkbox";
  cb.id = `mail-${nowyGuid()}`;
  const lbl = el("label", "crm-check") as HTMLLabelElement;
  lbl.htmlFor = cb.id;
  lbl.append(cb, el("span", undefined, "Zapisz otwarty mail w notatce"));
  w.appendChild(lbl);
  if (item?.subject != null) {
    w.appendChild(poleIk("crm-mail-opcja-info", "mail", item.subject || "(bez tematu)", "Temat"));
    const od = item.from?.emailAddress ?? "";
    if (od) w.appendChild(poleIk("crm-mail-opcja-info", "osoba", od, "Nadawca"));
  }
  box.appendChild(w);
  cb.onchange = zmiana;

  return {
    pokaz(rodzaj: string) {
      const ok = !!item?.body && !!rodzaj && rodzajNotatki(rodzaj).ik === "mail";
      w.style.display = ok ? "" : "none";
      if (!ok) cb.checked = false;
    },
    get wlaczony() {
      return cb.checked && w.style.display !== "none";
    },
    async pola(): Promise<Record<string, unknown>> {
      if (!item) return {};
      const tresc = await new Promise<string>((ok, blad) =>
        item.body.getAsync(Office.CoercionType.Text, (r) =>
          r.status === Office.AsyncResultStatus.Succeeded ? ok(r.value ?? "") : blad(new Error(r.error?.message ?? "nie odczytano treści maila"))
        )
      );
      const temat = item.subject ?? "";
      return {
        MailBody: tresc.length > MAX_MAIL_BODY ? tresc.slice(0, MAX_MAIL_BODY) + "\n[…treść skrócona]" : tresc,
        MailFrom: item.from?.emailAddress ?? "",
        MailTo: (item.to ?? []).map((a) => a.emailAddress).filter(Boolean).join("; "),
        MailSubject: temat,
        MailTitle: temat,
      };
    },
  };
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
  if (typeof it?.getSharedPropertiesAsync !== "function") return false;
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
  const wywolanie = ++pokazZglLicznik;

  const item = Office.context.mailbox.item as Office.MessageRead;
  const nadawca = (item?.from?.emailAddress ?? "").toLowerCase();
  const klucz = (await czySkrzynkaSupportu(item)) ? "support" : "wlasna";
  if (wywolanie !== pokazZglLicznik || widok !== "zgloszenia") return; // w miedzyczasie zmiana maila/widoku
  let zakresZgloszen = zakresy[klucz];
  root().innerHTML = "";

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
    { v: "moje", t: "Moje zgłoszenia" },
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
    if (zakresZgloszen === "mail" || zakresZgloszen === "firma") {
      zakres.value = "wszystkie";
      zakresZgloszen = "wszystkie";
    }
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

      if (zakresZgloszen === "moje") {
        await listaMoichZgloszen(lista, s, () => moje !== pokolenie);
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

  // filtr opiekuna nie ma sensu przy "Moje zgloszenia"
  const pokazOpiekuna = () => (wybor.style.display = zakresZgloszen === "moje" ? "none" : "");
  pokazOpiekuna();
  zakres.onchange = () => {
    zakresZgloszen = zakres.value as ZakresZgl;
    zakresy[klucz] = zakresZgloszen;
    pokazOpiekuna();
    void wczytaj();
  };
  wybor.onchange = () => {
    opiekunZgloszen = wybor.value;
    void wczytaj();
  };

  await wczytaj();
}


// ---------------------------------------------------------------- szanse firmy nadawcy

/** Firma nadawcy: z jego kontaktu, a gdy go nie ma - jedyna firma z ta domena. */
async function firmaNadawcy(nadawca: string): Promise<Company | null> {
  const kontakt = await kontaktDlaNadawcy(nadawca);
  const guid = String((kontakt?.fields as any)?.IdentyfikatorCompanies ?? "");
  if (guid) {
    const item = await getCompanyByGuid(guid);
    if (item) return toCompany(item);
  }
  const domena = nadawca.split("@")[1] ?? "";
  if (!domena) return null;
  const pasujace = (await getAllCompanies()).filter(
    (i) => String((i.fields as any).DomenA ?? "").trim().toLowerCase() === domena
  );
  return pasujace.length === 1 ? toCompany(pasujace[0]) : null;
}

let pokazSzansLicznik = 0;

async function pokazSzanse() {
  widok = "szanse";
  ustawNaglowek("Szanse", "nav-szanse");
  const wywolanie = ++pokazSzansLicznik;
  root().innerHTML = "";

  const item = Office.context.mailbox.item as Office.MessageRead;
  const nadawca = (item?.from?.emailAddress ?? "").toLowerCase();
  const s = status("Szukanie firmy nadawcy...");
  if (!nadawca) {
    s.textContent = "Wybierz wiadomość, aby zobaczyć szanse firmy nadawcy.";
    return;
  }

  try {
    const c = await firmaNadawcy(nadawca);
    if (wywolanie !== pokazSzansLicznik || widok !== "szanse") return;
    if (!c) {
      s.textContent = `${nadawca} nie jest powiązany z żadną firmą w CRM.`;
      return;
    }

    // naglowek z firma - klik otwiera pelna karte firmy na zakladce Szanse
    const firma = el("button", "crm-szanse-firma") as HTMLButtonElement;
    firma.append(ikona("firma"), el("span", "crm-ikona-txt", c.nazwa || "(bez nazwy)"));
    firma.title = "Otwórz kartę firmy";
    firma.onclick = () => {
      zakladkaFirmy = "Szanse";
      void pokazFirma(c);
    };
    root().insertBefore(firma, s);

    const box = el("div");
    root().appendChild(box);
    await zakladkaSzanse(box, s, c, () => void pokazSzanse());
  } catch (e: any) {
    console.error(e);
    s.textContent = `Błąd: ${e?.message ?? e}`;
    s.classList.add("crm-error");
  }
}

// ---------------------------------------------------------------- moje zgloszenia

// kolejnosc jak w obiegu zgloszenia; zalatwione zwiniete, bo bywa ich setki
const KATEGORIE_ZGLOSZEN = [
  { nazwa: "Przydzielony", otwarta: true },
  { nazwa: "Oczekuje na odpowiedź", otwarta: true },
  { nazwa: "Rozwiązany", otwarta: false },
  { nazwa: "Zamknięty", otwarta: false },
];
const ZGLOSZEN_NA_STRONE = 30;

function kategoriaZgloszenia(t: Ticket): string {
  if (t.status) return t.status;
  return t.zamkniete ? "Zamknięty" : "Bez statusu";
}

/** Wszystkie zgloszenia zalogowanego (MailCreator) w sekcjach wedlug StatusTicket. */
async function listaMoichZgloszen(lista: HTMLElement, s: HTMLElement, nieaktualne: () => boolean) {
  const mail = await getMojMail();
  const tickety = (await getZgloszeniaObslugujacego(mail)).map(toTicket);
  if (nieaktualne()) return;

  s.textContent = `Moje zgłoszenia (${tickety.length})`;
  if (!tickety.length) {
    pusto(lista, `Brak zgłoszeń, w których obsługującym jest ${mail}.`);
    return;
  }

  const grupy = new Map<string, Ticket[]>();
  tickety.forEach((t) => {
    const k = kategoriaZgloszenia(t);
    grupy.set(k, [...(grupy.get(k) ?? []), t]);
  });
  const znane = KATEGORIE_ZGLOSZEN.map((k) => k.nazwa);
  const kolejnosc = [...znane, ...[...grupy.keys()].filter((k) => !znane.includes(k))];

  kolejnosc.forEach((nazwa) => {
    const wGrupie = grupy.get(nazwa) ?? [];
    const def = KATEGORIE_ZGLOSZEN.find((k) => k.nazwa === nazwa);
    if (!wGrupie.length && !def) return;

    const sek = document.createElement("details");
    sek.className = "crm-kat";
    sek.open = !!def?.otwarta && wGrupie.length > 0;
    const naglowek = document.createElement("summary");
    naglowek.className = "crm-kat-naglowek";
    const kropka = el("span", "crm-kat-kropka");
    kropka.style.background = KOLOR_STATUSU[nazwa] ?? "#3860B2";
    naglowek.append(kropka, el("span", "crm-kat-nazwa", nazwa), el("span", "crm-kat-ile", String(wGrupie.length)));
    sek.appendChild(naglowek);

    const tresc = el("div", "crm-kat-lista");
    sek.appendChild(tresc);
    if (!wGrupie.length) tresc.appendChild(el("div", "crm-pusto", "Brak zgłoszeń."));

    const wiecej = el("button", "crm-btn crm-btn-cichy crm-kat-wiecej") as HTMLButtonElement;
    wiecej.style.display = "none";
    tresc.appendChild(wiecej);
    let pokazane = 0;
    const dorysuj = () => {
      wGrupie.slice(pokazane, pokazane + ZGLOSZEN_NA_STRONE).forEach((t) => tresc.insertBefore(karta(t), wiecej));
      pokazane = Math.min(pokazane + ZGLOSZEN_NA_STRONE, wGrupie.length);
      const zostalo = wGrupie.length - pokazane;
      wiecej.textContent = `Pokaż kolejne (${Math.min(ZGLOSZEN_NA_STRONE, zostalo)} z ${zostalo})`;
      wiecej.style.display = zostalo > 0 ? "" : "none";
    };
    wiecej.onclick = dorysuj;

    // karty rysujemy dopiero przy pierwszym rozwinieciu sekcji
    const start = () => { if (!pokazane && wGrupie.length) dorysuj(); };
    if (sek.open) start();
    sek.addEventListener("toggle", () => { if (sek.open) start(); });

    lista.appendChild(sek);
  });
}

// ---------------------------------------------------------------- galeria firm

/** 1 support, 2-4 supporty, 5+ supportow (12-14 tez supportow) */
function supportyOdmiana(n: number): string {
  if (n === 1) return "1 support";
  const d = n % 10, s = n % 100;
  return `${n} ${d >= 2 && d <= 4 && (s < 12 || s > 14) ? "supporty" : "supportów"}`;
}

function firmyOdmiana(n: number): string {
  if (n === 1) return "1 firma";
  const d = n % 10, s = n % 100;
  return `${n} ${d >= 2 && d <= 4 && (s < 12 || s > 14) ? "firmy" : "firm"}`;
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

  // dodawanie firm tylko dla osob z MEScoStaff - jak w Power Apps; do sprawdzenia przycisk ukryty
  const nowa = przyciskIk("crm-btn crm-nf-otworz", "firma", "Nowa firma");
  nowa.style.display = "none";
  nowa.onclick = () => {
    root().innerHTML = "";
    ustawNaglowek("Nowa firma", "nav-companies");
    void formularzNowejFirmy(root(), () => void pokazFirmy(szukaj.value, wybor.value));
  };
  filtry.appendChild(nowa);
  czyPracownikMESco().then(
    (ok) => { if (ok) nowa.style.display = ""; },
    (e) => console.error("MEScoStaff:", e)
  );

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
        : `Największe zużycie TECSu (${firmyOdmiana(firmy.length)} w spisie)`;

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
      const b = chipRodzaju(o.etykieta);
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
  const mail = opcjaMaila(box, () => sprawdz());

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
    mail.pokaz(rodzaj);
    // dolaczony mail wystarczy jako tresc
    if (!ta.value.trim() && !mail.wlaczony) brak.push("treść");
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
      if (mail.wlaczony) Object.assign(pola, await mail.pola());
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

/** Edycja wlasnej notatki: data, osoba, rodzaj, tresc, follow-up. Pola maila zostaja bez zmian. */
function edycjaNotatki(nt: Nota, poZapisie: (nowa: Nota) => void, anuluj: () => void): HTMLElement {
  const box = el("div", "crm-nowe crm-demo");
  box.appendChild(poleIk("crm-demo-tytul", "pioro", "Edycja notatki"));

  const data = poleFormularza(box, "Data wydarzenia *", "date");
  data.value = nt.data ? nt.data.slice(0, 10) : dzisISO();
  const osoba = poleFormularza(box, "Osoba");
  osoba.value = nt.osoba;

  box.appendChild(el("label", "crm-label", "Rodzaj interakcji *"));
  const rodzaje = el("div", "crm-chipy crm-chipy-rowne");
  box.appendChild(rodzaje);
  let rodzaj = nt.rodzaj;
  let rodzajId = nt.rodzajId;
  getLookupOptions("activities", "Notes_types").then(
    (opcje) => opcje.forEach((op) => {
      const b = chipRodzaju(op.etykieta);
      if (op.id === rodzajId || (!rodzajId && op.etykieta === rodzaj)) {
        b.classList.add("crm-chip-akt");
        rodzajId = op.id;
      }
      b.onclick = () => {
        rodzaj = op.etykieta;
        rodzajId = op.id;
        rodzaje.querySelectorAll(".crm-chip-wybor").forEach((x) => x.classList.remove("crm-chip-akt"));
        b.classList.add("crm-chip-akt");
        sprawdz();
      };
      rodzaje.appendChild(b);
    }),
    (e) => rodzaje.appendChild(el("div", "crm-nowe-info", `Nie pobrano rodzajów: ${e?.message ?? e}`))
  );

  box.appendChild(el("label", "crm-label", "Notatka"));
  const ta = document.createElement("textarea");
  ta.className = "crm-textarea";
  ta.rows = 5;
  ta.value = nt.tresc;
  box.appendChild(ta);

  const fu = document.createElement("input");
  fu.type = "checkbox";
  fu.id = `fu-ed-${nt.id}`;
  fu.checked = !!nt.followUp;
  const fuLbl = el("label", "crm-check") as HTMLLabelElement;
  fuLbl.htmlFor = fu.id;
  fuLbl.append(fu, el("span", undefined, "Follow-up (przypomnienie)"));
  box.appendChild(fuLbl);
  const fuData = document.createElement("input");
  fuData.type = "date";
  fuData.className = "crm-input";
  fuData.value = nt.followUp ? nt.followUp.slice(0, 10) : plusDni(data.value, 7);
  fuData.style.display = fu.checked ? "" : "none";
  box.appendChild(fuData);
  fu.onchange = () => (fuData.style.display = fu.checked ? "" : "none");

  const brakujace = el("div", "crm-podpowiedz");
  box.appendChild(brakujace);
  const zapisz = el("button", "crm-btn", "Zapisz zmiany") as HTMLButtonElement;
  const wstecz = el("button", "crm-btn crm-btn-cichy", "Anuluj") as HTMLButtonElement;
  const komunikat = el("span", "crm-komunikat");
  const rzad = el("div", "crm-akcje");
  rzad.append(zapisz, wstecz, komunikat);
  box.appendChild(rzad);
  wstecz.onclick = anuluj;

  function sprawdz(): boolean {
    const brak: string[] = [];
    if (!data.value) brak.push("data");
    if (!rodzaj) brak.push("rodzaj");
    const ok = brak.length === 0;
    zapisz.disabled = !ok;
    brakujace.textContent = ok ? "" : `Uzupełnij: ${brak.join(", ")}.`;
    brakujace.style.display = ok ? "none" : "block";
    return ok;
  }
  data.onchange = () => sprawdz();
  sprawdz();

  zapisz.onclick = async () => {
    if (!sprawdz()) return;
    zapisz.disabled = true;
    komunikat.textContent = "Zapisywanie...";
    komunikat.className = "crm-komunikat";

    const followUp = fu.checked && fuData.value ? dzienSP(fuData.value) : null;
    const pola: Record<string, unknown> = {
      Data: dzienSP(data.value),
      Osoba: osoba.value.trim(),
      Rodzaj1: rodzaj,
      Totatka: ta.value,            // uwaga: literowka w nazwie kolumny
      Follow_up: followUp,          // null czysci przypomnienie
    };
    if (rodzajId) pola["TypeLookupId"] = rodzajId;

    try {
      await updateTicketFields(nt.id, pola, undefined, "Notes");
      poZapisie({
        ...nt,
        data: String(pola.Data ?? ""),
        osoba: osoba.value.trim(),
        rodzaj,
        rodzajId,
        tresc: ta.value,
        followUp: followUp ?? "",
      });
    } catch (e: any) {
      komunikat.textContent = e?.message ?? String(e);
      komunikat.className = "crm-komunikat crm-blad";
      zapisz.disabled = false;
    }
  };

  return box;
}

/** Karta notatki z koszem przy wlasnych - w zakladce Notatki i przy szansach. */
function kartaNotatki(nt: Nota, ja: string, poUsunieciu?: () => void): HTMLElement {
  const k = el("div", "crm-nota");
  const top = el("div", "crm-nota-top");
  top.appendChild(poleIk("crm-nota-data", "kalendarz", dataPL(nt.data), "Data"));

  // typ notatki w prawym rogu
  const prawo = el("div", "crm-nota-prawo");
  if (nt.rodzaj) {
    const { tekst, ik } = rodzajNotatki(nt.rodzaj);
    const typ = el("span", "crm-nota-typ");
    typ.title = tekst;
    typ.append(ikona(ik), el("span", undefined, tekst));
    prawo.appendChild(typ);
  }
  const moja = !!ja && nt.autor.trim().toLowerCase() === ja;
  const edytuj = el("button", "crm-usun crm-edytuj") as HTMLButtonElement;
  edytuj.appendChild(ikona("pioro"));
  edytuj.title = "Edytuj notatkę";
  const kosz = el("button", "crm-usun") as HTMLButtonElement;
  kosz.appendChild(ikona("kosz"));
  kosz.title = "Usuń notatkę";
  if (moja) prawo.append(edytuj, kosz);
  top.appendChild(prawo);
  k.appendChild(top);

  // edycja w miejscu karty; po zapisie karta rysowana od nowa z nowymi danymi
  edytuj.onclick = () => {
    const forma = edycjaNotatki(
      nt,
      (nowa) => forma.replaceWith(kartaNotatki(nowa, ja, poUsunieciu)),
      () => forma.replaceWith(k)
    );
    k.replaceWith(forma);
  };

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
      poUsunieciu?.();
    } catch (e: any) {
      tak.textContent = "Usuń";
      tak.disabled = false;
      nie.disabled = false;
      pytanie.appendChild(el("div", "crm-komunikat crm-blad", e?.message ?? String(e)));
    }
  };

  return k;
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

  lista.forEach((nt) => box.appendChild(kartaNotatki(nt, ja)));
}

async function zakladkaTecs(box: HTMLElement, s: HTMLElement, c: Company) {
  segment(box, ["Bieżące", "Archiwalne", "Demo"], tecsTryb, (v) => {
    tecsTryb = v as typeof tecsTryb;
    void pokazFirma(c);
  });

  // nowa szansa TECS - przycisk nad lista albo od razu po wygranej szansie
  const nowyBox = el("div");
  box.appendChild(nowyBox);
  const otworzFormTecs = (zSzansy: Szansa | null = null) => {
    nowyBtn.disabled = true;
    void formularzNowegoTecs(nowyBox, c, zSzansy, () => {
      tecsTryb = "Bieżące";
      void pokazFirma(c);
    }, () => {
      nowyBox.innerHTML = "";
      nowyBox.appendChild(nowyBtn);
      nowyBtn.disabled = false;
    });
  };
  const nowyBtn = przyciskIk("crm-btn crm-tecs-nowy", "licencja", "Nowa szansa TECS");
  nowyBtn.onclick = () => {
    nowyBox.innerHTML = "";
    otworzFormTecs();
  };
  if (otworzNowyTecs) {
    otworzNowyTecs = false;
    otworzFormTecs(nowyTecsZSzansy);
    nowyTecsZSzansy = null;
  } else nowyBox.appendChild(nowyBtn);

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
      ? wszystkie.filter((t) => ETAPY_TECS_BIEZACE.includes(t.etap))
      : tecsTryb === "Archiwalne"
      ? wszystkie.filter((t) => ETAPY_TECS_ARCHIWALNE.includes(t.etap))
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
    const wartEl = poleIk("crm-kafel-wartosc", ikWart, wart, "Wartość");
    k.appendChild(wartEl);

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

    if (!t.demo && t.guid) {
      const rzadP = el("div", "crm-akcje crm-szansa-akcje");
      const szcz = przyciskIk("crm-btn crm-btn-cichy", "flaga", "Szczegóły i etap");
      rzadP.appendChild(szcz);
      const panelSzcz = el("div");
      const pokazSzczegoly = () => {
        panelSzcz.appendChild(panelTecs(t, c));
        szcz.classList.add("crm-btn-wcisniety");
        otwartySzczegolTecs = t.id;
      };
      szcz.onclick = () => {
        if (panelSzcz.childElementCount) {
          panelSzcz.innerHTML = "";
          szcz.classList.remove("crm-btn-wcisniety");
          otwartySzczegolTecs = "";
        } else pokazSzczegoly();
      };
      if (otwartySzczegolTecs === t.id) pokazSzczegoly();
      const prod = przyciskIk("crm-btn crm-btn-cichy", "paczka", "Produkty i ceny");
      rzadP.appendChild(prod);
      k.appendChild(rzadP);
      k.appendChild(panelSzcz);
      const panel = el("div");
      k.appendChild(panel);
      prod.onclick = () => {
        if (panel.childElementCount) {
          panel.innerHTML = "";
          prod.classList.remove("crm-btn-wcisniety");
          return;
        }
        prod.classList.add("crm-btn-wcisniety");
        void produktyTecs(panel, t, c, (nowa) => {
          t.wartosc = nowa;
          if (!t.tp) wartEl.querySelector(".crm-ikona-txt")!.textContent = euro(nowa);
        });
      };
    }

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

// ---------------------------------------------------------------- szczegoly TECS (conTECSView)

// etapy TECS (Etap1, tekst) - sciezka zwykla i osobna dla Tech Poola, jak w Power Apps
const ETAPY_TECS = ["Przygotowanie oferty", "Oferta wysłana", "Zamówione", "Wystawiona faktura"];
const ETAPY_TP = ["Tech Pool", "Koniec Tech Poola"];
const REZYGNACJA_TECS = "Rezygnacja z TECS";
// "Wystawiona faktura" i etapy Tech Poola nie trafialy dotad do zadnej listy - dopisane
const ETAPY_TECS_BIEZACE = ["Przygotowanie oferty", "Oferta wysłana", "Tech Pool"];
const ETAPY_TECS_ARCHIWALNE = ["Oferta wysłana", "Zamówione", "Wystawiona faktura", REZYGNACJA_TECS, "Koniec Tech Poola"];

// panel szczegolow otwarty przed przerysowaniem zakladki (zmiana etapu, edycja)
let otwartySzczegolTecs = "";

const linkZPola = (v: unknown) => String((v as any)?.Url ?? v ?? "").trim();

/** Szczegoly TECS: kafle (wartosc, odbiorca, okres, follow-up), etap, uwagi, linki SF i edycja. */
function panelTecs(t: TecsRow, c: Company): HTMLElement {
  const box = el("div", "crm-szansa-panel crm-tecs-szcz");
  const komunikat = el("div", "crm-komunikat");

  const opis = [
    `Klient Ansys: ${String(t.pola.Numer_klienta ?? "").trim() || "—"}`,
    `Quote: ${String(t.pola.QuoteNumber ?? "").trim() || "—"}`,
  ];
  if (t.tp) opis.push("TECH POOL");
  box.appendChild(el("div", "crm-tecs-szcz-opis", opis.join("  ·  ")));

  // linki do Salesforce - tylko gdy sa
  const linki = el("div", "crm-akcje crm-szansa-akcje");
  ([["Opp w SF", t.pola.LinkdoOppsawSF], ["Quote w SF", t.pola.QuoteNumberLink]] as [string, unknown][]).forEach(([tekst, v]) => {
    const url = linkZPola(v);
    if (!url) return;
    const a = document.createElement("a");
    a.href = url;
    a.target = "_blank";
    a.rel = "noopener";
    a.className = "crm-btn crm-btn-cichy crm-tecs-link";
    a.textContent = `↗ ${tekst}`;
    linki.appendChild(a);
  });
  if (linki.childElementCount) box.appendChild(linki);

  // --- kafle ---
  const kafle = el("div", "crm-tecs-kafle");
  box.appendChild(kafle);
  const kafel = (ik: keyof typeof IKONY, kol: string, etykieta: string, wartosc: string, dopisek = "", kolDopisku = "") => {
    const k = el("div", "crm-tecs-kafel");
    const ikona_ = ikona(ik);
    ikona_.style.color = kol;
    k.appendChild(ikona_);
    const tresc = el("div", "crm-tecs-kafel-tresc");
    tresc.append(el("div", "crm-tecs-kafel-et", etykieta), el("div", "crm-tecs-kafel-war", wartosc));
    if (dopisek) {
      const d = el("div", "crm-tecs-kafel-dop", dopisek);
      if (kolDopisku) d.style.color = kolDopisku;
      tresc.appendChild(d);
    }
    k.appendChild(tresc);
    k.title = `${etykieta}: ${wartosc}${dopisek ? ` (${dopisek})` : ""}`;
    kafle.appendChild(k);
  };
  kafel("wykres", "#16A34A", t.tp ? "Wartość Tech Poola" : "Wartość TECS", euro(t.wartosc));
  kafel("osoba", "#2F6FD6", "Oferta do", t.odbiorca || "—");
  const dk = dniDo(t.koniec);
  kafel("kalendarz", "#7C3AED", t.tp ? "Okres Tech Poola" : "Okres TECS",
    `${dataPL(t.poczatek)} – ${dataPL(t.koniec)}`,
    dk === null ? "" : dk < 0 ? `zakończony ${-dk} dni temu` : `do końca ${dk} dni`,
    dk !== null && dk < 30 ? "#DC2626" : "");
  const df = dniDo(t.followUp);
  kafel("zegar", "#D97706", "Follow-up", t.followUp ? dataPL(t.followUp) : "—",
    df === null ? "nie ustawiono" : df < 0 ? "po terminie" : df === 0 ? "dzisiaj" : `za ${df} dni`,
    df !== null && df < 0 ? "#DC2626" : "");

  // --- etap: klik zapisuje od razu (z kontrola wspolbieznosci) ---
  box.appendChild(el("div", "crm-nf-sekcja", "Etap"));
  const kroki = el("div", "crm-szansa-etapy");
  box.appendChild(kroki);
  const sciezka = ETAPY_TP.includes(t.etap) ? ETAPY_TP : ETAPY_TECS;
  const biezacy = sciezka.indexOf(t.etap) + 1;
  const zapiszEtap = async (etap: string, b: HTMLButtonElement) => {
    b.disabled = true;
    komunikat.textContent = "Zapisywanie...";
    komunikat.className = "crm-komunikat";
    try {
      await updateTicketFields(t.id, { Etap1: etap }, t.etag || undefined, "Opps_TECS");
      // karta przechodzi do listy, na ktorej ten etap jest widoczny
      if (!ETAPY_TECS_BIEZACE.includes(etap) && tecsTryb === "Bieżące") tecsTryb = "Archiwalne";
      if (!ETAPY_TECS_ARCHIWALNE.includes(etap) && tecsTryb === "Archiwalne") tecsTryb = "Bieżące";
      otwartySzczegolTecs = t.id;
      await pokazFirma(c);
    } catch (e: any) {
      komunikat.textContent = e?.message ?? String(e);
      komunikat.className = "crm-komunikat crm-blad";
      b.disabled = false;
    }
  };
  sciezka.forEach((etap, i) => {
    const nr = i + 1;
    const ukonczony = nr < biezacy && t.etap !== REZYGNACJA_TECS;
    const b = el("button", "crm-szansa-krok", `${ukonczony ? "✓" : nr + "."} ${etap}`) as HTMLButtonElement;
    b.title = etap;
    b.classList.add(nr === biezacy ? "crm-krok-biez" : ukonczony ? "crm-krok-ok" : "crm-krok-przysz");
    b.onclick = () => { if (nr !== biezacy) void zapiszEtap(etap, b); };
    kroki.appendChild(b);
  });
  if (sciezka === ETAPY_TECS) {
    const rez = el("button", "crm-szansa-krok crm-krok-rezygnacja", REZYGNACJA_TECS) as HTMLButtonElement;
    rez.classList.toggle("crm-krok-rez-akt", t.etap === REZYGNACJA_TECS);
    rez.onclick = () => { if (t.etap !== REZYGNACJA_TECS) void zapiszEtap(REZYGNACJA_TECS, rez); };
    kroki.appendChild(rez);
  }

  // --- uwagi ---
  box.appendChild(el("div", "crm-nf-sekcja", "Uwagi"));
  const uwagi = String(t.pola.UWAGI ?? "").trim();
  box.appendChild(el("div", uwagi ? "crm-tecs-uwagi" : "crm-tecs-uwagi crm-tecs-uwagi-brak", uwagi || "Brak uwag"));

  // --- edycja danych TECS - ten sam formularz co nowy TECS ---
  const edycjaBox = el("div");
  const rzad = el("div", "crm-akcje");
  const edytuj = przyciskIk("crm-btn crm-btn-cichy", "pioro", "Edytuj dane TECS");
  const odnow = przyciskIk("crm-btn crm-btn-cichy", "wykres", "Renewal");
  odnow.title = "Nowy TECS na kolejny rok na podstawie tego";
  rzad.append(edytuj, odnow);
  box.append(rzad, edycjaBox, komunikat);
  odnow.onclick = () => {
    rzad.style.display = "none";
    void formularzNowegoTecs(edycjaBox, c, null, () => {
      tecsTryb = "Bieżące";
      void pokazFirma(c);
    }, () => {
      edycjaBox.innerHTML = "";
      rzad.style.display = "";
    }, null, t);
  };
  edytuj.onclick = () => {
    rzad.style.display = "none";
    void formularzNowegoTecs(edycjaBox, c, null, () => {
      otwartySzczegolTecs = t.id;
      void pokazFirma(c);
    }, () => {
      edycjaBox.innerHTML = "";
      rzad.style.display = "";
    }, t);
  };

  return box;
}

// ---------------------------------------------------------------- produkty TECS

const rodzajProduktuTecs = (p: ProduktTecs) =>
  p.tp ? "Tech Pool" : p.lease ? `Dzierżawa ${p.leaseDni} dni` : "TECS";

const r2 = (v: number) => Math.round(v * 100) / 100;

/** SUMA pozycji TECS (tvFormProdukt): cena x ilosc po rabacie + wartosc dzierzawy. Power Apps przy ↻ pomijal rabat. */
const sumaTecs = (cena: number, ile: number, rabat: number, leaseValue = 0) =>
  r2(cena * ile * (1 - (rabat || 0) / 100) + leaseValue);

/** Mnoznik krotkiej dzierzawy: do 62 dni x2, do 270 dni x1,35, dluzej x1. */
const mnoznikDzierzawy = (dni: number) => (dni <= 62 ? 2 : dni <= 270 ? 1.35 : 1);

/** LeaseValue: ilosc x Annual Lease / 365 x mnoznik x dni, po rabacie. */
const wartoscDzierzawy = (ile: number, annualLease: number, dni: number, rabat: number) =>
  r2(ile * annualLease / 365 * mnoznikDzierzawy(dni) * dni * (1 - (rabat || 0) / 100));

type RodzajTecs = "TECS" | "TP" | "Lease";

/**
 * Produkty i ceny TECS (conTECSProdukty z Power Apps): lista pozycji Oferty_TECS,
 * cena z cennika, usuwanie, przeliczenie wartosci i przepisanie z poprzedniego TECS-u.
 * Zmiany tylko dla osob z MEScoStaff, jak w Power Apps.
 */
async function produktyTecs(miejsce: HTMLElement, t: TecsRow, c: Company, nowaWartosc: (v: number) => void) {
  const box = el("div", "crm-tecs-produkty");
  miejsce.appendChild(box);
  const stan = el("div", "crm-komunikat", "Wczytywanie produktów...");
  box.appendChild(stan);

  let pracownik = false;
  let ceny = new Map<string, CenaTecs>();
  let pozycje: ProduktTecs[] = [];
  let poprzednie: ProduktTecs[] = [];
  try {
    const [p, cn, biez, stare] = await Promise.all([
      czyPracownikMESco().catch(() => false),
      getCenyTecs().catch((e) => {
        console.warn("cennik TECS:", e);
        return new Map<string, CenaTecs>();
      }),
      getByCompany("Oferty_TECS", t.guid, { pole: "IdentyfikatorTECS", top: 200 }),
      t.guidOld
        ? getByCompany("Oferty_TECS", t.guidOld, { pole: "IdentyfikatorTECS", top: 200 })
        : Promise.resolve([] as ListItem[]),
    ]);
    pracownik = p;
    ceny = cn;
    pozycje = biez.map(toProduktTecs);
    poprzednie = stare.map(toProduktTecs);
  } catch (e: any) {
    stan.textContent = `Błąd: ${e?.message ?? e}`;
    stan.className = "crm-komunikat crm-blad";
    return;
  }
  stan.remove();

  const lista = el("div", "crm-tecs-lista");
  const stopka = el("div", "crm-tecs-stopka");
  const komunikat = el("div", "crm-komunikat");
  const dodajBox = el("div", "crm-tecs-dodaj");
  box.append(lista, dodajBox, stopka, komunikat);

  if (pracownik) {
    const dodaj = przyciskIk("crm-btn crm-tecs-nowy", "paczka", "Dodaj produkt");
    dodajBox.appendChild(dodaj);
    dodaj.onclick = () => {
      dodaj.disabled = true;
      const forma = formularzProduktuTecs(t, c, ceny, null, (nowy) => {
        pozycje.push(nowy);
        forma.remove();
        dodaj.disabled = false;
        ok(`Dodano ${nowy.produkt}. Pamiętaj o „Przelicz wartość”.`);
        rysuj();
      }, () => {
        forma.remove();
        dodaj.disabled = false;
      });
      dodajBox.appendChild(forma);
    };
  }

  const blad = (e: any) => {
    komunikat.textContent = e?.message ?? String(e);
    komunikat.className = "crm-komunikat crm-blad";
  };
  const ok = (tekst: string) => {
    komunikat.textContent = tekst;
    komunikat.className = "crm-komunikat crm-ok";
  };

  const rysuj = () => {
    lista.innerHTML = "";
    if (!pozycje.length) lista.appendChild(el("div", "crm-pusto", "Brak produktów w tym TECS."));
    pozycje.forEach((p) => lista.appendChild(wiersz(p)));
    rysujStopke();
  };

  const wiersz = (p: ProduktTecs): HTMLElement => {
    const w = el("div", "crm-tecs-poz");
    const gora = el("div", "crm-tecs-poz-gora");
    gora.appendChild(poleIk("crm-tecs-poz-nazwa", "licencja", p.produkt || "(bez nazwy)", p.produkt));
    gora.appendChild(el("span", "crm-tecs-poz-suma", euro(p.suma)));
    w.appendChild(gora);

    // cena na czerwono = nizsza niz w cenniku, na zielono = zgodna
    const cennik = ceny.get(p.produkt.trim())?.tecs;
    const linia = el("div", "crm-tecs-poz-linia");
    linia.appendChild(el("span", undefined, `${rodzajProduktuTecs(p)} · ${p.ile} szt.`));
    const cena = el("span", "crm-tecs-cena", euro(p.cena));
    if (cennik !== undefined) {
      cena.title = `Cennik: ${euro(cennik)}`;
      cena.classList.add(p.cena < cennik ? "crm-tecs-cena-nizsza" : p.cena === cennik ? "crm-tecs-cena-ok" : "");
    } else cena.title = "Brak produktu w cenniku";
    linia.append(cena);
    if (p.rabat) linia.appendChild(el("span", undefined, `rabat ${p.rabat}%`));
    if (p.tpValue) linia.appendChild(el("span", undefined, `TP ${euro(p.tpValue)}`));
    if (p.leaseValue) linia.appendChild(el("span", undefined, `dzierżawa ${euro(p.leaseValue)}`));
    w.appendChild(linia);

    if (!pracownik) return w;
    const akcje = el("div", "crm-tecs-poz-akcje");
    w.appendChild(akcje);

    // ↻ cena TECS z cennika; dzierzawa i TP zostaja jak byly (licza sie z Annual Lease)
    {
      const sync = el("button", "crm-usun crm-edytuj") as HTMLButtonElement;
      sync.title = cennik === undefined ? "Brak produktu w cenniku" : `Pobierz cenę z cennika (${euro(cennik)})`;
      sync.textContent = "↻ cena z cennika";
      sync.disabled = cennik === undefined;
      sync.onclick = async () => {
        if (cennik === undefined) return;
        sync.disabled = true;
        try {
          const suma = sumaTecs(cennik, p.ile, p.rabat, p.lease ? p.leaseValue : 0);
          await updateTicketFields(p.id, { Cena_TECS: cennik, SUMA: suma }, undefined, "Oferty_TECS");
          Object.assign(p, { cena: cennik, suma });
          ok(`Zaktualizowano cenę ${p.produkt} z cennika (z rabatem ${p.rabat || 0}%). Pamiętaj o „Przelicz wartość”.`);
          rysuj();
        } catch (e) {
          blad(e);
          sync.disabled = false;
        }
      };
      akcje.appendChild(sync);
    }

    const edytuj = el("button", "crm-usun crm-edytuj") as HTMLButtonElement;
    edytuj.appendChild(ikona("pioro"));
    edytuj.title = "Edytuj produkt";
    akcje.appendChild(edytuj);
    edytuj.onclick = () => {
      const forma = formularzProduktuTecs(t, c, ceny, p, (nowy) => {
        Object.assign(p, nowy);
        ok(`Zapisano ${p.produkt}. Pamiętaj o „Przelicz wartość”.`);
        rysuj();
      }, () => forma.replaceWith(w));
      w.replaceWith(forma);
    };

    const kosz = el("button", "crm-usun") as HTMLButtonElement;
    kosz.appendChild(ikona("kosz"));
    kosz.title = "Usuń produkt";
    akcje.appendChild(kosz);
    kosz.onclick = () => {
      akcje.innerHTML = "";
      akcje.appendChild(el("span", "crm-tecs-pytanie", `Usunąć ${p.produkt}?`));
      const tak = el("button", "crm-btn crm-btn-ostrz", "Usuń") as HTMLButtonElement;
      const nie = el("button", "crm-btn crm-btn-cichy", "Anuluj") as HTMLButtonElement;
      akcje.append(tak, nie);
      nie.onclick = rysuj;
      tak.onclick = async () => {
        tak.disabled = nie.disabled = true;
        try {
          await deleteListItem("Oferty_TECS", p.id);
          pozycje = pozycje.filter((x) => x !== p);
          ok(`Usunięto ${p.produkt}. Pamiętaj o „Przelicz wartość”.`);
          rysuj();
        } catch (e) {
          blad(e);
          rysuj();
        }
      };
    };
    return w;
  };

  const rysujStopke = () => {
    stopka.innerHTML = "";
    const razem = Math.round(pozycje.reduce((a, p) => a + p.suma, 0) * 100) / 100;
    const zgodne = Math.round(razem) === Math.round(t.wartosc || 0);
    const kol = (etykieta: string, wartosc: string, klasa = "") => {
      const k = el("div", "crm-tecs-stopka-kol");
      k.append(el("div", "crm-tecs-stopka-et", etykieta), el("div", `crm-tecs-stopka-war ${klasa}`.trim(), wartosc));
      return k;
    };
    stopka.append(
      kol("Suma pozycji", euro(razem)),
      kol("Wartość TECS w systemie", euro(t.wartosc || 0), zgodne ? "crm-ok" : "crm-nf-uwaga")
    );
    const przelicz = przyciskIk("crm-btn crm-btn-cichy", "wykres", "Przelicz wartość");
    przelicz.disabled = zgodne;
    przelicz.title = zgodne ? "Wartość zgodna z sumą pozycji" : "Zapisz sumę pozycji jako wartość TECS";
    przelicz.onclick = async () => {
      przelicz.disabled = true;
      try {
        await updateTicketFields(t.id, { [await nazwaWewnetrzna("Opps_TECS", "Wartość")]: razem }, undefined, "Opps_TECS");
        nowaWartosc(razem);
        ok(`Przeliczono wartość TECS: ${euro(razem)}`);
        rysujStopke();
      } catch (e) {
        blad(e);
        przelicz.disabled = false;
      }
    };
    stopka.appendChild(przelicz);
  };

  rysuj();

  // --- z poprzedniego TECS-u ---
  if (!t.guidOld) return;
  box.appendChild(el("div", "crm-nf-sekcja", "Z poprzedniego TECS-u"));
  if (!poprzednie.length) {
    box.appendChild(el("div", "crm-pusto", "Poprzedni TECS nie ma produktów."));
    return;
  }
  poprzednie.forEach((p) => {
    const w = el("div", "crm-tecs-poz crm-tecs-stary");
    const gora = el("div", "crm-tecs-poz-gora");
    gora.appendChild(poleIk("crm-tecs-poz-nazwa", "licencja", p.produkt, p.produkt));
    w.appendChild(gora);
    w.appendChild(el("div", "crm-tecs-poz-linia", `${p.ile} szt. · do ${dataPL(p.koniec)}`));
    if (pracownik) {
      const przepisz = el("button", "crm-btn crm-btn-cichy", "Przepisz") as HTMLButtonElement;
      gora.appendChild(przepisz);
      przepisz.onclick = async () => {
        przepisz.disabled = true;
        // Power Apps przepisywal bez ceny i sumy - bierzemy aktualna cene z cennika (albo stara, gdy produktu brak)
        const cena = ceny.get(p.produkt.trim())?.tecs ?? p.cena;
        const suma = sumaTecs(cena, p.ile, p.rabat);
        try {
          const nowy = await createListItem("Oferty_TECS", {
            IdentyfikatorCompanies: c.guid,
            IdentyfikatorTECS: t.guid,
            ProducID: p.produkt,
            Liczba_licencji: p.ile,
            Pocz_x0105_tek_TECS: t.poczatek || null,
            Koniec_TECS: t.koniec || null,
            FirmaIDofertaTECS: p.firmaId,
            TECSOppsIDoferta: p.oppsId,
            Rabat: p.rabat,
            Cena_TECS: cena,
            SUMA: suma,
          });
          pozycje.push(toProduktTecs(nowy));
          ok(`Przepisano ${p.produkt} (${euro(suma)}). Pamiętaj o „Przelicz wartość”.`);
          rysuj();
          przepisz.textContent = "Przepisano";
        } catch (e) {
          blad(e);
          przepisz.disabled = false;
        }
      };
    }
    box.appendChild(w);
  });
}

/**
 * Dodanie / edycja pozycji Oferty_TECS (tvFormProdukt z Power Apps). Rodzaj TECS / Tech Pool / Dzierzawa,
 * wzory: SUMA = cena x ilosc x (1 - rabat) [+ LeaseValue], TPValue = ilosc x Annual Lease (bez rabatu).
 */
function formularzProduktuTecs(
  t: TecsRow,
  c: Company,
  ceny: Map<string, CenaTecs>,
  p: ProduktTecs | null,
  poZapisie: (nowy: ProduktTecs) => void,
  anuluj: () => void
): HTMLElement {
  const box = el("div", "crm-nowe crm-demo crm-tecs-forma");
  box.appendChild(poleIk("crm-demo-tytul", "paczka", p ? `Edycja: ${p.produkt}` : "Nowy produkt"));

  // --- rodzaj: domyslnie jak w Power Apps (TP gdy TECS jest TechPoolem) ---
  box.appendChild(el("label", "crm-label", "Rodzaj"));
  let rodzaj: RodzajTecs = p ? (p.tp ? "TP" : p.lease ? "Lease" : "TECS") : t.tp ? "TP" : "TECS";
  chipyWyboru<RodzajTecs>(box, [
    { tekst: "TECS", wartosc: "TECS" },
    { tekst: "Tech Pool", wartosc: "TP" },
    { tekst: "Dzierżawa", wartosc: "Lease" },
  ], rodzaj, (v) => {
    rodzaj = v;
    dniBox.style.display = v === "Lease" ? "" : "none";
    przelicz();
  });

  // --- produkt z cennika (datalist - wpisz lub wybierz) ---
  const produkt = poleFormularza(box, "Produkt *");
  produkt.placeholder = "zacznij wpisywać nazwę licencji…";
  const lista = document.createElement("datalist");
  lista.id = `cennik-${nowyGuid()}`;
  [...ceny.keys()].sort((a, b) => a.localeCompare(b, "pl")).forEach((k) => {
    const o = document.createElement("option");
    o.value = k;
    lista.appendChild(o);
  });
  box.appendChild(lista);
  produkt.setAttribute("list", lista.id);
  produkt.value = p?.produkt ?? "";

  const rzad = el("div", "crm-demo-okres");
  box.appendChild(rzad);
  const kol = (etykieta: string, wartosc: string) => {
    const k = el("div");
    k.appendChild(el("label", "crm-label", etykieta));
    const i = document.createElement("input");
    i.type = "text";
    i.inputMode = "decimal";
    i.className = "crm-input";
    i.value = wartosc;
    k.appendChild(i);
    rzad.appendChild(k);
    return i;
  };
  const ilosc = kol("Liczba licencji *", String(p?.ile ?? 1));
  const rabat = kol("Rabat %", String(p?.rabat ?? 0));
  const cena = poleFormularza(box, "Cena TECS / szt. (€)");
  cena.inputMode = "decimal";
  cena.value = p ? String(p.cena) : "";
  let cenaReczna = !!p;
  const dniBox = el("div");
  box.appendChild(dniBox);
  const dni = poleFormularza(dniBox, "Dni dzierżawy");
  dni.inputMode = "numeric";
  dni.value = String(p?.leaseDni || 30);
  dniBox.style.display = rodzaj === "Lease" ? "" : "none";

  const podglad = el("div", "crm-tecs-podglad");
  box.appendChild(podglad);
  const brakujace = el("div", "crm-podpowiedz");
  box.appendChild(brakujace);
  const zapisz = el("button", "crm-btn", p ? "Zapisz zmiany" : "Dodaj produkt") as HTMLButtonElement;
  const wstecz = el("button", "crm-btn crm-btn-cichy", "Anuluj") as HTMLButtonElement;
  const komunikat = el("span", "crm-komunikat");
  const akcje = el("div", "crm-akcje");
  akcje.append(zapisz, wstecz, komunikat);
  box.appendChild(akcje);
  wstecz.onclick = anuluj;

  const num = (i: HTMLInputElement) => {
    const v = Number(i.value.replace(/\s/g, "").replace(",", "."));
    return Number.isFinite(v) ? v : 0;
  };
  const wylicz = () => {
    const cn = ceny.get(produkt.value.trim());
    const q = num(ilosc), r = num(rabat), d = num(dni);
    const leaseValue = rodzaj === "Lease" ? wartoscDzierzawy(q, cn?.lease ?? 0, d, r) : 0;
    return {
      cn, q, r, d, leaseValue,
      tpValue: rodzaj === "TP" ? r2(q * (cn?.lease ?? 0)) : 0,
      suma: sumaTecs(num(cena), q, r, leaseValue),
    };
  };
  function przelicz() {
    const w = wylicz();
    podglad.innerHTML = "";
    const linia = (et: string, v: string) => podglad.appendChild(el("div", undefined, `${et}: ${v}`));
    linia("Suma", euro(w.suma));
    if (rodzaj === "TP") linia("Wartość TP", euro(w.tpValue));
    if (rodzaj === "Lease") linia(`Dzierżawa (×${mnoznikDzierzawy(w.d)})`, euro(w.leaseValue));
    if (w.cn) linia("Cennik TECS", euro(w.cn.tecs));
    sprawdz();
  }
  function sprawdz(): boolean {
    const brak: string[] = [];
    if (!produkt.value.trim()) brak.push("produkt");
    if (num(ilosc) <= 0) brak.push("liczba licencji");
    if (rodzaj === "Lease" && num(dni) <= 0) brak.push("dni dzierżawy");
    const ok = brak.length === 0;
    zapisz.disabled = !ok;
    brakujace.textContent = ok ? "" : `Uzupełnij: ${brak.join(", ")}.`;
    brakujace.style.display = ok ? "none" : "block";
    return ok;
  }
  // nowa pozycja: cena podpowiada sie z cennika, dopoki ktos jej nie wpisze recznie
  produkt.oninput = () => {
    const cn = ceny.get(produkt.value.trim());
    if (cn && !cenaReczna) cena.value = String(cn.tecs);
    przelicz();
  };
  cena.oninput = () => {
    cenaReczna = true;
    przelicz();
  };
  [ilosc, rabat, dni].forEach((i) => (i.oninput = przelicz));
  przelicz();

  zapisz.onclick = async () => {
    if (!sprawdz()) return;
    zapisz.disabled = true;
    komunikat.textContent = "Zapisywanie...";
    komunikat.className = "crm-komunikat";
    const w = wylicz();
    const pola: Record<string, unknown> = {
      ProducID: produkt.value.trim(),
      Liczba_licencji: w.q,
      Rabat: w.r,
      Cena_TECS: num(cena),
      SUMA: w.suma,
      Lease: rodzaj === "Lease",
      LeaseDays: rodzaj === "Lease" ? w.d : 0,
      LeaseValue: w.leaseValue,
      TP: rodzaj === "TP",
      TPValue: w.tpValue,
      FirmaIDofertaTECS: t.firmaId || firmaPelna?.skrocona || c.nazwa,
      TECSOppsIDoferta: t.nazwa,
      Koniec_TECS: t.koniec || null,
      Pocz_x0105_tek_TECS: t.poczatek || null,
      IdentyfikatorTECS: t.guid,
      IdentyfikatorCompanies: c.guid,
    };
    try {
      if (p) {
        await updateTicketFields(p.id, pola, undefined, "Oferty_TECS");
        poZapisie({
          ...p, produkt: String(pola.ProducID), ile: w.q, rabat: w.r, cena: num(cena), suma: w.suma,
          lease: rodzaj === "Lease", leaseDni: rodzaj === "Lease" ? w.d : 0, leaseValue: w.leaseValue,
          tp: rodzaj === "TP", tpValue: w.tpValue,
        });
      } else {
        poZapisie(toProduktTecs(await createListItem("Oferty_TECS", pola)));
      }
    } catch (e: any) {
      komunikat.textContent = e?.message ?? String(e);
      komunikat.className = "crm-komunikat crm-blad";
      zapisz.disabled = false;
    }
  };

  return box;
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

/** iso + lata i dni, np. koniec TECS = poczatek + 1 rok - 1 dzien (jak DateAdd w Power Apps). */
const plusLatDni = (iso: string, lata: number, dni: number) => {
  const d = new Date((iso || dzisISO()) + "T12:00:00");
  d.setFullYear(d.getFullYear() + lata);
  d.setDate(d.getDate() + dni);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
};

/** Nowa szansa TECS w Opps_TECS (conTECSForm z Power Apps), Etap1 = "Przygotowanie oferty". */
async function formularzNowegoTecs(
  miejsce: HTMLElement,
  c: Company,
  zSzansy: Szansa | null,
  poZapisie: () => void,
  anuluj: () => void,
  edycja: TecsRow | null = null,
  renewal: TecsRow | null = null
) {
  // renewal: nazwa jak w Power Apps z FirmaID poprzedniego TECS-u
  const skrocona = (renewal?.firmaId) || firmaPelna?.skrocona || c.nazwa;
  const box = el("div", "crm-nowe crm-demo");
  miejsce.appendChild(box);
  box.appendChild(poleIk("crm-demo-tytul", renewal ? "wykres" : "dokument",
    edycja ? "Edycja szansy TECS" : renewal ? "Renewal TECS" : "Nowa szansa TECS"));
  box.appendChild(el("div", "crm-nowe-info",
    (firmaPelna?.pelna || skrocona) + (zSzansy ? ` · z wygranej szansy: ${zSzansy.nazwa}` : "")
    + (renewal ? ` · na podstawie: ${renewal.nazwa}` : "")));
  if (renewal) {
    const st = el("div", "crm-tecs-poprzedni");
    st.append(el("div", "crm-tecs-kafel-et", "POPRZEDNI TECS"), el("div", undefined,
      `Okres: ${dataPL(renewal.poczatek)} – ${dataPL(renewal.koniec)} · Wartość: ${euro(renewal.wartosc)} · ` +
      `Etap: ${renewal.etap || "—"}${renewal.tp ? " · Tech Pool" : ""}`));
    box.appendChild(st);
  }

  // --- okres ---
  box.appendChild(el("div", "crm-nf-sekcja", "Okres TECS"));
  const okres = el("div", "crm-demo-okres");
  box.appendChild(okres);
  const kolumna = (rodzic: HTMLElement, etykieta: string) => {
    const k = el("div");
    k.appendChild(el("label", "crm-label", etykieta));
    const i = document.createElement("input");
    i.type = "date";
    i.className = "crm-input";
    k.appendChild(i);
    rodzic.appendChild(k);
    return i;
  };
  const poczatek = kolumna(okres, "Początek TECS *");
  const koniec = kolumna(okres, "Koniec TECS *");
  const follow = poleFormularza(box, "Data follow-up", "date");
  poczatek.value = dzisISO();
  koniec.value = plusLatDni(poczatek.value, 1, -1);
  follow.value = plusLatDni(poczatek.value, 1, -92);

  // nazwa "Q<kwartal konca>_<rok konca>_<skrocona>" - podpowiadana, dopoki ktos jej nie zmieni
  const nazwa = poleFormularza(box, "Nazwa szansy *");
  let nazwaReczna = false;
  const podpowiedzNazwy = () => {
    if (nazwaReczna || !koniec.value) return;
    const [rok, mies] = koniec.value.split("-").map(Number);
    nazwa.value = `Q${Math.ceil(mies / 3)}_${rok}_${skrocona}`;
  };
  podpowiedzNazwy();
  nazwa.oninput = () => {
    nazwaReczna = true;
    sprawdz();
  };
  poczatek.onchange = () => {
    if (poczatek.value) {
      koniec.value = plusLatDni(poczatek.value, 1, -1);
      follow.value = plusLatDni(poczatek.value, 1, -92);
    }
    podpowiedzNazwy();
    sprawdz();
  };
  koniec.onchange = () => {
    podpowiedzNazwy();
    sprawdz();
  };

  const wartosc = poleFormularza(box, "Wartość – End User Price (€) *");
  wartosc.inputMode = "decimal";
  wartosc.placeholder = "0";
  let wartoscReczna = false;
  const zmianaWartosci = el("div", "crm-komunikat crm-nf-info");
  if (renewal?.wartosc) box.appendChild(zmianaWartosci);
  const opiszZmiane = () => {
    if (!renewal?.wartosc) return;
    const v = Number(wartosc.value.replace(/\s/g, "").replace(",", "."));
    const proc = ((Number.isFinite(v) ? v : 0) / renewal.wartosc - 1) * 100;
    zmianaWartosci.textContent = `zmiana: ${proc >= 0 ? "+" : ""}${proc.toFixed(1).replace(".", ",")}% względem poprzedniego`;
  };
  wartosc.oninput = () => {
    wartoscReczna = true;
    opiszZmiane();
    sprawdz();
  };

  // --- szczegoly ---
  box.appendChild(el("div", "crm-nf-sekcja", "Szczegóły"));
  const klient = poleFormularza(box, "Numer klienta Ansys");
  klient.placeholder = "np. 1234567";

  const tp = document.createElement("input");
  tp.type = "checkbox";
  tp.id = `tecs-tp-${nowyGuid()}`;
  const tpLbl = el("label", "crm-check") as HTMLLabelElement;
  tpLbl.htmlFor = tp.id;
  tpLbl.append(tp, el("span", undefined, "TechPool"));
  box.appendChild(tpLbl);

  box.appendChild(el("label", "crm-label", "Opiekun *"));
  const opiekun = document.createElement("select");
  opiekun.className = "crm-select";
  box.appendChild(opiekun);

  box.appendChild(el("label", "crm-label", "Odbiorca oferty"));
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

  const linkSF = poleFormularza(box, "Link do Oppsa w SF");
  linkSF.placeholder = "https://…";

  // --- quote ---
  box.appendChild(el("div", "crm-nf-sekcja", "Quote"));
  const quote = poleFormularza(box, "Numer Quota");
  const quoteLink = poleFormularza(box, "Link do Quota");
  quoteLink.placeholder = "https://…";
  box.appendChild(el("label", "crm-label", "Uwagi"));
  const uwagi = document.createElement("textarea");
  uwagi.className = "crm-textarea";
  uwagi.rows = 3;
  uwagi.placeholder = "Szczegóły oferty, licencje, ustalenia…";
  box.appendChild(uwagi);

  // --- produkty z wygranej szansy (Oferty po IdentyfikatorOpps) - zaznaczone trafia do Oferty_TECS ---
  type Propozycja = { produkt: string; ile: number; typ: string; cena: number; rabat: number; wybrana: boolean };
  let propozycje: Propozycja[] = [];
  const produktyBox = el("div");
  const zrodloProduktow = zSzansy?.guid ? "szansy" : renewal?.guid ? "poprzedniego TECS-u" : "";
  if (zrodloProduktow) {
    box.appendChild(el("div", "crm-nf-sekcja", renewal ? "Licencje z poprzedniego TECS-u" : "Produkty z szansy"));
    box.appendChild(produktyBox);
    produktyBox.appendChild(el("div", "crm-komunikat", `Wczytywanie produktów ${zrodloProduktow}...`));
  }
  const sumaPropozycji = () =>
    r2(propozycje.filter((x) => x.wybrana).reduce((a, x) => a + sumaTecs(x.cena, x.ile, x.rabat), 0));
  const rysujPropozycje = () => {
    produktyBox.innerHTML = "";
    if (!propozycje.length) {
      produktyBox.appendChild(el("div", "crm-pusto", renewal ? "Poprzedni TECS nie ma licencji." : "Szansa nie ma produktów."));
      return;
    }
    produktyBox.appendChild(el("div", "crm-nowe-info", renewal
      ? "Zaznaczone przepiszemy do nowego TECS-u z aktualną ceną z cennika i dotychczasowym rabatem."
      : "Zaznaczone dodamy do TECS-u z ceną TECS z cennika (bez rabatu – zmienisz później w „Produkty i ceny”)."));
    propozycje.forEach((x, i) => {
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.id = `tecs-prop-${i}-${nowyGuid()}`;
      cb.checked = x.wybrana;
      cb.onchange = () => {
        x.wybrana = cb.checked;
        if (!wartoscReczna) wartosc.value = String(sumaPropozycji());
        if (renewal) sumaRenewal.textContent = `Suma zaznaczonych licencji wg cennika: ${euro(sumaPropozycji())}`;
        sprawdz();
      };
      const lbl = el("label", "crm-check crm-tecs-prop") as HTMLLabelElement;
      lbl.htmlFor = cb.id;
      lbl.append(cb, el("span", "crm-tecs-prop-nazwa", `${x.produkt} ×${x.ile}`),
        el("span", "crm-tecs-prop-cena",
          `${x.typ}${x.rabat ? ` · −${x.rabat}%` : ""} · ${x.cena ? euro(sumaTecs(x.cena, x.ile, x.rabat)) : "brak w cenniku"}`));
      produktyBox.appendChild(lbl);
    });
    if (renewal) {
      sumaRenewal.textContent = `Suma zaznaczonych licencji wg cennika: ${euro(sumaPropozycji())}`;
      produktyBox.appendChild(sumaRenewal);
    }
  };
  const sumaRenewal = el("div", "crm-komunikat crm-nf-info");

  // --- przyciski ---
  const brakujace = el("div", "crm-podpowiedz");
  box.appendChild(brakujace);
  const zapisz = el("button", "crm-btn", edycja ? "Zapisz zmiany" : renewal ? "Utwórz renewal" : "Zapisz szansę TECS") as HTMLButtonElement;
  const wstecz = el("button", "crm-btn crm-btn-cichy", "Anuluj") as HTMLButtonElement;
  const komunikat = el("span", "crm-komunikat");
  const rzad = el("div", "crm-akcje");
  rzad.append(zapisz, wstecz, komunikat);
  box.appendChild(rzad);
  wstecz.onclick = anuluj;

  const liczba = () => Number(wartosc.value.replace(/\s/g, "").replace(",", "."));
  function sprawdz(): boolean {
    const brak: string[] = [];
    if (!nazwa.value.trim()) brak.push("nazwa");
    if (!poczatek.value) brak.push("początek");
    if (!koniec.value) brak.push("koniec");
    else if (poczatek.value && koniec.value < poczatek.value) brak.push("koniec przed początkiem");
    if (!opiekun.value) brak.push("opiekun");
    if (!wartosc.value.trim() || !Number.isFinite(liczba())) brak.push("wartość");
    const ok = brak.length === 0;
    zapisz.disabled = !ok;
    brakujace.textContent = ok ? "" : `Uzupełnij: ${brak.join(", ")}.`;
    brakujace.style.display = ok ? "none" : "block";
    return ok;
  }
  opiekun.onchange = () => sprawdz();
  sprawdz();

  // --- dane do podpowiedzi: opiekunowie, kontakty, ostatni numer klienta Ansys ---
  let ja = "";
  let kontakty: Contact[] = [];
  let wybrana: Contact | null = null;
  try {
    const [osoby, mojaNazwa, kt, tecsy, oferty, ceny] = await Promise.all([
      getOpiekunowie().catch(() => [] as string[]),
      getMojaNazwa().catch(() => ""),
      getContactsByCompany(c.guid).catch(() => [] as ListItem[]),
      getByCompany("Opps_TECS", c.guid, { top: 200 }).catch(() => [] as ListItem[]),
      zSzansy?.guid
        ? getByCompany("Oferty", zSzansy.guid, { pole: "IdentyfikatorOpps", top: 200 }).catch((e) => {
          console.warn("Oferty szansy:", e);
          return [] as ListItem[];
        })
        : Promise.resolve([] as ListItem[]),
      zSzansy?.guid || renewal?.guid
        ? getCenyTecs().catch(() => new Map<string, CenaTecs>())
        : Promise.resolve(new Map<string, CenaTecs>()),
    ]);
    // TECS dotyczy licencji kupionych na wlasnosc (Paid-up) - dzierzawa i AEC/ALH maja serwis w cenie, wiec domyslnie odznaczone
    propozycje = oferty.map((it) => {
      const f = it.fields as any;
      const produkt = String(f.ProductID ?? "").trim();
      const typ = f.LeaseDays === true ? "Lease" : f.AECYN === true ? "AEC" : f.ALHYN === true ? "ALH" : "Paid-up";
      return { produkt, ile: Number(f.Liczba_licencji) || 1, typ, cena: ceny.get(produkt)?.tecs ?? 0, rabat: 0, wybrana: typ === "Paid-up" };
    }).filter((x) => x.produkt);
    if (renewal?.guid) {
      // jak "Przepisz": licencje poprzedniego TECS-u, cena z cennika (albo stara), rabat bez zmian
      const stare = await getByCompany("Oferty_TECS", renewal.guid, { pole: "IdentyfikatorTECS", top: 200 })
        .catch(() => [] as ListItem[]);
      propozycje = stare.map(toProduktTecs).filter((x) => x.produkt).map((x) => ({
        produkt: x.produkt, ile: x.ile || 1, typ: rodzajProduktuTecs(x),
        cena: ceny.get(x.produkt.trim())?.tecs ?? x.cena, rabat: x.rabat, wybrana: true,
      }));
    }
    if (zrodloProduktow) {
      rysujPropozycje();
      if (!wartoscReczna && sumaPropozycji()) wartosc.value = String(sumaPropozycji());
    }
    ja = mojaNazwa;
    const lista = ja && !osoby.includes(ja) ? [ja, ...osoby] : osoby;
    ["", ...lista].forEach((o) => {
      const opt = document.createElement("option");
      opt.value = o;
      opt.textContent = o || "wybierz…";
      opt.selected = o === ja;
      opiekun.appendChild(opt);
    });
    kontakty = kt.map(toContact).filter((x) => x.aktywny)
      .sort((a, b) => a.imieNazwisko.localeCompare(b.imieNazwisko, "pl"));
    kontakty.forEach((x, i) => {
      const opt = document.createElement("option");
      opt.value = String(i);
      opt.textContent = x.imieNazwisko + (x.stanowisko ? ` · ${x.stanowisko}` : "");
      osobaSel.appendChild(opt);
    });
    // jak w Power Apps: numer klienta z najnowszego TECS-u firmy, ktory go ma
    const zNumerem = tecsy
      .filter((i) => String((i.fields as any).Numer_klienta ?? "").trim())
      .sort((a, b) => Number(b.id) - Number(a.id))[0];
    if (zNumerem) klient.value = String((zNumerem.fields as any).Numer_klienta).trim();
  } catch (e: any) {
    box.insertBefore(el("div", "crm-komunikat crm-blad", `Nie pobrano podpowiedzi: ${e?.message ?? e}`), brakujace);
  }
  osobaSel.onchange = () => {
    wybrana = osobaSel.value ? kontakty[Number(osobaSel.value)] : null;
    if (wybrana) odbiorca.value = wybrana.imieNazwisko;
  };

  // edycja: pola z zapisanego TECS-u zamiast podpowiedzi
  if (edycja) {
    const f = edycja.pola;
    const dzien = (v: unknown) => String(v ?? "").slice(0, 10);
    poczatek.value = dzien(f.Pocz_x0105_tek_TECS);
    koniec.value = dzien(f.Koniec_TECS);
    follow.value = dzien(f.Data_Follow_Up);
    nazwa.value = edycja.nazwa;
    nazwaReczna = true;
    wartosc.value = String(edycja.wartosc || "");
    wartoscReczna = true;
    klient.value = String(f.Numer_klienta ?? "");
    tp.checked = edycja.tp;
    odbiorca.value = edycja.odbiorca;
    linkSF.value = linkZPola(f.LinkdoOppsawSF);
    quote.value = String(f.QuoteNumber ?? "");
    quoteLink.value = linkZPola(f.QuoteNumberLink);
    uwagi.value = String(f.UWAGI ?? "");
    const kto = String(f.TECScreator ?? "").trim();
    if (kto) {
      if (!Array.from(opiekun.options).some((o) => o.value === kto)) {
        const opt = document.createElement("option");
        opt.value = opt.textContent = kto;
        opiekun.appendChild(opt);
      }
      opiekun.value = kto;
    }
  }

  // renewal (conTECSRenewal): wszystko +1 rok, wartosc +4%, link SF i quote od nowa
  if (renewal) {
    const f = renewal.pola;
    const rokPozniej = (v: unknown) => (String(v ?? "") ? plusLatDni(String(v).slice(0, 10), 1, 0) : "");
    poczatek.value = rokPozniej(f.Pocz_x0105_tek_TECS);
    koniec.value = rokPozniej(f.Koniec_TECS);
    follow.value = rokPozniej(f.Data_Follow_Up);
    podpowiedzNazwy();
    wartosc.value = renewal.wartosc ? String(r2(renewal.wartosc * 1.04)) : "";
    wartoscReczna = true;
    opiszZmiane();
    klient.value = String(f.Numer_klienta ?? "");
    tp.checked = renewal.tp;
    odbiorca.value = renewal.odbiorca;
    uwagi.value = String(f.UWAGI ?? "");
  }
  sprawdz();

  zapisz.onclick = async () => {
    if (!sprawdz()) return;
    zapisz.disabled = true;
    komunikat.textContent = "Zapisywanie...";
    komunikat.className = "crm-komunikat";

    try {
      // pole Osoba wymaga Id uzytkownika witryny - mail wybranego opiekuna
      const mail = opiekun.value === ja
        ? await getMojMail()
        : (await getSerwisanci().catch(() => [])).find((x) => x.nazwa === opiekun.value)?.mail ?? "";
      const opiekunId = mail ? await getSiteUserId(mail) : null;
      // w edycji niezmieniony opiekun zostaje jak byl - nie wymagamy, zeby byl na liscie pracownikow
      const bezZmianyOpiekuna = !!edycja && opiekun.value === String(edycja.pola.TECScreator ?? "").trim();
      if (opiekunId === null && !bezZmianyOpiekuna) {
        throw new Error(`nie znaleziono użytkownika witryny dla „${opiekun.value}” – pole Opiekun jest wymagane`);
      }

      const pola: Record<string, unknown> = {
        Nazwaszansysprzeda_x017c_y_TECS: nazwa.value.trim(),
        Numer_klienta: klient.value.trim(),
        Pocz_x0105_tek_TECS: dzienSP(poczatek.value),
        Koniec_TECS: dzienSP(koniec.value),
        Data_Follow_Up: dzienSP(follow.value),
        Warto_x015b__x0107_: liczba(),
        Odbiorca_oferty: odbiorca.value.trim(),
        TP: tp.checked,
        QuoteNumber: quote.value.trim(),
        UWAGI: uwagi.value,
        FirmaID: skrocona,
        Etap1: "Przygotowanie oferty",
        TECScreator: opiekun.value,
        TECSGUID: nowyGuid(),
        IdentyfikatorCompanies: c.guid,
      };
      if (opiekunId !== null) pola.OpiekunLookupId = String(opiekunId);
      if (renewal) {
        pola.TECSGUID_OLD = renewal.guid;
        pola.IdentyfikatorCompanies = renewal.pola.IdentyfikatorCompanies || c.guid;
      }
      const guidTecs = String(pola.TECSGUID);

      // edycja: bez nowego GUID, etapu i powiazania z firma; zapis z kontrola wspolbieznosci
      if (edycja) {
        ["TECSGUID", "Etap1", "IdentyfikatorCompanies", "FirmaID"].forEach((k) => delete pola[k]);
        const linkiEd: Record<string, unknown> = {};
        if (linkSF.value.trim() !== linkZPola(edycja.pola.LinkdoOppsawSF)) linkiEd.LinkdoOppsawSF = linkSF.value.trim() || null;
        if (quoteLink.value.trim() !== linkZPola(edycja.pola.QuoteNumberLink)) linkiEd.QuoteNumberLink = quoteLink.value.trim() || null;
        try {
          await updateTicketFields(edycja.id, { ...pola, ...linkiEd }, edycja.etag || undefined, "Opps_TECS");
        } catch (e: any) {
          if (!Object.keys(linkiEd).length || /zmienił/.test(e?.message ?? "")) throw e;
          console.warn("Opps_TECS z linkami odrzucone, zapis bez linkow:", e);
          await updateTicketFields(edycja.id, pola, edycja.etag || undefined, "Opps_TECS");
          komunikat.textContent = "Zapisano bez linków (SF / Quote) – uzupełnij je w Power Apps.";
          komunikat.className = "crm-komunikat crm-nf-uwaga";
          setTimeout(poZapisie, 2500);
          return;
        }
        poZapisie();
        return;
      }
      // kolumny Hyperlink - Graph bywa na nie oporny, wiec przy bledzie zapisujemy bez linkow
      const linki: Record<string, unknown> = {};
      if (linkSF.value.trim()) linki.LinkdoOppsawSF = linkSF.value.trim();
      if (quoteLink.value.trim()) linki.QuoteNumberLink = quoteLink.value.trim();

      let bezLinkow = false;
      let nowy: ListItem;
      try {
        nowy = await createListItem("Opps_TECS", { ...pola, ...linki });
      } catch (e) {
        if (!Object.keys(linki).length) throw e;
        console.warn("Opps_TECS z linkami odrzucone, zapis bez linkow:", e);
        nowy = await createListItem("Opps_TECS", pola);
        bezLinkow = true;
      }
      // po odswiezeniu listy nowy TECS od razu z otwartymi szczegolami
      otwartySzczegolTecs = String(nowy.id ?? "");
      // produkty z szansy - TECS juz jest, wiec blad pojedynczej pozycji nie cofa zapisu
      const uwagiZapisu: string[] = [];
      if (bezLinkow) uwagiZapisu.push("bez linków (SF / Quote) – uzupełnij je w Power Apps");
      const doDodania = propozycje.filter((x) => x.wybrana);
      let dodane = 0;
      for (const x of doDodania) {
        komunikat.textContent = `Dodawanie produktów: ${x.produkt}...`;
        try {
          await createListItem("Oferty_TECS", {
            ProducID: x.produkt,
            Liczba_licencji: x.ile,
            Cena_TECS: x.cena,
            Rabat: x.rabat,
            SUMA: sumaTecs(x.cena, x.ile, x.rabat),
            TP: tp.checked,
            TPValue: 0,
            Lease: false,
            LeaseDays: 0,
            LeaseValue: 0,
            FirmaIDofertaTECS: skrocona,
            TECSOppsIDoferta: String(pola.Nazwaszansysprzeda_x017c_y_TECS),
            Pocz_x0105_tek_TECS: pola.Pocz_x0105_tek_TECS,
            Koniec_TECS: pola.Koniec_TECS,
            IdentyfikatorTECS: guidTecs,
            IdentyfikatorCompanies: c.guid,
          });
          dodane++;
        } catch (e) {
          console.error("produkt TECS z szansy:", x.produkt, e);
        }
      }
      if (dodane < doDodania.length) uwagiZapisu.push(`dodano ${dodane} z ${doDodania.length} produktów – resztę dodaj w „Produkty i ceny”`);

      if (uwagiZapisu.length) {
        // komunikat zostaje widoczny chwile przed przerysowaniem listy
        komunikat.textContent = `Zapisano TECS, ale: ${uwagiZapisu.join("; ")}.`;
        komunikat.className = "crm-komunikat crm-nf-uwaga";
        setTimeout(poZapisie, 4000);
      } else poZapisie();
    } catch (e: any) {
      komunikat.textContent = e?.message ?? String(e);
      komunikat.className = "crm-komunikat crm-blad";
      zapisz.disabled = false;
    }
  };
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
        guid, firmaId: skrocona, odbiorca: odbiorca.value.trim(), guidOld: "",
        followUp: "", etag: "", pola: {},
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

async function zakladkaSzanse(
  box: HTMLElement,
  s: HTMLElement,
  c: Company,
  odswiezWidok: () => void = () => void pokazFirma(c)
) {
  segment(box, ["Aktywne", "Zamknięte"], szanseZamkniete ? "Zamknięte" : "Aktywne", (v) => {
    szanseZamkniete = v === "Zamknięte";
    odswiezWidok();
  });

  // notatki firmy jednym zapytaniem - przy szansach grupujemy je po Opps_ID
  const notatki = getByCompany("Notes", c.guid, { top: 500 }).then((it) => it.map(toNota));
  notatki.catch((e) => console.error("notatki szans:", e));
  const ja = getMojaNazwa().then((n) => n.trim().toLowerCase(), () => "");

  const items = await getByCompany("Opps", c.guid, { top: 100 });
  const lista = items
    .map(toSzansa)
    .filter((o) => (szanseZamkniete ? o.status.startsWith("Zamknięty") : o.status === "Aktywny"));

  s.textContent = `Szanse — ${lista.length}`;
  if (!lista.length) return pusto(box, szanseZamkniete ? "Brak zamkniętych szans." : "Brak aktywnych szans.");

  lista.forEach((o) => box.appendChild(kartaSzansy(o, c, notatki, ja)));
}

// --- status i etapy szansy (DetailScreen z Power Apps) ---

const STATUSY_SZANSY = [
  { v: "Aktywny", krotko: "Aktywny", kol: "#5B6CDB", tlo: "#EEF0FC" },
  { v: "Zamknięty wygrany", krotko: "Wygrany", kol: "#16A34A", tlo: "#E6F6EE" },
  { v: "Zamknięty przegrany", krotko: "Przegrany", kol: "#DC2626", tlo: "#FEE2E2" },
];

// Etap1 to tekst - dokladne wartosci jak w Power Apps, kolejnosc wyznacza numer kroku
const ETAPY_SZANSY: [string, string][] = [
  ["Pierwszy kontakt", "Kontakt"],
  ["Discovery", "Discovery"],
  ["Oferta", "Oferta"],
  ["Spotkanie techniczne", "Spotk. tech."],
  ["Demo/Benchmark", "Demo"],
  ["Decyzja", "Decyzja"],
  ["Rewizja oferty", "Rewizja"],
  ["Wdrożenie/Realizacja", "Wdrożenie"],
];

const POWODY_PRZEGRANEJ: [string, string][] = [
  ["DR_Odrzucony", "DR odrzucony"],
  ["Konkurencja", "Konkurencja"],
  ["BrakDofinansowania", "Brak dofinansowania"],
];
const NOTATKA_PRZEGRANEJ = "ZamkniętyOppsNotes"; // nazwa wyswietlana - wewnetrzna ma _x0119_

// pola "Dane etapu" (DetailScreen, zakladki 1-8); klucze jak w Power Apps, wewnetrzne ustala nazwaWewnetrzna
type TypPolaEtapu = "date" | "text" | "multiline" | "number" | "bool" | "link";
const POLA_ETAPOW: { klucz: string; etykieta: string; typ: TypPolaEtapu }[][] = [
  [
    { klucz: "FTDate", etykieta: "Data spotkania", typ: "date" },
    { klucz: "TerminDR", etykieta: "Termin DR z SF", typ: "date" },   // display FTDR_date
  ],
  [
    { klucz: "DISCH", etykieta: "Champion", typ: "text" },
    { klucz: "DISB", etykieta: "Budżet (EUR)", typ: "number" },
    { klucz: "DISP", etykieta: "Pain", typ: "multiline" },
  ],
  [
    { klucz: "OFP", etykieta: "Czy oferta została wysłana?", typ: "bool" },
    { klucz: "OFSF", etykieta: "Czy zrobiłeś quote w SF?", typ: "bool" },
    { klucz: "QuoteNumber", etykieta: "Numer quota", typ: "text" },
    { klucz: "OFDate", etykieta: "Data spotkania", typ: "date" },
    { klucz: "QuoteNumberLink", etykieta: "Link do quota", typ: "link" },
  ],
  [
    { klucz: "TMFU", etykieta: "Data follow-up", typ: "date" },
    { klucz: "TMEB", etykieta: "Economic Buyer – Eb", typ: "text" },
    { klucz: "TMDP", etykieta: "Proces decyzji – Dp", typ: "multiline" },
    { klucz: "TMDC", etykieta: "Kryteria – Dc", typ: "multiline" },
    { klucz: "TMM", etykieta: "Metryki – M", typ: "multiline" },
  ],
  [
    { klucz: "DBDE", etykieta: "Czy będzie Demo/Benchmark?", typ: "bool" },
    { klucz: "DBDate", etykieta: "Data demo / benchmarku", typ: "date" },
  ],
  [
    { klucz: "DSM", etykieta: "Spotkanie podsumowujące", typ: "date" },
    { klucz: "DEB", etykieta: "Potwierdzenie Eb", typ: "multiline" },
  ],
  [
    { klucz: "RTL", etykieta: "Czy potrzebne są szkolenia?", typ: "multiline" },
    { klucz: "RHW", etykieta: "Czy potrzebny jest sprzęt?", typ: "multiline" },
    { klucz: "RCR", etykieta: "Czy będzie zmiana konfiguracji?", typ: "multiline" },
  ],
  [
    { klucz: "WRDL", etykieta: "Data deadline", typ: "date" },
    { klucz: "WROF", etykieta: "Czy jest Order Form od klienta?", typ: "bool" },
    { klucz: "WRO", etykieta: "Zamówienie od klienta", typ: "multiline" },
  ],
];

// ostatnio ogladana zakladka danych etapu - przezywa przerysowanie karty po zapisie
const zakladkaEtapuSzansy = new Map<string, number>();

/** Wartosc pola etapu w postaci do kontrolki formularza. */
function wartoscDoPola(typ: TypPolaEtapu, v: unknown): string | boolean {
  if (typ === "bool") return v === true;
  if (v == null) return "";
  if (typ === "date") return String(v).slice(0, 10);
  if (typ === "link") return String((v as any)?.Url ?? v);
  return String(v);
}

/** Wartosc z kontrolki w postaci do zapisu w SharePoint (puste -> null czysci pole). */
function wartoscDoZapisu(typ: TypPolaEtapu, v: string | boolean): unknown {
  if (typ === "bool") return v === true;
  const t = String(v);
  if (typ === "date") return t ? dzienSP(t) : null;
  if (typ === "number") {
    const n = Number(t.replace(/\s/g, "").replace(",", "."));
    return t.trim() && Number.isFinite(n) ? n : null;
  }
  return t;
}

/**
 * Dane etapu: zakladki 1-8, domyslnie biezacy etap. Zmiany zbieramy ze wszystkich
 * zakladek i zapisujemy razem - Power Apps zapisywal tylko widoczna i reszta przepadala.
 */
async function sekcjaDanychEtapu(box: HTMLElement, o: Szansa, zapisz: (pola: Record<string, unknown>, b?: HTMLButtonElement) => Promise<void>) {
  const nazwy = new Map<string, string>();
  await Promise.all(POLA_ETAPOW.flat().map(async (p) => nazwy.set(p.klucz, await nazwaWewnetrzna("Opps", p.klucz))));
  const oryginal = (p: { klucz: string; typ: TypPolaEtapu }) => wartoscDoPola(p.typ, o.pola[nazwy.get(p.klucz)!]);

  const biezacy = ETAPY_SZANSY.findIndex(([v]) => v === o.etap) + 1;
  let zakladka = zakladkaEtapuSzansy.get(o.id) ?? Math.max(biezacy, 1);
  const zmiany = new Map<string, string | boolean>();

  const pasek = el("div", "crm-szansa-zakladki");
  const pola = el("div", "crm-szansa-pola");
  const zapiszBtn = el("button", "crm-btn") as HTMLButtonElement;
  const info = el("div", "crm-komunikat");
  const rzad = el("div", "crm-akcje");
  rzad.append(zapiszBtn, info);
  box.append(pasek, pola, rzad);

  const odswiezPrzycisk = () => {
    zapiszBtn.textContent = zmiany.size ? `Zapisz dane etapu (${zmiany.size})` : "Zapisz dane etapu";
    zapiszBtn.disabled = !zmiany.size;
    // zakladki z niezapisanymi zmianami oznaczone kropka
    pasek.querySelectorAll<HTMLButtonElement>(".crm-szansa-zakl").forEach((b, i) => {
      const brudna = POLA_ETAPOW[i].some((p) => zmiany.has(p.klucz));
      b.classList.toggle("crm-zakl-zmiany", brudna);
    });
  };

  const rysujPola = () => {
    pola.innerHTML = "";
    pasek.querySelectorAll(".crm-szansa-zakl").forEach((b, i) => b.classList.toggle("crm-zakl-akt", i + 1 === zakladka));
    POLA_ETAPOW[zakladka - 1].forEach((p) => {
      const start = oryginal(p);
      const teraz = zmiany.has(p.klucz) ? zmiany.get(p.klucz)! : start;
      const ustaw = (v: string | boolean) => {
        if (v === start) zmiany.delete(p.klucz);
        else zmiany.set(p.klucz, v);
        odswiezPrzycisk();
      };

      if (p.typ === "bool") {
        const cb = document.createElement("input");
        cb.type = "checkbox";
        cb.id = `etap-${o.id}-${p.klucz}`;
        cb.checked = teraz === true;
        cb.onchange = () => ustaw(cb.checked);
        const lbl = el("label", "crm-check") as HTMLLabelElement;
        lbl.htmlFor = cb.id;
        lbl.append(cb, el("span", undefined, p.etykieta));
        pola.appendChild(lbl);
        return;
      }

      pola.appendChild(el("label", "crm-label", p.etykieta));
      // link (kolumna Hyperlink) tylko do odczytu - zapis Hyperlinka przez Graph jest zawodny
      if (p.typ === "link") {
        const url = String(teraz);
        if (url) {
          const a = document.createElement("a");
          a.href = url;
          a.target = "_blank";
          a.rel = "noopener";
          a.className = "crm-szansa-link";
          a.textContent = url;
          pola.appendChild(a);
        } else pola.appendChild(el("div", "crm-komunikat", "brak – link uzupełnisz w Power Apps"));
        return;
      }

      const pole = p.typ === "multiline" ? document.createElement("textarea") : document.createElement("input");
      if (pole instanceof HTMLTextAreaElement) {
        pole.className = "crm-textarea";
        pole.rows = 3;
      } else {
        pole.className = "crm-input";
        pole.type = p.typ === "date" ? "date" : "text";
        if (p.typ === "number") pole.inputMode = "decimal";
      }
      pole.value = String(teraz);
      pole.oninput = () => ustaw(pole.value);
      pola.appendChild(pole);
    });
  };

  ETAPY_SZANSY.forEach(([v, krotko], i) => {
    const nr = i + 1;
    const b = el("button", "crm-szansa-zakl", `${nr}. ${krotko}${nr === biezacy ? " ●" : ""}`) as HTMLButtonElement;
    b.title = nr === biezacy ? `${v} – bieżący etap` : v;
    b.onclick = () => {
      zakladka = nr;
      zakladkaEtapuSzansy.set(o.id, nr);
      rysujPola();
    };
    pasek.appendChild(b);
  });

  zapiszBtn.onclick = () => {
    const doZapisu: Record<string, unknown> = {};
    POLA_ETAPOW.flat().forEach((p) => {
      if (zmiany.has(p.klucz)) doZapisu[p.klucz] = wartoscDoZapisu(p.typ, zmiany.get(p.klucz)!);
    });
    zakladkaEtapuSzansy.set(o.id, zakladka);
    void zapisz(doZapisu, zapiszBtn);
  };

  rysujPola();
  odswiezPrzycisk();
}

function kolorStatusuSzansy(status: string): [string, string] {
  const st = STATUSY_SZANSY.find((x) => x.v === status);
  return st ? [st.kol, st.tlo] : ["#6B7280", "#F3F4F6"];
}

/**
 * Czesciowy zapis szansy z kontrola wspolbieznosci (If-Match) - Power Apps jej nie mial.
 * Klucze moga byc nazwami wyswietlanymi, tlumaczymy je na wewnetrzne.
 */
async function zapiszSzanse(o: Szansa, pola: Record<string, unknown>): Promise<Szansa> {
  const doZapisu: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(pola)) doZapisu[await nazwaWewnetrzna("Opps", k)] = v;
  const wynik = await updateTicketFields(o.id, doZapisu, o.etag || undefined, "Opps");
  // bez nowego etagu nie wysylamy starego - inaczej kolejny zapis dostalby 412
  const f = { ...o.pola, ...wynik, "@odata.etag": wynik["@odata.etag"] ?? "" };
  return toSzansa({ id: o.id, fields: f } as unknown as ListItem);
}

/** Wartosc pola szansy po nazwie wyswietlanej (np. ZamkniętyOppsNotes). */
async function poleSzansy(o: Szansa, klucz: string): Promise<unknown> {
  return o.pola[await nazwaWewnetrzna("Opps", klucz)];
}

/** Panel zarzadzania szansa: status (z data wygranej / powodami przegranej) i etap. */
function panelSzansy(o: Szansa, c: Company, zmieniona: (nowa: Szansa) => void): HTMLElement {
  const box = el("div", "crm-szansa-panel");
  const komunikat = el("div", "crm-komunikat");

  const zapisz = async (pola: Record<string, unknown>, przycisk?: HTMLButtonElement) => {
    if (przycisk) przycisk.disabled = true;
    komunikat.textContent = "Zapisywanie...";
    komunikat.className = "crm-komunikat";
    try {
      zmieniona(await zapiszSzanse(o, pola));
    } catch (e: any) {
      komunikat.textContent = e?.message ?? String(e);
      komunikat.className = "crm-komunikat crm-blad";
      if (przycisk) przycisk.disabled = false;
    }
  };

  // --- status ---
  box.appendChild(el("div", "crm-nf-sekcja", "Status"));
  const statusy = el("div", "crm-szansa-statusy");
  box.appendChild(statusy);
  const podforma = el("div");
  box.appendChild(podforma);

  STATUSY_SZANSY.forEach((st) => {
    const akt = o.status === st.v;
    const b = el("button", "crm-szansa-status", (akt ? "✓ " : "") + st.krotko) as HTMLButtonElement;
    b.style.setProperty("--kol", st.kol);
    b.style.setProperty("--tlo", st.tlo);
    b.classList.toggle("crm-szansa-status-akt", akt);
    b.onclick = () => {
      podforma.innerHTML = "";
      komunikat.textContent = "";
      if (st.v === "Aktywny") formaAktywny(b);
      else if (st.v === "Zamknięty wygrany") formaWygrany();
      else void formaPrzegrany();
    };
    statusy.appendChild(b);
  });

  const przyciski = (tekst: string, styl = "crm-btn") => {
    const ok = el("button", styl, tekst) as HTMLButtonElement;
    const nie = el("button", "crm-btn crm-btn-cichy", "Anuluj") as HTMLButtonElement;
    nie.onclick = () => (podforma.innerHTML = "");
    const rzad = el("div", "crm-akcje");
    rzad.append(ok, nie);
    return { ok, rzad };
  };

  // powrot do aktywnych czysci dane zamkniecia - Power Apps je zostawial i po ponownym zamknieciu wisialy stare
  function formaAktywny(b: HTMLButtonElement) {
    if (o.status === "Aktywny") return;
    if (!o.status) return void zapisz({ Status1: "Aktywny" }, b);
    const f = el("div", "crm-nowe crm-szansa-forma");
    f.appendChild(el("div", "crm-nowe-info",
      "Przywrócić szansę jako aktywną? Data wygranej i zaznaczone powody przegranej zostaną wyczyszczone (opis przegranej zostaje)."));
    const { ok, rzad } = przyciski("Przywróć");
    ok.onclick = () => void zapisz({
      Status1: "Aktywny",
      ClosedDateWin: null,
      DR_Odrzucony: false,
      Konkurencja: false,
      BrakDofinansowania: false,
    }, ok);
    f.appendChild(rzad);
    podforma.appendChild(f);
  }

  function formaWygrany() {
    const f = el("div", "crm-nowe crm-szansa-forma");
    const data = poleFormularza(f, "Data zamknięcia *", "date");
    const bylo = String(o.pola.ClosedDateWin ?? "");
    data.value = bylo ? bylo.slice(0, 10) : dzisISO();
    const { ok, rzad } = przyciski(o.status === "Zamknięty wygrany" ? "Zapisz datę" : "Oznacz jako wygraną");
    // jak w Power Apps: wygrana i od razu formularz nowej szansy TECS dla firmy
    const zTecs = el("button", "crm-btn crm-btn-cichy", "Zapisz i utwórz TECS") as HTMLButtonElement;
    rzad.insertBefore(zTecs, rzad.lastChild);
    data.oninput = () => (ok.disabled = zTecs.disabled = !data.value);
    const pola = () => ({ Status1: "Zamknięty wygrany", ClosedDateWin: dzienSP(data.value) });
    ok.onclick = () => void zapisz(pola(), ok);
    zTecs.onclick = async () => {
      zTecs.disabled = ok.disabled = true;
      komunikat.textContent = "Zapisywanie...";
      komunikat.className = "crm-komunikat";
      try {
        nowyTecsZSzansy = await zapiszSzanse(o, pola());
        zakladkaFirmy = "TECS";
        tecsTryb = "Bieżące";
        otworzNowyTecs = true;
        await pokazFirma(c);
      } catch (e: any) {
        komunikat.textContent = e?.message ?? String(e);
        komunikat.className = "crm-komunikat crm-blad";
        zTecs.disabled = ok.disabled = false;
      }
    };
    f.appendChild(rzad);
    podforma.appendChild(f);
  }

  async function formaPrzegrany() {
    const f = el("div", "crm-nowe crm-szansa-forma");
    podforma.appendChild(f);
    f.appendChild(el("label", "crm-label", "Powód przegranej"));
    const wybrane = new Set<string>();
    const chipy = el("div", "crm-chipy");
    f.appendChild(chipy);
    for (const [pole, tekst] of POWODY_PRZEGRANEJ) {
      const b = el("button", "crm-chip-wybor", tekst) as HTMLButtonElement;
      if ((await poleSzansy(o, pole)) === true) wybrane.add(pole);
      b.classList.toggle("crm-chip-akt", wybrane.has(pole));
      b.onclick = () => {
        if (wybrane.has(pole)) wybrane.delete(pole);
        else wybrane.add(pole);
        b.classList.toggle("crm-chip-akt", wybrane.has(pole));
        sprawdz();
      };
      chipy.appendChild(b);
    }
    f.appendChild(el("label", "crm-label", "Opis"));
    const ta = document.createElement("textarea");
    ta.className = "crm-textarea";
    ta.rows = 3;
    ta.placeholder = "Dlaczego szansa została przegrana?";
    ta.value = String((await poleSzansy(o, NOTATKA_PRZEGRANEJ)) ?? "");
    ta.oninput = () => sprawdz();
    f.appendChild(ta);

    // Power Apps pozwalal zamknac bez powodu - wymagamy powodu albo opisu
    const brak = el("div", "crm-podpowiedz");
    f.appendChild(brak);
    const { ok, rzad } = przyciski(o.status === "Zamknięty przegrany" ? "Zapisz zmiany" : "Zamknij jako przegraną",
      "crm-btn crm-btn-ostrz");
    f.appendChild(rzad);
    function sprawdz() {
      const pusto = !wybrane.size && !ta.value.trim();
      ok.disabled = pusto;
      brak.textContent = pusto ? "Zaznacz powód albo opisz, dlaczego szansa przepadła." : "";
      brak.style.display = pusto ? "block" : "none";
    }
    sprawdz();
    ok.onclick = () => {
      const pola: Record<string, unknown> = { Status1: "Zamknięty przegrany", [NOTATKA_PRZEGRANEJ]: ta.value };
      POWODY_PRZEGRANEJ.forEach(([pole]) => (pola[pole] = wybrane.has(pole)));
      void zapisz(pola, ok);
    };
  }

  // --- etap: klik zapisuje od razu, mozna cofac (jak w Power Apps); status sie nie zmienia ---
  box.appendChild(el("div", "crm-nf-sekcja", "Etap"));
  const kroki = el("div", "crm-szansa-etapy");
  box.appendChild(kroki);
  const biezacy = ETAPY_SZANSY.findIndex(([v]) => v === o.etap) + 1;
  ETAPY_SZANSY.forEach(([v, krotko], i) => {
    const nr = i + 1;
    const b = el("button", "crm-szansa-krok", `${nr < biezacy ? "✓" : nr + "."} ${krotko}`) as HTMLButtonElement;
    b.title = v;
    b.classList.add(nr === biezacy ? "crm-krok-biez" : nr < biezacy ? "crm-krok-ok" : "crm-krok-przysz");
    b.onclick = () => {
      if (nr === biezacy) return;
      void zapisz({ Etap1: v }, b);
    };
    kroki.appendChild(b);
  });

  box.appendChild(el("div", "crm-nf-sekcja", "Dane etapu"));
  const dane = el("div");
  box.appendChild(dane);
  sekcjaDanychEtapu(dane, o, zapisz).catch((e) =>
    dane.appendChild(el("div", "crm-komunikat crm-blad", `Nie wczytano pól etapu: ${e?.message ?? e}`))
  );

  box.appendChild(komunikat);
  return box;
}

function kartaSzansy(
  o: Szansa,
  c: Company,
  notatki: Promise<Nota[]>,
  ja: Promise<string>,
  panelOtwarty = false
): HTMLElement {
  const k = el("div", "crm-kafel");
  const [kol, tlo] = kolorStatusuSzansy(o.status);
  k.style.borderLeftColor = kol;

  const gora = el("div", "crm-kafel-rzad");
  gora.appendChild(poleIk("crm-kafel-nazwa", "cel", o.nazwa, "Nazwa szansy"));
  const status = pastylka(o.status || "—", kol, tlo, "flaga");
  status.title = "Status";
  gora.appendChild(status);
  k.appendChild(gora);

  const rzad = el("div", "crm-kafel-rzad");
  rzad.appendChild(poleIk("crm-kafel-wartosc", "euro", euro(o.wartosc), "Wartość"));
  if (o.typ) rzad.appendChild(pastylka(o.typ, "#2F6FD6", "#E8F0FD", "teczka"));
  k.appendChild(rzad);

  if (o.etap) {
    const nr = ETAPY_SZANSY.findIndex(([v]) => v === o.etap) + 1;
    k.appendChild(poleIk("crm-kafel-meta", "wykres", `Etap: ${nr ? `${nr}/8 · ` : ""}${o.etap}`, "Etap sprzedaży"));
    if (nr) {
      const postep = el("div", "crm-szansa-postep");
      postep.title = `Etap ${nr} z ${ETAPY_SZANSY.length}`;
      ETAPY_SZANSY.forEach((_, i) => {
        const odc = el("span", "crm-szansa-odc");
        if (i < nr) odc.style.background = kol;
        postep.appendChild(odc);
      });
      k.appendChild(postep);
    }
  }
  k.appendChild(poleIk("crm-kafel-meta", "kalendarz", `Otwarta ${dataPL(o.otwarcie)}`, "Data otwarcia"));

  const akcje = el("div", "crm-akcje crm-szansa-akcje");
  k.appendChild(akcje);
  const btnPanel = przyciskIk("crm-btn crm-btn-cichy", "flaga", "Status i etap");
  akcje.appendChild(btnPanel);
  const panelBox = el("div");
  k.appendChild(panelBox);

  // po zapisie karta rysowana od nowa (kolor, status, etap) z otwartym panelem
  const pokazPanel = () => {
    panelBox.appendChild(panelSzansy(o, c, (nowa) => k.replaceWith(kartaSzansy(nowa, c, notatki, ja, true))));
    btnPanel.classList.add("crm-btn-wcisniety");
  };
  btnPanel.onclick = () => {
    if (panelBox.childElementCount) {
      panelBox.innerHTML = "";
      btnPanel.classList.remove("crm-btn-wcisniety");
    } else pokazPanel();
  };
  if (panelOtwarty) pokazPanel();

  // notatki szansy - bez OPPSGUID nie da sie ich powiazac
  if (!o.guid) return k;

  const btnNoty = przyciskIk("crm-btn crm-btn-cichy", "dymek", "Notatki …");
  const btnNowa = przyciskIk("crm-btn", "pioro", "Dodaj notatkę");
  akcje.append(btnNoty, btnNowa);
  const formBox = el("div");
  k.appendChild(formBox);
  const listaNot = el("div", "crm-szansa-noty");
  listaNot.style.display = "none";
  k.appendChild(listaNot);

  let moje: Nota[] = [];
  let mojeJa = "";
  const etykieta = btnNoty.querySelector("span")!;
  const rysuj = () => {
    etykieta.textContent = `Notatki (${moje.length})`;
    listaNot.innerHTML = "";
    if (!moje.length) listaNot.appendChild(el("div", "crm-pusto", "Brak notatek do tej szansy."));
    moje.forEach((nt) => listaNot.appendChild(kartaNotatki(nt, mojeJa, () => {
      moje = moje.filter((x) => x !== nt);
      // wspolna lista firmy tez - karta po zmianie statusu rysuje sie z niej od nowa
      void notatki.then((a) => {
        const i = a.indexOf(nt);
        if (i >= 0) a.splice(i, 1);
      });
      rysuj();
    })));
  };
  Promise.all([notatki, ja]).then(
    ([wszystkie, j]) => {
      mojeJa = j;
      moje = wszystkie.filter((n) => n.opps === o.guid).sort((a, b) => (b.data || "").localeCompare(a.data || ""));
      rysuj();
    },
    () => (etykieta.textContent = "Notatki (błąd)")
  );

  btnNoty.onclick = () => {
    const widac = listaNot.style.display !== "none";
    listaNot.style.display = widac ? "none" : "";
    btnNoty.classList.toggle("crm-btn-wcisniety", !widac);
  };
  btnNowa.onclick = () => {
    btnNowa.disabled = true;
    const zamknij = () => {
      formBox.innerHTML = "";
      btnNowa.disabled = false;
    };
    void formularzNotatkiSzansy(formBox, o, c, (nt) => {
      zamknij();
      moje = [nt, ...moje];
      void notatki.then((a) => a.push(nt));
      rysuj();
      listaNot.style.display = "";
      btnNoty.classList.add("crm-btn-wcisniety");
    }, zamknij);
  };

  return k;
}

/** Nowa notatka do szansy (Notes.Opps_ID = OPPSGUID). Odpowiednik conNotatkaOpp z Power Apps. */
async function formularzNotatkiSzansy(
  miejsce: HTMLElement,
  o: Szansa,
  c: Company,
  poZapisie: (nt: Nota) => void,
  anuluj: () => void
) {
  const box = el("div", "crm-nowe crm-demo");
  miejsce.appendChild(box);
  box.appendChild(poleIk("crm-demo-tytul", "dokument", "Nowa notatka"));
  box.appendChild(el("div", "crm-nowe-info", `${o.firma || c.nazwa} — ${o.nazwa}`));

  const data = poleFormularza(box, "Data wydarzenia *", "date");
  data.value = dzisISO();

  // --- osoba: z kontaktow firmy albo wpisana recznie ---
  box.appendChild(el("label", "crm-label", "Osoba *"));
  const osobaSel = document.createElement("select");
  osobaSel.className = "crm-select";
  const brakOsoby = document.createElement("option");
  brakOsoby.value = "";
  brakOsoby.textContent = "wybierz osobę z firmy…";
  osobaSel.appendChild(brakOsoby);
  box.appendChild(osobaSel);
  const osoba = document.createElement("input");
  osoba.type = "text";
  osoba.className = "crm-input";
  osoba.placeholder = "imię i nazwisko";
  box.appendChild(osoba);

  let kontakty: Contact[] = [];
  let wybrana: Contact | null = null;
  try {
    kontakty = (await getContactsByCompany(o.companyGuid || c.guid)).map(toContact)
      .filter((x) => x.aktywny)
      .sort((a, b) => a.imieNazwisko.localeCompare(b.imieNazwisko, "pl"));
    kontakty.forEach((x, i) => {
      const opt = document.createElement("option");
      opt.value = String(i);
      opt.textContent = x.imieNazwisko + (x.stanowisko ? ` · ${x.stanowisko}` : "");
      osobaSel.appendChild(opt);
    });
  } catch (e: any) {
    brakOsoby.textContent = `kontakty: ${e?.message ?? e}`;
  }
  osobaSel.onchange = () => {
    wybrana = osobaSel.value ? kontakty[Number(osobaSel.value)] : null;
    if (wybrana) osoba.value = wybrana.imieNazwisko;
    sprawdz();
  };
  osoba.oninput = () => {
    if (wybrana && osoba.value.trim() !== wybrana.imieNazwisko) {
      wybrana = null;
      osobaSel.value = "";
    }
    sprawdz();
  };

  // domyslnie nadawca otwartego maila, jesli jest w tej firmie
  const nadawca = ((Office.context.mailbox.item as Office.MessageRead)?.from?.emailAddress ?? "").toLowerCase();
  const iNadawcy = kontakty.findIndex((x) => nadawca && x.email.toLowerCase() === nadawca);
  if (iNadawcy >= 0) {
    osobaSel.value = String(iNadawcy);
    wybrana = kontakty[iNadawcy];
    osoba.value = wybrana.imieNazwisko;
  }

  // --- rodzaj interakcji (lookup do listy activities) ---
  box.appendChild(el("label", "crm-label", "Rodzaj interakcji *"));
  const rodzaje = el("div", "crm-chipy crm-chipy-rowne");
  box.appendChild(rodzaje);
  let rodzaj = "";
  let rodzajId = "";
  try {
    const opcje = await getLookupOptions("activities", "Notes_types");
    opcje.forEach((op) => {
      const b = chipRodzaju(op.etykieta);
      b.onclick = () => {
        rodzaj = op.etykieta;
        rodzajId = op.id;
        rodzaje.querySelectorAll(".crm-chip-wybor").forEach((x) => x.classList.remove("crm-chip-akt"));
        b.classList.add("crm-chip-akt");
        sprawdz();
      };
      rodzaje.appendChild(b);
    });
  } catch (e: any) {
    box.appendChild(el("div", "crm-nowe-info", `Nie pobrano rodzajów: ${e?.message ?? e}`));
  }
  const mail = opcjaMaila(box, () => sprawdz());

  box.appendChild(el("label", "crm-label", "Notatka *"));
  const ta = document.createElement("textarea");
  ta.className = "crm-textarea";
  ta.rows = 5;
  ta.placeholder = "Co ustaliliście?";
  ta.oninput = () => sprawdz();
  box.appendChild(ta);

  // --- follow-up ---
  const fu = document.createElement("input");
  fu.type = "checkbox";
  fu.id = `fu-opp-${o.id}`;
  const fuLbl = el("label", "crm-check") as HTMLLabelElement;
  fuLbl.htmlFor = fu.id;
  fuLbl.append(fu, el("span", undefined, "Follow-up (przypomnienie)"));
  box.appendChild(fuLbl);

  const fuBox = el("div");
  fuBox.style.display = "none";
  box.appendChild(fuBox);
  const fuData = document.createElement("input");
  fuData.type = "date";
  fuData.className = "crm-input";
  fuData.value = plusDni(data.value, 7);
  fuBox.appendChild(fuData);
  let fuDni = 7;
  chipyWyboru(
    fuBox,
    [7, 14, 21, 28, 56].map((d) => ({ tekst: d === 56 ? "+2 mies." : `+${d} dni`, wartosc: d })),
    fuDni,
    (d) => {
      fuDni = d;
      fuData.value = plusDni(data.value, d);
      opiszFu();
    }
  );
  const fuInfo = el("div", "crm-komunikat");
  fuBox.appendChild(fuInfo);
  const opiszFu = () => {
    if (!fuData.value) return (fuInfo.textContent = "");
    const d = new Date(fuData.value + "T12:00:00");
    const za = Math.round((d.getTime() - new Date(dzisISO() + "T12:00:00").getTime()) / 86400000);
    fuInfo.textContent = `Przypomnienie: ${d.toLocaleDateString("pl-PL", {
      weekday: "long", day: "numeric", month: "long", year: "numeric",
    })} (za ${za} dni)`;
  };
  opiszFu();
  fuData.onchange = opiszFu;
  data.onchange = () => {
    fuData.value = plusDni(data.value, fuDni);
    opiszFu();
    sprawdz();
  };
  fu.onchange = () => (fuBox.style.display = fu.checked ? "" : "none");

  // --- przyciski ---
  const brakujace = el("div", "crm-podpowiedz");
  box.appendChild(brakujace);
  const zapisz = el("button", "crm-btn", "Zapisz notatkę") as HTMLButtonElement;
  const wstecz = el("button", "crm-btn crm-btn-cichy", "Anuluj") as HTMLButtonElement;
  const komunikat = el("span", "crm-komunikat");
  const rzad = el("div", "crm-akcje");
  rzad.append(zapisz, wstecz, komunikat);
  box.appendChild(rzad);
  wstecz.onclick = anuluj;

  function sprawdz(): boolean {
    const brak: string[] = [];
    if (!data.value) brak.push("data");
    if (!osoba.value.trim()) brak.push("osoba");
    if (!rodzajId) brak.push("rodzaj");
    mail.pokaz(rodzaj);
    if (!ta.value.trim() && !mail.wlaczony) brak.push("treść");
    const ok = brak.length === 0;
    zapisz.disabled = !ok;
    brakujace.textContent = ok ? "" : `Uzupełnij: ${brak.join(", ")}.`;
    brakujace.style.display = ok ? "none" : "block";
    return ok;
  }
  sprawdz();

  zapisz.onclick = async () => {
    if (!sprawdz()) return;
    zapisz.disabled = true;
    komunikat.textContent = "Zapisywanie...";
    komunikat.className = "crm-komunikat";

    let autor = "";
    try {
      autor = await getMojaNazwa();
    } catch {
      autor = Office.context.mailbox.userProfile.displayName ?? "";
    }

    const followUp = fu.checked && fuData.value ? dzienSP(fuData.value) : null;
    const pola: Record<string, unknown> = {
      Data: dzienSP(data.value),
      Osoba: osoba.value.trim(),
      Telefon: wybrana?.telefon ?? "",
      TypeLookupId: rodzajId,
      Rodzaj1: rodzaj,
      Totatka: ta.value,            // uwaga: literowka w nazwie kolumny
      Firma: o.firma || firmaPelna?.skrocona || c.nazwa,
      Creator: autor,
      Opps_ID: o.guid,
      Etap: o.etap,
      IdentyfikatorCompanies: o.companyGuid || c.guid,
    };
    if (followUp) pola["Follow_up"] = followUp;

    try {
      if (mail.wlaczony) Object.assign(pola, await mail.pola());
      const nowy = await createListItem("Notes", pola);
      poZapisie({
        id: String(nowy.id ?? ""),
        data: String(pola.Data ?? ""),
        osoba: osoba.value.trim(),
        rodzaj,
        rodzajId,
        tresc: ta.value,
        autor,
        followUp: followUp ?? "",
        followDone: false,
        opps: o.guid,
      });
    } catch (e: any) {
      komunikat.textContent = e?.message ?? String(e);
      komunikat.className = "crm-komunikat crm-blad";
      zapisz.disabled = false;
    }
  };
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
  if (widok === "szanse") return pokazSzanse();
  return pokazZgloszenia();
}
