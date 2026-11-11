// src/taskpane/outlook.ts
/* global document, Office, console, setTimeout, window, navigator, clearTimeout */
import {
  getListItems, getTicketsByField, updateTicketFields, getCompanies2, getOpiekunowie,
  wyczyscCacheFirm, getContactByEmail, createTicket, getSerwisanci, getContactsByCompany,
  getTicketsFiltered, getMojMail, getByCompany, getCompanyByGuid,createListItem, getLookupOptions, getMojaNazwa, deleteListItem,
  getTicketByNumber, getAllCompanies, getChoiceOptions, getAllContacts, wyczyscCacheKontaktow,
  getLiczbaSupportowRok, getSiteUserId, dodajDoKontaktowOutlook, getZalaczniki, dodajZalacznik, usunZalacznik, pobierzZalacznik, getPlikiFolderu, getUrlFolderu, utworzFolder, wyslijPlik, usunPlik, getCennik, getOpcjeLookupu, czyPracownikMESco, getZgloszeniaObslugujacego, nazwaWewnetrzna, getCenyTecs, getProgiRabatowe,
} from "./graphService";
import type { ListItem, PozycjaCennika, CenaTecs, ProgRabatu } from "./graphService";
import {
  toTicket, toCompany, formatData, dniOdZmiany, KOLOR_STATUSU, ticketIdZTematu, Ticket, Company,
  toContact, Contact, toCompanyFull, CompanyFull, toNota, Nota, toTecs, toSzansa, Szansa,
  toLicencjaTecs, toLicencjaBez, Licencja, toSzkolenie, dniDo, kolorTerminu,
  ETAP_KOLORY, ETAP_SKROTY, rozbijNazwe, plecZImienia, nowyGuid, TecsRow, toProduktTecs, ProduktTecs, toProduktSzansy, ProduktSzansy,
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
  // wejscie w Szanse zawsze od szans sprzedazy - TECS to przelacznik w srodku
  document.getElementById("nav-szanse")!.onclick = () => {
    szanseRodzaj = "Sprzedaż";
    void pokazSzanse();
  };

  Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => {
    if (widok === "zgloszenia") void pokazZgloszenia();
    if (widok === "szanse") {
      szanseRodzaj = "Sprzedaż";
      void pokazSzanse();
    }
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
  | "Szanse" | "Posiadane" | "Szkolenia" | "Uwagi" | "Dokumenty";

let zakladkaFirmy: Zakladka = "Kontakty";
let tecsTryb: "Bieżące" | "Archiwalne" | "Demo" = "Bieżące";
let otworzKodyDemo = false; // kafelek "Nowe kody demo" -> TECS/Demo z otwartym formularzem
let otworzNowaSzanse = false; // "Nowy Opps Soft" z paska firmy -> Szanse z otwartym formularzem
let otworzNowyTecs = false; // "Zapisz i utworz TECS" przy wygranej -> TECS/Biezace z otwartym formularzem
let nowyTecsZSzansy: Szansa | null = null; // wygrana szansa - jej produkty podpowiadamy do TECS-u
let szanseZamkniete = false;
// jak OnVisible w Power Apps: po wejsciu w firme bez biezacych TECS-ow / aktywnych szans
// zakladka sama pokazuje archiwalne / zamkniete - tylko przy pierwszym wyswietleniu
let szczegolyFirmy = false; // karta firmy: adres, NIP, status, NDA, godziny zwiniete pod "Pokaż szczegóły"
let autoTecs = false;
let autoSzanse = false;
let szanseRodzaj: "Sprzedaż" | "TECS" = "Sprzedaż"; // widok maila handlowca: szanse sprzedazy albo TECS
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
  hasz: '<line x1="4" x2="20" y1="9" y2="9"/><line x1="4" x2="20" y1="15" y2="15"/><line x1="10" x2="8" y1="3" y2="21"/><line x1="16" x2="14" y1="3" y2="21"/>',
  ludzie: '<path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M22 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/>',
  tarczaOk: '<path d="M20 13c0 5-3.5 7.5-7.66 8.95a1 1 0 0 1-.67-.01C7.5 20.5 4 18 4 13V6a1 1 0 0 1 1-1c2 0 4.5-1.2 6.24-2.72a1.17 1.17 0 0 1 1.52 0C14.51 3.81 17 5 19 5a1 1 0 0 1 1 1z"/><path d="m9 12 2 2 4-4"/>',
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
  spinacz: '<path d="m21.44 11.05-9.19 9.19a6 6 0 0 1-8.49-8.49l8.57-8.57A4 4 0 1 1 18 8.84l-8.59 8.57a2 2 0 0 1-2.83-2.83l8.49-8.48"/>',
  wyslij: '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" x2="12" y1="3" y2="15"/>',
  nowaSzansa: '<circle cx="12" cy="12" r="10"/><path d="M8 12h8"/><path d="M12 8v8"/>',
  kalkulator: '<rect width="16" height="20" x="4" y="2" rx="2"/><line x1="8" x2="16" y1="6" y2="6"/><line x1="16" x2="16" y1="14" y2="18"/><path d="M16 10h.01"/><path d="M12 10h.01"/><path d="M8 10h.01"/><path d="M12 14h.01"/><path d="M8 14h.01"/><path d="M12 18h.01"/><path d="M8 18h.01"/>',
  blyskawica: '<path d="M4 14a1 1 0 0 1-.78-1.63l9.9-10.2a.5.5 0 0 1 .86.46l-1.92 6.02A1 1 0 0 0 13 10h7a1 1 0 0 1 .78 1.63l-9.9 10.2a.5.5 0 0 1-.86-.46l1.92-6.02A1 1 0 0 0 11 14z"/>',
  dodajOsobe: '<path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><line x1="19" x2="19" y1="8" y2="14"/><line x1="22" x2="16" y1="11" y2="11"/>',
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

/** Podswietla opcje w pierwszym przelaczniku w box - gdy tryb zmienil sie po narysowaniu. */
function zaznaczSegment(box: HTMLElement, wartosc: string) {
  box.querySelector(".crm-segment")?.querySelectorAll(".crm-segment-btn").forEach((b) =>
    b.classList.toggle("crm-segment-akt", b.textContent === wartosc)
  );
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

  // usuwanie - tylko wlasciciel zgloszenia (obslugujacy, MailCreator), z potwierdzeniem
  const usun = przyciskIk("crm-btn crm-btn-cichy crm-zgl-usun", "kosz", "Usuń");
  usun.title = "Usuń zgłoszenie";
  usun.hidden = true;
  pasek.insertBefore(usun, komunikat);
  if (t.opiekunMail) {
    void getMojMail().then((m) => (usun.hidden = m.trim().toLowerCase() !== t.opiekunMail), () => undefined);
  }
  usun.onclick = () => {
    ostrzezenie.innerHTML = "";
    ostrzezenie.appendChild(el("div", undefined,
      `Usunąć zgłoszenie MESco-ID-${t.ticketId}${t.osoba ? ` (${t.osoba})` : ""}? Tej operacji nie da się cofnąć z dodatku.`));
    const tak = el("button", "crm-btn crm-btn-ostrz", "Usuń zgłoszenie") as HTMLButtonElement;
    const nie = el("button", "crm-btn crm-btn-cichy", "Anuluj") as HTMLButtonElement;
    const rzad = el("div", "crm-akcje");
    rzad.append(tak, nie);
    ostrzezenie.appendChild(rzad);
    ostrzezenie.style.display = "block";
    nie.onclick = () => { ostrzezenie.style.display = "none"; };
    tak.onclick = async () => {
      tak.disabled = nie.disabled = usun.disabled = zapisz.disabled = true;
      try {
        await deleteListItem("Support", t.id);
        const karta = box.closest(".crm-card");
        if (karta) karta.remove();
        else box.remove();
      } catch (e: any) {
        ostrzezenie.style.display = "none";
        komunikat.textContent = e?.message ?? String(e);
        komunikat.className = "crm-komunikat crm-blad";
        usun.disabled = zapisz.disabled = false;
      }
    };
  };

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
  anuluj?: () => void,
  edycja: Contact | null = null
) {
  if (edycja) box.appendChild(el("div", "crm-nowe-info", `Edycja: ${edycja.imieNazwisko || "kontakt"}`));
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
  const telefon = pole("Telefon", edycja?.telefon ?? "", "tel");
  const stanowisko = pole("Stanowisko", edycja?.stanowisko ?? "");

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
  if (edycja?.plec) {
    plecRecznie = true;
    zaznaczPlec(edycja.plec);
  } else podpowiedzPlci();

  const wybraneRole = new Set<string>(edycja?.role ?? []);
  if (role.length) {
    box.appendChild(el("label", "crm-label", "Rola (MEDDIC)"));
    const rolaBox = el("div", "crm-chipy");
    role.forEach((o) => {
      const b = el("button", "crm-chip-wybor", o) as HTMLButtonElement;
      b.classList.toggle("crm-chip-akt", wybraneRole.has(o));
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
  stala.checked = edycja?.stala ?? false;
  tymczasowa.checked = edycja?.tymczasowa ?? false;
  // Status = true znaczy "juz nie pracuje" - przy nowym kontakcie bez sensu
  const niePracuje = edycja ? przelacznik("Nie pracuje już w firmie") : null;
  if (niePracuje) niePracuje.checked = !edycja!.aktywny;

  box.appendChild(el("label", "crm-label", "Uwagi"));
  const uwagi = document.createElement("textarea");
  uwagi.className = "crm-textarea";
  uwagi.rows = 2;
  uwagi.value = edycja?.uwagi ?? "";
  box.appendChild(uwagi);

  const brakujace = el("div", "crm-podpowiedz");
  box.appendChild(brakujace);

  const zapisz = el("button", "crm-btn", edycja ? "Zapisz zmiany" : "Zapisz kontakt") as HTMLButtonElement;
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
    // stare kontakty bywaja bez plci i roli - edycja ich nie wymusza
    if (plcie.length && !plec && !edycja) brak.push("płeć");
    if (role.length && !wybraneRole.size && !edycja) brak.push("rola");
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
      const zmienionyMail = !edycja || mail !== edycja.email.trim().toLowerCase();
      if (mail && zmienionyMail && (await getContactByEmail(mail))) {
        throw new Error(`Kontakt z adresem ${mail} już jest w bazie.`);
      }

      if (edycja) {
        const zmiany: Record<string, unknown> = {
          Imi_x0119_: imie.value.trim(),
          Nazwisko: nazwisko.value.trim(),
          Email: mail,
          Phone: telefon.value.trim(),
          Stanowisko: stanowisko.value.trim(),
          UWAGI: uwagi.value.trim(),
          Sta_x0142_awysy_x0142_ka: stala.checked,
          Tymczasowawysy_x0142_ka: tymczasowa.checked,
          Status: niePracuje!.checked,
          "RolawMEDDIC@odata.type": "Collection(Edm.String)",
          RolawMEDDIC: [...wybraneRole],
        };
        if (plec) zmiany["P_x0142_e_x0107_"] = plec;
        const pola = await updateTicketFields(edycja.id, zmiany, edycja.etag || undefined, "Contacts");
        komunikat.textContent = "Zmiany zapisane";
        komunikat.className = "crm-komunikat crm-ok";
        poZapisie({ id: edycja.id, fields: { ...pola, id: edycja.id } });
        return;
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

    // szanse sprzedazy albo TECS tej firmy - przelacznik nad lista
    const przel = el("div");
    root().insertBefore(przel, s);
    segment(przel, ["Sprzedaż", "TECS"], szanseRodzaj, (v) => {
      szanseRodzaj = v as typeof szanseRodzaj;
      void pokazSzanse();
    });

    const box = el("div");
    root().appendChild(box);
    if (szanseRodzaj === "TECS") await zakladkaTecs(box, s, c, false);
    else await zakladkaSzanse(box, s, c, () => void pokazSzanse(), true);
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

// kosz przy kontakcie - tylko dla administratora CRM, jak w Power Apps
const ADMIN_KONTAKTOW = "mnowakowski@mesco.com.pl";

/** Dane firmy do formularza kontaktu - pelne, gdy firma jest otwarta, inaczej z listy. */
function firmaDoFormularza(c: Company): CompanyFull {
  return firmaPelna && firmaPelna.guid === c.guid ? firmaPelna : {
    id: "", guid: c.guid, skrocona: c.nazwa, pelna: "", nip: "", ulica: c.ulica,
    miasto: c.miasto, kod: c.kod, status: "", dataNDA: "", opiekun: c.opiekun,
    tecsH: 0, supportDay: 0, preSalesH: 0, uwagi: "", domena: "",
  };
}

/** Mala ikona-przycisk karty kontaktu (galeria Contacts z Power Apps). */
function ikonaAkcji(ik: keyof typeof IKONY, tytul: string, klasa = ""): HTMLButtonElement {
  const b = el("button", `crm-kk-akcja ${klasa}`.trim()) as HTMLButtonElement;
  b.title = tytul;
  b.setAttribute("aria-label", tytul);
  b.appendChild(ikona(ik));
  return b;
}

/** Wiersz "ikona  wartosc  [kopiuj]" - mail albo telefon. */
function liniaKontaktu(ik: keyof typeof IKONY, wartosc: string, brak: string, kopiujTytul: string): {
  linia: HTMLElement;
  ikonaEl: HTMLElement;
} {
  const linia = el("div", "crm-kk-linia");
  const ikonaEl = el("span", "crm-kk-ik");
  ikonaEl.appendChild(ikona(ik));
  linia.append(ikonaEl, el("span", wartosc ? "crm-kk-wartosc" : "crm-kk-wartosc crm-kk-brak", wartosc || brak));
  if (wartosc) {
    const kop = ikonaAkcji("kopiuj", kopiujTytul, "crm-kk-kopiuj");
    kop.onclick = async (e) => {
      e.stopPropagation();
      const ok = await kopiuj(wartosc);
      blysk(linia, ok ? "skopiowano" : "nie udało się");
    };
    linia.appendChild(kop);
  }
  return { linia, ikonaEl };
}

function kartaKontaktu(k: Contact, c: Company, poUsunieciu?: () => void, linkFirmy = false): HTMLElement {
  const card = el("div", k.aktywny ? "crm-kontakt-card crm-kk" : "crm-kontakt-card crm-kk crm-kk-byly");

  // --- naglowek: inicjaly, imie i nazwisko, stanowisko, edycja ---
  const top = el("div", "crm-kk-top");
  const inicjaly = ((k.imie || k.imieNazwisko).charAt(0) + (k.nazwisko || k.imieNazwisko.split(" ")[1] || "").charAt(0))
    .toUpperCase();
  top.appendChild(el("span", "crm-kk-avatar", inicjaly || "?"));
  const kto = el("div", "crm-kk-kto");
  kto.appendChild(el("div", "crm-kk-nazwa", k.imieNazwisko || "(bez nazwiska)"));
  kto.appendChild(el("div", k.stanowisko ? "crm-kk-stanowisko" : "crm-kk-stanowisko crm-kk-brak", k.stanowisko || "Brak stanowiska"));
  top.appendChild(kto);
  const edytuj = ikonaAkcji("pioro", "Edytuj kontakt");
  top.appendChild(edytuj);
  card.appendChild(top);

  // wyszukiwarka kontaktow - klik w firme prowadzi do pelnego widoku firmy
  if (linkFirmy && k.companyGuid) {
    const doFirmy = el("button", "crm-kontakt-firma", `↗ ${k.firmaId || "firma"}`) as HTMLButtonElement;
    doFirmy.title = "Otwórz firmę";
    doFirmy.onclick = () => {
      zakladkaFirmy = "Kontakty";
      void pokazFirma(c);
    };
    card.appendChild(doFirmy);
  }

  // --- pastylki: rola, nie pracuje, MESsenger ---
  const pigulki = el("div", "crm-kk-pigulki");
  if (k.role.length) pigulki.appendChild(pastylka(k.role.join(", "), "#5B6CDB", "#EEF0FC"));
  if (!k.aktywny) pigulki.appendChild(pastylka("Nie pracuje", "#DC2626", "#FEE2E2"));
  if (k.stala || k.tymczasowa) {
    const m = pastylka(`MESsenger${k.ileMessengerow > 0 ? " x" + k.ileMessengerow : ""}`, "#16A34A", "#E6F6EE");
    m.title = k.stala ? "Stała wysyłka" : "Tymczasowa wysyłka";
    pigulki.appendChild(m);
  }
  if (pigulki.childElementCount) card.appendChild(pigulki);

  card.appendChild(el("div", "crm-kk-kreska"));

  // --- mail i telefon, kazdy z kopiowaniem; ikona telefonu dzwoni ---
  card.appendChild(liniaKontaktu("mail", k.email, "brak e-maila", "Kopiuj adres").linia);
  const tel = liniaKontaktu("telefon", k.telefon, "brak telefonu", "Kopiuj numer");
  if (k.telefon) {
    const zadzwon = document.createElement("a");
    zadzwon.className = "crm-kk-ik crm-kk-dzwon";
    zadzwon.href = `tel:${k.telefon.replace(/[^\d+]/g, "")}`;
    zadzwon.title = "Zadzwoń";
    zadzwon.appendChild(ikona("telefon"));
    tel.ikonaEl.replaceWith(zadzwon);
  }
  card.appendChild(tel.linia);

  // --- akcje ---
  const akcje = el("div", "crm-kk-akcje");
  const nowyMail = ikonaAkcji("mail", k.email ? "Napisz e-mail" : "Brak adresu e-mail", "crm-kk-niebieski");
  nowyMail.disabled = !k.email;
  nowyMail.onclick = () =>
    Office.context.mailbox.displayNewMessageForm({
      toRecipients: [k.email],
    });
  const zgloszenie = ikonaAkcji("klucz", "Nowe zgłoszenie Support", "crm-kk-pomaranczowy");
  const szybka = ikonaAkcji("blyskawica", "Szybka akcja — notatka", "crm-kk-zolty");
  const doOutlooka = ikonaAkcji("dodajOsobe", "Dodaj do kontaktów Outlook", "crm-kk-niebieski");
  const usun = ikonaAkcji("kosz", "Usuń kontakt", "crm-kk-czerwony crm-kk-prawo");
  usun.hidden = true;
  akcje.append(nowyMail, zgloszenie, szybka, doOutlooka, usun);
  card.appendChild(akcje);
  const komunikat = el("span", "crm-komunikat");
  card.appendChild(komunikat);

  zgloszenie.onclick = () => void formularzReczny(card, zgloszenie, komunikat, k, c);
  szybka.onclick = () => void formularzNotatki(card, szybka, komunikat, k, c);

  doOutlooka.onclick = async () => {
    doOutlooka.disabled = true;
    komunikat.className = "crm-komunikat";
    komunikat.textContent = "Dodawanie do kontaktów Outlook...";
    try {
      const nowy = await dodajDoKontaktowOutlook({
        imie: k.imie || rozbijNazwe(k.imieNazwisko).imie,
        nazwisko: k.nazwisko || rozbijNazwe(k.imieNazwisko).nazwisko,
        email: k.email,
        telefon: k.telefon,
        firma: k.firmaId || c.nazwa,
        stanowisko: k.stanowisko,
      });
      komunikat.textContent = nowy
        ? `Dodano ${k.imieNazwisko} do kontaktów Outlook`
        : `${k.imieNazwisko} jest już w kontaktach Outlook`;
      komunikat.className = nowy ? "crm-komunikat crm-ok" : "crm-komunikat";
    } catch (e: any) {
      komunikat.textContent = e?.message ?? String(e);
      komunikat.className = "crm-komunikat crm-blad";
    } finally {
      doOutlooka.disabled = false;
    }
  };

  // --- edycja: formularz w miejscu karty ---
  edytuj.onclick = () => {
    const box = el("div", "crm-nowe");
    card.replaceWith(box);
    void formularzKontaktu(box, firmaDoFormularza(c), { imie: k.imie, nazwisko: k.nazwisko, email: k.email }, (nowy) => {
      Object.assign(k, toContact(nowy));
      kontaktCache.clear();
      wyczyscCacheKontaktow();
      box.replaceWith(kartaKontaktu(k, c, poUsunieciu, linkFirmy));
    }, () => box.replaceWith(card), k);
  };

  // --- usuwanie: tylko administrator, z potwierdzeniem ---
  void getMojMail().then((m) => (usun.hidden = m.trim().toLowerCase() !== ADMIN_KONTAKTOW), () => undefined);
  usun.onclick = () => {
    usun.disabled = true;
    const pyt = el("div", "crm-kk-potwierdz");
    pyt.appendChild(el("span", undefined, `Usunąć ${k.imieNazwisko || "kontakt"} z CRM?`));
    const tak = el("button", "crm-btn crm-btn-ostrz", "Usuń") as HTMLButtonElement;
    const nie = el("button", "crm-btn crm-btn-cichy", "Anuluj") as HTMLButtonElement;
    pyt.append(tak, nie);
    card.appendChild(pyt);
    nie.onclick = () => {
      pyt.remove();
      usun.disabled = false;
    };
    tak.onclick = async () => {
      tak.disabled = nie.disabled = true;
      try {
        await deleteListItem("Contacts", k.id);
        kontaktCache.clear();
        wyczyscCacheKontaktow();
        card.remove();
        poUsunieciu?.();
      } catch (e: any) {
        pyt.remove();
        usun.disabled = false;
        komunikat.textContent = e?.message ?? String(e);
        komunikat.className = "crm-komunikat crm-blad";
      }
    };
  };

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
  // karta firmy jak hdCard w Power Apps: nazwa i opiekun, adres i NIP, status, NDA, godziny
  const karta = el("div", "crm-firma-karta crm-hd");
  box.appendChild(karta);

  const gora = el("div", "crm-hd-gora");
  const bud = el("span", "crm-hd-bud");
  bud.appendChild(ikona("firma"));
  const kto = el("div", "crm-hd-kto");
  const pelna = f?.pelna || f?.skrocona || c.nazwa;
  const nazwaEl = el("div", "crm-hd-nazwa", pelna);
  nazwaEl.title = pelna;
  kto.appendChild(nazwaEl);
  const opiekunNazwa = f?.opiekun || c.opiekun;
  if (opiekunNazwa) {
    const op = pastylka(`Opiekun: ${opiekunNazwa}`, "#5B6CDB", "#EEF0FC", "osoba");
    op.classList.add("crm-hd-opiekun");
    kto.appendChild(op);
  }
  gora.append(bud, kto);
  karta.appendChild(gora);

  // reszta karty zwinieta - zajmowala za duzo miejsca nad zakladkami
  const szczegoly = el("div", "crm-hd-szczegoly");
  const rozwin = el("button", "crm-hd-rozwin") as HTMLButtonElement;
  const ustawRozwin = () => {
    szczegoly.hidden = !szczegolyFirmy;
    rozwin.textContent = szczegolyFirmy ? "▴ Ukryj szczegóły" : "▾ Pokaż szczegóły";
    rozwin.setAttribute("aria-expanded", String(szczegolyFirmy));
  };
  rozwin.onclick = () => {
    szczegolyFirmy = !szczegolyFirmy;
    ustawRozwin();
  };
  kto.appendChild(rozwin);

  // --- adres (nawigacja) i NIP (kopiowanie) ---
  const chipy = el("div", "crm-hd-chipy");
  if (f?.ulica || f?.miasto) {
    const u = (f.ulica || "").trim().replace(/^ul\.?\s*/i, "");
    const tekst = [u ? `ul. ${u}` : "", [f.kod, f.miasto].filter(Boolean).join(" ")].filter(Boolean).join(", ");
    const adres = przyciskIk("crm-hd-chip crm-hd-adres", "pinezka", tekst);
    adres.title = "Nawiguj w Google Maps";
    adres.onclick = () => {
      const cel = encodeURIComponent(`${(f.ulica || "").trim()}, ${f.kod} ${(f.miasto || "").trim()}, Polska`);
      window.open(`https://www.google.com/maps/dir/?api=1&travelmode=driving&destination=${cel}`, "_blank");
    };
    chipy.appendChild(adres);
  }
  if (f?.nip) {
    const czysty = f.nip.toUpperCase().replace(/[\s-]/g, "").replace(/^PL/, "");
    const widoczny =
      czysty.length === 10
        ? `${czysty.slice(0, 3)}-${czysty.slice(3, 5)}-${czysty.slice(5, 7)}-${czysty.slice(7)}`
        : czysty;
    const nipBtn = przyciskIk("crm-hd-chip crm-hd-nip", "hasz", `NIP ${widoczny}`);
    nipBtn.appendChild(ikona("kopiuj"));
    nipBtn.title = "Kopiuj NIP";
    nipBtn.onclick = async () => {
      const ok = await kopiuj(czysty);
      blysk(nipBtn, ok ? `skopiowano NIP: ${czysty}` : "nie udało się");
    };
    chipy.appendChild(nipBtn);
  }
  if (chipy.childElementCount) szczegoly.appendChild(chipy);

  // ostrzezenie o TECS zostaje na wierzchu - ma byc widoczne bez rozwijania
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
    p.classList.add("crm-hd-alert");
    karta.appendChild(p);
  }
  karta.appendChild(szczegoly);
  ustawRozwin();

  // --- kafelki: status i NDA po lewej, wykorzystane godziny po prawej ---
  const kafle = el("div", "crm-hd-kafle");
  const kafelek = (ik: keyof typeof IKONY, kolIk: string, tloIk: string, etykieta: string, wartosc: string, kolor: string) => {
    const k = el("div", "crm-hd-kafel");
    const i = el("span", "crm-hd-kafel-ik");
    i.style.color = kolIk;
    i.style.background = tloIk;
    i.appendChild(ikona(ik));
    const t = el("div", "crm-hd-kafel-txt");
    const v = el("div", "crm-hd-kafel-val", `● ${wartosc}`);
    v.style.color = kolor;
    t.append(el("div", "crm-hd-kafel-lbl", etykieta), v);
    k.append(i, t);
    return k;
  };
  const kolStatusu: Record<string, string> = {
    Klient: "#16A34A", Prospekt: "#D97706", Prospect: "#D97706", GSS: "#2F6FD6", Nieaktywny: "#9CA3AF",
  };
  const st = f?.status || "";
  const dniNDA = dniDo(f?.dataNDA ?? "");
  const kolNDA = !f?.dataNDA ? "#9CA3AF" : dniNDA !== null && dniNDA < 0 ? "#DC2626"
    : dniNDA !== null && dniNDA <= 30 ? "#D97706" : "#16A34A";
  const lewa = el("div", "crm-hd-lewa");
  lewa.append(
    kafelek("ludzie", "#7C3AED", "#F1EAFE", "Status", st || "brak", kolStatusu[st] ?? "#5B6CDB"),
    kafelek("tarczaOk", "#16A34A", "#E6F6EE", "NDA do", f?.dataNDA ? dataPL(f.dataNDA) : "brak", kolNDA)
  );

  const godz = el("div", "crm-hd-godz");
  godz.appendChild(poleIk("crm-hd-godz-lbl", "zegar", "Wykorzystane godziny"));
  ([
    ["TECS", f?.tecsH ?? 0, "#2F6FD6"],
    ["Support Day", f?.supportDay ?? 0, "#D97706"],
    ["PreSales", f?.preSalesH ?? 0, "#5B6CDB"],
  ] as [string, number, string][]).forEach(([lbl, val, kol]) => {
    const r = el("div", "crm-hd-godz-rzad");
    const kolor = val === 0 ? "#9CA3AF" : kol;
    const kropka = el("span", "crm-hd-kropka", "●");
    kropka.style.color = kolor;
    const v = el("span", "crm-hd-godz-val", godziny(val));
    v.style.color = kolor;
    r.append(kropka, el("span", "crm-hd-godz-nazwa", lbl), v);
    godz.appendChild(r);
  });
  kafle.append(lewa, godz);
  szczegoly.appendChild(kafle);

  // --- zakladki jak pasek w Power Apps: ikona + nazwa, akcje wypelnione na koncu ---
  const pasek = el("div", "crm-zakladki");
  const nazwy: [Zakladka, string, keyof typeof IKONY][] = [
    ["Kontakty", "Kontakty", "osoba"],
    ["Notatki", "Notatki", "dokument"],
    ["Supporty", "Supporty", "sluchawki"],
    ["Szkolenia", "Szkolenia", "czapka"],
    ["Szanse", "Szanse", "cel"],
    ["TECS", "TECS", "tarcza"],
    ["Posiadane", "Klient posiada", "paczka"],
    ["Uwagi", "Uwagi", "dymek"],
    ["Dokumenty", "Dokumenty", "folder"],
  ];
  // Szanse i TECS w jednej grupie - zawsze obok siebie, nawet gdy pasek sie zawija
  const sprzedaz = el("div", "crm-zakladki-para");
  nazwy.forEach(([z, etykieta, ik]) => {
    const b = przyciskIk("crm-zakladka", ik, etykieta);
    if (z === zakladkaFirmy) b.classList.add("crm-zakladka-akt");
    b.onclick = () => {
      zakladkaFirmy = z;
      void pokazFirma(c);
    };
    if (z === "Szanse" || z === "TECS") {
      sprzedaz.appendChild(b);
      if (z === "Szanse") pasek.appendChild(sprzedaz);
    } else pasek.appendChild(b);
  });


  // przyciski akcji w osobnym rzedzie pod zakladkami - zakladki maja rowna siatke
  const akcjePaska = el("div", "crm-zakladki-akcje");

  // kody demo laduja w TECS / Demo z otwartym formularzem
  const demo = przyciskIk("crm-zakladka", "licencja", "Kody demo");
  demo.onclick = () => {
    zakladkaFirmy = "TECS";
    tecsTryb = "Demo";
    otworzKodyDemo = true;
    void pokazFirma(c);
  };
  // akcje zalezne od zakladki: w TECS tylko TECS i kody demo, w Szansach tylko nowa szansa,
  // a na zakladkach bez zwiazku ze sprzedaza zadnych - mniej balaganu
  const bezAkcji = (["Supporty", "Szkolenia", "Uwagi", "Dokumenty"] as Zakladka[]).includes(zakladkaFirmy);
  if (zakladkaFirmy !== "Szanse" && !bezAkcji) akcjePaska.appendChild(demo);

  const nowaSzansa = przyciskIk("crm-zakladka crm-zakladka-glowna", "nowaSzansa", "Nowy Opps Soft");
  nowaSzansa.onclick = () => {
    zakladkaFirmy = "Szanse";
    szanseZamkniete = false;
    otworzNowaSzanse = true;
    void pokazFirma(c);
  };
  if (zakladkaFirmy !== "TECS" && !bezAkcji) akcjePaska.appendChild(nowaSzansa);

  const nowyTecs = przyciskIk("crm-zakladka crm-zakladka-glowna", "nowaSzansa", "Nowy Opps TECS");
  nowyTecs.onclick = () => {
    zakladkaFirmy = "TECS";
    tecsTryb = "Bieżące";
    otworzNowyTecs = true;
    void pokazFirma(c);
  };
  if (zakladkaFirmy !== "Szanse" && !bezAkcji) akcjePaska.appendChild(nowyTecs);
  box.appendChild(pasek);
  if (akcjePaska.childElementCount) box.appendChild(akcjePaska);

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
    if (!otworzNowyTecs && !otworzKodyDemo) tecsTryb = "Bieżące";
    if (!otworzNowaSzanse) szanseZamkniete = false;
    autoTecs = autoSzanse = true;
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
      case "Dokumenty": await zakladkaDokumenty(tresc, s, c); break;
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
  // po nazwisku, jak galeria w Power Apps
  const wszyscy = items
    .map(toContact)
    .sort((a, b) =>
      (a.nazwisko || a.imieNazwisko).localeCompare(b.nazwisko || b.imieNazwisko, "pl") ||
      a.imie.localeCompare(b.imie, "pl")
    );

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
    const f = firmaDoFormularza(c);
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
    widoczni.forEach((k) => lista.appendChild(kartaKontaktu(k, c, () => {
      wszyscy.splice(wszyscy.indexOf(k), 1);
      rysuj();
    })));
  };

  let timer: any;
  szukaj.oninput = () => {
    clearTimeout(timer);
    timer = setTimeout(rysuj, 150);
  };

  rysuj();
}

// filtry zakladki Supporty - pamietane w obrebie firmy
let filtrSupportow = { firma: "", obslugujacy: "", dlaKogo: "" };

async function zakladkaSupporty(box: HTMLElement, s: HTMLElement, c: Company) {
  // wiecej niz 50 - filtrujemy po stronie dodatku, wiec potrzebna jest cala historia firmy
  const wszystkie = (await getTicketsFiltered({ companyGuid: c.guid, top: 999 })).map(toTicket);
  if (!wszystkie.length) {
    s.textContent = "Zgłoszenia — 0";
    return pusto(box, "Brak zgłoszeń.");
  }
  if (filtrSupportow.firma !== c.guid) filtrSupportow = { firma: c.guid, obslugujacy: "", dlaKogo: "" };

  // opcje z liczba zgloszen
  type Opcja = [string, { nazwa: string; ile: number }];
  const opcje = (zgl: Ticket[], klucz: (t: Ticket) => string, nazwa: (t: Ticket) => string): Opcja[] => {
    const m = new Map<string, { nazwa: string; ile: number }>();
    zgl.forEach((t) => {
      const k = klucz(t);
      if (!k) return;
      const o = m.get(k) ?? { nazwa: nazwa(t) || k, ile: 0 };
      o.ile++;
      m.set(k, o);
    });
    return [...m.entries()].sort((a, b) => a[1].nazwa.localeCompare(b[1].nazwa, "pl"));
  };
  // obslugujacy tylko po nazwie (Dodatkowy_opiekun) - MailCreator bywa rozny dla tej samej osoby
  const kluczObsl = (t: Ticket) => t.opiekun.trim().replace(/\s+/g, " ").toLowerCase();
  // osoba tez tylko po nazwie (Osoba) - mail zglaszajacego bywa rozny lub pusty
  const kluczOsoby = (t: Ticket) => t.osoba.trim().replace(/\s+/g, " ").toLowerCase();

  const filtry = el("div", "crm-sup-filtry");
  const wybor = (etykieta: string) => {
    const k = el("div");
    k.appendChild(el("label", "crm-label", etykieta));
    const sel = document.createElement("select");
    sel.className = "crm-select";
    k.appendChild(sel);
    filtry.appendChild(k);
    return sel;
  };
  /** Opcje listy od nowa; wybor zostaje, jesli nadal wystepuje. */
  const wypelnij = (sel: HTMLSelectElement, wszyscy: string, zgl: Ticket[], lista: Opcja[], wartosc: string) => {
    sel.innerHTML = "";
    ([["", { nazwa: wszyscy, ile: zgl.length }]] as Opcja[]).concat(lista).forEach(([v, o]) => {
      const opt = document.createElement("option");
      opt.value = v;
      opt.textContent = `${o.nazwa} (${o.ile})`;
      sel.appendChild(opt);
    });
    sel.value = lista.some(([v]) => v === wartosc) ? wartosc : "";
  };
  const obslugujacy = wybor("Obsługujący");
  const dlaKogo = wybor("Dla kogo");

  // filtr kaskadowy: kazda lista pokazuje tylko osoby z zgloszen pasujacych do drugiego filtra
  const odswiezOpcje = () => {
    const dlaObsl = wszystkie.filter((t) => !filtrSupportow.dlaKogo || kluczOsoby(t) === filtrSupportow.dlaKogo);
    const dlaOsob = wszystkie.filter((t) => !filtrSupportow.obslugujacy || kluczObsl(t) === filtrSupportow.obslugujacy);
    wypelnij(obslugujacy, "wszyscy", dlaObsl, opcje(dlaObsl, kluczObsl, (t) => t.opiekun), filtrSupportow.obslugujacy);
    wypelnij(dlaKogo, "wszystkie osoby", dlaOsob, opcje(dlaOsob, kluczOsoby, (t) => t.osoba), filtrSupportow.dlaKogo);
  };
  odswiezOpcje();
  const wyczysc = el("button", "crm-btn crm-btn-cichy crm-sup-wyczysc", "Wyczyść filtry") as HTMLButtonElement;
  filtry.appendChild(wyczysc);
  box.appendChild(filtry);

  const lista = el("div");
  box.appendChild(lista);
  const NA_STRONE = 50;

  const rysuj = () => {
    filtrSupportow.obslugujacy = obslugujacy.value;
    filtrSupportow.dlaKogo = dlaKogo.value;
    const widoczne = wszystkie.filter((t) =>
      (!obslugujacy.value || kluczObsl(t) === obslugujacy.value) &&
      (!dlaKogo.value || kluczOsoby(t) === dlaKogo.value)
    );
    const filtrowane = !!(obslugujacy.value || dlaKogo.value);
    wyczysc.hidden = !filtrowane;
    s.textContent = filtrowane
      ? `Zgłoszenia — ${widoczne.length} z ${wszystkie.length}`
      : `Zgłoszenia — ${wszystkie.length}`;
    lista.innerHTML = "";
    if (!widoczne.length) return pusto(lista, "Brak zgłoszeń dla wybranych filtrów.");

    let pokazane = 0;
    const wiecej = el("button", "crm-btn crm-btn-cichy crm-sup-wiecej") as HTMLButtonElement;
    const dorysuj = () => {
      widoczne.slice(pokazane, pokazane + NA_STRONE).forEach((t) => lista.insertBefore(karta(t), wiecej));
      pokazane = Math.min(pokazane + NA_STRONE, widoczne.length);
      wiecej.textContent = `Pokaż więcej (${widoczne.length - pokazane})`;
      wiecej.hidden = pokazane >= widoczne.length;
    };
    lista.appendChild(wiecej);
    wiecej.onclick = dorysuj;
    dorysuj();
  };
  const zmiana = () => {
    filtrSupportow.obslugujacy = obslugujacy.value;
    filtrSupportow.dlaKogo = dlaKogo.value;
    odswiezOpcje();
    rysuj();
  };
  obslugujacy.onchange = zmiana;
  dlaKogo.onchange = zmiana;
  wyczysc.onclick = () => {
    obslugujacy.value = dlaKogo.value = "";
    zmiana();
  };
  rysuj();
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

async function zakladkaTecs(box: HTMLElement, s: HTMLElement, c: Company, zDemo = true) {
  // widok maila handlowca pokazuje tylko szanse TECS, bez kodow demo
  if (!zDemo && tecsTryb === "Demo") tecsTryb = "Bieżące";
  segment(box, zDemo ? ["Bieżące", "Archiwalne", "Demo"] : ["Bieżące", "Archiwalne"], tecsTryb, (v) => {
    tecsTryb = v as typeof tecsTryb;
    void odswiezTecs(c);
  });

  // notatki firmy jednym zapytaniem - przy TECS-ach grupujemy je po Opps_ID = TECSGUID
  const notatki = getByCompany("Notes", c.guid, { top: 500 }).then((it) => it.map(toNota));
  notatki.catch((e) => console.error("notatki TECS:", e));
  const ja = getMojaNazwa().then((n) => n.trim().toLowerCase(), () => "");

  // nowa szansa TECS - z paska firmy, po wygranej szansie albo przyciskiem nad lista w widoku maila
  const nowyBox = el("div");
  box.appendChild(nowyBox);
  const otworzFormTecs = (zSzansy: Szansa | null = null) => {
    nowyBtn.disabled = true;
    void formularzNowegoTecs(nowyBox, c, zSzansy, () => {
      tecsTryb = "Bieżące";
      void odswiezTecs(c);
    }, () => {
      nowyBox.innerHTML = "";
      if (!zDemo) nowyBox.appendChild(nowyBtn);
      nowyBtn.disabled = false;
    });
  };
  const nowyBtn = przyciskIk("crm-btn crm-tecs-nowy crm-nowa-szansa", "nowaSzansa", "Nowy Opps TECS");
  nowyBtn.onclick = () => {
    nowyBox.innerHTML = "";
    otworzFormTecs();
  };
  if (otworzNowyTecs) {
    otworzNowyTecs = false;
    otworzFormTecs(nowyTecsZSzansy);
    nowyTecsZSzansy = null;
  } else if (!zDemo) nowyBox.appendChild(nowyBtn); // w karcie firmy jest "Nowy Opps TECS" w pasku

  // nowe kody demo - formularz nad lista, otwierany kafelkiem pod zakladkami
  if (otworzKodyDemo) {
    otworzKodyDemo = false;
    const miejsce = el("div");
    box.appendChild(miejsce);
    void formularzKodowDemo(miejsce, c, () => void pokazFirma(c), () => miejsce.remove());
  }

  const items = await getByCompany("Opps_TECS", c.guid, { top: 100 });
  const wszystkie = items.map(toTecs);
  if (autoTecs) {
    autoTecs = false;
    if (tecsTryb === "Bieżące" && !wszystkie.some((t) => ETAPY_TECS_BIEZACE.includes(t.etap))) {
      tecsTryb = "Archiwalne";
      zaznaczSegment(box, tecsTryb);
    }
  }
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
      const rzadP = el("div", "crm-akcje crm-szansa-akcje crm-karta-akcje");
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
      const prod = przyciskIk("crm-btn crm-btn-cichy", "paczka", "Produkty");
      rzadP.appendChild(prod);
      k.appendChild(rzadP);
      k.appendChild(panelSzcz);
      const panel = el("div");
      k.appendChild(panel);
      sekcjaNotatek(k, rzadP, {
        id: t.id, guid: t.guid, nazwa: t.nazwa, firma: t.firmaId, companyGuid: c.guid, etap: t.etap,
      }, c, notatki, ja, "Brak notatek do tego TECS-u.");
      sekcjaZalacznikow(k, rzadP, "Opps_TECS", t.id, folderFirmy(c));
      przyciskGenerujOferte(rzadP);
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

/** Po zapisie TECS-u: odswiezamy ten widok, w ktorym jestesmy - szanse maila albo karte firmy. */
function odswiezTecs(c: Company): Promise<void> {
  if (widok === "szanse") {
    szanseRodzaj = "TECS";
    return pokazSzanse();
  }
  return pokazFirma(c);
}

// ---------------------------------------------------------------- szczegoly TECS (conTECSView)

// etapy TECS (Etap1, tekst) - sciezka zwykla i osobna dla Tech Poola, jak w Power Apps
const ETAPY_TECS = ["Przygotowanie oferty", "Oferta wysłana", "Zamówione", "Wystawiona faktura"];
const ETAPY_TP = ["Tech Pool", "Koniec Tech Poola"];
const REZYGNACJA_TECS = "Rezygnacja z TECS";
// dodawanie, edycja i usuwanie produktow TECS-u - tylko przed wyslaniem zamowienia
const ETAPY_EDYCJI_PRODUKTOW_TECS = ["Przygotowanie oferty", "Oferta wysłana"];
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
      await odswiezTecs(c);
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
      void odswiezTecs(c);
    }, () => {
      edycjaBox.innerHTML = "";
      rzad.style.display = "";
    }, null, t);
  };
  edytuj.onclick = () => {
    rzad.style.display = "none";
    void formularzNowegoTecs(edycjaBox, c, null, () => {
      otwartySzczegolTecs = t.id;
      void odswiezTecs(c);
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
    // produkty zmieniamy tylko przed zamowieniem - potem lista jest zamrozona
    pracownik = p && ETAPY_EDYCJI_PRODUKTOW_TECS.includes(t.etap);
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
  if (!ETAPY_EDYCJI_PRODUKTOW_TECS.includes(t.etap)) {
    dodajBox.appendChild(el("div", "crm-nowe-info",
      `Etap „${t.etap || "—"}” — produkty można zmieniać tylko na etapach: ${ETAPY_EDYCJI_PRODUKTOW_TECS.join(", ")}.`));
  }

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
    if (!pracownik) return;
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
    edycja ? "Edycja szansy TECS" : renewal ? "Renewal TECS" : "Nowy Opps TECS"));
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

// ---------------------------------------------------------------- nowa szansa sprzedazy (conOppsForm)

// domyslne Źródło szansy wg zalogowanego - jak DefaultSelectedItems w Power Apps
// (inne niz przy nowej firmie: tam ZRODLO_HANDLOWCA)
const ZRODLO_SZANSY: Record<string, string> = {
  "Mateusz Nowakowski": "MNow",
  "Łukasz Borgul": "LBor",
  "Przemysław Draga": "PDra",
  "Grzegorz Woźniak": "PDra",
};

/** Nowa szansa sprzedazy Soft w liscie Opps - odpowiednik FormNowyOpps z Power Apps. */
async function formularzNowejSzansy(
  miejsce: HTMLElement,
  c: Company,
  poZapisie: (nowa?: Szansa) => void,
  anuluj: () => void,
  edycja: Szansa | null = null // edycja istniejacej szansy - bez statusu i etapu (te w panelu "Status i etap")
) {
  const box = el("div", "crm-nowe crm-demo crm-ns");
  miejsce.appendChild(box);
  const naglowek = el("div", "crm-ns-naglowek");
  const ik = el("span", "crm-ns-ik");
  ik.appendChild(ikona("wykres"));
  const tytul = el("div");
  tytul.append(
    el("div", "crm-ns-tytul", edycja ? "Edycja szansy sprzedaży" : "Nowa szansa sprzedaży – Soft"),
    el("div", "crm-nowe-info", edycja ? `${edycja.firma || c.nazwa} — ${edycja.nazwa}` : firmaPelna?.pelna || firmaPelna?.skrocona || c.nazwa)
  );
  naglowek.append(ik, tytul);
  box.appendChild(naglowek);

  const skrocona = firmaPelna?.skrocona || c.nazwa;
  const sekcja = (t: string) => box.appendChild(el("div", "crm-nf-sekcja", t));
  const wybor = (etykieta: string) => {
    box.appendChild(el("label", "crm-label", etykieta));
    const sel = document.createElement("select");
    sel.className = "crm-select";
    box.appendChild(sel);
    return sel;
  };
  const wypelnij = (sel: HTMLSelectElement, opcje: string[], domyslna: string, pusta = "wybierz…") => {
    sel.innerHTML = "";
    ["", ...opcje].forEach((o) => {
      const opt = document.createElement("option");
      opt.value = o;
      opt.textContent = o || pusta;
      sel.appendChild(opt);
    });
    // wielkosc liter w opcjach kolumny bywa inna niz w Power Apps ("New logo" / "New Logo")
    sel.value = opcje.find((o) => o.toLowerCase() === domyslna.toLowerCase()) ?? "";
  };

  // --- SZANSA ---
  sekcja("Szansa");
  const nazwa = poleFormularza(box, "Nazwa szansy");
  const d = new Date();
  nazwa.value = `Q${Math.ceil((d.getMonth() + 1) / 3)}_${d.getFullYear()}_${skrocona}`;
  const wartosc = poleFormularza(box, "Wartość (EUR)", "number");
  wartosc.min = "0";
  wartosc.step = "0.01";
  wartosc.placeholder = "0";
  const etap = wybor("Etap *");
  const status = wybor("Status *");
  const typ = wybor("Typ biznesu");

  // --- TERMINY ---
  sekcja("Terminy");
  const otwarcie = poleFormularza(box, "Data otwarcia", "date");
  otwarcie.value = dzisISO();
  const zamkniecie = poleFormularza(box, "Spodziewane zamknięcie", "date");
  const zamkHint = el("div", "crm-nowe-info", `np. ${dataPL(plusDni(dzisISO(), 183))} (za pół roku)`);
  box.appendChild(zamkHint);
  const dr = poleFormularza(box, "Data DR", "date");

  // --- PRZYPISANIE ---
  sekcja("Przypisanie");
  const opiekun = wybor("Opiekun Oppsa *");
  const zrodlo = wybor("Źródło *");
  const rodzaj = wybor("Rodzaj biznesu *");

  // --- Champion: z kontaktow firmy albo wpisany ---
  box.appendChild(el("label", "crm-label", "Champion *"));
  const champion = document.createElement("input");
  champion.type = "text";
  champion.className = "crm-input";
  champion.placeholder = "Imię i nazwisko osoby decyzyjnej";
  box.appendChild(champion);
  const wyborOsoby = el("div", "crm-ns-osoby");
  const szukaj = document.createElement("input");
  szukaj.type = "search";
  szukaj.className = "crm-szukaj";
  szukaj.placeholder = "Szukaj osoby w kontaktach firmy";
  const listaOsob = el("div", "crm-ns-lista");
  wyborOsoby.append(szukaj, listaOsob);
  box.appendChild(wyborOsoby);

  box.appendChild(el("label", "crm-label", "Notatki"));
  const notatki = document.createElement("textarea");
  notatki.className = "crm-textarea";
  notatki.rows = 4;
  notatki.placeholder = "Kontekst szansy, potrzeby klienta, konkurencja…";
  box.appendChild(notatki);

  // --- przyciski ---
  const brakujace = el("div", "crm-podpowiedz");
  box.appendChild(brakujace);
  const zapisz = el("button", "crm-btn", edycja ? "Zapisz zmiany" : "Zapisz szansę") as HTMLButtonElement;
  const wstecz = el("button", "crm-btn crm-btn-cichy", "Anuluj") as HTMLButtonElement;
  const komunikat = el("span", "crm-komunikat");
  const rzad = el("div", "crm-akcje");
  rzad.append(zapisz, wstecz, komunikat);
  box.appendChild(rzad);
  wstecz.onclick = anuluj;

  function sprawdz(): boolean {
    const brak: string[] = [];
    if (!etap.value && !edycja) brak.push("Etap");
    if (!status.value && !edycja) brak.push("Status");
    if (!opiekun.value) brak.push("Opiekun");
    if (!zrodlo.value) brak.push("Źródło");
    if (!rodzaj.value) brak.push("Rodzaj biznesu");
    if (!champion.value.trim()) brak.push("Champion");
    const ok = brak.length === 0;
    zapisz.disabled = !ok;
    brakujace.textContent = ok ? "Wszystko gotowe do zapisu" : `Uzupełnij: ${brak.join(", ")}`;
    brakujace.classList.toggle("crm-ns-gotowe", ok);
    return ok;
  }
  [etap, status, opiekun, zrodlo, rodzaj].forEach((x) => (x.onchange = () => sprawdz()));
  champion.oninput = () => {
    rysujOsoby();
    sprawdz();
  };
  sprawdz();

  // --- opcje z definicji kolumn Opps, opiekunowie, kontakty ---
  let ja = "";
  let kluczZrodla = "";
  let idZrodel = new Map<string, string>(); // lookup: etykieta -> Id; pusta mapa = kolumna wyboru
  let kontakty: Contact[] = [];
  try {
    const [etapy, statusy, typy, rodzaje, osoby, mojaNazwa, kt, nazwaZrodla] = await Promise.all([
      getChoiceOptions("Opps", "Etap"),
      getChoiceOptions("Opps", "Status"),
      getChoiceOptions("Opps", "TypBiznesu"),
      getChoiceOptions("Opps", "Rodzajbiznesu"),
      getOpiekunowie().catch(() => [] as string[]),
      getMojaNazwa().catch(() => ""),
      getContactsByCompany(c.guid).catch(() => [] as ListItem[]),
      nazwaWewnetrzna("Opps", "Źródło"),
    ]);
    // firma bez zadnej szansy TECS = nowy klient; przy bledzie pobrania nie zgadujemy
    const tecsy = await getByCompany("Opps_TECS", c.guid, { top: 1 }).catch(() => null);
    ja = mojaNazwa.trim();
    wypelnij(etap, etapy, "Pierwszy kontakt", "Wybierz etap");
    wypelnij(status, statusy, "Aktywny");
    wypelnij(typ, typy, "Paid-up");
    wypelnij(rodzaj, rodzaje, tecsy && !tecsy.length ? "New Logo" : "", "Wybierz rodzaj");
    wypelnij(opiekun, ja && !osoby.includes(ja) ? [ja, ...osoby] : osoby, ja);

    // Źródło: kolumna wyboru, a gdyby byla lookupem (jak w Companies) - Id pozycji
    let zrodla = await getChoiceOptions("Opps", nazwaZrodla);
    if (!zrodla.length) {
      const lookup = await getOpcjeLookupu("Opps", "Źródło").catch(() => null);
      if (lookup) {
        idZrodel = new Map(lookup.opcje.map((o) => [o.etykieta, o.id]));
        zrodla = lookup.opcje.map((o) => o.etykieta);
      }
    }
    kluczZrodla = nazwaZrodla;
    wypelnij(zrodlo, zrodla, ZRODLO_SZANSY[ja] ?? "");

    if (edycja) wypelnijZSzansy(edycja, zrodla);

    kontakty = kt.map(toContact).filter((x) => x.aktywny)
      .sort((a, b) => (a.nazwisko || a.imieNazwisko).localeCompare(b.nazwisko || b.imieNazwisko, "pl"));
  } catch (e: any) {
    komunikat.textContent = `Nie pobrano opcji: ${e?.message ?? e}`;
    komunikat.className = "crm-komunikat crm-blad";
  }

  // edycja: status i etap zmienia sie w panelu "Status i etap" (zamkniecie wymaga daty / powodu)
  if (edycja) {
    [etap, status].forEach((x) => {
      x.hidden = true;
      (x.previousElementSibling as HTMLElement | null)?.setAttribute("hidden", "");
    });
  }

  /** Wartosci istniejacej szansy w polach formularza. */
  function wypelnijZSzansy(o: Szansa, zrodla: string[]) {
    const p = o.pola as Record<string, unknown>;
    const dzien = (v: unknown) => {
      if (!v) return "";
      const d = new Date(String(v));
      return isNaN(d.getTime()) ? "" : `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
    };
    const ustaw = (sel: HTMLSelectElement, v: string) => {
      if (!v) return;
      if (!Array.from(sel.options).some((x) => x.value === v)) {
        const opt = document.createElement("option");
        opt.value = opt.textContent = v;
        sel.appendChild(opt);
      }
      sel.value = v;
    };
    nazwa.value = o.nazwa;
    wartosc.value = p["Warto_x015b__x0107_wEUR"] == null ? "" : String(p["Warto_x015b__x0107_wEUR"]);
    otwarcie.value = dzien(p.Dataotwarcia);
    zamkniecie.value = dzien(p["Spodziewanezamkni_x0119_cie"]);
    dr.value = dzien(p.TerminDR);
    ustaw(typ, String(p.TypBiznesu ?? ""));
    ustaw(rodzaj, String(p.Rodzajbiznesu ?? ""));
    ustaw(opiekun, String(p.Creator ?? "").trim());
    champion.value = String(p.DISCH ?? "");
    notatki.value = String(p.Notatki ?? "");
    // Źródło: lookup trzyma Id pozycji, wybor - tekst; bez wartosci nie podstawiamy domyslnej handlowca
    zrodlo.value = "";
    const surowe = p[`${kluczZrodla}LookupId`] ?? p[`OData_${kluczZrodla}LookupId`] ?? p[kluczZrodla] ?? p[`OData_${kluczZrodla}`];
    if (surowe != null) {
      const etykieta = idZrodel.size
        ? [...idZrodel.entries()].find(([, id]) => String(id) === String(surowe))?.[0] ?? ""
        : String(surowe);
      if (etykieta && zrodla.includes(etykieta)) zrodlo.value = etykieta;
    }
  }

  function rysujOsoby() {
    const q = szukaj.value.trim().toLowerCase();
    const wybrany = champion.value.trim().toLowerCase();
    listaOsob.innerHTML = "";
    const pasuja = kontakty.filter((k) => !q || [k.imie, k.nazwisko, k.imieNazwisko].some((x) => x.toLowerCase().startsWith(q)));
    if (!pasuja.length) return pusto(listaOsob, kontakty.length ? "Nikt nie pasuje." : "Brak kontaktów");
    pasuja.forEach((k) => {
      const b = el("button", "crm-ns-osoba") as HTMLButtonElement;
      if (k.imieNazwisko.trim().toLowerCase() === wybrany) b.classList.add("crm-ns-osoba-akt");
      b.appendChild(ikona("osoba"));
      const t = el("span");
      t.append(el("span", "crm-ns-osoba-nazwa", k.imieNazwisko), el("span", "crm-ns-osoba-stan", k.stanowisko || " "));
      b.appendChild(t);
      b.onclick = () => {
        champion.value = k.imieNazwisko;
        rysujOsoby();
        sprawdz();
      };
      listaOsob.appendChild(b);
    });
  }
  szukaj.oninput = rysujOsoby;
  rysujOsoby();
  sprawdz();

  zapisz.onclick = async () => {
    if (!sprawdz()) return;
    zapisz.disabled = wstecz.disabled = true;
    komunikat.textContent = "Zapisywanie...";
    komunikat.className = "crm-komunikat";
    try {
      // pole Osoba wymaga Id uzytkownika witryny - mail z MEScoStaff albo wlasny
      const mail = opiekun.value === ja
        ? await getMojMail().catch(() => Office.context.mailbox.userProfile.emailAddress ?? "")
        : (await getSerwisanci().catch(() => [])).find((x) => x.nazwa === opiekun.value)?.mail ?? "";
      const opiekunId = mail ? await getSiteUserId(mail) : null;
      // w edycji niezmieniony opiekun zostaje jak byl - nie musi byc na liscie pracownikow
      const bezZmianyOpiekuna = !!edycja && opiekun.value === String(edycja.pola.Creator ?? "").trim();
      if (opiekunId === null && !bezZmianyOpiekuna) {
        throw new Error(`nie znaleziono użytkownika witryny dla „${opiekun.value}”`);
      }

      const pola: Record<string, unknown> = {
        Nazwaszansysprzeda_x017c_y: nazwa.value.trim(),
        Etap: etap.value,
        Etap1: etap.value,
        Status: status.value,
        Status1: status.value,
        Rodzajbiznesu: rodzaj.value,
        OpiekunOppsaLookupId: opiekunId,
        DISCH: champion.value.trim(),
        Notatki: notatki.value,
        Firma: skrocona,
        Creator: opiekun.value,
        OPPSGUID: nowyGuid(),
        IdentyfikatorCompanies: c.guid,
        OppsDomen_x0040_: firmaPelna?.guid === c.guid ? firmaPelna.domena : "",
      };
      if (typ.value) pola.TypBiznesu = typ.value;
      if (wartosc.value !== "") pola["Warto_x015b__x0107_wEUR"] = Number(wartosc.value);
      if (otwarcie.value) pola.Dataotwarcia = dzienSP(otwarcie.value);
      if (zamkniecie.value) pola["Spodziewanezamkni_x0119_cie"] = dzienSP(zamkniecie.value);
      if (dr.value) pola.TerminDR = dzienSP(dr.value);
      if (edycja) {
        // tylko pola formularza - status, etap, GUID i powiazanie z firma bez zmian; puste daty czyscimy
        ["Etap", "Etap1", "Status", "Status1", "OPPSGUID", "Firma", "IdentyfikatorCompanies", "OppsDomen_x0040_"]
          .forEach((k) => delete pola[k]);
        if (opiekunId === null) delete pola.OpiekunOppsaLookupId;
        if (wartosc.value === "") pola["Warto_x015b__x0107_wEUR"] = null;
        if (!otwarcie.value) pola.Dataotwarcia = null;
        if (!zamkniecie.value) pola["Spodziewanezamkni_x0119_cie"] = null;
        if (!dr.value) pola.TerminDR = null;
      }
      // nazwa pola Źródło w Graph nie jest pewna: kolumna zaczyna sie od podkreslnika,
      // a Graph raz chce prefiks "OData_", raz nie - probujemy wariantow po kolei
      const sufiks = idZrodel.size ? "LookupId" : "";
      const warianty = [`${kluczZrodla}${sufiks}`, `OData_${kluczZrodla}${sufiks}`];
      const wartoscZrodla = idZrodel.size ? idZrodel.get(zrodlo.value) : zrodlo.value;
      let nowa: Szansa | undefined;
      for (let i = 0; ; i++) {
        warianty.forEach((w) => delete pola[w]);
        pola[warianty[i]] = wartoscZrodla;
        try {
          if (edycja) nowa = await zapiszSzanse(edycja, pola);
          else await createListItem("Opps", pola);
          break;
        } catch (e: any) {
          const nieznane = String(e?.message ?? "").includes(`'${warianty[i]}' is not recognized`);
          if (!nieznane || i === warianty.length - 1) throw e;
        }
      }
      komunikat.textContent = edycja ? "Zmiany zapisane" : `Dodano szansę ${nazwa.value.trim()}`;
      komunikat.className = "crm-komunikat crm-ok";
      poZapisie(nowa);
    } catch (e: any) {
      komunikat.textContent = `Nie udało się zapisać szansy: ${e?.message ?? e}`;
      komunikat.className = "crm-komunikat crm-blad";
      zapisz.disabled = wstecz.disabled = false;
    }
  };
}

// ---------------------------------------------------------------- produkty szansy (conDsOferta / conDsProdukt)

const TYPY_PRODUKTU_SZANSY: { tekst: string; wartosc: ProduktSzansy["typ"] }[] = [
  { tekst: "Paid-up + TECS", wartosc: "PU" },
  { tekst: "Lease", wartosc: "L" },
  { tekst: "AEC", wartosc: "AEC" },
  { tekst: "ALH", wartosc: "ALH" },
];

const opisTypuProduktu = (p: ProduktSzansy) =>
  p.typ === "PU" ? "Paid-up" : p.typ === "L" ? `Lease ${p.dni} dni` : p.typ || "—";

/** Rabat z progow: najwyzszy prog schematu, ktorego "Od" nie przekracza liczby licencji. */
function rabatZProgow(progi: ProgRabatu[], schemat: string, ile: number): number {
  if (!schemat) return 0;
  const pasujace = progi.filter((p) => p.schemat === schemat && p.od <= ile).sort((a, b) => b.od - a.od);
  return pasujace[0]?.rabat ?? 0;
}

/** Wyliczenie pozycji jak w Power Apps - ceny za licencje i suma z rabatem. */
function wyliczProduktSzansy(cena: CenaTecs | undefined, typ: ProduktSzansy["typ"], ile: number, rabatPr: number, dni: number) {
  const pu = cena?.paidUp ?? 0;
  const r = 1 - rabatPr / 100;
  const lease = typ === "L" ? r2(((cena?.lease ?? 0) / 365) * mnoznikDzierzawy(dni) * dni * r) : 0;
  const suma = r2(
    (typ === "PU" ? (pu + (cena?.tecs ?? 0)) * ile * r : 0) +
    (typ === "AEC" || typ === "ALH" ? ile * pu * r : 0) +
    (typ === "L" ? ile * lease : 0)
  );
  return {
    paidUp: typ === "PU" ? pu : 0,
    tecs: typ === "PU" ? cena?.tecs ?? 0 : 0,
    lease,
    aec: typ === "AEC" ? pu : 0,
    alh: typ === "ALH" ? pu : 0,
    suma,
  };
}

/** Lista produktow szansy z suma i "Przelicz wartosc". Zmiany tylko przy aktywnej szansie. */
async function produktySzansy(miejsce: HTMLElement, o: Szansa, c: Company, nowaWartosc: (nowa: Szansa) => void) {
  const box = el("div", "crm-tecs-produkty");
  miejsce.appendChild(box);
  const stan = el("div", "crm-komunikat", "Wczytywanie produktów...");
  box.appendChild(stan);

  let edycja = false;
  let ceny = new Map<string, CenaTecs>();
  let progi: ProgRabatu[] = [];
  let pozycje: ProduktSzansy[] = [];
  try {
    const [p, cn, pr, it] = await Promise.all([
      czyPracownikMESco().catch(() => false),
      getCenyTecs().catch((e) => {
        console.warn("cennik:", e);
        return new Map<string, CenaTecs>();
      }),
      getProgiRabatowe().catch((e) => {
        console.warn("progi rabatowe:", e);
        return [] as ProgRabatu[];
      }),
      getByCompany("Oferty", o.guid, { pole: "IdentyfikatorOpps", top: 200 }),
    ]);
    edycja = p && o.status === "Aktywny";
    ceny = cn;
    progi = pr;
    pozycje = it.map(toProduktSzansy);
  } catch (e: any) {
    stan.textContent = `Błąd: ${e?.message ?? e}`;
    stan.className = "crm-komunikat crm-blad";
    return;
  }
  stan.remove();

  const lista = el("div", "crm-tecs-lista");
  const dodajBox = el("div", "crm-tecs-dodaj");
  const stopka = el("div", "crm-tecs-stopka");
  const komunikat = el("div", "crm-komunikat");
  box.append(lista, dodajBox, stopka, komunikat);

  const ok = (t: string) => {
    komunikat.textContent = t;
    komunikat.className = "crm-komunikat crm-ok";
  };
  const blad = (e: any) => {
    komunikat.textContent = e?.message ?? String(e);
    komunikat.className = "crm-komunikat crm-blad";
  };

  if (edycja) {
    const dodaj = przyciskIk("crm-btn crm-tecs-nowy", "paczka", "Dodaj produkt");
    dodajBox.appendChild(dodaj);
    dodaj.onclick = () => {
      dodaj.disabled = true;
      const forma = formularzProduktuSzansy(o, c, ceny, progi, null, (nowy) => {
        pozycje.push(nowy);
        forma.remove();
        dodaj.disabled = false;
        ok(`Zapisano ${nowy.produkt}. Pamiętaj o „Przelicz wartość”.`);
        rysuj();
      }, () => {
        forma.remove();
        dodaj.disabled = false;
      });
      dodajBox.appendChild(forma);
    };
  } else {
    dodajBox.appendChild(el("div", "crm-nowe-info", o.status === "Aktywny"
      ? "Produkty mogą zmieniać tylko osoby z MEScoStaff."
      : `Szansa „${o.status || "—"}” — produkty można zmieniać tylko przy aktywnej szansie.`));
  }

  const wiersz = (p: ProduktSzansy): HTMLElement => {
    const w = el("div", "crm-tecs-poz");
    const gora = el("div", "crm-tecs-poz-gora");
    gora.appendChild(poleIk("crm-tecs-poz-nazwa", "licencja", p.produkt || "(bez produktu)", p.produkt));
    gora.appendChild(el("span", "crm-tecs-poz-suma", euro(p.suma)));
    w.appendChild(gora);
    w.appendChild(el("div", "crm-tecs-poz-linia",
      `${p.ile} szt. · ${opisTypuProduktu(p)}${p.rabat ? ` · rabat ${p.rabat}%` : ""}`));
    const kwoty = [
      p.paidUp ? `Paid-up ${euro(p.paidUp)}` : "",
      p.tecs ? `TECS ${euro(p.tecs)}` : "",
      p.lease ? `Lease ${euro(p.lease)}` : "",
      p.aec + p.alh ? `AEC/ALH ${euro(p.aec + p.alh)}` : "",
    ].filter(Boolean).join(" · ");
    if (kwoty) w.appendChild(el("div", "crm-tecs-poz-linia crm-tecs-poz-ceny", `${kwoty} (za licencję)`));

    if (!edycja) return w;
    const akcje = el("div", "crm-tecs-poz-akcje");
    w.appendChild(akcje);
    const edytuj = el("button", "crm-usun crm-edytuj") as HTMLButtonElement;
    edytuj.appendChild(ikona("pioro"));
    edytuj.title = "Edytuj produkt";
    akcje.appendChild(edytuj);
    edytuj.onclick = () => {
      const forma = formularzProduktuSzansy(o, c, ceny, progi, p, (nowy) => {
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
      const pyt = el("div", "crm-zal-potwierdz");
      pyt.appendChild(el("span", undefined, `Usunąć ${p.produkt}?`));
      const tak = el("button", "crm-btn crm-btn-ostrz", "Usuń") as HTMLButtonElement;
      const nie = el("button", "crm-btn crm-btn-cichy", "Anuluj") as HTMLButtonElement;
      pyt.append(tak, nie);
      w.appendChild(pyt);
      kosz.disabled = true;
      nie.onclick = () => {
        pyt.remove();
        kosz.disabled = false;
      };
      tak.onclick = async () => {
        tak.disabled = nie.disabled = true;
        try {
          await deleteListItem("Oferty", p.id);
          pozycje = pozycje.filter((x) => x !== p);
          ok(`Usunięto ${p.produkt}. Pamiętaj o „Przelicz wartość”.`);
          rysuj();
        } catch (e) {
          pyt.remove();
          kosz.disabled = false;
          blad(e);
        }
      };
    };
    return w;
  };

  const rysuj = () => {
    lista.innerHTML = "";
    if (!pozycje.length) {
      pusto(lista, edycja ? "Brak produktów — dodaj pierwszy przyciskiem „Dodaj produkt”." : "Brak produktów.");
    }
    pozycje.forEach((p) => lista.appendChild(wiersz(p)));
    rysujStopke();
  };

  const rysujStopke = () => {
    stopka.innerHTML = "";
    const razem = r2(pozycje.reduce((a, p) => a + (p.suma || 0), 0));
    const zgodne = Math.round(razem) === Math.round(o.wartosc || 0);
    const kol = (etykieta: string, wartosc: string, klasa = "") => {
      const k = el("div", "crm-tecs-stopka-kol");
      k.append(el("div", "crm-tecs-stopka-et", etykieta), el("div", `crm-tecs-stopka-war ${klasa}`.trim(), wartosc));
      return k;
    };
    stopka.append(
      kol("Suma pozycji", euro(razem)),
      kol("Wartość oppsa w systemie", euro(o.wartosc || 0), zgodne ? "crm-ok" : "crm-nf-uwaga")
    );
    if (!zgodne) {
      stopka.appendChild(el("div", "crm-nowe-info crm-tecs-legenda",
        "Pomarańczowa wartość = suma pozycji różni się od wartości oppsa. Kliknij „Przelicz wartość”."));
    }
    if (!edycja) return;
    const przelicz = przyciskIk("crm-btn crm-btn-cichy", "kalkulator", "Przelicz wartość");
    przelicz.disabled = zgodne;
    przelicz.title = zgodne ? "Wartość zgodna z sumą pozycji" : "Zapisz sumę pozycji jako wartość oppsa";
    przelicz.onclick = async () => {
      przelicz.disabled = true;
      try {
        const nowa = await zapiszSzanse(o, { "Warto_x015b__x0107_wEUR": razem });
        Object.assign(o, nowa);
        nowaWartosc(nowa);
        ok(`Przeliczono wartość oppsa: ${euro(razem)}`);
        rysujStopke();
      } catch (e) {
        blad(e);
        przelicz.disabled = false;
      }
    };
    stopka.appendChild(przelicz);
  };

  rysuj();
}

/** Nowy produkt / edycja produktu szansy - odpowiednik dsFormOferta z Power Apps. */
function formularzProduktuSzansy(
  o: Szansa,
  c: Company,
  ceny: Map<string, CenaTecs>,
  progi: ProgRabatu[],
  p: ProduktSzansy | null,
  poZapisie: (nowy: ProduktSzansy) => void,
  anuluj: () => void
): HTMLElement {
  const box = el("div", "crm-nowe crm-demo crm-tecs-forma");
  box.appendChild(poleIk("crm-demo-tytul", "paczka", p ? "Edycja produktu" : "Nowy produkt w ofercie"));
  box.appendChild(el("div", "crm-nowe-info", `${o.firma || c.nazwa} — ${o.nazwa}`));

  // --- produkt z cennika (wyszukiwanie po Title) ---
  box.appendChild(el("div", "crm-nf-sekcja", "Produkt z cennika"));
  const produkt = document.createElement("input");
  produkt.type = "text";
  produkt.className = "crm-input";
  produkt.placeholder = "Wyszukaj licencję w cenniku…";
  const lista = document.createElement("datalist");
  lista.id = `cennik-${nowyGuid()}`;
  const tytuly = new Map<string, CenaTecs>();
  ceny.forEach((v) => v.tytul && tytuly.set(v.tytul, v));
  [...tytuly.keys()].sort((a, b) => a.localeCompare(b, "pl")).forEach((k) => {
    const opt = document.createElement("option");
    opt.value = k;
    lista.appendChild(opt);
  });
  produkt.setAttribute("list", lista.id);
  produkt.value = p?.produkt ?? "";
  box.append(produkt, lista);

  // --- rodzaj: domyslnie z TypBiznesu szansy ---
  box.appendChild(el("div", "crm-nf-sekcja", "Rodzaj"));
  let typ: ProduktSzansy["typ"] = p?.typ ||
    (o.typ === "Lease" ? "L" : o.typ === "AEC" ? "AEC" : o.typ === "ALH" ? "ALH" : "PU");
  chipyWyboru(box, TYPY_PRODUKTU_SZANSY, typ, (v) => {
    typ = v;
    dniBox.style.display = v === "L" ? "" : "none";
    przelicz();
  });

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
  // rabat podpowiadany z progow, dopoki ktos go recznie nie zmieni (jak Default w Power Apps)
  let rabatReczny = !!p;
  const dniBox = el("div");
  box.appendChild(dniBox);
  const dni = poleFormularza(dniBox, "Dni dzierżawy");
  dni.inputMode = "numeric";
  dni.value = String(p && p.dni > 0 ? p.dni : 30);
  dniBox.style.display = typ === "L" ? "" : "none";

  const podglad = el("div", "crm-tecs-podglad crm-sz-podglad");
  box.appendChild(podglad);
  const brakujace = el("div", "crm-podpowiedz");
  box.appendChild(brakujace);
  const zapisz = el("button", "crm-btn", "Zapisz") as HTMLButtonElement;
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
  const cenaProduktu = () => ceny.get(produkt.value.trim());

  let wynik = wyliczProduktSzansy(undefined, typ, 0, 0, 0);
  function przelicz() {
    const cena = cenaProduktu();
    if (!rabatReczny) rabat.value = String(rabatZProgow(progi, cena?.schemat ?? "", num(ilosc)));
    wynik = wyliczProduktSzansy(cena, typ, num(ilosc), num(rabat), num(dni));
    podglad.innerHTML = "";
    const linia = (etykieta: string, wartosc: string, klasa = "") => {
      const r = el("div", `crm-tecs-podglad-rzad ${klasa}`.trim());
      r.append(el("span", undefined, etykieta), el("span", undefined, wartosc));
      podglad.appendChild(r);
    };
    if (produkt.value.trim() && !cena) linia("Brak produktu w cenniku", "ceny 0 €", "crm-nf-uwaga");
    linia("Cennik: Paid-Up / TECS / Annual Lease",
      `${euro(cena?.paidUp ?? 0)} / ${euro(cena?.tecs ?? 0)} / ${euro(cena?.lease ?? 0)}`);
    linia("Paid-up (za licencję)", euro(wynik.paidUp));
    linia("TECS (za licencję)", euro(wynik.tecs));
    linia("Lease (za licencję, okres)", euro(wynik.lease));
    linia("AEC / ALH (za licencję)", euro(wynik.aec + wynik.alh));
    linia("Suma pozycji", euro(wynik.suma), "crm-tecs-podglad-suma");
    sprawdz();
  }
  function sprawdz(): boolean {
    const okForm = !!produkt.value.trim() && num(ilosc) > 0;
    zapisz.disabled = !okForm;
    brakujace.textContent = okForm ? "" : "Wybierz produkt i podaj liczbę licencji";
    brakujace.style.display = okForm ? "none" : "block";
    return okForm;
  }
  produkt.oninput = przelicz;
  ilosc.oninput = przelicz;
  dni.oninput = przelicz;
  rabat.oninput = () => {
    rabatReczny = true;
    przelicz();
  };
  przelicz();

  zapisz.onclick = async () => {
    if (!sprawdz()) return;
    zapisz.disabled = wstecz.disabled = true;
    komunikat.textContent = "Zapisywanie...";
    komunikat.className = "crm-komunikat";
    const pola: Record<string, unknown> = {
      PaidUP0: typ === "PU",
      LeaseDays: typ === "L",
      AECYN: typ === "AEC",
      ALHYN: typ === "ALH",
      Liczba_licencji: num(ilosc),
      Paidup: wynik.paidUp,
      TECS: wynik.tecs,
      Lease: wynik.lease,
      Lease_Days: typ === "L" ? num(dni) : 0,
      Suma: wynik.suma,
      ProductID: produkt.value.trim(),
      ALH: wynik.alh,
      AEC: wynik.aec,
      Rabat: num(rabat),
      FirmaID: o.firma || c.nazwa,
      SzansaID: o.nazwa,
      IdentyfikatorCompanies: o.companyGuid || c.guid,
      IdentyfikatorOpps: o.guid,
    };
    try {
      const it = p
        ? { id: p.id, fields: { ...(await updateTicketFields(p.id, pola, p.etag || undefined, "Oferty")), id: p.id } }
        : await createListItem("Oferty", pola);
      poZapisie(toProduktSzansy(it as ListItem));
    } catch (e: any) {
      komunikat.textContent = `Nie udało się zapisać: ${e?.message ?? e}`;
      komunikat.className = "crm-komunikat crm-blad";
      zapisz.disabled = wstecz.disabled = false;
    }
  };
  return box;
}

async function zakladkaSzanse(
  box: HTMLElement,
  s: HTMLElement,
  c: Company,
  odswiezWidok: () => void = () => void pokazFirma(c),
  zPrzyciskiem = false // widok maila nie ma paska firmy - przycisk nad lista
) {
  // "Nowy Opps Soft" - z paska firmy albo przyciskiem nad lista; formularz nad lista
  const miejsce = el("div");
  box.appendChild(miejsce);
  const nowyBtn = przyciskIk("crm-btn crm-tecs-nowy crm-nowa-szansa", "nowaSzansa", "Nowy Opps Soft");
  const otworzForm = () => {
    miejsce.innerHTML = "";
    void formularzNowejSzansy(miejsce, c, () => {
      szanseZamkniete = false;
      odswiezWidok();
    }, () => {
      miejsce.innerHTML = "";
      if (zPrzyciskiem) miejsce.appendChild(nowyBtn);
    });
  };
  nowyBtn.onclick = otworzForm;
  if (otworzNowaSzanse) {
    otworzNowaSzanse = false;
    otworzForm();
  } else if (zPrzyciskiem) miejsce.appendChild(nowyBtn);

  segment(box, ["Aktywne", "Zamknięte"], szanseZamkniete ? "Zamknięte" : "Aktywne", (v) => {
    szanseZamkniete = v === "Zamknięte";
    odswiezWidok();
  });

  // notatki firmy jednym zapytaniem - przy szansach grupujemy je po Opps_ID
  const notatki = getByCompany("Notes", c.guid, { top: 500 }).then((it) => it.map(toNota));
  notatki.catch((e) => console.error("notatki szans:", e));
  const ja = getMojaNazwa().then((n) => n.trim().toLowerCase(), () => "");

  const items = await getByCompany("Opps", c.guid, { top: 100 });
  const szanse = items.map(toSzansa);
  if (autoSzanse) {
    autoSzanse = false;
    if (!szanseZamkniete && !szanse.some((o) => o.status === "Aktywny")) {
      szanseZamkniete = true;
      zaznaczSegment(box, "Zamknięte");
    }
  }
  const lista = szanse
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

  // --- edycja danych - tylko aktywna szansa ---
  if (o.status === "Aktywny") {
    const edytuj = przyciskIk("crm-btn crm-btn-cichy crm-szansa-edytuj", "pioro", "Edytuj dane szansy");
    const edycjaBox = el("div");
    box.append(edytuj, edycjaBox);
    edytuj.onclick = () => {
      edytuj.hidden = true;
      void formularzNowejSzansy(edycjaBox, c, (nowa) => {
        if (nowa) zmieniona(nowa);
      }, () => {
        edycjaBox.innerHTML = "";
        edytuj.hidden = false;
      }, o);
    };
  }

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
        await odswiezTecs(c);
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
  const wartoscEl = poleIk("crm-kafel-wartosc", "euro", euro(o.wartosc), "Wartość");
  rzad.appendChild(wartoscEl);
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
  // zamkniecie: wygrana - data wygranej, aktywna - spodziewane zamkniecie z liczba dni
  const pz = o.pola as Record<string, unknown>;
  const wygrana = String(pz.ClosedDateWin ?? "");
  const spodziewane = String(pz["Spodziewanezamkni_x0119_cie"] ?? "");
  if (o.status === "Zamknięty wygrany" && wygrana) {
    k.appendChild(poleIk("crm-kafel-meta", "flaga", `Zamknięta ${dataPL(wygrana)}`, "Data wygranej"));
  } else if (spodziewane) {
    const d = dniDo(spodziewane);
    const zam = poleIk("crm-kafel-meta", "klepsydra",
      `Zamknięcie ${dataPL(spodziewane)}` +
        (o.status === "Aktywny" && d !== null ? ` · ${d < 0 ? "po terminie" : `za ${d} ${d === 1 ? "dzień" : "dni"}`}` : ""),
      "Spodziewane zamknięcie");
    if (o.status === "Aktywny") zam.style.color = kolorTerminu(d);
    k.appendChild(zam);
  }

  const akcje = el("div", "crm-akcje crm-szansa-akcje crm-karta-akcje");
  k.appendChild(akcje);
  const btnPanel = przyciskIk("crm-btn crm-btn-cichy", "flaga", "Status i etap");
  akcje.appendChild(btnPanel);
  const panelBox = el("div");
  k.appendChild(panelBox);

  // produkty i ceny (lista Oferty) - bez OPPSGUID nie da sie ich powiazac
  if (o.guid) {
    const btnProd = przyciskIk("crm-btn crm-btn-cichy", "paczka", "Produkty");
    akcje.appendChild(btnProd);
    const prodBox = el("div");
    k.appendChild(prodBox);
    btnProd.onclick = () => {
      if (prodBox.childElementCount) {
        prodBox.innerHTML = "";
        btnProd.classList.remove("crm-btn-wcisniety");
        return;
      }
      btnProd.classList.add("crm-btn-wcisniety");
      void produktySzansy(prodBox, o, c, (nowa) => {
        const w = wartoscEl.querySelector(".crm-ikona-txt");
        if (w) w.textContent = euro(nowa.wartosc);
      });
    };
  }

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
  if (o.guid) sekcjaNotatek(k, akcje, o, c, notatki, ja, "Brak notatek do tej szansy.");
  if (o.id) sekcjaZalacznikow(k, akcje, "Opps", o.id, folderFirmy(c));
  przyciskGenerujOferte(akcje);
  return k;
}

// ---------------------------------------------------------------- zalaczniki szansy / TECS-u

const MAX_ZALACZNIK = 50 * 1024 * 1024;

/** Plik z otwartego maila jako bajty - Mailbox 1.8, tylko zalaczniki-pliki. */
function bajtyZalacznikaMaila(id: string): Promise<ArrayBuffer> {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item as Office.MessageRead;
    item.getAttachmentContentAsync(id, (r) => {
      if (r.status !== Office.AsyncResultStatus.Succeeded) return reject(new Error(r.error?.message ?? "nie pobrano"));
      if (r.value.format !== Office.MailboxEnums.AttachmentContentFormat.Base64) {
        return reject(new Error("to nie jest plik"));
      }
      const bin = atob(r.value.content);
      const bajty = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bajty[i] = bin.charCodeAt(i);
      resolve(bajty.buffer);
    });
  });
}

type DoWyslania = { nazwa: string; tresc: () => Promise<ArrayBuffer>; rozmiar: number };

/**
 * Pole "upusc pliki albo kliknij" + zalaczniki otwartego maila jako przyciski.
 * Zalacznikow maila nie da sie przeciagnac - Outlook nie oddaje ich dodatkowi jako plikow.
 */
function polePrzesylania(
  wyslij: (pliki: DoWyslania[]) => Promise<void>,
  komunikat: HTMLElement,
  tytulMaila: string
): [HTMLElement, HTMLElement] {
  const strefa = el("label", "crm-zal-strefa") as HTMLLabelElement;
  const wybor = document.createElement("input");
  wybor.type = "file";
  wybor.multiple = true;
  wybor.hidden = true;
  strefa.append(ikona("wyslij"), el("span", undefined, "Upuść pliki tutaj albo kliknij, aby wybrać"), wybor);

  const zPlikow = (lista: FileList | null) =>
    void wyslij(Array.from(lista ?? []).map((f) => ({ nazwa: f.name, tresc: () => f.arrayBuffer(), rozmiar: f.size })));
  wybor.onchange = () => {
    zPlikow(wybor.files);
    wybor.value = "";
  };
  strefa.ondragover = (e) => {
    e.preventDefault();
    strefa.classList.add("crm-zal-nad");
  };
  strefa.ondragleave = () => strefa.classList.remove("crm-zal-nad");
  strefa.ondrop = (e) => {
    e.preventDefault();
    strefa.classList.remove("crm-zal-nad");
    if (!e.dataTransfer?.files.length) {
      komunikat.textContent = "Tu można upuścić pliki z dysku. Załączniki maila dodasz przyciskami poniżej.";
      komunikat.className = "crm-komunikat";
      return;
    }
    zPlikow(e.dataTransfer.files);
  };

  const zMaila = el("div", "crm-zal-mail");
  const item = Office.context.mailbox.item as Office.MessageRead | undefined;
  const plikiMaila = (item?.attachments ?? []).filter(
    (a) => a.attachmentType === Office.MailboxEnums.AttachmentType.File && !a.isInline
  );
  if (plikiMaila.length && Office.context.requirements.isSetSupported("Mailbox", "1.8")) {
    zMaila.appendChild(el("div", "crm-label", "Z otwartego maila"));
    plikiMaila.forEach((a) => {
      const b = przyciskIk("crm-chip-wybor crm-zal-chip", "spinacz", a.name);
      b.title = tytulMaila;
      b.onclick = async () => {
        b.disabled = true;
        try {
          await wyslij([{ nazwa: a.name, tresc: () => bajtyZalacznikaMaila(a.id), rozmiar: a.size }]);
        } finally {
          b.disabled = false;
        }
      };
      zMaila.appendChild(b);
    });
  }
  return [strefa, zMaila];
}

/** Rozmiar pliku po ludzku. */
const rozmiarPL = (b: number) =>
  b < 1024 ? `${b} B` : b < 1048576 ? `${Math.round(b / 1024)} KB` : `${(b / 1048576).toLocaleString("pl-PL", { maximumFractionDigits: 1 })} MB`;

// podfolder otwarty w zakladce Dokumenty - wzgledem folderu firmy
let dokSciezka: { firma: string; podfolder: string[] } = { firma: "", podfolder: [] };
const KATALOG_KLIENTOW = "DokumentyKlienci";
const MAX_PLIK_DOK = 250 * 1024 * 1024;

/** Folder firmy w SharePoint (Shared Documents/DokumentyKlienci/<firma>) w dodatku. */
async function zakladkaDokumenty(box: HTMLElement, s: HTMLElement, c: Company) {
  const firma = (firmaPelna?.guid === c.guid ? firmaPelna.skrocona : "") || c.nazwa;
  if (dokSciezka.firma !== c.guid) dokSciezka = { firma: c.guid, podfolder: [] };
  const glowny = `${KATALOG_KLIENTOW}/${firma}`;
  const sciezka = [glowny, ...dokSciezka.podfolder].join("/");
  const odswiez = () => void pokazFirma(c);

  // --- okruszki: firma / podfolder / ... + Otworz w SharePoint ---
  const gora = el("div", "crm-dok-gora");
  const okruszki = el("div", "crm-dok-okruszki");
  [firma, ...dokSciezka.podfolder].forEach((n, i, wsz) => {
    if (i) okruszki.appendChild(el("span", "crm-dok-sep", "/"));
    const b = el("button", i === wsz.length - 1 ? "crm-dok-okr crm-dok-okr-akt" : "crm-dok-okr", n) as HTMLButtonElement;
    b.onclick = () => {
      dokSciezka.podfolder = dokSciezka.podfolder.slice(0, i);
      odswiez();
    };
    okruszki.appendChild(b);
  });
  const wSP = przyciskIk("crm-btn crm-btn-cichy", "folder", "Otwórz w SharePoint");
  gora.append(okruszki, wSP);
  box.appendChild(gora);
  wSP.onclick = async () => {
    const url = await getUrlFolderu(sciezka).catch(() => null);
    window.open(url ?? `https://mescocompl.sharepoint.com/sites/CRMSystemPOC/Shared%20Documents/${KATALOG_KLIENTOW}/` +
      encodeURIComponent(firma), "_blank");
  };

  const komunikat = el("div", "crm-komunikat");
  const pliki = await getPlikiFolderu(sciezka);

  // --- brak folderu firmy ---
  if (pliki === null) {
    s.textContent = "Dokumenty";
    if (dokSciezka.podfolder.length) {
      dokSciezka.podfolder = [];
      return odswiez();
    }
    pusto(box, `Firma nie ma jeszcze folderu „${firma}” w ${KATALOG_KLIENTOW}.`);
    const utworz = przyciskIk("crm-btn", "folder", "Utwórz folder");
    box.append(utworz, komunikat);
    utworz.onclick = async () => {
      utworz.disabled = true;
      try {
        await utworzFolder(KATALOG_KLIENTOW, firma);
        odswiez();
      } catch (e: any) {
        komunikat.textContent = e?.message ?? String(e);
        komunikat.className = "crm-komunikat crm-blad";
        utworz.disabled = false;
      }
    };
    return;
  }

  // foldery najpierw, potem pliki od najnowszych
  pliki.sort((a, b) => Number(b.folder) - Number(a.folder) ||
    (a.folder ? a.nazwa.localeCompare(b.nazwa, "pl") : b.zmieniony.localeCompare(a.zmieniony)));
  s.textContent = `Dokumenty — ${pliki.length}`;

  async function wyslij(doWyslania: DoWyslania[]) {
    let ok = 0;
    const bledy: string[] = [];
    for (const p of doWyslania) {
      komunikat.className = "crm-komunikat";
      komunikat.textContent = `Wysyłanie: ${p.nazwa}...`;
      try {
        if (p.rozmiar > MAX_PLIK_DOK) throw new Error("plik większy niż 250 MB");
        if (pliki!.some((z) => z.nazwa.toLowerCase() === p.nazwa.toLowerCase())) throw new Error("plik o tej nazwie już jest");
        await wyslijPlik(sciezka, p.nazwa, await p.tresc());
        ok++;
      } catch (e: any) {
        bledy.push(`${p.nazwa}: ${e?.message ?? e}`);
      }
    }
    if (!bledy.length) return odswiez();
    komunikat.textContent = [ok ? `Zapisano: ${ok}` : "", ...bledy].filter(Boolean).join(" · ");
    komunikat.className = "crm-komunikat crm-blad";
    if (ok) setTimeout(odswiez, 2500);
  }
  const [strefa, zMaila] = polePrzesylania(wyslij, komunikat, "Zapisz ten załącznik w folderze firmy");
  box.append(strefa, zMaila, komunikat);

  // --- nowy podfolder ---
  const nowyRzad = el("div", "crm-dok-nowy");
  const nowyBtn = przyciskIk("crm-btn crm-btn-cichy", "folder", "Nowy folder");
  nowyRzad.appendChild(nowyBtn);
  box.appendChild(nowyRzad);
  nowyBtn.onclick = () => {
    nowyBtn.style.display = "none";
    const pole = document.createElement("input");
    pole.className = "crm-input";
    pole.placeholder = "nazwa folderu";
    const ok = el("button", "crm-btn", "Utwórz") as HTMLButtonElement;
    const nie = el("button", "crm-btn crm-btn-cichy", "Anuluj") as HTMLButtonElement;
    nowyRzad.append(pole, ok, nie);
    pole.focus();
    nie.onclick = odswiez;
    ok.onclick = async () => {
      if (!pole.value.trim()) return;
      ok.disabled = true;
      try {
        await utworzFolder(sciezka, pole.value.trim());
        odswiez();
      } catch (e: any) {
        komunikat.textContent = e?.message ?? String(e);
        komunikat.className = "crm-komunikat crm-blad";
        ok.disabled = false;
      }
    };
  };

  // --- lista ---
  const lista = el("div", "crm-dok-lista");
  box.appendChild(lista);
  if (dokSciezka.podfolder.length) {
    const wyzej = el("button", "crm-dok-wiersz crm-dok-wyzej") as HTMLButtonElement;
    wyzej.append(ikona("folder"), el("span", "crm-dok-nazwa", ".. (wyżej)"));
    wyzej.onclick = () => {
      dokSciezka.podfolder.pop();
      odswiez();
    };
    lista.appendChild(wyzej);
  }
  if (!pliki.length) pusto(lista, "Folder jest pusty.");

  pliki.forEach((p) => {
    const wiersz = el("div", "crm-dok-pozycja");
    const otworz = (p.folder ? el("button", "crm-dok-wiersz") : document.createElement("a")) as HTMLElement;
    if (!p.folder) {
      otworz.className = "crm-dok-wiersz";
      const a = otworz as HTMLAnchorElement;
      a.href = p.url;
      a.target = "_blank";
      a.rel = "noopener";
      a.title = "Otwórz";
    } else {
      otworz.onclick = () => {
        dokSciezka.podfolder.push(p.nazwa);
        odswiez();
      };
    }
    const opis = el("span", "crm-dok-opis");
    opis.append(
      el("span", "crm-dok-nazwa", p.nazwa),
      el("span", "crm-dok-meta", [
        p.zmieniony ? new Date(p.zmieniony).toLocaleDateString("pl-PL") : "",
        p.folder ? "" : rozmiarPL(p.rozmiar),
        p.kto,
      ].filter(Boolean).join(" · "))
    );
    otworz.append(ikona(p.folder ? "folder" : "dokument"), opis);
    if (p.folder) otworz.classList.add("crm-dok-folder");

    const kosz = el("button", "crm-zal-kosz") as HTMLButtonElement;
    kosz.title = p.folder ? "Usuń folder" : "Usuń plik";
    kosz.appendChild(ikona("kosz"));
    wiersz.append(otworz, kosz);
    lista.appendChild(wiersz);

    kosz.onclick = () => {
      kosz.disabled = true;
      const pyt = el("div", "crm-zal-potwierdz");
      pyt.appendChild(el("span", undefined, p.folder
        ? `Usunąć folder „${p.nazwa}” razem z zawartością?`
        : `Usunąć „${p.nazwa}”?`));
      const tak = el("button", "crm-btn crm-btn-ostrz", "Usuń") as HTMLButtonElement;
      const nie = el("button", "crm-btn crm-btn-cichy", "Anuluj") as HTMLButtonElement;
      pyt.append(tak, nie);
      wiersz.after(pyt);
      nie.onclick = () => {
        pyt.remove();
        kosz.disabled = false;
      };
      tak.onclick = async () => {
        tak.disabled = nie.disabled = true;
        try {
          await usunPlik(p.id);
          odswiez();
        } catch (e: any) {
          pyt.remove();
          kosz.disabled = false;
          komunikat.textContent = e?.message ?? String(e);
          komunikat.className = "crm-komunikat crm-blad";
        }
      };
    };
  });
}

/** Nazwa folderu firmy w DokumentyKlienci - skrocona nazwa, jak w zakladce Dokumenty. */
const folderFirmy = (c: Company) => (firmaPelna?.guid === c.guid ? firmaPelna.skrocona : "") || c.nazwa;

/** Wolna nazwa w folderze: "oferta.pdf" -> "oferta (1).pdf", gdy pierwsza jest zajeta. */
function wolnaNazwa(nazwa: string, zajete: string[]): string {
  const male = new Set(zajete.map((z) => z.toLowerCase()));
  if (!male.has(nazwa.toLowerCase())) return nazwa;
  const kropka = nazwa.lastIndexOf(".");
  const [rdzen, rozsz] = kropka > 0 ? [nazwa.slice(0, kropka), nazwa.slice(kropka)] : [nazwa, ""];
  for (let i = 1; ; i++) {
    const kandydat = `${rdzen} (${i})${rozsz}`;
    if (!male.has(kandydat.toLowerCase())) return kandydat;
  }
}

/**
 * Wybor folderu w DokumentyKlienci/<firma> - podfoldery, nowy folder, "Kopiuj tutaj" / "Przenies tutaj".
 * cel() dostaje sciezke folderu i nazwy plikow, ktore juz w nim sa.
 */
function wyborFolderu(
  miejsce: HTMLElement,
  firma: string,
  cel: (sciezka: string, zajete: string[], przenies: boolean) => Promise<void>,
  zamknij: () => void
) {
  const box = el("div", "crm-nowe crm-wybor-folderu");
  miejsce.appendChild(box);
  const glowny = `${KATALOG_KLIENTOW}/${firma}`;
  let podfolder: string[] = [];

  const rysuj = async () => {
    box.innerHTML = "";
    box.appendChild(poleIk("crm-demo-tytul", "folder", "Kopiuj do Dokumentów"));
    const okruszki = el("div", "crm-dok-okruszki");
    [firma, ...podfolder].forEach((n, i, wsz) => {
      if (i) okruszki.appendChild(el("span", "crm-dok-sep", "/"));
      const b = el("button", i === wsz.length - 1 ? "crm-dok-okr crm-dok-okr-akt" : "crm-dok-okr", n) as HTMLButtonElement;
      b.onclick = () => {
        podfolder = podfolder.slice(0, i);
        void rysuj();
      };
      okruszki.appendChild(b);
    });
    box.appendChild(okruszki);
    const komunikat = el("div", "crm-komunikat", "Wczytywanie folderu...");
    box.appendChild(komunikat);
    const anuluj = el("button", "crm-btn crm-btn-cichy", "Anuluj") as HTMLButtonElement;
    anuluj.onclick = zamknij;

    const sciezka = [glowny, ...podfolder].join("/");
    let pliki: Awaited<ReturnType<typeof getPlikiFolderu>>;
    try {
      pliki = await getPlikiFolderu(sciezka);
    } catch (e: any) {
      komunikat.textContent = e?.message ?? String(e);
      komunikat.className = "crm-komunikat crm-blad";
      box.appendChild(anuluj);
      return;
    }
    komunikat.textContent = "";

    // firma bez folderu - propozycja utworzenia
    if (pliki === null) {
      komunikat.textContent = `Firma nie ma jeszcze folderu „${firma}” w ${KATALOG_KLIENTOW}.`;
      const utworz = przyciskIk("crm-btn", "folder", "Utwórz folder");
      const rzad = el("div", "crm-akcje");
      rzad.append(utworz, anuluj);
      box.appendChild(rzad);
      utworz.onclick = async () => {
        utworz.disabled = true;
        try {
          await utworzFolder(KATALOG_KLIENTOW, firma);
          void rysuj();
        } catch (e: any) {
          komunikat.textContent = e?.message ?? String(e);
          komunikat.className = "crm-komunikat crm-blad";
          utworz.disabled = false;
        }
      };
      return;
    }

    const lista = el("div", "crm-dok-lista");
    box.appendChild(lista);
    if (podfolder.length) {
      const wyzej = el("button", "crm-dok-wiersz crm-dok-wyzej") as HTMLButtonElement;
      wyzej.append(ikona("folder"), el("span", "crm-dok-nazwa", ".. (wyżej)"));
      wyzej.onclick = () => {
        podfolder.pop();
        void rysuj();
      };
      lista.appendChild(wyzej);
    }
    const foldery = pliki.filter((p) => p.folder).sort((a, b) => a.nazwa.localeCompare(b.nazwa, "pl"));
    if (!foldery.length && !podfolder.length) lista.appendChild(el("div", "crm-pusto", "Brak podfolderów."));
    foldery.forEach((f) => {
      const b = el("button", "crm-dok-wiersz crm-dok-folder") as HTMLButtonElement;
      b.append(ikona("folder"), el("span", "crm-dok-nazwa", f.nazwa));
      b.onclick = () => {
        podfolder.push(f.nazwa);
        void rysuj();
      };
      lista.appendChild(b);
    });

    // akcje w dwoch rzedach: kopiuj / przenies, pod spodem nowy folder / anuluj
    const kopiujBtn = przyciskIk("crm-btn", "kopiuj", "Kopiuj tutaj");
    const przeniesBtn = przyciskIk("crm-btn", "wyslij", "Przenieś tutaj");
    przeniesBtn.title = "Skopiuj do tego folderu i usuń z załączników";
    const rzad = el("div", "crm-wybor-akcje");
    rzad.append(kopiujBtn, przeniesBtn);
    box.appendChild(rzad);

    const nowyRzad = el("div", "crm-wybor-akcje");
    const nowyBtn = przyciskIk("crm-btn crm-btn-cichy", "folder", "Nowy folder");
    nowyRzad.append(nowyBtn, anuluj);
    box.appendChild(nowyRzad);
    const nowyPole = el("div", "crm-dok-nowy");
    box.appendChild(nowyPole);
    nowyBtn.onclick = () => {
      nowyBtn.disabled = true;
      const pole = document.createElement("input");
      pole.className = "crm-input";
      pole.placeholder = "nazwa folderu";
      const ok = el("button", "crm-btn", "Utwórz") as HTMLButtonElement;
      nowyPole.append(pole, ok);
      pole.focus();
      ok.onclick = async () => {
        if (!pole.value.trim()) return;
        ok.disabled = true;
        try {
          await utworzFolder(sciezka, pole.value.trim());
          podfolder.push(pole.value.trim());
          void rysuj();
        } catch (e: any) {
          komunikat.textContent = e?.message ?? String(e);
          komunikat.className = "crm-komunikat crm-blad";
          ok.disabled = false;
        }
      };
    };

    const zajete = pliki.filter((p) => !p.folder).map((p) => p.nazwa);
    const wykonaj = async (przenies: boolean) => {
      kopiujBtn.disabled = przeniesBtn.disabled = anuluj.disabled = nowyBtn.disabled = true;
      komunikat.className = "crm-komunikat";
      komunikat.textContent = przenies ? "Przenoszenie..." : "Kopiowanie...";
      try {
        await cel(sciezka, zajete, przenies);
      } catch (e: any) {
        komunikat.textContent = e?.message ?? String(e);
        komunikat.className = "crm-komunikat crm-blad";
        kopiujBtn.disabled = przeniesBtn.disabled = anuluj.disabled = nowyBtn.disabled = false;
      }
    };
    kopiujBtn.onclick = () => void wykonaj(false);
    przeniesBtn.onclick = () => void wykonaj(true);
  };
  void rysuj();
}

/**
 * Przycisk "Załączniki" przy karcie - lista plikow z SharePointa (oferty z generatora),
 * dodawanie przeciaganiem, z dysku albo z zalacznikow otwartego maila.
 */
function sekcjaZalacznikow(k: HTMLElement, akcje: HTMLElement, listName: string, itemId: string, firma: string) {
  const btn = przyciskIk("crm-btn crm-btn-cichy", "spinacz", "Załączniki");
  akcje.appendChild(btn);
  const panel = el("div", "crm-zal");
  panel.style.display = "none";
  k.appendChild(panel);
  const etykieta = btn.querySelector("span")!;

  const lista = el("div", "crm-zal-lista");
  const komunikat = el("div", "crm-komunikat");

  const [strefa, zMaila] = polePrzesylania((p) => wyslij(p), komunikat, "Zapisz ten załącznik w SharePoint");

  panel.append(lista, strefa, zMaila, komunikat);

  let pliki: { nazwa: string; url: string }[] = [];
  const rysuj = () => {
    etykieta.textContent = `Załączniki (${pliki.length})`;
    lista.innerHTML = "";
    if (!pliki.length) lista.appendChild(el("div", "crm-pusto", "Brak plików w SharePoint."));
    pliki.forEach((z) => {
      const a = document.createElement("a");
      a.className = "crm-zal-plik";
      // dokumenty Office otwieraja sie w przegladarce, reszta pobiera
      a.href = /\.(docx?|xlsx?|pptx?)$/i.test(z.nazwa) ? `${z.url}?web=1` : z.url;
      a.target = "_blank";
      a.rel = "noopener";
      a.title = "Otwórz";
      a.append(ikona("dokument"), el("span", undefined, z.nazwa));

      // kosz z potwierdzeniem - oferta z generatora nie ma kopii w koszu dodatku
      const wiersz = el("div", "crm-zal-wiersz");
      const doDok = el("button", "crm-zal-kosz crm-zal-dok") as HTMLButtonElement;
      doDok.title = "Kopiuj do Dokumentów firmy";
      doDok.setAttribute("aria-label", `Kopiuj ${z.nazwa} do Dokumentów`);
      doDok.appendChild(ikona("folder"));
      doDok.onclick = () => {
        if (wiersz.nextElementSibling?.classList.contains("crm-wybor-folderu")) return;
        doDok.disabled = true;
        const miejsce = el("div");
        wiersz.after(miejsce);
        const zamknijWybor = () => {
          miejsce.remove();
          doDok.disabled = false;
        };
        wyborFolderu(miejsce, firma, async (sciezka, zajete, przenies) => {
          const nazwa = wolnaNazwa(z.nazwa, zajete);
          await wyslijPlik(sciezka, nazwa, await pobierzZalacznik(listName, itemId, z.nazwa));
          const gdzie = sciezka.replace(`${KATALOG_KLIENTOW}/`, "");
          if (przenies) {
            await usunZalacznik(listName, itemId, z.nazwa);
            pliki = pliki.filter((x) => x !== z);
            rysuj();
          } else zamknijWybor();
          komunikat.textContent = `${przenies ? "Przeniesiono" : "Skopiowano"} do ${gdzie}/${nazwa}`;
          komunikat.className = "crm-komunikat crm-ok";
        }, zamknijWybor);
      };
      const kosz = el("button", "crm-zal-kosz") as HTMLButtonElement;
      kosz.title = "Usuń plik z SharePoint";
      kosz.setAttribute("aria-label", `Usuń ${z.nazwa}`);
      kosz.appendChild(ikona("kosz"));
      wiersz.append(a, doDok, kosz);
      lista.appendChild(wiersz);

      kosz.onclick = () => {
        kosz.disabled = true;
        const pyt = el("div", "crm-zal-potwierdz");
        pyt.appendChild(el("span", undefined, `Usunąć „${z.nazwa}”?`));
        const tak = el("button", "crm-btn crm-btn-ostrz", "Usuń") as HTMLButtonElement;
        const nie = el("button", "crm-btn crm-btn-cichy", "Anuluj") as HTMLButtonElement;
        pyt.append(tak, nie);
        wiersz.after(pyt);
        nie.onclick = () => {
          pyt.remove();
          kosz.disabled = false;
        };
        tak.onclick = async () => {
          tak.disabled = nie.disabled = true;
          komunikat.className = "crm-komunikat";
          komunikat.textContent = `Usuwanie: ${z.nazwa}...`;
          try {
            await usunZalacznik(listName, itemId, z.nazwa);
            pliki = pliki.filter((x) => x !== z);
            rysuj();
            komunikat.textContent = `Usunięto: ${z.nazwa}`;
            komunikat.className = "crm-komunikat crm-ok";
          } catch (e: any) {
            pyt.remove();
            kosz.disabled = false;
            komunikat.textContent = e?.message ?? String(e);
            komunikat.className = "crm-komunikat crm-blad";
          }
        };
      };
    });
  };

  let wczytane = false;
  const wczytaj = async () => {
    lista.innerHTML = "";
    lista.appendChild(el("div", "crm-pusto", "Wczytywanie plików..."));
    try {
      pliki = (await getZalaczniki(listName, itemId)).sort((a, b) => a.nazwa.localeCompare(b.nazwa, "pl"));
      wczytane = true;
      rysuj();
    } catch (e: any) {
      lista.innerHTML = "";
      etykieta.textContent = "Załączniki (błąd)";
      komunikat.textContent = e?.message ?? String(e);
      komunikat.className = "crm-komunikat crm-blad";
    }
  };

  async function wyslij(doWyslania: DoWyslania[]) {
    let ok = 0;
    const bledy: string[] = [];
    for (const p of doWyslania) {
      komunikat.className = "crm-komunikat";
      komunikat.textContent = `Wysyłanie: ${p.nazwa}...`;
      try {
        if (p.rozmiar > MAX_ZALACZNIK) throw new Error("plik większy niż 50 MB");
        if (pliki.some((z) => z.nazwa.toLowerCase() === p.nazwa.toLowerCase())) {
          throw new Error("plik o tej nazwie już jest");
        }
        await dodajZalacznik(listName, itemId, p.nazwa, await p.tresc());
        ok++;
      } catch (e: any) {
        bledy.push(`${p.nazwa}: ${e?.message ?? e}`);
      }
    }
    await wczytaj();
    komunikat.textContent = [ok ? `Zapisano w SharePoint: ${ok}` : "", ...bledy].filter(Boolean).join(" · ");
    komunikat.className = bledy.length ? "crm-komunikat crm-blad" : "crm-komunikat crm-ok";
  }
  btn.onclick = () => {
    const widac = panel.style.display !== "none";
    panel.style.display = widac ? "none" : "";
    btn.classList.toggle("crm-btn-wcisniety", !widac);
    if (!widac && !wczytane) void wczytaj();
  };
}

/** "Generuj ofertę" na karcie szansy i TECS-u - na razie bez akcji, generator dojdzie pozniej. */
function przyciskGenerujOferte(akcje: HTMLElement) {
  const b = przyciskIk("crm-btn crm-btn-cichy", "dokument", "Generuj ofertę");
  b.title = "Generowanie oferty — wkrótce";
  b.onclick = () => blysk(b, "wkrótce");
  akcje.appendChild(b);
}

/** Cel notatki: szansa sprzedazy (OPPSGUID) albo TECS (TECSGUID) - oba ida do Notes.Opps_ID. */
type CelNotatki = Pick<Szansa, "id" | "guid" | "nazwa" | "firma" | "companyGuid" | "etap">;

/** Przyciski "Notatki (n)" i "Dodaj notatkę" z lista pod karta - wspolne dla szans i TECS-ow. */
function sekcjaNotatek(
  k: HTMLElement,
  akcje: HTMLElement,
  o: CelNotatki,
  c: Company,
  notatki: Promise<Nota[]>,
  ja: Promise<string>,
  brakTekst: string
) {
  const btnNoty = przyciskIk("crm-btn crm-btn-cichy", "dymek", "Notatki …");
  const btnNowa = przyciskIk("crm-btn crm-btn-cichy", "pioro", "Dodaj notatkę");
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
    if (!moje.length) listaNot.appendChild(el("div", "crm-pusto", brakTekst));
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
  // drugi klik zamyka formularz jak przy "Notatki"; wpisana tresc wymaga potwierdzenia
  const zamknij = () => {
    formBox.innerHTML = "";
    btnNowa.classList.remove("crm-btn-wcisniety");
  };
  const probaZamkniecia = () => {
    const tresc = formBox.querySelector("textarea")?.value.trim() ?? "";
    if (!tresc) return zamknij();
    if (formBox.querySelector(".crm-zamknij-pyt")) return;
    const pyt = el("div", "crm-zal-potwierdz crm-zamknij-pyt");
    pyt.appendChild(el("span", undefined, "Notatka ma wpisaną treść. Na pewno zamknąć bez zapisu?"));
    const tak = el("button", "crm-btn crm-btn-ostrz", "Zamknij bez zapisu") as HTMLButtonElement;
    const nie = el("button", "crm-btn crm-btn-cichy", "Wróć do notatki") as HTMLButtonElement;
    pyt.append(tak, nie);
    formBox.prepend(pyt);
    pyt.scrollIntoView({ block: "nearest", behavior: "smooth" });
    tak.onclick = zamknij;
    nie.onclick = () => pyt.remove();
  };
  btnNowa.onclick = () => {
    if (formBox.childElementCount) return probaZamkniecia();
    btnNowa.classList.add("crm-btn-wcisniety");
    void formularzNotatkiSzansy(formBox, o, c, (nt) => {
      zamknij();
      moje = [nt, ...moje];
      void notatki.then((a) => a.push(nt));
      rysuj();
      listaNot.style.display = "";
      btnNoty.classList.add("crm-btn-wcisniety");
    }, probaZamkniecia);
  };
}

/** Nowa notatka do szansy albo TECS-u (Notes.Opps_ID = OPPSGUID / TECSGUID). Odpowiednik conNotatkaOpp z Power Apps. */
async function formularzNotatkiSzansy(
  miejsce: HTMLElement,
  o: CelNotatki,
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
      lista.appendChild(kartaKontaktu(k, c, () => {
        wszyscy.splice(wszyscy.indexOf(k), 1);
        rysuj();
      }, true));
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
