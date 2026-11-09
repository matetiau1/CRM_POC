// src/taskpane/crm.ts
import { ListItem } from "./graphService";

export type TicketStatus =
  "Przydzielony" | "W toku" | "Oczekuje na odpowiedź" | "Rozwiązany" | "Zamknięty" | "";

export type Ticket = {
  id: string;            // id elementu listy (do edycji)
  ticketId: number;      // MESco-ID
  firma: string;
  osoba: string;
  osobaMail: string;
  opiekun: string;       // nazwa wyswietlana
  opiekunMail: string;   // stabilniejszy klucz niz nazwa
  opis: string;
  data: string;          // ISO
  czas: number;          // godziny
  zamkniete: boolean;    // Status = true -> klodka zamknieta
  status: TicketStatus;  // StatusTicket
  modified: string;      // ISO - do wykrywania "utknietych" zgloszen
  companyId: string;
  contactId: string;
  etag: string;          // @odata.etag - optymistyczna blokada przy zapisie
};

const str = (v: unknown) => (v == null ? "" : String(v));
const num = (v: unknown) => (v == null || v === "" ? 0 : Number(v));

export function toTicket(item: ListItem): Ticket {
  const f = item.fields as Record<string, unknown>;
  return {
    id: str(f.id),
    ticketId: num(f.TicketID),
    firma: str(f.Firma),
    osoba: str(f.Osoba),
    osobaMail: str(f.OsobaMail).toLowerCase(),
    opiekun: str(f.Dodatkowy_opiekun),
    opiekunMail: str(f.MailCreator).toLowerCase(),
    opis: str(f.Opis),
    data: str(f.Data),
    czas: num(f.Czas),
    zamkniete: f.Status === true,
    status: str(f.StatusTicket) as TicketStatus,
    modified: str(f.Modified),
    companyId: str(f.IdentyfikatorCompanies),
    contactId: str(f.IdentyfikatorContact),
    etag: str(f["@odata.etag"] ?? (item as any)["@odata.etag"]),
  };
}

/** 05.sierpień.2026 - jak w Power Apps */
export function formatData(iso: string): string {
  if (!iso) return "";
  const d = new Date(iso);
  const m = [
    "styczeń", "luty", "marzec", "kwiecień", "maj", "czerwiec",
    "lipiec", "sierpień", "wrzesień", "październik", "listopad", "grudzień",
  ][d.getMonth()];
  return `${String(d.getDate()).padStart(2, "0")}.${m}.${d.getFullYear()}`;
}

/** Ile dni od ostatniej zmiany - sygnal, ze status moze byc nieaktualny. */
export function dniOdZmiany(iso: string): number {
  if (!iso) return 0;
  return Math.floor((Date.now() - new Date(iso).getTime()) / 86400000);
}

export const KOLOR_STATUSU: Record<string, string> = {
  "Przydzielony": "#DC2626",
  "W toku": "#F59E0B",
  "Oczekuje na odpowiedź": "#10B981",
  "Rozwiązany": "#2563EB",
  "Zamknięty": "#6B7280",
};

/** Numer zgloszenia z tematu wiadomosci: [MESco-ID-14487] */
export function ticketIdZTematu(subject: string): number | null {
  const m = /\[MESco-ID-(\d+)\]/i.exec(subject || "");
  return m ? Number(m[1]) : null;
}
// ===== Dopisz na koncu src/taskpane/crm.ts =====

export type Company = {
  id: string;
  guid: string;        // CompaniesGUID -> Support.IdentyfikatorCompanies
  nazwa: string;
  ulica: string;
  kod: string;
  miasto: string;
  tecs: number;
  opiekun: string;
};

export function toCompany(item: ListItem): Company {
  const f = item.fields as Record<string, unknown>;
  const s = (v: unknown) => (v == null ? "" : String(v));
  const n = (v: unknown) => (v == null || v === "" ? 0 : Number(v));
  return {
    id: s(f.id),
    guid: s(f.CompaniesGUID),
    nazwa: s(f.Nazwa) || s(f.Pe_x0142_nanazwa),
    ulica: s(f.Ulica),
    kod: s(f.Kodpocztowy),
    miasto: s(f.Miasto),
    tecs: n(f.TECSh),
    opiekun: s(f.Creator),
  };
}
// ===== Dopisz na koncu src/taskpane/crm.ts =====

export type Contact = {
  id: string;
  guid: string;          // ContactGUID
  imieNazwisko: string;
  email: string;
  telefon: string;
  stanowisko: string;
  firmaId: string;       // FirmaID - etykieta firmy
  companyGuid: string;   // IdentyfikatorCompanies
  aktywny: boolean;      // Status
};

export function toContact(item: ListItem): Contact {
  const f = item.fields as Record<string, unknown>;
  const s = (v: unknown) => (v == null ? "" : String(v));

  const pelne = s(f["Imi_x0119__x0020_i_x0020_nazwisk"])
    || [s(f["Imi_x0119_"]), s(f["Nazwisko"])].filter(Boolean).join(" ");

  return {
    id: s(f.id),
    guid: s(f.ContactGUID),
    imieNazwisko: pelne,
    email: s(f.Email),
    telefon: s(f.Phone),
    stanowisko: s(f.Stanowisko),
    firmaId: s(f.FirmaID),
    companyGuid: s(f.IdentyfikatorCompanies),
    aktywny: f.Status !== true,
  };
}// ===== Dopisz na koncu src/taskpane/crm.ts =====

const s = (v: unknown) => (v == null ? "" : String(v));
const n = (v: unknown) => (v == null || v === "" ? 0 : Number(v));

/** Pelne dane firmy - wiecej niz lista firm, uzywane w naglowku. */
export type CompanyFull = {
  id: string;
  guid: string;
  skrocona: string;
  pelna: string;
  nip: string;
  ulica: string;
  miasto: string;
  kod: string;
  status: string;
  dataNDA: string;
  opiekun: string;
  tecsH: number;
  supportDay: number;
  preSalesH: number;
  uwagi: string;
  domena: string;      // DomenA - sama domena, do dopasowania nadawcy
};

export function toCompanyFull(item: ListItem): CompanyFull {
  const f = item.fields as Record<string, unknown>;
  return {
    id: s(f.id),
    guid: s(f.CompaniesGUID),
    skrocona: s(f.Nazwa),
    pelna: s(f["Pe_x0142_nanazwa"]),
    nip: s(f.VAT),                       // uwaga: wewnetrzna nazwa to VAT
    ulica: s(f.Ulica),
    miasto: s(f.Miasto),
    kod: s(f.Kodpocztowy),
    status: s(f.Status1),
    dataNDA: s(f.DataNDA),
    opiekun: s(f.Creator),
    tecsH: n(f.TECSh),
    supportDay: n(f.SupportDay),
    preSalesH: n(f.PreSalesh),
    uwagi: s(f.Uwagi),
    domena: s(f.DomenA).trim().toLowerCase(),
  };
}

/** Imie i nazwisko z nazwy wyswietlanej nadawcy ("Jan Kowalski" albo "Kowalski, Jan"). */
export function rozbijNazwe(nazwa: string): { imie: string; nazwisko: string } {
  let n = (nazwa || "").replace(/["']/g, "").trim();
  if (n.includes("@")) n = ""; // brak nazwy - Outlook podstawia adres
  if (n.includes(",")) {
    const [naz, im] = n.split(",").map((x) => x.trim());
    return { imie: im ?? "", nazwisko: naz ?? "" };
  }
  const cz = n.split(/\s+/).filter(Boolean);
  return { imie: cz[0] ?? "", nazwisko: cz.slice(1).join(" ") };
}

/** Jak w Power Apps: imie zakonczone na "a" -> K, inaczej M. */
export function plecZImienia(imie: string): "K" | "M" {
  return imie.trim().toLowerCase().endsWith("a") ? "K" : "M";
}

/** Odpowiednik GUID() z Power Apps - male litery, z myslnikami. */
export function nowyGuid(): string {
  const c = crypto as any;
  if (typeof c.randomUUID === "function") return c.randomUUID();
  const b = crypto.getRandomValues(new Uint8Array(16));
  b[6] = (b[6] & 0x0f) | 0x40;
  b[8] = (b[8] & 0x3f) | 0x80;
  const h = Array.from(b, (x) => x.toString(16).padStart(2, "0")).join("");
  return `${h.slice(0, 8)}-${h.slice(8, 12)}-${h.slice(12, 16)}-${h.slice(16, 20)}-${h.slice(20)}`;
}

export type Nota = {
  id: string;
  data: string;
  osoba: string;
  rodzaj: string;
  rodzajId: string;     // TypeLookupId - do edycji
  tresc: string;
  autor: string;
  followUp: string;
  followDone: boolean;
  opps: string;         // Opps_ID - GUID szansy, gdy notatka dotyczy szansy
};

export function toNota(item: ListItem): Nota {
  const f = item.fields as Record<string, unknown>;
  return {
    id: s(f.id),
    data: s(f.Data),
    osoba: s(f.Osoba),
    rodzaj: s(f.Rodzaj1) || s(f.Type),
    rodzajId: s(f.TypeLookupId),
    tresc: s(f.Totatka),               // uwaga: literowka w nazwie kolumny
    autor: s(f.Creator),
    followUp: s(f.Follow_up),
    followDone: f.FollowUpStatus === true,
    opps: s(f.Opps_ID),
  };
}

export type TecsRow = {
  id: string;
  nazwa: string;
  wartosc: number;
  etap: string;
  poczatek: string;
  koniec: string;
  tp: boolean;
  demo: boolean;
  klient: string;
  guid: string;       // TECSGUID - laczy z licencjami w Oferty_TECS
  firmaId: string;    // FirmaID
  odbiorca: string;   // Odbiorca_oferty - dla kogo kody demo
  guidOld: string;    // TECSGUID_OLD - poprzedni TECS, z ktorego mozna przepisac licencje
  followUp: string;   // Data_Follow_Up
  etag: string;       // @odata.etag - kontrola wspolbieznosci przy zmianie etapu / edycji
  pola: Record<string, unknown>; // surowe pola - do widoku szczegolow i edycji
};

export function toTecs(item: ListItem): TecsRow {
  const f = item.fields as Record<string, unknown>;
  return {
    id: s(f.id),
    nazwa: s(f["Nazwaszansysprzeda_x017c_y_TECS"]),
    wartosc: n(f["Warto_x015b__x0107_"]),
    etap: s(f.Etap1),
    poczatek: s(f["Pocz_x0105_tek_TECS"]),
    koniec: s(f.Koniec_TECS),
    tp: f.TP === true,
    demo: f.Demo === true,
    klient: s(f.Numer_klienta),
    guid: s(f.TECSGUID),
    firmaId: s(f.FirmaID),
    odbiorca: s(f.Odbiorca_oferty),
    guidOld: s(f.TECSGUID_OLD),
    followUp: s(f.Data_Follow_Up),
    etag: s(f["@odata.etag"] ?? (item as any)["@odata.etag"]),
    pola: f,
  };
}

/** Pozycja produktowa TECS (Oferty_TECS). */
export type ProduktTecs = {
  id: string;
  produkt: string;     // ProducID (sic) - nazwa licencji z cennika
  ile: number;
  cena: number;        // Cena_TECS - cena jednostkowa
  rabat: number;
  tpValue: number;
  leaseValue: number;
  suma: number;        // SUMA
  tp: boolean;
  lease: boolean;
  leaseDni: number;    // LeaseDays
  koniec: string;
  firmaId: string;     // FirmaIDofertaTECS
  oppsId: string;      // TECSOppsIDoferta
};

export function toProduktTecs(item: ListItem): ProduktTecs {
  const f = item.fields as Record<string, unknown>;
  return {
    id: s(f.id),
    produkt: s(f.ProducID),
    ile: n(f.Liczba_licencji),
    cena: n(f.Cena_TECS),
    rabat: n(f.Rabat),
    tpValue: n(f.TPValue),
    leaseValue: n(f.LeaseValue),
    suma: n(f.SUMA),
    tp: f.TP === true,
    lease: f.Lease === true,
    leaseDni: n(f.LeaseDays),
    koniec: s(f.Koniec_TECS),
    firmaId: s(f.FirmaIDofertaTECS),
    oppsId: s(f.TECSOppsIDoferta),
  };
}

export type Szansa = {
  id: string;
  nazwa: string;
  wartosc: number;
  status: string;
  typ: string;
  otwarcie: string;
  guid: string;         // OPPSGUID -> Notes.Opps_ID
  firma: string;        // Firma - etykieta firmy, idzie do notatki
  etap: string;         // Etap1 - etap sprzedazy (stepper), idzie tez do notatki
  companyGuid: string;  // IdentyfikatorCompanies
  etag: string;         // @odata.etag - kontrola wspolbieznosci przy zmianie statusu/etapu
  pola: Record<string, unknown>; // surowe pola - status zamkniecia czytamy po nazwie wewnetrznej
};

export function toSzansa(item: ListItem): Szansa {
  const f = item.fields as Record<string, unknown>;
  return {
    id: s(f.id),
    nazwa: s(f["Nazwaszansysprzeda_x017c_y"]),
    wartosc: n(f["Warto_x015b__x0107_wEUR"]),
    status: s(f.Status1),
    typ: s(f.TypBiznesu),
    otwarcie: s(f.Dataotwarcia),
    guid: s(f.OPPSGUID),
    firma: s(f.Firma),
    etap: s(f.Etap1),
    companyGuid: s(f.IdentyfikatorCompanies),
    etag: s(f["@odata.etag"] ?? (item as any)["@odata.etag"]),
    pola: f,
  };
}

export type Licencja = {
  id: string;
  nazwa: string;
  ile: number;
  koniec: string;
  tp: boolean;
  demo: boolean;
  opps: string;
  zamkniete: boolean;
};

/** Z listy Oferty_TECS - licencje powiazane z TECS. */
export function toLicencjaTecs(item: ListItem): Licencja {
  const f = item.fields as Record<string, unknown>;
  return {
    id: s(f.id),
    nazwa: s(f.ProducID),
    ile: n(f.Liczba_licencji),
    koniec: s(f.Koniec_TECS),
    tp: f.TP === true,
    demo: f.Demo === true,
    opps: s(f.TECSOppsIDoferta),
    zamkniete: s(f.TECSclosed) !== "NIE",   // uwaga: tekst, nie boolean
  };
}

/** Z listy StanPosiadania - licencje bez TECS. */
export function toLicencjaBez(item: ListItem): Licencja {
  const f = item.fields as Record<string, unknown>;
  return {
    id: s(f.id),
    nazwa: s(f.NazwaLicencji),
    ile: n(f.LiczbaLicencji),
    koniec: s(f.KoniecTECS),
    tp: false,
    demo: false,
    opps: "",
    zamkniete: true,
  };
}

export type Szkolenie = {
  id: string;
  nazwa: string;
  data: string;
  osoba: string;
};

export function toSzkolenie(item: ListItem): Szkolenie {
  const f = item.fields as Record<string, unknown>;
  return {
    id: s(f.id),
    nazwa: s(f.NazwaSzkolenia),
    data: s(f.DataSzkolenia),
    osoba: s(f.ImieNazwisko),
  };
}

/** Ile dni do daty. Ujemne = po terminie. */
export function dniDo(iso: string): number | null {
  if (!iso) return null;
  const d = new Date(iso);
  if (isNaN(d.getTime())) return null;
  const dzis = new Date();
  dzis.setHours(0, 0, 0, 0);
  d.setHours(0, 0, 0, 0);
  return Math.round((d.getTime() - dzis.getTime()) / 86400000);
}

/** Kolor paska i tekstu wedlug liczby dni do konca. */
export function kolorTerminu(dni: number | null): string {
  if (dni === null) return "#9CA3AF";
  if (dni <= 30) return "#DC2626";
  if (dni <= 90) return "#D97706";
  return "#16A34A";
}

export const ETAP_KOLORY: Record<string, [string, string]> = {
  "Przygotowanie oferty": ["#D97706", "#FEF3E2"],
  "Oferta wysłana": ["#2F6FD6", "#E8F0FD"],
  "Zamówione": ["#16A34A", "#E6F6EE"],
  "Wystawiona faktura": ["#0D9488", "#E0F5F3"],
  "Rezygnacja z TECS": ["#DC2626", "#FEE2E2"],
  "Tech Pool": ["#5B6CDB", "#EEF0FC"],
  "Koniec Tech Poola": ["#6B7280", "#F3F4F6"],
  "Demo": ["#7C3AED", "#F1EAFE"],
};

export const ETAP_SKROTY: Record<string, string> = {
  "Przygotowanie oferty": "W przygotowaniu",
  "Oferta wysłana": "Wysłana",
  "Rezygnacja z TECS": "Rezygnacja",
};
