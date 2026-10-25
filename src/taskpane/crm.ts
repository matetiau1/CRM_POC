// src/taskpane/crm.ts
import { ListItem } from "./graphService";

export type TicketStatus = "Przydzielony" | "W toku" | "Oczekuje na odpowiedź" | "";

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
  };
}

export type Nota = {
  id: string;
  data: string;
  osoba: string;
  rodzaj: string;
  tresc: string;
  autor: string;
  followUp: string;
  followDone: boolean;
};

export function toNota(item: ListItem): Nota {
  const f = item.fields as Record<string, unknown>;
  return {
    id: s(f.id),
    data: s(f.Data),
    osoba: s(f.Osoba),
    rodzaj: s(f.Rodzaj1) || s(f.Type),
    tresc: s(f.Totatka),               // uwaga: literowka w nazwie kolumny
    autor: s(f.Creator),
    followUp: s(f.Follow_up),
    followDone: f.FollowUpStatus === true,
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
  };
}

export type Szansa = {
  id: string;
  nazwa: string;
  wartosc: number;
  status: string;
  typ: string;
  otwarcie: string;
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
  "Rezygnacja z TECS": ["#DC2626", "#FEE2E2"],
  "Demo": ["#7C3AED", "#F1EAFE"],
};

export const ETAP_SKROTY: Record<string, string> = {
  "Przygotowanie oferty": "W przygotowaniu",
  "Oferta wysłana": "Wysłana",
  "Rezygnacja z TECS": "Rezygnacja",
};
