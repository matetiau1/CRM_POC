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
    aktywny: f.Status === true,
  };
}