// src/taskpane/outlook.ts
/* global document, Office, console, setTimeout, window, navigator, clearTimeout */
import {
  getListItems, getTicketsByField, updateTicketFields, getCompanies2, getOpiekunowie,
  wyczyscCacheFirm, getContactByEmail, createTicket, getSerwisanci, getContactsByCompany,
  getTicketsFiltered, getMojMail, getByCompany, getCompanyByGuid,createListItem, getLookupOptions, getMojaNazwa, deleteListItem,
  getTicketByNumber
} from "./graphService";
import {
  toTicket, toCompany, formatData, dniOdZmiany, KOLOR_STATUSU, ticketIdZTematu, Ticket, Company,
  toContact, Contact, toCompanyFull, CompanyFull, toNota, toTecs, toSzansa,
  toLicencjaTecs, toLicencjaBez, Licencja, toSzkolenie, dniDo, kolorTerminu,
  ETAP_KOLORY, ETAP_SKROTY,
} from "./crm";

// ---------------------------------------------------------------- start

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) return;

  document.getElementById("sideload-msg")!.style.display = "none";
  document.getElementById("app-body")!.style.display = "block";

  document.getElementById("run")!.onclick = () => void odswiez();
  document.getElementById("nav-tickets")!.onclick = () => void pokazZgloszenia();
  document.getElementById("nav-companies")!.onclick = () => void pokazFirmy();

  Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => {
    if (widok === "zgloszenia") void pokazZgloszenia();
  });

  void pokazZgloszenia();
});

// ---------------------------------------------------------------- stan

type Widok = "zgloszenia" | "firmy" | "firma";
let widok: Widok = "zgloszenia";
let ostatniaFirma: Company | null = null;

let zakresZgloszen: "mail" | "firma" | "wszystkie" = "mail";
let opiekunZgloszen = "";

type Zakladka =
  | "Kontakty" | "Supporty" | "Notatki" | "TECS"
  | "Szanse" | "Posiadane" | "Szkolenia" | "Uwagi";

let zakladkaFirmy: Zakladka = "Kontakty";
let tecsTryb: "Bieżące" | "Archiwalne" | "Demo" = "Bieżące";
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

function ustawNaglowek(tytul: string, aktywny: "nav-tickets" | "nav-companies") {
  document.getElementById("nav-title")!.textContent = tytul;
  ["nav-tickets", "nav-companies"].forEach((id) =>
    document.getElementById(id)!.classList.toggle("crm-nav-aktywny", id === aktywny)
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

/** Mala pastylka z kolorem tekstu i tla. */
function pastylka(tekst: string, kolor: string, tlo: string): HTMLElement {
  const p = el("span", "crm-pill", tekst);
  p.style.color = kolor;
  p.style.background = tlo;
  return p;
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

// ---------------------------------------------------------------- zgloszenia

function karta(t: Ticket): HTMLElement {
  const kolor = KOLOR_STATUSU[t.status] ?? "#3860B2";

  const card = el("div", "crm-card");
  card.style.borderLeftColor = kolor;

  const head = el("div", "crm-card-head");
  head.tabIndex = 0;
  head.setAttribute("role", "button");

  const top = el("div", "crm-card-top");
  top.appendChild(el("span", "crm-firma", t.firma || "(brak firmy)"));
  const idEl = el("span", "crm-id", `MESco-ID-${t.ticketId}`);
  zrobKopiowalny(idEl, `[MESco-ID-${t.ticketId}]`, "skopiowano numer");
  top.appendChild(idEl);
  head.appendChild(top);

  const r1 = el("div", "crm-row");
  r1.appendChild(el("span", "crm-osoba", t.osoba || "—"));
  const czasEtykieta = el("span", "crm-czas", t.czas ? `${t.czas} h` : "—");
  r1.appendChild(czasEtykieta);
  head.appendChild(r1);

  head.appendChild(el("div", "crm-opiekun", t.opiekun || "—"));
  
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
  foot.appendChild(el("span", "crm-data", formatData(t.data)));
  const lock = el("span", "crm-lock", t.zamkniete ? "zamknięte 🔒" : "otwarte 🔓");
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
    lock.textContent = t.zamkniete ? "zamknięte 🔒" : "otwarte 🔓";
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

  zapisz.onclick = () => {
    const stary = t.opis.trim();
    const nowy = ta.value.trim();
    if (stary.length > 0 && !nowy.includes(stary)) {
      ostrzezenie.innerHTML = "";
      ostrzezenie.appendChild(el("div", undefined, nowy
        ? "Zmieniasz istniejący opis — poprzedni zniknie bez śladu."
        : "Kasujesz cały opis."));
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
    box.appendChild(el("div", "crm-nowe-info",
      `${nadawca} nie figuruje w kontaktach — dodaj go w CRM, żeby założyć zgłoszenie.`));
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

  btn.onclick = async () => {
    const wybrany = serwisanci[Number(wybor.value)];
    if (!wybrany) {
      komunikat.textContent = "Wybierz opiekuna supportu";
      komunikat.className = "crm-komunikat crm-blad";
      return;
    }

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
      MailSubject: temat,
      MailID: mailIdRest(item),
      MailCreator: wybrany.mail,
      Dodatkowy_opiekun: wybrany.nazwa,
    };

    if (wybrany.siteId !== null) {
      pola["Opiekun_supportuLookupId"] = String(wybrany.siteId);
    }

    try {
      await createTicket(pola);

      box.className = "crm-nowe crm-nowe-ok";
      box.innerHTML = "";
      box.appendChild(el("div", "crm-nowe-info",
        `Zgłoszenie założone dla ${wybrany.nazwa}. Numer nadawany przez przepływ — pojawi się w ciągu kilku minut.`));
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

  const wczytaj = async () => {
    lista.innerHTML = "";
    s.textContent = "Pobieranie zgłoszeń...";
    s.classList.remove("crm-error");

    try {
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

function kartaFirmy(c: Company): HTMLElement {
  const card = el("div", "crm-firma-card");
  card.tabIndex = 0;
  card.setAttribute("role", "button");

  const top = el("div", "crm-firma-top");
  top.appendChild(el("span", "crm-firma-nazwa", c.nazwa || "(bez nazwy)"));
  top.appendChild(el("span", "crm-tecs", `TECS: ${c.tecs}`));
  card.appendChild(top);

  const adres = [c.ulica, [c.kod, c.miasto].filter(Boolean).join(" ")].filter(Boolean);
  if (adres.length) card.appendChild(el("div", "crm-adres", adres.join("\n")));

  if (c.opiekun) card.appendChild(el("div", "crm-firma-opiekun", `Opiekun: ${c.opiekun}`));

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
      firmy.forEach((c) => lista.appendChild(kartaFirmy(c)));
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

// ---------------------------------------------------------------- kontakty firmy

function kartaKontaktu(k: Contact, c: Company): HTMLElement {
  const card = el("div", "crm-kontakt-card");

  const top = el("div", "crm-firma-top");
  top.appendChild(el("span", "crm-firma-nazwa", k.imieNazwisko || "(bez nazwiska)"));
  if (!k.aktywny) top.appendChild(el("span", "crm-nieaktywny", "nie pracuje"));
  card.appendChild(top);

  if (k.stanowisko) card.appendChild(el("div", "crm-adres", k.stanowisko));

  if (k.email) {
    const mail = el("div", "crm-kontakt-linia", k.email);
    zrobKopiowalny(mail, k.email, "skopiowano adres");
    card.appendChild(mail);
  }

  if (k.telefon) {
    const tel = document.createElement("a");
    tel.className = "crm-kontakt-linia crm-tel";
    tel.href = `tel:${k.telefon.replace(/[^\d+]/g, "")}`;
    tel.textContent = k.telefon;
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
  const rodzaje = el("div", "crm-chipy");
  box.appendChild(rodzaje);

  let rodzaj = "";
  let rodzajId = "";
  try {
    const opcje = await getLookupOptions("activities", "Notes_types");
    if (!opcje.length) throw new Error("lista rodzajów jest pusta");
    opcje.forEach((o) => {
      const b = el("button", "crm-chip-wybor", o.etykieta) as HTMLButtonElement;
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
  box.appendChild(el("div", "crm-firma-tytul", f?.skrocona || c.nazwa));
  if (f?.pelna && f.pelna !== f.skrocona) box.appendChild(el("div", "crm-firma-pelna", f.pelna));

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
  pills.appendChild(pastylka(st || "brak statusu", sc, sb));

  const dniNDA = dniDo(f?.dataNDA ?? "");
  const ndaKolor: [string, string] =
    !f?.dataNDA ? ["#6B7280", "#F3F4F6"]
    : dniNDA !== null && dniNDA < 0 ? ["#DC2626", "#FEE2E2"]
    : dniNDA !== null && dniNDA <= 30 ? ["#D97706", "#FEF3E2"]
    : ["#16A34A", "#E6F6EE"];
  pills.appendChild(
    pastylka("NDA " + (f?.dataNDA ? "do " + dataPL(f.dataNDA) : "brak"), ndaKolor[0], ndaKolor[1])
  );

  if (f?.opiekun) pills.appendChild(pastylka(f.opiekun, "#3860B2", "#EEF0FC"));
  if (tecsAlert) {
    const p = pastylka(
      `⚠ TECS kończy się za ${tecsAlert.dni} dni` +
        (tecsAlert.ile > 1 ? ` (${tecsAlert.ile} poz.)` : ""),
      "#DC2626",
      "#FEE2E2"
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
  box.appendChild(pills);

  // --- przyciski akcji ---
  const akcje = el("div", "crm-firma-akcje");

  if (f?.ulica || f?.miasto) {
    const mapa = el("button", "crm-akcja", "📍 Nawiguj") as HTMLButtonElement;
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
    const nipBtn = el("button", "crm-akcja", `⧉ ${widoczny}`) as HTMLButtonElement;
    nipBtn.title = "Kopiuj NIP";
    nipBtn.onclick = async () => {
      const ok = await kopiuj(czysty);
      blysk(nipBtn, ok ? "skopiowano" : "nie udało się");
    };
    akcje.appendChild(nipBtn);
  }

  const dok = el("button", "crm-akcja", "📁 Dokumenty") as HTMLButtonElement;
  dok.onclick = () =>
    window.open(
      "https://mescocompl.sharepoint.com/sites/CRMSystemPOC/Shared%20Documents/DokumentyKlienci/" +
        encodeURIComponent(f?.skrocona || c.nazwa),
      "_blank"
    );
  akcje.appendChild(dok);
  box.appendChild(akcje);

  // --- liczniki godzin ---
  const liczniki = el("div", "crm-liczniki");
  ([
    ["TECS", f?.tecsH ?? 0, "#2F6FD6"],
    ["Support Day", f?.supportDay ?? 0, "#D97706"],
    ["PreSales", f?.preSalesH ?? 0, "#3860B2"],
  ] as [string, number, string][]).forEach(([lbl, val, kol]) => {
    const kafel = el("div", "crm-licznik");
    kafel.appendChild(el("div", "crm-licznik-lbl", lbl));
    const v = el("div", "crm-licznik-val", godziny(val));
    v.style.color = val === 0 ? "#9CA3AF" : kol;
    kafel.appendChild(v);
    liczniki.appendChild(kafel);
  });
  box.appendChild(liczniki);

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

  if (!wszyscy.length) {
    s.textContent = "Kontakty — 0";
    return pusto(box, "Brak kontaktów dla tej firmy.");
  }

  // --- szukanie ---
  const szukaj = document.createElement("input");
  szukaj.type = "search";
  szukaj.className = "crm-szukaj";
  szukaj.placeholder = "szukaj osoby, stanowiska, maila";
  box.appendChild(szukaj);

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
    top.appendChild(el("span", "crm-nota-data", `${dataPL(nt.data)} · ${nt.rodzaj}`));

    const moja = !!ja && nt.autor.trim().toLowerCase() === ja;
    const kosz = el("button", "crm-usun", "🗑") as HTMLButtonElement;
    kosz.title = "Usuń notatkę";
    if (moja) top.appendChild(kosz);
    k.appendChild(top);

    if (nt.osoba) k.appendChild(el("div", "crm-nota-osoba", nt.osoba));
    if (nt.followUp) {
      const po = (dniDo(nt.followUp) ?? 0) < 0 && !nt.followDone;
      k.appendChild(
        pastylka(
          (nt.followDone ? "✓ " : "") + "Follow-up " + dataPL(nt.followUp),
          po ? "#DC2626" : "#D97706",
          po ? "#FEE2E2" : "#FEF3E2"
        )
      );
    }
    if (nt.tresc) k.appendChild(el("div", "crm-nota-tresc", nt.tresc));
    if (nt.autor) k.appendChild(el("div", "crm-nota-autor", nt.autor));

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
    k.appendChild(el("div", "crm-kafel-nazwa", t.nazwa));

    const rzad = el("div", "crm-kafel-rzad");
    const wart =
      t.demo ? (t.klient ? "Klient " + t.klient : "Kody demo") : t.tp ? "TechPool" : euro(t.wartosc);
    rzad.appendChild(el("span", "crm-kafel-wartosc", wart));
    const [kol, tlo] = ETAP_KOLORY[t.etap] ?? ["#6B7280", "#F3F4F6"];
    rzad.appendChild(pastylka(ETAP_SKROTY[t.etap] ?? t.etap, kol, tlo));
    k.appendChild(rzad);

    const d = dniDo(t.koniec);
    const meta = el(
      "div",
      "crm-kafel-meta",
      (t.demo
        ? `${dataPL(t.poczatek)} → ${dataPL(t.koniec)}`
        : `${t.tp ? "TechPool" : "TECS"} do ${dataPL(t.koniec)}`) +
        (d === null ? "" : ` · ${d < 0 ? "po terminie" : d + (d === 1 ? " dzień" : " dni")}`)
    );
    meta.style.color = kolorTerminu(d);
    k.appendChild(meta);
    k.appendChild(pasekTerminu(d));
    box.appendChild(k);
  });
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
    k.appendChild(el("div", "crm-kafel-nazwa", o.nazwa));
    const rzad = el("div", "crm-kafel-rzad");
    rzad.appendChild(el("span", "crm-kafel-wartosc", euro(o.wartosc)));
    if (o.typ) rzad.appendChild(pastylka(o.typ, "#2F6FD6", "#E8F0FD"));
    k.appendChild(rzad);

    const meta = el(
      "div",
      "crm-kafel-meta",
      `Otwarta ${dataPL(o.otwarcie)}` + (o.status === "Aktywny" ? "" : ` · ${o.status}`)
    );
    meta.style.color =
      o.status === "Zamknięty wygrany" ? "#16A34A" : o.status === "Zamknięty przegrany" ? "#DC2626" : "#6B7280";
    k.appendChild(meta);
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
    const top = el("div", "crm-kafel-rzad");
    top.appendChild(el("span", "crm-kafel-nazwa", l.nazwa));
    top.appendChild(el("span", "crm-licencje", `${l.ile} lic.`));
    k.appendChild(top);

    if (l.opps) k.appendChild(el("div", "crm-kafel-opps", l.opps));

    const d = dniDo(l.koniec);
    const meta = el(
      "div",
      "crm-kafel-meta",
      !l.koniec
        ? "bez TECS"
        : `TECS do ${dataPL(l.koniec)}` + (d === null ? "" : ` · ${d < 0 ? "wygasł" : d + " dni"}`)
    );
    meta.style.color = d !== null && d < 0 ? "#6B7280" : kolorTerminu(d);
    k.appendChild(meta);

    if (l.tp || l.demo) k.appendChild(pastylka(l.tp ? "TechPool" : "Demo", "#7C3AED", "#F1EAFE"));
    if (d !== null && d >= 0) k.appendChild(pasekTerminu(d));
    box.appendChild(k);
  });
}

async function zakladkaSzkolenia(box: HTMLElement, s: HTMLElement, c: Company) {
  // uwaga: ta lista uzywa GUIDFirmy, nie IdentyfikatorCompanies
  const items = await getByCompany("UczestnicySzkolen", c.guid, { pole: "GUIDFirmy", top: 100 });
  const lista = items.map(toSzkolenie).sort((a, b) => (a.data || "").localeCompare(b.data || ""));
  s.textContent = `Szkolenia — uczestników: ${lista.length}`;
  if (!lista.length) return pusto(box, "Brak szkoleń dla tej firmy.");

  lista.forEach((z) => {
    const k = el("div", "crm-kafel");
    k.appendChild(el("div", "crm-kafel-nazwa", z.nazwa));
    k.appendChild(el("div", "crm-kafel-meta", `${z.data} · ${z.osoba}`));
    box.appendChild(k);
  });
}

function zakladkaUwagi(box: HTMLElement, s: HTMLElement) {
  s.textContent = "Uwagi do firmy";
  const t = (firmaPelna?.uwagi ?? "").trim();
  box.appendChild(el("div", t ? "crm-uwagi" : "crm-pusto", t || "Brak uwag dla tej firmy."));
}

// ---------------------------------------------------------------- odswiezanie

async function odswiez() {
  if (widok === "firmy") {
    wyczyscCacheFirm();
    return pokazFirmy();
  }
  if (widok === "firma" && ostatniaFirma) {
    firmaPelna = null; // wymus ponowne pobranie danych firmy
    return pokazFirma(ostatniaFirma);
  }
  return pokazZgloszenia();
}
