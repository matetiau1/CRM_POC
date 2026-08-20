// src/taskpane/outlook.ts
/* global document, Office, console, setTimeout */
import { getListItems, getTicketsByField, updateTicketFields, getCompanies2, getOpiekunowie, wyczyscCacheFirm, getContactByEmail, createTicket, getSerwisanci, getContactsByCompany,getTicketsFiltered } from "./graphService";
import { toTicket, toCompany, formatData, dniOdZmiany, KOLOR_STATUSU, ticketIdZTematu, Ticket, Company, toContact, Contact } from "./crm";

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

type Widok = "zgloszenia" | "firmy" | "firma" | "kontakty";
let widok: Widok = "zgloszenia";
let ostatniaFirma: Company | null = null;
let zakresZgloszen: "mail" | "firma" | "wszystkie" = "mail";
let opiekunZgloszen = "";


const kontaktCache = new Map<string, any>();
 
async function kontaktDlaNadawcy(email: string): Promise<any | null> {
  const klucz = email.toLowerCase();
  if (kontaktCache.has(klucz)) return kontaktCache.get(klucz);
  const k = await getContactByEmail(klucz);
  kontaktCache.set(klucz, k);
  return k;
}
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

      // data zamknięcia: ustawiamy przy zamykaniu, czyścimy przy ponownym otwarciu
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
/** Identyfikator wiadomości w formacie REST - Office.js domyślnie zwraca EWS. */
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

async function panelNowego() {
  const item = Office.context.mailbox.item as Office.MessageRead;
  if (!item) return;

  const temat = item.subject ?? "";
  if (ticketIdZTematu(temat) !== null) return; // wątek istniejącego zgłoszenia

  const nadawca = (item.from?.emailAddress ?? "").toLowerCase();
  if (!nadawca) return;
  if (nadawca.endsWith("@mesco.com.pl")) return; // support zakładamy dla klienta

  const box = el("div", "crm-nowe");
  root().insertBefore(box, root().children[1] ?? null);

  let kontakt: any = null;
  try {
    kontakt = await getContactByEmail(nadawca);
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

  // --- wybór opiekuna supportu ---
  box.appendChild(el("label", "crm-label", "Opiekun supportu"));
  const wybor = document.createElement("select");
  wybor.className = "crm-select";
  box.appendChild(wybor);

  const zalogowany = (Office.context.mailbox.userProfile.emailAddress ?? "").toLowerCase();
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

  // zalogowanego nie ma w MEScoStaff - wymuszamy świadomy wybór
  const jestNaLiscie = serwisanci.some((s) => s.mail === zalogowany);
  if (!jestNaLiscie) {
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

    // pole Osoba - tylko gdy mamy Id uzytkownika witryny
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
  const opcje = [
    { v: "mail", t: nadawca ? `Osoba: ${nadawca}` : "Z tej wiadomości" },
    { v: "firma", t: "Firma z tej wiadomości" },
    { v: "wszystkie", t: "Wszystkie zgłoszenia" },
  ];
  opcje.forEach((o) => {
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

// ---------------------------------------------------------------- firmy

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

  const otworz = () => void pokazZgloszeniaFirmy(c);
  card.onclick = otworz;
  card.onkeydown = (e: KeyboardEvent) => {
    if (e.key === "Enter" || e.key === " ") { e.preventDefault(); otworz(); }
  };

  return card;
}

// ===== Zamień funkcję pokazFirmy w src/taskpane/outlook.ts na poniższą =====
// Zmień też import na górze pliku:
//   import { getListItems, getTicketsByField, updateTicketFields, getCompanies2, getOpiekunowie } from "./graphService";

async function pokazFirmy(fraza = "", opiekunWybrany = "") {
  widok = "firmy";
  ustawNaglowek("Firmy", "nav-companies");
  root().innerHTML = "";

  // --- pasek filtrów ---
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

  // --- lista opiekunów (raz na sesję, potem z pamięci) ---
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

async function pokazZgloszeniaFirmy(c: Company) {
  widok = "firma";
  ostatniaFirma = c;
  ustawNaglowek(c.nazwa || "Firma", "nav-companies");
  root().innerHTML = "";

  root().appendChild(naglowekFirmy(c, "zgloszenia"));

  const s = status("Pobieranie zgłoszeń firmy...");

  try {
    if (!c.guid) {
      s.textContent = "Firma nie ma CompaniesGUID — brak powiązania ze zgłoszeniami.";
      s.classList.add("crm-error");
      return;
    }
    const items = await getTicketsByField("IdentyfikatorCompanies", c.guid, 50);
    s.textContent = `${c.nazwa} — zgłoszeń: ${items.length}`;
    if (!items.length) root().appendChild(el("div", "crm-pusto", "Brak zgłoszeń dla tej firmy."));
    items.map(toTicket).forEach((t) => root().appendChild(karta(t)));
  } catch (e: any) {
    console.error(e);
    s.textContent = `Błąd: ${e?.message ?? e}`;
    s.classList.add("crm-error");
  }
}

// ---------------------------------------------------------------- odswiezanie

async function odswiez() {
  if (widok === "firmy") {
    wyczyscCacheFirm();
    return pokazFirmy();
  }
  if (widok === "kontakty" && ostatniaFirma) return pokazKontaktyFirmy(ostatniaFirma);
  if (widok === "firma" && ostatniaFirma) return pokazZgloszeniaFirmy(ostatniaFirma);
  return pokazZgloszenia();
}


/** Schowek z zapasową ścieżką - navigator.clipboard bywa blokowany w panelu. */
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

/** Krótki komunikat przy elemencie, po kliknięciu "skopiowano". */
function blysk(obok: HTMLElement, tekst: string) {
  const b = el("span", "crm-blysk", tekst);
  obok.appendChild(b);
  setTimeout(() => b.remove(), 1400);
}

/** Zamienia element w przycisk kopiujący podaną wartość. */
function zrobKopiowalny(elem: HTMLElement, wartosc: string, etykieta = "skopiowano") {
  elem.classList.add("crm-kopiuj");
  elem.title = "Kliknij, aby skopiować";
  elem.onclick = async (e: MouseEvent) => {
    e.stopPropagation(); // żeby nie rozwinąć karty
    const ok = await kopiuj(wartosc);
    blysk(elem, ok ? etykieta : "nie udało się");
  };
}


// ============================================================ kontakty firmy

function kartaKontaktu(k: Contact, c: Company): HTMLElement {
  const card = el("div", "crm-kontakt-card");

  const top = el("div", "crm-firma-top");
  top.appendChild(el("span", "crm-firma-nazwa", k.imieNazwisko || "(bez nazwiska)"));
  if (!k.aktywny) top.appendChild(el("span", "crm-nieaktywny", "nieaktywny"));
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
  const komunikat = el("span", "crm-komunikat");
  const rzad = el("div", "crm-akcje");
  rzad.appendChild(btn);
  rzad.appendChild(komunikat);
  card.appendChild(rzad);

  btn.onclick = () => void formularzReczny(card, btn, komunikat, k, c);

  return card;
}

/** Ręczne zgłoszenie - bez MailID i MailSubject, bo nie wychodzi z wiadomości. */
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

  const zalogowany = (Office.context.mailbox.userProfile.emailAddress ?? "").toLowerCase();
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


// ============================================================ przelacznik w widoku firmy

/** Pasek: wróć + przełącznik Zgłoszenia / Kontakty. */
function naglowekFirmy(c: Company, aktywna: "zgloszenia" | "kontakty"): HTMLElement {
  const box = el("div");

  const powrot = el("button", "crm-powrot", "← wróć do firm") as HTMLButtonElement;
  powrot.onclick = () => void pokazFirmy();
  box.appendChild(powrot);

  const przelacznik = el("div", "crm-podnav");

  const bZgl = el("button", "crm-podnav-btn", "Zgłoszenia") as HTMLButtonElement;
  const bKon = el("button", "crm-podnav-btn", "Kontakty") as HTMLButtonElement;
  (aktywna === "zgloszenia" ? bZgl : bKon).classList.add("crm-podnav-aktywny");

  bZgl.onclick = () => void pokazZgloszeniaFirmy(c);
  bKon.onclick = () => void pokazKontaktyFirmy(c);

  przelacznik.appendChild(bZgl);
  przelacznik.appendChild(bKon);
  box.appendChild(przelacznik);

  return box;
}

async function pokazKontaktyFirmy(c: Company) {
  widok = "kontakty";
  ostatniaFirma = c;
  ustawNaglowek(c.nazwa || "Firma", "nav-companies");
  root().innerHTML = "";
  root().appendChild(naglowekFirmy(c, "kontakty"));

  const s = status("Pobieranie kontaktów...");

  try {
    if (!c.guid) {
      s.textContent = "Firma nie ma CompaniesGUID — brak powiązania z kontaktami.";
      s.classList.add("crm-error");
      return;
    }
    const items = await getContactsByCompany(c.guid);
    const kontakty = items.map(toContact)
      .sort((a, b) => a.imieNazwisko.localeCompare(b.imieNazwisko, "pl"));

    s.textContent = `${c.nazwa} — kontaktów: ${kontakty.length}`;
    if (!kontakty.length) root().appendChild(el("div", "crm-pusto", "Brak kontaktów dla tej firmy."));
    kontakty.forEach((k) => root().appendChild(kartaKontaktu(k, c)));
  } catch (e: any) {
    console.error(e);
    s.textContent = `Błąd: ${e?.message ?? e}`;
    s.classList.add("crm-error");
  }
}