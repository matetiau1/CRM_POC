// src/taskpane/rejestr.ts
/* global fetch, AbortController, setTimeout, clearTimeout */

/**
 * Dane firmy po NIP z Wykazu podatnikow VAT (Ministerstwo Finansow).
 * Usluga GUS (BIR) odrzuca zapytania z przegladarki - brak CORS - wiec
 * dodatek nie moze jej wolac bez posrednika. API MF ma CORS i nie wymaga klucza.
 * Limit: ok. 100 wyszukiwan dziennie z jednego adresu IP.
 */
const API_MF = "https://wl-api.mf.gov.pl/api/search/nip";

export type DaneFirmy = {
  nazwa: string;
  ulica: string;
  kod: string;
  miasto: string;
  statusVat: string;
};

export async function pobierzPoNip(nip: string): Promise<DaneFirmy | null> {
  const d = new Date();
  const dzis = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;

  const stop = new AbortController();
  const t = setTimeout(() => stop.abort(), 20000);
  try {
    const r = await fetch(`${API_MF}/${nip}?date=${dzis}`, { signal: stop.signal });
    const dane = (await r.json().catch(() => ({}))) as any;
    if (!r.ok) throw new Error(dane?.message || `rejestr VAT odpowiedział ${r.status}`);

    // brak w wykazie (np. firma niebedaca podatnikiem VAT) -> subject = null
    const s = dane?.result?.subject;
    if (!s?.name) return null;

    const adres = rozbijAdres(String(s.workingAddress || s.residenceAddress || ""));
    return { nazwa: String(s.name), statusVat: String(s.statusVat ?? ""), ...adres };
  } catch (e: any) {
    if (e?.name === "AbortError") throw new Error("rejestr VAT nie odpowiedział w ciągu 20 s");
    if (e instanceof TypeError) throw new Error("brak połączenia z rejestrem VAT");
    throw e;
  } finally {
    clearTimeout(t);
  }
}

/** "UL. DŁUGA 5/3, 00-001 WARSZAWA" -> ulica, kod, miasto (z normalna wielkoscia liter). */
function rozbijAdres(adres: string): { ulica: string; kod: string; miasto: string } {
  const i = adres.lastIndexOf(",");
  const przed = (i >= 0 ? adres.slice(0, i) : adres).trim();
  const po = (i >= 0 ? adres.slice(i + 1) : "").trim();
  const m = po.match(/^(\d{2}-\d{3})\s+(.*)$/);
  return {
    ulica: rzymskie(proper(przed.replace(/^ul\.?\s+/i, ""))),
    kod: m ? m[1] : "",
    miasto: rzymskie(proper(m ? m[2] : po)),
  };
}

// litery lacinskie i polskie - target es6 nie zna \p{L}
const LITERA = "A-Za-zÀ-ÖØ-öø-ÿĄĆĘŁŃŚŹŻąćęłńśźż";
const PROPER = new RegExp(`(^|[^${LITERA}])([${LITERA}])`, "g");

/** Odpowiednik Proper() z Power Apps: kazde slowo wielka litera, reszta mala. */
function proper(t: string): string {
  return t.toLowerCase().replace(PROPER, (_, a: string, b: string) => a + b.toUpperCase());
}

/** "Jana Pawła Ii" -> "Jana Pawła II" - tylko w adresach, w nazwach firm byloby ryzykowne. */
function rzymskie(t: string): string {
  return t.replace(/\b[IVX]{1,4}\b/gi, (w) => (/^[ivx]+$/i.test(w) && w.length > 1 ? w.toUpperCase() : w));
}

const FORMY_PRAWNE: [string, string][] = [
  ["Spółka Z Ograniczoną Odpowiedzialnością", "Sp. z o.o."],
  ["Spółka Jawna", "sp.j."],
  ["Spółka Akcyjna", "S.A."],
];

/** Pelna nazwa jak w Power Apps: rejestr daje wersaliki, skracamy forme prawna. */
export function pelnaNazwa(surowa: string): string {
  return FORMY_PRAWNE.reduce((n, [dlugo, krotko]) => n.split(dlugo).join(krotko), proper(surowa));
}

/** Suma kontrolna NIP - tylko ostrzezenie, Power Apps jej nie sprawdzal. */
export function nipPoprawny(nip: string): boolean {
  if (!/^\d{10}$/.test(nip)) return false;
  const wagi = [6, 5, 7, 2, 3, 4, 5, 6, 7];
  const suma = wagi.reduce((a, w, i) => a + w * Number(nip[i]), 0);
  return suma % 11 === Number(nip[9]);
}
