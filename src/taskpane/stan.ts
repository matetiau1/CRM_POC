// src/taskpane/stan.ts
// Stan wspoldzielony miedzy modulami. Eksportowany przez funkcje (getX/setX),
// nie jako gole "export let" - przy "let" import bywa kopia, nie referencja.

// karta firmy otwarta z paska klienta w Support - "wroc" prowadzi z powrotem do zgloszenia
let powrotDoSupportu = false;

export function getPowrotDoSupportu(): boolean {
  return powrotDoSupportu;
}

export function setPowrotDoSupportu(v: boolean): void {
  powrotDoSupportu = v;
}
