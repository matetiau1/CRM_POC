# CRM MESco — dodatek do Outlooka

Specyfikacja projektu. Plik do wrzucania na początku każdej nowej rozmowy z Claude,
żeby nie odtwarzać kontekstu od zera.

Ostatnia aktualizacja: 15.08.2026

---

## 1. Cel i uzasadnienie

Zastąpienie dotychczasowego dodatku do Outlooka, który osadzał aplikację Power Apps
w iframe. Tamto rozwiązanie przestało działać na Windows 11 i macOS z powodu
blokowania ciasteczek trzeciej strony (`thirdPartyCookiesBlocked`) — logowanie zapętla
się albo próbuje otworzyć popup, który nie ma czym uwierzytelnić sesji.

Objawy w wersji developerskiej:
```
{code: 'UserNotLoggedIn', source: 'host', message: 'Popup was blocked by the browser'}
ssoSilent -> acquireToken -> executeCodeFlow -> _loginInteractiveWithCustomPopupAsync
```

Działa: Outlook on the web. Nie działa: Outlook desktop na W11 i macOS.
Na W10 działało (starszy WebView2).

**Nowe podejście:** natywny dodatek Office (HTML/TypeScript) czytający dane wprost
z SharePointa przez Microsoft Graph. Token przekazywany w nagłówku `Authorization`,
nie w ciasteczku — dlatego blokada ciasteczek trzeciej strony nie ma zastosowania.

Docelowo: pełny transfer funkcji z Power Apps do dodatku.

---

## 2. Środowisko developerskie

| Element | Wartość |
|---|---|
| System | macOS (Apple Silicon) |
| Edytor | VS Code |
| Node | v24.19.0 (uwaga: narzędzia Office testowane na LTS 20/22) |
| Szkielet | `yo office` → Office Add-in Task Pane, TypeScript, Outlook, **Add-in only manifest (XML)** |
| Katalog | `~/Documents/CRM/crm-mesco` |
| Serwer dev | `npm start` → `https://localhost:3000` |
| Repozytorium | https://github.com/matetiau1/CRM_MESco_POC |

### Sideload na macOS

Automatyczny sideload nie jest wspierany (`Sideload to the Outlook app is not supported`).
Manifest kopiuje się ręcznie:

```bash
cp manifest.xml ~/Library/Containers/com.microsoft.Outlook/Data/Documents/wef/
```
Potem pełne zamknięcie Outlooka (`pkill -x "Microsoft Outlook"`) i uruchomienie ponownie.

### Skrypt do package.json (zalecany)

```json
"scripts": {
  "sideload": "xmllint --noout manifest.xml && office-addin-manifest validate manifest.xml && cp manifest.xml ~/Library/Containers/com.microsoft.Outlook/Data/Documents/wef/"
}
```

`&&` gwarantuje, że zepsuty manifest nie zostanie skopiowany.

---

## 3. Wyciągnięte wnioski — czytać przed debugowaniem

Rzeczy, które kosztowały kilka godzin przy pierwszym podejściu:

1. **Outlook cicho odrzuca niepoprawny manifest** i zostawia poprzednią wersję.
   Objawia się jako „cache się nie aktualizuje". Zawsze najpierw:
   `xmllint --noout manifest.xml` (wskazuje element i linię),
   potem `npx office-addin-manifest validate manifest.xml`.
2. **`raw.githubusercontent.com` serwuje HTML jako `text/plain`** — Outlook nie
   wyrenderuje takiego taskpane'a. Nie nadaje się do hostowania dodatku.
3. **GitHub Pages wymaga `index.html`** w katalogu głównym. Plik o nazwie `index`
   (bez rozszerzenia) blokuje build.
4. **`SupportsPinning`** działa tylko w zagnieżdżonym bloku
   `VersionOverridesV1_1`, który musi mieć własną sekcję `Resources`.
5. **Zmiana `<Id>` (GUID)** to najpewniejszy sposób wymuszenia przeładowania —
   dla Outlooka to nowy dodatek, więc cache starego nie ma znaczenia.
6. **Dodatki instalowane przez Get Add-ins → My Add-ins zapisują się na skrzynce
   Exchange**, nie lokalnie, i mają pierwszeństwo nad plikiem z `wef`. Stare wpisy
   trzeba usunąć ręcznie.
7. Zmiany w kodzie: wystarczy zapis (webpack przeładuje). Zmiany w manifeście:
   kopiowanie + restart Outlooka.

---

## 4. Azure / uwierzytelnianie

| Element | Wartość |
|---|---|
| Nazwa aplikacji | SharePoint-List-AddIn-POC |
| Application (client) ID | `7277cae3-d05a-4a6f-8afe-2edb6b9b6365` |
| Directory (tenant) ID | `4ad84be0-f918-4dfa-8006-a408554ec666` |
| Object ID | `10ad251f-7e96-40c5-8780-832608f4b1b2` |
| Platforma | **Single-page application** (nie „Web" — SPA daje PKCE, wymagane przez MSAL.js) |
| Redirect URI | `https://localhost:3000/taskpane.html` |
| Biblioteka | `@azure/msal-browser` |

### Uprawnienia (delegowane, Microsoft Graph)

- `User.Read`
- `email`
- `Sites.Read.All`
- `Sites.ReadWrite.All` — **będzie potrzebne**, bo dodatek edytuje `Status` i `Uwagi`

### STATUS: ZGODA ADMINA UZYSKANA (15.08.2026) — dodatek czyta prawdziwe dane

Poniższy opis zostaje jako kontekst historyczny.

### Było: brak zgody admina

Tenant ma wyłączoną zgodę użytkownika (Enterprise applications → Consent and
permissions → User consent settings). Dlatego mimo „Admin consent required: No"
przy uprawnieniach, ekran logowania mówi „Wymaga zatwierdzenia administratora".
Nie da się tego obejść kodem — MSAL, Office SSO i SharePoint REST prowadzą przez
ten sam ekran zgody.

Link dla admina:
```
https://login.microsoftonline.com/4ad84be0-f918-4dfa-8006-a408554ec666/adminconsent?client_id=7277cae3-d05a-4a6f-8afe-2edb6b9b6365
```

Alternatywa węższa, łatwiejsza do obrony przed działem bezpieczeństwa:
`Sites.Selected` — uprawnienie ograniczone do konkretnych witryn, przypisywane
przez admina osobno.

### Logowanie: Office Dialog API, NIE loginPopup

`loginPopup` z MSAL **nie działa w panelu Outlooka** — okno się otwiera, ale nie ma
kanału powrotnego, więc MSAL kończy błędem `timed_out`. Objaw pokrewny:
`interaction_in_progress`, gdy kilka wywołań naraz prosi o token.

Rozwiązanie: `Office.context.ui.displayDialogAsync` → strona `dialog.html` →
MSAL `loginRedirect` → token wraca przez `Office.context.ui.messageParent`.

Wymagania:
- drugi SPA redirect URI w Entra: `https://localhost:3000/dialog.html`
- `src/dialog/dialog.html` + `dialog.ts`
- wpis `dialog` w `entry` i trzeci `HtmlWebpackPlugin` w `webpack.config.js`
- w `graphService.ts`: wspólna obietnica tokenu (blokada równoległych logowań)
  oraz `handleRedirectPromise()` przy inicjalizacji

Odświeżanie tokenu działa cicho przez refresh token (MSAL SPA + PKCE), bez ukrytej
ramki i bez ciasteczek — czyli bez problemu, który zabił Power Apps.
Logowanie pojawia się realnie raz na dobę.

---

## 5. Model danych

Witryna: `https://mescocompl.sharepoint.com/sites/CRMSystemPOC`

**UWAGA:** Graph filtruje listy po **nazwie wyświetlanej**, nie po nazwie w adresie URL.

| Nazwa wyświetlana (do kodu) | Nazwa wewnętrzna / URL | GUID |
|---|---|---|
| `Support` | Support | `1b5d3dea-d1eb-4170-8848-cbd26bc5a688` |
| `Contacts` | Contacts | — |
| **`Companies`** | **Firmy** ← inna niż widoczna! | — |

Na witrynie jest 41 list. Uwaga na myląco podobne: `Support1`, `SupportTickets`,
`SupportMail`, `SupportAttachments`, `FirmyTest`.

Graph **pomija puste pola** w odpowiedzi — brak pola w JSON nie znaczy, że kolumny nie ma.
Pełny schemat: `GET /sites/{siteId}/lists/{listId}/columns`.

---

### 5.1 Support — kolumny istotne

| Nazwa widoczna | Nazwa wewnętrzna | Typ | Uwagi |
|---|---|---|---|
| Data | `Data` | dateTime | data zgłoszenia |
| Firma | `Firma` | text | nazwa firmy |
| Osoba | `Osoba` | text | dla kogo support |
| OsobaMail | `OsobaMail` | text | **mail kontaktu — pozwala dopasować nadawcę bez zaglądania do Contacts** |
| Status | `Status` | boolean | `true` = zamknięte (kłódka) |
| StatusTicket | `StatusTicket` | choice | Przydzielony / W toku / Oczekuje na odpowiedź |
| Opis | `Opis` | text | opis zgłoszenia (prawdopodobnie to pole edytowane jako „uwagi") |
| Czas | `Czas` | number | **godziny** |
| TicketID | `TicketID` | number | numer bez prefiksu |
| Opiekun_supportu | `Opiekun_supportu` | **personOrGroup** | w danych jako `Opiekun_supportuLookupId` |
| Creator | **`Dodatkowy_opiekun`** | text | mylące: nazwa widoczna „Creator", wewnętrzna „Dodatkowy_opiekun" |
| MailCreator | `MailCreator` | text | **mail opiekuna — stabilny klucz do „moje zgłoszenia"** |
| MailSubject | `MailSubject` | text | już istnieje |
| MailID | `MailID` | text | już istnieje, format EWS |
| IdentyfikatorCompanies | `IdentyfikatorCompanies` | text | klucz do `Companies.CompaniesGUID` |
| IdentyfikatorContact | `IdentyfikatorContact` | text | klucz do `Contacts.ContactGUID` |
| Zmodyfikowane | `Modified` | dateTime | do wykrywania „utkniętych" statusów |

Pozostałe: `Datazako_x0144_czenia` (Data zakończenia), `Dla_sprzeda_x017c_y` (PreSales,
boolean), `AnsysGPT`, `AnsysHelp`, `SupportDay` (boolean), `Zlecil` (choice),
`Cel` (choice), `supportTime` i `Warto_x015b__x0107__x0020_w_x002` (Wartość w EUR) —
oba **calculated, tylko odczyt**.

**Nie ma kolumny `Uwagi`.** Do potwierdzenia, czy komentarz z Power Apps to `Opis`.

**Dla migracji przepływu:** `MailSubject`, `MailID`, `OsobaMail`, `MailCreator`
już istnieją — zmiana triggera na `when item created` nie wymaga nowych kolumn.

---

### 5.2 Companies (URL: /Lists/Firmy/) — kolumny istotne

| Nazwa widoczna | Nazwa wewnętrzna | Typ |
|---|---|---|
| Skrócona nazwa | `Nazwa` | text |
| Pełna nazwa | `Pe_x0142_nanazwa` | text |
| CompaniesGUID | `CompaniesGUID` | text |
| **Domen@** | `Domen_x0040_` | text |
| **DomenA** | `DomenA` | text |
| NIP | `VAT` | text |
| Status | `Status` | choice |
| Branża | `Bran_x017c_a` | choice |
| Uwagi | `Uwagi` | text |
| Opiekun | `Utworzoneprzez_x002f_Opiekun` | personOrGroup |
| Creator | `Creator` | text |

Adresowe: `Ulica`, `Miasto`, `Kodpocztowy`.
Pozostałe: `TECSh`, `PreSalesh`, `TECStotalGodziny` (calculated), `Oddzia_x0142_`,
`NIPjednostkiNadrzednej`, `DataNDA`, `LinkedIn`, `Soft`, `BizRaport`,
`WartoscTECS`, `MarzaTECS`, `SzkoleniaGratis`, `SzkolenieGratis`,
`SupportDay`, `ProspectStatus`, `Status1`, `Zrodlo1`,
`_x0179_r_x00f3_d_x0142_o` (Źródło, lookup), `Tw_x00f3_rcaFirmy`.

**DO ROZSTRZYGNIĘCIA:** dwie kolumny na domenę — `Domen_x0040_` czy `DomenA`?
Potrzebne do trzeciego poziomu kaskady zakresu.

---

### 5.3 Contacts — kolumny istotne

| Nazwa widoczna | Nazwa wewnętrzna | Typ |
|---|---|---|
| E-mail | `Email` | text |
| Imię | `Imi_x0119_` | text |
| Nazwisko | `Nazwisko` | text |
| Imię i nazwisko | `Imi_x0119__x0020_i_x0020_nazwisk` | **calculated, tylko odczyt** |
| ContactGUID | `ContactGUID` | text |
| IdentyfikatorCompanies | `IdentyfikatorCompanies` | text |
| FirmaID | `FirmaID` | text (etykieta, nie klucz) |
| Telefon | `Phone` | text |
| Stanowisko | `Stanowisko` | text |
| Rola | `RolawMEDDIC` | choice |
| Opiekun | `Opiekun` | personOrGroup |
| UWAGI | `UWAGI` | text |
| Status | `Status` | boolean |

Pozostałe: `P_x0142_e_x0107_` (Płeć), `Utworzonyprzez`, `Datautworzenia`,
`KontaktUtworzy_x0142_`, `Sta_x0142_awysy_x0142_ka`, `Tymczasowawysy_x0142_ka`,
`InnyAdresWysylki`, `nazwafirmyKontakt`, `ulicakontakt`, `miastoKontakt`,
`kodpocztowykontakt`, `LiczbaPolaczen`, `Wydzial`, `Katedra`.

Nazwy wewnętrzne SharePointa są **obcinane do 32 znaków** — stąd
`Imi_x0119__x0020_i_x0020_nazwisk` bez końcowego „o".

---

### 5.4 Relacje

```
Support.IdentyfikatorCompanies  →  Companies.CompaniesGUID
Support.IdentyfikatorContact    →  Contacts.ContactGUID
Contacts.IdentyfikatorCompanies →  Companies.CompaniesGUID
Support.MailCreator             →  mail opiekuna (filtr „moje zgłoszenia")
Support.OsobaMail               →  mail kontaktu (dopasowanie do nadawcy maila)
```

## 6. Logika biznesowa

### Statusy — ping-pong odpowiedzialności

| Status | Kto ma ruch | Znaczenie |
|---|---|---|
| Przydzielony | my | ma opiekuna, **nie było jeszcze odpowiedzi do klienta** |
| W toku | my | klient odpisał, czeka na naszą reakcję |
| Oczekuje na odpowiedź | klient | odpisaliśmy, czekamy |

Statusy zmieniają **przepływy reagujące na maile**, nie aplikacja (patrz sekcja 8).
Dodatek ich nie dotyka.

### Paleta kolorów — NIE ZMIENIAĆ (powiązana z innymi miejscami systemu)

Kolory to **tło aktywnej zakładki**; tekst aktywnej biały, nieaktywnej `#3860B2`.

| Zakładka | RGBA | HEX |
|---|---|---|
| Przydzielone | 220, 38, 38 | `#DC2626` |
| W toku | 245, 158, 11 | `#F59E0B` |
| Oczekują | 16, 185, 129 | `#10B981` |
| Wszystkie / marka | 56, 96, 178 | `#3860B2` |

Kolejność czerwony → pomarańczowy → zielony wyraża malejącą pilność.
Domyślna zakładka: **Przydzielone**.

### Kaskada zakresu — automatyczny wybór kontekstu

Panel zna otwartą wiadomość, więc wybiera zakres sam:

1. temat zawiera `[MESco-ID-XXXXX]` → **pokaż to jedno zgłoszenie**
2. adres nadawcy = `Support.OsobaMail` → **zgłoszenia tej osoby**
   (uproszczenie względem Power Apps: bez zaglądania do `Contacts`)
3. nadawca nieznany, domena pasuje do kolumny domeny w `Companies` →
   **zgłoszenia firmy** (po `IdentyfikatorCompanies`)
4. nic nie pasuje → **wszystkie od najnowszych**, z wyszukiwaniem

Użytkownik może zakres nadpisać ręcznie (przełącznik jak w Power Apps),
ale nie musi wybierać przy każdym mailu.

### Wykrywanie numeru zgłoszenia

Wzorzec: `\[MESco-ID-(\d+)\]`, bez rozróżniania wielkości liter, szukany w dowolnym
miejscu tematu (tematy mają przedrostki `RE:`, `ODP:`).

**UWAGA — niespójność w Power Apps do naprawy:** `OnStart` używa `\d{5}`, co łapie
dowolne pięć cyfr (kwota, data, numer oferty) i wstawia śmieć do `gblticketid`.
Widoczność przycisku „+" używa już poprawnego `\[MESco-ID-\d+\]`.

### Zakładki i liczniki

Zakładki filtrują po `StatusTicket`. W Power Apps trzy pierwsze pokazują **tylko
zgłoszenia zalogowanego użytkownika** (`Creator = User().FullName`), a „Wszystkie"
zdejmuje ten filtr i dodaje filtry po opiekunie i nazwie firmy.

Różnica w dodatku: **liczniki widoczne na wszystkich zakładkach naraz**
(w Power Apps tylko na aktywnej, bo `CountRows(Gallery.AllItems)`).

Filtr „moje zgłoszenia": w Power Apps `Creator = User().FullName` (porównanie nazw).
W dodatku **`MailCreator` = adres zalogowanego użytkownika** — stabilne, odporne na
zmianę nazwiska.

### Wyszukiwanie

- po firmie: `StartsWith(Firma, tekst)`
- po numerze: pole tekstowe zamiast listy rozwijanej — użytkownik wpisuje numer
  albo wkleja cały `[MESco-ID-14487]`, kod wyciąga cyfry.
  (W Power Apps `ForAll` po wszystkich zgłoszeniach z roku — zbyt ciężkie)
- po opiekunie: lista rozwijana; **zalecana osobna lista pomocnicza** `SupportOwners`
  z flagą `Aktywny`, zamiast `Distinct(Creator)` z roku

### Karta zgłoszenia — pola

| Pole | Źródło | W Power Apps |
|---|---|---|
| Firma | `Firma` | `🏢` |
| Osoba | `Osoba` | `🙋` |
| Opiekun | `Creator` | `🎧` |
| Data | `Data`, format `dd.mmmm.yyyy` pl-PL | `🗓️` |
| Numer | `MESco-ID-` + `TicketID` | — |
| Czas | `Czas` (godziny) | liczba w zegarze |
| Zamknięte | `Status` true/false | `🔒` / `🔓` |
| Status | `StatusTicket` | `🙋` / `📥` / `📤` |

**Propozycja:** pokazywać też datę ostatniej zmiany (`Modified`), bo statusy potrafią
utknąć przy awarii przepływu — „Przydzielony, zmiana 12 dni temu" to sygnał ostrzegawczy.

### Zapis z dodatku

Tylko dwa pola: `Status` (true/false) i `Uwagi` (nadpisanie, bez historii).

---

## 7. Nowe zgłoszenie (przycisk „+")

Widoczny gdy: temat **nie** zawiera `[MESco-ID-...]` **i** nadawca jest w `Contacts`.

Wypełnienie formularza:

| Pole | Wartość |
|---|---|
| Firma | `LookUp(Contacts, E-mail = nadawca).FirmaID` |
| Osoba | `LookUp(Contacts, E-mail = nadawca).'Imię i nazwisko'` |
| Opiekun | zalogowany użytkownik (DisplayName) |
| Data | dziś |
| Status | `false` |

Walidacja: opiekun nie może być pusty.

Po zapisie Power Apps wywołuje przepływ `Support_TicketCRM_forward`.

---

## 8. Przepływy Power Automate

### 8.1 `Support_TicketCRM_forward`

ID: `2e8091ad-a8f9-4957-a7a7-ee8639c07d54`
Trigger: **PowerAppV2** (wywoływany z Power Apps, 6 parametrów)

Kroki:
1. `Get_items` — najwyższy `TicketID` (sort desc, top 1)
2. `Compose` — nowy numer: `TicketID + 1`, albo `10000` jeśli lista pusta
3. `Compose_1` — temat `[MESco-ID-00000] oryginalny temat` (`formatNumber '00000'`)
4. `Update_item` — zapis `TicketID` i `StatusTicket = Przydzielony`
5. `Send_an_HTTP_request` — Graph PATCH na wiadomości w `support@mesco.com.pl`:
   zmiana `subject` **oraz** `singleValueExtendedProperties` `String 0x0070`
   (temat wątku — żeby odpowiedzi zachowały numer)
6. `Get_email_(V2)` + `Compose_2` — odczyt subject / conversationId / internetMessageId
7. `Post_message_in_a_chat_or_channel` — bot Teams do opiekuna
8. `Send_an_email_(V2)` — HTML potwierdzenie do klienta z numerem
9. `Mark_as_read_or_unread_(V3)` — oznaczenie jako nieprzeczytana
10. `Move_email` ×3 z opóźnieniami 1 s — tymczasowe obejście zalecone przez Microsoft
11. `Delay_2_1` 30 s → `Update_item_1` — **CELOWE**: mail rejestracyjny trafia do
    Elementów wysłanych i przepływ 8.2 ustawia „Oczekuje na odpowiedź";
    te 30 s cofają status na „Przydzielony", bo automat nie jest odpowiedzią
    merytoryczną

Parametry triggera: `text` = recordId, `text_1` = mailSubject, `text_2` = osoba
(w praktyce **nazwa folderu**), `text_3` = mailItemId, `text_4` = mailbox klienta,
`text_5` = creatorMail.

**Usterka:** `Move_email` używa `@triggerBody()?['mailbox']`, ale w schemacie pola
nazywają się `text_4` — „mailbox" to tylko etykieta w UI. Wyrażenie zwraca `null`
i konektor używa domyślnej skrzynki połączenia. Działa przypadkiem.

**Ryzyko:** numeracja „najwyższy + 1" to wyścig. Dziś nieszkodliwy (wywołania
pojedyncze), po zmianie triggera na `when item created` — realny.
Zalecenie: `concurrency = 1` w ustawieniach przepływu.

### 8.2 `SupportTicketAutomaterSupport_Out`

ID: `38da1bd0-184b-42fc-acc8-516de9306978`
Trigger: nowy mail w folderze Elementy wysłane skrzynki `support@mesco.com.pl`

Kroki: wyciąga numer z tematu → znajduje zgłoszenie → ustawia
`StatusTicket = Oczekuje na odpowiedź`.

Istnieje analogiczny przepływ dla maili **przychodzących** od klienta,
ustawiający `W toku`.

**WYŁĄCZONY przez Microsoft 29.07.2026** — `AllActionsFailingDetected`,
„failed consistently for 14 days".

Prawdopodobne przyczyny:
- trigger reaguje na **każdy** mail w Elementach wysłanych; gdy w temacie nie ma
  `MESco-ID-`, `split(subject, 'MESco-ID-')[1]` wywala błąd
  → naprawa: warunek na triggerze
  `@contains(triggerOutputs()?['body/subject'], 'MESco-ID-')`
- filtr porównuje liczbę z tekstem: `TicketID eq '10001'` (apostrofy)
  → naprawa: `TicketID eq @{outputs('Compose')}`
- brak zabezpieczenia gdy `Get_items` nic nie znajdzie (`first()` na pustym zbiorze)

Status: system jest na wdrożeniu, to środowisko testowe — **nie priorytet**.
Priorytet to dodatek bez ciasteczek.

### 8.3 Plan zmiany triggera dla dodatku

Przepływ z triggerem PowerAppV2 **nie da się wywołać z dodatku do Outlooka** —
to inny mechanizm, nie kwestia uprawnień.

Wybrane rozwiązanie: **trigger `when an item is created`** w liście `Support`.
Dodatek tworzy wiersz przez Graph, przepływ startuje sam i odczytuje dane z rekordu.

Nowe kolumny w `Support`:

| Parametr dziś | Po zmianie |
|---|---|
| `text` (recordId) | ID elementu z triggera (automatycznie) |
| `text_1` (mailSubject) | nowa kolumna `MailSubject` |
| `text_2` (folder) | istniejąca `Creator` |
| `text_3` (mailItemId) | nowa kolumna `MailID` |
| `text_4` (mail klienta) | nowa kolumna `ClientEmail` |
| `text_5` (mail opiekuna) | nowa kolumna `CreatorEmail` |

Koszt: opóźnienie 1–3 min (trigger sonduje listę). W interfejsie pokazać
„numer nadawany…" w miejscu ID.

**Identyfikator wiadomości:** przepływ robi dziś `replace(text_3, '/', '-')` —
obejście różnicy między formatem EWS (`item.itemId` z Office.js) i REST (Graph).
W dodatku zrobić poprawnie: `Office.context.mailbox.convertToRestId()`
i zapisać gotowy identyfikator REST do `MailID`. Wtedy `replace` z przepływu wypada.

**Skrzynka `support@mesco.com.pl` jest współdzielona** — istotne dla `MailID`
i uprawnień do wiadomości.

---

## 9. Plan etapów

- [x] Szkielet dodatku działa na macOS, sideload, pinowanie
- [x] Odczyt otwartej wiadomości (temat, nadawca, odbiorcy)
- [x] Szkielet MSAL + Graph, tryb mock
- [x] Logowanie przez Office Dialog API — działa, token wraca do panelu
- [x] Odczyt prawdziwych danych z SharePointa przez Graph
- [x] Pełne schematy list Support / Companies / Contacts
- [ ] **Etap 1: lista zgłoszeń** — karty, zakładki, liczniki, kaskada zakresu, wyszukiwanie
- [ ] **Etap 2: szczegóły + edycja** — `Status` i `Uwagi` (pierwszy zapis, potrzebne `Sites.ReadWrite.All`)
- [ ] **Etap 3: nowe zgłoszenie** — po zmianie triggera przepływu
- [ ] **Etap 4: galeria firm i galeria kontaktów** — nieomówione, do zrobienia później

Do omówienia w przyszłości: ekrany `HomeCRM` i `CompaniesAll`.
Tymczasowo mogą otwierać Power Apps w przeglądarce.

---

## 10. Otwarte kwestie

| # | Kwestia | Status |
|---|---|---|
| 1 | Wewnętrzne nazwy kolumn | **rozwiązane 15.08** — pełne schematy w sekcji 5 |
| 2 | Zgoda admina na Graph | **rozwiązane 15.08** |
| 3 | Dwie kolumny domeny w Companies: `Domen_x0040_` czy `DomenA`? | **blokuje kaskadę** |
| 4 | Czy komentarz z Power Apps to kolumna `Opis`? | do potwierdzenia |
| 5 | Kto może zamykać zgłoszenie — tylko opiekun czy każdy? | do decyzji |
| 6 | Które z kolumn `Zlecil`, `Cel`, `Data zakończenia`, `AnsysGPT`, `AnsysHelp`, `SupportDay`, `PreSales` pokazać na karcie | do decyzji |
| 7 | Lista pomocnicza `SupportOwners` | zaakceptowane, do zrobienia |
| 8 | Monitoring przepływów | odłożone, środowisko testowe |
| 9 | Założenie „jedna domena na firmę" | świadomy dług |

## 11. Struktura projektu

| Plik | Odpowiedzialność |
|---|---|
| `manifest.xml` | rejestracja dodatku: nazwa, adres taskpane'a, przycisk, uprawnienia do skrzynki. Zero logiki |
| `src/taskpane/taskpane.html` | struktura panelu; kluczowe `id`: `sideload-msg`, `app-body`, `item-subject`, `run` |
| `src/taskpane/taskpane.css` | wygląd, klasy `crm-*` |
| `src/taskpane/taskpane.ts` | tylko `import "./outlook"` |
| `src/taskpane/outlook.ts` | dyrygent: `Office.onReady`, `ItemChanged`, odczyt wiadomości, renderowanie |
| `src/taskpane/authConfig.ts` | stałe: client ID, tenant, adresy SharePointa, nazwy list |
| `src/taskpane/graphService.ts` | token (MSAL, cicho → popup), zapytania do Graph, `USE_MOCK` |
| `src/taskpane/crm.ts` | logika biznesowa: mapowanie kolumn, kaskada zakresu, dopasowanie po domenie |
| `commands.html` / `commands.ts` | nieużywane, ale manifest się do nich odwołuje |

Zasada: manifest → kopiowanie + restart Outlooka. Pozostałe pliki → wystarczy zapis.
