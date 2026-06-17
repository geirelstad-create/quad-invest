# Quad AS — Investeringsoversikt på invest.quad.no (Dropbox-versjon)

Render-tjeneste som leser Excel-arbeidsboken **fra Dropbox**, henter de fem WebFeed-tabellene (`tbl_kpi`, `tbl_aktiva`, `tbl_eiendom`, `tbl_aksjer`, `tbl_historikk`) og viser dem på et innlogget dashboard.

```
Excel-fil i Dropbox
        │  (Dropbox API: last ned .xlsx, les tabellene med SheetJS)
        ▼
Render-webservice  ──  Dropbox-nøkler ligger trygt som miljøvariabler
        │  (JSON, kun for innlogget bruker)
        ▼
invest.quad.no  ──  innlogging med brukernavn + passord
```

Du oppdaterer bare detalj-fanene i Excel som før (WebFeed-tabellene følger med), lagrer fila i Dropbox, og dashboardet henter ferske tall (innenfor 5-min cache, eller med «Oppdater»-knappen).

---

## Del 1 — Opprett en Dropbox-app (App folder)

1. Gå til **dropbox.com/developers/apps** → **Create app**.
   - **Choose an API:** Scoped access
   - **Type of access:** **App folder** (appen ser kun sin egen mappe — tryggest)
   - **Navn:** `quad-invest` (dette navnet blir også navnet på mappa Dropbox lager)
2. På appens side → fanen **Permissions** → huk av:
   - `files.metadata.read`
   - `files.content.read`
   - Klikk **Submit** nederst for å lagre tillatelsene.
3. Fanen **Settings** → noter:
   - **App key** → blir `DROPBOX_APP_KEY`
   - **App secret** → blir `DROPBOX_APP_SECRET`

Når appen er opprettet, lager Dropbox automatisk en mappe for den:
`Dropbox/Apps/quad-invest/`. **Legg en kopi av Excel-arbeidsboken der inne** (f.eks. `Dropbox/Apps/quad-invest/Formue.xlsx`). Det er denne kopien tjenesten leser — du kan oppdatere den ved å lagre/synke arbeidsboken hit.

## Del 2 — Hent et refresh token (varer evig)

Et refresh token lar tjenesten hente ferske tilgangstokener selv, uten at du må fornye noe manuelt. Dette er en engangsjobb. Den enkleste veien er det vedlagte hjelpescriptet:

1. Kjør (krever Node 18+, ingen pakker trengs):
   ```
   node hent-refresh-token.js DIN_APP_KEY DIN_APP_SECRET
   ```
2. Scriptet skriver ut en URL. Åpne den, logg inn og godkjenn appen.
3. Kopier «authorization code» du får, lim den inn i terminalen og trykk Enter.
4. Scriptet skriver ut `DROPBOX_REFRESH_TOKEN=...` — den verdien bruker du i Render.

> Authorization-koden kan kun brukes én gang og må brukes raskt, så lim den inn med en gang.

**Uten scriptet (manuelt):** åpne
`https://www.dropbox.com/oauth2/authorize?client_id=DIN_APP_KEY&response_type=code&token_access_type=offline`,
godkjenn, og bytt koden mot et refresh token med:
```
curl https://api.dropboxapi.com/oauth2/token \
  -d code=AUTH_CODE \
  -d grant_type=authorization_code \
  -u DIN_APP_KEY:DIN_APP_SECRET
```
Svaret inneholder `"refresh_token": "..."`.

## Del 3 — Stien til Excel-fila (viktig med App folder)

`DROPBOX_FILE_PATH` regnes **fra appmappa**, ikke fra Dropbox-roten. Dette er den vanligste feilen, så merk deg:

- La du fila rett i appmappa (`Dropbox/Apps/quad-invest/Formue.xlsx`)
  → stien er **`/Formue.xlsx`** (IKKE `/Apps/quad-invest/Formue.xlsx`).
- La du den i en undermappe (`Dropbox/Apps/quad-invest/data/Formue.xlsx`)
  → stien er `/data/Formue.xlsx`.

Pass på eksakt skrivemåte (store/små bokstaver, æøå). Stien skal alltid begynne med `/`.

## Del 4 — Deploy til Render

1. Legg disse filene i et GitHub-repo: `server.js`, `package.json`, `public/index.html`, `public/login.html`.
2. På **render.com** → **New** → **Web Service** → velg repoet.
   - **Build command:** `npm install`
   - **Start command:** `npm start`
3. Under **Environment**, legg inn:

   | Variabel | Verdi |
   |---|---|
   | `DROPBOX_APP_KEY` | fra Dropbox (Del 1) |
   | `DROPBOX_APP_SECRET` | fra Dropbox (Del 1) |
   | `DROPBOX_REFRESH_TOKEN` | fra Del 2 |
   | `DROPBOX_FILE_PATH` | stien fra Del 3, f.eks. `/Quad/Formue.xlsx` |
   | `AUTH_USER` | brukernavn for innlogging |
   | `AUTH_PASS` | et sterkt passord |
   | `SESSION_SECRET` | lang tilfeldig streng |

4. **Deploy.** Test på Render-URL-en (f.eks. `quad-invest.onrender.com`) → logg inn → tallene skal vises.

> Render Starter kan «sovne» ved inaktivitet, så første lasting kan ta noen sekunder.

## Del 5 — Koble invest.quad.no

Samme mønster som du gjorde med www:

1. I **Render** → tjenesten → Settings → Custom Domains → legg til `invest.quad.no`. Render gir en CNAME-verdi.
2. Hos **Domeneshop** → fanen **DNS-pekere** → ny rad:
   ```
   Navn:   invest
   Peker:  quad-invest.onrender.com   (verdien Render oppgir)
   ```
3. Render utsteder HTTPS-sertifikat automatisk når DNS har spredd seg. La MX og de andre postene stå.

---

## Sikkerhet

- Dropbox-nøklene og refresh token ligger kun som miljøvariabler i Render, aldri i koden eller mot nettleseren.
- Tallene er sensitive — hovedvernet er innloggingen.
- `App folder`-tilgang er tryggest: appen kan kun lese sin egen mappe i Dropbox, ikke resten.
- Vil du senere ha ekte innlogging (flere brukere) i stedet for ett felles brukernavn/passord, kan det utvides.

## Daglig bruk

1. Oppdater detalj-fanene i Excel som vanlig. WebFeed-tabellene følger med automatisk.
2. Lagre fila i Dropbox (samme sti).
3. Åpne `invest.quad.no`, logg inn. Trykk **Oppdater** for ferske tall (ellers 5-min cache).
