/**
 * Quad AS — Investeringsoversikt
 * Render-webservice: leser Excel-fila fra Dropbox + innlogging + dashboard.
 *
 * Token og hemmeligheter ligger som Environment Variables i Render —
 * ALDRI i koden og aldri eksponert mot nettleseren.
 *
 * Påkrevde env-variabler (sett i Render → Environment):
 *   DROPBOX_APP_KEY        App key fra Dropbox-appen
 *   DROPBOX_APP_SECRET     App secret fra Dropbox-appen
 *   DROPBOX_REFRESH_TOKEN  Refresh token (varer evig; fornyer tilgang selv)
 *   DROPBOX_FILE_PATH      Sti til .xlsx i Dropbox, f.eks. "/Quad/Formue.xlsx"
 *   AUTH_USER              brukernavn for innlogging
 *   AUTH_PASS              passord for innlogging
 *   SESSION_SECRET         lang tilfeldig streng for signering av sesjon
 */

const express = require("express");
const session = require("express-session");
const path = require("path");
const XLSX = require("xlsx");

const app = express();
const PORT = process.env.PORT || 3000;

// ---- Konfig fra miljøet ----
const {
  DROPBOX_APP_KEY,
  DROPBOX_APP_SECRET,
  DROPBOX_REFRESH_TOKEN,
  DROPBOX_FILE_PATH,
  AUTH_USER,
  AUTH_PASS,
  SESSION_SECRET,
} = process.env;

// Navngitte tabeller på WebFeed-arket (laget av Claude for Excel)
const TABLES = ["tbl_kpi", "tbl_aktiva", "tbl_eiendom", "tbl_aksjer", "tbl_historikk"];

app.set("trust proxy", 1); // Render sitter bak en proxy — kreves for secure cookies
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

app.use(
  session({
    secret: SESSION_SECRET || "bytt-meg-i-render",
    resave: false,
    saveUninitialized: false,
    cookie: {
      httpOnly: true,
      secure: true,        // kun over HTTPS (Render kjører HTTPS)
      sameSite: "lax",
      maxAge: 1000 * 60 * 60 * 8, // 8 timer
    },
  })
);

// ---------- Innlogging ----------
function krevInnlogging(req, res, next) {
  if (req.session && req.session.innlogget) return next();
  if (req.accepts("html")) return res.redirect("/login");
  return res.status(401).json({ feil: "Ikke innlogget" });
}

app.get("/login", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "login.html"));
});

app.post("/login", (req, res) => {
  const { brukernavn, passord } = req.body;
  if (brukernavn === AUTH_USER && passord === AUTH_PASS) {
    req.session.innlogget = true;
    return res.redirect("/");
  }
  res.redirect("/login?feil=1");
});

app.post("/logout", (req, res) => {
  req.session.destroy(() => res.redirect("/login"));
});

// ---------- Dropbox (refresh token -> kort tilgangstoken) ----------
let _token = { value: null, exp: 0 };

async function hentDropboxToken() {
  const naa = Date.now();
  if (_token.value && naa < _token.exp - 60_000) return _token.value;

  // Bytt refresh token mot et ferskt, kortlevd tilgangstoken
  const basic = Buffer.from(`${DROPBOX_APP_KEY}:${DROPBOX_APP_SECRET}`).toString("base64");
  const body = new URLSearchParams({
    grant_type: "refresh_token",
    refresh_token: DROPBOX_REFRESH_TOKEN,
  });
  const r = await fetch("https://api.dropboxapi.com/oauth2/token", {
    method: "POST",
    headers: {
      Authorization: `Basic ${basic}`,
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body,
  });
  if (!r.ok) {
    const t = await r.text();
    throw new Error(`Dropbox token-feil ${r.status}: ${t}`);
  }
  const j = await r.json();
  _token = { value: j.access_token, exp: naa + (j.expires_in || 14400) * 1000 };
  return _token.value;
}

// Last ned selve .xlsx-fila fra Dropbox som en Buffer
async function lastNedFil(token) {
  const r = await fetch("https://content.dropboxapi.com/2/files/download", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Dropbox-API-Arg": JSON.stringify({ path: DROPBOX_FILE_PATH }),
    },
  });
  if (!r.ok) {
    const t = await r.text();
    throw new Error(`Dropbox nedlasting-feil ${r.status}: ${t}`);
  }
  const ab = await r.arrayBuffer();
  return Buffer.from(ab);
}

// Les en navngitt tabell fra arbeidsboken (via SheetJS).
// I denne fila ligger ikke tabellene som "Defined Names", men som
// seksjoner på WebFeed-arket markert med rader som "1) NØKKELTALL  → tbl_kpi".
// Vi parser arket én gang og deler det i seksjoner ut fra disse markørene.
function parseWebFeed(workbook) {
  const ws = workbook.Sheets["WebFeed"];
  if (!ws) return {};
  // Les bare kolonne A–H (kolonne I inneholder forklaringstekst, ikke data)
  const alle = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false, defval: "" });
  const rows = alle.map((r) => r.slice(0, 8)); // A–H

  const seksjoner = {};
  let aktiv = null;     // navnet på tabellen vi fyller nå (f.eks. "tbl_kpi")
  let ventHeader = false; // neste rad er kolonneoverskrift
  let header = null;
  let data = [];

  const lagre = () => {
    if (aktiv && header) {
      seksjoner[aktiv] = [header, ...data];
    }
    header = null; data = [];
  };

  for (const rad of rows) {
    const forste = String(rad[0] || "");
    // Markørrad? f.eks. "1) NØKKELTALL  → tbl_kpi"
    const m = forste.match(/→\s*(tbl_[a-zæøå_]+)/i);
    if (m) {
      lagre();                 // lagre forrige seksjon
      aktiv = m[1].toLowerCase();
      ventHeader = true;
      continue;
    }
    if (!aktiv) continue;      // før første markør (intro-tekst) hopper vi over

    // Tom rad avslutter en seksjons data
    const heltTom = rad.every((c) => c === "" || c === null || c === undefined);
    if (heltTom) { continue; }

    if (ventHeader) {
      header = rad.map((c) => String(c).trim());
      ventHeader = false;
      continue;
    }
    data.push(rad);
  }
  lagre(); // siste seksjon
  return seksjoner;
}

// Gjør [header, ...rader] om til liste av objekter (kun ikke-tomme kolonner i header)
function tilObjekter(values) {
  if (!values || !values.length) return [];
  const [header, ...rader] = values;
  // Indekser der header faktisk har et navn
  const kol = header
    .map((h, i) => ({ navn: String(h).trim(), i }))
    .filter((x) => x.navn !== "");
  return rader
    .filter((rad) => kol.some(({ i }) => rad[i] !== "" && rad[i] !== null && rad[i] !== undefined))
    .map((rad) => Object.fromEntries(kol.map(({ navn, i }) => [navn, rad[i]])));
}

// ---------- Data-endepunkt ----------
let _cache = { data: null, exp: 0 };

app.get("/data.json", krevInnlogging, async (req, res) => {
  try {
    const naa = Date.now();
    const ferskt = req.query.fersk === "1";
    if (!ferskt && _cache.data && naa < _cache.exp) {
      return res.json(_cache.data);
    }

    const token = await hentDropboxToken();
    const buf = await lastNedFil(token);
    const workbook = XLSX.read(buf, { type: "buffer" });

    const seksjoner = parseWebFeed(workbook);
    const resultater = {};
    for (const navn of TABLES) {
      resultater[navn] = tilObjekter(seksjoner[navn] || []);
    }

    const payload = {
      oppdatert: new Date().toISOString(),
      valuta: "NOK",
      kpi: resultater.tbl_kpi,
      aktiva: resultater.tbl_aktiva,
      eiendom: resultater.tbl_eiendom,
      aksjer: resultater.tbl_aksjer,
      historikk: resultater.tbl_historikk,
    };

    _cache = { data: payload, exp: naa + 5 * 60 * 1000 }; // 5 min cache
    res.json(payload);
  } catch (e) {
    console.error(e);
    res.status(502).json({ feil: "Kunne ikke hente data fra Dropbox", detalj: String(e.message) });
  }
});

// ---------- Dashboard (bak innlogging) ----------
app.get("/", krevInnlogging, (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

// Helsesjekk for Render
app.get("/healthz", (req, res) => res.send("ok"));

app.listen(PORT, () => console.log(`Quad invest kjører på :${PORT}`));
