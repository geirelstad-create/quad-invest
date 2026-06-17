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
// Finner tabellen på tvers av ark ut fra dens definerte område (!ref).
function lesTabell(workbook, navn) {
  // Excel-tabeller havner i workbook.Workbook.Names som "navn" -> "Ark!$A$1:$F$9"
  const names = (workbook.Workbook && workbook.Workbook.Names) || [];
  let omr = names.find((n) => (n.Name || "").toLowerCase() === navn.toLowerCase());
  let arkNavn, ref;

  if (omr && omr.Ref) {
    // Ref kan være "WebFeed!$A$1:$F$9" eller med apostrof "'Ark 1'!$A$1:..."
    const m = String(omr.Ref).match(/^'?([^'!]+)'?!(.+)$/);
    if (m) { arkNavn = m[1]; ref = m[2].replace(/\$/g, ""); }
  }

  // Fallback: hvis tabellnavn ikke ligger i Names, prøv et ark som heter "WebFeed"
  let ws;
  if (arkNavn && workbook.Sheets[arkNavn]) {
    ws = workbook.Sheets[arkNavn];
  } else {
    ws = workbook.Sheets["WebFeed"] || workbook.Sheets[workbook.SheetNames[0]];
  }
  if (!ws) return [];

  const opts = { header: 1, blankrows: false, defval: "" };
  if (ref) opts.range = ref;
  const rows = XLSX.utils.sheet_to_json(ws, opts);
  return rows; // rad 0 = header
}

// Gjør [[header...],[rad...]] om til liste av objekter
function tilObjekter(values) {
  if (!values || !values.length) return [];
  const [header, ...rader] = values;
  return rader
    .filter((rad) => rad.some((c) => c !== "" && c !== null && c !== undefined))
    .map((rad) => Object.fromEntries(header.map((h, i) => [String(h), rad[i]])));
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

    const resultater = {};
    for (const navn of TABLES) {
      resultater[navn] = tilObjekter(lesTabell(workbook, navn));
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
