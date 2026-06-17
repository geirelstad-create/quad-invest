#!/usr/bin/env node
/**
 * Engangs-hjelper for å hente et Dropbox refresh token.
 *
 * Bruk:
 *   1. Kjør:  node hent-refresh-token.js DIN_APP_KEY DIN_APP_SECRET
 *   2. Åpne URL-en scriptet skriver ut, logg inn og godkjenn appen.
 *   3. Lim inn koden du får tilbake, trykk Enter.
 *   4. Scriptet skriver ut ditt DROPBOX_REFRESH_TOKEN.
 *
 * Krever ingen pakker — bruker innebygd fetch (Node 18+) og readline.
 */

const readline = require("readline");

const [appKey, appSecret] = process.argv.slice(2);

if (!appKey || !appSecret) {
  console.error("\nBruk: node hent-refresh-token.js DIN_APP_KEY DIN_APP_SECRET\n");
  process.exit(1);
}

const authUrl =
  `https://www.dropbox.com/oauth2/authorize` +
  `?client_id=${encodeURIComponent(appKey)}` +
  `&response_type=code&token_access_type=offline`;

console.log("\n1) Åpne denne URL-en i nettleseren, logg inn og godkjenn appen:\n");
console.log("   " + authUrl + "\n");
console.log("2) Kopier 'authorization code' du får, og lim den inn her.\n");

const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
rl.question("Authorization code: ", async (code) => {
  code = code.trim();
  if (!code) { console.error("Ingen kode oppgitt."); process.exit(1); }

  const basic = Buffer.from(`${appKey}:${appSecret}`).toString("base64");
  const body = new URLSearchParams({
    code,
    grant_type: "authorization_code",
  });

  try {
    const r = await fetch("https://api.dropboxapi.com/oauth2/token", {
      method: "POST",
      headers: {
        Authorization: `Basic ${basic}`,
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body,
    });
    const j = await r.json();
    if (!r.ok) {
      console.error("\nFeil fra Dropbox:", JSON.stringify(j, null, 2));
      process.exit(1);
    }
    console.log("\n✅ Ferdig! Sett denne i Render:\n");
    console.log("   DROPBOX_REFRESH_TOKEN=" + j.refresh_token + "\n");
    if (j.account_id) console.log("   (konto: " + j.account_id + ")\n");
  } catch (e) {
    console.error("\nNoe gikk galt:", e.message);
    console.error("Merk: authorization-koden kan kun brukes én gang og må brukes raskt.");
    process.exit(1);
  } finally {
    rl.close();
  }
});
