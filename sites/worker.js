/**
 * OpdrachtZeker — Worker voor de gratis beta.
 *
 * Routes:
 *   POST /api/wachtlijst   inschrijving valideren en centraal opslaan (D1)
 *   GET  /api/wachtlijst   export van alle inschrijvingen, alleen met admin-token
 *   al het overige         statische bestanden (landingspagina, risicoscan)
 *
 * De statische bestanden komen uit sites/public/ (zie scripts/bouw_site.py).
 * Nooit vertrouwen op client-validatie: alles wordt hier opnieuw gecontroleerd.
 */

const MAX_BODY = 4000;

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "no-store",
      "X-Content-Type-Options": "nosniff",
    },
  });
}

/** Vergelijkt twee strings zonder vroegtijdig af te breken (tegen timing-lekken). */
function veiligGelijk(a, b) {
  if (typeof a !== "string" || typeof b !== "string") return false;
  if (a.length !== b.length) return false;
  let verschil = 0;
  for (let i = 0; i < a.length; i++) verschil |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return verschil === 0;
}

/** Mutaties alleen vanaf de eigen site accepteren. */
function zelfdeOorsprong(request) {
  const origin = request.headers.get("Origin");
  if (!origin) return true; // geen browser-context (curl, server-side client)
  try {
    return new URL(origin).host === new URL(request.url).host;
  } catch {
    return false;
  }
}

function geldigEmail(v) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v);
}

// Schema eenmalig per isolate aanmaken: gecontroleerde migratie in plaats van losse SQL.
let schemaKlaar = false;
async function zorgVoorSchema(db) {
  if (schemaKlaar) return;
  await db
    .prepare(
      `CREATE TABLE IF NOT EXISTS wachtlijst (
         id INTEGER PRIMARY KEY AUTOINCREMENT,
         naam TEXT NOT NULL,
         email TEXT NOT NULL UNIQUE,
         bedrijf TEXT NOT NULL,
         zzp TEXT NOT NULL,
         bron TEXT,
         tijdstip TEXT NOT NULL
       )`
    )
    .run();
  schemaKlaar = true;
}

async function aanmelden(request, env) {
  if (!env.DB) return json({ fout: "Database niet beschikbaar" }, 503);
  if (!zelfdeOorsprong(request)) return json({ fout: "Ongeldige oorsprong" }, 403);

  const ruw = await request.text();
  if (ruw.length > MAX_BODY) return json({ fout: "Verzoek te groot" }, 413);

  let invoer;
  try {
    invoer = JSON.parse(ruw);
  } catch {
    return json({ fout: "Ongeldige JSON" }, 400);
  }

  const naam = String(invoer?.naam ?? "").trim().slice(0, 120);
  const email = String(invoer?.email ?? "").trim().toLowerCase().slice(0, 200);
  const bedrijf = String(invoer?.bedrijf ?? "").trim().slice(0, 160);
  const zzp = String(invoer?.zzp ?? "").trim().slice(0, 60);
  const bron = String(invoer?.bron ?? "onbekend").trim().slice(0, 60);

  const velden = {};
  if (naam.length < 2) velden.naam = "Vul je naam in.";
  if (!geldigEmail(email)) velden.email = "Vul een geldig e-mailadres in.";
  if (bedrijf.length < 2) velden.bedrijf = "Vul je bedrijfsnaam in.";
  if (!zzp) velden.zzp = "Geef aan met hoeveel ZZP'ers je werkt.";
  if (Object.keys(velden).length) return json({ fout: "Validatie mislukt", velden }, 400);

  await zorgVoorSchema(env.DB);

  const res = await env.DB.prepare(
    `INSERT INTO wachtlijst (naam, email, bedrijf, zzp, bron, tijdstip)
     VALUES (?, ?, ?, ?, ?, ?)
     ON CONFLICT(email) DO NOTHING`
  )
    .bind(naam, email, bedrijf, zzp, bron, new Date().toISOString())
    .run();

  // Positie op de lijst = aantal inschrijvingen tot en met deze rij (stabiel bij herinschrijving).
  const rij = await env.DB.prepare(
    `SELECT COUNT(*) AS n FROM wachtlijst
     WHERE id <= (SELECT id FROM wachtlijst WHERE email = ?)`
  )
    .bind(email)
    .first();

  return json(
    {
      ok: true,
      plek: rij?.n ?? null,
      alBekend: res.meta?.changes === 0,
    },
    201
  );
}

async function exporteer(request, env) {
  if (!env.DB) return json({ fout: "Database niet beschikbaar" }, 503);

  const verwacht = env.ADMIN_TOKEN;
  if (!verwacht) return json({ fout: "Export is niet geconfigureerd (ADMIN_TOKEN ontbreekt)." }, 503);

  const url = new URL(request.url);
  const meegegeven =
    (request.headers.get("Authorization") || "").replace(/^Bearer\s+/i, "") ||
    url.searchParams.get("token") ||
    "";
  if (!veiligGelijk(meegegeven, verwacht)) {
    return json({ fout: "Niet geautoriseerd" }, 401, { "WWW-Authenticate": "Bearer" });
  }

  await zorgVoorSchema(env.DB);
  const { results } = await env.DB.prepare(
    `SELECT id, naam, email, bedrijf, zzp, bron, tijdstip FROM wachtlijst ORDER BY id ASC`
  ).all();
  const rijen = results ?? [];

  if (url.searchParams.get("formaat") === "csv") {
    const kop = ["id", "naam", "email", "bedrijf", "zzp", "bron", "tijdstip"];
    const cel = (v) => `"${String(v ?? "").replace(/"/g, '""')}"`;
    const csv = [kop.join(","), ...rijen.map((r) => kop.map((k) => cel(r[k])).join(","))].join("\n");
    return new Response(csv, {
      headers: {
        "Content-Type": "text/csv; charset=utf-8",
        "Content-Disposition": 'attachment; filename="wachtlijst.csv"',
        "Cache-Control": "no-store",
      },
    });
  }

  return json({ aantal: rijen.length, inschrijvingen: rijen });
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (url.pathname === "/api/wachtlijst" || url.pathname === "/api/wachtlijst/") {
      if (request.method === "POST") return aanmelden(request, env);
      if (request.method === "GET") return exporteer(request, env);
      return json({ fout: "Methode niet toegestaan" }, 405);
    }

    if (url.pathname.startsWith("/api/")) return json({ fout: "Onbekend eindpunt" }, 404);

    return env.ASSETS.fetch(request);
  },
};
