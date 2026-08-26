/**
 * Extraction serveur de la campagne intégrée RR-POLIO depuis le DHIS2 de
 * campagne (rdccampagne.hispwca.org, dataset « PEV_Campagne … RR et Polio »).
 *
 * L'API analytics est utilisée au niveau Aire de Santé (LEVEL-4) : vérifié le
 * 26/08/2026, elle restitue exactement les mêmes totaux que les valeurs brutes
 * `/api/dataValueSets` (les saisies FOSA remontent à l'AS parente), pour un
 * volume ~20× moindre (2 requêtes par province au lieu de 2 par ZS).
 *
 * Conventions alignées sur le tableau de bord Power BI officiel (Bloc 3) :
 *  - cible RR = « RR_Cible attendue RR » telle que saisie (46 % pop) ;
 *  - cible polio = cible RR × (18,9/46) — cible 0-59 mois — uniquement pour
 *    les AS qui participent au volet polio (« RR-POLIO ») ; les provinces /
 *    antennes en campagne RR seule gardent une cible polio nulle et leurs
 *    indicateurs polio s'affichent « — » ;
 *  - population totale = cible RR ÷ 0,46 ;
 *  - J1 est propre à chaque province (lancements décalés : 11/08 pour la
 *    plupart, 20/08 au Nord Kivu…) : détecté comme le premier jour dont le
 *    volume vacciné atteint 5 % du meilleur jour ; J1+5 et suivants = Ratissage.
 */

import https from "node:https";
import { ANTIGENES, sanitizeRecords, type ASRecord, type VaccineStats } from "./parse-masque";
import { antenneForZS, normZS } from "./antennes";
import { DHIS2_BLOCK_SCHEMA, JOUR_LABELS, NB_JOURS, type ProvinceBlock, type ProvinceInfo } from "./dhis2-shared";
import { kvGetJSON, kvSetJSON, readNationalBlocks } from "./kv-store";

/* ─── Connexion ────────────────────────────────────────────────────────── */

const DHIS_HOST = "rdccampagne.hispwca.org";
/** Repli IP directe : certains résolveurs DNS ne connaissent pas le domaine. */
const DHIS_IP = "161.97.129.236";
const DHIS_AUTH = "Basic " + Buffer.from("RR_Polio:Snis@2026").toString("base64");
const TIMEOUT_MS = 120_000;

function dhisOnce(pathname: string, host: string): Promise<unknown> {
  return new Promise((resolve, reject) => {
    const req = https.request(
      {
        host,
        servername: DHIS_HOST,
        path: "/dhis/api" + pathname,
        method: "GET",
        headers: { Host: DHIS_HOST, Authorization: DHIS_AUTH, Accept: "application/json" },
        timeout: TIMEOUT_MS,
      },
      (res) => {
        let d = "";
        res.on("data", (c) => (d += c));
        res.on("end", () => {
          if ((res.statusCode ?? 0) >= 400) return reject(new Error(`DHIS2 ${res.statusCode}: ${d.slice(0, 300)}`));
          try { resolve(JSON.parse(d)); } catch (e) { reject(e); }
        });
      }
    );
    req.on("error", reject);
    req.on("timeout", () => req.destroy(new Error("DHIS2 timeout")));
    req.end();
  });
}

async function dhis(pathname: string): Promise<unknown> {
  try {
    return await dhisOnce(pathname, DHIS_HOST);
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    if (/ENOTFOUND|EAI_AGAIN|ECONNREFUSED|ETIMEDOUT|timeout/i.test(msg)) {
      return dhisOnce(pathname, DHIS_IP);
    }
    throw e;
  }
}

/* ─── Éléments de données / combos (identiques au rapport Lubutu) ───────── */

const PART_RR = 0.46; // cible RR = 6 mois - 14 ans
const PART_POLIO = 0.189; // cible polio = 0 - 59 mois
const START_DATE = "2026-08-01";
const DEFAULT_J1 = "2026-08-11";

const COC = {
  zeroM: "wx9PDftv9VU",
  zeroF: "WCHQizINNMq",
  plusM: "dNY0P6Nezy0",
  plusF: "eJdLLsSvmFA",
  recue: "Iez1cqvzyKD",
  utilisee: "Iy27KpBZ9aF",
  rendue: "ZdlGwU1SLuo",
  perdue: "z9e4HQlBLra",
  ident: "ZItCz2m93Hb",
  recup: "xEoj2QGn4hw",
  notifiees: "Wm3bOvTQPhi",
} as const;

const AGE_DE = {
  nvpo2: ["hcjjuPEmau2", "trxWpua3GB1", "KhMa1ehdwvY", null],
  vpob: ["bWZF88KmH7E", "bVoRjDT5DUN", "NJ7bAfA9lUY", null],
  rr: ["LUAiDML2q1j", "iRJtRKAEOPQ", "ztjsflVdf1c"],
} as const;
const ALL_AGE_DE: string[] = ([...AGE_DE.nvpo2, ...AGE_DE.vpob, ...AGE_DE.rr] as (string | null)[]).filter(
  (x): x is string => x != null
);
const POLIO_AGE_DE = new Set(
  ([...AGE_DE.nvpo2, ...AGE_DE.vpob] as (string | null)[]).filter((x): x is string => x != null)
);

const FLACON_DE = { nvpo2: "nYq937kfajz", vpob: "V38AKsoLCvy", rr: "Xgw4c99kK0p" } as const;
const CIBLE_RR_DE = "BJ6CjSbAiLJ";
const CIBLE_POLIO_DE = "GnRVkCsFDZE";
const SURV_DE = { pfa: "ZexanWv3fFM", rougeole: "dYhGRWbkXbT", fj: "n6JhuMhdGbf", tnn: "jMJ7abE4cKb" } as const;
const MAPI_DE = { graves: "YG0JCwIikUU", nonGraves: "gF1so9TfqUX" } as const;
/** Récupération pendant les AVS — ordre du masque : [nVPO2, VPOb, RR]. */
const AVS_DE: string[][] = [
  ["qZ6OXVLAZoH", "VVTa5KoZfek"],
  ["XUKHSDhV4np", "vKW3SByKOz3"],
  ["Tj6Uug4aaQj", "WQOKvgWStMm"],
];
const PEV_DE: Record<string, string> = {
  BCG: "ATZtz3vX9Cj", DTC1: "iFW7yj3ftkK", PCV1: "eGWjLe4JqkZ", ROTA1: "aKsA7JcExZw",
  DTC2: "d9gOS5k17rX", PCV2: "LnynmJNoki9", ROTA2: "Smopcheplu8", DTC3: "ZHaMTWh7MHh",
  PCV3: "l09sbfgjBF8", ROTA3: "ptTjjrd4RBa", VPI1: "sZaMowoYB8o", VAP1: "SYaFLJIuU96",
  VAP2: "yKinSHZrRhP", VAP3: "tPqHdo4GoYa", VPI2: "WDNopj5AcpF", VAA: "dJqSnLSMKik",
  VAP4: "Pfc354bq0cE", TD1: "kNdUxoWhUS5", TD2: "ZzEhrEZu261", TD3: "Vt354qVK08v",
  TD4: "uryaNxO3R1F", TD5: "x3AP46l37ge",
};

/** DE cumulables (les cibles sont traitées à part : déduplication par FOSA/AS). */
const ALL_DE: string[] = [
  ...ALL_AGE_DE,
  FLACON_DE.nvpo2, FLACON_DE.vpob, FLACON_DE.rr,
  SURV_DE.pfa, SURV_DE.rougeole, SURV_DE.fj, SURV_DE.tnn,
  MAPI_DE.graves, MAPI_DE.nonGraves,
  ...AVS_DE.flat(),
  ...Object.values(PEV_DE),
];

/* ─── Helpers ──────────────────────────────────────────────────────────── */

interface AnalyticsResult { headers: { name: string }[]; rows: string[][] }

async function analytics(dims: string[]): Promise<{ rows: string[][]; col: Record<string, number> }> {
  const qs = dims.map((d) => `dimension=${encodeURIComponent(d)}`).join("&");
  const res = (await dhis(`/analytics.json?${qs}&skipMeta=true&skipRounding=true`)) as AnalyticsResult;
  const col: Record<string, number> = {};
  (res.headers ?? []).forEach((h, i) => { col[h.name] = i; });
  return { rows: res.rows ?? [], col };
}

function cleanName(n: string): string {
  return (n || "")
    .replace(/^[a-z]{2}\s+/i, "")
    .replace(/\s+(Aire|Zone) de Sant[eé]\s*$/i, "")
    .replace(/\s+Province\s*$/i, "")
    .trim();
}

function emptyVacc(nbAges: number): VaccineStats {
  return {
    vacc: 0, zeroDose: 0, garcons: 0, filles: 0, ages: new Array(nbAges).fill(0),
    flaconsRecus: 0, flaconsUtil: 0, flaconsRendus: 0, flaconsPerdus: 0,
    daily: new Array(NB_JOURS).fill(0),
  };
}

/** Périodes journalières « yyyymmdd » du 01/08/2026 à aujourd'hui (+1 pour les fuseaux). */
function dailyPeriods(): string[] {
  const out: string[] = [];
  const d = new Date(`${START_DATE}T00:00:00Z`);
  const end = new Date();
  end.setUTCDate(end.getUTCDate() + 1);
  const max = new Date("2026-10-15T00:00:00Z");
  const stop = end < max ? end : max;
  while (d <= stop) {
    out.push(d.toISOString().slice(0, 10).replace(/-/g, ""));
    d.setUTCDate(d.getUTCDate() + 1);
  }
  return out;
}

function isoOf(period: string): string {
  return `${period.slice(0, 4)}-${period.slice(4, 6)}-${period.slice(6, 8)}`;
}

/** Index de jour de campagne (0..5) d'une période, selon le J1 de la province. */
function dayIndex(period: string, j1Period: string): number {
  if (period < j1Period) return -1;
  const diff = Math.round((Date.parse(isoOf(period)) - Date.parse(isoOf(j1Period))) / 86_400_000);
  return diff <= 4 ? diff : 5;
}

/* ─── Liste des provinces de la campagne ───────────────────────────────── */

export async function listCampaignProvinces(): Promise<ProvinceInfo[]> {
  const meta = (await dhis(`/organisationUnits.json?level=2&fields=id,name&paging=false`)) as {
    organisationUnits: { id: string; name: string }[];
  };
  const names = new Map(meta.organisationUnits.map((o) => [o.id, cleanName(o.name)]));

  // Vaccinés : cumul mensuel classique par province.
  const { rows, col } = await analytics([`dx:${ALL_AGE_DE.join(";")}`, "ou:LEVEL-2", "pe:202607;202608;202609"]);
  const acc = new Map<string, { cibleRR: number; ciblePolio: number; rrVacc: number; polioVacc: number }>();
  const get = (ou: string) => {
    let a = acc.get(ou);
    if (!a) { a = { cibleRR: 0, ciblePolio: 0, rrVacc: 0, polioVacc: 0 }; acc.set(ou, a); }
    return a;
  };
  for (const r of rows) {
    const a = get(r[col.ou]);
    const v = parseFloat(r[col.value]) || 0;
    if (POLIO_AGE_DE.has(r[col.dx])) a.polioVacc += v;
    else a.rrVacc += v;
  }

  // Cibles : même déduplication que buildProvinceBlock (max par FOSA, valeurs
  // distinctes par AS) — la somme brute doublait le Sud Kivu / Nord Kivu.
  const qc = await analytics([`dx:${CIBLE_RR_DE};${CIBLE_POLIO_DE}`, "ou:LEVEL-5", `pe:${dailyPeriods().join(";")}`]);
  const fosas = (await dhis(`/organisationUnits.json?level=5&fields=id,path&paging=false`)) as {
    organisationUnits: { id: string; path: string }[];
  };
  const pathOf = new Map(fosas.organisationUnits.map((f) => [f.id, f.path]));
  const fosaMax = new Map<string, { rr: number; polio: number }>();
  for (const r of qc.rows) {
    const fosa = r[qc.col.ou];
    const v = parseFloat(r[qc.col.value]) || 0;
    let m = fosaMax.get(fosa);
    if (!m) { m = { rr: 0, polio: 0 }; fosaMax.set(fosa, m); }
    if (r[qc.col.dx] === CIBLE_RR_DE) m.rr = Math.max(m.rr, v);
    else m.polio = Math.max(m.polio, v);
  }
  const asVals = new Map<string, { prov: string; rr: Set<number>; polio: Set<number> }>();
  for (const [fosa, m] of fosaMax) {
    // path = /pays/province/zs/as/fosa
    const seg = (pathOf.get(fosa) ?? "").split("/").filter(Boolean);
    if (seg.length < 4) continue;
    const prov = seg[1];
    const as = seg[3];
    let a = asVals.get(as);
    if (!a) { a = { prov, rr: new Set(), polio: new Set() }; asVals.set(as, a); }
    if (m.rr > 0) a.rr.add(m.rr);
    if (m.polio > 0) a.polio.add(m.polio);
  }
  for (const a of asVals.values()) {
    const p = get(a.prov);
    for (const v of a.rr) p.cibleRR += v;
    for (const v of a.polio) p.ciblePolio += v;
  }

  const out: ProvinceInfo[] = [];
  for (const [id, a] of acc) {
    if (a.cibleRR <= 0 && a.rrVacc <= 0 && a.polioVacc <= 0) continue;
    // Province RR-POLIO : des vaccinés polio, ou une cible polio saisie de façon
    // substantielle (Kasaï Central : cibles saisies, résultats via le masque).
    // Une cible parasite sur 1-2 AS (Sankuru) ne suffit pas.
    const polio = a.polioVacc > 0 || (a.cibleRR > 0 && a.ciblePolio >= 0.1 * a.cibleRR);
    out.push({ id, name: names.get(id) ?? id, cibleRR: a.cibleRR, rrVacc: a.rrVacc, polioVacc: a.polioVacc, polio });
  }
  return out.sort((a, b) => a.name.localeCompare(b.name, "fr"));
}

/* ─── Extraction d'une province ────────────────────────────────────────── */

export async function buildProvinceBlock(provinceId: string): Promise<ProvinceBlock> {
  // 1. Hiérarchie : province, ZS (niveau 3), AS (niveau 4).
  const ous = (await dhis(
    `/organisationUnits.json?filter=path:like:${provinceId}&filter=level:in:[2,3,4,5]&fields=id,name,level,parent[id]&paging=false`
  )) as { organisationUnits: { id: string; name: string; level: number; parent?: { id: string } }[] };
  const units = ous.organisationUnits ?? [];
  const provName = cleanName(units.find((u) => u.level === 2 && u.id === provinceId)?.name ?? provinceId);
  const zsName = new Map<string, string>();
  for (const u of units) if (u.level === 3) zsName.set(u.id, cleanName(u.name));
  /** FOSA (niveau 5) → Aire de Santé parente. */
  const fosaParent = new Map<string, string>();
  for (const u of units) if (u.level === 5 && u.parent?.id) fosaParent.set(u.id, u.parent.id);

  const records = new Map<string, ASRecord>();
  const polioSeen = new Set<string>();
  for (const u of units) {
    if (u.level !== 4) continue;
    const zs = zsName.get(u.parent?.id ?? "") ?? "—";
    records.set(u.id, {
      province: provName, antenne: antenneForZS(zs, provName), zs, as: cleanName(u.name),
      popTotale: 0, menagesPrevus: 0, menagesVisites: 0, mosoAttendus: 0, mosoRecus: 0,
      pers15: 0, refusSignales: 0, refusGeres: 0,
      vaccAttendus: NB_JOURS, vaccRecus: 0,
      dailyReports: JOUR_LABELS.map((label) => ({ label, attendus: 1, recus: 0 })),
      ciblePolio: 0, cibleRR: 0,
      nvpo2: emptyVacc(4), vpob: emptyVacc(4), rr: emptyVacc(3),
      survPFA: 0, survRougeole: 0, survFJ: 0, survTNN: 0,
      mapiNonGraves: 0, mapiGraves: 0,
      aidantsIdent: [0, 0, 0], aidantsRecup: [0, 0, 0],
      pevIdent: new Array(ANTIGENES.length).fill(0),
      pevRecup: new Array(ANTIGENES.length).fill(0),
    });
  }

  // 2. Cumuls mensuels avec ventilation par combinaison de catégories.
  const q1 = await analytics([`dx:${ALL_DE.join(";")}`, "co", `ou:LEVEL-4;${provinceId}`, "pe:202607;202608;202609"]);
  for (const row of q1.rows) {
    const r = records.get(row[q1.col.ou]);
    if (!r) continue;
    const de = row[q1.col.dx];
    const coc = row[q1.col.co];
    const val = parseFloat(row[q1.col.value]) || 0;
    if (!val) continue;

    let handled = false;
    for (const key of ["nvpo2", "vpob", "rr"] as const) {
      const idx = (AGE_DE[key] as readonly (string | null)[]).indexOf(de);
      if (idx < 0) continue;
      const s = r[key];
      s.vacc += val;
      s.ages[idx] += val;
      if (coc === COC.zeroM || coc === COC.zeroF) s.zeroDose += val;
      if (coc === COC.zeroM || coc === COC.plusM) s.garcons += val;
      if (coc === COC.zeroF || coc === COC.plusF) s.filles += val;
      if (key !== "rr") polioSeen.add(row[q1.col.ou]);
      handled = true;
      break;
    }
    if (handled) continue;

    for (const key of ["nvpo2", "vpob", "rr"] as const) {
      if (de !== FLACON_DE[key]) continue;
      const s = r[key];
      if (coc === COC.recue) s.flaconsRecus += val;
      else if (coc === COC.utilisee) s.flaconsUtil += val;
      else if (coc === COC.rendue) s.flaconsRendus += val;
      else if (coc === COC.perdue) s.flaconsPerdus += val;
      if (key !== "rr") polioSeen.add(row[q1.col.ou]);
      handled = true;
      break;
    }
    if (handled) continue;

    if (de === SURV_DE.pfa) { r.survPFA += val; continue; }
    if (de === SURV_DE.rougeole) { r.survRougeole += val; continue; }
    if (de === SURV_DE.fj) { r.survFJ += val; continue; }
    if (de === SURV_DE.tnn) { r.survTNN += val; continue; }

    // MAPI : seule la ventilation « Notifiées » compte (convention Power BI).
    if (de === MAPI_DE.graves) { if (coc === COC.notifiees) r.mapiGraves += val; continue; }
    if (de === MAPI_DE.nonGraves) { if (coc === COC.notifiees) r.mapiNonGraves += val; continue; }

    const avs = AVS_DE.findIndex((pair) => pair.includes(de));
    if (avs >= 0) {
      if (coc === COC.ident) r.aidantsIdent[avs] += val;
      else if (coc === COC.recup) r.aidantsRecup[avs] += val;
      continue;
    }

    const ag = ANTIGENES.findIndex((a) => PEV_DE[a.key] === de);
    if (ag >= 0) {
      if (coc === COC.ident) r.pevIdent[ag] += val;
      else if (coc === COC.recup) r.pevRecup[ag] += val;
    }
  }

  // 2 bis. Cibles — convention du tableau de bord officiel (vérifiée le 26/08 sur
  //    le Power BI Bloc 3) : la cible est SAISIE PAR FOSA et ne s'additionne pas
  //    aveuglément — au Sud Kivu (231 AS) et au Nord Kivu, plusieurs FOSA d'une
  //    même AS ont chacune saisi la cible de TOUTE l'AS, parfois sur plusieurs
  //    jours (la somme brute donnait 7,35 M au lieu de ~4,37 M au Sud Kivu).
  //    Règle : max par FOSA (tous jours), puis somme des valeurs DISTINCTES par
  //    AS — exacte pour Kasaï Or./Lomami/Maniema/Sankuru, ±1 % pour SK/NK.
  const periods = dailyPeriods();
  {
    const q3 = await analytics([`dx:${CIBLE_RR_DE};${CIBLE_POLIO_DE}`, `ou:LEVEL-5;${provinceId}`, `pe:${periods.join(";")}`]);
    const fosaMax = new Map<string, { rr: number; polio: number }>();
    for (const row of q3.rows) {
      const fosa = row[q3.col.ou];
      const val = parseFloat(row[q3.col.value]) || 0;
      let m = fosaMax.get(fosa);
      if (!m) { m = { rr: 0, polio: 0 }; fosaMax.set(fosa, m); }
      if (row[q3.col.dx] === CIBLE_RR_DE) m.rr = Math.max(m.rr, val);
      else m.polio = Math.max(m.polio, val);
    }
    const asVals = new Map<string, { rr: Set<number>; polio: Set<number> }>();
    for (const [fosa, m] of fosaMax) {
      const as = fosaParent.get(fosa) ?? fosa; // valeur saisie directement à l'AS
      let a = asVals.get(as);
      if (!a) { a = { rr: new Set(), polio: new Set() }; asVals.set(as, a); }
      if (m.rr > 0) a.rr.add(m.rr);
      if (m.polio > 0) a.polio.add(m.polio);
    }
    for (const [as, a] of asVals) {
      const r = records.get(as);
      if (!r) continue;
      for (const v of a.rr) r.cibleRR += v;
      if (a.polio.size > 0) polioSeen.add(as);
    }
  }

  const q2 = await analytics([`dx:${ALL_AGE_DE.join(";")}`, `ou:LEVEL-4;${provinceId}`, `pe:${periods.join(";")}`]);
  const byDay = new Map<string, number>();
  for (const row of q2.rows) {
    const pe = row[q2.col.pe];
    byDay.set(pe, (byDay.get(pe) ?? 0) + (parseFloat(row[q2.col.value]) || 0));
  }
  const maxDay = Math.max(0, ...byDay.values());
  const threshold = Math.max(100, maxDay * 0.05);
  const j1Period =
    [...byDay.entries()].filter(([, v]) => v >= threshold).map(([p]) => p).sort()[0] ??
    DEFAULT_J1.replace(/-/g, "");

  const reported = new Map<string, Set<number>>();
  for (const row of q2.rows) {
    const ou = row[q2.col.ou];
    const r = records.get(ou);
    if (!r) continue;
    const day = dayIndex(row[q2.col.pe], j1Period);
    if (day < 0) continue;
    const de = row[q2.col.dx];
    const val = parseFloat(row[q2.col.value]) || 0;
    if (!val) continue;
    for (const key of ["nvpo2", "vpob", "rr"] as const) {
      if ((AGE_DE[key] as readonly (string | null)[]).includes(de)) {
        r[key].daily[day] += val;
        break;
      }
    }
    let set = reported.get(ou);
    if (!set) { set = new Set(); reported.set(ou, set); }
    set.add(day);
  }

  // 4. Dérivés : cible polio (convention 0-59 mois) pour les AS RR-POLIO,
  //    population, complétude journalière. Une province n'est retenue RR-POLIO
  //    que si elle a des vaccinés polio ou des cibles polio saisies sur une part
  //    substantielle de ses AS (une cible parasite sur 1-2 AS ne suffit pas).
  const polioVaccTotal = Array.from(records.values()).reduce((a, r) => a + r.nvpo2.vacc + r.vpob.vacc, 0);
  const provincePolio = polioVaccTotal > 0 || polioSeen.size >= 0.3 * Math.max(1, records.size);
  if (!provincePolio) polioSeen.clear();
  for (const [uid, r] of records) {
    if (polioSeen.has(uid)) r.ciblePolio = Math.round((r.cibleRR * PART_POLIO) / PART_RR);
    r.popTotale = Math.round(r.cibleRR / PART_RR);
    const days = reported.get(uid);
    if (days) {
      r.vaccRecus = days.size;
      for (const d of days) r.dailyReports[d].recus = 1;
    }
  }

  const recs = Array.from(records.values()).sort(
    (a, b) => a.zs.localeCompare(b.zs, "fr") || a.as.localeCompare(b.as, "fr")
  );
  return {
    ok: true,
    schema: DHIS2_BLOCK_SCHEMA,
    fetchedAt: new Date().toISOString(),
    provinceId,
    province: provName,
    j1: isoOf(j1Period),
    polio: recs.some((r) => r.ciblePolio > 0),
    records: recs,
  };
}

/* ─── Accès mutualisé aux blocs provinciaux (routes dhis2 + supervision) ── */

/** Noms propres des provinces DHIS2 (niveau 2) — jointure avec le masque importé. */
export const PROVINCE_NAMES: Record<string, string> = {
  rWrCdr321Qu: "Bas Uele", XjeRGfqHMrl: "Equateur", F9w3VW1cQmb: "Haut Katanga",
  fEKDiQIuqeE: "Haut Lomami", wy1lwIP18SL: "Haut Uele", Q4cbnIAo10f: "Ituri",
  dKdrd8HqZWz: "Kongo Central", fgHCmGhaP2X: "Kasai Oriental", PvtAI4RUMkr: "Kwango",
  BmKjwqc6BEw: "Kwilu", TwSa8zUu09Q: "Kinshasa", I8CuQpdBQfP: "Kasai Central",
  D15NtionqkH: "Kasai", dJ3v8xc6ZIK: "Lualaba", an1cK6GbbVw: "Lomami",
  u0vP3ZicczY: "Maindombe", krWZMdwGDIf: "Mongala", uyuwe6bqphf: "Maniema",
  pIAYIpy4hiH: "Nord Kivu", iu4Zj3Zq39m: "Nord Ubangi", GnLX8MNgxZw: "Sud Kivu",
  ybgmW3kIGuq: "Sankuru", JkIljbLc4Ny: "Sud Ubangi", hyvduSNKvfe: "Tanganyika",
  mnOXJ2Oa5U7: "Tshopo", ym2K6YcSNl9: "Tshuapa",
};

function mostCommon(arr: string[]): string {
  const c = new Map<string, number>();
  for (const v of arr) c.set(v, (c.get(v) ?? 0) + 1);
  let best = "";
  let max = -1;
  for (const [k, n] of c) if (n > max) { max = n; best = k; }
  return best;
}

/**
 * Résultats du dernier masque de saisie importé pour cette province (compilation
 * partagée) — TOUJOURS prioritaires sur DHIS2 : certaines provinces (Kasaï
 * Central : Kananga, Luiza) saisissent leurs résultats dans le masque et
 * n'alimentent pas DHIS2. Null si aucun masque n'est importé pour la province.
 */
export async function masqueBlockFor(provinceId: string): Promise<ProvinceBlock | null> {
  const name = PROVINCE_NAMES[provinceId];
  if (!name) return null;
  try {
    const blocks = await readNationalBlocks();
    const match = blocks.filter((b) => normZS(b.province) === normZS(name));
    if (match.length === 0) return null;
    const records = sanitizeRecords(match.flatMap((b) => b.records));
    if (records.length === 0) return null;
    // J1 : le Kasaï Central a lancé sa campagne le 17/08/2026 (confirmé) — les
    // dates du masque ne sont pas fiables (05/08 = valeur par défaut du modèle).
    const dates = match
      .map((b) => b.dateDebut ?? "")
      .filter((d) => d >= "2026-08-10" && d <= "2026-09-30");
    const j1 = normZS(name) === "KASAICENTRAL" ? "2026-08-17" : mostCommon(dates) || "2026-08-17";
    const latest = match.reduce((m, b) => (b.importedAt > m ? b.importedAt : m), "");
    return {
      ok: true,
      schema: DHIS2_BLOCK_SCHEMA,
      fetchedAt: latest || new Date().toISOString(),
      provinceId,
      province: mostCommon(records.map((r) => r.province)) || name,
      j1,
      polio: records.some((r) => r.ciblePolio > 0),
      source: "masque",
      records,
    };
  } catch {
    return null; // KV indisponible → repli DHIS2
  }
}

/** Bloc provincial : masque prioritaire, puis cache KV, puis extraction DHIS2 live. */
export async function getProvinceBlockCached(provinceId: string): Promise<ProvinceBlock | null> {
  const masque = await masqueBlockFor(provinceId);
  if (masque) return masque;
  const key = `rrpolio:dhis2:v${DHIS2_BLOCK_SCHEMA}:prov:${provinceId}`;
  const saved = await kvGetJSON<ProvinceBlock>(key);
  if (saved?.ok && saved.schema === DHIS2_BLOCK_SCHEMA) return saved;
  try {
    const built = await buildProvinceBlock(provinceId);
    await kvSetJSON(key, built);
    return built;
  } catch {
    return null;
  }
}
