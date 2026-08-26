/**
 * Génération des rapports PPTX de l'Antenne PEV Lubutu (Maniema) à partir de DHIS2.
 *
 *   npx tsx scripts/rapport-dhis2-lubutu.ts [--no-odk] [--sortie=C:\...\Downloads]
 *
 * Contrairement au Kasaï Central (masque Excel), les données Maniema viennent du
 * DHIS2 de campagne (rdccampagne.hispwca.org, dataset « PEV_Campagne … RR et Polio »,
 * valeurs brutes /api/dataValueSets — pas les analytics). Le script reconstruit un
 * `MasqueData` équivalent au masque puis réutilise le pipeline habituel
 * (buildReportData → exportReportPPT). Supervision : ODK (env ODK_PROVINCE=Maniema,
 * ODK_DATE_MIN=2026-08-11).
 *
 * Conventions alignées sur le tableau de bord Power BI officiel du Bloc 3 :
 *  - cible RR = « RR_Cible attendue RR » telle que saisie (46 % pop) ;
 *  - cible polio = cible RR × (18,9/46) — cible 0-59 mois, PAS le DE
 *    « Polio_Cible attendue Polio » (0-9 ans) que le tableau de bord n'utilise pas ;
 *  - population totale = cible RR ÷ 0,46 ;
 *  - J1..J5 = 11..15/08/2026, tout jour ≥ 16/08 = Ratissage.
 */

import { readFileSync, writeFileSync, existsSync, statSync } from "node:fs";
import path from "node:path";
import https from "node:https";
import type { ASRecord, MasqueData, VaccineStats, DailyReport } from "../lib/parse-masque";
import { MASQUE_SCHEMA, ANTIGENES } from "../lib/parse-masque";
import { buildReportData } from "../lib/report-data";
import { exportReportPPT } from "../lib/export-report-pptx";
import { fetchSupervision } from "../lib/odk-server";
import type { SupervisionPayload } from "../lib/odk-supervision";

/* ─── DHIS2 (connexion directe IP : le DNS local ne résout pas le domaine) ─ */

const DHIS_HOST = "rdccampagne.hispwca.org";
const DHIS_IP = "161.97.129.236";
const DHIS_AUTH = "Basic " + Buffer.from("RR_Polio:Snis@2026").toString("base64");

function dhis(pathname: string): Promise<any> {
  return new Promise((resolve, reject) => {
    const req = https.request(
      {
        host: DHIS_IP,
        servername: DHIS_HOST,
        path: "/dhis/api" + pathname,
        method: "GET",
        headers: { Host: DHIS_HOST, Authorization: DHIS_AUTH, Accept: "application/json" },
        timeout: 180000,
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

/* ─── Paramètres de la campagne Antenne Lubutu ─────────────────────────── */

const PROVINCE = "Maniema";
const ANTENNE = "Lubutu";
const ZONES: Record<string, string> = {
  Ferekeni: "YuGk2dXTvKj",
  Lubutu: "HtQZ7yf79SN",
  Obokote: "BsSB1PA58Qp",
  Punia: "EyMlz2Qvcmt",
};
const DATASET = "HEcwCd5Rd8U";
const DATE_DEBUT = "2026-08-11"; // J1
const START = "2026-08-01";
const END = "2026-08-25";
const JOUR_LABELS = ["J1", "J2", "J3", "J4", "J5", "Ratissage"];
const NB_JOURS = 6;
/** Conventions cibles du tableau de bord officiel (Power BI Bloc 3). */
const PART_RR = 0.46; // 6 mois - 14 ans
const PART_POLIO = 0.189; // 0 - 59 mois

/** Jour de campagne (0..5) d'une période DHIS2 « yyyymmdd », ou -1 hors campagne. */
function dayIndex(period: string): number {
  if (period < "20260811") return -1;
  if (period <= "20260815") return Number(period.slice(6)) - 11; // 11..15 → 0..4
  return 5; // 16/08 et suivants → ratissage
}

/* ─── Éléments de données / combos ─────────────────────────────────────── */

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
  default: "HllvX50cXC0",
} as const;

/** DE des vaccinés par tranche d'âge, dans l'ordre des tableaux `ages` du masque. */
const AGE_DE = {
  // POLIO_AGE_LABELS = [0-5 mois, 6-11 mois, 12-59 mois, 5-9 ans] — pas de DE 5-9 ans
  nvpo2: ["hcjjuPEmau2", "trxWpua3GB1", "KhMa1ehdwvY", null],
  vpob: ["bWZF88KmH7E", "bVoRjDT5DUN", "NJ7bAfA9lUY", null],
  // RR_AGE_LABELS = [6-11 mois, 12-59 mois, 5-14 ans]
  rr: ["LUAiDML2q1j", "iRJtRKAEOPQ", "ztjsflVdf1c"],
} as const;

const FLACON_DE = { nvpo2: "nYq937kfajz", vpob: "V38AKsoLCvy", rr: "Xgw4c99kK0p" } as const;
const CIBLE_RR_DE = "BJ6CjSbAiLJ";
const SURV_DE = { pfa: "ZexanWv3fFM", rougeole: "dYhGRWbkXbT", fj: "n6JhuMhdGbf", tnn: "jMJ7abE4cKb" } as const;
const MAPI_DE = { graves: "YG0JCwIikUU", nonGraves: "gF1so9TfqUX" } as const;
/** Récupération pendant les AVS — ordre du masque : [nVPO2, VPOb, RR]. */
const AVS_DE: string[][] = [
  ["qZ6OXVLAZoH", "VVTa5KoZfek"],
  ["XUKHSDhV4np", "vKW3SByKOz3"],
  ["Tj6Uug4aaQj", "WQOKvgWStMm"],
];
/** PEV systématique : DE par antigène, dans l'ordre de ANTIGENES. */
const PEV_DE: Record<string, string> = {
  BCG: "ATZtz3vX9Cj", DTC1: "iFW7yj3ftkK", PCV1: "eGWjLe4JqkZ", ROTA1: "aKsA7JcExZw",
  DTC2: "d9gOS5k17rX", PCV2: "LnynmJNoki9", ROTA2: "Smopcheplu8", DTC3: "ZHaMTWh7MHh",
  PCV3: "l09sbfgjBF8", ROTA3: "ptTjjrd4RBa", VPI1: "sZaMowoYB8o", VAP1: "SYaFLJIuU96",
  VAP2: "yKinSHZrRhP", VAP3: "tPqHdo4GoYa", VPI2: "WDNopj5AcpF", VAA: "dJqSnLSMKik",
  VAP4: "Pfc354bq0cE", TD1: "kNdUxoWhUS5", TD2: "ZzEhrEZu261", TD3: "Vt354qVK08v",
  TD4: "uryaNxO3R1F", TD5: "x3AP46l37ge",
};

/* ─── Construction du MasqueData ───────────────────────────────────────── */

interface RawValue { dataElement: string; period: string; orgUnit: string; categoryOptionCombo: string; value: string }

function cleanName(n: string): string {
  return n.replace(/^mn\s+/i, "").replace(/\s+(Aire|Zone) de Sant[eé]$/i, "").trim();
}

function emptyVacc(nbAges: number): VaccineStats {
  return {
    vacc: 0, zeroDose: 0, garcons: 0, filles: 0, ages: new Array(nbAges).fill(0),
    flaconsRecus: 0, flaconsUtil: 0, flaconsRendus: 0, flaconsPerdus: 0,
    daily: new Array(NB_JOURS).fill(0),
  };
}

async function buildMasqueFromDHIS2(): Promise<MasqueData> {
  const records = new Map<string, ASRecord>(); // clé = uid AS
  const reported = new Map<string, Set<number>>(); // uid AS → jours avec données de vaccination

  for (const [zsName, zsId] of Object.entries(ZONES)) {
    const meta = await dhis(`/organisationUnits/${zsId}.json?fields=id,name,children[id,name,level]`);
    const aires: { id: string; name: string }[] = meta.children
      .filter((c: any) => c.level === 4)
      .map((c: any) => ({ id: c.id, name: cleanName(c.name) }))
      .sort((a: any, b: any) => a.name.localeCompare(b.name, "fr"));

    for (const as of aires) {
      records.set(as.id, {
        province: PROVINCE, antenne: ANTENNE, zs: zsName, as: as.name,
        popTotale: 0, menagesPrevus: 0, menagesVisites: 0, mosoAttendus: 0, mosoRecus: 0,
        pers15: 0, refusSignales: 0, refusGeres: 0,
        vaccAttendus: NB_JOURS, vaccRecus: 0,
        dailyReports: JOUR_LABELS.map((label): DailyReport => ({ label, attendus: 1, recus: 0 })),
        ciblePolio: 0, cibleRR: 0,
        nvpo2: emptyVacc(4), vpob: emptyVacc(4), rr: emptyVacc(3),
        survPFA: 0, survRougeole: 0, survFJ: 0, survTNN: 0,
        mapiNonGraves: 0, mapiGraves: 0,
        aidantsIdent: [0, 0, 0], aidantsRecup: [0, 0, 0],
        pevIdent: new Array(ANTIGENES.length).fill(0),
        pevRecup: new Array(ANTIGENES.length).fill(0),
      });
      reported.set(as.id, new Set());
    }

    // Les données sont saisies au niveau FOSA (niveau 5) : on rattache chaque
    // formation sanitaire à son Aire de Santé parente.
    const fosas = await dhis(`/organisationUnits.json?filter=path:like:${zsId}&filter=level:eq:5&fields=id,parent[id]&paging=false`);
    const parentOf = new Map<string, string>();
    for (const f of fosas.organisationUnits ?? []) parentOf.set(f.id, f.parent?.id ?? "");

    const dvs = await dhis(`/dataValueSets.json?dataSet=${DATASET}&orgUnit=${zsId}&children=true&startDate=${START}&endDate=${END}`);
    const values: RawValue[] = dvs.dataValues ?? [];
    console.log(`DHIS2 ${zsName}: ${aires.length} AS, ${fosas.organisationUnits?.length ?? 0} FOSA, ${values.length} valeurs`);

    for (const v of values) {
      const ouAS = records.has(v.orgUnit) ? v.orgUnit : parentOf.get(v.orgUnit) ?? "";
      const r = records.get(ouAS);
      if (!r) continue;
      const val = parseFloat(v.value) || 0;
      const de = v.dataElement;
      const coc = v.categoryOptionCombo;
      const day = dayIndex(v.period);

      // Cibles (saisies à J-1)
      if (de === CIBLE_RR_DE) { r.cibleRR += val; continue; }

      // Vaccinés par tranche d'âge / sexe / dose
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
        if (day >= 0) {
          s.daily[day] += val;
          reported.get(ouAS)!.add(day);
        }
        handled = true;
        break;
      }
      if (handled) continue;

      // Gestion des flacons
      for (const key of ["nvpo2", "vpob", "rr"] as const) {
        if (de !== FLACON_DE[key]) continue;
        const s = r[key];
        if (coc === COC.recue) s.flaconsRecus += val;
        else if (coc === COC.utilisee) s.flaconsUtil += val;
        else if (coc === COC.rendue) s.flaconsRendus += val;
        else if (coc === COC.perdue) s.flaconsPerdus += val;
        handled = true;
        break;
      }
      if (handled) continue;

      // Surveillance MPV
      if (de === SURV_DE.pfa) { r.survPFA += val; continue; }
      if (de === SURV_DE.rougeole) { r.survRougeole += val; continue; }
      if (de === SURV_DE.fj) { r.survFJ += val; continue; }
      if (de === SURV_DE.tnn) { r.survTNN += val; continue; }

      // MAPI (le tableau de bord officiel n'utilise que « Notifiées »)
      if (de === MAPI_DE.graves) { if (coc === COC.notifiees) r.mapiGraves += val; continue; }
      if (de === MAPI_DE.nonGraves) { if (coc === COC.notifiees) r.mapiNonGraves += val; continue; }

      // Récupération pendant les AVS ([nVPO2, VPOb, RR])
      const avs = AVS_DE.findIndex((pair) => pair.includes(de));
      if (avs >= 0) {
        if (coc === COC.ident) r.aidantsIdent[avs] += val;
        else if (coc === COC.recup) r.aidantsRecup[avs] += val;
        continue;
      }

      // PEV systématique
      const ag = ANTIGENES.findIndex((a) => PEV_DE[a.key] === de);
      if (ag >= 0) {
        if (coc === COC.ident) r.pevIdent[ag] += val;
        else if (coc === COC.recup) r.pevRecup[ag] += val;
      }
    }
  }

  // Dérivés : cible polio (convention 0-59 mois du tableau de bord), population,
  // complétude journalière (rapport reçu = au moins une donnée de vaccination le jour J).
  for (const [uid, r] of records) {
    r.ciblePolio = Math.round((r.cibleRR * PART_POLIO) / PART_RR);
    r.popTotale = Math.round(r.cibleRR / PART_RR);
    const days = reported.get(uid)!;
    r.vaccRecus = days.size;
    for (const d of days) r.dailyReports[d].recus = 1;
  }

  const recs = Array.from(records.values());
  return {
    meta: {
      pays: "RD CONGO",
      periode: "du 11 au 15 Août 2026 (+ ratissage)",
      dateDebut: DATE_DEBUT,
      dateFin: "2026-08-17",
      province: PROVINCE,
      antennes: [ANTENNE],
      zones: Object.keys(ZONES).sort((a, b) => a.localeCompare(b, "fr")),
      importedAt: new Date().toISOString(),
      fileName: "DHIS2 rdccampagne.hispwca.org — dataset PEV_Campagne RR-Polio",
      nbAires: recs.length,
      nbJours: NB_JOURS,
      jourLabels: JOUR_LABELS,
      schema: MASQUE_SCHEMA,
    },
    records: recs,
  };
}

/* ─── Supervision ODK (province Maniema, avec cache disque) ────────────── */

const SUP_CACHE = path.join(process.cwd(), "scripts", ".sup-maniema-cache.json");
const SUP_CACHE_MAX_AGE_MS = 6 * 3600_000;

async function getSupervision(): Promise<{ sup: SupervisionPayload | null; reason?: string }> {
  if (existsSync(SUP_CACHE) && Date.now() - statSync(SUP_CACHE).mtimeMs < SUP_CACHE_MAX_AGE_MS) {
    const sup = JSON.parse(readFileSync(SUP_CACHE, "utf8")) as SupervisionPayload;
    if (sup.ok && !sup.partial) {
      console.log(`ODK (cache disque): ${sup.total} supervisions`);
      return { sup };
    }
  }
  try {
    let sup = await fetchSupervision();
    let rounds = 1;
    while (sup.partial && rounds < 8) {
      console.log(`ODK: extraction partielle (${sup.total} enregistrements) — reprise ${rounds}…`);
      sup = await fetchSupervision({ force: true });
      rounds++;
    }
    writeFileSync(SUP_CACHE, JSON.stringify(sup));
    console.log(`ODK: ${sup.total} supervisions depuis ${sup.dateMin}${sup.partial ? " (encore partiel)" : ""}`);
    return { sup };
  } catch (e) {
    const reason = e instanceof Error ? e.message : String(e);
    console.log("ODK erreur:", reason);
    if (existsSync(SUP_CACHE)) {
      const sup = JSON.parse(readFileSync(SUP_CACHE, "utf8")) as SupervisionPayload;
      console.log(`ODK: repli sur le cache disque (${sup.total})`);
      return { sup };
    }
    return { sup: null, reason };
  }
}

/* ─── Génération ───────────────────────────────────────────────────────── */

/**
 * Mode --powerbi : reproduit l'état affiché par le tableau de bord Power BI officiel
 * (non resynchronisé depuis la correction DHIS2 de l'AS Osso du 25/08 à 08h41).
 * Seuls les vaccinés/flacons polio d'Osso diffèrent ; les cibles restent celles du
 * DHIS2 (demande utilisateur : « pour RR considère cible DHIS2 »).
 * Valeurs cibles vérifiées dans le modèle PBI : Osso nVPO2 1854 (M 1048 / F 806),
 * VPOb 1852 (M 1048 / F 804), tranche 12-59 mois 1501, doses utilisées 1950 (39
 * flacons nVPO2) et 1980 (99 flacons VPOb).
 */
function applyPowerBIState(data: MasqueData): void {
  const r = data.records.find((x) => x.zs === "Lubutu" && x.as === "Osso");
  if (!r) throw new Error("AS Osso introuvable");
  for (const key of ["nvpo2", "vpob"] as const) {
    const s = r[key];
    s.vacc += 200;
    s.ages[2] += 200; // 12-59 mois (J1)
    s.garcons += 100;
    s.filles += 100;
    s.daily[0] += 200;
  }
  const setUtil = (key: "nvpo2" | "vpob", pbiUtil: number) => {
    const s = r[key];
    const delta = pbiUtil - s.flaconsUtil;
    s.flaconsUtil += delta;
    s.flaconsRecus += delta;
  };
  setUtil("nvpo2", 39);
  setUtil("vpob", 99);
}

async function main() {
  const noOdk = process.argv.includes("--no-odk");
  const powerbi = process.argv.includes("--powerbi");
  const outDir = process.argv.find((a) => a.startsWith("--sortie="))?.slice(9) ?? "C:\\Users\\felly\\Downloads";

  const data = await buildMasqueFromDHIS2();
  if (powerbi) {
    applyPowerBIState(data);
    console.log("Mode Power BI : état pré-correction de l'AS Osso appliqué (+200 nVPO2, +200 VPOb).");
  }
  const t = { rr: 0, nvpo2: 0, vpob: 0, cibR: 0, cibP: 0 };
  for (const r of data.records) {
    t.rr += r.rr.vacc; t.nvpo2 += r.nvpo2.vacc; t.vpob += r.vpob.vacc;
    t.cibR += r.cibleRR; t.cibP += r.ciblePolio;
  }
  console.log(
    `Antenne ${ANTENNE}: ${data.records.length} AS — RR ${t.rr}/${t.cibR} (${((100 * t.rr) / t.cibR).toFixed(1)} %), ` +
    `nVPO2 ${t.nvpo2}/${t.cibP} (${((100 * t.nvpo2) / t.cibP).toFixed(1)} %), VPOb ${t.vpob}/${t.cibP} (${((100 * t.vpob) / t.cibP).toFixed(1)} %)`
  );

  let sup: SupervisionPayload | null = null;
  let reason: string | undefined;
  if (!noOdk) ({ sup, reason } = await getSupervision());

  const loader = async (p: string): Promise<string | null> => {
    try {
      const f = path.join(process.cwd(), "public", p);
      const b = readFileSync(f);
      const ext = path.extname(f).slice(1).toLowerCase();
      const mime = ext === "jpg" || ext === "jpeg" ? "image/jpeg" : "image/png";
      return `data:${mime};base64,${b.toString("base64")}`;
    } catch {
      return null;
    }
  };

  const actionsPC = [
    { activite: "Consolider les actions du poste de commandement (à compléter par le PC)", responsable: "PC Antenne Lubutu", statut: "En cours", echeance: "—" },
  ];

  const suffix = powerbi ? "_PowerBI" : "";
  const sourceText = powerbi
    ? "Power BI officiel de la campagne (Bloc 3) — cibles RR : DHIS2"
    : undefined;
  const jobs: { label: string; filters: any; file: string }[] = [
    {
      label: `Antenne ${ANTENNE}`,
      filters: { provinces: [], antenne: ANTENNE, zs: null, as: null },
      file: `Resultats_partiels_RR_Polio_Antenne_LUBUTU${suffix}_25-08-2026.pptx`,
    },
    {
      label: "ZS Lubutu",
      filters: { provinces: [], antenne: null, zs: "Lubutu", as: null },
      file: `Resultats_partiels_RR_Polio_Zone_de_Sante_LUBUTU${suffix}_25-08-2026.pptx`,
    },
  ];

  for (const job of jobs) {
    const report = buildReportData({
      data,
      filters: job.filters,
      supervision: sup,
      supervisionReason: reason,
      actionsPC,
      dateLancement: DATE_DEBUT,
    });
    report.titre = "Campagne intégrée RR-POLIO Maniema";
    console.log(
      `${job.label}: units=${report.units.length} (${report.byUnitLabel}), total RR ${report.total.rr.vacc}/${report.total.rr.cible} ` +
      `CV=${report.total.rr.cv?.toFixed(1)} %, nVPO2 CV=${report.total.nvpo2.cv?.toFixed(1)} %, compl=${report.total.completude?.toFixed(1)} %, ` +
      `sup=${report.supervision.total}, problemes=${report.problemes.length}`
    );
    const out = (await exportReportPPT(report, { loader, output: "buffer", sourceText })) as Uint8Array;
    const dest = path.join(outDir, job.file);
    writeFileSync(dest, Buffer.from(out));
    console.log(`  → ${dest} (${(out.length / 1024).toFixed(0)} Ko)`);
  }
}

main().catch((e) => { console.error(e); process.exit(1); });
