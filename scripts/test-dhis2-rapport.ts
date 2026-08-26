/**
 * Banc d'essai du moteur « RR-polio DHIS2 Bloc 3 » :
 *   npx tsx scripts/test-dhis2-rapport.ts [--pptx]
 *
 * 1. Liste les provinces de la campagne ;
 * 2. Extrait Maniema et compare l'Antenne Lubutu à l'étalon du 25/08
 *    (script rapport-dhis2-lubutu.ts : RR 63 335 / 74 172, nVPO2 61 232 / 30 477) ;
 * 3. Génère les PPTX de contrôle (multi-provinces, province, antenne, ZS).
 */

import { writeFileSync, readFileSync } from "node:fs";
import path from "node:path";
import { buildProvinceBlock, listCampaignProvinces } from "../lib/dhis2-campagne";
import { mergeBlocks, type ProvinceBlock } from "../lib/dhis2-shared";
import { buildReportData } from "../lib/report-data";
import { exportReportPPT } from "../lib/export-report-pptx";
import type { Filters } from "../lib/store";

const OUT = "C:\\Users\\felly\\Downloads";

async function main() {
  const doPptx = process.argv.includes("--pptx");

  const provinces = await listCampaignProvinces();
  console.log(`Provinces de la campagne : ${provinces.length}`);
  for (const p of provinces) {
    console.log(
      `  - ${p.name.padEnd(16)} ${p.polio ? "RR-POLIO" : "RR seule"}  cibleRR=${p.cibleRR}  rrVacc=${p.rrVacc}  polioVacc=${p.polioVacc}`
    );
  }

  const blocks: ProvinceBlock[] = [];
  for (const p of provinces) {
    const t0 = Date.now();
    const b = await buildProvinceBlock(p.id);
    blocks.push(b);
    const tot = b.records.reduce(
      (a, r) => ({
        rr: a.rr + r.rr.vacc, cibR: a.cibR + r.cibleRR,
        nv: a.nv + r.nvpo2.vacc, cibP: a.cibP + r.ciblePolio,
        recus: a.recus + r.vaccRecus,
      }),
      { rr: 0, cibR: 0, nv: 0, cibP: 0, recus: 0 }
    );
    console.log(
      `${b.province.padEnd(16)} J1=${b.j1} AS=${b.records.length} RR=${tot.rr}/${tot.cibR}` +
      ` (${tot.cibR ? ((100 * tot.rr) / tot.cibR).toFixed(1) : "—"} %)` +
      ` nVPO2=${tot.nv}/${tot.cibP} polio=${b.polio} — ${((Date.now() - t0) / 1000).toFixed(1)} s`
    );
  }

  // Étalon : Antenne Lubutu (Maniema) — comparer au script d'hier.
  const maniema = blocks.find((b) => /maniema/i.test(b.province));
  if (maniema) {
    const lubutu = maniema.records.filter((r) => r.antenne === "Lubutu");
    const t = lubutu.reduce(
      (a, r) => ({ rr: a.rr + r.rr.vacc, cibR: a.cibR + r.cibleRR, nv: a.nv + r.nvpo2.vacc, vb: a.vb + r.vpob.vacc, cibP: a.cibP + r.ciblePolio }),
      { rr: 0, cibR: 0, nv: 0, vb: 0, cibP: 0 }
    );
    console.log(
      `\nÉTALON Antenne Lubutu : ${lubutu.length} AS — RR ${t.rr}/${t.cibR} (${((100 * t.rr) / t.cibR).toFixed(1)} %), ` +
      `nVPO2 ${t.nv}/${t.cibP} (${((100 * t.nv) / t.cibP).toFixed(1)} %), VPOb ${t.vb}/${t.cibP} (${((100 * t.vb) / t.cibP).toFixed(1)} %)`
    );
    const zsSet = new Set(lubutu.map((r) => r.zs));
    console.log(`ZS de l'antenne : ${Array.from(zsSet).sort().join(", ")}`);
  }

  if (!doPptx) return;

  const loader = async (p: string): Promise<string | null> => {
    try {
      const f = path.join(process.cwd(), "public", p);
      const b = readFileSync(f);
      const ext = path.extname(f).slice(1).toLowerCase();
      return `data:${ext === "jpg" || ext === "jpeg" ? "image/jpeg" : "image/png"};base64,${b.toString("base64")}`;
    } catch {
      return null;
    }
  };
  const sourceText = "DHIS2 de campagne (rdccampagne.hispwca.org) — dataset PEV_Campagne RR et Polio";

  const jobs: { label: string; blocks: ProvinceBlock[]; filters: Filters; titre: string; file: string }[] = [
    {
      label: "Toutes provinces",
      blocks,
      filters: { provinces: blocks.map((b) => b.province), antenne: null, zs: null, as: null },
      titre: "Campagne intégrée RR-POLIO RD Congo",
      file: "TEST_DHIS2_Toutes_Provinces.pptx",
    },
    {
      label: "Province Sud Kivu",
      blocks: blocks.filter((b) => /sud kivu/i.test(b.province)),
      filters: { provinces: ["Sud Kivu"], antenne: null, zs: null, as: null },
      titre: "Campagne intégrée RR-POLIO Sud Kivu",
      file: "TEST_DHIS2_Province_Sud_Kivu.pptx",
    },
    {
      label: "Antenne Lubutu",
      blocks: blocks.filter((b) => /maniema/i.test(b.province)),
      filters: { provinces: ["Maniema"], antenne: "Lubutu", zs: null, as: null },
      titre: "Campagne intégrée RR-POLIO Maniema",
      file: "TEST_DHIS2_Antenne_Lubutu.pptx",
    },
    {
      label: "ZS Lubutu",
      blocks: blocks.filter((b) => /maniema/i.test(b.province)),
      filters: { provinces: ["Maniema"], antenne: null, zs: "Lubutu", as: null },
      titre: "Campagne intégrée RR-POLIO Maniema",
      file: "TEST_DHIS2_ZS_Lubutu.pptx",
    },
    {
      label: "Province RR seule (Lomami)",
      blocks: blocks.filter((b) => /lomami/i.test(b.province) && !/haut/i.test(b.province)),
      filters: { provinces: ["Lomami"], antenne: null, zs: null, as: null },
      titre: "Campagne RR Lomami",
      file: "TEST_DHIS2_Province_Lomami_RRseule.pptx",
    },
  ];

  for (const job of jobs) {
    const data = mergeBlocks(job.blocks);
    if (!data) { console.log(`${job.label}: pas de données`); continue; }
    const report = buildReportData({
      data,
      filters: job.filters,
      supervision: null,
      supervisionReason: "test sans ODK",
      actionsPC: [],
      titre: job.titre,
      dateLancement: job.blocks.map((b) => b.j1).sort()[0],
    });
    console.log(
      `\n${job.label}: unités=${report.units.length} (par ${report.byUnitLabel}), antennes=${report.antennes.length}, ` +
      `RR ${report.total.rr.vacc}/${report.total.rr.cible} CV=${report.total.rr.cv?.toFixed(1)} %, ` +
      `nVPO2 CV=${report.total.nvpo2.cv?.toFixed(1) ?? "—"} %, compl=${report.total.completude?.toFixed(1)} %, problèmes=${report.problemes.length}`
    );
    const out = (await exportReportPPT(report, { loader, output: "buffer", sourceText })) as Uint8Array;
    const dest = path.join(OUT, job.file);
    writeFileSync(dest, Buffer.from(out));
    console.log(`  → ${dest} (${(out.length / 1024).toFixed(0)} Ko)`);
  }
}

main().catch((e) => { console.error(e); process.exit(1); });
