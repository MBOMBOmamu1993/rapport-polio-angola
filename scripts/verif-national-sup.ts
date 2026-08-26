/**
 * Vérification du rapport national avec supervision par province (données de
 * production ; cartes omises en Node — rendues par le navigateur en réel).
 */
import { writeFileSync, readFileSync } from "node:fs";
import path from "node:path";
import { mergeBlocks, type ProvinceBlock } from "../lib/dhis2-shared";
import { buildReportData, prettyProvince, type SupervisionProvince } from "../lib/report-data";
import { exportReportPPT } from "../lib/export-report-pptx";
import type { SupervisionCompact } from "../lib/odk-supervision";

const IDS = ["D15NtionqkH","fgHCmGhaP2X","an1cK6GbbVw","uyuwe6bqphf","pIAYIpy4hiH","ybgmW3kIGuq","GnLX8MNgxZw","I8CuQpdBQfP"];
const DM: Record<string, string> = { I8CuQpdBQfP: "2026-08-17", pIAYIpy4hiH: "2026-08-17" };

async function main() {
  const blocks: ProvinceBlock[] = [];
  const supEntries: SupervisionProvince[] = [];
  for (const id of IDS) {
    const rb = await fetch(`https://rapport-jnv-polio.vercel.app/api/dhis2?province=${id}`);
    const b = (await rb.json()) as ProvinceBlock;
    blocks.push(b);
    const qs = new URLSearchParams({ compact: "1", provinceId: id, province: b.province, dateMin: DM[id] ?? "2026-08-11" });
    const rc = await fetch(`https://rapport-jnv-polio.vercel.app/api/supervision?${qs}`);
    const c = (await rc.json()) as SupervisionCompact;
    if (c.ok) {
      supEntries.push({
        province: b.province,
        provinceLabel: prettyProvince(b.province),
        data: {
          available: true, fetchedAt: c.fetchedAt, dateMin: c.dateMin, formTitle: c.formTitle,
          total: c.total, byZS: c.byZS, conformity: c.conformity, points: c.points,
        },
      });
    }
    console.log(`${b.province}: bloc ok, supervision ${c.ok ? c.total : "ÉCHEC"}`);
  }
  const data = mergeBlocks(blocks)!;
  const report = buildReportData({
    data,
    filters: { provinces: blocks.map((b) => b.province), antenne: null, zs: null, as: null },
    supervision: null,
    supervisionReason: "supervision détaillée par province",
    supervisionParProvince: supEntries,
    actionsPC: [],
    titre: "Campagne intégrée RR-POLIO RD Congo",
    dateLancement: blocks.map((b) => b.j1).sort()[0],
  });
  const loader = async (p: string) => {
    try { const f = path.join(process.cwd(), "public", p); const b = readFileSync(f); const e = path.extname(f).slice(1).toLowerCase();
      return `data:${e === "jpg" || e === "jpeg" ? "image/jpeg" : "image/png"};base64,${b.toString("base64")}`; } catch { return null; }
  };
  const out = (await exportReportPPT(report, { loader, output: "buffer", sourceText: "DHIS2 de campagne (rdccampagne.hispwca.org) + masque de saisie importé (Kasaï Central)" })) as Uint8Array;
  writeFileSync("C:/Users/felly/Downloads/VERIF_national_supervision.pptx", Buffer.from(out));
  console.log("→ VERIF_national_supervision.pptx", (out.length / 1024).toFixed(0), "Ko");
}
main().catch((e) => { console.error(e); process.exit(1); });
