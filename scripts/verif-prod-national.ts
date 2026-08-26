import { writeFileSync, readFileSync } from "node:fs";
import path from "node:path";
import { mergeBlocks, type ProvinceBlock } from "../lib/dhis2-shared";
import { buildReportData } from "../lib/report-data";
import { exportReportPPT } from "../lib/export-report-pptx";

const IDS = ["D15NtionqkH","I8CuQpdBQfP","fgHCmGhaP2X","an1cK6GbbVw","uyuwe6bqphf","pIAYIpy4hiH","ybgmW3kIGuq","GnLX8MNgxZw"];

async function main() {
  const blocks: ProvinceBlock[] = [];
  for (const id of IDS) {
    for (let t = 0; t < 20; t++) {
      const res = await fetch(`https://rapport-jnv-polio.vercel.app/api/dhis2?province=${id}`);
      if (res.status === 200) { blocks.push((await res.json()) as ProvinceBlock); break; }
      await new Promise((r) => setTimeout(r, 6000));
    }
  }
  const data = mergeBlocks(blocks)!;
  console.log("provinces:", blocks.map((b) => `${b.province}(${b.source ?? "dhis2"}, J1 ${b.j1})`).join(" | "));
  console.log("antennes:", data.meta.antennes.join(", "));
  if (data.meta.antennes.some((a) => /kisangani/i.test(a))) throw new Error("KISANGANI ENCORE PRESENT !");
  const report = buildReportData({
    data,
    filters: { provinces: blocks.map((b) => b.province), antenne: null, zs: null, as: null },
    supervision: null, supervisionReason: "vérification", actionsPC: [],
    titre: "Campagne intégrée RR-POLIO RD Congo", dateLancement: blocks.map((b) => b.j1).sort()[0],
  });
  const kc = report.units.find((u) => /kasai.?central/i.test(u.unit));
  console.log("KC dans le tableau provinces:", kc ? `${kc.unit} — RR ${kc.rr.vacc}/${kc.rr.cible} CV=${kc.rr.cv?.toFixed(1)}% compl=${kc.completude?.toFixed(1)}%` : "ABSENT");
  console.log("total: RR", report.total.rr.vacc, "/", report.total.rr.cible, "CV", report.total.rr.cv?.toFixed(1), "%");
  const loader = async (p: string) => {
    try { const f = path.join(process.cwd(), "public", p); const b = readFileSync(f); const e = path.extname(f).slice(1).toLowerCase();
      return `data:${e === "jpg" || e === "jpeg" ? "image/jpeg" : "image/png"};base64,${b.toString("base64")}`; } catch { return null; }
  };
  const out = (await exportReportPPT(report, { loader, output: "buffer", sourceText: "DHIS2 de campagne (rdccampagne.hispwca.org) + masque de saisie importé (Kasaï Central)" })) as Uint8Array;
  const dest = "C:/Users/felly/Downloads/VERIF_Bloc3_national_prod.pptx";
  writeFileSync(dest, Buffer.from(out));
  console.log(dest, (out.length / 1024).toFixed(0), "Ko");
}
main().catch((e) => { console.error(e); process.exit(1); });
