import { writeFileSync, readFileSync } from "node:fs";
import path from "node:path";
import { buildProvinceBlock } from "../lib/dhis2-campagne";
import { mergeBlocks, type ProvinceBlock } from "../lib/dhis2-shared";
import { buildReportData } from "../lib/report-data";
import { exportReportPPT } from "../lib/export-report-pptx";

const EXTRA: Record<string, [string, string]> = {
  ke: ["fgHCmGhaP2X", "Kasai Oriental"],
  nk: ["pIAYIpy4hiH", "Nord Kivu"],
  sk: ["GnLX8MNgxZw", "Sud Kivu"],
};
const IDS: [string, string][] = [
  ["D15NtionqkH", "Kasai"],
  ["an1cK6GbbVw", "Lomami"],
  ["ybgmW3kIGuq", "Sankuru"],
  ["I8CuQpdBQfP", "Kasai Central"],
  ["uyuwe6bqphf", "Maniema"],
];

async function main() {
  const extra = process.argv[2] ? EXTRA[process.argv[2]] : null;
  const ids = extra ? [...IDS, extra] : IDS;
  const blocks: ProvinceBlock[] = [];
  for (const [id] of ids) blocks.push(await buildProvinceBlock(id));
  const data = mergeBlocks(blocks)!;
  const report = buildReportData({
    data,
    filters: { provinces: ids.map(([, n]) => n), antenne: null, zs: null, as: null },
    supervision: null,
    supervisionReason: "supervision ODK affichée pour une province à la fois",
    actionsPC: [],
    titre: "Campagne intégrée RR-POLIO RD Congo",
    dateLancement: blocks.map((b) => b.j1).sort()[0],
  });
  const loader = async (p: string): Promise<string | null> => {
    try {
      const f = path.join(process.cwd(), "public", p);
      const b = readFileSync(f);
      const ext = path.extname(f).slice(1).toLowerCase();
      return `data:${ext === "jpg" || ext === "jpeg" ? "image/jpeg" : "image/png"};base64,${b.toString("base64")}`;
    } catch { return null; }
  };
  const out = (await exportReportPPT(report, { loader, output: "buffer", sourceText: "DHIS2 de campagne (rdccampagne.hispwca.org) — dataset PEV_Campagne RR et Polio" })) as Uint8Array;
  const dest = "C:/Users/felly/Downloads/REPRO_multi_" + (process.argv[2] ?? "5prov") + ".pptx";
  writeFileSync(dest, Buffer.from(out));
  console.log(dest, (out.length / 1024).toFixed(0), "Ko — units:", report.units.length, "antennes:", report.antennes.length);
}
main().catch((e) => { console.error(e); process.exit(1); });
