/**
 * Test hors navigateur : masque → parse → données du rapport → PPTX.
 *   npx tsx scripts/test-report.ts <masque.xlsx> <sortie.pptx> [--no-odk] [--zs=NOM] [--antenne=NOM]
 */
import { readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { parseMasque } from "../lib/parse-masque";
import { buildReportData } from "../lib/report-data";
import { exportReportPPT } from "../lib/export-report-pptx";
import { fetchSupervision } from "../lib/odk-server";
import type { SupervisionPayload } from "../lib/odk-supervision";

async function main() {
  const [, , input, output, ...rest] = process.argv;
  if (!input || !output) {
    console.error("usage: tsx scripts/test-report.ts <masque.xlsx> <sortie.pptx> [--no-odk] [--zs=NOM] [--antenne=NOM]");
    process.exit(1);
  }
  const noOdk = rest.includes("--no-odk");
  const zs = rest.find((a) => a.startsWith("--zs="))?.slice(5) ?? null;
  const antenne = rest.find((a) => a.startsWith("--antenne="))?.slice(10) ?? null;

  const buf = readFileSync(input);
  const t0 = Date.now();
  const data = parseMasque(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength), path.basename(input));
  console.log(`parse: ${data.records.length} AS, ${data.meta.zones.length} ZS, antennes=${data.meta.antennes.join("/")}, jours=${data.meta.jourLabels.join(",")}, province=${data.meta.province}, periode=${data.meta.periode}, dateDebut=${data.meta.dateDebut} (${Date.now() - t0} ms)`);
  const r0 = data.records[0];
  console.log("sample:", r0.zs, r0.as, "pop", r0.popTotale, "ciblePolio", r0.ciblePolio, "cibleRR", r0.cibleRR, "rr", r0.rr.vacc, "nvpo2", r0.nvpo2.vacc, "vpob", r0.vpob.vacc, "daily", r0.rr.daily, r0.dailyReports.map((d) => d.recus));

  let sup: SupervisionPayload | null = null;
  let reason: string | undefined;
  if (!noOdk) {
    try {
      sup = await fetchSupervision();
      console.log(`odk: ${sup.total} supervisions depuis ${sup.dateMin}`);
    } catch (e) {
      reason = e instanceof Error ? e.message : String(e);
      console.log("odk error:", reason);
    }
  }

  const report = buildReportData({
    data,
    filters: { provinces: [], antenne, zs, as: null },
    supervision: sup,
    supervisionReason: reason,
    actionsPC: [
      { activite: "Partager le rapport d'investigation des MAPI graves", responsable: "MCA Kananga", statut: "En cours", echeance: "Continu" },
      { activite: "Corriger la liste agrégée des collecteurs (1 numéro par collecteur)", responsable: "MCA Luiza", statut: "Réalisée", echeance: "14 Août 2026" },
      { activite: "Procéder au paiement des collecteurs", responsable: "AFENET CDC", statut: "En retard", echeance: "13 août 2026" },
    ],
  });
  console.log(`report: units=${report.units.length} (${report.byUnitLabel}), antennes=${report.antennes.length}, total RR ${report.total.rr.vacc}/${report.total.rr.cible} CV=${report.total.rr.cv?.toFixed(1)}, compl=${report.total.completude?.toFixed(1)}, sup byZS=${report.supervision.byZS.length}, problemes=${report.problemes.length}`);
  for (const z of report.supervision.byZS.filter((z) => z.nbSupervisions > 0)) console.log("  sup", z.zs, z.nbSupervisions, `${z.nbASVisitees}/${z.nbASTotal}`, z.score?.toFixed(0), "visitées:", z.asVisitees.join("|"));

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
  const t1 = Date.now();
  const out = (await exportReportPPT(report, { loader, output: "buffer" })) as Uint8Array;
  writeFileSync(output, Buffer.from(out));
  console.log(`pptx: ${output} (${(out.length / 1024).toFixed(0)} Ko, ${Date.now() - t1} ms)`);
}

main().catch((e) => { console.error(e); process.exit(1); });
