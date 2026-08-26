import { type ProvinceBlock } from "../lib/dhis2-shared";
const IDS = ["D15NtionqkH","I8CuQpdBQfP","fgHCmGhaP2X","an1cK6GbbVw","uyuwe6bqphf","pIAYIpy4hiH","ybgmW3kIGuq","GnLX8MNgxZw"];
async function main() {
  let cibR = 0, vacR = 0, cibP = 0, vacN = 0, vacB = 0;
  for (const id of IDS) {
    const res = await fetch(`https://rapport-jnv-polio.vercel.app/api/dhis2?province=${id}`);
    if (res.status !== 200) { console.log(id, "HTTP", res.status); continue; }
    const b = (await res.json()) as ProvinceBlock;
    const t = b.records.reduce((a, r) => ({
      cibR: a.cibR + r.cibleRR, vacR: a.vacR + r.rr.vacc, cibP: a.cibP + r.ciblePolio,
      vacN: a.vacN + r.nvpo2.vacc, vacB: a.vacB + r.vpob.vacc,
    }), { cibR: 0, vacR: 0, cibP: 0, vacN: 0, vacB: 0 });
    console.log(
      b.province.padEnd(15),
      `cibleRR=${Math.round(t.cibR).toLocaleString("fr-FR")}`.padEnd(22),
      `ciblePolio=${Math.round(t.cibP).toLocaleString("fr-FR")}`.padEnd(24),
      `${b.polio ? "RR-POLIO" : "RR seule"}`
    );
    cibR += t.cibR; vacR += t.vacR; cibP += t.cibP; vacN += t.vacN; vacB += t.vacB;
  }
  console.log("\nNATIONAL :");
  console.log(`  CV RR    = ${Math.round(vacR).toLocaleString("fr-FR")} / ${Math.round(cibR).toLocaleString("fr-FR")} (cible des 8 provinces) = ${((100*vacR)/cibR).toFixed(1)} %`);
  console.log(`  CV nVPO2 = ${Math.round(vacN).toLocaleString("fr-FR")} / ${Math.round(cibP).toLocaleString("fr-FR")} (cible des seules provinces polio) = ${((100*vacN)/cibP).toFixed(1)} %`);
  console.log(`  CV VPOb  = ${Math.round(vacB).toLocaleString("fr-FR")} / ${Math.round(cibP).toLocaleString("fr-FR")} = ${((100*vacB)/cibP).toFixed(1)} %`);
}
main().catch((e) => { console.error(e); process.exit(1); });
