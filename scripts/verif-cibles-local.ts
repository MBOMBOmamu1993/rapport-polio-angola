import { buildProvinceBlock } from "../lib/dhis2-campagne";
const IDS: [string, string, number][] = [
  ["D15NtionqkH", "Kasai", 3128640],
  ["fgHCmGhaP2X", "Kasai Oriental", 2921898],
  ["an1cK6GbbVw", "Lomami", 2348889],
  ["uyuwe6bqphf", "Maniema", 1637831],
  ["pIAYIpy4hiH", "Nord Kivu", 5320263],
  ["ybgmW3kIGuq", "Sankuru", 1402218],
  ["GnLX8MNgxZw", "Sud Kivu", 4370433],
];
async function main() {
  for (const [id, name, pbi] of IDS) {
    const b = await buildProvinceBlock(id);
    const t = b.records.reduce((a, r) => ({ c: a.c + r.cibleRR, v: a.v + r.rr.vacc, p: a.p + r.ciblePolio }), { c: 0, v: 0, p: 0 });
    const cv = t.c ? (100 * t.v) / t.c : NaN;
    console.log(
      `${name.padEnd(15)} cibleRR=${Math.round(t.c).toLocaleString("fr-FR").padStart(11)}  PBI=${pbi.toLocaleString("fr-FR").padStart(11)}  écart=${((100 * (t.c - pbi)) / pbi).toFixed(2).padStart(6)} %  CV RR=${cv.toFixed(2)} %  ciblePolio=${Math.round(t.p).toLocaleString("fr-FR")}`
    );
  }
}
main().catch((e) => { console.error(e); process.exit(1); });
