// Test navigateur (Playwright) : import du masque → page rapport → téléchargement du PPTX.
//   NODE_PATH=<chemin node_modules contenant playwright> node scripts/e2e-browser.mjs <url> <masque.xlsx> <dossier_sortie>
import { createRequire } from "node:module";
import path from "node:path";
import fs from "node:fs";
const require = createRequire(process.env.PW_REQUIRE_FROM || import.meta.url);
const { chromium } = require("playwright");

const [, , base, masque, outDir] = process.argv;
fs.mkdirSync(outDir, { recursive: true });
const browser = await chromium.launch({ headless: true });
const ctx = await browser.newContext({ viewport: { width: 1400, height: 1000 }, acceptDownloads: true });
const page = await ctx.newPage();
page.on("console", (m) => { if (m.type() === "error") console.log("console.error:", m.text().slice(0, 300)); });
page.on("pageerror", (e) => console.log("pageerror:", e.message));

await page.goto(`${base}/import`, { waitUntil: "networkidle" });
await page.screenshot({ path: path.join(outDir, "import-1.png"), fullPage: true });
await page.setInputFiles('input[type="file"]', masque);
await page.waitForSelector("text=Aperçu instantané", { timeout: 120000 });
await page.waitForTimeout(1500);
await page.screenshot({ path: path.join(outDir, "import-2.png"), fullPage: true });

await page.goto(`${base}/rapport`, { waitUntil: "load" });
await page.waitForSelector("text=Connecté à ODK", { timeout: 150000 }).catch(() => console.log("ODK badge not ok"));
await page.waitForTimeout(1000);
await page.screenshot({ path: path.join(outDir, "rapport-1.png"), fullPage: true });

const [download] = await Promise.all([
  page.waitForEvent("download", { timeout: 180000 }),
  page.click("text=Télécharger le rapport en PowerPoint"),
]);
const target = path.join(outDir, "browser_out.pptx");
await download.saveAs(target);
console.log("downloaded:", target, fs.statSync(target).size);
await page.waitForTimeout(500);
await page.screenshot({ path: path.join(outDir, "rapport-2.png"), fullPage: true });
await browser.close();
