/**
 * Génération du rapport PowerPoint « Résultats partiels de la campagne intégrée
 * Rougeole‑Rubéole et Poliomyélite — Kasaï Central ».
 *
 * Reproduit le modèle de référence (présentation Bloc 3, août 2026) : page de
 * garde, plan, suivi des points d'action, points saillants, synthèse des
 * indicateurs, complétude, couvertures et taux de perte (RR, nVPO2, VPOb),
 * MAPI, MPV, récupération PEV, tranches d'âge, sexe, supervision ODK,
 * problèmes / actions, remerciements — avec commentaires automatiques.
 */

import PptxGenJS from "pptxgenjs";
import type { UnitAgg, VaccineAgg } from "./analytics";
import { ANTIGENES, POLIO_AGE_LABELS, RR_AGE_LABELS, SEUIL_PERTE, VACCINE_LABELS, type VaccineKey } from "./parse-masque";
import type { ReportData } from "./report-data";
import { supervisionColor, type ZSSupervision } from "./odk-supervision";

export type { ReportData } from "./report-data";

/* ─── Design tokens (modèle Power BI Bloc 3) ───────────────────────────── */

const NAVY = "0D4A82";
const NAVY_DK = "06457F";
const NAVY_TABLE = "003F7A";
const TEAL = "035D7E";      // RR
const PURPLE = "563C64";    // nVPO2
const INDIGO = "2D3D89";    // VPOb
const MAPI_BLUE = "0E1A77";
const MAROON = "6B2328";
const TITLE_BLUE = "1155CC";
const GREY = "595959";
const GREY_LT = "F2F2F2";
const LINE = "BFBFBF";
const WHITE = "FFFFFF";
const BLACK = "1A1A1A";

const RED = "DF3817";
const YELLOW = "EFF751";
const GREEN = "37D23A";
const BLUE = "0D6ABF";
const NONE = "D9D9D9";

const SUP_GREEN = "59A251";
const SUP_ORANGE = "F38E2C";
const SUP_RED = "E15759";

const AGE_COLORS = ["0D6ABF", "035D7E", "A38600", "5B8C5A"];
const SEX_COLORS = { fille: "50005C", garcon: "6C7BC4" };

const VACC_COLOR: Record<VaccineKey, string> = { rr: TEAL, nvpo2: PURPLE, vpob: INDIGO };

const W = 13.333;
const H = 7.5;
const FONT = "Calibri";
const FONT_TITLE = "Cambria";

/* ─── Helpers ──────────────────────────────────────────────────────────── */

export type AssetLoader = (path: string) => Promise<string | null>;

async function browserLoader(path: string): Promise<string | null> {
  try {
    const res = await fetch(path);
    if (!res.ok) return null;
    const blob = await res.blob();
    return await new Promise((resolve) => {
      const fr = new FileReader();
      fr.onload = () => resolve(fr.result as string);
      fr.onerror = () => resolve(null);
      fr.readAsDataURL(blob);
    });
  } catch {
    return null;
  }
}

function fmtPct(n: number | null | undefined, d = 2, suffix = " %"): string {
  if (n == null || !Number.isFinite(n)) return "—";
  return `${n.toFixed(d).replace(".", ",")}${suffix}`;
}
function fmtNum(n: number | null | undefined, d = 2): string {
  if (n == null || !Number.isFinite(n)) return "—";
  return n.toFixed(d).replace(".", ",");
}
function fmtInt(n: number | null | undefined): string {
  if (n == null || !Number.isFinite(n)) return "0";
  return Math.round(n).toLocaleString("fr-FR").replace(/ /g, " ").replace(/ /g, " ");
}
function r2(n: number | null | undefined): number {
  return n == null || !Number.isFinite(n) ? 0 : Math.round(n * 100) / 100;
}

/** Couleur de seuil couverture / complétude : < 80 rouge, 80‑95 jaune, 95‑100 vert, > 100 bleu. */
function covColor(v: number | null | undefined): string {
  if (v == null || !Number.isFinite(v)) return NONE;
  if (v > 100) return BLUE;
  if (v >= 95) return GREEN;
  if (v >= 80) return YELLOW;
  return RED;
}
/** Couleur de seuil taux de perte : > 10 rouge, 5‑10 jaune, 0‑5 vert, < 0 bleu. */
function lossColor(v: number | null | undefined): string {
  if (v == null || !Number.isFinite(v)) return NONE;
  if (v < 0) return BLUE;
  if (v > 10) return RED;
  if (v > 5) return YELLOW;
  return GREEN;
}
/** Texte lisible sur une pastille colorée. */
function onColor(bg: string): string {
  return bg === YELLOW || bg === NONE || bg === GREEN ? BLACK : WHITE;
}
function slug(s: string): string {
  return s.normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[^a-zA-Z0-9]+/g, "_").replace(/^_|_$/g, "").slice(0, 50) || "rapport";
}
function joinAnd(names: string[]): string {
  if (names.length <= 1) return names.join("");
  return `${names.slice(0, -1).join(", ")} et ${names[names.length - 1]}`;
}
/** « Zone de Santé » → « Zones de Santé ». */
function pluralLabel(label: string): string {
  return label.replace(/^(\S+)/, "$1s");
}
function plural(n: number, s: string, p: string): string {
  return n > 1 ? p : s;
}
function chunk<T>(rows: T[], per: number): T[][] {
  if (per <= 0 || rows.length <= per) return [rows];
  const out: T[][] = [];
  for (let i = 0; i < rows.length; i += per) out.push(rows.slice(i, i + per));
  return out;
}
function suite(idx: number, total: number): string {
  return total <= 1 ? "" : ` (${idx + 1}/${total})`;
}
function colW(total: number, weights: number[]): number[] {
  const s = weights.reduce((a, b) => a + b, 0);
  return weights.map((w) => (w / s) * total);
}

/* ─── Contexte ─────────────────────────────────────────────────────────── */

interface Ctx {
  pptx: PptxGenJS;
  d: ReportData;
  assets: Record<string, string | null>;
}

export interface ExportOptions {
  /** Chargeur d'images (par défaut : fetch navigateur des fichiers /cover/…). */
  loader?: AssetLoader;
  /** "download" (navigateur) ou "buffer" (tests Node). */
  output?: "download" | "buffer";
  fileName?: string;
}

const ASSET_PATHS: Record<string, string> = {
  photo1: "/cover/photo-1.jpg",
  photo2: "/cover/photo-2.jpg",
  minsante: "/cover/logo-minsante.png",
  pev: "/logo/pev-officiel.png",
  partenaires: "/cover/partenaires.png",
  merci: "/cover/merci.png",
};

export async function exportReportPPT(d: ReportData, opts: ExportOptions = {}): Promise<Uint8Array | void> {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: "WIDE", width: W, height: H });
  pptx.layout = "WIDE";
  pptx.author = "PEV — RD Congo";
  pptx.company = "Programme Élargi de Vaccination";
  pptx.title = `${d.titre} — ${d.scopeLabel}`;

  const loader = opts.loader ?? browserLoader;
  const assets: Record<string, string | null> = {};
  await Promise.all(Object.entries(ASSET_PATHS).map(async ([k, p]) => { assets[k] = await loader(p); }));

  const ctx: Ctx = { pptx, d, assets };

  buildCover(ctx);
  buildPlan(ctx);
  buildActionsPC(ctx);
  buildPointsSaillants(ctx);
  buildSynthese(ctx);
  buildCompletude(ctx);
  buildCouverture(ctx, "rr");
  buildCouverture(ctx, "nvpo2");
  buildCouverture(ctx, "vpob");
  buildMapi(ctx);
  buildMapiChart(ctx);
  buildMPV(ctx);
  buildRecup(ctx);
  buildAges(ctx);
  buildSexe(ctx);
  buildSupervision(ctx);
  buildProblemes(ctx);
  buildMerci(ctx);

  const fileName = opts.fileName ?? `Resultats_partiels_RR_Polio_${slug(d.scopeLabel)}_${d.dateMaj.replace(/\//g, "-")}.pptx`;
  if (opts.output === "buffer") {
    const out = await pptx.write({ outputType: "nodebuffer" });
    return out as unknown as Uint8Array;
  }
  await pptx.writeFile({ fileName });
}

/* ─── Éléments communs ─────────────────────────────────────────────────── */

/** Bandeau de titre du modèle : carré navy à gauche + cadre bleu + titre serif navy. */
function header(ctx: Ctx, s: PptxGenJS.Slide, title: string, opts: { h?: number } = {}): void {
  const { pptx } = ctx;
  const h = opts.h ?? 0.62;
  s.background = { color: WHITE };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0.08, w: 0.45, h, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
  s.addShape(pptx.ShapeType.rect, { x: 0.45, y: 0.08, w: W - 0.6, h, fill: { color: WHITE }, line: { color: "2E75B6", width: 1 } });
  const fontSize = title.length > 95 ? 15 : title.length > 78 ? 17 : 20;
  s.addText(title, {
    x: 0.55, y: 0.08, w: W - 0.8, h, fontSize, bold: true, color: NAVY, fontFace: FONT_TITLE, valign: "middle", align: "left",
  });
}

/** Titre bleu vif (diapos supervision, style « accroche »). */
function headerAccroche(ctx: Ctx, s: PptxGenJS.Slide, title: string): void {
  const { pptx } = ctx;
  s.background = { color: WHITE };
  s.addShape(pptx.ShapeType.rect, { x: 0.08, y: 0.22, w: 0.22, h: 0.75, fill: { color: TITLE_BLUE }, line: { color: TITLE_BLUE, width: 0 } });
  s.addShape(pptx.ShapeType.rect, { x: 0.45, y: 0.15, w: W - 0.6, h: 0.9, fill: { color: WHITE }, line: { color: "2E75B6", width: 1 } });
  const fontSize = title.length > 150 ? 16 : title.length > 110 ? 18 : 21;
  s.addText(title, {
    x: 0.55, y: 0.15, w: W - 0.8, h: 0.9, fontSize, bold: true, color: TITLE_BLUE, fontFace: FONT_TITLE, valign: "middle",
  });
}

/** Cartouche « Commentaire : » — colonne de droite ou bandeau bas. */
function commentBox(ctx: Ctx, s: PptxGenJS.Slide, comment: string, pos: { x: number; y: number; w: number; h: number }): void {
  const { pptx } = ctx;
  s.addShape(pptx.ShapeType.rect, { ...pos, fill: { color: WHITE }, line: { color: NAVY, width: 1.25 } });
  const vertical = pos.h > 2;
  s.addText(
    [
      { text: "Commentaire : ", options: { bold: true, fontSize: vertical ? 15 : 13, color: BLACK, breakLine: vertical } },
      { text: comment, options: { fontSize: vertical ? 12 : 12, color: BLACK } },
    ],
    { x: pos.x + 0.1, y: pos.y + 0.05, w: pos.w - 0.2, h: pos.h - 0.1, valign: "top", fontFace: FONT, fit: "shrink" }
  );
}

function sourceLine(s: PptxGenJS.Slide, text = "Source : Masque de saisie de la campagne (Kasaï Central)"): void {
  s.addText(text, { x: 0.1, y: H - 0.3, w: 12, h: 0.25, fontSize: 11, color: TITLE_BLUE, fontFace: FONT_TITLE });
}

/** Barre de titre de bloc (bandeau plein sur fond coloré). */
function blockTitle(ctx: Ctx, s: PptxGenJS.Slide, text: string, x: number, y: number, w: number, color = NAVY, h = 0.36, fontSize = 13): void {
  s.addShape(ctx.pptx.ShapeType.rect, { x, y, w, h, fill: { color }, line: { color, width: 0 } });
  s.addText(text, { x, y, w, h, fontSize, bold: true, color: WHITE, align: "center", valign: "middle", fontFace: FONT });
}

/** Cadre de bloc (contour) autour d'un visuel. */
function frame(ctx: Ctx, s: PptxGenJS.Slide, x: number, y: number, w: number, h: number, color = NAVY): void {
  s.addShape(ctx.pptx.ShapeType.rect, { x, y, w, h, fill: { color: WHITE }, line: { color, width: 1 } });
}

/** Légendes des seuils (barre à cases colorées, style modèle). */
function legendBar(
  ctx: Ctx,
  s: PptxGenJS.Slide,
  x: number,
  y: number,
  w: number,
  title: string,
  items: { c: string; t: string }[],
  titleColor = NAVY,
  h = 0.27
): void {
  const { pptx } = ctx;
  s.addShape(pptx.ShapeType.rect, { x, y, w, h, fill: { color: titleColor }, line: { color: titleColor, width: 0 } });
  s.addText(title, { x, y, w, h, fontSize: 11, bold: true, color: WHITE, align: "center", valign: "middle", fontFace: FONT });
  const cw = w / items.length;
  items.forEach((it, i) => {
    s.addShape(pptx.ShapeType.rect, { x: x + i * cw, y: y + h, w: cw, h, fill: { color: it.c }, line: { color: WHITE, width: 0.5 } });
    s.addText(it.t, { x: x + i * cw, y: y + h, w: cw, h, fontSize: 11, bold: true, color: onColor(it.c), align: "center", valign: "middle", fontFace: FONT });
  });
}
const LEG_COV = [{ c: RED, t: "0%-80%" }, { c: YELLOW, t: "80-95%" }, { c: GREEN, t: "95-100%" }, { c: BLUE, t: "> 100 %" }];
const LEG_LOSS = [{ c: RED, t: ">10%" }, { c: YELLOW, t: "5%-10%" }, { c: GREEN, t: "0%-5%" }, { c: BLUE, t: "<0%" }];

function legendsBottom(ctx: Ctx, s: PptxGenJS.Slide, y = 6.85, x = 0.4, w = W - 0.8, cov = true, loss = true, compl = false): void {
  const parts = [compl, cov, loss].filter(Boolean).length;
  const gap = 0.08;
  const bw = (w - gap * (parts - 1)) / parts;
  let cx = x;
  if (compl) { legendBar(ctx, s, cx, y, bw, "Critères de graduation des complétudes", LEG_COV); cx += bw + gap; }
  if (cov) { legendBar(ctx, s, cx, y, bw, "Critères des couvertures vaccinales", LEG_COV); cx += bw + gap; }
  if (loss) { legendBar(ctx, s, cx, y, bw, "Critères des taux de perte", LEG_LOSS); }
}

/* Cellules de tableau */
type CellOpts = Partial<PptxGenJS.TableCellProps>;
function th(fill = NAVY, extra: CellOpts = {}): PptxGenJS.TableCellProps {
  return { bold: true, color: WHITE, fill: { color: fill }, fontSize: 10, align: "center", valign: "middle", fontFace: FONT, margin: [2, 3, 2, 3], ...extra };
}
function td(extra: CellOpts = {}): PptxGenJS.TableCellProps {
  return { color: BLACK, fontSize: 10, valign: "middle", fontFace: FONT, fill: { color: WHITE }, margin: [2, 3, 2, 3], ...extra };
}
function tdNum(extra: CellOpts = {}): PptxGenJS.TableCellProps {
  return td({ align: "right", ...extra });
}
function tdColored(v: number | null, kind: "cov" | "loss", extra: CellOpts = {}): PptxGenJS.TableCellProps {
  const c = kind === "cov" ? covColor(v) : lossColor(v);
  return td({ align: "right", bold: true, fill: { color: c }, color: onColor(c), ...extra });
}
function ttotal(fill = NAVY_DK, extra: CellOpts = {}): PptxGenJS.TableCellProps {
  return { bold: true, color: WHITE, fill: { color: fill }, fontSize: 10, align: "right", valign: "middle", fontFace: FONT, margin: [2, 3, 2, 3], ...extra };
}
function zebra(i: number): CellOpts {
  return i % 2 === 1 ? { fill: { color: GREY_LT } } : {};
}

const TABLE_BORDER: PptxGenJS.TableProps["border"] = { type: "solid", color: LINE, pt: 0.5 };

/** Graphique en colonnes verticales avec une couleur par barre. */
function colChart(
  ctx: Ctx,
  s: PptxGenJS.Slide,
  pos: { x: number; y: number; w: number; h: number },
  labels: string[],
  values: number[],
  colors: string[],
  o: { valAxisTitle?: string; catAxisTitle?: string; fmt?: string; max?: number; min?: number; fontSize?: number; labelFont?: number; barDir?: "col" | "bar"; showAxis?: boolean; gap?: number } = {}
): void {
  if (labels.length === 0) return;
  // PowerPoint trace la première catégorie en bas des barres horizontales : on
  // inverse pour lire du haut vers le bas dans l'ordre fourni.
  if ((o.barDir ?? "col") === "bar") {
    labels = [...labels].reverse();
    values = [...values].reverse();
    colors = [...colors].reverse();
  }
  const cols = colors.length > 1 ? colors : [...colors, ...colors, NAVY];
  const vmin = Math.min(0, ...values);
  const vmax = Math.max(0, ...values);
  const autoMin = vmin < 0 ? Math.floor(vmin / 10) * 10 - 5 : 0;
  const autoMax = vmax <= 0 ? 10 : Math.ceil((vmax * 1.12) / 10) * 10;
  s.addChart(
    ctx.pptx.ChartType.bar,
    [{ name: "Valeur", labels, values }],
    {
      ...pos,
      barDir: o.barDir ?? "col",
      chartColors: cols,
      showValue: true,
      dataLabelPosition: "outEnd",
      dataLabelFormatCode: o.fmt ?? "0.00",
      dataLabelFontSize: o.labelFont ?? 9,
      dataLabelColor: "404040",
      dataLabelFontBold: true,
      catAxisLabelFontSize: o.fontSize ?? 9,
      catAxisLabelFontBold: true,
      catAxisLabelColor: "262626",
      valAxisLabelFontSize: 9,
      valAxisMinVal: o.min ?? autoMin,
      valAxisMaxVal: o.max ?? autoMax,
      valAxisHidden: o.showAxis === false,
      showLegend: false,
      showTitle: false,
      valGridLine: { style: "dash", color: "E0E0E0", size: 0.5 },
      catGridLine: { style: "none" },
      barGapWidthPct: o.gap ?? 40,
      showValAxisTitle: Boolean(o.valAxisTitle),
      valAxisTitle: o.valAxisTitle,
      valAxisTitleFontSize: 9,
      showCatAxisTitle: Boolean(o.catAxisTitle),
      catAxisTitle: o.catAxisTitle,
      catAxisTitleFontSize: 9,
      catAxisLabelRotate: labels.length > 12 && (o.barDir ?? "col") === "col" ? -45 : 0,
    }
  );
}

/* ─── 1. Page de garde ─────────────────────────────────────────────────── */

function buildCover(ctx: Ctx): void {
  const { pptx, d, assets } = ctx;
  const s = pptx.addSlide();
  s.background = { color: WHITE };
  if (assets.minsante) s.addImage({ data: assets.minsante, x: 0.12, y: 0.12, w: 2.9, h: 0.71 });
  if (assets.pev) s.addImage({ data: assets.pev, x: 4.55, y: 0.25, w: 3.7, h: 0.567 });
  s.addText("CAMPAGNE DE VACCINATION INTEGREE\nCONTRE LA ROUGEOLE-RUBÉOLE ET POLIO", {
    x: 0.6, y: 1.05, w: 7.5, h: 0.95, fontSize: 19, bold: true, color: "C00000", fontFace: "Times New Roman", align: "center", valign: "middle",
  });
  s.addShape(pptx.ShapeType.rect, { x: 0.6, y: 2.4, w: 7.15, h: 2.5, fill: { color: WHITE }, line: { color: "2E75B6", width: 1.5 } });
  s.addText(`Résultats partiels de la campagne intégrée Rougeole-Rubéole et Poliomyélite`, {
    x: 0.65, y: 2.45, w: 7.05, h: 1.55, fontSize: 32, bold: true, color: TITLE_BLUE, fontFace: FONT, valign: "top", fit: "shrink",
  });
  s.addText(`${d.coverEntity}${d.periode ? ` — ${d.periode}` : ""}`, {
    x: 0.65, y: 4.05, w: 7.05, h: 0.75, fontSize: 18, bold: true, color: NAVY, fontFace: FONT, valign: "middle", fit: "shrink",
  });
  s.addText(`Mise à jour du ${d.dateMaj}`, {
    x: 0.6, y: 5.25, w: 7.2, h: 0.55, fontSize: 28, bold: true, color: BLACK, fontFace: "Times New Roman", align: "center",
  });
  s.addShape(pptx.ShapeType.line, { x: 8.42, y: 0.05, w: 0, h: 6.07, line: { color: TITLE_BLUE, width: 3 } });
  if (assets.photo1) s.addImage({ data: assets.photo1, x: 9.19, y: 0.13, w: 3.9, h: 2.78, sizing: { type: "cover", w: 3.9, h: 2.78 } });
  if (assets.photo2) s.addImage({ data: assets.photo2, x: 9.19, y: 3.05, w: 3.9, h: 2.98, sizing: { type: "cover", w: 3.9, h: 2.98 } });
  if (assets.partenaires) s.addImage({ data: assets.partenaires, x: 0.05, y: 6.12, w: W - 0.1, h: (W - 0.1) / 9.6 });
}

/* ─── 2. Plan ──────────────────────────────────────────────────────────── */

function buildPlan(ctx: Ctx): void {
  const { pptx, d } = ctx;
  const s = pptx.addSlide();
  header(ctx, s, "Résultats Partiels : Synthèse des principaux indicateurs de suivi de la campagne", { h: 0.77 });
  const jour = d.jourLabels.length ? d.jourLabels[d.jourLabels.length - 1] : "J1";
  const items = [
    "Suivi des actions du précédent poste de commandement",
    "Points saillants",
    `Résultats partiels de la campagne au ${jour === "Ratissage" ? "ratissage" : `jour ${jour.replace("J", "")}`}`,
    "Complétude, couvertures vaccinales et taux de perte (RR, nVPO2, VPOb)",
    "MAPI, surveillance des MPV et récupération en PEV de routine",
    "Vaccinés par tranche d'âge et par sexe",
    "Supervision des équipes de vaccination (ODK)",
    "Problèmes rencontrés / Actions correctrices",
  ];
  s.addText(
    items.map((t) => ({ text: t, options: { bullet: { code: "25AA" }, breakLine: true, paraSpaceAfter: 10 } })),
    { x: 0.7, y: 1.2, w: 12, h: 5.5, fontSize: 24, color: BLACK, fontFace: FONT, valign: "top" }
  );
}

/* ─── 3. Suivi des points d'action ─────────────────────────────────────── */

function statutColor(st: string): string {
  const t = st.toLowerCase();
  if (/r[ée]alis|termin|fait|achev|clos/.test(t)) return "D9EAD3";
  if (/retard|non r[ée]alis|bloqu/.test(t)) return "F4CCCC";
  if (/cours|partiel|en attente/.test(t)) return "FCE5CD";
  return WHITE;
}

function buildActionsPC(ctx: Ctx): void {
  const { pptx, d } = ctx;
  const rows = d.actionsPC.filter((r) => r.activite.trim() || r.responsable.trim());
  const head: PptxGenJS.TableRow = ["Activités prévues", "Responsables", "Statut", "Échéance"].map((t) => ({
    text: t, options: td({ bold: true, fontSize: 16, align: "center", fontFace: FONT_TITLE, margin: [6, 6, 6, 6] }),
  }));
  const pages = chunk(rows, 5);
  pages.forEach((page, idx) => {
    const s = pptx.addSlide();
    header(ctx, s, `Suivi des points d'action des précédents PC${suite(idx, pages.length)}`, { h: 0.81 });
    const body: PptxGenJS.TableRow[] = page.length
      ? page.map((r) => [
          { text: r.activite, options: td({ fontSize: 14, fontFace: FONT_TITLE, margin: [8, 8, 8, 8] }) },
          { text: r.responsable, options: td({ fontSize: 14, fontFace: FONT_TITLE, align: "center" }) },
          { text: r.statut, options: td({ fontSize: 14, fontFace: FONT_TITLE, align: "center", fill: { color: statutColor(r.statut) } }) },
          { text: r.echeance, options: td({ fontSize: 14, bold: true, fontFace: FONT_TITLE, align: "center" }) },
        ])
      : [[{ text: "Aucun point d'action reporté du précédent poste de commandement (à compléter dans l'application avant génération).", options: td({ fontSize: 14, italic: true, color: GREY, colspan: 4, align: "center", margin: [10, 10, 10, 10] }) }]];
    s.addTable([head, ...body], {
      x: 0.48, y: 1.17, w: W - 0.96, colW: colW(W - 0.96, [5.5, 2.3, 2.3, 2.5]),
      border: { type: "solid", color: LINE, pt: 0.5 }, rowH: page.length ? 1.05 : 0.6, valign: "middle", autoPage: false,
    });
  });
}

/* ─── 4. Points saillants ──────────────────────────────────────────────── */

function buildPointsSaillants(ctx: Ctx): void {
  const { pptx, d } = ctx;
  const t = d.total;
  const s = pptx.addSlide();
  s.background = { color: WHITE };
  s.addText("Points saillants", { x: 0.55, y: 0.3, w: 8, h: 0.5, fontSize: 28, bold: true, color: TITLE_BLUE, fontFace: FONT });
  s.addText(`Synthèse de la performance de la campagne intégrée — ${d.scopeLabel}`, { x: 0.57, y: 0.78, w: 10, h: 0.34, fontSize: 14, bold: true, color: BLACK, fontFace: FONT });
  s.addShape(pptx.ShapeType.rect, { x: 0.55, y: 1.15, w: 12.2, h: 0.03, fill: { color: TITLE_BLUE }, line: { color: TITLE_BLUE, width: 0 } });

  const kpi = (x: number, label: string, value: string, sub: string, valueColor = "C00000") => {
    s.addShape(pptx.ShapeType.roundRect, { x, y: 1.45, w: 2.85, h: 1.15, fill: { color: GREY_LT }, line: { color: "D0D0D0", width: 0.75 }, rectRadius: 0.15, shadow: { type: "outer", blur: 4, offset: 2, angle: 45, color: "9E9E9E", opacity: 0.35 } });
    s.addText(label, { x: x + 0.14, y: 1.52, w: 2.6, h: 0.4, fontSize: 16, bold: true, color: BLACK, fontFace: FONT });
    s.addText(value, { x: x + 0.18, y: 1.88, w: 2.6, h: 0.47, fontSize: 22, bold: true, color: valueColor, fontFace: FONT });
    s.addText(sub, { x: x + 0.18, y: 2.3, w: 2.6, h: 0.28, fontSize: 12, bold: true, color: BLACK, fontFace: FONT });
  };
  const jl = d.jourLabels;
  const jspan = jl.length ? (jl.length === 1 ? `Rapports ${jl[0]}` : `Rapports ${jl[0]}–${jl[jl.length - 1]}`) : "Rapports";
  kpi(0.55, "COMPLÉTUDE", fmtPct(t.completude, 2), jspan);
  kpi(3.63, "CV RR", fmtPct(t.rr.cv, 2, "%"), "Vaccination");
  kpi(6.71, "CV nVPO2", fmtPct(t.nvpo2.cv, 2, "%"), "Vaccination");
  kpi(9.79, "CV VPOb", fmtPct(t.vpob.cv, 2, "%"), "Vaccination");

  // Bloc VACCINATION
  const box = (x: number, w: number) => s.addShape(pptx.ShapeType.roundRect, { x, y: 2.78, w, h: 3.94, fill: { color: GREY_LT }, line: { color: "D0D0D0", width: 0.75 }, rectRadius: 0.35, shadow: { type: "outer", blur: 5, offset: 2, angle: 45, color: "9E9E9E", opacity: 0.35 } });
  box(0.55, 6.0);
  s.addText("VACCINATION", { x: 0.82, y: 3.05, w: 3, h: 0.4, fontSize: 18, bold: true, color: BLACK, fontFace: FONT });
  const vaccRow = (y: number, label: string, v: VaccineAgg) => {
    s.addText(label, { x: 0.6, y, w: 1.05, h: 0.44, fontSize: 20, bold: true, color: NAVY, fontFace: FONT });
    s.addText("Vaccinés", { x: 1.5, y: y + 0.03, w: 1.16, h: 0.37, fontSize: 16, color: GREY, fontFace: FONT });
    s.addText(fmtInt(v.vacc), { x: 2.5, y, w: 1.5, h: 0.44, fontSize: 18, bold: true, color: "7B2C2C", fontFace: FONT });
    s.addText("Cible", { x: 4.0, y: y + 0.03, w: 0.8, h: 0.37, fontSize: 16, color: GREY, fontFace: FONT });
    s.addText(fmtInt(v.cible), { x: 4.85, y, w: 1.6, h: 0.44, fontSize: 18, bold: true, color: BLACK, fontFace: FONT });
  };
  vaccRow(3.7, "RR", t.rr);
  vaccRow(4.5, "nVPO2", t.nvpo2);
  vaccRow(5.3, "VPOb", t.vpob);

  // Bloc GESTION & MAPI
  box(6.8, 5.95);
  s.addText("GESTION DES VACCINS & MAPI", { x: 7.07, y: 3.05, w: 5, h: 0.4, fontSize: 18, bold: true, color: BLACK, fontFace: FONT });
  const gestRow = (y: number, label: string, v: VaccineAgg) => {
    s.addText(label, { x: 6.9, y, w: 1.0, h: 0.44, fontSize: 20, bold: true, color: NAVY, fontFace: FONT });
    s.addText("Flacons utilisés", { x: 7.8, y: y + 0.03, w: 2.0, h: 0.37, fontSize: 16, color: GREY, fontFace: FONT });
    s.addText(fmtInt(v.flaconsUtil), { x: 9.45, y, w: 1.45, h: 0.44, fontSize: 18, bold: true, color: "7B2C2C", fontFace: FONT });
    s.addText("Perte", { x: 11.0, y: y + 0.03, w: 0.76, h: 0.37, fontSize: 16, color: GREY, fontFace: FONT });
    const lc = lossColor(v.tauxPerte);
    s.addText(fmtPct(v.tauxPerte, 2, "%"), { x: 11.72, y, w: 1.05, h: 0.44, fontSize: 16, bold: true, color: lc === YELLOW ? "B8860B" : lc === NONE ? GREY : lc, fontFace: FONT });
  };
  gestRow(3.7, "RR", t.rr);
  gestRow(4.5, "nVPO2", t.nvpo2);
  gestRow(5.3, "VPOb", t.vpob);
  s.addShape(pptx.ShapeType.roundRect, { x: 7.08, y: 5.85, w: 5.34, h: 0.77, fill: { color: "FFF4E6" }, line: { color: "F4B183", width: 1 }, rectRadius: 0.1 });
  s.addText(
    [
      { text: "MAPI :   ", options: { bold: true, color: "C00000", fontSize: 16 } },
      { text: `${fmtInt(t.mapiNonGraves)} non graves   |   `, options: { bold: true, color: "C00000", fontSize: 16 } },
      { text: `${fmtInt(t.mapiGraves)} graves`, options: { bold: true, color: "C00000", fontSize: 16 } },
    ],
    { x: 7.17, y: 5.86, w: 5.25, h: 0.4, valign: "middle", fontFace: FONT }
  );
  s.addText(`Proportion MAPI :  ${fmtNum(t.mapiPour100k, 2)}  MAPI / 100 000 doses`, { x: 7.39, y: 6.2, w: 4.9, h: 0.34, fontSize: 14, bold: true, color: BLACK, fontFace: FONT });

  // Légendes bas de page (couleurs du modèle)
  const legC = [{ c: "D9534F", t: "0%-80%" }, { c: "F0C419", t: "80-95%" }, { c: "50B266", t: "95-100%" }, { c: "317EB8", t: "> 100 %" }];
  const legL = [{ c: "D9534F", t: ">10%" }, { c: "F0C419", t: "5%-10%" }, { c: "50B266", t: "0%-5%" }, { c: "317EB8", t: "<0%" }];
  legendBar(ctx, s, 0.54, 6.85, 4.05, "Critères de graduation des Complétudes", legC, NAVY, 0.3);
  legendBar(ctx, s, 4.61, 6.85, 4.06, "Critères des couvertures vaccinales", legC, NAVY, 0.3);
  legendBar(ctx, s, 8.67, 6.85, 4.07, "Critères des taux de perte", legL, NAVY, 0.3);
}

/* ─── 5. Synthèse des principaux indicateurs ───────────────────────────── */

function syntheseRows(units: UnitAgg[], label: string, dateLancement: string, total: UnitAgg): PptxGenJS.TableRow[] {
  const head: PptxGenJS.TableRow = [
    label, "Date\nlancement", "Complé-\ntude", "Vaccinés\nnVPO2", "Vaccinés\nVPOb", "Vaccinés RR", "Couvert.\nRR%", "Couvert.\nnVPO2%", "Couvert.\nVPOb%",
    "Taux\nPerte RR", "Taux Perte\nnVPO2", "Taux Perte\nVPOb", "MAPI non\ngrave", "MAPI Grave",
  ].map((t, i) => ({ text: t, options: th(NAVY, { align: i === 0 ? "left" : "center", fontSize: 9 }) }));
  const row = (u: UnitAgg, i: number, isTotal = false): PptxGenJS.TableRow => {
    const base = isTotal ? ttotal(NAVY_DK) : tdNum(zebra(i));
    const name = isTotal ? ttotal(NAVY_DK, { align: "left" }) : td({ bold: true, ...zebra(i) });
    const gr = (v: number) => (isTotal ? ttotal(NAVY_DK) : tdNum({ bold: true, fill: { color: v > 0 ? RED : (i % 2 ? GREY_LT : WHITE) }, color: v > 0 ? WHITE : BLACK }));
    return [
      { text: u.unit, options: name },
      { text: dateLancement || "—", options: isTotal ? ttotal(NAVY_DK) : td({ align: "center", ...zebra(i) }) },
      { text: fmtPct(u.completude, 2), options: isTotal ? ttotal(NAVY_DK) : tdColored(u.completude, "cov", { align: "center" }) },
      { text: fmtInt(u.nvpo2.vacc), options: base },
      { text: fmtInt(u.vpob.vacc), options: base },
      { text: fmtInt(u.rr.vacc), options: base },
      { text: fmtNum(u.rr.cv), options: isTotal ? ttotal(NAVY_DK) : tdColored(u.rr.cv, "cov") },
      { text: fmtNum(u.nvpo2.cv), options: isTotal ? ttotal(NAVY_DK) : tdColored(u.nvpo2.cv, "cov") },
      { text: fmtNum(u.vpob.cv), options: isTotal ? ttotal(NAVY_DK) : tdColored(u.vpob.cv, "cov") },
      { text: fmtNum(u.rr.tauxPerte), options: isTotal ? ttotal(NAVY_DK) : tdColored(u.rr.tauxPerte, "loss") },
      { text: fmtNum(u.nvpo2.tauxPerte), options: isTotal ? ttotal(NAVY_DK) : tdColored(u.nvpo2.tauxPerte, "loss") },
      { text: fmtNum(u.vpob.tauxPerte), options: isTotal ? ttotal(NAVY_DK) : tdColored(u.vpob.tauxPerte, "loss") },
      { text: fmtInt(u.mapiNonGraves), options: base },
      { text: fmtInt(u.mapiGraves), options: gr(u.mapiGraves) },
    ];
  };
  return [head, ...units.map((u, i) => row(u, i)), row(total, 0, true)];
}

function buildSynthese(ctx: Ctx): void {
  const { pptx, d } = ctx;
  const draw = (rowsAll: PptxGenJS.TableRow[], label: string, perPage: number) => {
    const head = rowsAll[0];
    const body = rowsAll.slice(1, -1);
    const totalRow = rowsAll[rowsAll.length - 1];
    const pages = chunk(body, perPage);
    pages.forEach((page, idx) => {
      const s = pptx.addSlide();
      header(ctx, s, `Résultats partiels : Synthèse des principaux indicateurs de suivi de la campagne par ${label}${suite(idx, pages.length)}`, { h: 0.77 });
      frame(ctx, s, 0.15, 0.95, W - 0.3, 5.75);
      blockTitle(ctx, s, `Synthèse des principaux indicateurs de suivi de la campagne par ${label}`, 0.15, 0.95, W - 0.3, NAVY, 0.4, 14);
      const rows = idx === pages.length - 1 ? [head, ...page, totalRow] : [head, ...page];
      s.addTable(rows, {
        x: 0.2, y: 1.42, w: W - 0.4,
        colW: colW(W - 0.4, [1.6, 1.0, 0.85, 0.9, 0.9, 0.95, 0.85, 0.85, 0.85, 0.85, 0.85, 0.85, 0.85, 0.85]),
        border: TABLE_BORDER, rowH: 0.3, valign: "middle", autoPage: false,
      });
      const tableBottom = 1.42 + 0.31 * (rows.length + 0.6);
      if (6.65 - tableBottom > 1.0) {
        commentBox(ctx, s, syntheseComment(d, label), { x: 0.3, y: tableBottom + 0.15, w: W - 0.6, h: Math.min(2.6, 6.6 - tableBottom - 0.15) });
      }
      // Date de dernière actualisation + légendes (style modèle)
      s.addShape(pptx.ShapeType.rect, { x: 0.15, y: 6.78, w: 3.5, h: 0.6, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
      s.addText([
        { text: "DATE DERNIERE ACTUALISATION", options: { fontSize: 12, bold: true, color: WHITE, breakLine: true } },
        { text: d.dateMaj, options: { fontSize: 15, bold: true, color: WHITE } },
      ], { x: 0.2, y: 6.78, w: 3.4, h: 0.6, valign: "middle", fontFace: FONT });
      legendsBottom(ctx, s, 6.78, 3.75, W - 3.9, true, true, true);
    });
  };
  if (d.antennes.length > 1) draw(syntheseRows(d.antennes, "Antenne", d.dateLancement, d.total), "Antenne", 20);
  draw(syntheseRows(d.units, d.byUnitLabel, d.dateLancement, d.total), d.byUnitLabel, 15);
}

function syntheseComment(d: ReportData, label: string): string {
  const t = d.total;
  const started = t.vaccRecus > 0 || t.rr.vacc > 0;
  if (!started) {
    return `Aucun rapport de vaccination n'est encore saisi pour ${d.scopeLabel.toLowerCase()} (lancement prévu le ${d.dateLancement || "—"}). Les indicateurs (complétude, couvertures, taux de perte, MAPI) s'actualiseront automatiquement dès la remontée des premiers rapports journaliers.`;
  }
  const rows = label === "Antenne" ? d.antennes : d.units;
  const bestRR = [...rows].filter((u) => u.rr.vacc > 0).sort((a, b) => (b.rr.cv ?? 0) - (a.rr.cv ?? 0));
  const worstRR = [...bestRR].reverse();
  const lossHigh = rows.filter((u) => ["rr", "nvpo2", "vpob"].some((k) => (u[k as VaccineKey].tauxPerte ?? 0) > 10)).map((u) => u.unit);
  const parts: string[] = [
    `Au ${d.jourLabels[d.jourLabels.length - 1] ?? "J1"}, la complétude des rapports atteint ${fmtPct(t.completude)} ; les couvertures s'établissent à ${fmtPct(t.rr.cv)} pour le RR, ${fmtPct(t.nvpo2.cv)} pour le nVPO2 et ${fmtPct(t.vpob.cv)} pour le VPOb.`,
  ];
  if (bestRR.length) parts.push(`Meilleure couverture RR : ${bestRR[0].unit} (${fmtPct(bestRR[0].rr.cv, 1)}) ; la plus faible : ${worstRR[0].unit} (${fmtPct(worstRR[0].rr.cv, 1)}).`);
  parts.push(`Taux de perte : RR ${fmtPct(t.rr.tauxPerte)}, nVPO2 ${fmtPct(t.nvpo2.tauxPerte)}, VPOb ${fmtPct(t.vpob.tauxPerte)}${lossHigh.length ? ` — pertes > 10 % à surveiller : ${joinAnd(lossHigh.slice(0, 5))}${lossHigh.length > 5 ? "…" : ""}` : ""}.`);
  parts.push(`MAPI : ${fmtInt(t.mapiNonGraves)} non graves et ${fmtInt(t.mapiGraves)} graves (${fmtNum(t.mapiPour100k)} pour 100 000 doses).`);
  return parts.join(" ");
}

/* ─── 6. Complétude ────────────────────────────────────────────────────── */

function completudeComment(d: ReportData, units: UnitAgg[], label: string): string {
  const t = d.total;
  const started = units.filter((u) => u.vaccRecus > 0);
  if (started.length === 0) {
    return `Aucun rapport de vaccination n'est encore saisi pour ${d.scopeLabel.toLowerCase()} : la complétude est de 0 %. La saisie des rapports journaliers doit démarrer dès le lancement (${d.dateLancement || "date à confirmer"}).`;
  }
  const bad = units.filter((u) => u.vaccAttendus > 0 && (u.completude ?? 0) < 80).sort((a, b) => (a.completude ?? 0) - (b.completude ?? 0));
  const good = units.filter((u) => (u.completude ?? 0) >= 95);
  const lastIdx = d.jourLabels.length - 1;
  const lastJ = d.jourLabels[lastIdx];
  const lastRate = t.attendusDaily[lastIdx] > 0 ? (t.recusDaily[lastIdx] / t.attendusDaily[lastIdx]) * 100 : null;
  let txt = `Complétude globale de ${fmtPct(t.completude)} (${fmtInt(t.vaccRecus)} rapports reçus sur ${fmtInt(t.vaccAttendus)} attendus)`;
  if (lastRate != null) txt += ` ; ${lastJ} : ${fmtPct(lastRate)}`;
  txt += ". ";
  if (good.length) txt += `${good.length} ${plural(good.length, label.toLowerCase(), label.toLowerCase().replace(/^(\S+)/, "$1s"))} ≥ 95 %. `;
  if (bad.length) {
    txt += `${bad.length} ${plural(bad.length, "reste", "restent")} sous 80 % : ${joinAnd(bad.slice(0, 5).map((u) => `${u.unit} (${fmtPct(u.completude, 1)})`))}${bad.length > 5 ? ` et ${bad.length - 5} autres` : ""}. Relancer la saisie des rapports journaliers dans DHIS2/masque et le feedback du poste de commandement.`;
  } else {
    txt += "Aucune unité sous 80 % : maintenir la revue journalière des données.";
  }
  return txt;
}

function completudeTable(units: UnitAgg[], label: string, jours: string[], total: UnitAgg): PptxGenJS.TableRow[] {
  const head: PptxGenJS.TableRow = [
    { text: label, options: th(NAVY_TABLE, { align: "left" }) },
    { text: "Attendus\npar jour", options: th(NAVY_TABLE) },
    ...jours.flatMap((j) => [
      { text: `Reçu_${j}`, options: th(NAVY_TABLE) },
      { text: `Complét._${j}`, options: th(NAVY_TABLE) },
    ]),
    { text: "Reçus\ntotal", options: th(NAVY_TABLE) },
    { text: "Complét.\nglobale", options: th(NAVY_TABLE) },
  ];
  const row = (u: UnitAgg, i: number, isTotal = false): PptxGenJS.TableRow => {
    const perDay = u.attendusDaily[0] || u.vaccAttendus / Math.max(1, jours.length);
    const cells: PptxGenJS.TableRow = [
      { text: u.unit, options: isTotal ? ttotal(NAVY_DK, { align: "left" }) : td({ bold: true, ...zebra(i) }) },
      { text: fmtInt(perDay), options: isTotal ? ttotal(NAVY_DK) : tdNum({ bold: true, ...zebra(i) }) },
    ];
    jours.forEach((_, k) => {
      const att = u.attendusDaily[k] ?? 0;
      const rec = u.recusDaily[k] ?? 0;
      const c = att > 0 ? (rec / att) * 100 : null;
      cells.push({ text: fmtInt(rec), options: isTotal ? ttotal(NAVY_DK) : tdNum(zebra(i)) });
      cells.push({ text: fmtNum(c), options: isTotal ? ttotal(NAVY_DK) : tdColored(c, "cov") });
    });
    cells.push({ text: fmtInt(u.vaccRecus), options: isTotal ? ttotal(NAVY_DK) : tdNum(zebra(i)) });
    cells.push({ text: fmtNum(u.completude), options: isTotal ? ttotal(NAVY_DK) : tdColored(u.completude, "cov") });
    return cells;
  };
  return [head, ...units.map((u, i) => row(u, i)), row(total, 0, true)];
}

function buildCompletude(ctx: Ctx): void {
  const { pptx, d } = ctx;
  const jours = d.jourLabels;
  const nCols = 4 + jours.length * 2;
  const weights = [1.5, 0.8, ...jours.flatMap(() => [0.7, 0.75]), 0.75, 0.8];

  // A. Par antenne : tableau + graphique sur la même diapo (modèle « par province »).
  if (d.antennes.length > 1) {
    const s = pptx.addSlide();
    header(ctx, s, "Résultats Partiels : Complétude de saisie des rapports journalières et globale par Antenne", { h: 0.49 });
    frame(ctx, s, 0.05, 0.62, 9.75, 6.75);
    blockTitle(ctx, s, "Complétudes journalières de saisie des rapports des AS par Antenne", 0.05, 0.62, 9.75, NAVY_TABLE, 0.38, 13);
    const rows = completudeTable(d.antennes, "Antenne", jours, d.total);
    s.addTable(rows, { x: 0.1, y: 1.05, w: 9.65, colW: colW(9.65, weights.slice(0, nCols)), border: TABLE_BORDER, rowH: 0.28, valign: "middle", autoPage: false });
    const chartY = 1.05 + 0.3 * (rows.length + 1) + 0.15;
    blockTitle(ctx, s, "Complétudes globales de saisie des rapports journaliers des Aires de Santé par Antenne", 0.05, chartY, 9.75, NAVY_TABLE, 0.36, 13);
    const sorted = [...d.antennes].sort((a, b) => (b.completude ?? -1) - (a.completude ?? -1));
    colChart(ctx, s, { x: 0.1, y: chartY + 0.4, w: 9.65, h: 6.75 - (chartY - 0.62) - 0.95 }, sorted.map((u) => u.unit), sorted.map((u) => r2(u.completude)), sorted.map((u) => covColor(u.completude)), { valAxisTitle: "Complétude données", catAxisTitle: "Antenne", fmt: '0.00" %"', max: 100, min: 0 });
    legendBar(ctx, s, 0.05, 6.82, 9.75, "CRITERES DE GRADUATION DES COMPLETUDES", LEG_COV, NAVY_TABLE, 0.26);
    commentBox(ctx, s, completudeComment(d, d.antennes, "Antenne"), { x: 9.87, y: 0.62, w: 3.37, h: 6.75 });
  }

  // B. Par unité (ZS/AS) : tableau journalier paginé.
  const rowsAll = completudeTable(d.units, d.byUnitLabel, jours, d.total);
  const head = rowsAll[0];
  const body = rowsAll.slice(1, -1);
  const totalRow = rowsAll[rowsAll.length - 1];
  const pages = chunk(body, 17);
  pages.forEach((page, idx) => {
    const s = pptx.addSlide();
    header(ctx, s, `Résultats Partiels : Complétude de saisie des rapports journalières et globale par ${d.byUnitLabel}${suite(idx, pages.length)}`, { h: 0.49 });
    frame(ctx, s, 0.05, 0.62, 9.75, 6.75);
    blockTitle(ctx, s, `Complétudes journalières de saisie des rapports des AS par ${d.byUnitLabel}`, 0.05, 0.62, 9.75, NAVY_TABLE, 0.38, 13);
    const rows = idx === pages.length - 1 ? [head, ...page, totalRow] : [head, ...page];
    s.addTable(rows, { x: 0.1, y: 1.05, w: 9.65, colW: colW(9.65, weights.slice(0, nCols)), border: TABLE_BORDER, rowH: 0.28, valign: "middle", autoPage: false });
    legendBar(ctx, s, 0.05, 6.82, 9.75, "CRITERES DE GRADUATION DES COMPLETUDES", LEG_COV, NAVY_TABLE, 0.26);
    commentBox(ctx, s, completudeComment(d, d.units, d.byUnitLabel), { x: 9.87, y: 0.62, w: 3.37, h: 6.75 });
  });

  // C. Graphique de complétude globale par unité (modèle « par antenne »).
  {
    const s = pptx.addSlide();
    header(ctx, s, `Résultats Partiels : Complétude globale de saisie des rapports des AS par ${d.byUnitLabel}`, { h: 0.77 });
    frame(ctx, s, 0.05, 0.9, W - 0.1, 5.75);
    blockTitle(ctx, s, `Complétude de transmission de rapports des Aires de Santé par ${d.byUnitLabel}`, 0.05, 0.9, W - 0.1, NAVY_TABLE, 0.42, 14);
    const sorted = [...d.units].sort((a, b) => (b.completude ?? -1) - (a.completude ?? -1));
    colChart(ctx, s, { x: 0.1, y: 1.35, w: W - 0.2, h: 4.75 }, sorted.map((u) => u.unit), sorted.map((u) => r2(u.completude)), sorted.map((u) => covColor(u.completude)), { valAxisTitle: "Complétude de données", catAxisTitle: pluralLabel(d.byUnitLabel), fmt: '0.00" %"', max: 100, min: 0, fontSize: sorted.length > 20 ? 8 : 9, labelFont: sorted.length > 20 ? 7 : 9 });
    legendBar(ctx, s, 0.05, 6.1, W - 0.1, "CRITERES DE GRADUATION DES COMPLETUDES", LEG_COV, NAVY_TABLE, 0.27);
    commentBox(ctx, s, completudeComment(d, d.units, d.byUnitLabel), { x: 0.05, y: 6.72, w: W - 0.1, h: 0.5 });
    sourceLine(s);
  }
}

/* ─── 7‑9. Couverture vaccinale + taux de perte ────────────────────────── */

function couvComment(d: ReportData, k: VaccineKey, units: UnitAgg[], label: string): string {
  const t = d.total[k];
  const name = VACCINE_LABELS[k];
  const seuil = SEUIL_PERTE[k];
  if (t.vacc === 0) {
    return `Aucun vacciné ${name} n'est encore rapporté pour ${d.scopeLabel.toLowerCase()} (cible ${fmtInt(t.cible)}). Les couvertures et taux de perte seront calculés dès la remontée des premiers rapports journaliers.`;
  }
  const withData = units.filter((u) => u[k].vacc > 0);
  const best = [...withData].sort((a, b) => (b[k].cv ?? 0) - (a[k].cv ?? 0));
  const low = withData.filter((u) => (u[k].cv ?? 0) < 80).sort((a, b) => (a[k].cv ?? 0) - (b[k].cv ?? 0));
  const above = withData.filter((u) => (u[k].cv ?? 0) >= 95);
  let txt = `Couverture ${name} de ${fmtPct(t.cv)} (${fmtInt(t.vacc)} vaccinés / cible ${fmtInt(t.cible)}). `;
  if (best.length) txt += `Meilleure performance : ${best[0].unit} (${fmtPct(best[0][k].cv, 1)})`;
  if (above.length) txt += ` ; ${above.length} ${label.toLowerCase()}${above.length > 1 ? "s" : ""} ≥ 95 %`;
  txt += ". ";
  if (low.length) txt += `${low.length} sous 80 % : ${joinAnd(low.slice(0, 4).map((u) => `${u.unit} (${fmtPct(u[k].cv, 1)})`))}${low.length > 4 ? "…" : ""} → intensifier le porte‑à‑porte et prévoir le ratissage. `;
  const tp = t.tauxPerte;
  if (tp != null) {
    if (tp < 0) txt += `Taux de perte ${name} négatif (${fmtPct(tp)}) : incohérence entre vaccinés et flacons utilisés à corriger. `;
    else if (tp > seuil) txt += `Taux de perte ${name} de ${fmtPct(tp)}, au‑dessus du seuil de ${seuil} % : vérifier la gestion des flacons entamés et la saisie des flacons utilisés. `;
    else txt += `Taux de perte ${name} maîtrisé (${fmtPct(tp)}, seuil ${seuil} %). `;
    const hi = withData.filter((u) => u[k].tauxPerte != null && (u[k].tauxPerte as number) > seuil).sort((a, b) => (b[k].tauxPerte ?? 0) - (a[k].tauxPerte ?? 0));
    if (hi.length) txt += `Pertes élevées : ${joinAnd(hi.slice(0, 4).map((u) => `${u.unit} (${fmtPct(u[k].tauxPerte, 1)})`))}.`;
  } else {
    txt += `Flacons utilisés ${name} non renseignés : taux de perte non calculable.`;
  }
  return txt;
}

function buildCouverture(ctx: Ctx, k: VaccineKey): void {
  const { pptx, d } = ctx;
  const name = VACCINE_LABELS[k];
  const color = VACC_COLOR[k];
  const cibleLabel = k === "rr" ? "Cible RR" : "Cible Polio";

  // A. Par antenne : tableau + graphique CV + graphique perte (modèle « par province »).
  if (d.antennes.length > 1) {
    const s = pptx.addSlide();
    header(ctx, s, `Résultats Partiels : Couverture vaccinale globale ${name} et taux de perte par Antenne`, { h: 0.77 });
    frame(ctx, s, 0.1, 0.9, 10.2, 6.5, color);
    // Tableau
    blockTitle(ctx, s, `Couverture vaccinale ${name} par Antenne`, 0.12, 0.92, 3.3, color, 0.36, 12);
    const rows: PptxGenJS.TableRow[] = [
      ["Antenne", cibleLabel, `Vaccinés ${name}`, `Couv ${name}%`].map((t) => ({ text: t, options: th(color, { fontSize: 9 }) })),
      ...d.antennes.map((u, i): PptxGenJS.TableRow => [
        { text: u.unit, options: td({ bold: true, fontSize: 9, ...zebra(i) }) },
        { text: fmtInt(u[k].cible), options: tdNum({ fontSize: 9, ...zebra(i) }) },
        { text: fmtInt(u[k].vacc), options: tdNum({ fontSize: 9, ...zebra(i) }) },
        { text: fmtNum(u[k].cv), options: tdColored(u[k].cv, "cov", { fontSize: 9 }) },
      ]),
      [
        { text: "Total", options: ttotal(color, { align: "left", fontSize: 9 }) },
        { text: fmtInt(d.total[k].cible), options: ttotal(color, { fontSize: 9 }) },
        { text: fmtInt(d.total[k].vacc), options: ttotal(color, { fontSize: 9 }) },
        { text: fmtNum(d.total[k].cv), options: ttotal(color, { fontSize: 9 }) },
      ],
    ];
    s.addTable(rows, { x: 0.15, y: 1.32, w: 3.25, colW: colW(3.25, [1.1, 0.8, 0.8, 0.7]), border: TABLE_BORDER, rowH: 0.26, valign: "middle", autoPage: false });
    // Graphique CV
    blockTitle(ctx, s, `Répartition de la couverture vaccinale (%) globale de ${name} par Antenne`, 3.5, 0.92, 6.75, color, 0.36, 12);
    const byCV = [...d.antennes].sort((a, b) => (b[k].cv ?? -1) - (a[k].cv ?? -1));
    colChart(ctx, s, { x: 3.5, y: 1.3, w: 6.75, h: 2.7 }, byCV.map((u) => u.unit), byCV.map((u) => r2(u[k].cv)), byCV.map((u) => covColor(u[k].cv)), { valAxisTitle: "Couverture vaccinale Globale", catAxisTitle: "Antenne", min: 0 });
    // Graphique perte
    blockTitle(ctx, s, `Répartition du taux de perte (%) du vaccin ${name} par Antenne`, 0.12, 4.05, 10.13, color, 0.36, 12);
    const byLoss = [...d.antennes].sort((a, b) => (b[k].tauxPerte ?? -999) - (a[k].tauxPerte ?? -999));
    colChart(ctx, s, { x: 0.15, y: 4.45, w: 10.1, h: 2.3 }, byLoss.map((u) => u.unit), byLoss.map((u) => r2(u[k].tauxPerte)), byLoss.map((u) => lossColor(u[k].tauxPerte)), { valAxisTitle: "Taux de perte", catAxisTitle: "Antenne" });
    legendsBottom(ctx, s, 6.83, 0.1, 10.2, true, true, false);
    commentBox(ctx, s, couvComment(d, k, d.antennes, "Antenne"), { x: 10.4, y: 0.9, w: 2.85, h: 6.5 });
  }

  // B. Par unité (ZS/AS) : deux graphiques (CV, perte) — modèle « par antenne ».
  const units = d.units;
  const per = 30;
  const byCV = [...units].sort((a, b) => (b[k].cv ?? -1) - (a[k].cv ?? -1));
  const byLoss = [...units].sort((a, b) => (a[k].tauxPerte ?? -999) - (b[k].tauxPerte ?? -999));
  const pages = Math.max(1, Math.ceil(units.length / per));
  for (let p = 0; p < pages; p++) {
    const s = pptx.addSlide();
    header(ctx, s, `Résultats Partiels : Couverture vaccinale globale ${name} et taux de perte par ${d.byUnitLabel}${suite(p, pages)}`, { h: 0.49 });
    frame(ctx, s, 0.05, 0.62, 10.35, 6.75, color);
    blockTitle(ctx, s, `Répartition de la couverture vaccinale globale de ${name} par ${d.byUnitLabel}`, 0.07, 0.64, 10.31, color, 0.36, 12);
    const a = byCV.slice(p * per, (p + 1) * per);
    colChart(ctx, s, { x: 0.1, y: 1.02, w: 10.25, h: 2.6 }, a.map((u) => u.unit), a.map((u) => r2(u[k].cv)), a.map((u) => covColor(u[k].cv)), { valAxisTitle: "Couverture vaccinale", catAxisTitle: pluralLabel(d.byUnitLabel), min: 0, fontSize: 8, labelFont: 8 });
    blockTitle(ctx, s, `Répartition du taux de perte de vaccin ${name} par ${d.byUnitLabel}`, 0.07, 3.7, 10.31, color, 0.36, 12);
    const b = byLoss.slice(p * per, (p + 1) * per);
    colChart(ctx, s, { x: 0.1, y: 4.08, w: 10.25, h: 2.6 }, b.map((u) => u.unit), b.map((u) => r2(u[k].tauxPerte)), b.map((u) => lossColor(u[k].tauxPerte)), { valAxisTitle: "Taux de perte", catAxisTitle: pluralLabel(d.byUnitLabel), fontSize: 8, labelFont: 8 });
    legendsBottom(ctx, s, 6.83, 0.05, 10.35, true, true, false);
    commentBox(ctx, s, couvComment(d, k, units, d.byUnitLabel), { x: 10.5, y: 0.62, w: 2.75, h: 6.75 });
  }
}

/* ─── 10. MAPI ─────────────────────────────────────────────────────────── */

function mapiComment(d: ReportData): string {
  const t = d.total;
  const tot = t.mapiNonGraves + t.mapiGraves;
  const doses = t.rr.vacc + t.nvpo2.vacc + t.vpob.vacc;
  if (doses === 0) return "Aucune dose administrée rapportée : pas encore de notification MAPI attendue.";
  if (tot === 0) return `Aucune MAPI notifiée pour ${fmtInt(doses)} doses administrées : sous‑notification probable — rappeler la notification systématique (y compris « zéro cas ») et la disponibilité des fiches MAPI.`;
  const grave = d.units.filter((u) => u.mapiGraves > 0).sort((a, b) => b.mapiGraves - a.mapiGraves);
  const top = [...d.units].sort((a, b) => (b.mapiPour100k ?? 0) - (a.mapiPour100k ?? 0)).filter((u) => (u.mapiPour100k ?? 0) > 0).slice(0, 3);
  let txt = `${fmtInt(tot)} MAPI notifiées (${fmtInt(t.mapiNonGraves)} non graves, ${fmtInt(t.mapiGraves)} graves), soit ${fmtNum(t.mapiPour100k)} MAPI pour 100 000 doses. `;
  if (grave.length) txt += `MAPI graves : ${joinAnd(grave.slice(0, 4).map((u) => `${u.unit} (${u.mapiGraves})`))} — investigation sous 48 h et suivi par le poste de commandement. `;
  if (top.length) txt += `Taux de notification les plus élevés : ${joinAnd(top.map((u) => `${u.unit} (${fmtNum(u.mapiPour100k, 0)})`))}.`;
  return txt;
}

function buildMapi(ctx: Ctx): void {
  const { pptx, d } = ctx;
  const t = d.total;
  const s = pptx.addSlide();
  header(ctx, s, "Résultats Partiels : Notification des MAPI", { h: 0.49 });
  frame(ctx, s, 0.05, 0.62, 10.35, 6.75, MAPI_BLUE);
  const cards = [
    { l: "Total MAPI", v: fmtInt(t.mapiNonGraves + t.mapiGraves) },
    { l: "MAPI non Grave", v: fmtInt(t.mapiNonGraves) },
    { l: "MAPI Grave", v: fmtInt(t.mapiGraves) },
    { l: "Proport. MAPI pour 100 000 Doses", v: fmtNum(t.mapiPour100k, 2) },
  ];
  cards.forEach((c, i) => {
    const x = 0.12 + i * 2.58;
    s.addShape(pptx.ShapeType.rect, { x, y: 0.7, w: 2.5, h: 0.62, fill: { color: MAPI_BLUE }, line: { color: MAPI_BLUE, width: 0 } });
    s.addText([{ text: c.l, options: { fontSize: 12, italic: true, bold: true, color: WHITE, breakLine: true } }, { text: c.v, options: { fontSize: 15, bold: true, color: WHITE } }], { x, y: 0.7, w: 2.5, h: 0.62, align: "center", valign: "middle", fontFace: FONT });
  });
  const lvl1 = d.antennes.length > 1 ? { rows: d.antennes, label: "Antenne" } : { rows: d.zones, label: "ZS" };
  const lvl2 = d.antennes.length > 1 ? { rows: d.zones, label: "ZS" } : { rows: d.aires, label: "AS" };
  const lvl3 = { rows: d.aires, label: "AS" };
  const tbl = (x: number, w: number, title: string, rows: UnitAgg[], label: string, maxRows: number, fill = MAPI_BLUE, showTotal = true) => {
    s.addText(title, { x, y: 1.4, w, h: 0.32, fontSize: 12, bold: true, color: BLACK, fontFace: FONT, align: "center" });
    const sorted = [...rows].sort((a, b) => b.mapiGraves - a.mapiGraves || b.mapiNonGraves - a.mapiNonGraves).slice(0, maxRows);
    const head: PptxGenJS.TableRow = [label, "MAPI non\nGrave", "MAPI\nGrave", "%MAPI\n100.000doses"].map((h, i) => ({ text: h, options: th(fill, { fontSize: 8, align: i === 0 ? "left" : "center" }) }));
    const body: PptxGenJS.TableRow[] = sorted.map((u, i) => [
      { text: u.unit, options: td({ bold: true, fontSize: 8, ...zebra(i) }) },
      { text: fmtInt(u.mapiNonGraves), options: tdNum({ fontSize: 8, ...zebra(i) }) },
      { text: fmtInt(u.mapiGraves), options: tdNum({ fontSize: 8, bold: u.mapiGraves > 0, color: u.mapiGraves > 0 ? "C00000" : BLACK, ...zebra(i) }) },
      { text: fmtNum(u.mapiPour100k, 2), options: tdNum({ fontSize: 8, ...zebra(i) }) },
    ]);
    if (showTotal) body.push([
      { text: "Total", options: ttotal(fill, { align: "left", fontSize: 8 }) },
      { text: fmtInt(t.mapiNonGraves), options: ttotal(fill, { fontSize: 8 }) },
      { text: fmtInt(t.mapiGraves), options: ttotal(fill, { fontSize: 8 }) },
      { text: fmtNum(t.mapiPour100k, 2), options: ttotal(fill, { fontSize: 8 }) },
    ]);
    frame(ctx, s, x, 1.72, w, 5.55, MAPI_BLUE);
    s.addTable([head, ...body], { x: x + 0.03, y: 1.75, w: w - 0.06, colW: colW(w - 0.06, [1.3, 0.75, 0.6, 0.85]), border: TABLE_BORDER, rowH: 0.22, valign: "middle", autoPage: false });
  };
  tbl(0.12, 3.25, `1. Nombre de MAPI par ${lvl1.label === "ZS" ? "ZS" : "Antenne"}`, lvl1.rows, lvl1.label, 23);
  tbl(3.5, 3.4, `2. Nombre de MAPI par ${lvl2.label}`, lvl2.rows, lvl2.label, 23);
  tbl(7.02, 3.3, `3. Nombre de MAPI par ${lvl3.label}`, lvl3.rows, lvl3.label, 23);
  commentBox(ctx, s, mapiComment(d), { x: 10.5, y: 0.62, w: 2.75, h: 6.75 });
}

function buildMapiChart(ctx: Ctx): void {
  const { pptx, d } = ctx;
  const s = pptx.addSlide();
  header(ctx, s, "Résultats Partiels : Proportion de MAPI pour 100 000 doses", { h: 0.49 });
  frame(ctx, s, 0.05, 0.6, W - 0.1, 5.95, MAPI_BLUE);
  blockTitle(ctx, s, "Répartition de la proportion de MAPI notifiées pour 100 000 doses administrées", 0.07, 0.62, W - 0.14, MAPI_BLUE, 0.42, 14);
  const sorted = [...d.units].sort((a, b) => (b.mapiPour100k ?? 0) - (a.mapiPour100k ?? 0));
  colChart(ctx, s, { x: 0.1, y: 1.1, w: W - 0.2, h: 5.4 }, sorted.map((u) => u.unit), sorted.map((u) => r2(u.mapiPour100k)), sorted.map(() => MAPI_BLUE), { valAxisTitle: "Taux de MAPI pour 100 000 Doses", catAxisTitle: d.byUnitLabel, fmt: "0", fontSize: sorted.length > 20 ? 8 : 10, labelFont: 9 });
  commentBox(ctx, s, mapiComment(d), { x: 0.05, y: 6.65, w: W - 0.1, h: 0.55 });
  sourceLine(s);
}

/* ─── 11. Surveillance MPV ─────────────────────────────────────────────── */

function mpvComment(d: ReportData): string {
  const t = d.total;
  const tot = t.survPFA + t.survRougeole + t.survFJ + t.survTNN;
  if (tot === 0) return "Aucun cas de MPV (PFA, rougeole, fièvre jaune, TNN) notifié pendant la campagne : documenter la recherche active des cas dans chaque aire, y compris les notifications « zéro cas ».";
  const parts: string[] = [];
  if (t.survPFA) parts.push(`${fmtInt(t.survPFA)} PFA`);
  if (t.survRougeole) parts.push(`${fmtInt(t.survRougeole)} cas suspects de rougeole`);
  if (t.survFJ) parts.push(`${fmtInt(t.survFJ)} fièvre jaune`);
  if (t.survTNN) parts.push(`${fmtInt(t.survTNN)} TNN`);
  const top = [...d.units].map((u) => ({ u, n: u.survPFA + u.survRougeole + u.survFJ + u.survTNN })).filter((x) => x.n > 0).sort((a, b) => b.n - a.n).slice(0, 4);
  return `${fmtInt(tot)} cas de MPV notifiés pendant la campagne : ${joinAnd(parts)}. Unités les plus concernées : ${joinAnd(top.map((x) => `${x.u.unit} (${x.n})`))}. Chaque PFA doit être investigué avec prélèvement de selles ; les cas suspects de rougeole font l'objet d'un prélèvement sanguin.`;
}

function buildMPV(ctx: Ctx): void {
  const { pptx, d } = ctx;
  const t = d.total;
  const s = pptx.addSlide();
  header(ctx, s, "Résultats Partiels : Surveillance MPV", { h: 0.49 });
  frame(ctx, s, 0.05, 0.62, 10.1, 6.75, MAROON);
  const cards = [
    { l: "PFA – Cas", v: fmtInt(t.survPFA) },
    { l: "Rougeole‑Cas", v: fmtInt(t.survRougeole) },
    { l: "Fièvre Jaune‑Cas", v: fmtInt(t.survFJ) },
    { l: "TNN – Cas", v: fmtInt(t.survTNN) },
  ];
  cards.forEach((c, i) => {
    const x = 0.12 + i * 2.5;
    s.addShape(pptx.ShapeType.rect, { x, y: 0.7, w: 2.42, h: 0.62, fill: { color: MAROON }, line: { color: MAROON, width: 0 } });
    s.addText([{ text: c.l, options: { fontSize: 13, italic: true, bold: true, color: WHITE, breakLine: true } }, { text: c.v, options: { fontSize: 15, bold: true, color: WHITE } }], { x, y: 0.7, w: 2.42, h: 0.62, align: "center", valign: "middle", fontFace: FONT });
  });
  const lvl1 = d.antennes.length > 1 ? { rows: d.antennes, label: "Antenne" } : { rows: d.zones, label: "ZS" };
  const lvl2 = d.antennes.length > 1 ? { rows: d.zones, label: "ZS" } : { rows: d.aires, label: "AS" };
  const lvl3 = { rows: d.aires, label: "AS" };
  const tbl = (x: number, w: number, num: number, rows: UnitAgg[], label: string, maxRows: number) => {
    blockTitle(ctx, s, `${num}. MPV notifiées par ${label === "AS" ? "Aire de santé" : label === "ZS" ? "Zone de Santé" : "Antenne"}`, x, 1.4, w, MAROON, 0.34, 12);
    const sorted = [...rows].sort((a, b) => (b.survPFA + b.survRougeole + b.survFJ + b.survTNN) - (a.survPFA + a.survRougeole + a.survFJ + a.survTNN)).slice(0, maxRows);
    const head: PptxGenJS.TableRow = [label, "PFA", "RR", "FJ", "TNN"].map((h, i) => ({ text: h, options: th(MAROON, { fontSize: 9, align: i === 0 ? "left" : "center" }) }));
    const body: PptxGenJS.TableRow[] = sorted.map((u, i) => [
      { text: u.unit, options: td({ bold: true, fontSize: 8, ...zebra(i) }) },
      { text: fmtInt(u.survPFA), options: tdNum({ fontSize: 8, ...zebra(i) }) },
      { text: fmtInt(u.survRougeole), options: tdNum({ fontSize: 8, ...zebra(i) }) },
      { text: fmtInt(u.survFJ), options: tdNum({ fontSize: 8, ...zebra(i) }) },
      { text: fmtInt(u.survTNN), options: tdNum({ fontSize: 8, ...zebra(i) }) },
    ]);
    body.push([
      { text: "Total", options: ttotal(MAROON, { align: "left", fontSize: 8 }) },
      { text: fmtInt(t.survPFA), options: ttotal(MAROON, { fontSize: 8 }) },
      { text: fmtInt(t.survRougeole), options: ttotal(MAROON, { fontSize: 8 }) },
      { text: fmtInt(t.survFJ), options: ttotal(MAROON, { fontSize: 8 }) },
      { text: fmtInt(t.survTNN), options: ttotal(MAROON, { fontSize: 8 }) },
    ]);
    frame(ctx, s, x, 1.78, w, 5.5, MAROON);
    s.addTable([head, ...body], { x: x + 0.03, y: 1.81, w: w - 0.06, colW: colW(w - 0.06, [1.5, 0.55, 0.55, 0.55, 0.55]), border: TABLE_BORDER, rowH: 0.22, valign: "middle", autoPage: false });
  };
  tbl(0.12, 3.15, 1, lvl1.rows, lvl1.label, 23);
  tbl(3.4, 3.3, 2, lvl2.rows, lvl2.label, 23);
  tbl(6.82, 3.25, 3, lvl3.rows, lvl3.label, 23);
  commentBox(ctx, s, mpvComment(d), { x: 10.25, y: 0.62, w: 3.0, h: 6.75 });
}

/* ─── 12. Récupération PEV de routine ──────────────────────────────────── */

function recupComment(d: ReportData): string {
  const t = d.total;
  const ident = t.pevIdent.reduce((a, b) => a + b, 0);
  const rec = t.pevRecup.reduce((a, b) => a + b, 0);
  if (ident === 0 && rec === 0) return "Aucun enfant ou femme enceinte identifié / récupéré en PEV systématique n'est encore rapporté : rappeler aux équipes l'enregistrement systématique des récupérations pendant la campagne.";
  const idx = t.pevRecup.map((v, i) => ({ v, i })).sort((a, b) => b.v - a.v).slice(0, 3).filter((x) => x.v > 0);
  const dtc = ["DTC1", "DTC2", "DTC3"].map((k) => ANTIGENES.findIndex((a) => a.key === k)).reduce((s, i) => s + (t.pevRecup[i] ?? 0), 0);
  let txt = `${fmtInt(rec)} enfants et femmes récupérés en PEV de routine`;
  if (ident > 0) txt += ` sur ${fmtInt(ident)} identifiés (${fmtPct((rec / ident) * 100, 1)})`;
  txt += ". ";
  if (idx.length) txt += `Antigènes les plus récupérés : ${joinAnd(idx.map((x) => `${ANTIGENES[x.i].label} (${fmtInt(x.v)})`))}. `;
  txt += `Progression de la récupération d'enfants en PEV de routine avec un accent sur l'antigène DTC (${fmtInt(dtc)} doses DTC1‑3).`;
  return txt;
}

function buildRecup(ctx: Ctx): void {
  const { pptx, d } = ctx;
  const labels = ANTIGENES.map((a) => a.label);
  const nCols = 1 + labels.length;
  const weights = [1.45, ...labels.map(() => 0.5)];
  type Sel = (u: UnitAgg) => number[];
  const selRecup: Sel = (u) => u.pevRecup;
  const selIdent: Sel = (u) => u.pevIdent;
  const selPct: Sel = (u) => u.pevRecup.map((v, i) => (u.pevIdent[i] > 0 ? (v / u.pevIdent[i]) * 100 : NaN));
  const rowsFor = (units: UnitAgg[], label: string, sel: Sel, isPct = false): PptxGenJS.TableRow[] => {
    const fmt = (v: number) => (isPct ? (Number.isFinite(v) ? fmtPct(v, 0, "%") : "—") : fmtInt(v));
    const head: PptxGenJS.TableRow = [{ text: label, options: th(MAPI_BLUE, { fontSize: 8, align: "left" }) }, ...labels.map((l) => ({ text: l, options: th(MAPI_BLUE, { fontSize: 8 }) }))];
    const body: PptxGenJS.TableRow[] = units.map((u, i) => [
      { text: u.unit, options: td({ bold: true, fontSize: 8, ...zebra(i) }) },
      ...sel(u).map((v) => ({ text: fmt(v), options: isPct ? tdColored(Number.isFinite(v) ? v : null, "cov", { fontSize: 8 }) : tdNum({ fontSize: 8, ...zebra(i) }) })),
    ]);
    const tot: PptxGenJS.TableRow = [{ text: "Total", options: ttotal(MAPI_BLUE, { align: "left", fontSize: 8 }) }, ...sel(d.total).map((v) => ({ text: fmt(v), options: ttotal(MAPI_BLUE, { fontSize: 8 }) }))];
    return [head, ...body, tot];
  };
  const drawTable = (s: PptxGenJS.Slide, rows: PptxGenJS.TableRow[], y: number, title: string): number => {
    blockTitle(ctx, s, title, 0.07, y, W - 0.14, MAPI_BLUE, 0.4, 13);
    s.addTable(rows, { x: 0.1, y: y + 0.45, w: W - 0.2, colW: colW(W - 0.2, weights.slice(0, nCols)), border: TABLE_BORDER, rowH: 0.26, valign: "middle", autoPage: false });
    return y + 0.45 + 0.27 * rows.length + 0.15;
  };

  // A. Vue par antenne : récupérés, identifiés et % de récupération sur une diapo.
  if (d.antennes.length > 1) {
    const s = pptx.addSlide();
    header(ctx, s, "Résultats Partiels : Récupération PEV de routine par Antenne", { h: 0.49 });
    frame(ctx, s, 0.05, 0.6, W - 0.1, 5.35, MAPI_BLUE);
    let y = 0.62;
    y = drawTable(s, rowsFor(d.antennes, "Antenne", selRecup), y, "1. Données de récupération des enfants au PEV de routine par Antenne (récupérés)");
    y = drawTable(s, rowsFor(d.antennes, "Antenne", selIdent), y, "2. Enfants et femmes enceintes identifiés pour récupération par Antenne");
    drawTable(s, rowsFor(d.antennes, "Antenne", selPct, true), y, "3. Taux de récupération (récupérés ÷ identifiés) par Antenne");
    commentBox(ctx, s, recupComment(d), { x: 0.05, y: 6.05, w: W - 0.1, h: 0.85 });
    sourceLine(s);
  }

  // B. Vue par unité (ZS/AS) : récupérés (paginé).
  const all = rowsFor(d.units, d.byUnitLabel, selRecup);
  const head = all[0];
  const body = all.slice(1, -1);
  const tot = all[all.length - 1];
  const pages = chunk(body, 15);
  pages.forEach((page, idx) => {
    const s = pptx.addSlide();
    header(ctx, s, `Résultats Partiels : Récupération PEV de routine par ${d.byUnitLabel}${suite(idx, pages.length)}`, { h: 0.49 });
    frame(ctx, s, 0.05, 0.6, W - 0.1, 5.35, MAPI_BLUE);
    const rows = idx === pages.length - 1 ? [head, ...page, tot] : [head, ...page];
    drawTable(s, rows, 0.62, `${d.antennes.length > 1 ? "4" : "1"}. Données de récupération des enfants au PEV de routine par ${d.byUnitLabel} (récupérés)`);
    commentBox(ctx, s, recupComment(d), { x: 0.05, y: 6.05, w: W - 0.1, h: 0.85 });
    sourceLine(s);
  });
}

/* ─── 13. Tranches d'âge ───────────────────────────────────────────────── */

function pctParts(vals: number[]): number[] {
  const s = vals.reduce((a, b) => a + b, 0);
  return vals.map((v) => (s > 0 ? Math.round((v / s) * 10000) / 100 : 0));
}

function agePanel(ctx: Ctx, s: PptxGenJS.Slide, pos: { x: number; y: number; w: number; h: number }, k: VaccineKey, units: UnitAgg[]): void {
  const { pptx, d } = ctx;
  const color = VACC_COLOR[k];
  const labels = k === "rr" ? RR_AGE_LABELS : POLIO_AGE_LABELS;
  const tot = d.total[k].ages;
  frame(ctx, s, pos.x, pos.y, pos.w, pos.h, color);
  blockTitle(ctx, s, `Proportion des enfants vaccinés en ${VACCINE_LABELS[k]} par tranche d'âge : Globale et par ${d.byUnitLabel}`, pos.x, pos.y, pos.w, color, 0.32, 11);
  // Légende
  const lw = (pos.w * 0.38) / labels.length;
  labels.forEach((l, i) => {
    s.addShape(pptx.ShapeType.rect, { x: pos.x + i * lw, y: pos.y + 0.36, w: lw - 0.02, h: 0.26, fill: { color: AGE_COLORS[i] }, line: { color: WHITE, width: 0 } });
    s.addText(`% ${l}`, { x: pos.x + i * lw, y: pos.y + 0.36, w: lw - 0.02, h: 0.26, fontSize: 8, bold: true, color: WHITE, align: "center", valign: "middle", fontFace: FONT });
  });
  const total = tot.reduce((a, b) => a + b, 0);
  if (total > 0) {
    s.addChart(pptx.ChartType.doughnut, [{ name: "Tranches d'âge", labels, values: pctParts(tot) }], {
      x: pos.x, y: pos.y + 0.65, w: pos.w * 0.38, h: pos.h - 0.7,
      chartColors: AGE_COLORS.slice(0, labels.length), holeSize: 55, showLegend: false, showPercent: false, showValue: true,
      dataLabelFormatCode: '0.00"%"', dataLabelFontSize: 8, dataLabelColor: WHITE, dataLabelFontBold: true, showTitle: false,
    });
  } else {
    s.addText("Aucun vacciné", { x: pos.x, y: pos.y + 0.65, w: pos.w * 0.38, h: pos.h - 0.7, align: "center", valign: "middle", fontSize: 10, italic: true, color: GREY });
  }
  const withData = units.filter((u) => u[k].vacc > 0);
  const cats = withData.length ? withData : units.slice(0, 12);
  const series = labels.map((l, i) => ({ name: `% ${l}`, labels: cats.map((u) => u.unit), values: cats.map((u) => pctParts(u[k].ages)[i]) }));
  s.addChart(pptx.ChartType.bar, series, {
    x: pos.x + pos.w * 0.39, y: pos.y + 0.36, w: pos.w * 0.6, h: pos.h - 0.4,
    barDir: "col", barGrouping: "stacked", chartColors: AGE_COLORS.slice(0, labels.length),
    showValue: withData.length > 0 && cats.length <= 14, dataLabelFormatCode: '0.0"%"', dataLabelFontSize: 6, dataLabelColor: WHITE,
    catAxisLabelFontSize: cats.length > 14 ? 6 : 7, catAxisLabelRotate: cats.length > 10 ? -45 : 0, valAxisLabelFontSize: 7, valAxisMaxVal: 100, valAxisMinVal: 0,
    valAxisLabelFormatCode: '0"%"', showLegend: false, showValAxisTitle: true, valAxisTitle: "%Vaccinés par tranche d'âge", valAxisTitleFontSize: 7,
    showCatAxisTitle: true, catAxisTitle: d.byUnitLabel, catAxisTitleFontSize: 7, valGridLine: { style: "dash", color: "E0E0E0", size: 0.5 }, barGapWidthPct: 60,
  });
}

function ageComment(d: ReportData): string {
  const t = d.total;
  const parts: string[] = [];
  (["rr", "nvpo2", "vpob"] as VaccineKey[]).forEach((k) => {
    const labels = k === "rr" ? RR_AGE_LABELS : POLIO_AGE_LABELS;
    const ages = t[k].ages;
    const tot = ages.reduce((a, b) => a + b, 0);
    if (tot === 0) return;
    const order = ages.map((v, i) => ({ v, i })).sort((a, b) => b.v - a.v);
    parts.push(`${VACCINE_LABELS[k]} : la tranche ${labels[order[0].i]} est la plus vaccinée (${fmtPct((order[0].v / tot) * 100, 1)})${order[1] ? ` suivie de ${labels[order[1].i]} (${fmtPct((order[1].v / tot) * 100, 1)})` : ""}`);
  });
  if (parts.length === 0) return "Répartition par tranche d'âge non disponible : aucun vacciné rapporté à ce stade.";
  return parts.join(" ; ") + ".";
}

function buildAges(ctx: Ctx): void {
  const { pptx, d } = ctx;
  const s = pptx.addSlide();
  header(ctx, s, "Résultats Partiels : Proportion des enfants vaccinés par tranches d'âges", { h: 0.49 });
  const units = d.antennes.length > 1 && d.units.length > 16 ? d.units : d.units;
  agePanel(ctx, s, { x: 0.05, y: 0.75, w: 6.45, h: 2.6 }, "rr", units);
  agePanel(ctx, s, { x: 6.75, y: 0.75, w: 6.5, h: 2.6 }, "nvpo2", units);
  agePanel(ctx, s, { x: 3.1, y: 3.5, w: 7.2, h: 2.85 }, "vpob", units);
  commentBox(ctx, s, ageComment(d), { x: 0.2, y: 6.45, w: W - 0.4, h: 0.5 });
  sourceLine(s);
}

/* ─── 14. Sexe ─────────────────────────────────────────────────────────── */

function sexPanel(ctx: Ctx, s: PptxGenJS.Slide, pos: { x: number; y: number; w: number; h: number }, k: VaccineKey, units: UnitAgg[]): void {
  const { pptx, d } = ctx;
  const color = VACC_COLOR[k];
  const t = d.total[k];
  frame(ctx, s, pos.x, pos.y, pos.w, pos.h, color);
  blockTitle(ctx, s, `Proportion des enfants vaccinés en ${VACCINE_LABELS[k]} par Sexe et par ${d.byUnitLabel}`, pos.x, pos.y, pos.w, color, 0.32, 11);
  const lw = (pos.w * 0.38) / 3;
  const leg = [{ t: "Légende :", c: "595959" }, { t: "% Fille", c: SEX_COLORS.fille }, { t: "% Garçon", c: SEX_COLORS.garcon }];
  leg.forEach((l, i) => {
    s.addShape(pptx.ShapeType.rect, { x: pos.x + i * lw, y: pos.y + 0.36, w: lw - 0.02, h: 0.26, fill: { color: l.c }, line: { color: WHITE, width: 0 } });
    s.addText(l.t, { x: pos.x + i * lw, y: pos.y + 0.36, w: lw - 0.02, h: 0.26, fontSize: 8, bold: true, color: WHITE, align: "center", valign: "middle", fontFace: FONT });
  });
  const tot = t.filles + t.garcons;
  if (tot > 0) {
    s.addChart(pptx.ChartType.doughnut, [{ name: "Sexe", labels: ["% Fille", "% Garçon"], values: pctParts([t.filles, t.garcons]) }], {
      x: pos.x, y: pos.y + 0.65, w: pos.w * 0.38, h: pos.h - 0.7,
      chartColors: [SEX_COLORS.fille, SEX_COLORS.garcon], holeSize: 55, showLegend: false, showValue: true, showPercent: false,
      dataLabelFormatCode: '0.00"%"', dataLabelFontSize: 8, dataLabelColor: WHITE, dataLabelFontBold: true, showTitle: false,
    });
  } else {
    s.addText("Aucun vacciné", { x: pos.x, y: pos.y + 0.65, w: pos.w * 0.38, h: pos.h - 0.7, align: "center", valign: "middle", fontSize: 10, italic: true, color: GREY });
  }
  const withData = units.filter((u) => u[k].filles + u[k].garcons > 0);
  const cats = withData.length ? withData : units.slice(0, 12);
  const series = [
    { name: "% Fille", labels: cats.map((u) => u.unit), values: cats.map((u) => pctParts([u[k].filles, u[k].garcons])[0]) },
    { name: "% Garçon", labels: cats.map((u) => u.unit), values: cats.map((u) => pctParts([u[k].filles, u[k].garcons])[1]) },
  ];
  s.addChart(pptx.ChartType.bar, series, {
    x: pos.x + pos.w * 0.39, y: pos.y + 0.36, w: pos.w * 0.6, h: pos.h - 0.4,
    barDir: "col", barGrouping: "stacked", chartColors: [SEX_COLORS.fille, SEX_COLORS.garcon],
    showValue: withData.length > 0 && cats.length <= 14, dataLabelFormatCode: '0.0"%"', dataLabelFontSize: 6, dataLabelColor: WHITE,
    catAxisLabelFontSize: cats.length > 14 ? 6 : 7, catAxisLabelRotate: cats.length > 10 ? -45 : 0, valAxisLabelFontSize: 7, valAxisMaxVal: 100, valAxisMinVal: 0,
    valAxisLabelFormatCode: '0"%"', showLegend: false, showValAxisTitle: true, valAxisTitle: "%Vaccinés Sexe", valAxisTitleFontSize: 7,
    showCatAxisTitle: true, catAxisTitle: d.byUnitLabel, catAxisTitleFontSize: 7, valGridLine: { style: "dash", color: "E0E0E0", size: 0.5 }, barGapWidthPct: 60,
  });
}

function sexComment(d: ReportData): string {
  const t = d.total;
  const f = t.rr.filles + t.nvpo2.filles + t.vpob.filles;
  const g = t.rr.garcons + t.nvpo2.garcons + t.vpob.garcons;
  if (f + g === 0) return "Répartition par sexe non disponible : aucun vacciné rapporté à ce stade.";
  const pf = (f / (f + g)) * 100;
  const pg = 100 - pf;
  const parts = (["rr", "nvpo2", "vpob"] as VaccineKey[]).filter((k) => t[k].filles + t[k].garcons > 0).map((k) => `${VACCINE_LABELS[k]} ${fmtPct((t[k].filles / (t[k].filles + t[k].garcons)) * 100, 1)} de filles`);
  return `${pf >= pg ? "Les filles sont plus vaccinées que les garçons" : "Les garçons sont plus vaccinés que les filles"} avec une proportion de ${fmtPct(Math.max(pf, pg), 1)} contre ${fmtPct(Math.min(pf, pg), 1)} (${parts.join(", ")}). ${Math.abs(pf - pg) < 4 ? "L'équilibre entre les sexes est respecté." : "Un écart marqué entre les sexes mérite une vérification des données et de l'accès des filles/garçons aux sites."}`;
}

function buildSexe(ctx: Ctx): void {
  const { pptx, d } = ctx;
  const s = pptx.addSlide();
  header(ctx, s, "Résultats Partiels : Proportion des enfants vaccinés par sexe", { h: 0.49 });
  sexPanel(ctx, s, { x: 0.05, y: 0.65, w: 6.45, h: 2.6 }, "rr", d.units);
  sexPanel(ctx, s, { x: 6.75, y: 0.65, w: 6.5, h: 2.6 }, "nvpo2", d.units);
  sexPanel(ctx, s, { x: 3.1, y: 3.4, w: 7.2, h: 2.85 }, "vpob", d.units);
  commentBox(ctx, s, sexComment(d), { x: 0.2, y: 6.45, w: W - 0.4, h: 0.5 });
  sourceLine(s);
}

/* ─── 15. Supervision des équipes (ODK) ────────────────────────────────── */

function supColorHex(v: number | null): string {
  const c = supervisionColor(v);
  return c === "green" ? SUP_GREEN : c === "orange" ? SUP_ORANGE : c === "red" ? SUP_RED : NONE;
}

function buildSupervision(ctx: Ctx): void {
  const { pptx, d } = ctx;
  const sup = d.supervision;
  const label = d.byUnitLabel === "Aire de Santé" ? "Zone de Santé" : d.byUnitLabel;
  const jour = d.jourLabels.length ? d.jourLabels[d.jourLabels.length - 1] : "J1";
  const jourTxt = jour === "Ratissage" ? "au ratissage" : `au ${jour}`;
  const source = `Source : Questionnaire de supervision des équipes intégrées (ODK/OMS, formulaire ${sup.formTitle || "Bloc 3"}) — supervisions à partir du ${sup.dateMin ? sup.dateMin.split("-").reverse().join("/") : "17/08/2026"}`;

  if (!sup.available) {
    const s = pptx.addSlide();
    headerAccroche(ctx, s, "Supervision des équipes de vaccination — données ODK indisponibles");
    s.addText(`Les données de supervision n'ont pas pu être récupérées depuis le serveur ODK (${sup.reason ?? "erreur réseau"}). Régénérer le rapport avec une connexion Internet active.`, { x: 1, y: 3, w: W - 2, h: 1.2, fontSize: 16, italic: true, color: GREY, align: "center", fontFace: FONT });
    sourceLine(s, source);
    return;
  }

  const zsRows = [...sup.byZS].sort((a, b) => (b.pctASVisitees ?? -1) - (a.pctASVisitees ?? -1) || a.zs.localeCompare(b.zs));
  const nZS = zsRows.length;
  const under80 = zsRows.filter((z) => (z.pctASVisitees ?? 0) < 80);
  const under50 = zsRows.filter((z) => (z.pctASVisitees ?? 0) < 50);
  const totalAS = zsRows.reduce((a, z) => a + z.nbASTotal, 0);
  const visAS = zsRows.reduce((a, z) => a + z.nbASVisitees, 0);
  const pctGlobal = totalAS ? (visAS / totalAS) * 100 : 0;

  // A. Accroche : % d'AS visitées par ZS + carte.
  {
    const s = pptx.addSlide();
    const titre = sup.total === 0
      ? `Aucune supervision d'équipe n'est encore saisie dans ODK pour ${d.provinceLabel} ${jourTxt} de la campagne`
      : `${under80.length} ${label}${under80.length > 1 ? "s" : ""} sur ${nZS} ${under80.length > 1 ? "restent" : "reste"} en dessous de 80 % d'aires de santé visitées ${jourTxt} de la campagne (${fmtPct(pctGlobal, 1)} d'AS visitées, ${sup.total} supervisions)`;
    headerAccroche(ctx, s, titre);
    const per = 30;
    const shown = zsRows.slice(0, per);
    colChart(ctx, s, { x: 0.2, y: 1.15, w: 6.2, h: 5.6 }, shown.map((z) => z.zs), shown.map((z) => r2(z.pctASVisitees)), shown.map((z) => supColorHex(z.pctASVisitees)), { barDir: "bar", fmt: '0.0"%"', min: 0, max: 100, fontSize: 8, labelFont: 8, showAxis: false, gap: 60 });
    if (sup.mapPng) {
      s.addImage({ data: sup.mapPng, x: 6.6, y: 1.1, w: 5.6, h: 5.6 });
    } else {
      s.addText("Carte indisponible (fond cartographique non chargé).", { x: 6.6, y: 3.5, w: 5.6, h: 0.6, align: "center", fontSize: 12, italic: true, color: GREY });
    }
    // Légende de la carte
    const leg = [{ c: SUP_RED, t: "< 50%" }, { c: SUP_ORANGE, t: "50-80%" }, { c: SUP_GREEN, t: ">=80%" }];
    leg.forEach((l, i) => {
      s.addShape(pptx.ShapeType.ellipse, { x: 6.9 + i * 1.3, y: 6.78, w: 0.16, h: 0.16, fill: { color: l.c }, line: { color: l.c, width: 0 } });
      s.addText(l.t, { x: 7.1 + i * 1.3, y: 6.72, w: 1.1, h: 0.28, fontSize: 9, color: GREY, fontFace: FONT, valign: "middle" });
    });
    s.addText(`${source}. Le questionnaire de supervision est le questionnaire des sites.`, { x: 0.2, y: 7.02, w: W - 0.4, h: 0.4, fontSize: 10, color: BLACK, fontFace: FONT_TITLE, fit: "shrink" });
  }

  // B. Analyse province : barres « Antenne | ZS », carte, carte des points, tableau des taux.
  {
    const s = pptx.addSlide();
    const accroche = sup.total === 0
      ? `${d.provinceLabel} | Aucune donnée de supervision ODK depuis le lancement — les superviseurs doivent saisir chaque supervision d'équipe le jour même`
      : under50.length > 0
        ? `${d.provinceLabel} | L'analyse des données de supervision montre que ${under50.length} ZS (${joinAnd(under50.slice(0, 4).map((z) => z.zs))}${under50.length > 4 ? "…" : ""}) ${under50.length > 1 ? "ont" : "a"} moins de 50 % des aires de santé visitées ${jourTxt} de la campagne`
        : under80.length > 0
          ? `${d.provinceLabel} | La plupart des aires de santé ont été visitées ${jourTxt} ; ${under80.length} ZS (${joinAnd(under80.slice(0, 4).map((z) => z.zs))}${under80.length > 4 ? "…" : ""}) ${under80.length > 1 ? "restent" : "reste"} en dessous de 80 % d'AS visitées`
          : `${d.provinceLabel} | Toutes les ZS ont au moins 80 % de leurs aires de santé visitées ${jourTxt} de la campagne`;
    headerAccroche(ctx, s, accroche);
    const shown = zsRows.slice(0, 30);
    colChart(ctx, s, { x: 0.15, y: 1.1, w: 4.6, h: 4.3 }, shown.map((z) => `${z.antenne} | ${z.zs}`), shown.map((z) => r2(z.pctASVisitees)), shown.map((z) => supColorHex(z.pctASVisitees)), { barDir: "bar", fmt: '0.0"%"', min: 0, max: 100, fontSize: 7, labelFont: 7, showAxis: false, gap: 50 });
    if (sup.mapPng) s.addImage({ data: sup.mapPng, x: 4.85, y: 1.05, w: 4.1, h: 4.1 });
    if (sup.pointsMapPng) {
      s.addImage({ data: sup.pointsMapPng, x: 9.05, y: 1.05, w: 4.1, h: 4.1 });
      s.addText(`Points = ${fmtInt(sup.points.length)} soumissions ODK de supervision des équipes (GPS)`, { x: 9.05, y: 5.15, w: 4.1, h: 0.25, fontSize: 8, italic: true, color: GREY, fontFace: FONT, align: "center" });
    } else if (!sup.mapPng) s.addText("Cartes indisponibles", { x: 4.85, y: 2.5, w: 8, h: 0.6, align: "center", fontSize: 12, italic: true, color: GREY });
    // Parties non supervisées : ZS avec le plus d'AS non visitées
    const nonVis = zsRows.filter((z) => z.asNonVisitees.length > 0).sort((a, b) => b.asNonVisitees.length - a.asNonVisitees.length).slice(0, 3);
    const nonVisTxt = nonVis.length
      ? `Parties non supervisées : ${nonVis.map((z) => `${z.zs} (${z.asNonVisitees.length} AS)`).join(", ")}${zsRows.filter((z) => z.asNonVisitees.length > 0).length > 3 ? "…" : ""}`
      : "Toutes les aires de santé du périmètre ont été visitées.";
    s.addText(nonVisTxt, { x: 4.85, y: 5.4, w: 8.3, h: 0.4, fontSize: 12, color: BLACK, fontFace: FONT_TITLE, align: "center" });
    // Tableau des taux
    const t = d.total;
    const head: PptxGenJS.TableRow = ["Taux de couverture RR", "Taux de couverture nVPO2", "Taux de couverture VPOb", "Taux de complétude"].map((h) => ({ text: h, options: th(TITLE_BLUE, { fontSize: 12, fontFace: FONT_TITLE }) }));
    const val = (v: number | null): PptxGenJS.TableCell => ({ text: fmtPct(v, 0, "%"), options: td({ fontSize: 16, bold: true, align: "center", fontFace: FONT_TITLE, fill: { color: (v ?? 0) < 80 ? "F4CCCC" : "D9EAD3" } }) });
    s.addTable([head, [val(t.rr.cv), val(t.nvpo2.cv), val(t.vpob.cv), val(t.completude)]], { x: 0.45, y: 5.84, w: W - 0.9, colW: colW(W - 0.9, [1, 1, 1, 1]), border: { type: "solid", color: BLACK, pt: 0.75 }, rowH: 0.5, valign: "middle", autoPage: false });
    s.addText("Source : 1. ODK, 2. Masque de saisie de la campagne", { x: 0.13, y: 7.05, w: 12, h: 0.3, fontSize: 11, color: TITLE_BLUE, fontFace: FONT_TITLE });
  }

  // C. Détail par ZS (tableau) — supervisions, AS visitées, score.
  {
    const rows = zsRows;
    const pages = chunk(rows, 20);
    pages.forEach((page, idx) => {
      const s = pptx.addSlide();
      header(ctx, s, `Supervision des équipes de vaccination (ODK) : couverture et score par ${label}${suite(idx, pages.length)}`, { h: 0.62 });
      const head: PptxGenJS.TableRow = ["Antenne", label, "Supervisions", "AS totales", "AS visitées", "% AS visitées", "Score conformité", "AS non encore visitées"].map((h, i) => ({ text: h, options: th(TITLE_BLUE, { fontSize: 9, align: i < 2 || i === 7 ? "left" : "center" }) }));
      const body: PptxGenJS.TableRow[] = page.map((z, i) => {
        const pc = supColorHex(z.pctASVisitees);
        const sc = supColorHex(z.score);
        return [
          { text: z.antenne, options: td({ fontSize: 9, ...zebra(i) }) },
          { text: z.zs, options: td({ bold: true, fontSize: 9, ...zebra(i) }) },
          { text: fmtInt(z.nbSupervisions), options: tdNum({ fontSize: 9, align: "center", ...zebra(i) }) },
          { text: fmtInt(z.nbASTotal), options: tdNum({ fontSize: 9, align: "center", ...zebra(i) }) },
          { text: fmtInt(z.nbASVisitees), options: tdNum({ fontSize: 9, align: "center", ...zebra(i) }) },
          { text: fmtPct(z.pctASVisitees, 1), options: td({ fontSize: 9, bold: true, align: "center", fill: { color: pc }, color: onColor(pc) }) },
          { text: fmtPct(z.score, 1), options: td({ fontSize: 9, bold: true, align: "center", fill: { color: sc }, color: onColor(sc) }) },
          { text: z.asNonVisitees.length ? (z.asNonVisitees.length > 6 ? `${z.asNonVisitees.slice(0, 6).join(", ")} … (+${z.asNonVisitees.length - 6})` : z.asNonVisitees.join(", ")) : "—", options: td({ fontSize: 8, ...zebra(i) }) },
        ];
      });
      s.addTable([head, ...body], { x: 0.3, y: 0.85, w: W - 0.6, colW: colW(W - 0.6, [1.1, 1.5, 0.9, 0.8, 0.85, 0.95, 1.05, 5.5]), border: TABLE_BORDER, rowH: 0.27, valign: "middle", autoPage: false });
      const scoreLow = rows.filter((z) => z.score != null && (z.score as number) < 80).sort((a, b) => (a.score ?? 0) - (b.score ?? 0));
      const cmt = sup.total === 0
        ? "Aucune supervision saisie : le score de conformité ne peut pas être calculé."
        : `${sup.total} supervisions saisies dans ODK ; ${visAS} AS visitées sur ${totalAS} (${fmtPct(pctGlobal, 1)}). ${scoreLow.length ? `Scores de conformité < 80 % : ${joinAnd(scoreLow.slice(0, 5).map((z) => `${z.zs} (${fmtPct(z.score, 0)})`))}${scoreLow.length > 5 ? "…" : ""} — briefing correctif des équipes.` : "Tous les scores de conformité sont ≥ 80 %."}`;
      commentBox(ctx, s, cmt, { x: 0.3, y: 6.55, w: W - 0.6, h: 0.6 });
      s.addText(source, { x: 0.13, y: 7.18, w: 12.5, h: 0.28, fontSize: 10, color: TITLE_BLUE, fontFace: FONT_TITLE, fit: "shrink" });
    });
  }

  // D. Conformité par indicateur (qualité des équipes supervisées).
  {
    const s = pptx.addSlide();
    header(ctx, s, "Supervision des équipes (ODK) : conformité par indicateur du questionnaire", { h: 0.62 });
    const conf = sup.conformity.filter((c) => c.n > 0);
    if (conf.length === 0) {
      s.addText("Aucune supervision saisie : conformité non calculable.", { x: 1, y: 3, w: W - 2, h: 1, align: "center", fontSize: 14, italic: true, color: GREY });
    } else {
      const sorted = [...conf].sort((a, b) => (a.pct ?? 0) - (b.pct ?? 0));
      const half = Math.ceil(sorted.length / 2);
      const draw = (items: typeof sorted, x: number) => {
        colChart(ctx, s, { x, y: 0.85, w: (W - 0.7) / 2, h: 5.6 }, items.map((c) => (c.n < sup.total ? `${c.label} (n=${c.n})` : c.label)), items.map((c) => r2(c.pct)), items.map((c) => supColorHex(c.pct)), { barDir: "bar", fmt: '0"%"', min: 0, max: 100, fontSize: 7, labelFont: 7, showAxis: false, gap: 45 });
      };
      draw(sorted.slice(0, half), 0.2);
      draw(sorted.slice(half), 0.35 + (W - 0.7) / 2);
      const weak = sorted.filter((c) => (c.pct ?? 100) < 70).slice(0, 5);
      const strong = [...sorted].reverse().filter((c) => (c.pct ?? 0) >= 90).slice(0, 3);
      const cmt = `${sup.total} équipes supervisées. ${weak.length ? `Points faibles (< 70 % de conformité) : ${joinAnd(weak.map((c) => `${c.label} (${fmtPct(c.pct, 0)})`))}. ` : "Aucun indicateur sous 70 % de conformité. "}${strong.length ? `Points forts : ${joinAnd(strong.map((c) => c.label))}.` : ""}`;
      commentBox(ctx, s, cmt, { x: 0.3, y: 6.55, w: W - 0.6, h: 0.6 });
    }
    s.addText(source, { x: 0.13, y: 7.18, w: 12.5, h: 0.28, fontSize: 10, color: TITLE_BLUE, fontFace: FONT_TITLE, fit: "shrink" });
  }
}

/* ─── 16. Problèmes / Actions ──────────────────────────────────────────── */

function buildProblemes(ctx: Ctx): void {
  const { pptx, d } = ctx;
  const head: PptxGenJS.TableRow = ["Problèmes identifiés", `${pluralLabel(d.byUnitLabel)} concernées`, "Solutions proposées"].map((h) => ({
    text: h, options: td({ bold: true, fontSize: 15, fontFace: FONT_TITLE, margin: [4, 6, 4, 6] }),
  }));
  const rows = d.problemes;
  const pages = chunk(rows, 4);
  pages.forEach((page, idx) => {
    const s = pptx.addSlide();
    header(ctx, s, `Problèmes rencontrés / Actions correctrices${suite(idx, pages.length)}`, { h: 0.63 });
    const body: PptxGenJS.TableRow[] = page.length
      ? page.map((p) => [
          { text: [{ text: p.probleme, options: { bold: true, breakLine: true } }, { text: p.causes ? `Causes : ${p.causes}` : "", options: { fontSize: 11, color: GREY } }], options: td({ fontSize: 13, fontFace: FONT_TITLE, margin: [4, 6, 4, 6] }) },
          { text: p.zs, options: td({ fontSize: 12, fontFace: FONT_TITLE, margin: [4, 6, 4, 6] }) },
          { text: p.solutions, options: td({ fontSize: 13, fontFace: FONT_TITLE, margin: [4, 6, 4, 6], bullet: true }) },
        ])
      : [[{ text: "Aucun problème majeur détecté par l'analyse sur ce périmètre — indicateurs conformes aux seuils.", options: td({ fontSize: 14, italic: true, color: GREY, colspan: 3, align: "center", margin: [10, 10, 10, 10] }) }]];
    s.addTable([head, ...body], { x: 0.77, y: 0.93, w: W - 1.54, colW: colW(W - 1.54, [4.6, 2.6, 4.6]), border: { type: "solid", color: BLACK, pt: 0.75 }, rowH: 0.5, valign: "middle", autoPage: false });
  });
}

/* ─── 17. Merci ────────────────────────────────────────────────────────── */

function buildMerci(ctx: Ctx): void {
  const { pptx, assets } = ctx;
  const s = pptx.addSlide();
  s.background = { color: WHITE };
  if (assets.merci) s.addImage({ data: assets.merci, x: 0.32, y: 0.3, w: 7.0, h: 6.73 });
  s.addText("Merci", { x: 6.98, y: 3.05, w: 4.34, h: 1.39, fontSize: 88, bold: true, color: TITLE_BLUE, fontFace: FONT_TITLE, align: "center", valign: "middle" });
}

// Export utilitaire pour d'autres modules.
export { fmtInt, fmtPct, fmtNum, covColor, lossColor };
export type { ZSSupervision };
