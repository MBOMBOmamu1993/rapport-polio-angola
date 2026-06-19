/**
 * Génération du rapport PowerPoint « Journées Nationales de Vaccination (JNV)
 * contre la polio » — reproduit fidèlement le modèle officiel Kwango en ne gardant
 * que la composante polio (nVPO2 et VPOb, co-administration). Mise en page
 * Power BI : bandeaux navy, tableau par ZS × par jour avec coloration des cellules
 * selon les seuils 0–80 %, 80–95 %, 95–100 %, > 100 %.
 */

import PptxGenJS from "pptxgenjs";

/* ─── Types contrat ─────────────────────────────────────────────────────── */

export interface CoverageDailyRow {
  unit: string;
  cibleCampagne: number;
  cibleJournaliere: number;
  daily: { vaccines: number; couvJour: number | null }[];
  totalVacc: number;
  couvGlobale: number | null;
}

export interface GestionRow {
  unit: string;
  flaconsRecus: number;
  flaconsUtil: number;
  flaconsRendus: number;
  perdus: number;
  vacc: number;
  taux: number | null;
}

export interface UnitValue {
  unit: string;
  value: number | null;
}

export interface CompletudeRow {
  unit: string;
  attendus: number;
  recus: number;
  couv: number | null;
  daily: { recus: number; attendus: number; couv: number | null }[];
}

export interface ProblemeRow {
  probleme: string;
  causes: string;
  zs: string;
  solutions: string;
}

export interface ReportData {
  province: string;
  /** Entité de plus bas niveau filtrée par l'utilisateur (page de garde). */
  coverEntity: string;
  periode: string;
  dateLabel: string;
  scopeLabel: string;
  byUnitLabel: string;
  jourLabels: string[];
  nbJours: number;
  saillants: {
    completude: number | null;
    completudeAttendus: number;
    completudeRecus: number;
    nvpo2Vacc: number;
    nvpo2Cible: number;
    nvpo2CV: number | null;
    vpobVacc: number;
    vpobCible: number;
    vpobCV: number | null;
    nvpo2FlaconsRecus: number;
    nvpo2FlaconsUtil: number;
    nvpo2FlaconsRendus: number;
    nvpo2Perdus: number;
    nvpo2Perte: number | null;
    vpobFlaconsRecus: number;
    vpobFlaconsUtil: number;
    vpobFlaconsRendus: number;
    vpobPerdus: number;
    vpobPerte: number | null;
    recup: number;
    mapiMineures: number;
    mapiGraves: number;
  };
  completudeByUnit: CompletudeRow[];
  nvpo2Daily: CoverageDailyRow[];
  vpobDaily: CoverageDailyRow[];
  nvpo2Gestion: GestionRow[];
  vpobGestion: GestionRow[];
  recupByUnit: UnitValue[];
  /** Récupération PEV — enfants vaccinés par antigène. */
  antigenLabels: string[];
  /** Par unité : `ev` = enfants récupérés (vaccinés), `ident` = enfants identifiés — ordre = antigenLabels. */
  recupAntigenByUnit: { unit: string; ev: number[]; ident: number[] }[];
  /** Totaux « enfants récupérés » par antigène. */
  recupAntigenTotals: number[];
  /** Totaux « enfants identifiés » par antigène. */
  recupAntigenIdentTotals: number[];
  /** Surveillance des MPV (cas notifiés). */
  survByUnit: { unit: string; pfa: number; rougeole: number; fj: number; tnn: number }[];
  survTotals: { pfa: number; rougeole: number; fj: number; tnn: number };
  problemes: ProblemeRow[];
  /** Carte PNG de localisation : zones couvertes surlignées sur la RDC. */
  scopeMapPng?: string | null;
}

/* ─── Design tokens (navy / Power BI) ──────────────────────────────────── */

const NAVY = "002A72";      // bandeau principal (bleu OMS)
const NAVY_DEEP = "001B4D"; // bandeau secondaire
const NAVY_DARKER = "00112F";
const ACCENT = "2563EB";    // bleu vif (CTA)
const ACCENT_LIGHT = "DBEAFE";
const GREY = "64748B";
const GREY_BG = "EEF2F7";
const SOFT = "F6F8FB";

// Seuils de couverture / complétude — alignés sur le modèle :
//  rouge  < 80 %, jaune 80–95 %, vert 95–100 %, bleu > 100 %
const THR_LOW = "E23636";
const THR_MID = "F1C40F";
const THR_HIGH = "22B457";
const THR_FULL = "1D4ED8";
const THR_NONE = "CBD5E1";

const W = 13.333;
const H = 7.5;

/* ─── Helpers ──────────────────────────────────────────────────────────── */

async function toDataUrl(path: string): Promise<string | null> {
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

function fmtPct(n: number | null, d = 2): string {
  if (n === null || !Number.isFinite(n)) return "—";
  return `${n.toFixed(d).replace(".", ",")} %`;
}
function fmtInt(n: number): string {
  return Math.round(n || 0).toLocaleString("fr-FR").replace(/\u202f/g, " ");
}
function fmtNum(n: number): string {
  return (Math.round((n + Number.EPSILON) * 100) / 100).toLocaleString("fr-FR").replace(/\u202f/g, " ");
}

/**
 * Nomme le niveau d'agr\u00e9gation courant au lieu du mot g\u00e9n\u00e9rique \u00ab unit\u00e9 \u00bb.
 * `byUnitLabel` vaut \u00ab Province \u00bb, \u00ab Antenne \u00bb, \u00ab Zone de Sant\u00e9 \u00bb ou \u00ab Aire de
 * Sant\u00e9 \u00bb ; on renvoie ce nom au singulier ou au pluriel, en minuscules, pour
 * \u00e9crire des commentaires qui collent au tableau comment\u00e9.
 */
function levelWord(byUnitLabel: string, plural = false): string {
  const l = byUnitLabel.toLowerCase().trim();
  if (!plural) return l;
  // Pluralise le premier mot : \u00ab zone de sant\u00e9 \u00bb \u2192 \u00ab zones de sant\u00e9 \u00bb.
  return l.replace(/^(\S+)/, "$1s");
}

/** \u00ab 1 zone de sant\u00e9 \u00bb / \u00ab 3 zones de sant\u00e9 \u00bb avec accord du nombre. */
function countLevel(n: number, byUnitLabel: string): string {
  return `${n} ${levelWord(byUnitLabel, n > 1)}`;
}

/** Joint une liste de noms avec \u00ab et \u00bb avant le dernier \u00e9l\u00e9ment. */
function joinAnd(names: string[]): string {
  if (names.length <= 1) return names.join("");
  return `${names.slice(0, -1).join(", ")} et ${names[names.length - 1]}`;
}

function thresholdColor(v: number | null): string {
  if (v === null || !Number.isFinite(v)) return THR_NONE;
  if (v > 100) return THR_FULL;
  if (v >= 95) return THR_HIGH;
  if (v >= 80) return THR_MID;
  return THR_LOW;
}
function lossColor(v: number | null): string {
  if (v === null || !Number.isFinite(v)) return THR_NONE;
  if (v < 0 || v > 15) return THR_LOW;
  if (v > 10) return THR_MID;
  return THR_HIGH;
}
function slug(s: string): string {
  return s.replace(/[^a-zA-Z0-9]+/g, "_").replace(/^_|_$/g, "").slice(0, 40) || "rapport";
}

/* ─── Génération du rapport ────────────────────────────────────────────── */

export async function exportReportPPT(data: ReportData): Promise<void> {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: "WIDE", width: W, height: H });
  pptx.layout = "WIDE";
  pptx.author = "PEV — RD Congo";
  pptx.company = "Programme Élargi de Vaccination";
  pptx.title = `Rapport polio – ${data.province}`;

  const [cover, pev] = await Promise.all([
    toDataUrl("/cover-polio.png"),
    toDataUrl("/logo/pev-officiel.png"),
  ]);

  const ctx: SlideCtx = { pptx, data, pev, addHeader: addHeaderFactory(pptx, data, pev) };

  buildCover(ctx, cover);
  buildPlan(ctx);
  buildPointsSaillants(ctx);
  buildCompletude(ctx);
  buildCompletudeMap(ctx);
  buildCoverage(ctx, "nVPO2", data.nvpo2Daily, data.saillants.nvpo2CV);
  buildCoverage(ctx, "VPOb", data.vpobDaily, data.saillants.vpobCV);
  buildRecup(ctx);
  buildGestion(ctx, "nVPO2", data.nvpo2Gestion, data.saillants.nvpo2Perte, 11);
  buildGestion(ctx, "VPOb", data.vpobGestion, data.saillants.vpobPerte, 10);
  buildSurveillanceMPV(ctx);
  buildMapi(ctx);
  buildProblemes(ctx);
  buildMerci(ctx);

  await pptx.writeFile({ fileName: `Rapport_JNV_Polio_${slug(data.scopeLabel)}.pptx` });
}

/* ─── Layout commun ────────────────────────────────────────────────────── */

interface SlideCtx {
  pptx: PptxGenJS;
  data: ReportData;
  pev: string | null;
  addHeader: (s: PptxGenJS.Slide, title: string, subtitle?: string) => void;
}

function addHeaderFactory(
  pptx: PptxGenJS,
  data: ReportData,
  pev: string | null
): (s: PptxGenJS.Slide, title: string, subtitle?: string) => void {
  return (s, title, subtitle) => {
    s.background = { color: "FFFFFF" };
    // Bandeau navy plein largeur.
    s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: 1.05, fill: { color: NAVY } });
    s.addShape(pptx.ShapeType.rect, { x: 0, y: 1.05, w: W, h: 0.09, fill: { color: ACCENT } });
    s.addText(title, {
      x: 0.55, y: 0.06, w: W - 2.6, h: subtitle ? 0.6 : 0.95,
      fontSize: 26, bold: true, color: "FFFFFF", align: "left",
      valign: subtitle ? "top" : "middle", fontFace: "Calibri",
    });
    if (subtitle) {
      s.addText(subtitle, {
        x: 0.55, y: 0.66, w: W - 2.6, h: 0.36,
        fontSize: 14, color: "CCE4FF", align: "left", italic: true, fontFace: "Calibri",
      });
    }
    if (pev) {
      s.addShape(pptx.ShapeType.roundRect, { x: W - 2.42, y: 0.22, w: 2.17, h: 0.6, fill: { color: "FFFFFF" }, line: { color: "FFFFFF", width: 0 }, rectRadius: 0.05 });
      s.addImage({ data: pev, x: W - 2.32, y: 0.37, w: 1.97, h: 0.302 });
    }

    // Pied de page.
    s.addText("Rapport JNV Polio — nVPO2 & VPOb (co-administration)", {
      x: 0.55, y: H - 0.36, w: W - 4.2, h: 0.28, fontSize: 10, color: GREY, align: "left",
    });
    s.addText(data.scopeLabel, {
      x: W - 4.2, y: H - 0.36, w: 3.8, h: 0.28, fontSize: 10, color: NAVY, align: "right", bold: true,
    });
  };
}

/* ─── Slide 1 : Page de garde ──────────────────────────────────────────── */

function buildCover(ctx: SlideCtx, cover: string | null): void {
  const { pptx, data, pev } = ctx;
  const s = pptx.addSlide();
  s.background = { color: NAVY_DARKER };
  if (cover) {
    s.addImage({ data: cover, x: W * 0.40, y: 0, w: W * 0.60, h: H, sizing: { type: "cover", w: W * 0.60, h: H } });
    // Voile sombre côté droite pour lisibilité.
    s.addShape(pptx.ShapeType.rect, { x: W * 0.40, y: 0, w: W * 0.60, h: H, fill: { color: NAVY_DARKER, transparency: 70 } });
  }
  // Pan gauche navy.
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W * 0.46, h: H, fill: { color: NAVY_DEEP } });
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.16, h: H, fill: { color: ACCENT } });

  // Logo PEV (lockup horizontal) sur pastille blanche pour la lisibilité sur le navy.
  if (pev) {
    s.addShape(pptx.ShapeType.roundRect, { x: 0.55, y: 0.5, w: 3.8, h: 0.85, fill: { color: "FFFFFF" }, line: { color: "FFFFFF", width: 0 }, rectRadius: 0.08 });
    s.addImage({ data: pev, x: 0.75, y: 0.665, w: 3.4, h: 0.521 });
  }

  // Eyebrow.
  s.addText("RAPPORT DES RÉSULTATS", {
    x: 0.6, y: 1.95, w: 5.4, h: 0.45,
    fontSize: 18, color: ACCENT_LIGHT, bold: true, charSpacing: 4, fontFace: "Calibri",
  });
  // Titre principal.
  s.addText("Journées Nationales de Vaccination contre la polio", {
    x: 0.6, y: 2.45, w: 5.4, h: 1.75,
    fontSize: 32, color: "FFFFFF", bold: true, fontFace: "Calibri", lineSpacingMultiple: 1.05,
  });
  s.addShape(pptx.ShapeType.rect, { x: 0.62, y: 4.25, w: 1.6, h: 0.07, fill: { color: ACCENT } });

  s.addText(
    [
      { text: data.coverEntity, options: { bold: true, fontSize: 22, color: "FFFFFF", breakLine: true } },
      { text: data.periode || "Période de la campagne", options: { fontSize: 17, color: "CCE4FF", breakLine: true } },
      { text: "Vaccins : nVPO2 et VPOb (co-administration)", options: { fontSize: 15, color: "93C5FD" } },
    ],
    { x: 0.6, y: 4.5, w: 5.5, h: 1.6, valign: "top", lineSpacingMultiple: 1.2 }
  );

  s.addText("Programme Élargi de Vaccination — RD Congo", {
    x: 0.6, y: H - 0.72, w: 5.4, h: 0.4,
    fontSize: 14, color: "93C5FD", italic: true,
  });
  s.addText(data.dateLabel, {
    x: W * 0.42 + 0.2, y: H - 0.72, w: W * 0.58 - 0.4, h: 0.4,
    fontSize: 14, color: "FFFFFF", align: "right", bold: true,
  });
}

/* ─── Slide 2 : Plan ───────────────────────────────────────────────────── */

function buildPlan(ctx: SlideCtx): void {
  const { pptx } = ctx;
  const s = pptx.addSlide();
  ctx.addHeader(s, "Plan de présentation");
  const items = [
    "Points saillants de la campagne",
    "Complétude des rapports par Zone de Santé",
    "Couvertures vaccinales nVPO2 par jour de campagne",
    "Couvertures vaccinales VPOb par jour de campagne",
    "Récupération PEV de routine — enfants vaccinés par antigène",
    "Gestion du vaccin nVPO2 — flacons & taux de perte",
    "Gestion du vaccin VPOb — flacons & taux de perte",
    "Surveillance des MPV par Zone de Santé",
    "Synthèse Surveillance MAPI et récupération en routine",
    "Problèmes rencontrés / Actions correctrices",
  ];
  items.forEach((it, i) => {
    const y = 1.55 + i * 0.55;
    s.addShape(pptx.ShapeType.ellipse, { x: 1.2, y, w: 0.5, h: 0.5, fill: { color: ACCENT } });
    s.addText(String(i + 1), { x: 1.2, y, w: 0.5, h: 0.5, align: "center", valign: "middle", color: "FFFFFF", bold: true, fontSize: 18 });
    s.addText(it, { x: 1.9, y, w: 10.5, h: 0.5, valign: "middle", fontSize: 19, color: NAVY });
  });
}

/* ─── Slide 3 : Points saillants ──────────────────────────────────────── */

type SaillantLine =
  | { kind: "head"; text: string }
  | { kind: "item"; label: string; value: string; tone?: string };

function buildPointsSaillants(ctx: SlideCtx): void {
  const { pptx, data } = ctx;
  const s = pptx.addSlide();
  ctx.addHeader(s, "Points saillants", data.periode || undefined);
  const sa = data.saillants;

  const dosesTotal = sa.nvpo2Vacc + sa.vpobVacc;
  const mapiTotal = sa.mapiMineures + sa.mapiGraves;
  const mapiPour100k = dosesTotal > 0 ? (mapiTotal / dosesTotal) * 100000 : null;

  const left: SaillantLine[] = [
    { kind: "head", text: "COMPLÉTUDE" },
    { kind: "item", label: "Rapports attendus", value: fmtInt(sa.completudeAttendus) },
    { kind: "item", label: "Rapports reçus", value: fmtInt(sa.completudeRecus) },
    { kind: "item", label: "Complétude des rapports", value: fmtPct(sa.completude), tone: thresholdColor(sa.completude) },
    { kind: "head", text: "VACCINATION nVPO2" },
    { kind: "item", label: "Vaccinés", value: fmtInt(sa.nvpo2Vacc) },
    { kind: "item", label: "Cible", value: fmtInt(sa.nvpo2Cible) },
    { kind: "item", label: "CV", value: fmtPct(sa.nvpo2CV), tone: thresholdColor(sa.nvpo2CV) },
    { kind: "head", text: "VACCINATION VPOb" },
    { kind: "item", label: "Vaccinés", value: fmtInt(sa.vpobVacc) },
    { kind: "item", label: "Cible", value: fmtInt(sa.vpobCible) },
    { kind: "item", label: "CV", value: fmtPct(sa.vpobCV), tone: thresholdColor(sa.vpobCV) },
  ];
  const right: SaillantLine[] = [
    { kind: "head", text: "GESTION DE VACCIN nVPO2" },
    { kind: "item", label: "Flacons utilisés", value: fmtInt(sa.nvpo2FlaconsUtil) },
    { kind: "item", label: "Taux de perte", value: fmtPct(sa.nvpo2Perte), tone: lossColor(sa.nvpo2Perte) },
    { kind: "head", text: "GESTION DE VACCIN VPOb" },
    { kind: "item", label: "Flacons utilisés", value: fmtInt(sa.vpobFlaconsUtil) },
    { kind: "item", label: "Taux de perte", value: fmtPct(sa.vpobPerte), tone: lossColor(sa.vpobPerte) },
    { kind: "head", text: "NOTIFICATION MAPI" },
    { kind: "item", label: "MAPI non grave", value: fmtInt(sa.mapiMineures) },
    { kind: "item", label: "MAPI grave", value: fmtInt(sa.mapiGraves), tone: sa.mapiGraves ? THR_LOW : NAVY_DEEP },
    { kind: "item", label: "% MAPI pour 100 000 doses", value: mapiPour100k == null ? "—" : fmtNum(mapiPour100k) },
  ];

  const box = (x: number, w: number, lines: SaillantLine[]) => {
    const y0 = 1.3;
    const h = 5.7;
    s.addShape(pptx.ShapeType.roundRect, { x, y: y0, w, h, fill: { color: "FFFFFF" }, line: { color: NAVY, width: 1.5 }, rectRadius: 0.06 });
    const runs: PptxGenJS.TextProps[] = [];
    lines.forEach((ln) => {
      if (ln.kind === "head") {
        runs.push({ text: ln.text, options: { bold: true, underline: { style: "sng" }, fontSize: 18, color: NAVY, breakLine: true, paraSpaceBefore: 8 } as PptxGenJS.TextPropsOptions });
      } else {
        runs.push({ text: `   ❖  ${ln.label} : `, options: { fontSize: 16, color: NAVY_DEEP } });
        runs.push({ text: ln.value, options: { fontSize: 16, bold: true, color: ln.tone ?? NAVY_DEEP, breakLine: true } });
      }
    });
    s.addText(runs, { x: x + 0.25, y: y0 + 0.2, w: w - 0.5, h: h - 0.4, valign: "top", fontFace: "Calibri", lineSpacingMultiple: 1.18 });
  };

  box(0.45, 6.1, left);
  box(6.78, 6.1, right);
}

/* ─── Slide 4 : Complétude des rapports ────────────────────────────────── */

function buildCompletude(ctx: SlideCtx): void {
  const { pptx, data } = ctx;
  const sa = data.saillants;
  const globalCV = sa.completude;
  const arcColor = thresholdColor(globalCV);
  const days = data.jourLabels;

  const drawKpi = (s: PptxGenJS.Slide) => {
    s.addShape(pptx.ShapeType.roundRect, { x: 0.5, y: 1.95, w: 6.2, h: 4.0, fill: { color: SOFT }, line: { color: "DEE5EE", width: 1 }, rectRadius: 0.08 });
    s.addText("Complétude globale", { x: 0.7, y: 2.0, w: 5.8, h: 0.45, fontSize: 18, bold: true, color: NAVY });
    s.addShape(pptx.ShapeType.ellipse, { x: 0.95, y: 2.6, w: 2.4, h: 2.4, fill: { color: arcColor }, line: { color: arcColor, width: 0 } });
    s.addShape(pptx.ShapeType.ellipse, { x: 1.15, y: 2.8, w: 2.0, h: 2.0, fill: { color: "FFFFFF" }, line: { color: "FFFFFF", width: 0 } });
    s.addText(fmtPct(globalCV), { x: 0.95, y: 3.4, w: 2.4, h: 0.7, align: "center", valign: "middle", fontSize: 30, bold: true, color: arcColor });
    s.addText("complétude\nrapports", { x: 0.95, y: 4.05, w: 2.4, h: 0.5, align: "center", valign: "top", fontSize: 13, color: GREY });
    s.addShape(pptx.ShapeType.roundRect, { x: 3.55, y: 2.85, w: 3.05, h: 1.0, fill: { color: "FFFFFF" }, line: { color: "DEE5EE", width: 1 }, rectRadius: 0.06 });
    s.addText("Rapports reçus", { x: 3.7, y: 2.9, w: 2.8, h: 0.35, fontSize: 13, color: GREY });
    s.addText(fmtInt(sa.completudeRecus), { x: 3.7, y: 3.22, w: 2.8, h: 0.6, fontSize: 26, bold: true, color: NAVY_DEEP });
    s.addShape(pptx.ShapeType.roundRect, { x: 3.55, y: 3.95, w: 3.05, h: 1.0, fill: { color: "FFFFFF" }, line: { color: "DEE5EE", width: 1 }, rectRadius: 0.06 });
    s.addText("Rapports attendus", { x: 3.7, y: 4.0, w: 2.8, h: 0.35, fontSize: 13, color: GREY });
    s.addText(fmtInt(sa.completudeAttendus), { x: 3.7, y: 4.32, w: 2.8, h: 0.6, fontSize: 26, bold: true, color: NAVY_DEEP });
    s.addText("Une bonne complétude (≥ 95 %) conditionne la fiabilité des couvertures vaccinales calculées.", {
      x: 0.7, y: 5.1, w: 5.8, h: 0.75, fontSize: 13, color: GREY, italic: true,
    });
  };

  const head: PptxGenJS.TableCell[] = [
    { text: data.byUnitLabel, options: thHeader() },
    { text: "Attendus", options: thHeader() },
    ...days.flatMap((d) => [
      { text: `Reçus ${d}`, options: thHeader() } as PptxGenJS.TableCell,
      { text: `Compl. ${d}`, options: thHeader() } as PptxGenJS.TableCell,
    ]),
    { text: "Compl. Globale", options: thHeader() },
  ];
  const totalAttendus = data.completudeByUnit.reduce((a, r) => a + r.attendus, 0);
  const totalRecus = data.completudeByUnit.reduce((a, r) => a + r.recus, 0);
  const totalsDaily = days.map((_, i) =>
    data.completudeByUnit.reduce((a, r) => a + (r.daily[i]?.recus ?? 0), 0)
  );
  const totalsAttendusDaily = days.map((_, i) =>
    data.completudeByUnit.reduce((a, r) => a + (r.daily[i]?.attendus ?? 0), 0)
  );
  const bodyRows: PptxGenJS.TableRow[] = data.completudeByUnit.map((r) => {
    const cells: PptxGenJS.TableCell[] = [
      { text: r.unit, options: tdCell({ bold: true }) },
      { text: fmtInt(r.attendus), options: tdCell({ align: "right" }) },
    ];
    r.daily.forEach((d) => {
      cells.push({ text: fmtInt(d.recus), options: tdCell({ align: "right" }) });
      cells.push({ text: fmtPct(d.couv, 2), options: tdCell({ align: "right", bold: true, fill: { color: thresholdColor(d.couv) }, color: "FFFFFF" }) });
    });
    cells.push({ text: fmtPct(r.couv, 2), options: tdCell({ align: "right", bold: true, fill: { color: thresholdColor(r.couv) }, color: "FFFFFF" }) });
    return cells;
  });
  const totalRow: PptxGenJS.TableRow = [
    { text: "Total", options: thTotal() },
    { text: fmtInt(totalAttendus), options: thTotal({ align: "right" }) },
    ...days.flatMap((_, i): PptxGenJS.TableCell[] => {
      const attJour = totalsAttendusDaily[i];
      const couvJour = attJour > 0 ? (totalsDaily[i] / attJour) * 100 : null;
      return [
        { text: fmtInt(totalsDaily[i]), options: thTotal({ align: "right" }) },
        { text: fmtPct(couvJour, 2), options: thTotal({ align: "right" }) },
      ];
    }),
    { text: fmtPct(totalAttendus > 0 ? (totalRecus / totalAttendus) * 100 : null, 2), options: thTotal({ align: "right" }) },
  ];

  const nCols = 2 + days.length * 2 + 1;
  const weights = [1.4, 0.9, ...days.flatMap(() => [0.8, 0.8]), 0.9];
  const colWNarrow = computeColW(W - 7.4, nCols, weights);
  const colWWide = computeColW(W - 0.9, nCols, weights);

  const allRows = [...bodyRows, totalRow];
  const rowH = 0.34;
  // 1ère diapo : table à droite (colonne étroite) à côté du KPI ; suivantes : pleine largeur.
  const perPageFirst = rowsPerPage(2.55, 5.9, 0.6, rowH);
  const perPageNext = rowsPerPage(1.45, 6.4, 0.6, rowH);
  const pages: PptxGenJS.TableRow[][] = [];
  pages.push(allRows.slice(0, perPageFirst));
  for (let i = perPageFirst; i < allRows.length; i += perPageNext) pages.push(allRows.slice(i, i + perPageNext));

  pages.forEach((pageRows, idx) => {
    const s = pptx.addSlide();
    ctx.addHeader(s, `Complétude des rapports journaliers et globale${suiteSuffix(idx, pages.length)}`, "Source : Synthèse du masque de saisie");
    addLegend(pptx, s, 0.5, 1.2);
    if (idx === 0) {
      drawKpi(s);
      s.addShape(pptx.ShapeType.roundRect, { x: 6.9, y: 1.95, w: W - 7.4, h: 4.0, fill: { color: "FFFFFF" }, line: { color: "DEE5EE", width: 1 }, rectRadius: 0.08 });
      s.addText(`Complétude par ${data.byUnitLabel}`, { x: 7.05, y: 2.0, w: 5.5, h: 0.45, fontSize: 18, bold: true, color: NAVY });
      s.addTable([head, ...pageRows], {
        x: 7.0, y: 2.55, w: W - 7.4, colW: colWNarrow,
        border: { type: "solid", color: "DEE5EE", pt: 0.5 },
        rowH, valign: "middle", fontFace: "Calibri",
      });
    } else {
      s.addText(`Complétude par ${data.byUnitLabel}`, { x: 0.5, y: 1.62, w: 8, h: 0.4, fontSize: 16, bold: true, color: NAVY });
      s.addTable([head, ...pageRows], {
        x: 0.45, y: 2.0, w: W - 0.9, colW: colWWide,
        border: { type: "solid", color: "DEE5EE", pt: 0.5 },
        rowH, valign: "middle", fontFace: "Calibri",
      });
    }
  });
}

/* ─── Slide 4 bis : Spatialisation de la complétude (carte RDC / ZS) ────── */

function buildCompletudeMap(ctx: SlideCtx): void {
  const { pptx, data } = ctx;
  const s = pptx.addSlide();
  ctx.addHeader(s, "Spatialisation de la complétude", "Carte de la RD Congo — Zones de Santé du périmètre colorées selon leur complétude");

  // Légende des seuils de complétude.
  s.addText("Complétude :", { x: 0.55, y: 1.24, w: 1.7, h: 0.38, fontSize: 14, color: NAVY, bold: true, valign: "middle" });
  const leg: { c: string; t: string }[] = [
    { c: THR_LOW, t: "< 60 %" },
    { c: THR_MID, t: "60 – 79,9 %" },
    { c: THR_HIGH, t: "≥ 80 %" },
  ];
  leg.forEach((it, i) => {
    const cx = 2.3 + i * 2.1;
    s.addShape(pptx.ShapeType.rect, { x: cx, y: 1.28, w: 0.42, h: 0.3, fill: { color: it.c } });
    s.addText(it.t, { x: cx + 0.47, y: 1.24, w: 1.5, h: 0.38, fontSize: 13, color: NAVY, valign: "middle" });
  });
  s.addShape(pptx.ShapeType.rect, { x: 0.55, y: 1.68, w: 0.42, h: 0.3, fill: { color: "F1F5F9" }, line: { color: "94A3B8", width: 0.75 } });
  s.addText("Hors périmètre", { x: 1.05, y: 1.64, w: 3.0, h: 0.38, fontSize: 13, color: GREY, valign: "middle" });

  const png = data.scopeMapPng;
  if (png) {
    // Fond cartographique RDC (1100×1000 ≈ ratio 1.1).
    const h = 5.15;
    const w = h * 1.1;
    s.addImage({ data: png, x: (W - w) / 2, y: 2.1, w, h });
  } else {
    s.addShape(pptx.ShapeType.roundRect, {
      x: 2.0, y: 3.0, w: W - 4.0, h: 1.8,
      fill: { color: SOFT }, line: { color: "DEE5EE", width: 1 }, rectRadius: 0.1,
    });
    s.addText(
      "Carte indisponible — le fond cartographique des Zones de Santé n'a pas pu être chargé au moment de la génération (une connexion Internet est requise).",
      { x: 2.3, y: 3.0, w: W - 4.6, h: 1.8, align: "center", valign: "middle", fontSize: 15, italic: true, color: GREY }
    );
  }
}

function addLegend(pptx: PptxGenJS, s: PptxGenJS.Slide, x: number, y: number): void {
  const items: { c: string; t: string }[] = [
    { c: THR_LOW,  t: "< 80 %" },
    { c: THR_MID,  t: "80 – 95 %" },
    { c: THR_HIGH, t: "95 – 100 %" },
    { c: THR_FULL, t: "> 100 %" },
  ];
  s.addText("Critères de graduation des couvertures :", {
    x, y, w: 3.8, h: 0.38, fontSize: 14, color: NAVY, bold: true, valign: "middle",
  });
  items.forEach((it, i) => {
    const cx = x + 3.9 + i * 1.95;
    s.addShape(pptx.ShapeType.rect, { x: cx, y: y + 0.04, w: 0.42, h: 0.3, fill: { color: it.c } });
    s.addText(it.t, { x: cx + 0.47, y, w: 1.5, h: 0.38, fontSize: 13, color: NAVY, valign: "middle" });
  });
}

/* ─── Slides 5/6 : Couvertures par jour ────────────────────────────────── */

function buildCoverage(
  ctx: SlideCtx,
  vaccine: "nVPO2" | "VPOb",
  rows: CoverageDailyRow[],
  globalCV: number | null
): void {
  const { pptx, data } = ctx;
  const days = data.jourLabels;
  // En-tête du tableau, identique au modèle :
  // ZS | Cible Campagne | Cible Polio Journalière | Vacc. J1 | Couvert. J1 | … | Couvert. Globale
  const head: PptxGenJS.TableCell[] = [
    { text: data.byUnitLabel, options: thHeader() },
    { text: "Cible\nCampagne", options: thHeader() },
    { text: "Cible Polio\nJournalière", options: thHeader() },
    ...days.flatMap((d): PptxGenJS.TableCell[] => [
      { text: `Vaccinés\n${d}`, options: thHeader() },
      { text: `Couvert.\n${d}`, options: thHeader() },
    ]),
    { text: "Couvert.\nGlobale", options: thHeader() },
  ];

  // Totaux du tableau.
  const totalCible = rows.reduce((a, r) => a + r.cibleCampagne, 0);
  const totalCibleJ = rows.reduce((a, r) => a + r.cibleJournaliere, 0);
  const totalsDaily = days.map((_, i) => rows.reduce((a, r) => a + (r.daily[i]?.vaccines ?? 0), 0));
  const totalVacc = rows.reduce((a, r) => a + r.totalVacc, 0);
  const totalCouv = totalCible > 0 ? (totalVacc / totalCible) * 100 : null;

  const bodyRows: PptxGenJS.TableRow[] = rows.map((r): PptxGenJS.TableCell[] => {
    const cells: PptxGenJS.TableCell[] = [
      { text: r.unit, options: tdCell({ bold: true }) },
      { text: fmtInt(r.cibleCampagne), options: tdCell({ align: "right" }) },
      { text: fmtInt(r.cibleJournaliere), options: tdCell({ align: "right" }) },
    ];
    r.daily.forEach((d) => {
      cells.push({ text: fmtInt(d.vaccines), options: tdCell({ align: "right" }) });
      cells.push({ text: fmtPct(d.couvJour, 2), options: tdCell({ align: "right", bold: true, fill: { color: thresholdColor(d.couvJour) }, color: "FFFFFF" }) });
    });
    cells.push({ text: fmtPct(r.couvGlobale, 2), options: tdCell({ align: "right", bold: true, fill: { color: thresholdColor(r.couvGlobale) }, color: "FFFFFF" }) });
    return cells;
  });
  const totalRow: PptxGenJS.TableRow = [
    { text: "Total", options: thTotal() },
    { text: fmtInt(totalCible), options: thTotal({ align: "right" }) },
    { text: fmtInt(totalCibleJ), options: thTotal({ align: "right" }) },
    ...days.flatMap((_, i): PptxGenJS.TableCell[] => {
      const couvJ = totalCible > 0 ? (totalsDaily[i] / totalCible) * 100 : null;
      return [
        { text: fmtInt(totalsDaily[i]), options: thTotal({ align: "right" }) },
        { text: fmtPct(couvJ, 2), options: thTotal({ align: "right" }) },
      ];
    }),
    { text: fmtPct(totalCouv, 2), options: thTotal({ align: "right" }) },
  ];

  const nCols = 3 + days.length * 2 + 1;
  const weights = [1.6, 0.85, 0.95, ...days.flatMap(() => [0.85, 0.8]), 0.95];
  const colW = computeColW(W - 1.1, nCols, weights);

  const tableY = 1.65;
  const rowH = 0.36;
  const perPage = rowsPerPage(tableY, 5.9, 0.62, rowH);
  const pages = chunkRows([...bodyRows, totalRow], perPage);
  pages.forEach((pageRows, idx) => {
    const s = pptx.addSlide();
    ctx.addHeader(s, `Couvertures vaccinales ${vaccine}, par ${data.byUnitLabel}${suiteSuffix(idx, pages.length)}`, "Source : Synthèse du masque de saisie");
    addLegend(pptx, s, 0.5, 1.15);
    s.addTable([head, ...pageRows], {
      x: 0.55, y: tableY, w: W - 1.1, colW,
      border: { type: "solid", color: "DEE5EE", pt: 0.5 },
      rowH, valign: "middle", fontFace: "Calibri",
    });
    addCommentBar(pptx, s, coverageComment(rows, globalCV, vaccine, data.byUnitLabel));
  });
}

/* ─── Slide 7 : Récupération PEV de routine ────────────────────────────── */

function buildRecup(ctx: SlideCtx): void {
  const { pptx, data } = ctx;
  const ant = data.antigenLabels;
  const rows = data.recupAntigenByUnit;
  const totals = data.recupAntigenTotals;
  // On affiche le tableau dès qu'il y a des enfants identifiés OU récupérés : une
  // récupération nulle partout reste informative en regard des identifiés.
  const hasData = totals.some((v) => v > 0) || data.recupAntigenIdentTotals.some((v) => v > 0);
  const totalEnfants = totals.reduce((a, b) => a + b, 0);

  if (ant.length === 0 || !hasData) {
    const s = pptx.addSlide();
    ctx.addHeader(s, "Récupération des enfants en PEV de routine", "Enfants vaccinés (EV) par antigène pendant la campagne — toutes tranches d'âge");
    s.addText("Aucune récupération PEV par antigène saisie pour ce périmètre.", {
      x: 1, y: 3.2, w: W - 2, h: 1, align: "center", fontSize: 16, italic: true, color: GREY,
    });
    addCommentBar(pptx, s, `${fmtInt(data.saillants.recup)} enfants récupérés et orientés vers le PEV de routine pendant la campagne.`);
    return;
  }

  const ident = data.recupAntigenIdentTotals;
  const totalIdentifies = ident.reduce((a, b) => a + b, 0);
  const tauxRecupGlobal = totalIdentifies > 0 ? (totalEnfants / totalIdentifies) * 100 : null;

  // % récupéré, calculé EXACTEMENT comme la colonne « CV » du masque (feuille
  // Synthèse, colonnes HQ…) : récupéré (EV) ÷ identifié. Vérifié au cellule près
  // sur l'ensemble du masque. Affiché à 1 décimale comme le masque (« 42,9 % »).
  const pctRecup = (ev: number, id: number): string => {
    const t = id > 0 ? (ev / id) * 100 : null;
    return t === null ? "—" : `${t.toFixed(1).replace(".", ",")}%`;
  };

  // Totaux par entité, tous antigènes confondus (colonnes « Total » du tableau).
  const grandIdentByUnit = rows.map((r) => r.ident.reduce((a, b) => a + b, 0));
  const grandEvByUnit = rows.map((r) => r.ev.reduce((a, b) => a + b, 0));

  // Chaque antigène pèse 3 sous-colonnes (Nb identifiés | Nb récupéré | % récupéré) :
  // le tableau des 15 antigènes est trop large pour une seule diapo. On répartit
  // donc les antigènes sur plusieurs diapos (≤ 6 par diapo) pour garder une bonne
  // lisibilité. Les 3 colonnes « Total » (tous antigènes confondus) sont répétées
  // sur chaque diapo afin que chacune soit autoportante.
  const MAX_ANT_PER_SLIDE = 6;
  const groupCount = Math.max(1, Math.ceil(ant.length / MAX_ANT_PER_SLIDE));
  const baseSize = Math.floor(ant.length / groupCount);
  let rem = ant.length % groupCount;
  const groups: { start: number; end: number }[] = [];
  let cursor = 0;
  for (let g = 0; g < groupCount; g++) {
    const size = baseSize + (rem > 0 ? 1 : 0);
    if (rem > 0) rem--;
    groups.push({ start: cursor, end: cursor + size });
    cursor += size;
  }

  const subHead = (): PptxGenJS.TableCell[] => [
    { text: "Nb\nident.", options: thHeader({ fontSize: 10 }) },
    { text: "Nb\nrécup.", options: thHeader({ fontSize: 10 }) },
    { text: "%\nrécup.", options: thHeader({ fontSize: 10 }) },
  ];
  const recCell = (ev: number, id: number): PptxGenJS.TableCell[] => [
    { text: fmtInt(id), options: tdCell({ align: "right", fontSize: 12 }) },
    { text: fmtInt(ev), options: tdCell({ align: "right", fontSize: 12 }) },
    { text: pctRecup(ev, id), options: tdCell({ align: "right", bold: true, fontSize: 11, color: ACCENT }) },
  ];
  const totCell = (id: number, ev: number, total = false): PptxGenJS.TableCell[] => {
    const base = total ? thTotal : tdCell;
    const fill = total ? undefined : { color: GREY_BG };
    return [
      { text: fmtInt(id), options: base({ align: "right", bold: true, fontSize: 12, ...(fill ? { fill } : {}) }) },
      { text: fmtInt(ev), options: base({ align: "right", bold: true, fontSize: 12, ...(fill ? { fill } : {}) }) },
      { text: pctRecup(ev, id), options: base({ align: "right", bold: true, fontSize: 11, color: total ? "FFFFFF" : ACCENT, ...(fill ? { fill } : {}) }) },
    ];
  };

  groups.forEach((g, gi) => {
    const groupAnt = ant.slice(g.start, g.end);

    // En-tête sur deux niveaux : antigènes du groupe + bloc « TOTAL ».
    const headTop: PptxGenJS.TableCell[] = [
      { text: data.byUnitLabel, options: thHeader({ rowspan: 2, fontSize: 12 }) },
      ...groupAnt.map((a): PptxGenJS.TableCell => ({ text: a, options: thHeader({ colspan: 3, fontSize: 13 }) })),
      { text: "TOTAL (tous antigènes)", options: thHeader({ colspan: 3, fontSize: 11, fill: { color: NAVY_DEEP } }) },
    ];
    const headSub: PptxGenJS.TableCell[] = [
      ...groupAnt.flatMap(() => subHead()),
      { text: "Nb total\nidentifiés", options: thHeader({ fontSize: 10, fill: { color: NAVY_DEEP } }) },
      { text: "Nb total\nrécupéré", options: thHeader({ fontSize: 10, fill: { color: NAVY_DEEP } }) },
      { text: "%\nrécupération", options: thHeader({ fontSize: 10, fill: { color: NAVY_DEEP } }) },
    ];

    const bodyRows: PptxGenJS.TableRow[] = rows.map((r, ri): PptxGenJS.TableCell[] => [
      { text: r.unit, options: tdCell({ bold: true, fontSize: 11 }) },
      ...groupAnt.flatMap((_, k) => recCell(r.ev[g.start + k] ?? 0, r.ident?.[g.start + k] ?? 0)),
      ...totCell(grandIdentByUnit[ri], grandEvByUnit[ri]),
    ]);
    const totalRow: PptxGenJS.TableRow = [
      { text: "Total", options: thTotal({ fontSize: 12 }) },
      ...groupAnt.flatMap((_, k): PptxGenJS.TableCell[] => {
        const j = g.start + k;
        const id = ident[j] ?? 0;
        const ev = totals[j] ?? 0;
        return [
          { text: fmtInt(id), options: thTotal({ align: "right", fontSize: 12 }) },
          { text: fmtInt(ev), options: thTotal({ align: "right", fontSize: 12 }) },
          { text: pctRecup(ev, id), options: thTotal({ align: "right", fontSize: 11 }) },
        ];
      }),
      ...totCell(totalIdentifies, totalEnfants, true),
    ];

    const nCols = 1 + groupAnt.length * 3 + 3;
    const colW = computeColW(W - 0.5, nCols, [2.1, ...groupAnt.flatMap(() => [0.55, 0.55, 0.62]), 0.78, 0.78, 0.82]);

    const tableY = 1.5;
    const rowH = 0.42;
    // Pagination volontairement prudente : on réserve une marge sous le tableau
    // (≈ 0,5″/ligne pour les libellés sur deux lignes) afin que le tableau s'arrête
    // NETTEMENT au-dessus du bandeau « Commentaire » (y ≈ 6,05) et ne masque jamais
    // les dernières entités.
    const perPage = Math.max(4, Math.floor((5.75 - tableY - 0.95) / 0.46));
    const pages = chunkRows([...bodyRows, totalRow], perPage);
    pages.forEach((pageRows, idx) => {
      const s = pptx.addSlide();
      // Titre court (constant) : tout le contexte de pagination passe en sous-titre
      // pour éviter que le titre déborde sur deux lignes et chevauche le sous-titre.
      const grp = groupCount > 1 ? ` (groupe ${gi + 1}/${groupCount})` : "";
      const pg = pages.length > 1 ? ` · page ${idx + 1}/${pages.length}` : "";
      ctx.addHeader(
        s,
        "Récupération des enfants en PEV de routine",
        `Antigènes ${g.start + 1} à ${g.end}${grp}${pg} — toutes tranches d'âge`
      );
      s.addTable([headTop, headSub, ...pageRows], {
        x: 0.25, y: tableY, w: W - 0.5, colW,
        border: { type: "solid", color: "DEE5EE", pt: 0.5 },
        rowH, valign: "middle", fontFace: "Calibri", fontSize: 11,
        autoPage: false,
      });
      addCommentBar(
        pptx,
        s,
        `${fmtInt(totalIdentifies)} enfants identifiés pour tous les antigènes — ${fmtInt(totalEnfants)} récupérés au PEV de routine, soit ${fmtPct(tauxRecupGlobal, 1)} de récupération (co-administration polio + PEV).`
      );
    });
  });
}

/* ─── Slide : Surveillance des MPV par ZS ──────────────────────────────── */

const MAROON = "7B2D3A";
const MAROON_LIGHT = "F3E6E9";

function buildSurveillanceMPV(ctx: SlideCtx): void {
  const { pptx, data } = ctx;
  const t = data.survTotals;

  const cards: { label: string; value: number }[] = [
    { label: "PFA – Cas", value: t.pfa },
    { label: "Rougeole – Cas", value: t.rougeole },
    { label: "Fièvre Jaune – Cas", value: t.fj },
    { label: "TNN – Cas", value: t.tnn },
  ];
  const drawCards = (s: PptxGenJS.Slide) => {
    const cw = 2.9;
    const gap = 0.2;
    const startX = (W - (cards.length * cw + (cards.length - 1) * gap)) / 2;
    cards.forEach((c, i) => {
      const x = startX + i * (cw + gap);
      s.addShape(pptx.ShapeType.roundRect, { x, y: 1.3, w: cw, h: 1.55, fill: { color: MAROON }, line: { color: MAROON, width: 1 }, rectRadius: 0.08 });
      s.addText(c.label, { x, y: 1.37, w: cw, h: 0.5, align: "center", valign: "middle", color: "FFFFFF", bold: true, fontSize: 16 });
      s.addText(fmtInt(c.value), { x, y: 1.85, w: cw, h: 0.9, align: "center", valign: "middle", color: "FFFFFF", bold: true, fontSize: 40 });
    });
  };

  const head: PptxGenJS.TableCell[] = [
    { text: data.byUnitLabel, options: thHeader({ fill: { color: MAROON } }) },
    { text: "PFA", options: thHeader({ fill: { color: MAROON } }) },
    { text: "Rougeole", options: thHeader({ fill: { color: MAROON } }) },
    { text: "Fièvre Jaune", options: thHeader({ fill: { color: MAROON } }) },
    { text: "TNN", options: thHeader({ fill: { color: MAROON } }) },
  ];
  const rows = data.survByUnit
    .filter((r) => r.pfa || r.rougeole || r.fj || r.tnn)
    .sort((a, b) => (b.pfa + b.rougeole + b.fj + b.tnn) - (a.pfa + a.rougeole + a.fj + a.tnn));
  const bodyRows: PptxGenJS.TableRow[] = rows.map((r): PptxGenJS.TableCell[] => [
    { text: r.unit, options: tdCell({ bold: true }) },
    { text: fmtInt(r.pfa), options: tdCell({ align: "right" }) },
    { text: fmtInt(r.rougeole), options: tdCell({ align: "right" }) },
    { text: fmtInt(r.fj), options: tdCell({ align: "right" }) },
    { text: fmtInt(r.tnn), options: tdCell({ align: "right" }) },
  ]);
  const totalRow: PptxGenJS.TableRow = [
    { text: "Total", options: thTotal({ fill: { color: MAROON } }) },
    { text: fmtInt(t.pfa), options: thTotal({ align: "right", fill: { color: MAROON } }) },
    { text: fmtInt(t.rougeole), options: thTotal({ align: "right", fill: { color: MAROON } }) },
    { text: fmtInt(t.fj), options: thTotal({ align: "right", fill: { color: MAROON } }) },
    { text: fmtInt(t.tnn), options: thTotal({ align: "right", fill: { color: MAROON } }) },
  ];

  const colW = computeColW(W - 6.0, 5, [1.8, 1, 1.2, 1.2, 1]);
  const allRows = [...bodyRows, totalRow];
  // 1ère diapo : table sous les cartes (y 3.25). Diapos suivantes : table en haut.
  const perPageFirst = rowsPerPage(3.25, 6.9, 0.5, 0.36);
  const perPageNext = rowsPerPage(1.3, 6.9, 0.5, 0.36);
  const pages: PptxGenJS.TableRow[][] = [];
  pages.push(allRows.slice(0, perPageFirst));
  for (let i = perPageFirst; i < allRows.length; i += perPageNext) pages.push(allRows.slice(i, i + perPageNext));

  pages.forEach((pageRows, idx) => {
    const s = pptx.addSlide();
    ctx.addHeader(s, `Surveillance des MPV par ${data.byUnitLabel}${suiteSuffix(idx, pages.length)}`, "Maladies à potentiel épidémique — recherche active des cas de MEV");
    const tableY = idx === 0 ? 3.25 : 1.3;
    if (idx === 0) drawCards(s);
    s.addTable([head, ...pageRows], {
      x: 3.0, y: tableY, w: W - 6.0, colW,
      border: { type: "solid", color: "E7D6DB", pt: 0.5 },
      rowH: 0.36, valign: "middle", fontFace: "Calibri", fontSize: 12, autoPage: false,
    });
    if (idx === 0 && rows.length === 0) {
      s.addText("Aucun cas de MPV notifié sur ce périmètre pendant la campagne.", {
        x: 1, y: 6.1, w: W - 2, h: 0.5, align: "center", fontSize: 15, italic: true, color: GREY,
      });
    }
  });
}

/* ─── Slides 8/9 : Gestion vaccin ──────────────────────────────────────── */

function buildGestion(
  ctx: SlideCtx,
  vaccine: "nVPO2" | "VPOb",
  rows: GestionRow[],
  globalTaux: number | null,
  seuil: number
): void {
  const { pptx, data } = ctx;

  const head: PptxGenJS.TableCell[] = [
    { text: data.byUnitLabel, options: thHeader() },
    { text: "Flacons\nreçus", options: thHeader() },
    { text: "Flacons\nutilisés", options: thHeader() },
    { text: "Flacons\nrendus", options: thHeader() },
    { text: "Flacons\nperdus", options: thHeader() },
    { text: "Enfants\nvaccinés", options: thHeader() },
    { text: `% ${vaccine}\nPerte`, options: thHeader() },
  ];
  const sUtil = rows.reduce((a, r) => a + r.flaconsUtil, 0);
  const sRec = rows.reduce((a, r) => a + r.flaconsRecus, 0);
  const sRend = rows.reduce((a, r) => a + r.flaconsRendus, 0);
  const sPerd = rows.reduce((a, r) => a + r.perdus, 0);
  const sVacc = rows.reduce((a, r) => a + r.vacc, 0);

  const bodyRows: PptxGenJS.TableRow[] = rows.map((r): PptxGenJS.TableCell[] => [
    { text: r.unit, options: tdCell({ bold: true }) },
    { text: fmtInt(r.flaconsRecus), options: tdCell({ align: "right" }) },
    { text: fmtInt(r.flaconsUtil), options: tdCell({ align: "right" }) },
    { text: fmtInt(r.flaconsRendus), options: tdCell({ align: "right" }) },
    { text: fmtInt(r.perdus), options: tdCell({ align: "right" }) },
    { text: fmtInt(r.vacc), options: tdCell({ align: "right" }) },
    { text: fmtPct(r.taux, 2), options: tdCell({ align: "right", bold: true, color: lossColor(r.taux) }) },
  ]);
  const totalRow: PptxGenJS.TableRow = [
    { text: "Total", options: thTotal() },
    { text: fmtInt(sRec), options: thTotal({ align: "right" }) },
    { text: fmtInt(sUtil), options: thTotal({ align: "right" }) },
    { text: fmtInt(sRend), options: thTotal({ align: "right" }) },
    { text: fmtInt(sPerd), options: thTotal({ align: "right" }) },
    { text: fmtInt(sVacc), options: thTotal({ align: "right" }) },
    { text: fmtPct(globalTaux, 2), options: thTotal({ align: "right" }) },
  ];

  const tableY = 1.3;
  const rowH = 0.38;
  const perPage = rowsPerPage(tableY, 5.85, 0.6, rowH);
  const pages = chunkRows([...bodyRows, totalRow], perPage);
  // Pour garder un graphique lisible quand beaucoup d'entités sont présentées
  // (toutes les ZS d'une province, toutes les AS d'une ZS…), on ne trace plus un
  // unique graphique surchargé : chaque diapo porte le graphique des seules
  // entités de SON tableau. Le graphe se répartit ainsi sur autant de diapos que
  // nécessaire, en parallèle de la pagination du tableau.
  pages.forEach((pageRows, idx) => {
    // Entités réelles présentes sur cette diapo (hors ligne « Total » finale).
    const pageRaw = rows.slice(idx * perPage, idx * perPage + perPage);
    const labels = pageRaw.map((r) => r.unit);
    const values = pageRaw.map((r) => (r.taux == null ? 0 : Math.round(r.taux * 100) / 100));
    const pageHasData = labels.length > 0 && values.some((v) => v !== 0);

    const s = pptx.addSlide();
    ctx.addHeader(s, `Gestion du vaccin : ${vaccine}${suiteSuffix(idx, pages.length)}`, `Seuil acceptable de perte : ≤ ${seuil} %`);
    s.addTable([head, ...pageRows], {
      x: 0.45, y: tableY, w: 7.1, colW: [1.6, 0.95, 0.95, 0.95, 0.95, 1.05, 0.65],
      border: { type: "solid", color: "DEE5EE", pt: 0.5 },
      rowH, valign: "middle", fontFace: "Calibri",
    });

    // Graphique de répartition du taux de perte — limité aux entités de la diapo.
    s.addShape(pptx.ShapeType.roundRect, { x: 7.7, y: 1.3, w: W - 8.15, h: 4.6, fill: { color: "FFFFFF" }, line: { color: "DEE5EE", width: 1 }, rectRadius: 0.06 });
    const chartTitle = pages.length > 1
      ? `Taux de perte (%) de ${vaccine} — ${data.byUnitLabel} (${idx + 1}/${pages.length})`
      : `Répartition du taux de perte (%) de ${vaccine} par ${data.byUnitLabel}`;
    s.addText(chartTitle, {
      x: 7.8, y: 1.37, w: W - 8.35, h: 0.45, fontSize: 15, bold: true, color: NAVY, align: "center",
    });
    if (pageHasData) {
      s.addChart(
        pptx.ChartType.bar,
        [{ name: `Taux de perte ${vaccine}`, labels, values }],
        {
          x: 7.7, y: 1.85, w: W - 8.15, h: 3.95,
          barDir: "bar", chartColors: [NAVY], showValue: true,
          dataLabelColor: NAVY_DEEP, dataLabelFontSize: 11, dataLabelFormatCode: '0.00"%"',
          catAxisLabelFontSize: 11, valAxisLabelFontSize: 11, showLegend: false,
          valGridLine: { style: "dash", color: "E2E8F0", size: 1 },
        }
      );
    } else {
      s.addText("Aucune donnée disponible pour ce périmètre.", {
        x: 7.7, y: 3.3, w: W - 8.15, h: 1, align: "center", fontSize: 14, italic: true, color: GREY,
      });
    }

    addCommentBar(pptx, s, gestionComment(globalTaux, vaccine, seuil));
  });
}

/* ─── Slide 10 : Surveillance MAPI ─────────────────────────────────────── */

function buildMapi(ctx: SlideCtx): void {
  const { pptx, data } = ctx;
  const s = pptx.addSlide();
  ctx.addHeader(s, "Synthèse Surveillance MAPI et récupération en routine", "Manifestations adverses post-immunisation & récupération PEV de routine");
  const sa = data.saillants;

  const totalIdentifies = data.recupAntigenIdentTotals.reduce((a, b) => a + b, 0);
  const totalRecuperes = data.recupAntigenTotals.reduce((a, b) => a + b, 0);
  const tauxRecup = totalIdentifies > 0 ? (totalRecuperes / totalIdentifies) * 100 : null;

  // Carte générique paramétrable (position, taille, corps de police).
  const card = (
    x: number, y: number, w: number, h: number,
    label: string, value: string, color: string, valueFont = 54, labelFont = 15
  ) => {
    s.addShape(pptx.ShapeType.roundRect, { x, y, w, h, fill: { color: "FFFFFF" }, line: { color, width: 2 }, rectRadius: 0.1 });
    s.addShape(pptx.ShapeType.rect, { x, y, w, h: 0.62, fill: { color } });
    s.addText(label, { x: x + 0.1, y, w: w - 0.2, h: 0.62, valign: "middle", align: "center", color: "FFFFFF", bold: true, fontSize: labelFont });
    s.addText(value, { x, y: y + 0.68, w, h: h - 0.78, align: "center", valign: "middle", fontSize: valueFont, bold: true, color });
  };

  // Rangée 1 — surveillance MAPI + récupérations PEV (3 cartes).
  const r1y = 1.45, r1h = 2.2, r1w = 3.7, r1gap = 0.3;
  const r1x0 = (W - (3 * r1w + 2 * r1gap)) / 2;
  card(r1x0, r1y, r1w, r1h, "MAPI mineures", fmtInt(sa.mapiMineures), ACCENT, 52, 16);
  card(r1x0 + (r1w + r1gap), r1y, r1w, r1h, "MAPI graves", fmtInt(sa.mapiGraves), sa.mapiGraves ? THR_LOW : THR_HIGH, 52, 16);
  card(r1x0 + 2 * (r1w + r1gap), r1y, r1w, r1h, "Récupérations PEV", fmtInt(sa.recup), THR_HIGH, 52, 16);

  // Rangée 2 — synthèse récupération en routine, tous antigènes confondus (2 cartes).
  const r2y = 3.95, r2h = 2.05, r2w = 5.7, r2gap = 0.4;
  const r2x0 = (W - (2 * r2w + r2gap)) / 2;
  card(r2x0, r2y, r2w, r2h, "Nombre total enfants identifiés (tous antigènes confondus)", fmtInt(totalIdentifies), NAVY, 44, 14);
  card(r2x0 + (r2w + r2gap), r2y, r2w, r2h, "% Enfants récupérés (tous antigènes confondus)", fmtPct(tauxRecup, 1), THR_HIGH, 44, 14);

  s.addText(
    "Toute MAPI grave doit faire l'objet d'une investigation immédiate, d'une notification au niveau supérieur et d'une prise en charge médicale, conformément au guide de surveillance.",
    { x: 0.9, y: 6.15, w: 11.5, h: 0.7, fontSize: 13, italic: true, color: GREY, align: "center" }
  );
}

/* ─── Slide 11 : Problèmes / Actions ───────────────────────────────────── */

function buildProblemes(ctx: SlideCtx): void {
  const { pptx, data } = ctx;
  // Pluralise le premier mot du libellé d'unité (Aire de Santé → Aires de Santé).
  const concLabel = `${data.byUnitLabel.replace(/^(\S+)/, "$1s")} concernées`;
  const head: PptxGenJS.TableCell[] = ["Problèmes identifiés", "Causes", concLabel, "Solutions proposées"]
    .map((h) => ({ text: h, options: thHeader({ fontSize: 15 }) }));

  const TOP = 1.35;
  const colW = [3.4, 3.0, 2.0, 3.5];
  const cellOpts = { margin: [4, 5, 4, 5] as [number, number, number, number] };

  if (data.problemes.length === 0) {
    const s = pptx.addSlide();
    ctx.addHeader(s, "Problèmes rencontrés / Actions correctrices");
    s.addTable([head, [
      { text: "Aucun problème majeur détecté par l'analyse sur ce périmètre — indicateurs conformes aux seuils.", options: tdCell({ fontSize: 14, italic: true, color: GREY, colspan: 4, align: "center", ...cellOpts }) },
    ]], {
      x: 0.45, y: TOP, w: W - 0.9, colW,
      border: { type: "solid", color: "DEE5EE", pt: 0.5 },
      rowH: 0.5, valign: "middle", fontFace: "Calibri", autoPage: false,
    });
    return;
  }

  const bodyRows: PptxGenJS.TableRow[] = data.problemes.map((p): PptxGenJS.TableCell[] => [
    { text: p.probleme, options: tdCell({ bold: true, fontSize: 13, ...cellOpts }) },
    { text: p.causes, options: tdCell({ fontSize: 13, ...cellOpts }) },
    { text: p.zs, options: tdCell({ align: "center", fontSize: 13, ...cellOpts }) },
    { text: p.solutions, options: tdCell({ fontSize: 13, ...cellOpts }) },
  ]);

  // 3 problèmes par diapo + lignes compactes (rowH = minimum, la cellule
  // s'agrandit au besoin) : le tableau reste lisible sans déborder.
  const perPage = 3;
  const pages = chunkRows(bodyRows, perPage);
  pages.forEach((pageRows, idx) => {
    const s = pptx.addSlide();
    ctx.addHeader(s, `Problèmes rencontrés / Actions correctrices${suiteSuffix(idx, pages.length)}`);
    s.addTable([head, ...pageRows], {
      x: 0.45, y: TOP, w: W - 0.9, colW,
      border: { type: "solid", color: "DEE5EE", pt: 0.5 },
      rowH: 0.4, valign: "middle", fontFace: "Calibri", autoPage: false,
    });
  });
}

/* ─── Slide 12 : Merci ─────────────────────────────────────────────────── */

function buildMerci(ctx: SlideCtx): void {
  const { pptx, pev } = ctx;
  const s = pptx.addSlide();
  s.background = { color: NAVY };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: H / 2 - 1.0, w: W, h: 2.0, fill: { color: NAVY_DEEP } });
  s.addShape(pptx.ShapeType.rect, { x: 0, y: H / 2 - 1.0, w: W, h: 0.08, fill: { color: ACCENT } });
  s.addShape(pptx.ShapeType.rect, { x: 0, y: H / 2 + 0.92, w: W, h: 0.08, fill: { color: ACCENT } });
  s.addText("MERCI POUR VOTRE ATTENTION", {
    x: 0, y: H / 2 - 1.0, w: W, h: 2.0,
    align: "center", valign: "middle", fontSize: 44, bold: true, color: "FFFFFF", charSpacing: 4,
  });
  if (pev) {
    const cy = H / 2 + 1.4;
    s.addShape(pptx.ShapeType.roundRect, { x: W / 2 - 1.9, y: cy, w: 3.8, h: 0.85, fill: { color: "FFFFFF" }, line: { color: "FFFFFF", width: 0 }, rectRadius: 0.08 });
    s.addImage({ data: pev, x: W / 2 - 1.7, y: cy + 0.165, w: 3.4, h: 0.521 });
  }
}

/* ─── Sous-helpers tableau / cellules / commentaires ───────────────────── */

function thHeader(extra: Partial<PptxGenJS.TableCellProps> = {}): PptxGenJS.TableCellProps {
  return {
    bold: true, color: "FFFFFF", fill: { color: NAVY }, fontSize: 11,
    align: "center", valign: "middle", fontFace: "Calibri", ...extra,
  };
}
function thTotal(extra: Partial<PptxGenJS.TableCellProps> = {}): PptxGenJS.TableCellProps {
  return {
    bold: true, color: "FFFFFF", fill: { color: NAVY_DEEP }, fontSize: 11,
    align: "center", valign: "middle", fontFace: "Calibri", ...extra,
  };
}
function tdCell(extra: Partial<PptxGenJS.TableCellProps> = {}): PptxGenJS.TableCellProps {
  return {
    color: NAVY_DEEP, fontSize: 11, valign: "middle", fontFace: "Calibri",
    fill: { color: "FFFFFF" }, ...extra,
  };
}

/** Découpe un tableau (corps + total) en pages de `perPage` lignes max. */
function chunkRows<T>(rows: T[], perPage: number): T[][] {
  if (perPage <= 0 || rows.length <= perPage) return [rows];
  const out: T[][] = [];
  for (let i = 0; i < rows.length; i += perPage) out.push(rows.slice(i, i + perPage));
  return out;
}

/** Nombre de lignes de corps tenant dans la hauteur disponible. */
function rowsPerPage(tableTop: number, bottom: number, headerH: number, rowH: number): number {
  return Math.max(3, Math.floor((bottom - tableTop - headerH) / rowH));
}

/** Suffixe « (suite) » pour les diapos additionnelles d'un même tableau. */
function suiteSuffix(idx: number, total: number): string {
  if (total <= 1) return "";
  return idx === 0 ? "" : ` (suite ${idx + 1}/${total})`;
}

function computeColW(totalW: number, n: number, weights: number[]): number[] {
  const sum = weights.reduce((a, b) => a + b, 0);
  return weights.slice(0, n).map((w) => (w / sum) * totalW);
}

function addCommentBar(pptx: PptxGenJS, s: PptxGenJS.Slide, comment: string): void {
  s.addShape(pptx.ShapeType.roundRect, {
    x: 0.5, y: 6.05, w: W - 1, h: 0.85,
    fill: { color: ACCENT_LIGHT }, line: { color: ACCENT, width: 1 }, rectRadius: 0.05,
  });
  s.addText([
    { text: "Commentaire : ", options: { bold: true, color: ACCENT } },
    { text: comment, options: { color: NAVY_DEEP } },
  ], { x: 0.7, y: 6.05, w: W - 1.4, h: 0.85, valign: "middle", fontSize: 14 });
}

function coverageComment(
  rows: CoverageDailyRow[],
  globalCV: number | null,
  vaccine: string,
  byUnitLabel: string
): string {
  if (rows.length === 0) return `Aucune donnée de couverture ${vaccine} n'est disponible sur ce périmètre.`;
  const niveau = levelWord(byUnitLabel);
  const below = rows
    .filter((r) => r.couvGlobale != null && r.couvGlobale < 95)
    .sort((a, b) => (a.couvGlobale ?? 0) - (b.couvGlobale ?? 0));

  if (below.length === 0) {
    return `La couverture ${vaccine} atteint ${fmtPct(globalCV)} : l'objectif de 95 % est tenu dans chaque ${niveau} du périmètre. Il reste à maintenir ce niveau et à documenter les bonnes pratiques pour les prochains passages.`;
  }

  const shown = below.slice(0, 3).map((r) => `${r.unit} (${fmtPct(r.couvGlobale, 0)})`);
  const liste = below.length > 3
    ? `${joinAnd(shown)}, parmi ${countLevel(below.length, byUnitLabel)} concernées`
    : joinAnd(shown);
  const verbe = below.length > 1 ? "restent en deçà" : "reste en deçà";
  return `Avec ${fmtPct(globalCV)}, la couverture ${vaccine} reste à consolider sur le périmètre. Les ${levelWord(byUnitLabel, true)} les plus en retard ${verbe} de l'objectif de 95 % : ${liste}. Y prévoir des passages de rattrapage (porte-à-porte et sites fixes) en priorisant les plus faibles.`;
}

function gestionComment(globalTaux: number | null, vaccine: string, seuil: number): string {
  if (globalTaux == null) return `Les données de gestion du vaccin ${vaccine} ne sont pas renseignées sur ce périmètre.`;
  const taux = fmtPct(globalTaux);
  if (globalTaux < 0) {
    return `Le taux de perte ${vaccine} ressort à ${taux}, valeur négative qui signale une saisie incohérente des flacons (vaccinés supérieurs aux doses disponibles). À faire recroiser et corriger par les gestionnaires de données avant exploitation.`;
  }
  if (globalTaux > seuil) {
    return `Le taux de perte ${vaccine} s'établit à ${taux}, au-delà du seuil acceptable de ${seuil} %. À investiguer : maîtrise de la chaîne du froid, gestion des flacons entamés et fiabilité de la saisie des mouvements de flacons.`;
  }
  return `Le taux de perte ${vaccine} s'établit à ${taux}, sous le seuil de ${seuil} % : la gestion du vaccin est maîtrisée sur le périmètre. À maintenir lors des prochains passages.`;
}

// Le helper fmtNum reste exporté pour les usages avancés.
export { fmtNum };
