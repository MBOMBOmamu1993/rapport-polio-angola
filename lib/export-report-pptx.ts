/**
 * Génération du rapport PowerPoint « Campagne de vaccination polio synchronisée
 * avec l'Angola » — reproduit fidèlement le modèle officiel Kwango en ne gardant
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
  daily: { recus: number; couv: number | null }[];
}

export interface ProblemeRow {
  probleme: string;
  causes: string;
  zs: string;
  solutions: string;
}

export interface ReportData {
  province: string;
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
  recupAntigenByUnit: { unit: string; ev: number[] }[];
  recupAntigenTotals: number[];
  /** Surveillance des MPV (cas notifiés). */
  survByUnit: { unit: string; pfa: number; rougeole: number; fj: number; tnn: number }[];
  survTotals: { pfa: number; rougeole: number; fj: number; tnn: number };
  problemes: ProblemeRow[];
  completudeMapPng?: string | null;
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

/** Rasterise un SVG (fond transparent) en PNG data URL pour l'insertion dans le PPTX. */
async function svgToPngData(path: string, w: number, h: number): Promise<string | null> {
  try {
    const res = await fetch(path);
    if (!res.ok) return null;
    const svg = await res.text();
    return await new Promise((resolve) => {
      const img = new Image();
      img.onload = () => {
        const canvas = document.createElement("canvas");
        canvas.width = w;
        canvas.height = h;
        const ctx = canvas.getContext("2d");
        if (!ctx) return resolve(null);
        ctx.drawImage(img, 0, 0, w, h);
        resolve(canvas.toDataURL("image/png"));
      };
      img.onerror = () => resolve(null);
      img.src = "data:image/svg+xml;base64," + btoa(unescape(encodeURIComponent(svg)));
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
    svgToPngData("/logo/pev-logo-white.svg", 930, 198),
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

  await pptx.writeFile({ fileName: `Rapport_Polio_Angola_${slug(data.scopeLabel)}.pptx` });
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
    s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: 0.95, fill: { color: NAVY } });
    s.addShape(pptx.ShapeType.rect, { x: 0, y: 0.95, w: W, h: 0.08, fill: { color: ACCENT } });
    s.addText(title, {
      x: 0.55, y: 0.08, w: W - 2.4, h: subtitle ? 0.55 : 0.83,
      fontSize: 20, bold: true, color: "FFFFFF", align: "left",
      valign: subtitle ? "top" : "middle", fontFace: "Calibri",
    });
    if (subtitle) {
      s.addText(subtitle, {
        x: 0.55, y: 0.6, w: W - 2.4, h: 0.32,
        fontSize: 11, color: "CCE4FF", align: "left", italic: true, fontFace: "Calibri",
      });
    }
    if (pev) s.addImage({ data: pev, x: W - 2.25, y: 0.26, w: 2.0, h: 0.43 });

    // Pied de page.
    s.addText("Campagne de vaccination polio synchronisée avec l'Angola — nVPO2 & VPOb (co-administration)", {
      x: 0.55, y: H - 0.36, w: W - 4.2, h: 0.28, fontSize: 8, color: GREY, align: "left",
    });
    s.addText(data.scopeLabel, {
      x: W - 4.2, y: H - 0.36, w: 3.8, h: 0.28, fontSize: 8, color: NAVY, align: "right", bold: true,
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

  // Logo PEV (sans fond) — lockup horizontal.
  if (pev) s.addImage({ data: pev, x: 0.6, y: 0.55, w: 3.5, h: 0.745 });

  // Eyebrow.
  s.addText("RAPPORT DES RÉSULTATS", {
    x: 0.6, y: 2.05, w: 5.4, h: 0.4,
    fontSize: 14, color: ACCENT_LIGHT, bold: true, charSpacing: 4, fontFace: "Calibri",
  });
  // Titre principal.
  s.addText("Campagne de vaccination polio synchronisée avec l'Angola", {
    x: 0.6, y: 2.5, w: 5.4, h: 1.7,
    fontSize: 28, color: "FFFFFF", bold: true, fontFace: "Calibri", lineSpacingMultiple: 1.05,
  });
  s.addShape(pptx.ShapeType.rect, { x: 0.62, y: 4.2, w: 1.6, h: 0.06, fill: { color: ACCENT } });

  s.addText(
    [
      { text: `Province : ${data.province}`, options: { bold: true, fontSize: 16, color: "FFFFFF", breakLine: true } },
      { text: data.periode || "Période de la campagne", options: { fontSize: 13, color: "CCE4FF", breakLine: true } },
      { text: "Vaccins : nVPO2 et VPOb (co-administration)", options: { fontSize: 12, color: "93C5FD" } },
    ],
    { x: 0.6, y: 4.45, w: 5.4, h: 1.5, valign: "top" }
  );

  s.addText("Programme Élargi de Vaccination — RD Congo", {
    x: 0.6, y: H - 0.7, w: 5.4, h: 0.4,
    fontSize: 11, color: "93C5FD", italic: true,
  });
  s.addText(data.dateLabel, {
    x: W * 0.42 + 0.2, y: H - 0.7, w: W * 0.58 - 0.4, h: 0.4,
    fontSize: 11, color: "FFFFFF", align: "right", bold: true,
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
    "Surveillance des MAPI",
    "Problèmes rencontrés / Actions correctrices",
  ];
  items.forEach((it, i) => {
    const y = 1.45 + i * 0.6;
    s.addShape(pptx.ShapeType.ellipse, { x: 1.2, y, w: 0.45, h: 0.45, fill: { color: ACCENT } });
    s.addText(String(i + 1), { x: 1.2, y, w: 0.45, h: 0.45, align: "center", valign: "middle", color: "FFFFFF", bold: true, fontSize: 14 });
    s.addText(it, { x: 1.85, y, w: 9.8, h: 0.45, valign: "middle", fontSize: 15, color: NAVY });
  });
}

/* ─── Slide 3 : Points saillants ──────────────────────────────────────── */

function buildPointsSaillants(ctx: SlideCtx): void {
  const { pptx, data } = ctx;
  const s = pptx.addSlide();
  ctx.addHeader(s, "Points saillants", data.periode || undefined);
  const sa = data.saillants;
  const card = (x: number, w: number, h: number, header: string, rows: { label: string; value: string; tone?: string }[]) => {
    s.addShape(pptx.ShapeType.roundRect, { x, y: 1.3, w, h, fill: { color: SOFT }, line: { color: "DEE5EE", width: 1 }, rectRadius: 0.1 });
    s.addShape(pptx.ShapeType.rect, { x, y: 1.3, w, h: 0.5, fill: { color: NAVY } });
    s.addText(header, { x: x + 0.2, y: 1.3, w: w - 0.4, h: 0.5, valign: "middle", color: "FFFFFF", bold: true, fontSize: 13 });
    rows.forEach((ln, i) => {
      const y = 2.05 + i * 0.55;
      s.addText(ln.label, { x: x + 0.25, y, w: w - 2.0, h: 0.5, valign: "middle", fontSize: 12, color: NAVY });
      s.addText(ln.value, { x: x + w - 2.0, y, w: 1.75, h: 0.5, valign: "middle", align: "right", fontSize: 14, bold: true, color: ln.tone ?? NAVY_DEEP });
    });
  };
  card(0.45, 6.1, 5.6, "COMPLÉTUDE & VACCINATION", [
    { label: "Rapports attendus", value: fmtInt(sa.completudeAttendus) },
    { label: "Rapports reçus", value: fmtInt(sa.completudeRecus) },
    { label: "Complétude des rapports", value: fmtPct(sa.completude), tone: thresholdColor(sa.completude) },
    { label: "Vaccinés nVPO2 / Cible", value: `${fmtInt(sa.nvpo2Vacc)} / ${fmtInt(sa.nvpo2Cible)}` },
    { label: "Couverture nVPO2", value: fmtPct(sa.nvpo2CV), tone: thresholdColor(sa.nvpo2CV) },
    { label: "Vaccinés VPOb / Cible", value: `${fmtInt(sa.vpobVacc)} / ${fmtInt(sa.vpobCible)}` },
    { label: "Couverture VPOb", value: fmtPct(sa.vpobCV), tone: thresholdColor(sa.vpobCV) },
    { label: "Récupérations PEV", value: fmtInt(sa.recup) },
  ]);
  card(6.78, 6.1, 5.6, "GESTION VACCINS & MAPI", [
    { label: "Flacons nVPO2 reçus", value: fmtInt(sa.nvpo2FlaconsRecus) },
    { label: "Flacons nVPO2 utilisés", value: fmtInt(sa.nvpo2FlaconsUtil) },
    { label: "Taux de perte nVPO2", value: fmtPct(sa.nvpo2Perte), tone: lossColor(sa.nvpo2Perte) },
    { label: "Flacons VPOb reçus", value: fmtInt(sa.vpobFlaconsRecus) },
    { label: "Flacons VPOb utilisés", value: fmtInt(sa.vpobFlaconsUtil) },
    { label: "Taux de perte VPOb", value: fmtPct(sa.vpobPerte), tone: lossColor(sa.vpobPerte) },
    { label: "MAPI mineures notifiées", value: fmtInt(sa.mapiMineures) },
    { label: "MAPI graves notifiées", value: fmtInt(sa.mapiGraves), tone: sa.mapiGraves ? THR_LOW : GREY },
  ]);
}

/* ─── Slide 4 : Complétude des rapports ────────────────────────────────── */

function buildCompletude(ctx: SlideCtx): void {
  const { pptx, data } = ctx;
  const s = pptx.addSlide();
  ctx.addHeader(s, "Complétude des rapports journaliers et globale", "Source : Synthèse du masque de saisie");

  const sa = data.saillants;
  const globalCV = sa.completude;

  // ── Légende des seuils en haut.
  addLegend(pptx, s, 0.5, 1.2);

  // ── Bloc gauche : KPI principal + carte d'identité de la complétude.
  const arcColor = thresholdColor(globalCV);
  s.addShape(pptx.ShapeType.roundRect, { x: 0.5, y: 1.95, w: 6.2, h: 4.0, fill: { color: SOFT }, line: { color: "DEE5EE", width: 1 }, rectRadius: 0.08 });
  s.addText("Complétude globale", { x: 0.7, y: 2.05, w: 5.8, h: 0.4, fontSize: 14, bold: true, color: NAVY });

  // KPI géant — cercle plein avec valeur centrée + libellé.
  s.addShape(pptx.ShapeType.ellipse, { x: 0.95, y: 2.6, w: 2.4, h: 2.4, fill: { color: arcColor }, line: { color: arcColor, width: 0 } });
  s.addShape(pptx.ShapeType.ellipse, { x: 1.15, y: 2.8, w: 2.0, h: 2.0, fill: { color: "FFFFFF" }, line: { color: "FFFFFF", width: 0 } });
  s.addText(fmtPct(globalCV), {
    x: 0.95, y: 3.45, w: 2.4, h: 0.7,
    align: "center", valign: "middle", fontSize: 24, bold: true, color: arcColor,
  });
  s.addText("complétude\nrapports", {
    x: 0.95, y: 4.0, w: 2.4, h: 0.5,
    align: "center", valign: "top", fontSize: 10, color: GREY,
  });

  // KPIs à droite (rapports reçus / attendus).
  s.addShape(pptx.ShapeType.roundRect, { x: 3.55, y: 2.85, w: 3.05, h: 1.0, fill: { color: "FFFFFF" }, line: { color: "DEE5EE", width: 1 }, rectRadius: 0.06 });
  s.addText("Rapports reçus", { x: 3.7, y: 2.92, w: 2.8, h: 0.3, fontSize: 10, color: GREY });
  s.addText(fmtInt(sa.completudeRecus), { x: 3.7, y: 3.2, w: 2.8, h: 0.6, fontSize: 22, bold: true, color: NAVY_DEEP });

  s.addShape(pptx.ShapeType.roundRect, { x: 3.55, y: 3.95, w: 3.05, h: 1.0, fill: { color: "FFFFFF" }, line: { color: "DEE5EE", width: 1 }, rectRadius: 0.06 });
  s.addText("Rapports attendus", { x: 3.7, y: 4.02, w: 2.8, h: 0.3, fontSize: 10, color: GREY });
  s.addText(fmtInt(sa.completudeAttendus), { x: 3.7, y: 4.3, w: 2.8, h: 0.6, fontSize: 22, bold: true, color: NAVY_DEEP });

  s.addText("Une bonne complétude (≥ 95 %) conditionne la fiabilité des couvertures vaccinales calculées.", {
    x: 0.7, y: 5.15, w: 5.8, h: 0.7, fontSize: 10, color: GREY, italic: true,
  });

  // ── Bloc droit : tableau jour-par-jour (par unité d'agrégation).
  s.addShape(pptx.ShapeType.roundRect, { x: 6.9, y: 1.95, w: W - 7.4, h: 4.0, fill: { color: "FFFFFF" }, line: { color: "DEE5EE", width: 1 }, rectRadius: 0.08 });
  s.addText(`Complétude par ${data.byUnitLabel}`, { x: 7.05, y: 2.05, w: 5.5, h: 0.4, fontSize: 14, bold: true, color: NAVY });

  const days = data.jourLabels;
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
  const trows: PptxGenJS.TableRow[] = [
    head,
    ...data.completudeByUnit.map((r) => {
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
    }),
    [
      { text: "Total", options: thTotal() },
      { text: fmtInt(totalAttendus), options: thTotal({ align: "right" }) },
      ...days.flatMap((_, i): PptxGenJS.TableCell[] => {
        const couvJour = totalAttendus > 0 ? (totalsDaily[i] / totalAttendus) * 100 : null;
        return [
          { text: fmtInt(totalsDaily[i]), options: thTotal({ align: "right" }) },
          { text: fmtPct(couvJour, 2), options: thTotal({ align: "right" }) },
        ];
      }),
      { text: fmtPct(totalAttendus > 0 ? (totalRecus / totalAttendus) * 100 : null, 2), options: thTotal({ align: "right" }) },
    ],
  ];
  const colW = computeColW(W - 7.6, 2 + days.length * 2 + 1, [1.4, 0.9, ...days.flatMap(() => [0.8, 0.8]), 0.9]);
  s.addTable(trows, {
    x: 7.0, y: 2.5, w: W - 7.4, colW,
    border: { type: "solid", color: "DEE5EE", pt: 0.5 },
    rowH: 0.28, valign: "middle", fontFace: "Calibri",
  });

}

/* ─── Slide 4 bis : Spatialisation de la complétude (carte RDC / ZS) ────── */

function buildCompletudeMap(ctx: SlideCtx): void {
  const { pptx, data } = ctx;
  const s = pptx.addSlide();
  ctx.addHeader(s, "Spatialisation de la complétude globale", "Carte des Zones de Santé de la RDC — Source : Synthèse du masque de saisie");
  addLegend(pptx, s, 0.5, 1.15);

  const png = data.completudeMapPng;
  if (png) {
    // Fond cartographique RDC (1100×1000 ≈ ratio 1.1).
    const h = 5.45;
    const w = h * 1.1;
    s.addImage({ data: png, x: (W - w) / 2, y: 1.7, w, h });
  } else {
    s.addShape(pptx.ShapeType.roundRect, {
      x: 2.0, y: 2.8, w: W - 4.0, h: 1.8,
      fill: { color: SOFT }, line: { color: "DEE5EE", width: 1 }, rectRadius: 0.1,
    });
    s.addText(
      "Carte indisponible — le fond cartographique des Zones de Santé n'a pas pu être chargé au moment de la génération (une connexion Internet est requise).",
      { x: 2.3, y: 2.8, w: W - 4.6, h: 1.8, align: "center", valign: "middle", fontSize: 13, italic: true, color: GREY }
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
    x, y, w: 3.5, h: 0.35, fontSize: 11, color: NAVY, bold: true, valign: "middle",
  });
  items.forEach((it, i) => {
    const cx = x + 3.6 + i * 1.95;
    s.addShape(pptx.ShapeType.rect, { x: cx, y: y + 0.04, w: 0.4, h: 0.28, fill: { color: it.c } });
    s.addText(it.t, { x: cx + 0.45, y, w: 1.5, h: 0.35, fontSize: 10, color: NAVY, valign: "middle" });
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
  const s = pptx.addSlide();
  ctx.addHeader(s, `Couvertures vaccinales ${vaccine}, par ${data.byUnitLabel}`, "Source : Synthèse du masque de saisie");
  addLegend(pptx, s, 0.5, 1.15);

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

  const trows: PptxGenJS.TableRow[] = [
    head,
    ...rows.map((r): PptxGenJS.TableCell[] => {
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
    }),
    [
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
    ],
  ];

  const nCols = 3 + days.length * 2 + 1;
  const weights = [1.6, 0.85, 0.95, ...days.flatMap(() => [0.85, 0.8]), 0.95];
  const colW = computeColW(W - 1.1, nCols, weights);
  s.addTable(trows, {
    x: 0.55, y: 1.65, w: W - 1.1, colW,
    border: { type: "solid", color: "DEE5EE", pt: 0.5 },
    rowH: 0.30, valign: "middle", fontFace: "Calibri",
  });

  addCommentBar(pptx, s, coverageComment(rows, globalCV, vaccine));
}

/* ─── Slide 7 : Récupération PEV de routine ────────────────────────────── */

function buildRecup(ctx: SlideCtx): void {
  const { pptx, data } = ctx;
  const s = pptx.addSlide();
  ctx.addHeader(s, "Récupération des enfants en PEV de routine", "Enfants vaccinés (EV) par antigène pendant la campagne — toutes tranches d'âge");

  const ant = data.antigenLabels;
  const rows = data.recupAntigenByUnit;
  const totals = data.recupAntigenTotals;
  const hasData = totals.some((v) => v > 0);

  if (ant.length === 0 || !hasData) {
    s.addText("Aucune récupération PEV par antigène saisie pour ce périmètre.", {
      x: 1, y: 3.2, w: W - 2, h: 1, align: "center", fontSize: 14, italic: true, color: GREY,
    });
    addCommentBar(pptx, s, `${fmtInt(data.saillants.recup)} enfants récupérés et orientés vers le PEV de routine pendant la campagne.`);
    return;
  }

  // Tableau matriciel : unité (ZS) × antigène (EV).
  const head: PptxGenJS.TableCell[] = [
    { text: data.byUnitLabel, options: thHeader() },
    ...ant.map((a): PptxGenJS.TableCell => ({ text: a, options: thHeader() })),
  ];
  const trows: PptxGenJS.TableRow[] = [
    head,
    ...rows.map((r): PptxGenJS.TableCell[] => [
      { text: r.unit, options: tdCell({ bold: true }) },
      ...r.ev.map((v): PptxGenJS.TableCell => ({ text: fmtInt(v), options: tdCell({ align: "right" }) })),
    ]),
    [
      { text: "Total", options: thTotal() },
      ...totals.map((v): PptxGenJS.TableCell => ({ text: fmtInt(v), options: thTotal({ align: "right" }) })),
    ],
  ];

  const nCols = 1 + ant.length;
  const colW = computeColW(W - 0.9, nCols, [1.7, ...ant.map(() => 0.75)]);
  s.addTable(trows, {
    x: 0.45, y: 1.25, w: W - 0.9, colW,
    border: { type: "solid", color: "DEE5EE", pt: 0.5 },
    rowH: 0.3, valign: "middle", fontFace: "Calibri", fontSize: 8,
    autoPage: false,
  });

  const totalEnfants = totals.reduce((a, b) => a + b, 0);
  addCommentBar(pptx, s, `${fmtInt(data.saillants.recup)} enfants récupérés au PEV de routine — ${fmtInt(totalEnfants)} doses d'antigènes administrées (co-administration polio + PEV).`);
}

/* ─── Slide : Surveillance des MPV par ZS ──────────────────────────────── */

const MAROON = "7B2D3A";
const MAROON_LIGHT = "F3E6E9";

function buildSurveillanceMPV(ctx: SlideCtx): void {
  const { pptx, data } = ctx;
  const s = pptx.addSlide();
  ctx.addHeader(s, "Surveillance des MPV par Zone de Santé", "Maladies à potentiel épidémique — recherche active des cas de MEV");
  const t = data.survTotals;

  // Cartes de synthèse (PFA / Rougeole / FJ / TNN).
  const cards: { label: string; value: number }[] = [
    { label: "PFA – Cas", value: t.pfa },
    { label: "Rougeole – Cas", value: t.rougeole },
    { label: "Fièvre Jaune – Cas", value: t.fj },
    { label: "TNN – Cas", value: t.tnn },
  ];
  const cw = 2.9;
  const gap = 0.2;
  const startX = (W - (cards.length * cw + (cards.length - 1) * gap)) / 2;
  cards.forEach((c, i) => {
    const x = startX + i * (cw + gap);
    s.addShape(pptx.ShapeType.roundRect, { x, y: 1.25, w: cw, h: 1.5, fill: { color: MAROON }, line: { color: MAROON, width: 1 }, rectRadius: 0.08 });
    s.addText(c.label, { x, y: 1.32, w: cw, h: 0.45, align: "center", valign: "middle", color: "FFFFFF", bold: true, fontSize: 13 });
    s.addText(fmtInt(c.value), { x, y: 1.75, w: cw, h: 0.9, align: "center", valign: "middle", color: "FFFFFF", bold: true, fontSize: 34 });
  });

  // Tableau par unité (ZS) : PFA / Rougeole / FJ / TNN.
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
  const body: PptxGenJS.TableRow[] = rows.map((r): PptxGenJS.TableCell[] => [
    { text: r.unit, options: tdCell({ bold: true }) },
    { text: fmtInt(r.pfa), options: tdCell({ align: "right" }) },
    { text: fmtInt(r.rougeole), options: tdCell({ align: "right" }) },
    { text: fmtInt(r.fj), options: tdCell({ align: "right" }) },
    { text: fmtInt(r.tnn), options: tdCell({ align: "right" }) },
  ]);
  const trows: PptxGenJS.TableRow[] = [
    head,
    ...body,
    [
      { text: "Total", options: thTotal({ fill: { color: MAROON } }) },
      { text: fmtInt(t.pfa), options: thTotal({ align: "right", fill: { color: MAROON } }) },
      { text: fmtInt(t.rougeole), options: thTotal({ align: "right", fill: { color: MAROON } }) },
      { text: fmtInt(t.fj), options: thTotal({ align: "right", fill: { color: MAROON } }) },
      { text: fmtInt(t.tnn), options: thTotal({ align: "right", fill: { color: MAROON } }) },
    ],
  ];
  s.addShape(pptx.ShapeType.roundRect, { x: 2.4, y: 3.0, w: W - 4.8, h: 0.02, fill: { color: MAROON_LIGHT } });
  s.addTable(trows, {
    x: 3.0, y: 3.1, w: W - 6.0, colW: computeColW(W - 6.0, 5, [1.8, 1, 1.2, 1.2, 1]),
    border: { type: "solid", color: "E7D6DB", pt: 0.5 },
    rowH: 0.3, valign: "middle", fontFace: "Calibri", fontSize: 9, autoPage: false,
  });

  if (rows.length === 0) {
    s.addText("Aucun cas de MPV notifié sur ce périmètre pendant la campagne.", {
      x: 1, y: 6.1, w: W - 2, h: 0.5, align: "center", fontSize: 12, italic: true, color: GREY,
    });
  }
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
  const s = pptx.addSlide();
  ctx.addHeader(s, `Gestion du vaccin : ${vaccine}`, `Seuil acceptable de perte : ≤ ${seuil} %`);

  // ── Tableau gauche : Flacons reçus / utilisés / rendus / perdus / vaccinés / %perte.
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

  const trows: PptxGenJS.TableRow[] = [
    head,
    ...rows.map((r): PptxGenJS.TableCell[] => [
      { text: r.unit, options: tdCell({ bold: true }) },
      { text: fmtInt(r.flaconsRecus), options: tdCell({ align: "right" }) },
      { text: fmtInt(r.flaconsUtil), options: tdCell({ align: "right" }) },
      { text: fmtInt(r.flaconsRendus), options: tdCell({ align: "right" }) },
      { text: fmtInt(r.perdus), options: tdCell({ align: "right" }) },
      { text: fmtInt(r.vacc), options: tdCell({ align: "right" }) },
      { text: fmtPct(r.taux, 2), options: tdCell({ align: "right", bold: true, color: lossColor(r.taux) }) },
    ]),
    [
      { text: "Total", options: thTotal() },
      { text: fmtInt(sRec), options: thTotal({ align: "right" }) },
      { text: fmtInt(sUtil), options: thTotal({ align: "right" }) },
      { text: fmtInt(sRend), options: thTotal({ align: "right" }) },
      { text: fmtInt(sPerd), options: thTotal({ align: "right" }) },
      { text: fmtInt(sVacc), options: thTotal({ align: "right" }) },
      { text: fmtPct(globalTaux, 2), options: thTotal({ align: "right" }) },
    ],
  ];
  s.addTable(trows, {
    x: 0.45, y: 1.25, w: 7.1, colW: [1.6, 0.95, 0.95, 0.95, 0.95, 1.05, 0.65],
    border: { type: "solid", color: "DEE5EE", pt: 0.5 },
    rowH: 0.32, valign: "middle", fontFace: "Calibri",
  });

  // ── Bar chart horizontal à droite : Répartition du taux de perte.
  const labels = rows.map((r) => r.unit);
  const values = rows.map((r) => (r.taux == null ? 0 : Math.round(r.taux * 100) / 100));
  const hasData = rows.length > 0 && values.some((v) => v !== 0);
  s.addShape(pptx.ShapeType.roundRect, { x: 7.7, y: 1.25, w: W - 8.15, h: 4.6, fill: { color: "FFFFFF" }, line: { color: "DEE5EE", width: 1 }, rectRadius: 0.06 });
  s.addText(`Répartition du taux de perte (%) de ${vaccine} par ${data.byUnitLabel}`, {
    x: 7.8, y: 1.32, w: W - 8.35, h: 0.4, fontSize: 12, bold: true, color: NAVY, align: "center",
  });
  if (hasData) {
    s.addChart(
      pptx.ChartType.bar,
      [{ name: `Taux de perte ${vaccine}`, labels, values }],
      {
        x: 7.7, y: 1.75, w: W - 8.15, h: 4.05,
        barDir: "bar", chartColors: [NAVY], showValue: true,
        dataLabelColor: NAVY_DEEP, dataLabelFontSize: 9, dataLabelFormatCode: '0.00"%"',
        catAxisLabelFontSize: 9, valAxisLabelFontSize: 9, showLegend: false,
        valGridLine: { style: "dash", color: "E2E8F0", size: 1 },
      }
    );
  } else {
    s.addText("Aucune donnée disponible pour ce périmètre.", {
      x: 7.7, y: 3.3, w: W - 8.15, h: 1, align: "center", fontSize: 12, italic: true, color: GREY,
    });
  }

  addCommentBar(pptx, s, gestionComment(globalTaux, vaccine, seuil));
}

/* ─── Slide 10 : Surveillance MAPI ─────────────────────────────────────── */

function buildMapi(ctx: SlideCtx): void {
  const { pptx, data } = ctx;
  const s = pptx.addSlide();
  ctx.addHeader(s, "Surveillance des MAPI", "Manifestations adverses post-immunisation");
  const sa = data.saillants;
  const card = (x: number, label: string, value: string, color: string) => {
    s.addShape(pptx.ShapeType.roundRect, { x, y: 2.2, w: 3.6, h: 2.5, fill: { color: "FFFFFF" }, line: { color, width: 2 }, rectRadius: 0.1 });
    s.addShape(pptx.ShapeType.rect, { x, y: 2.2, w: 3.6, h: 0.5, fill: { color } });
    s.addText(label, { x: x + 0.15, y: 2.2, w: 3.3, h: 0.5, valign: "middle", align: "center", color: "FFFFFF", bold: true, fontSize: 13 });
    s.addText(value, { x, y: 2.85, w: 3.6, h: 1.6, align: "center", valign: "middle", fontSize: 56, bold: true, color });
  };
  card(0.9, "MAPI mineures", fmtInt(sa.mapiMineures), ACCENT);
  card(4.85, "MAPI graves", fmtInt(sa.mapiGraves), sa.mapiGraves ? THR_LOW : THR_HIGH);
  card(8.8, "Récupérations PEV", fmtInt(sa.recup), THR_HIGH);

  s.addText(
    "Toute MAPI grave doit faire l'objet d'une investigation immédiate, d'une notification au niveau supérieur et d'une prise en charge médicale, conformément au guide de surveillance.",
    { x: 0.9, y: 5.0, w: 11.5, h: 0.9, fontSize: 12, italic: true, color: GREY, align: "center" }
  );
}

/* ─── Slide 11 : Problèmes / Actions ───────────────────────────────────── */

function buildProblemes(ctx: SlideCtx): void {
  const { pptx, data } = ctx;
  const s = pptx.addSlide();
  ctx.addHeader(s, "Problèmes rencontrés / Actions correctrices");
  const head = ["Problèmes identifiés", "Causes", "ZS concernées", "Solutions proposées"];
  const rows: PptxGenJS.TableRow[] = [
    head.map((h) => ({ text: h, options: thHeader({ fontSize: 12 }) })),
    ...data.problemes.map((p): PptxGenJS.TableCell[] => [
      { text: p.probleme, options: tdCell({ bold: true }) },
      { text: p.causes, options: tdCell() },
      { text: p.zs, options: tdCell({ align: "center" }) },
      { text: p.solutions, options: tdCell() },
    ]),
  ];
  s.addTable(rows, {
    x: 0.45, y: 1.3, w: W - 0.9, colW: [3.4, 3.0, 2.0, 3.5],
    border: { type: "solid", color: "DEE5EE", pt: 0.5 },
    rowH: 0.55, valign: "middle", fontFace: "Calibri",
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
    align: "center", valign: "middle", fontSize: 38, bold: true, color: "FFFFFF", charSpacing: 4,
  });
  if (pev) s.addImage({ data: pev, x: W / 2 - 1.65, y: H / 2 + 1.45, w: 3.3, h: 0.703 });
}

/* ─── Sous-helpers tableau / cellules / commentaires ───────────────────── */

function thHeader(extra: Partial<PptxGenJS.TableCellProps> = {}): PptxGenJS.TableCellProps {
  return {
    bold: true, color: "FFFFFF", fill: { color: NAVY }, fontSize: 9,
    align: "center", valign: "middle", fontFace: "Calibri", ...extra,
  };
}
function thTotal(extra: Partial<PptxGenJS.TableCellProps> = {}): PptxGenJS.TableCellProps {
  return {
    bold: true, color: "FFFFFF", fill: { color: NAVY_DEEP }, fontSize: 9,
    align: "center", valign: "middle", fontFace: "Calibri", ...extra,
  };
}
function tdCell(extra: Partial<PptxGenJS.TableCellProps> = {}): PptxGenJS.TableCellProps {
  return {
    color: NAVY_DEEP, fontSize: 9, valign: "middle", fontFace: "Calibri",
    fill: { color: "FFFFFF" }, ...extra,
  };
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
  ], { x: 0.7, y: 6.05, w: W - 1.4, h: 0.85, valign: "middle", fontSize: 12 });
}

function coverageComment(rows: CoverageDailyRow[], globalCV: number | null, vaccine: string): string {
  if (rows.length === 0) return `Aucune donnée de couverture ${vaccine} disponible pour ce périmètre.`;
  const below = rows.filter((r) => r.couvGlobale != null && r.couvGlobale < 95);
  const base = `Couverture vaccinale ${vaccine} de ${fmtPct(globalCV)} sur le périmètre sélectionné`;
  if (below.length === 0) return `${base} : objectif ≥ 95 % atteint dans toutes les unités.`;
  const names = below.map((r) => r.unit).slice(0, 4).join(", ");
  const extra = below.length > 4 ? `… (${below.length} unités au total)` : "";
  return `${base}. ${below.length} unité(s) sous le seuil de 95 % : ${names}${extra} — à renforcer par des passages de rattrapage.`;
}

function gestionComment(globalTaux: number | null, vaccine: string, seuil: number): string {
  if (globalTaux == null) return `Données de gestion du vaccin ${vaccine} non disponibles pour ce périmètre.`;
  const base = `Taux de perte ${vaccine} de ${fmtPct(globalTaux)} (seuil acceptable ≤ ${seuil} %)`;
  if (globalTaux < 0) return `${base}. Un taux négatif traduit une saisie irrégulière des flacons utilisés — à corriger dans l'outil de collecte.`;
  if (globalTaux > seuil) return `${base} : dépassement à surveiller, vérifier la chaîne du froid et la saisie des flacons.`;
  return `${base} : performance conforme au seuil.`;
}

// Le helper fmtNum reste exporté pour les usages avancés.
export { fmtNum };
