/**
 * Génération du rapport PowerPoint « Campagne de vaccination polio synchronisée
 * avec l'Angola » — reproduit le modèle officiel en ne gardant que la composante
 * polio (nVPO2 et VPOb), co-administration incluse. Chaque indicateur est présenté
 * par Zone de Santé sous forme de tableau + graphique + commentaire dynamique.
 */

import PptxGenJS from "pptxgenjs";

export interface UnitValue {
  unit: string;
  value: number | null;
}

export interface CoverageRow {
  unit: string;
  cible: number;
  vacc: number;
  cv: number | null;
}

export interface GestionRow {
  unit: string;
  flaconsUtil: number;
  perdus: number;
  vacc: number;
  taux: number | null;
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
    nvpo2Flacons: number;
    nvpo2Perte: number | null;
    vpobFlacons: number;
    vpobPerte: number | null;
    recup: number;
    mapiMineures: number;
    mapiGraves: number;
  };
  completudeByUnit: UnitValue[];
  nvpo2Coverage: CoverageRow[];
  vpobCoverage: CoverageRow[];
  nvpo2Gestion: GestionRow[];
  vpobGestion: GestionRow[];
  recupByUnit: UnitValue[];
  problemes: ProblemeRow[];
  completudeMapPng?: string | null;
}

// ─── Design tokens (OMS) ────────────────────────────────────────────────────
const OMS = "0093D5";
const OMS_DARK = "005A82";
const OMS_DEEP = "003D57";
const GREY = "64748B";
const GREEN = "22B457";
const RED = "E23636";
const ORANGE = "F29E0B";
const LIGHT = "F1F5F9";

const W = 13.333;
const H = 7.5;

async function toDataUrl(path: string): Promise<string | null> {
  try {
    const res = await fetch(path);
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

function fmtPct(n: number | null, d = 1): string {
  if (n === null || !Number.isFinite(n)) return "—";
  return `${n.toFixed(d).replace(".", ",")} %`;
}
function fmtInt(n: number): string {
  return Math.round(n || 0).toLocaleString("fr-FR");
}

function covTone(v: number | null): string {
  if (v === null) return GREY;
  if (v >= 95) return GREEN;
  if (v >= 90) return ORANGE;
  return RED;
}
function lossTone(v: number | null): string {
  if (v === null) return GREY;
  if (v < 0 || v > 15) return RED;
  if (v > 10) return ORANGE;
  return GREEN;
}
function slug(s: string): string {
  return s.replace(/[^a-zA-Z0-9]+/g, "_").replace(/^_|_$/g, "").slice(0, 40) || "rapport";
}

export async function exportReportPPT(data: ReportData): Promise<void> {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: "WIDE", width: W, height: H });
  pptx.layout = "WIDE";
  pptx.author = "PEV — RD Congo";
  pptx.company = "Programme Élargi de Vaccination";

  const [cover, pev, oms] = await Promise.all([
    toDataUrl("/cover-polio.png"),
    toDataUrl("/logo/pev.png"),
    toDataUrl("/logo/oms.png"),
  ]);

  // En-tête + pied réutilisables sur les slides de contenu.
  const addHeader = (slide: PptxGenJS.Slide, title: string) => {
    slide.background = { color: "FFFFFF" };
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: 0.95, fill: { color: OMS } });
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0.95, w: W, h: 0.06, fill: { color: OMS_DEEP } });
    slide.addText(title, {
      x: 0.55, y: 0.06, w: W - 2.2, h: 0.83, fontSize: 20, bold: true, color: "FFFFFF",
      align: "left", valign: "middle", fontFace: "Calibri",
    });
    if (pev) slide.addImage({ data: pev, x: W - 1.45, y: 0.12, w: 0.72, h: 0.72 });
    if (oms) slide.addImage({ data: oms, x: W - 0.72, y: 0.18, w: 0.6, h: 0.6 });
    slide.addText("Campagne de vaccination polio synchronisée avec l'Angola", {
      x: 0.55, y: H - 0.36, w: W - 3, h: 0.28, fontSize: 8, color: GREY, align: "left",
    });
    slide.addText(data.scopeLabel, {
      x: W - 4.2, y: H - 0.36, w: 4, h: 0.28, fontSize: 8, color: OMS_DARK, align: "right", bold: true,
    });
  };

  // ── Slide 1 : Page de garde ────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.background = { color: OMS_DEEP };
    if (cover) s.addImage({ data: cover, x: W * 0.42, y: 0, w: W * 0.58, h: H, sizing: { type: "cover", w: W * 0.58, h: H } });
    s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W * 0.46, h: H, fill: { color: OMS_DEEP, transparency: 8 } });

    if (pev) s.addImage({ data: pev, x: 0.6, y: 0.5, w: 1.0, h: 1.0 });
    if (oms) s.addImage({ data: oms, x: 1.75, y: 0.62, w: 0.78, h: 0.78 });

    s.addText("RAPPORT DES RÉSULTATS", {
      x: 0.6, y: 2.0, w: 5.4, h: 0.5, fontSize: 16, color: OMS, bold: true, charSpacing: 2,
    });
    s.addText("Campagne de vaccination polio synchronisée avec l'Angola", {
      x: 0.6, y: 2.5, w: 5.4, h: 1.5, fontSize: 30, color: "FFFFFF", bold: true, fontFace: "Calibri",
      lineSpacingMultiple: 1.0,
    });
    s.addShape(pptx.ShapeType.rect, { x: 0.62, y: 4.05, w: 1.6, h: 0.06, fill: { color: OMS } });
    s.addText(
      [
        { text: `Province : ${data.province}`, options: { bold: true, fontSize: 16, color: "FFFFFF", breakLine: true } },
        { text: data.periode || "Période de la campagne", options: { fontSize: 13, color: "CCE9F7", breakLine: true } },
        { text: "Vaccins : nVPO2 et VPOb (co-administration)", options: { fontSize: 12, color: "99D4EF" } },
      ],
      { x: 0.6, y: 4.3, w: 5.4, h: 1.4, valign: "top" }
    );
    s.addText("Programme Élargi de Vaccination — RD Congo", {
      x: 0.6, y: H - 0.7, w: 5.4, h: 0.4, fontSize: 11, color: "99D4EF", italic: true,
    });
    s.addText(data.dateLabel, {
      x: W * 0.42 + 0.2, y: H - 0.7, w: W * 0.58 - 0.4, h: 0.4, fontSize: 11, color: "FFFFFF", align: "right", bold: true,
    });
  }

  // ── Slide 2 : Plan de présentation ──────────────────────────────────────────
  {
    const s = pptx.addSlide();
    addHeader(s, "Plan de présentation");
    const items = [
      "Points saillants de la campagne",
      "Complétude des rapports par Zone de Santé",
      "Couvertures vaccinales nVPO2",
      "Couvertures vaccinales VPOb",
      "Récupération des enfants en PEV de routine (co-administration)",
      "Gestion du vaccin nVPO2",
      "Gestion du vaccin VPOb",
      "Surveillance des MAPI",
      "Problèmes rencontrés / Actions correctrices",
    ];
    items.forEach((it, i) => {
      const y = 1.45 + i * 0.6;
      s.addShape(pptx.ShapeType.ellipse, { x: 1.2, y, w: 0.42, h: 0.42, fill: { color: OMS } });
      s.addText(String(i + 1), { x: 1.2, y, w: 0.42, h: 0.42, align: "center", valign: "middle", color: "FFFFFF", bold: true, fontSize: 14 });
      s.addText(it, { x: 1.8, y, w: 9.8, h: 0.42, valign: "middle", fontSize: 15, color: OMS_DEEP });
    });
  }

  // ── Slide 3 : Points saillants ───────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    addHeader(s, "Points saillants");
    const sa = data.saillants;
    const col = (x: number, w: number, header: string, lines: { label: string; value: string; tone?: string }[]) => {
      s.addShape(pptx.ShapeType.roundRect, { x, y: 1.3, w, h: 5.5, fill: { color: LIGHT }, line: { color: "E2E8F0", width: 1 }, rectRadius: 0.08 });
      s.addShape(pptx.ShapeType.rect, { x, y: 1.3, w, h: 0.5, fill: { color: OMS } });
      s.addText(header, { x: x + 0.2, y: 1.3, w: w - 0.4, h: 0.5, valign: "middle", color: "FFFFFF", bold: true, fontSize: 13 });
      lines.forEach((ln, i) => {
        const y = 2.0 + i * 0.62;
        s.addText(ln.label, { x: x + 0.25, y, w: w - 2.0, h: 0.55, valign: "middle", fontSize: 12, color: OMS_DEEP });
        s.addText(ln.value, { x: x + w - 2.0, y, w: 1.75, h: 0.55, valign: "middle", align: "right", fontSize: 14, bold: true, color: ln.tone ?? OMS_DARK });
      });
    };
    col(0.45, 6.1, "COMPLÉTUDE & VACCINATION", [
      { label: "Rapports attendus", value: fmtInt(sa.completudeAttendus) },
      { label: "Rapports reçus", value: fmtInt(sa.completudeRecus) },
      { label: "Complétude des rapports", value: fmtPct(sa.completude), tone: covTone(sa.completude) },
      { label: "Vaccinés nVPO2 / Cible", value: `${fmtInt(sa.nvpo2Vacc)} / ${fmtInt(sa.nvpo2Cible)}` },
      { label: "Couverture nVPO2", value: fmtPct(sa.nvpo2CV), tone: covTone(sa.nvpo2CV) },
      { label: "Vaccinés VPOb / Cible", value: `${fmtInt(sa.vpobVacc)} / ${fmtInt(sa.vpobCible)}` },
      { label: "Couverture VPOb", value: fmtPct(sa.vpobCV), tone: covTone(sa.vpobCV) },
    ]);
    col(6.75, 6.1, "GESTION DES VACCINS & MAPI", [
      { label: "Flacons nVPO2 utilisés", value: fmtInt(sa.nvpo2Flacons) },
      { label: "Taux de perte nVPO2", value: fmtPct(sa.nvpo2Perte), tone: lossTone(sa.nvpo2Perte) },
      { label: "Flacons VPOb utilisés", value: fmtInt(sa.vpobFlacons) },
      { label: "Taux de perte VPOb", value: fmtPct(sa.vpobPerte), tone: lossTone(sa.vpobPerte) },
      { label: "Récupérations PEV de routine", value: fmtInt(sa.recup) },
      { label: "MAPI mineures", value: fmtInt(sa.mapiMineures), tone: GREY },
      { label: "MAPI graves", value: fmtInt(sa.mapiGraves), tone: sa.mapiGraves ? RED : GREY },
    ]);
  }

  // ── Slide 4 : Complétude des rapports par ZS ──────────────────────────────────
  addBarSlide(
    pptx, addHeader,
    "Complétude des rapports par Zone de Santé",
    data.completudeByUnit, OMS, "%", 100,
    `Complétude globale de ${fmtPct(data.saillants.completude)} (${fmtInt(data.saillants.completudeRecus)} rapports reçus sur ${fmtInt(data.saillants.completudeAttendus)} attendus).`
  );

  // ── Slide : Spatialisation de la complétude (carte choroplèthe des ZS) ─────────
  if (data.completudeMapPng) {
    const s = pptx.addSlide();
    addHeader(s, "Spatialisation de la complétude par Zone de Santé");
    s.addImage({ data: data.completudeMapPng, x: 1.4, y: 1.15, w: 7.2, h: 5.4, sizing: { type: "contain", w: 7.2, h: 5.4 } });
    // Légende.
    const legend: { c: string; t: string }[] = [
      { c: "E23636", t: "< 80 %" },
      { c: "F29E0B", t: "80 – 89 %" },
      { c: "66BEE7", t: "90 – 99 %" },
      { c: "0093D5", t: "100 %" },
      { c: "E2E8F0", t: "Pas de donnée" },
    ];
    legend.forEach((l, i) => {
      const y = 2.2 + i * 0.55;
      s.addShape(pptx.ShapeType.rect, { x: 9.4, y, w: 0.4, h: 0.36, fill: { color: l.c }, line: { color: "CBD5E1", width: 0.5 } });
      s.addText(l.t, { x: 9.95, y, w: 2.8, h: 0.36, valign: "middle", fontSize: 12, color: OMS_DEEP });
    });
    s.addText("Complétude des rapports de vaccination par ZS sur le périmètre sélectionné.", {
      x: 9.4, y: 5.2, w: 3.4, h: 1.0, fontSize: 10, italic: true, color: GREY,
    });
  }

  // ── Slide 5 : Couvertures vaccinales nVPO2 (tableau + graphique + commentaire) ─
  addCoverageSlide(pptx, addHeader, "Couvertures vaccinales nVPO2, par Zone de Santé", data.nvpo2Coverage, data.byUnitLabel, "nVPO2", OMS, data.saillants.nvpo2CV);

  // ── Slide 6 : Couvertures vaccinales VPOb ─────────────────────────────────────
  addCoverageSlide(pptx, addHeader, "Couvertures vaccinales VPOb, par Zone de Santé", data.vpobCoverage, data.byUnitLabel, "VPOb", OMS_DARK, data.saillants.vpobCV);

  // ── Slide 7 : Récupération PEV de routine ─────────────────────────────────────
  addBarSlide(
    pptx, addHeader,
    "Récupération des enfants en PEV de routine (co-administration)",
    data.recupByUnit, GREEN, "", null,
    `${fmtInt(data.saillants.recup)} enfants récupérés et orientés vers le PEV de routine pendant la campagne polio.`
  );

  // ── Slide 8 : Gestion du vaccin nVPO2 (tableau + graphique + commentaire) ──────
  addGestionSlide(pptx, addHeader, "Gestion des vaccins : nVPO2", data.nvpo2Gestion, data.byUnitLabel, "nVPO2", data.saillants.nvpo2Perte, 11);

  // ── Slide 9 : Gestion du vaccin VPOb ──────────────────────────────────────────
  addGestionSlide(pptx, addHeader, "Gestion des vaccins : VPOb", data.vpobGestion, data.byUnitLabel, "VPOb", data.saillants.vpobPerte, 10);

  // ── Slide 10 : Surveillance des MAPI ──────────────────────────────────────────
  {
    const s = pptx.addSlide();
    addHeader(s, "Surveillance des MAPI");
    const sa = data.saillants;
    const card = (x: number, label: string, value: string, color: string) => {
      s.addShape(pptx.ShapeType.roundRect, { x, y: 2.4, w: 3.6, h: 2.4, fill: { color: LIGHT }, line: { color, width: 2 }, rectRadius: 0.1 });
      s.addText(value, { x, y: 2.7, w: 3.6, h: 1.2, align: "center", valign: "middle", fontSize: 48, bold: true, color });
      s.addText(label, { x: x + 0.2, y: 3.95, w: 3.2, h: 0.6, align: "center", valign: "top", fontSize: 14, color: OMS_DEEP });
    };
    card(1.1, "MAPI mineures notifiées", fmtInt(sa.mapiMineures), OMS);
    card(5.05, "MAPI graves notifiées", fmtInt(sa.mapiGraves), sa.mapiGraves ? RED : GREEN);
    card(9.0, "Récupérations PEV", fmtInt(sa.recup), GREEN);
    s.addText(
      "Toute MAPI grave doit faire l'objet d'une investigation immédiate et d'une notification au niveau supérieur conformément au guide de surveillance.",
      { x: 1.1, y: 5.2, w: 11.4, h: 0.8, fontSize: 12, italic: true, color: GREY, align: "center" }
    );
  }

  // ── Slide 11 : Problèmes / Actions ────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    addHeader(s, "Problèmes rencontrés / Actions correctrices");
    const head = ["Problèmes identifiés", "Causes", "ZS concernées", "Solutions proposées"];
    const rows: PptxGenJS.TableRow[] = [
      head.map((h) => ({ text: h, options: { bold: true, color: "FFFFFF", fill: { color: OMS_DARK }, fontSize: 12, valign: "middle", align: "center" } })),
      ...data.problemes.map((p) => [
        { text: p.probleme, options: { fontSize: 11, color: OMS_DEEP } },
        { text: p.causes, options: { fontSize: 11, color: OMS_DEEP } },
        { text: p.zs, options: { fontSize: 11, color: OMS_DEEP, align: "center" as const } },
        { text: p.solutions, options: { fontSize: 11, color: OMS_DEEP } },
      ]),
    ];
    s.addTable(rows, {
      x: 0.45, y: 1.3, w: W - 0.9, colW: [3.4, 3.0, 2.0, 3.5],
      border: { type: "solid", color: "CBD5E1", pt: 0.5 },
      rowH: 0.5, valign: "middle", fill: { color: "FFFFFF" },
    });
  }

  // ── Slide 12 : Merci ──────────────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.background = { color: OMS };
    s.addShape(pptx.ShapeType.rect, { x: 0, y: H / 2 - 0.9, w: W, h: 1.8, fill: { color: OMS_DARK } });
    s.addText("MERCI POUR VOTRE ATTENTION", {
      x: 0, y: H / 2 - 0.9, w: W, h: 1.8, align: "center", valign: "middle", fontSize: 36, bold: true, color: "FFFFFF", charSpacing: 2,
    });
    if (pev) s.addImage({ data: pev, x: W / 2 - 0.5, y: H / 2 + 1.2, w: 1.0, h: 1.0 });
  }

  await pptx.writeFile({ fileName: `Rapport_Polio_Angola_${slug(data.scopeLabel)}.pptx` });
}

// ── Slide « couverture » : graphique (gauche) + tableau (droite) + commentaire ──
function addCoverageSlide(
  pptx: PptxGenJS,
  addHeader: (s: PptxGenJS.Slide, t: string) => void,
  title: string,
  rows: CoverageRow[],
  unitLabel: string,
  vaccine: string,
  color: string,
  globalCV: number | null
): void {
  const s = pptx.addSlide();
  addHeader(s, title);

  const labels = rows.map((r) => r.unit);
  const values = rows.map((r) => (r.cv == null ? 0 : +r.cv.toFixed(1)));
  const hasData = rows.length > 0 && values.some((v) => v !== 0);

  // Graphique (colonnes) — couverture % par unité.
  if (hasData) {
    s.addChart(pptx.ChartType.bar, [{ name: "Couverture", labels, values }], {
      x: 0.4, y: 1.2, w: 6.3, h: 4.6,
      barDir: "col", chartColors: [color], showValue: true,
      dataLabelColor: OMS_DEEP, dataLabelFontSize: 8, dataLabelFormatCode: '0"%"',
      valAxisMaxVal: 100, valAxisMinVal: 0,
      catAxisLabelFontSize: 8, catAxisLabelRotate: labels.length > 6 ? 45 : 0,
      valAxisLabelFontSize: 8, showLegend: false,
      valGridLine: { style: "dash", color: "E23636", size: 1 },
    });
  } else {
    s.addText("Aucune donnée disponible pour ce périmètre.", {
      x: 0.4, y: 3.0, w: 6.3, h: 1, align: "center", fontSize: 13, italic: true, color: GREY,
    });
  }

  // Tableau (droite) — Cible / Vaccinés / CV % par unité + Total.
  const sumCible = rows.reduce((a, r) => a + r.cible, 0);
  const sumVacc = rows.reduce((a, r) => a + r.vacc, 0);
  const totalCV = sumCible ? (sumVacc / sumCible) * 100 : null;
  const head = [unitLabel, "Cible", "Vaccinés", "CV"];
  const trows: PptxGenJS.TableRow[] = [
    head.map((h) => ({ text: h, options: { bold: true, color: "FFFFFF", fill: { color: OMS_DARK }, fontSize: 9, align: "center" as const, valign: "middle" as const } })),
    ...rows.map((r) => [
      { text: r.unit, options: { fontSize: 8, color: OMS_DEEP } },
      { text: fmtInt(r.cible), options: { fontSize: 8, color: OMS_DEEP, align: "right" as const } },
      { text: fmtInt(r.vacc), options: { fontSize: 8, color: OMS_DEEP, align: "right" as const } },
      { text: fmtPct(r.cv), options: { fontSize: 8, color: covTone(r.cv), align: "right" as const, bold: true } },
    ]),
    [
      { text: "Total", options: { fontSize: 9, color: "FFFFFF", fill: { color: OMS }, bold: true } },
      { text: fmtInt(sumCible), options: { fontSize: 9, color: "FFFFFF", fill: { color: OMS }, align: "right" as const, bold: true } },
      { text: fmtInt(sumVacc), options: { fontSize: 9, color: "FFFFFF", fill: { color: OMS }, align: "right" as const, bold: true } },
      { text: fmtPct(totalCV), options: { fontSize: 9, color: "FFFFFF", fill: { color: OMS }, align: "right" as const, bold: true } },
    ],
  ];
  s.addTable(trows, {
    x: 6.9, y: 1.2, w: 6.0, colW: [2.4, 1.2, 1.4, 1.0],
    border: { type: "solid", color: "CBD5E1", pt: 0.5 },
    autoPage: false, valign: "middle", fontFace: "Calibri",
  });

  addCommentBar(pptx, s, coverageComment(rows, globalCV, vaccine));
}

// ── Slide « gestion vaccin » : tableau (gauche) + graphique perte (droite) ──────
function addGestionSlide(
  pptx: PptxGenJS,
  addHeader: (s: PptxGenJS.Slide, t: string) => void,
  title: string,
  rows: GestionRow[],
  unitLabel: string,
  vaccine: string,
  globalTaux: number | null,
  seuil: number
): void {
  const s = pptx.addSlide();
  addHeader(s, title);

  // Tableau (gauche) : Flacons utilisés / Perdus / Enfants vaccinés / Taux de perte.
  const sumUtil = rows.reduce((a, r) => a + r.flaconsUtil, 0);
  const sumPerdus = rows.reduce((a, r) => a + r.perdus, 0);
  const sumVacc = rows.reduce((a, r) => a + r.vacc, 0);
  const head = [unitLabel, "Flac. util.", "Perdus", "Vaccinés", "Perte"];
  const trows: PptxGenJS.TableRow[] = [
    head.map((h) => ({ text: h, options: { bold: true, color: "FFFFFF", fill: { color: OMS_DARK }, fontSize: 9, align: "center" as const, valign: "middle" as const } })),
    ...rows.map((r) => [
      { text: r.unit, options: { fontSize: 8, color: OMS_DEEP } },
      { text: fmtInt(r.flaconsUtil), options: { fontSize: 8, color: OMS_DEEP, align: "right" as const } },
      { text: fmtInt(r.perdus), options: { fontSize: 8, color: OMS_DEEP, align: "right" as const } },
      { text: fmtInt(r.vacc), options: { fontSize: 8, color: OMS_DEEP, align: "right" as const } },
      { text: fmtPct(r.taux), options: { fontSize: 8, color: lossTone(r.taux), align: "right" as const, bold: true } },
    ]),
    [
      { text: "Total", options: { fontSize: 9, color: "FFFFFF", fill: { color: OMS }, bold: true } },
      { text: fmtInt(sumUtil), options: { fontSize: 9, color: "FFFFFF", fill: { color: OMS }, align: "right" as const, bold: true } },
      { text: fmtInt(sumPerdus), options: { fontSize: 9, color: "FFFFFF", fill: { color: OMS }, align: "right" as const, bold: true } },
      { text: fmtInt(sumVacc), options: { fontSize: 9, color: "FFFFFF", fill: { color: OMS }, align: "right" as const, bold: true } },
      { text: fmtPct(globalTaux), options: { fontSize: 9, color: "FFFFFF", fill: { color: OMS }, align: "right" as const, bold: true } },
    ],
  ];
  s.addTable(trows, {
    x: 0.35, y: 1.2, w: 6.7, colW: [1.9, 1.3, 1.0, 1.4, 1.1],
    border: { type: "solid", color: "CBD5E1", pt: 0.5 },
    autoPage: false, valign: "middle", fontFace: "Calibri",
  });

  // Graphique (droite) : taux de perte (%) par unité — barres horizontales.
  const labels = rows.map((r) => r.unit);
  const values = rows.map((r) => (r.taux == null ? 0 : +r.taux.toFixed(2)));
  const hasData = rows.length > 0 && values.some((v) => v !== 0);
  if (hasData) {
    s.addChart(pptx.ChartType.bar, [{ name: `Taux de perte ${vaccine}`, labels, values }], {
      x: 7.2, y: 1.2, w: 5.8, h: 4.6,
      barDir: "bar", chartColors: [ORANGE], showValue: true,
      dataLabelColor: OMS_DEEP, dataLabelFontSize: 8, dataLabelFormatCode: '0.0"%"',
      valAxisMinVal: 0, catAxisLabelFontSize: 8, valAxisLabelFontSize: 8, showLegend: false,
    });
  } else {
    s.addText("Aucune donnée disponible pour ce périmètre.", {
      x: 7.2, y: 3.0, w: 5.8, h: 1, align: "center", fontSize: 13, italic: true, color: GREY,
    });
  }

  addCommentBar(pptx, s, gestionComment(globalTaux, vaccine, seuil));
}

function addBarSlide(
  pptx: PptxGenJS,
  addHeader: (s: PptxGenJS.Slide, t: string) => void,
  title: string,
  units: UnitValue[],
  color: string,
  unitSuffix: string,
  max: number | null,
  comment: string
): void {
  const s = pptx.addSlide();
  addHeader(s, title);

  const labels = units.map((u) => u.unit);
  const values = units.map((u) => (u.value == null ? 0 : +u.value.toFixed(1)));

  if (labels.length === 0 || values.every((v) => v === 0)) {
    s.addText("Aucune donnée disponible pour ce périmètre (les chiffres seront affichés dès la saisie dans le masque).", {
      x: 1, y: 3.2, w: W - 2, h: 1, align: "center", fontSize: 14, italic: true, color: GREY,
    });
  } else {
    s.addChart(pptx.ChartType.bar, [{ name: title, labels, values }], {
      x: 0.5, y: 1.25, w: W - 1, h: comment ? 4.6 : 5.4,
      barDir: "col", chartColors: [color], showValue: true,
      dataLabelColor: OMS_DEEP, dataLabelFontSize: 9,
      dataLabelFormatCode: unitSuffix === "%" ? '0.0"%"' : "0",
      valAxisMaxVal: max ?? undefined, valAxisMinVal: 0,
      catAxisLabelFontSize: 9, catAxisLabelRotate: labels.length > 8 ? 45 : 0,
      valAxisLabelFontSize: 9, showLegend: false,
      ...(max === 100 ? { valGridLine: { style: "dash", color: "E23636", size: 1 } } : {}),
    });
  }

  if (comment) addCommentBar(pptx, s, comment);
}

function addCommentBar(pptx: PptxGenJS, s: PptxGenJS.Slide, comment: string): void {
  s.addShape(pptx.ShapeType.roundRect, { x: 0.5, y: 6.0, w: W - 1, h: 0.9, fill: { color: "E6F5FC" }, line: { color: OMS, width: 1 }, rectRadius: 0.05 });
  s.addText([
    { text: "Commentaire : ", options: { bold: true, color: OMS_DARK } },
    { text: comment, options: { color: OMS_DEEP } },
  ], { x: 0.7, y: 6.0, w: W - 1.4, h: 0.9, valign: "middle", fontSize: 12 });
}

function coverageComment(rows: CoverageRow[], globalCV: number | null, vaccine: string): string {
  if (rows.length === 0) return `Aucune donnée de couverture ${vaccine} disponible pour ce périmètre.`;
  const below = rows.filter((r) => r.cv != null && r.cv < 95);
  const base = `Couverture vaccinale ${vaccine} de ${fmtPct(globalCV)} sur le périmètre sélectionné`;
  if (below.length === 0) return `${base} : objectif de 95 % atteint dans toutes les zones.`;
  const names = below.map((r) => r.unit).slice(0, 4).join(", ");
  const extra = below.length > 4 ? `… (${below.length} ZS au total)` : "";
  return `${base}. ${below.length} zone(s) sous le seuil de 95 % : ${names}${extra} — à renforcer par des passages de rattrapage.`;
}

function gestionComment(globalTaux: number | null, vaccine: string, seuil: number): string {
  const base = `Taux de perte ${vaccine} de ${fmtPct(globalTaux)} (seuil acceptable ≤ ${seuil} %)`;
  if (globalTaux == null) return `Données de gestion du vaccin ${vaccine} non disponibles pour ce périmètre.`;
  if (globalTaux < 0) return `${base}. Un taux négatif traduit une saisie irrégulière des flacons utilisés — à corriger dans l'outil de collecte.`;
  if (globalTaux > seuil) return `${base} : dépassement à surveiller, vérifier la chaîne du froid et la saisie des flacons.`;
  return `${base} : performance conforme au seuil.`;
}
