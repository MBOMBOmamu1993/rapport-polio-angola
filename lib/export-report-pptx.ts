/**
 * Génération du rapport PowerPoint « Campagne de vaccination polio synchronisée
 * avec l'Angola » — reproduit le modèle fourni en ne gardant que la composante
 * polio (nVPO2 et VPOb), co-administration incluse.
 */

import PptxGenJS from "pptxgenjs";

export interface UnitValue {
  unit: string;
  value: number | null;
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
  nvpo2CVByUnit: UnitValue[];
  vpobCVByUnit: UnitValue[];
  recupByUnit: UnitValue[];
  nvpo2PerteByUnit: UnitValue[];
  vpobPerteByUnit: UnitValue[];
  problemes: ProblemeRow[];
  commentNvpo2: string;
  commentVpob: string;
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
      x: 0.55, y: 0.06, w: W - 2.2, h: 0.83, fontSize: 21, bold: true, color: "FFFFFF",
      align: "left", valign: "middle", fontFace: "Calibri",
    });
    if (pev) slide.addImage({ data: pev, x: W - 1.45, y: 0.12, w: 0.72, h: 0.72 });
    if (oms) slide.addImage({ data: oms, x: W - 0.72, y: 0.18, w: 0.6, h: 0.6 });
    slide.addText("Campagne de vaccination polio synchronisée avec l'Angola", {
      x: 0.55, y: H - 0.38, w: W - 3, h: 0.3, fontSize: 8, color: GREY, align: "left",
    });
    slide.addText(data.scopeLabel, {
      x: W - 4.2, y: H - 0.38, w: 4, h: 0.3, fontSize: 8, color: OMS_DARK, align: "right", bold: true,
    });
  };

  // ── Slide 1 : Page de garde ────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.background = { color: OMS_DEEP };
    s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W * 0.42, h: H, fill: { color: OMS_DEEP } });
    if (cover) s.addImage({ data: cover, x: W * 0.42, y: 0, w: W * 0.58, h: H, sizing: { type: "cover", w: W * 0.58, h: H } });
    // Voile dégradé pour lisibilité du titre.
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
  addBarSlide(pptx, addHeader, "Complétude des rapports par Zone de Santé", data.completudeByUnit, OMS, "%", 100, "Source : masque de saisie de la campagne");

  // ── Slide 5 : Couverture nVPO2 ────────────────────────────────────────────────
  addBarSlide(pptx, addHeader, "Couvertures vaccinales nVPO2, par Zone de Santé", data.nvpo2CVByUnit, OMS, "%", 100, data.commentNvpo2, true);

  // ── Slide 6 : Couverture VPOb ─────────────────────────────────────────────────
  addBarSlide(pptx, addHeader, "Couvertures vaccinales VPOb, par Zone de Santé", data.vpobCVByUnit, OMS_DARK, "%", 100, data.commentVpob, true);

  // ── Slide 7 : Récupération PEV routine ────────────────────────────────────────
  addBarSlide(pptx, addHeader, "Récupération des enfants en PEV de routine (co-administration)", data.recupByUnit, GREEN, "", null, "Enfants récupérés et orientés vers le PEV de routine pendant la campagne polio");

  // ── Slide 8 : Gestion vaccin nVPO2 ────────────────────────────────────────────
  addBarSlide(pptx, addHeader, "Gestion des vaccins : nVPO2 (taux de perte par ZS)", data.nvpo2PerteByUnit, ORANGE, "%", null, "Seuil acceptable du taux de perte nVPO2 : ≤ 11 %");

  // ── Slide 9 : Gestion vaccin VPOb ─────────────────────────────────────────────
  addBarSlide(pptx, addHeader, "Gestion des vaccins : VPOb (taux de perte par ZS)", data.vpobPerteByUnit, ORANGE, "%", null, "Seuil acceptable du taux de perte VPOb : ≤ 10 %");

  // ── Slide 10 : Surveillance MAPI ──────────────────────────────────────────────
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

function addBarSlide(
  pptx: PptxGenJS,
  addHeader: (s: PptxGenJS.Slide, t: string) => void,
  title: string,
  units: UnitValue[],
  color: string,
  unitSuffix: string,
  max: number | null,
  comment: string,
  threshold = false
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
      barDir: "col",
      chartColors: [color],
      showValue: true,
      dataLabelColor: OMS_DEEP,
      dataLabelFontSize: 9,
      dataLabelFormatCode: unitSuffix === "%" ? '0.0"%"' : "0",
      valAxisMaxVal: max ?? undefined,
      valAxisMinVal: 0,
      catAxisLabelFontSize: 9,
      catAxisLabelRotate: labels.length > 8 ? 45 : 0,
      valAxisLabelFontSize: 9,
      showLegend: false,
      ...(threshold && max === 100
        ? { valGridLine: { style: "dash", color: "E23636", size: 1 } }
        : {}),
    });
  }

  if (comment) {
    s.addShape(pptx.ShapeType.roundRect, { x: 0.5, y: 6.0, w: W - 1, h: 0.85, fill: { color: "E6F5FC" }, line: { color: OMS, width: 1 }, rectRadius: 0.05 });
    s.addText([
      { text: "Commentaire : ", options: { bold: true, color: OMS_DARK } },
      { text: comment, options: { color: OMS_DEEP } },
    ], { x: 0.7, y: 6.0, w: W - 1.4, h: 0.85, valign: "middle", fontSize: 12 });
  }
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
