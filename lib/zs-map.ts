/**
 * Rendu d'une carte des Zones de Santé de la RDC : les ZS du périmètre
 * sélectionné sont colorées selon leur complétude obtenue
 * (< 60 % rouge, 60–79,9 % jaune, ≥ 80 % vert), le reste du territoire en gris.
 *
 * Le rendu est entièrement côté navigateur : on récupère le TopoJSON, on le
 * projette en SVG (d3-geo), puis on convertit le SVG en PNG (data URL) pour
 * l'embarquer dans le PowerPoint. En cas d'échec (réseau, format), on renvoie
 * null et la carte est simplement omise du rapport.
 */

import { geoMercator, geoPath } from "d3-geo";
import { feature } from "topojson-client";
import type { Topology, GeometryCollection } from "topojson-specification";
import type { Feature, FeatureCollection, Geometry } from "geojson";

const TOPOJSON_URL =
  "https://gist.githubusercontent.com/MBOMBOmamu1993/1297c206c046ee018d5ed6c392d6c20f/raw/24ce95b2935d2b4cc4ef71701138218ca870ff01/rdc_zs.topojson";

const RED = "#E23636";
const YELLOW = "#F1C40F";
const GREEN = "#22B457";
const BASE_FILL = "#F1F5F9"; // reste du territoire (hors périmètre)
const STROKE = "#64748B";

export function normalizeZS(name: string): string {
  return (name || "")
    .toString()
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .toUpperCase()
    .replace(/ZONE DE SANTE/g, "")
    .replace(/[^A-Z0-9]/g, "");
}

/** Couleur de la complétude (%) selon les seuils du rapport. */
export function colorForCompletude(v: number | undefined): string {
  if (v == null || !Number.isFinite(v)) return BASE_FILL;
  if (v >= 80) return GREEN;
  if (v >= 60) return YELLOW;
  return RED;
}

/** Distance d'édition (Levenshtein) bornée — utilisée pour harmoniser les noms. */
function levenshtein(a: string, b: string): number {
  const m = a.length;
  const n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;
  let prev = Array.from({ length: n + 1 }, (_, j) => j);
  let cur = new Array<number>(n + 1).fill(0);
  for (let i = 1; i <= m; i++) {
    cur[0] = i;
    for (let j = 1; j <= n; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      cur[j] = Math.min(cur[j - 1] + 1, prev[j] + 1, prev[j - 1] + cost);
    }
    [prev, cur] = [cur, prev];
  }
  return prev[n];
}

/**
 * Harmonise un nom de ZS du masque avec les noms du fond cartographique :
 * essaie une correspondance exacte (normalisée) puis, à défaut, la ZS la plus
 * proche par distance d'édition. Renvoie l'index de la feature ou -1.
 */
function matchFeature(target: string, featNames: string[]): number {
  const exact = featNames.indexOf(target);
  if (exact >= 0) return exact;
  // Tolérance : un nom court accepte 1–2 différences, un nom long un peu plus.
  const thr = Math.max(2, Math.floor(target.length * 0.25));
  let best = -1;
  let bestD = Infinity;
  for (let i = 0; i < featNames.length; i++) {
    const fn = featNames[i];
    if (!fn) continue;
    // Court-circuit : inclusion mutuelle (ex. préfixe/suffixe commun).
    if (fn.includes(target) || target.includes(fn)) {
      const d = Math.abs(fn.length - target.length);
      if (d < bestD) { bestD = d; best = i; }
      continue;
    }
    const d = levenshtein(target, fn);
    if (d < bestD) { bestD = d; best = i; }
  }
  return best >= 0 && bestD <= thr ? best : -1;
}

/** Détecte la propriété portant le nom de la ZS dans les features du TopoJSON. */
function detectNameKey(features: Feature[], dataKeys: Set<string>): string | null {
  if (features.length === 0) return null;
  const props = features[0].properties ?? {};
  const candidates = Object.keys(props).filter((k) => typeof props[k] === "string");
  let best: string | null = null;
  let bestScore = -1;
  for (const key of candidates) {
    let score = 0;
    for (const f of features) {
      const val = f.properties?.[key];
      if (typeof val === "string" && dataKeys.has(normalizeZS(val))) score++;
    }
    if (score > bestScore) {
      bestScore = score;
      best = key;
    }
  }
  if (bestScore > 0) return best;
  // Repli : noms de propriété usuels.
  return candidates.find((k) => /nom|name|zs|zone/i.test(k)) ?? candidates[0] ?? null;
}

function svgToPng(svg: string, width: number, height: number): Promise<string | null> {
  return new Promise((resolve) => {
    const img = new Image();
    const url = "data:image/svg+xml;base64," + btoa(unescape(encodeURIComponent(svg)));
    img.onload = () => {
      const canvas = document.createElement("canvas");
      canvas.width = width;
      canvas.height = height;
      const ctx = canvas.getContext("2d");
      if (!ctx) return resolve(null);
      ctx.fillStyle = "#FFFFFF";
      ctx.fillRect(0, 0, width, height);
      ctx.drawImage(img, 0, 0, width, height);
      resolve(canvas.toDataURL("image/png"));
    };
    img.onerror = () => resolve(null);
    img.src = url;
  });
}

/**
 * @param completudeByZS clé = nom de ZS normalisé (via normalizeZS), valeur = % complétude.
 *   Seules les ZS présentes dans la map (= périmètre sélectionné) sont colorées ;
 *   le reste du territoire reste en gris.
 * @returns data URL PNG, ou null si la carte ne peut pas être produite.
 */
export async function renderZSMap(completudeByZS: Map<string, number>): Promise<string | null> {
  try {
    const res = await fetch(TOPOJSON_URL, { cache: "force-cache" });
    if (!res.ok) return null;
    const topo = (await res.json()) as Topology;
    const objNames = Object.keys(topo.objects ?? {});
    if (objNames.length === 0) return null;
    const objName = objNames.find((n) => /zone/i.test(n)) ?? objNames[0];
    const fc = feature(topo, topo.objects[objName] as GeometryCollection) as FeatureCollection<Geometry>;
    if (!fc.features?.length) return null;

    const nameKey = detectNameKey(fc.features, new Set(completudeByZS.keys()));

    // Noms de ZS du fond, normalisés, indexés par feature.
    const featNames = fc.features.map((f) =>
      nameKey ? normalizeZS((f.properties?.[nameKey] as string) ?? "") : ""
    );

    // Pour chaque ZS du masque, on retrouve la feature correspondante (avec
    // harmonisation tolérante des orthographes) et on lui affecte sa complétude.
    const valueByFeature = new Map<number, number>();
    for (const [key, v] of completudeByZS) {
      const idx = matchFeature(key, featNames);
      if (idx >= 0) valueByFeature.set(idx, v);
    }

    const width = 1100;
    const height = 1000;
    const projection = geoMercator().fitSize([width, height], fc);
    const path = geoPath(projection);

    // On dessine d'abord les ZS de fond, puis les ZS du périmètre par-dessus
    // (contour net) pour que leur couleur de complétude ressorte.
    const base: string[] = [];
    const hi: string[] = [];
    fc.features.forEach((f, i) => {
      const d = path(f);
      if (!d) return;
      const v = valueByFeature.get(i);
      if (v != null) {
        hi.push(`<path d="${d}" fill="${colorForCompletude(v)}" stroke="${STROKE}" stroke-width="1.1"/>`);
      } else {
        base.push(`<path d="${d}" fill="${BASE_FILL}" stroke="${STROKE}" stroke-width="0.6"/>`);
      }
    });

    const svg =
      `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">` +
      `<rect width="${width}" height="${height}" fill="#FFFFFF"/>${base.join("")}${hi.join("")}</svg>`;

    return await svgToPng(svg, width, height);
  } catch {
    return null;
  }
}
