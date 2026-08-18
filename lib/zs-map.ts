/**
 * Cartes des Zones de Santé (fond DHIS2 de la RDC, TopoJSON) rendues côté
 * navigateur : SVG (d3-geo) → PNG (data URL) embarqué dans le PowerPoint.
 *
 * Deux usages :
 *  - carte de la province (Kasaï Central) avec les ZS colorées selon un indicateur
 *    (complétude, % d'AS visitées…) ;
 *  - carte des points GPS des supervisions par‑dessus les contours des ZS.
 * En cas d'échec (réseau, format), on renvoie null et la carte est omise.
 */

import { geoMercator, geoPath } from "d3-geo";
import { feature } from "topojson-client";
import type { Topology, GeometryCollection } from "topojson-specification";
import type { Feature, FeatureCollection, Geometry } from "geojson";

const LOCAL_URL = "/geo/rdc_zs.topojson";
const REMOTE_URL =
  "https://gist.githubusercontent.com/MBOMBOmamu1993/1297c206c046ee018d5ed6c392d6c20f/raw/24ce95b2935d2b4cc4ef71701138218ca870ff01/rdc_zs.topojson";

export const MAP_RED = "#EA7C77";
export const MAP_ORANGE = "#F5B77A";
export const MAP_GREEN = "#91BE8A";
export const MAP_BASE = "#F1F5F9";
const STROKE = "#7A8699";

export function normalizeZS(name: string): string {
  return (name || "")
    .toString()
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .toUpperCase()
    .replace(/ZONE DE SANTE/g, "")
    .replace(/PROVINCE/g, "")
    .replace(/[^A-Z0-9]/g, "");
}

/** Nom DHIS2 « kr Bena Leka Zone de Santé » → « BENALEKA ». */
function dhisName(name: string): string {
  const cleaned = (name || "").replace(/^[a-z]{2}\s+/i, "");
  return normalizeZS(cleaned);
}

let topoCache: Promise<FeatureCollection<Geometry> | null> | null = null;

async function loadFeatures(): Promise<FeatureCollection<Geometry> | null> {
  if (topoCache) return topoCache;
  topoCache = (async () => {
    for (const url of [LOCAL_URL, REMOTE_URL]) {
      try {
        const res = await fetch(url, { cache: "force-cache" });
        if (!res.ok) continue;
        const topo = (await res.json()) as Topology;
        const objNames = Object.keys(topo.objects ?? {});
        if (objNames.length === 0) continue;
        const objName = objNames.find((n) => /zone/i.test(n)) ?? objNames[0];
        const fc = feature(topo, topo.objects[objName] as GeometryCollection) as FeatureCollection<Geometry>;
        if (fc.features?.length) return fc;
      } catch {
        /* essai suivant */
      }
    }
    return null;
  })();
  return topoCache;
}

function provinceFeatures(fc: FeatureCollection<Geometry>, province: string): Feature<Geometry>[] {
  const target = normalizeZS(province);
  return fc.features.filter((f) => {
    // « kr Kasai Central Province » → « KASAICENTRAL » (préfixe DHIS2 de 2 lettres retiré).
    const p = normalizeZS(String(f.properties?.parentName ?? "").replace(/^[a-z]{2}\s+/i, ""));
    return p === target;
  });
}

function levenshtein(a: string, b: string): number {
  const m = a.length, n = b.length;
  if (!m) return n;
  if (!n) return m;
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

/** Retrouve l'index de la feature correspondant à un nom de ZS (tolérant). */
function matchFeature(target: string, featNames: string[]): number {
  const exact = featNames.indexOf(target);
  if (exact >= 0) return exact;
  let best = -1;
  let bestD = Infinity;
  for (let i = 0; i < featNames.length; i++) {
    const fn = featNames[i];
    if (!fn) continue;
    if (fn.includes(target) || target.includes(fn)) {
      const d = Math.abs(fn.length - target.length);
      if (d < bestD) { bestD = d; best = i; }
      continue;
    }
    const d = levenshtein(target, fn);
    if (d < bestD) { bestD = d; best = i; }
  }
  const thr = Math.max(2, Math.floor(target.length * 0.25));
  return best >= 0 && bestD <= thr ? best : -1;
}

function svgToPng(svg: string, width: number, height: number): Promise<string | null> {
  return new Promise((resolve) => {
    if (typeof Image === "undefined" || typeof document === "undefined") return resolve(null);
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

function esc(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

export interface ProvinceMapOptions {
  /** Province (nom du masque, ex. « KASAI-CENTRAL »). */
  province: string;
  /** Couleur de remplissage par ZS (clé = normalizeZS(nom)). Absente → gris clair. */
  fillByZS?: Map<string, string>;
  /** Points GPS à dessiner (supervisions). */
  points?: { lat: number; lon: number; color?: string }[];
  /** Afficher le nom des ZS. */
  labels?: boolean;
  /** Fond satellite‑like sombre (pour la carte des points). */
  dark?: boolean;
  width?: number;
  height?: number;
  /** Renvoie le SVG (tests) au lieu du PNG. */
  svgOnly?: boolean;
}

/**
 * Carte de la province : ZS colorées + points optionnels. Renvoie une data URL PNG
 * (ou le SVG si `svgOnly`), null si le fond n'est pas disponible.
 */
export async function renderProvinceMap(o: ProvinceMapOptions): Promise<string | null> {
  try {
    const fc = await loadFeatures();
    if (!fc) return null;
    let feats = provinceFeatures(fc, o.province);
    if (feats.length === 0) feats = fc.features;
    const width = o.width ?? 1000;
    const height = o.height ?? 1000;
    const sub: FeatureCollection<Geometry> = { type: "FeatureCollection", features: feats };
    const projection = geoMercator().fitExtent([[20, 20], [width - 20, height - 20]], sub);
    const path = geoPath(projection);

    const names = feats.map((f) => dhisName(String(f.properties?.name ?? "")));
    const fills = new Map<number, string>();
    if (o.fillByZS) {
      for (const [k, color] of o.fillByZS) {
        const idx = matchFeature(k, names);
        if (idx >= 0) fills.set(idx, color);
      }
    }
    const bg = o.dark ? "#2F3E2C" : "#FFFFFF";
    const baseFill = o.dark ? "#3E5039" : MAP_BASE;
    const stroke = o.dark ? "#FFFFFF" : STROKE;
    const parts: string[] = [`<rect width="${width}" height="${height}" fill="${bg}"/>`];
    feats.forEach((f, i) => {
      const d = path(f);
      if (!d) return;
      parts.push(`<path d="${d}" fill="${fills.get(i) ?? baseFill}" stroke="${stroke}" stroke-width="${o.dark ? 1.4 : 1.6}" stroke-linejoin="round"/>`);
    });
    if (o.labels) {
      feats.forEach((f) => {
        const c = path.centroid(f);
        if (!c || !Number.isFinite(c[0])) return;
        const raw = String(f.properties?.name ?? "").replace(/^[a-z]{2}\s+/i, "").replace(/zone de sant[eé]/i, "").trim();
        parts.push(
          `<text x="${c[0]}" y="${c[1]}" text-anchor="middle" font-family="Calibri, Arial, sans-serif" font-size="17" font-weight="600" fill="${o.dark ? "#FFFFFF" : "#1E293B"}" stroke="${o.dark ? "#000000" : "#FFFFFF"}" stroke-width="3" paint-order="stroke" stroke-linejoin="round">${esc(raw)}</text>`
        );
      });
    }
    if (o.points?.length) {
      for (const p of o.points) {
        const xy = projection([p.lon, p.lat]);
        if (!xy || !Number.isFinite(xy[0]) || !Number.isFinite(xy[1])) continue;
        parts.push(`<circle cx="${xy[0].toFixed(1)}" cy="${xy[1].toFixed(1)}" r="7" fill="${p.color ?? "#F28E2A"}" fill-opacity="0.85" stroke="#FFFFFF" stroke-width="1.2"/>`);
      }
    }
    const svg =
      `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">` +
      parts.join("") +
      `</svg>`;
    if (o.svgOnly) return svg;
    return await svgToPng(svg, width, height);
  } catch {
    return null;
  }
}

/** Couleur d'un pourcentage (< 50 rouge, 50‑80 orange, ≥ 80 vert). */
export function colorFor3(v: number | null | undefined): string {
  if (v == null || !Number.isFinite(v)) return MAP_BASE;
  if (v >= 80) return MAP_GREEN;
  if (v >= 50) return MAP_ORANGE;
  return MAP_RED;
}
