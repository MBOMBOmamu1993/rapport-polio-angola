/**
 * Rendu d'une carte de localisation des Zones de Santé de la RDC : les ZS
 * couvertes par le périmètre sélectionné sont surlignées (jaune) sur le fond
 * national, à la manière de la carte « locator » du modèle officiel.
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

const HILITE = "#F4E03A";   // zone(s) couverte(s)
const BASE_FILL = "#F1F5F9"; // reste du territoire
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
 * @param highlightZS ensemble de noms de ZS normalisés (via normalizeZS) à surligner.
 * @returns data URL PNG, ou null si la carte ne peut pas être produite.
 */
export async function renderZSMap(highlightZS: Set<string>): Promise<string | null> {
  try {
    const res = await fetch(TOPOJSON_URL, { cache: "force-cache" });
    if (!res.ok) return null;
    const topo = (await res.json()) as Topology;
    const objNames = Object.keys(topo.objects ?? {});
    if (objNames.length === 0) return null;
    const objName = objNames.find((n) => /zone/i.test(n)) ?? objNames[0];
    const fc = feature(topo, topo.objects[objName] as GeometryCollection) as FeatureCollection<Geometry>;
    if (!fc.features?.length) return null;

    const nameKey = detectNameKey(fc.features, highlightZS);

    const width = 1100;
    const height = 1000;
    const projection = geoMercator().fitSize([width, height], fc);
    const path = geoPath(projection);

    // On dessine d'abord les ZS de fond, puis les ZS surlignées par-dessus pour
    // que leur contour reste net.
    const base: string[] = [];
    const hi: string[] = [];
    for (const f of fc.features) {
      const d = path(f);
      if (!d) continue;
      const rawName = nameKey ? (f.properties?.[nameKey] as string | undefined) : undefined;
      const on = rawName ? highlightZS.has(normalizeZS(rawName)) : false;
      if (on) {
        hi.push(`<path d="${d}" fill="${HILITE}" stroke="${STROKE}" stroke-width="1.1"/>`);
      } else {
        base.push(`<path d="${d}" fill="${BASE_FILL}" stroke="${STROKE}" stroke-width="0.6"/>`);
      }
    }

    const svg =
      `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">` +
      `<rect width="${width}" height="${height}" fill="#FFFFFF"/>${base.join("")}${hi.join("")}</svg>`;

    return await svgToPng(svg, width, height);
  } catch {
    return null;
  }
}
