import type { Config } from "tailwindcss";

/**
 * Palette « navy » inspirée du modèle Power BI : bleu marine profond pour les
 * bandeaux et titres, accent bleu vif pour les actions. Conserve les nuances
 * de complétude (rouge / jaune / vert / bleu) du rapport officiel.
 */

const config: Config = {
  content: ["./app/**/*.{ts,tsx}", "./components/**/*.{ts,tsx}", "./lib/**/*.{ts,tsx}"],
  theme: {
    extend: {
      colors: {
        // Bleu navy / marine — palette principale.
        navy: {
          50: "#eef2f9",
          100: "#dce5f3",
          200: "#b3c4e3",
          300: "#869fd0",
          400: "#5a7abc",
          500: "#2e55a8",
          600: "#1f4490",
          700: "#1f3864", // navy header du modèle
          800: "#162a4d",
          900: "#0a1733",
          950: "#040b1c",
        },
        // Accent bleu vif (boutons primaires, liens).
        accent: {
          50: "#eff6ff",
          100: "#dbeafe",
          200: "#bfdbfe",
          400: "#60a5fa",
          500: "#2563eb",
          600: "#1d4ed8",
          700: "#1e40af",
        },
        // Sémantique complétude / couverture — seuils 80 / 95 / 100.
        threshold: {
          low: "#e23636",     // < 80 %  — rouge
          mid: "#f1c40f",     // 80–95 % — jaune
          high: "#22b457",    // 95–100 % — vert
          full: "#1d4ed8",    // > 100 % — bleu profond
          none: "#cbd5e1",    // pas de donnée
        },
        danger: { 50: "#fff1f1", 100: "#ffdede", 500: "#e23636", 600: "#c81e1e", 700: "#9b1616" },
        warn: { 50: "#fffbeb", 100: "#fef3c7", 500: "#f1c40f", 600: "#d4a40b" },
        good: { 50: "#effdf3", 100: "#d6f9e0", 500: "#22b457", 600: "#178a44" },
        surface: {
          0: "#ffffff",
          50: "#f6f8fb",
          100: "#eef2f7",
          200: "#dee5ee",
          300: "#c0cad7",
          400: "#94a3b8",
          500: "#64748b",
          700: "#334155",
          800: "#1e293b",
          900: "#0f172a",
        },
      },
      fontFamily: {
        sans: ["ui-sans-serif", "system-ui", "-apple-system", "Segoe UI", "Roboto", "Helvetica", "Arial", "sans-serif"],
      },
      boxShadow: {
        card: "0 1px 2px 0 rgba(15, 23, 42, 0.04), 0 4px 14px -4px rgba(15, 23, 42, 0.10)",
        navy: "0 8px 24px -12px rgba(31, 56, 100, 0.45)",
      },
      borderRadius: { xl: "0.9rem", "2xl": "1.1rem" },
      keyframes: {
        slidein: {
          "0%": { transform: "translateY(-12px)", opacity: "0" },
          "100%": { transform: "translateY(0)", opacity: "1" },
        },
        shimmer: {
          "0%": { backgroundPosition: "-200% 0" },
          "100%": { backgroundPosition: "200% 0" },
        },
      },
      animation: {
        slidein: "slidein 0.3s ease-out",
        shimmer: "shimmer 2.5s ease-in-out infinite",
      },
    },
  },
  plugins: [],
};

export default config;
