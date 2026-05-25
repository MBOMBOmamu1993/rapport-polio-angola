import type { Metadata } from "next";
import "./globals.css";
import { Nav } from "@/components/Nav";

export const metadata: Metadata = {
  title: "Rapport — Campagne polio synchronisée avec l'Angola",
  description:
    "Importez le masque de saisie et téléchargez le rapport PowerPoint de la campagne de vaccination polio synchronisée avec l'Angola.",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="fr">
      <body>
        <Nav />
        <main className="mx-auto max-w-6xl px-4 py-6 md:px-6 md:py-8">{children}</main>
      </body>
    </html>
  );
}
