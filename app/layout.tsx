import type { Metadata } from "next";
import "./globals.css";
import { Nav } from "@/components/Nav";

export const metadata: Metadata = {
  title: "Campagne intégrée RR-POLIO Kasaï Central",
  description:
    "Importez le masque de saisie intégré Rougeole-Rubéole / Polio du Kasaï Central et téléchargez le rapport PowerPoint des résultats partiels (données du masque + supervision ODK).",
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
