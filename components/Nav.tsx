"use client";

import Link from "next/link";
import Image from "next/image";
import { usePathname } from "next/navigation";
import { clsx } from "clsx";

const links = [
  { href: "/import", label: "Importer le masque de saisie", icon: "📥" },
  { href: "/rapport", label: "Télécharger le rapport", icon: "📊" },
];

export function Nav() {
  const pathname = usePathname();
  return (
    <header className="sticky top-0 z-30 border-b border-surface-200 bg-white/90 backdrop-blur">
      <div className="mx-auto flex max-w-6xl items-center gap-3 px-4 py-2.5 md:px-6">
        <Image src="/logo/pev.png" alt="PEV" width={40} height={40} className="h-10 w-10 object-contain" priority />
        <div className="mr-auto leading-tight">
          <div className="text-sm font-bold text-oms-800 md:text-base">
            Campagne polio synchronisée avec l&apos;Angola
          </div>
          <div className="text-[11px] text-surface-500">Programme Élargi de Vaccination — RD Congo</div>
        </div>
        <nav className="flex gap-1.5">
          {links.map((l) => {
            const active = pathname === l.href || (pathname === "/" && l.href === "/import");
            return (
              <Link
                key={l.href}
                href={l.href}
                className={clsx(
                  "rounded-lg px-3 py-2 text-xs font-semibold transition md:text-sm",
                  active ? "bg-oms-500 text-white" : "text-oms-700 hover:bg-oms-50"
                )}
              >
                <span className="mr-1">{l.icon}</span>
                <span className="hidden sm:inline">{l.label}</span>
              </Link>
            );
          })}
        </nav>
      </div>
    </header>
  );
}
