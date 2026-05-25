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
    <header className="sticky top-0 z-30 navy-bar shadow-navy">
      <div className="mx-auto flex max-w-6xl items-center gap-3 px-4 py-3 md:px-6">
        <Image
          src="/logo/pev-logo-white.svg"
          alt="PEV — Programme Élargi de Vaccination — RD Congo"
          width={300}
          height={64}
          className="h-9 w-auto shrink-0 md:h-11"
          priority
        />
        <div className="mr-auto leading-tight">
          <div className="text-sm font-bold text-white md:text-base">
            Campagne polio synchronisée avec l&apos;Angola
          </div>
          <div className="text-[11px] text-accent-100/80 md:text-xs">
            nVPO2 &amp; VPOb (co-administration)
          </div>
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
                  active
                    ? "bg-white text-navy-700 shadow-md"
                    : "text-white/85 hover:bg-white/10 hover:text-white"
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
