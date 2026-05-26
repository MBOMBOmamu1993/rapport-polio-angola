"use client";

import { useEffect, useRef, useState } from "react";
import Link from "next/link";
import Image from "next/image";
import { useApp } from "@/lib/store";
import { parseMasque } from "@/lib/parse-masque";
import { totals } from "@/lib/analytics";
import { fmtInt, fmtPct } from "@/lib/format";
import { fetchNational, pushImport, resetNational, type EntityInfo } from "@/lib/national";

export default function ImportPage() {
  const { data, setData, clearData } = useApp();
  const inputRef = useRef<HTMLInputElement>(null);
  const [toast, setToast] = useState<{ kind: "ok" | "err"; msg: string } | null>(null);
  const [busy, setBusy] = useState(false);
  const [dragOver, setDragOver] = useState(false);
  const [sync, setSync] = useState<"idle" | "syncing" | "ok" | "local">("idle");
  const [entities, setEntities] = useState<EntityInfo[] | null>(null);

  async function refreshEntities() {
    const nat = await fetchNational();
    setEntities(nat ? nat.entities : null);
  }
  useEffect(() => { void refreshEntities(); }, []);

  async function handleFile(file: File) {
    setBusy(true);
    setToast(null);
    try {
      const buf = await file.arrayBuffer();
      const parsed = parseMasque(buf, file.name);
      setData(parsed);
      setToast({ kind: "ok", msg: "Masque de saisie importé avec succès" });
      setSync("syncing");
      const r = await pushImport(parsed);
      setSync(r.ok ? "ok" : "local");
      if (r.ok) void refreshEntities();
    } catch (e) {
      setToast({ kind: "err", msg: e instanceof Error ? e.message : "Échec de l'import du fichier." });
    } finally {
      setBusy(false);
      setTimeout(() => setToast(null), 6000);
    }
  }

  function onPick(e: React.ChangeEvent<HTMLInputElement>) {
    const f = e.target.files?.[0];
    if (f) void handleFile(f);
    e.target.value = "";
  }

  const t = data ? totals(data.records) : null;

  return (
    <div className="space-y-6">
      {/* Toast */}
      {toast && (
        <div
          className={`animate-slidein fixed left-1/2 top-20 z-50 flex -translate-x-1/2 items-center gap-2 rounded-xl px-5 py-3 text-sm font-semibold text-white shadow-navy ${
            toast.kind === "ok" ? "bg-good-500" : "bg-danger-500"
          }`}
          role="status"
        >
          <span className="text-lg">{toast.kind === "ok" ? "✅" : "⚠️"}</span>
          {toast.msg}
        </div>
      )}

      {/* Hero navy */}
      <section className="overflow-hidden rounded-2xl navy-bar text-white shadow-navy">
        <div className="grid grid-cols-1 items-center gap-4 px-6 py-8 md:grid-cols-[auto_1fr] md:gap-6">
          <div className="flex items-center justify-center gap-3 md:justify-start">
            <div className="inline-flex items-center rounded-xl bg-white px-4 py-3 shadow-card">
              <Image src="/logo/pev-officiel.png" alt="PEV — Programme Élargi de Vaccination" width={326} height={50} className="h-11 w-auto md:h-14" priority />
            </div>
          </div>
          <div className="text-center md:text-left">
            <div className="mb-1 inline-block rounded-full bg-white/15 px-3 py-1 text-[11px] font-semibold uppercase tracking-wider text-accent-100">
              Étape 1 / 2
            </div>
            <h1 className="text-2xl font-bold md:text-3xl">Importer le masque de saisie</h1>
            <p className="mt-1 max-w-2xl text-sm text-accent-100/90">
              Sélectionnez le fichier Excel de votre province / antenne / zone de santé.
              Les analyses du rapport polio (nVPO2 et VPOb) seront disponibles instantanément.
              Réimporter le masque remplace automatiquement l&apos;ancienne version.
            </p>
          </div>
        </div>
      </section>

      {/* Dropzone */}
      <section
        onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
        onDragLeave={() => setDragOver(false)}
        onDrop={(e) => {
          e.preventDefault();
          setDragOver(false);
          const f = e.dataTransfer.files?.[0];
          if (f) void handleFile(f);
        }}
        className={`relative rounded-2xl border-2 border-dashed bg-white p-10 text-center transition shadow-card ${
          dragOver ? "border-navy-500 bg-navy-50" : "border-navy-200"
        }`}
      >
        <input ref={inputRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={onPick} />
        <div className="mx-auto mb-4 flex h-16 w-16 items-center justify-center rounded-2xl bg-navy-50 text-3xl">📄</div>
        <p className="mb-1 text-sm font-medium text-surface-700">
          Glissez-déposez le fichier Excel ici, ou
        </p>
        <button
          onClick={() => inputRef.current?.click()}
          disabled={busy}
          className="mt-3 inline-flex items-center gap-2 rounded-xl bg-navy-700 px-7 py-3.5 text-sm font-semibold text-white shadow-navy transition hover:bg-navy-800 disabled:opacity-50"
        >
          {busy ? (
            <>
              <span className="h-4 w-4 animate-spin rounded-full border-2 border-white border-t-transparent" />
              Lecture du fichier…
            </>
          ) : (
            <>📥 Choisir le masque de saisie</>
          )}
        </button>
        <p className="mt-4 text-[11px] text-surface-400">
          Formats acceptés : .xlsx, .xls — feuille « Synthèse » requise
        </p>
      </section>

      {/* État courant + aperçu instantané */}
      {data && t && (
        <section className="space-y-4">
          <div className="flex flex-wrap items-center justify-between gap-3 rounded-xl border border-good-100 bg-good-50 px-4 py-3">
            <div className="flex items-center gap-2 text-sm text-good-600">
              <span className="text-lg">✅</span>
              <span>
                <strong>{data.meta.fileName}</strong> — {fmtInt(data.meta.nbAires)} aires de santé ·{" "}
                {data.meta.province} · {data.meta.nbJours} jours saisis · importé le{" "}
                {new Date(data.meta.importedAt).toLocaleString("fr-FR")}
              </span>
            </div>
            <div className="flex items-center gap-2">
              <SyncBadge sync={sync} />
              <Link href="/rapport" className="rounded-lg bg-navy-700 px-4 py-2 text-xs font-semibold text-white shadow hover:bg-navy-800">
                📊 Aller au rapport
              </Link>
              <button onClick={() => { clearData(); setToast({ kind: "ok", msg: "Données effacées." }); }} className="rounded-lg border border-surface-300 px-4 py-2 text-xs font-medium text-surface-700 hover:bg-surface-100">
                Effacer
              </button>
            </div>
          </div>

          <h2 className="text-sm font-bold uppercase tracking-wide text-surface-500">
            Aperçu instantané des analyses
          </h2>
          <div className="grid grid-cols-2 gap-3 md:grid-cols-4">
            <Kpi label="Complétude rapports" value={fmtPct(t.completude)} icon="📋" tone={t.completude} />
            <Kpi label="Couverture nVPO2" value={fmtPct(t.nvpo2CV)} icon="💧" tone={t.nvpo2CV} />
            <Kpi label="Couverture VPOb" value={fmtPct(t.vpobCV)} icon="💧" tone={t.vpobCV} />
            <Kpi label="Récupérations PEV" value={fmtInt(t.recup)} icon="🔁" />
            <Kpi label="Vaccinés nVPO2" value={fmtInt(t.nvpo2Vacc)} icon="👶" />
            <Kpi label="Vaccinés VPOb" value={fmtInt(t.vpobVacc)} icon="👶" />
            <Kpi label="Taux perte nVPO2" value={fmtPct(t.nvpo2TauxPerte)} icon="🧪" />
            <Kpi label="MAPI graves" value={fmtInt(t.mapiGraves)} icon="⚕️" />
          </div>

          <div className="rounded-xl border border-surface-200 bg-white p-4 shadow-card">
            <h3 className="mb-3 text-sm font-semibold text-navy-700">Zones de Santé de votre import</h3>
            <div className="flex flex-wrap gap-2 text-xs">
              {data.meta.zones.map((z) => (
                <span key={z} className="rounded-full bg-navy-50 px-3 py-1 font-medium text-navy-700">{z}</span>
              ))}
            </div>
          </div>
        </section>
      )}

      <NationalPanel entities={entities} />

      {entities !== null && <AdminReset onDone={refreshEntities} />}
    </div>
  );
}

function AdminReset({ onDone }: { onDone: () => Promise<void> | void }) {
  const [open, setOpen] = useState(false);
  const [code, setCode] = useState("");
  const [confirm, setConfirm] = useState("");
  const [busy, setBusy] = useState(false);
  const [msg, setMsg] = useState<{ kind: "ok" | "err"; text: string } | null>(null);

  const phrase = "REINITIALISER";
  const ready = code.trim().length > 0 && confirm.trim().toUpperCase() === phrase;

  async function run() {
    if (!ready || busy) return;
    setBusy(true);
    setMsg(null);
    const r = await resetNational(code.trim());
    setBusy(false);
    if (r.ok) {
      setMsg({ kind: "ok", text: `Compilation nationale réinitialisée — ${r.deleted ?? 0} zone(s) de santé supprimée(s). Chaque province peut réimporter à zéro.` });
      setCode("");
      setConfirm("");
      await onDone();
    } else {
      const text =
        r.reason === "bad_code" ? "Code administrateur incorrect." :
        r.reason === "admin_not_configured" ? "Réinitialisation non configurée : définissez la variable d'environnement ADMIN_RESET_CODE chez l'hébergeur (Vercel), puis réessayez." :
        r.reason === "kv_unavailable" ? "Stockage national indisponible." :
        `Échec de la réinitialisation${r.reason ? ` (${r.reason})` : ""}.`;
      setMsg({ kind: "err", text });
    }
  }

  return (
    <section className="rounded-2xl border border-danger-100 bg-danger-50/40 p-4 shadow-card">
      <button
        type="button"
        onClick={() => setOpen((o) => !o)}
        className="flex w-full items-center justify-between text-left"
      >
        <span className="text-sm font-bold text-danger-600">🔒 Zone administrateur — réinitialiser la compilation nationale</span>
        <span className="text-surface-400">{open ? "▴" : "▾"}</span>
      </button>

      {open && (
        <div className="mt-4 space-y-3">
          <p className="text-xs text-surface-600">
            Cette action <strong>supprime définitivement</strong> tous les imports de toutes les provinces
            au niveau national. Chacun devra réimporter son masque à zéro. Les imports locaux sur les
            postes ne sont pas affectés.
          </p>
          <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
            <label className="block">
              <span className="mb-1 block text-[11px] font-semibold uppercase tracking-wide text-surface-400">Code administrateur</span>
              <input
                type="password"
                value={code}
                onChange={(e) => setCode(e.target.value)}
                placeholder="Code secret"
                className="w-full rounded-lg border border-surface-300 px-3 py-2.5 text-sm text-navy-700 focus:border-danger-400 focus:outline-none focus:ring-2 focus:ring-danger-100"
              />
            </label>
            <label className="block">
              <span className="mb-1 block text-[11px] font-semibold uppercase tracking-wide text-surface-400">Tapez « {phrase} » pour confirmer</span>
              <input
                value={confirm}
                onChange={(e) => setConfirm(e.target.value)}
                placeholder={phrase}
                className="w-full rounded-lg border border-surface-300 px-3 py-2.5 text-sm text-navy-700 focus:border-danger-400 focus:outline-none focus:ring-2 focus:ring-danger-100"
              />
            </label>
          </div>
          <button
            onClick={run}
            disabled={!ready || busy}
            className="inline-flex items-center gap-2 rounded-xl bg-danger-500 px-5 py-2.5 text-sm font-semibold text-white shadow-card transition hover:bg-danger-600 disabled:cursor-not-allowed disabled:opacity-40"
          >
            {busy ? (
              <>
                <span className="h-4 w-4 animate-spin rounded-full border-2 border-white border-t-transparent" />
                Réinitialisation…
              </>
            ) : (
              <>🗑️ Tout réinitialiser au niveau national</>
            )}
          </button>
          {msg && (
            <p className={`text-xs font-medium ${msg.kind === "ok" ? "text-good-600" : "text-danger-600"}`}>
              {msg.text}
            </p>
          )}
        </div>
      )}
    </section>
  );
}

function SyncBadge({ sync }: { sync: "idle" | "syncing" | "ok" | "local" }) {
  if (sync === "syncing")
    return <span className="rounded-lg bg-navy-50 px-3 py-2 text-xs font-medium text-navy-700">⏳ Synchronisation…</span>;
  if (sync === "ok")
    return <span className="rounded-lg bg-good-50 px-3 py-2 text-xs font-medium text-good-600">🌍 Synchronisé au niveau national</span>;
  if (sync === "local")
    return <span className="rounded-lg bg-warn-50 px-3 py-2 text-xs font-medium text-warn-600">💾 Enregistré localement</span>;
  return null;
}

function NationalPanel({ entities }: { entities: EntityInfo[] | null }) {
  if (entities === null) {
    return (
      <section className="rounded-xl border border-warn-100 bg-warn-50 p-4 text-sm text-warn-600">
        <strong>Compilation nationale non activée.</strong> Pour que les 5 provinces alimentent une vue pays
        commune, activez le stockage partagé (Vercel KV) — voir le README. En attendant, chaque import reste
        disponible localement.
      </section>
    );
  }
  if (entities.length === 0) {
    return (
      <section className="rounded-xl border border-surface-200 bg-white p-4 text-sm text-surface-500 shadow-card">
        🌍 Compilation nationale activée — aucune entité importée pour le moment.
      </section>
    );
  }
  const provinces = Array.from(new Set(entities.map((e) => e.province)));
  return (
    <section className="rounded-2xl border border-navy-100 bg-white p-4 shadow-card">
      <div className="mb-3 flex items-center justify-between">
        <h2 className="text-sm font-bold text-navy-700">🌍 Compilation nationale — entités déjà importées</h2>
        <span className="text-xs text-surface-500">{provinces.length} province(s) · {entities.length} ZS</span>
      </div>
      <div className="max-h-72 overflow-auto rounded-lg border border-surface-200">
        <table className="w-full text-left text-xs">
          <thead className="sticky top-0 bg-surface-100 text-surface-500">
            <tr>
              <th className="px-3 py-2 font-semibold">Province</th>
              <th className="px-3 py-2 font-semibold">Antenne</th>
              <th className="px-3 py-2 font-semibold">Zone de Santé</th>
              <th className="px-3 py-2 font-semibold">Aires</th>
              <th className="px-3 py-2 font-semibold">Dernière mise à jour</th>
            </tr>
          </thead>
          <tbody>
            {entities.map((e, i) => (
              <tr key={`${e.province}-${e.antenne}-${e.zs}-${i}`} className="border-t border-surface-100">
                <td className="px-3 py-1.5 font-medium text-navy-700">{e.province}</td>
                <td className="px-3 py-1.5 text-surface-700">{e.antenne}</td>
                <td className="px-3 py-1.5 text-surface-700">{e.zs}</td>
                <td className="px-3 py-1.5 text-surface-700">{e.nbAires}</td>
                <td className="px-3 py-1.5 text-surface-500">{new Date(e.importedAt).toLocaleString("fr-FR")}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  );
}

function Kpi({ label, value, icon, tone }: { label: string; value: string; icon: string; tone?: number | null }) {
  const color =
    tone == null ? "text-navy-700" :
    tone > 100 ? "text-accent-600" :
    tone >= 95 ? "text-good-600" :
    tone >= 80 ? "text-warn-600" : "text-danger-500";
  return (
    <div className="rounded-xl border border-surface-200 bg-white p-3 shadow-card">
      <div className="mb-1 flex items-center gap-1.5 text-[11px] uppercase tracking-wide text-surface-400">
        <span>{icon}</span>
        {label}
      </div>
      <div className={`text-2xl font-bold ${color}`}>{value}</div>
    </div>
  );
}
