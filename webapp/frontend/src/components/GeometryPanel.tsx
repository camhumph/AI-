import { useCallback, useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import {
  Circle, Square, CircleDot, Download, RefreshCw, Ruler, Type, TriangleAlert,
} from "lucide-react";
import { Card, Badge, Button, EmptyState } from "./ui";
import { geometryFromStls } from "../lib/datumAdapter";
import type { AssetRef, GeometryPart, GeometryReport, QuoteLineItem } from "../api/client";

/**
 * Every feature on every plate, with sizes.
 *
 * Deliberately separate from the Machining tab. That one answers "how long",
 * which involves calibration, derating and assumption. This one answers "what is
 * actually on the part" -- nothing here is adjusted, so it is the page to check
 * the model against, and the page to hand someone who wants to work out the time
 * themselves.
 *
 * Exports to CSV because that is what actually gets used: pasted into a
 * spreadsheet next to a real routing sheet.
 */
export default function GeometryPanel({
  models,
  quoteItems,
}: {
  models?: AssetRef[];
  quoteItems?: QuoteLineItem[];
}) {
  const [data, setData] = useState<GeometryReport | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  const inputSig = useMemo(() => {
    const m = (models || []).map((x) => x.name).sort().join("|");
    const q = (quoteItems || [])
      .map((r) => `${r.role_label || r.component}:${r.thickness}x${r.width}x${r.length}`)
      .sort()
      .join("|");
    return `${m}##${q}`;
  }, [models, quoteItems]);

  const latest = useRef({ models, quoteItems });
  latest.current = { models, quoteItems };

  const load = useCallback(() => {
    setLoading(true);
    setError("");
    const { models: mdls, quoteItems: items } = latest.current;
    geometryFromStls(mdls || [], items || [])
      .then(setData)
      .catch((e) => setError(e instanceof Error ? e.message : "Analysis failed."))
      .finally(() => setLoading(false));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [inputSig]);

  useEffect(load, [load]);

  if (loading) {
    return (
      <div className="space-y-4">
        <div className="grid grid-cols-2 gap-4 lg:grid-cols-5">
          {[0, 1, 2, 3, 4].map((i) => <div key={i} className="skeleton h-24" />)}
        </div>
        <div className="skeleton h-96" />
      </div>
    );
  }

  if (error) {
    return (
      <EmptyState
        icon={<TriangleAlert className="h-8 w-8" />}
        title="Could not analyse geometry"
        description={error}
        action={<Button onClick={load}><RefreshCw className="h-4 w-4" /> Retry</Button>}
      />
    );
  }

  const live = (data?.parts || []).filter((p) => !p.skipped);
  if (live.length === 0) {
    return (
      <EmptyState
        icon={<Ruler className="h-8 w-8" />}
        title="No readable plate STLs"
        description="Geometry is measured from the six per-plate STLs Module6121 exports. Re-run the macro so the stl folder is populated, then reload."
      />
    );
  }

  const t = data!.totals;

  return (
    <div className="space-y-5">
      <div className="grid grid-cols-2 gap-4 lg:grid-cols-5">
        <Stat className="rise-in rise-in-1" icon={<Circle className="h-4 w-4" />} label="Holes" value={t.holes} rail="rail-brand" />
        <Stat className="rise-in rise-in-2" icon={<CircleDot className="h-4 w-4" />} label="Tapped" value={t.tapped} rail="rail-green" />
        <Stat className="rise-in rise-in-3" icon={<Square className="h-4 w-4" />} label="Pockets" value={t.pockets} />
        <Stat className="rise-in rise-in-4" icon={<CircleDot className="h-4 w-4" />} label="Bores" value={t.bores} />
        <Stat className="rise-in rise-in-5" icon={<Ruler className="h-4 w-4" />} label="Chamfer" value={t.chamfer_len_in} unit="in" />
      </div>

      <div className="flex justify-end">
        <button
          onClick={() => downloadCsv(data!)}
          className="flex items-center gap-1.5 rounded-full border border-white/10 px-3.5 py-1.5 text-[11px] font-semibold uppercase tracking-wider text-ink-400 transition hover:border-white/20 hover:text-ink-200"
        >
          <Download className="h-3.5 w-3.5" /> Export CSV
        </button>
      </div>

      {data!.parts.map((p, i) => <PlateCard key={i} part={p} index={i} />)}

      <Card className="p-4">
        <div className="section-label mb-2">Reading this page</div>
        <ul className="space-y-1.5">
          {data!.notes.map((n, i) => (
            <li key={i} className="flex gap-2 text-xs leading-relaxed text-ink-400">
              <span className="mt-1.5 h-1 w-1 shrink-0 rounded-full bg-ink-600" />
              {n}
            </li>
          ))}
        </ul>
      </Card>
    </div>
  );
}

function PlateCard({ part, index }: { part: GeometryPart; index: number }) {
  if (part.skipped) {
    return (
      <Card className="p-4">
        <div className="flex items-center gap-2">
          <span className="text-sm font-semibold text-ink-100">{part.role_label}</span>
          <Badge tone="warning">not analysed</Badge>
        </div>
        <p className="mt-1.5 text-xs text-ink-400">{part.reason}</p>
      </Card>
    );
  }

  return (
    <Card className={`overflow-hidden rise-in rise-in-${Math.min(6, index + 1)}`}>
      <div className="flex flex-wrap items-center justify-between gap-3 border-b border-white/10 bg-ink-800/60 px-4 py-3">
        <div className="min-w-0">
          <div className="flex items-center gap-2">
            <span className="text-sm font-semibold uppercase tracking-wider text-ink-100">
              {part.role_label}
            </span>
            <Badge>{part.material}</Badge>
            {part.orientations > 1 && <Badge tone="brand">{part.orientations} orientations</Badge>}
          </div>
          <div className="truncate font-mono text-[10px] text-ink-500">{part.component}</div>
        </div>
        <div className="num text-[11px] text-ink-400">
          {part.stock_in && (
            <>
              stock {fmt(part.stock_in.thickness, 3)} × {fmt(part.stock_in.width, 3)} ×{" "}
              {fmt(part.stock_in.length, 3)} in ·{" "}
            </>
          )}
          finished {fmt(part.finished_volume_cuin, 1)} in³
        </div>
      </div>

      <div className="space-y-5 p-4">
        {part.engraving && part.engraving.length > 0 && (
          <Section icon={<Type className="h-3.5 w-3.5" />} title="Engraving">
            <div className="rounded-xl border border-white/10 bg-ink-900/50 p-3 font-mono text-xs leading-relaxed text-ink-200">
              {part.engraving.map((line, i) => <div key={i}>{line}</div>)}
            </div>
          </Section>
        )}

        {part.holes.length > 0 && (
          <Section icon={<Circle className="h-3.5 w-3.5" />} title={`Holes (${part.holes.reduce((s, h) => s + h.count, 0)})`}>
            <table className="data-table text-sm">
              <thead>
                <tr>
                  <th className="text-right">Ø in</th>
                  <th className="text-right">Depth</th>
                  <th className="text-right">Qty</th>
                  <th className="text-right">L/D</th>
                  <th>End</th>
                  <th>Purpose</th>
                  <th>Thread</th>
                  <th>Tool</th>
                </tr>
              </thead>
              <tbody>
                {part.holes.map((h, i) => (
                  <tr key={i}>
                    <td className="num text-right font-medium text-ink-100">{fmt(h.dia_in, 4)}</td>
                    <td className="num text-right text-ink-300">{fmt(h.depth_in, 3)}</td>
                    <td className="num text-right text-ink-200">{h.count}</td>
                    <td className="num text-right">
                      <span className={h.ld > 8 ? "text-accent-rose" : h.ld > 5 ? "text-accent-amber" : "text-ink-400"}>
                        {fmt(h.ld, 1)}
                      </span>
                    </td>
                    <td className="text-xs text-ink-400">{h.blind ? "blind" : "through"}</td>
                    <td className="text-xs text-ink-400">{h.purpose}</td>
                    <td className="text-xs">
                      {h.thread ? <span className="text-accent-green">{h.thread}</span> : <span className="text-ink-600">—</span>}
                    </td>
                    <td className="text-[10px] text-ink-500">
                      {/* Amber means "read this", not "we cannot do it". A reach
                          limit needs a peck strategy or a gun drill on any
                          machine. A tool that simply lives on another machine is
                          ordinary — shown plainly, priced normally. */}
                      {h.tool.includes("reach") ? (
                        <span className="text-accent-amber">{h.tool}</span>
                      ) : h.tool.includes("another machine") ? (
                        <span className="text-ink-400">{h.tool}</span>
                      ) : h.tool}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </Section>
        )}

        {part.taps.length > 0 && (
          <Section icon={<CircleDot className="h-3.5 w-3.5" />} title="Threads">
            <div className="flex flex-wrap gap-2">
              {part.taps.map((tp, i) => (
                <div
                  key={i}
                  className={`rounded-xl border px-3 py-2 ${
                    tp.tool.startsWith("NOT IN CRIB")
                      ? "border-accent-amber/40 bg-accent-amber/10"
                      : "border-white/10 bg-white/5"
                  }`}
                >
                  <div className="text-xs font-semibold text-ink-100">
                    {tp.spec} <span className="num text-ink-400">×{tp.count}</span>
                  </div>
                  <div className="mt-0.5 text-[10px] text-ink-500">{tp.tool}</div>
                </div>
              ))}
            </div>
          </Section>
        )}

        {part.bores.length > 0 && (
          <Section icon={<CircleDot className="h-3.5 w-3.5" />} title="Bores">
            <table className="data-table text-sm">
              <thead><tr><th className="text-right">Ø in</th><th className="text-right">Qty</th><th>Method</th></tr></thead>
              <tbody>
                {part.bores.map((b, i) => (
                  <tr key={i}>
                    <td className="num text-right font-medium text-ink-100">{fmt(b.dia_in, 4)}</td>
                    <td className="num text-right text-ink-200">{b.count}</td>
                    <td className="text-xs text-ink-400">{b.method}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </Section>
        )}

        {part.pockets.length > 0 && (
          <Section icon={<Square className="h-3.5 w-3.5" />} title={`Pockets (${part.pockets.reduce((s, k) => s + k.count, 0)})`}>
            <table className="data-table text-sm">
              <thead><tr><th className="text-right">Depth in</th><th className="text-right">Area in²</th><th className="text-right">Qty</th><th>Type</th></tr></thead>
              <tbody>
                {part.pockets.map((k, i) => (
                  <tr key={i}>
                    <td className="num text-right font-medium text-ink-100">{fmt(k.depth_in, 3)}</td>
                    <td className="num text-right text-ink-300">{fmt(k.area_sqin, 2)}</td>
                    <td className="num text-right text-ink-200">{k.count}</td>
                    <td className="text-xs text-ink-400">{k.cbore ? "counterbore seat" : "pocket"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </Section>
        )}

        <div className="flex flex-wrap gap-x-6 gap-y-1 border-t border-white/5 pt-3 text-[11px] text-ink-500">
          <span>Chamfer: <span className="num text-ink-300">{fmt(part.chamfer_len_in, 1)} in</span> over {part.chamfer_count} band(s)</span>
          <span>Radii / blends: <span className="num text-ink-300">{part.radii_count}</span></span>
          <span>Distinct tools: <span className="num text-ink-300">{part.distinct_tools}</span></span>
        </div>
      </div>
    </Card>
  );
}

function Section({ icon, title, children }: { icon: ReactNode; title: string; children: ReactNode }) {
  return (
    <div>
      <div className="mb-2 flex items-center gap-2 text-ink-500">
        {icon}
        <span className="section-label">{title}</span>
      </div>
      <div className="scrollbar-thin overflow-x-auto">{children}</div>
    </div>
  );
}

function Stat({
  icon, label, value, unit, rail, className,
}: {
  icon: ReactNode; label: string; value: number; unit?: string; rail?: string; className?: string;
}) {
  return (
    <div className={`glass-panel card-lift rounded-2xl p-4 ${rail || ""} ${className || ""}`}>
      <div className="flex items-center gap-2 text-ink-500">{icon}<span className="section-label">{label}</span></div>
      <div className="mt-2 flex items-baseline gap-1.5">
        <span className="num-lg">{value.toLocaleString()}</span>
        {unit && <span className="text-xs text-ink-500">{unit}</span>}
      </div>
    </div>
  );
}

/** Flat CSV, one row per feature, so it pastes straight into a routing sheet. */
function downloadCsv(r: GeometryReport) {
  const rows: string[][] = [
    ["Plate", "Feature", "Dia_in", "Depth_in", "Area_sqin", "Qty", "LD", "End", "Purpose", "Thread", "Tool"],
  ];
  for (const p of r.parts) {
    if (p.skipped) {
      rows.push([p.role_label, "NOT ANALYSED", "", "", "", "", "", "", p.reason || "", "", ""]);
      continue;
    }
    for (const h of p.holes) {
      rows.push([p.role_label, "Hole", String(h.dia_in), String(h.depth_in), "", String(h.count),
        String(h.ld), h.blind ? "blind" : "through", h.purpose, h.thread || "", h.tool]);
    }
    for (const b of p.bores) {
      rows.push([p.role_label, "Bore", String(b.dia_in), "", "", String(b.count), "", "", b.method, "", ""]);
    }
    for (const k of p.pockets) {
      rows.push([p.role_label, k.cbore ? "Counterbore" : "Pocket", "", String(k.depth_in),
        String(k.area_sqin), String(k.count), "", "", "", "", ""]);
    }
    for (const tp of p.taps) {
      rows.push([p.role_label, "Thread", "", "", "", String(tp.count), "", "", "", tp.spec, tp.tool]);
    }
    rows.push([p.role_label, "Chamfer", "", "", "", String(p.chamfer_count), "", "", `${p.chamfer_len_in} in of edge`, "", ""]);
    for (const line of p.engraving || []) {
      rows.push([p.role_label, "Engraving", "", "", "", "", "", "", line, "", ""]);
    }
  }

  const csv = rows
    .map((r2) => r2.map((c) => (/[",\n]/.test(c) ? `"${c.replace(/"/g, '""')}"` : c)).join(","))
    .join("\r\n");

  const url = URL.createObjectURL(new Blob([csv], { type: "text/csv;charset=utf-8" }));
  const a = document.createElement("a");
  a.href = url;
  a.download = "geometry.csv";
  a.click();
  URL.revokeObjectURL(url);
}

function fmt(v: number | null | undefined, digits = 2): string {
  const n = Number(v);
  if (!Number.isFinite(n)) return "--";
  return n.toLocaleString(undefined, { minimumFractionDigits: digits, maximumFractionDigits: digits });
}
