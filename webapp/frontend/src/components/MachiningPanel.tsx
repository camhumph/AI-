import { Fragment, useCallback, useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import {
  Clock, ChevronDown, ChevronRight, Gauge, RefreshCw,
  TriangleAlert, Ruler, DollarSign,
} from "lucide-react";
import { Card, Badge, EmptyState, Button } from "./ui";
import { useCountUp } from "../lib/useCountUp";
import { estimateFromStls } from "../lib/datumAdapter";
import {
  api,
  type AssetRef,
  type MachiningEstimate,
  type MachiningPart,
  type QuoteLineItem,
} from "../api/client";

/**
 * Machining time, shown as an argument rather than an answer.
 *
 * A foreman has to be able to push back on this, so the panel always exposes
 * three things: which operations came from measured geometry, which came from an
 * assumed hole pattern, and what share of the total each represents. A single
 * opaque hour figure would get quoted off and then be wrong.
 */
export default function MachiningPanel({
  jobId,
  models,
  quoteItems,
}: {
  jobId: string;
  /** Per-plate STLs from Module6121. When present, they are measured directly. */
  models?: AssetRef[];
  quoteItems?: QuoteLineItem[];
}) {
  const [data, setData] = useState<MachiningEstimate | null>(null);
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(true);
  const [open, setOpen] = useState<Record<string, boolean>>({});
  const [engine, setEngine] = useState<"mesh" | "pattern">("pattern");

  // QuoteDetailPage rebuilds `quoteItems` as a fresh array literal on every
  // render (it sits after an early return, so it cannot be memoized there
  // without breaking hook order). Depending on array IDENTITY would restart the
  // whole mesh analysis -- minutes of main-thread work across six STLs -- on any
  // parent re-render. Depend on a cheap content signature instead.
  const inputSig = useMemo(() => {
    const m = (models || []).map((x) => x.name).sort().join("|");
    const q = (quoteItems || [])
      .map((r) => `${r.role_label || r.component}:${r.thickness}x${r.width}x${r.length}:${r.material}`)
      .sort()
      .join("|");
    return `${m}##${q}`;
  }, [models, quoteItems]);

  // Latest values without making them dependencies.
  const latest = useRef({ models, quoteItems });
  latest.current = { models, quoteItems };

  const load = useCallback(() => {
    setLoading(true);
    setError("");

    // Prefer measuring the mesh. The Datum engine finds the actual holes,
    // pockets and counterbores in the six per-plate STLs, so nothing has to be
    // guessed from a per-role pattern table. Fall back to the server's coarse
    // estimate only when the STLs are missing or unreadable.
    const { models: mdls, quoteItems: items } = latest.current;
    const stlCount = (mdls || []).filter((m) => m.name.toLowerCase().endsWith(".stl")).length;

    const serverEstimate = () =>
      api.machining(jobId).then((d) => {
        setEngine("pattern");
        setData(d);
      });

    if (stlCount > 0 && items && items.length > 0) {
      estimateFromStls(mdls || [], items)
        .then((d) => {
          if (d.summary.part_count > 0) {
            setEngine("mesh");
            setData(d);
            return;
          }
          return serverEstimate();
        })
        .catch(() => serverEstimate())
        .catch((e) => setError(e instanceof Error ? e.message : "Failed to estimate."))
        .finally(() => setLoading(false));
      return;
    }

    serverEstimate()
      .catch((e) => setError(e instanceof Error ? e.message : "Failed to estimate."))
      .finally(() => setLoading(false));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [jobId, inputSig]);

  useEffect(load, [load]);

  const priced = useMemo(() => (data?.parts || []).filter((p) => !p.skipped), [data]);
  const skipped = useMemo(() => (data?.parts || []).filter((p) => p.skipped), [data]);

  // Longest part drives the bar scale, so the chart is comparative not absolute.
  const maxMin = useMemo(
    () => priced.reduce((m, p) => Math.max(m, p.total_min), 0) || 1,
    [priced],
  );

  if (loading) {
    return (
      <div className="space-y-4">
        <div className="grid grid-cols-2 gap-4 lg:grid-cols-4">
          {[0, 1, 2, 3].map((i) => (
            <div key={i} className="skeleton h-28" />
          ))}
        </div>
        <div className="skeleton h-64" />
      </div>
    );
  }

  if (error) {
    return (
      <EmptyState
        icon={<TriangleAlert className="h-8 w-8" />}
        title="Could not estimate"
        description={error}
        action={<Button onClick={load}><RefreshCw className="h-4 w-4" /> Retry</Button>}
      />
    );
  }

  if (!data || priced.length === 0) {
    return (
      <EmptyState
        icon={<Clock className="h-8 w-8" />}
        title="Nothing to estimate yet"
        description="Machining time is derived from the quoted steel rows. Once the quote sheet has plates with stock dimensions, the estimate appears here."
      />
    );
  }

  const s = data.summary;
  const confTone = s.confidence_pct >= 70 ? "success" : s.confidence_pct >= 50 ? "warning" : "danger";
  const confFill =
    s.confidence_pct >= 70 ? "meter-fill-green" : s.confidence_pct >= 50 ? "meter-fill-amber" : "meter-fill-rose";

  // Job-level confidence band. Present only on the mesh path.
  const band = s.confidence;
  const gradeTone =
    band?.grade === "A" ? "success" : band?.grade === "B" ? "brand" : band?.grade === "C" ? "warning" : "danger";

  return (
    <div className="space-y-5">
      {/* ---- headline numbers ------------------------------------------ */}
      <div className="grid grid-cols-2 gap-4 lg:grid-cols-4">
        <Kpi
          className="rise-in rise-in-1"
          icon={<Clock className="h-4 w-4" />}
          label="Total time"
          amount={s.total_hours}
          format={(v) => fmt(v, 1)}
          unit="hrs"
          hint={
            band
              ? `${fmt((s.low_minutes ?? 0) / 60, 1)}–${fmt((s.high_minutes ?? 0) / 60, 1)} hrs across ` +
                `${s.part_count} part${s.part_count === 1 ? "" : "s"}`
              : `${fmt(s.total_minutes, 0)} min across ${s.part_count} part${s.part_count === 1 ? "" : "s"}`
          }
          rail="rail-brand"
        />
        <Kpi
          className="rise-in rise-in-2"
          icon={<DollarSign className="h-4 w-4" />}
          label={engine === "mesh" ? "Full cost" : "Machine + setup"}
          amount={s.estimated_cost}
          format={(v) => `$${fmt(v, 0)}`}
          hint={
            engine === "mesh"
              ? "incl. stock, programming, inspection, tooling"
              : `hours x $${fmt(s.shop_rate_per_hour, 0)}/hr only`
          }
          rail="rail-green"
        />
        <Kpi
          className="rise-in rise-in-3"
          icon={<Ruler className="h-4 w-4" />}
          label={engine === "mesh" ? "Measured in mesh" : "From geometry"}
          amount={s.measured_minutes / 60}
          format={(v) => fmt(v, 1)}
          unit="hrs"
          hint={engine === "mesh" ? "found holes, pockets, faces" : "measured volume + area"}
        />
        <Kpi
          className="rise-in rise-in-4"
          icon={<Gauge className="h-4 w-4" />}
          label="Assumed"
          amount={s.assumed_minutes / 60}
          format={(v) => fmt(v, 1)}
          unit="hrs"
          hint={engine === "mesh" ? "setup + fixturing only" : "holes, taps, setup"}
          rail="rail-amber"
        />
      </div>

      {/* ---- quote band: the number to actually quote off ---------------- */}
      {band && (
        <Card className="rise-in rise-in-5 p-5">
          <div className="mb-4 flex flex-wrap items-start justify-between gap-3">
            <div>
              <div className="flex items-center gap-2">
                <span className="section-label">Quote range</span>
                <Badge tone={gradeTone}>
                  Grade {band.grade} · {band.score}/100
                </Badge>
              </div>
              <p className="mt-1.5 max-w-2xl text-xs leading-relaxed text-ink-400">
                {band.guidance}
              </p>
            </div>
            <div className="text-right">
              <div className="num-money text-2xl text-ink-100">
                ${fmt(s.low_cost ?? 0, 0)} – ${fmt(s.high_cost ?? 0, 0)}
              </div>
              <div className="mt-0.5 text-[11px] text-ink-400">
                point estimate ${fmt(s.estimated_cost, 0)}
              </div>
            </div>
          </div>

          {/* Band bar. The point estimate sits where it falls inside the range,
              which is deliberately NOT the middle -- the high side opens wider
              because unplanned work gets discovered, not removed. */}
          <div className="relative h-9">
            <div className="absolute inset-x-0 top-3 h-3 overflow-hidden rounded-full bg-ink-800">
              <div
                className="h-full rounded-full bg-gradient-to-r from-accent-green/50 via-brand-500/50 to-accent-amber/60"
                style={{ width: "100%" }}
              />
            </div>
            <div
              className="absolute top-1 h-7 w-0.5 rounded bg-ink-100"
              style={{
                left: `${clamp(
                  (100 * (1 - band.lowFactor)) / Math.max(band.highFactor - band.lowFactor, 0.001),
                )}%`,
              }}
              title={`Point estimate ${fmt(s.total_hours, 1)} hrs`}
            />
          </div>
          <div className="flex justify-between text-[11px] text-ink-400">
            <span>
              {fmt((s.low_minutes ?? 0) / 60, 1)} hrs
              <span className="ml-1 text-ink-500">
                (x{fmt(band.lowFactor, 2)})
              </span>
            </span>
            <span className="text-ink-300">{fmt(s.total_hours, 1)} hrs estimated</span>
            <span>
              {fmt((s.high_minutes ?? 0) / 60, 1)} hrs
              <span className="ml-1 text-ink-500">(x{fmt(band.highFactor, 2)})</span>
            </span>
          </div>

          {band.drivers.length > 0 && (
            <div className="mt-4 border-t border-white/10 pt-3">
              <div className="section-label mb-2">What widened this range</div>
              <ul className="space-y-1.5">
                {band.drivers.map((d, i) => (
                  <li key={i} className="flex gap-2 text-[11px] leading-relaxed text-ink-300">
                    <span className="mt-1.5 h-1 w-1 shrink-0 rounded-full bg-accent-amber" />
                    <span>{d}</span>
                  </li>
                ))}
              </ul>
              <p className="mt-3 text-[11px] leading-relaxed text-ink-500">
                Each line is a named assumption, not a fudge factor. Resolve the ones that
                matter and the range narrows — that is the whole point of showing them.
              </p>
            </div>
          )}
        </Card>
      )}

      {/* ---- how much of this is real ----------------------------------- */}
      <Card className="rise-in rise-in-5 p-5">
        <div className="mb-3 flex flex-wrap items-center justify-between gap-3">
          <div>
            <div className="flex items-center gap-2">
              <span className="section-label">Confidence</span>
              <Badge tone={engine === "mesh" ? "brand" : "neutral"}>
                {engine === "mesh" ? "measured from mesh" : "pattern estimate"}
              </Badge>
            </div>
            <p className="mt-1.5 max-w-2xl text-xs leading-relaxed text-ink-400">
              {engine === "mesh" ? (
                <>
                  Holes, pockets, counterbores and chamfers were <strong className="text-ink-200">found
                  in the STL</strong> — diameters fitted from the facets, depths measured, and
                  through-vs-blind resolved by ray cast. Only setup and fixturing are assumed.
                </>
              ) : (
                <>
                  No readable plate STLs, so this is the server's coarse estimate: volume
                  and area are measured, but hole and tap counts come from a per-role
                  pattern table. Re-run Module6121 to export the six plate STLs and this
                  becomes a measured number.
                </>
              )}
            </p>
          </div>
          <Badge tone={confTone}>{fmt(s.confidence_pct, 0)}% measured</Badge>
        </div>
        <div className="meter">
          <div className={`meter-fill ${confFill}`} style={{ width: `${clamp(s.confidence_pct)}%` }} />
        </div>
        {s.confidence_pct < 60 && (
          <p className="mt-3 text-[11px] leading-relaxed text-accent-amber">
            Under 60% measured — treat this as a rough order of magnitude, not a
            quotable figure, until the hole counts are checked against the prints.
          </p>
        )}
      </Card>

      {/* ---- per part ---------------------------------------------------- */}
      <Card className="rise-in rise-in-6 overflow-hidden">
        <div className="flex items-center justify-between border-b border-white/10 bg-ink-800/60 px-4 py-3">
          <div className="flex items-center gap-2">
            <span className="text-sm font-semibold text-ink-100">By part</span>
            <Badge>{priced.length}</Badge>
            {skipped.length > 0 && (
              <Badge tone="danger">{skipped.length} not quoted</Badge>
            )}
          </div>
          <button
            onClick={load}
            className="flex items-center gap-1.5 rounded-full border border-white/10 px-3 py-1.5 text-[11px] font-semibold uppercase tracking-wider text-ink-400 transition hover:border-white/20 hover:text-ink-200"
          >
            <RefreshCw className="h-3.5 w-3.5" /> Recalculate
          </button>
        </div>

        <div className="scrollbar-thin overflow-x-auto">
          <table className="data-table text-sm">
            <thead>
              <tr>
                <th>Part</th>
                <th>Material</th>
                <th className="text-right">Qty</th>
                <th className="text-right">Stock in³</th>
                <th className="text-right">Removed</th>
                <th>Share of total</th>
                <th className="text-right">Hours</th>
                <th className="w-8" />
              </tr>
            </thead>
            <tbody>
              {/* Refused plates render IN the table, not in a footnote below it.
                  A plate that silently vanishes reads as "missing" and sends the
                  reader hunting; a red row with a reason tells them what to fix. */}
              {skipped.map((p, i) => (
                <tr key={`skip-${i}`} className="bg-accent-rose/5">
                  <td className="max-w-[18rem]">
                    <div className="truncate font-medium text-ink-200" title={p.component}>
                      {p.role_label || p.component || "--"}
                    </div>
                    <div className="truncate font-mono text-[10px] text-ink-500" title={p.component}>
                      {p.component}
                    </div>
                  </td>
                  <td className="text-xs text-ink-400">{p.material}</td>
                  <td className="num text-right text-ink-400">{p.qty}</td>
                  <td colSpan={4} className="text-xs leading-relaxed text-accent-rose">
                    NOT QUOTED — {p.reason}
                  </td>
                  <td />
                </tr>
              ))}
              {priced.map((p, i) => {
                const key = `${p.component}-${i}`;
                const isOpen = open[key];
                return (
                  <Fragment key={key}>
                    <tr
                      className="cursor-pointer"
                      onClick={() => setOpen((o) => ({ ...o, [key]: !o[key] }))}
                    >
                      <td className="max-w-[18rem]">
                        <div className="truncate font-medium text-ink-100" title={p.component}>
                          {p.role_label || p.component || "--"}
                        </div>
                        {p.role_label && p.component && p.role_label !== p.component && (
                          <div className="truncate font-mono text-[10px] text-ink-500" title={p.component}>
                            {p.component}
                          </div>
                        )}
                      </td>
                      <td className="text-xs text-ink-300">{p.material}</td>
                      <td className="num text-right text-ink-200">{p.qty}</td>
                      <td className="num text-right text-xs text-ink-400">
                        {fmt(p.stock_volume_cuin ?? 0, 0)}
                      </td>
                      <td className="num text-right text-xs">
                        <span className="text-ink-200">{fmt(p.removed_pct ?? 0, 0)}%</span>
                        <span className="ml-1 text-ink-500">
                          ({fmt(p.removed_cuin ?? 0, 0)} in³)
                        </span>
                      </td>
                      <td className="min-w-[7rem]">
                        <div className="meter">
                          <div
                            className="meter-fill"
                            style={{ width: `${clamp((p.total_min / maxMin) * 100)}%` }}
                          />
                        </div>
                      </td>
                      <td className="num-money text-right text-ink-100">{fmt(p.total_hours, 2)}</td>
                      <td className="text-ink-500">
                        {isOpen ? <ChevronDown className="h-4 w-4" /> : <ChevronRight className="h-4 w-4" />}
                      </td>
                    </tr>

                    {isOpen && (
                      <tr>
                        <td colSpan={8} className="bg-ink-900/50 px-4 py-4">
                          <OpBreakdown part={p} />
                        </td>
                      </tr>
                    )}
                  </Fragment>
                );
              })}
            </tbody>
            <tfoot>
              <tr className="border-t border-white/15 bg-ink-850/50">
                <td colSpan={6} className="px-4 py-2.5 text-xs font-semibold uppercase tracking-wider text-ink-300">
                  Total
                </td>
                <td className="num-money px-4 py-2.5 text-right text-base text-ink-100">
                  {fmt(s.total_hours, 2)}
                </td>
                <td />
              </tr>
            </tfoot>
          </table>
        </div>
      </Card>


      <Card className="p-4">
        <div className="section-label mb-2">How this is calculated</div>
        <ul className="space-y-1.5">
          {data.notes.map((n, i) => (
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

function OpBreakdown({ part }: { part: MachiningPart }) {
  const ops = part.operations || [];
  const total = ops.reduce((sum, o) => sum + o.min, 0) || 1;

  return (
    <div className="space-y-4">
      <div className="grid gap-4 sm:grid-cols-3">
        <Fact label="Stock" value={
          part.stock
            ? `${fmt(part.stock.thickness, 3)} × ${fmt(part.stock.width, 3)} × ${fmt(part.stock.length, 3)} in`
            : "--"
        } />
        <Fact
          label="Finished volume"
          value={`${fmt(part.finished_volume_cuin ?? 0, 1)} in³`}
          hint={part.finished_basis}
        />
        <Fact
          label="Cut + setup"
          value={`${fmt(part.cut_min ?? 0, 0)} + ${fmt(
            Math.max(0, (part.per_part_min ?? 0) - (part.cut_min ?? 0)),
            0,
          )} min`}
          hint={
            // Only shown when a real factor was applied. The mesh path leaves it
            // undefined on purpose -- Datum declares a duty factor but plan()
            // never reads it, so printing "efficiency 0%" was worse than silence.
            part.efficiency_factor
              ? `efficiency ${fmt(part.efficiency_factor * 100, 0)}%`
              : "raw cutting time — add your own shop factor"
          }
        />
      </div>

      <div className="space-y-1.5">
        {ops.map((o, i) => (
          <div key={i} className="flex items-center gap-3">
            <div className="w-32 shrink-0 text-xs font-medium text-ink-200">{o.op}</div>
            <div className="flex-1">
              <div className="meter">
                <div
                  className={`meter-fill ${o.measured ? "" : "meter-fill-amber"}`}
                  style={{ width: `${clamp((o.min / total) * 100)}%` }}
                />
              </div>
            </div>
            <div className="num w-16 shrink-0 text-right text-xs text-ink-200">
              {fmt(o.min, 1)} m
            </div>
            <div className="hidden w-64 shrink-0 text-[10px] text-ink-500 lg:block">{o.detail}</div>
            <Badge tone={o.measured ? "success" : "warning"}>
              {o.measured ? "measured" : "assumed"}
            </Badge>
          </div>
        ))}
      </div>

      <p className="text-[10px] leading-relaxed text-ink-500">
        Amber bars are pattern estimates. They scale with the part's role, not its
        actual print, so they are the first thing to correct if the total looks off.
      </p>
    </div>
  );
}

function Fact({ label, value, hint }: { label: string; value: string; hint?: string }) {
  return (
    <div>
      <div className="section-label">{label}</div>
      <div className="num mt-1 text-sm text-ink-100">{value}</div>
      {hint && <div className="mt-0.5 text-[10px] text-ink-500">{hint}</div>}
    </div>
  );
}

/**
 * A headline figure that counts up to its value.
 *
 * The count-up writes straight to the DOM node via a ref, so a 620 ms animation
 * costs one paint per frame instead of ~37 React re-renders.
 */
function Kpi({
  icon, label, amount, format, unit, hint, rail, className,
}: {
  icon: ReactNode;
  label: string;
  amount: number;
  format: (v: number) => string;
  unit?: string;
  hint?: string;
  rail?: string;
  className?: string;
}) {
  const ref = useCountUp(amount, format);
  return (
    <div
      className={`glass-panel card-lift relative overflow-hidden rounded-2xl p-4 ${rail || ""} ${className || ""}`}
    >
      <div className="flex items-center gap-2 text-ink-500">
        {icon}
        <span className="section-label">{label}</span>
      </div>
      <div className="mt-2 flex items-baseline gap-1.5">
        <span ref={ref} className="num-lg">{format(amount)}</span>
        {unit && <span className="text-xs font-medium text-ink-500">{unit}</span>}
      </div>
      {hint && <div className="mt-1 truncate text-[10px] text-ink-500">{hint}</div>}
    </div>
  );
}

function fmt(v: number | null | undefined, digits = 1): string {
  const n = Number(v);
  if (!Number.isFinite(n)) return "--";
  return n.toLocaleString(undefined, {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
  });
}

function clamp(pct: number): number {
  if (!Number.isFinite(pct)) return 0;
  return Math.max(0, Math.min(100, pct));
}
