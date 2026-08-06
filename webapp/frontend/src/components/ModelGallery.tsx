import { Suspense, useEffect, useMemo, useState } from "react";
import { Box, Download } from "lucide-react";
import StlViewer, { type StlInfo } from "./StlViewer";
import { Card, Spinner, Badge } from "./ui";
import { classifyComponentFilename, KIND_ORDER, type ComponentKind } from "../lib/componentKind";
import type { AssetRef } from "../api/client";

interface Props {
  models: AssetRef[];
}

const ALL = "__ALL__" as const;
type KindFilter = ComponentKind | typeof ALL;

/**
 * Per-part model browser, laid out like the shop's Match Studio: component
 * pills across the top, the matching files down the left, and the selected
 * plate in a large viewer with its live bounding box and triangle count.
 *
 * Module6121 writes one STL per quoted plate into <job>\stl (see
 * ExportPlateStlsForComparison) and the backend hoists those into models/,
 * so on a standard base this fills up with A Plate, B Plate, Rails 1..n,
 * the ejector stack, and so on -- all already rotated into the same
 * corrected CMS Top/Front frame, so switching between them holds orientation.
 */
export default function ModelGallery({ models }: Props) {
  const grouped = useMemo(() => {
    const map = new Map<ComponentKind, AssetRef[]>();
    for (const m of models) {
      const kind = classifyComponentFilename(m.name);
      const list = map.get(kind) || [];
      list.push(m);
      map.set(kind, list);
    }
    return map;
  }, [models]);

  const kindsWithModels = useMemo(
    () => KIND_ORDER.filter((k) => (grouped.get(k)?.length ?? 0) > 0),
    [grouped],
  );

  // Start on "All" -- on a full standard base, seeing every plate at once is
  // the useful first view, and the pills are right there to narrow it down.
  const [filter, setFilter] = useState<KindFilter>(ALL);
  const [activeUrl, setActiveUrl] = useState("");
  const [info, setInfo] = useState<StlInfo | null>(null);

  const visible = useMemo(
    () => (filter === ALL ? kindsWithModels.flatMap((k) => grouped.get(k) || []) : grouped.get(filter) || []),
    [filter, grouped, kindsWithModels],
  );

  // Keep the filter and selection valid as the job reloads (an upload or a
  // re-run of the macro can change the model list under us).
  useEffect(() => {
    setFilter((prev) =>
      prev === ALL || (grouped.get(prev)?.length ?? 0) > 0 ? prev : kindsWithModels[0] ?? ALL,
    );
  }, [grouped, kindsWithModels]);

  useEffect(() => {
    setActiveUrl((prev) => (visible.some((m) => m.url === prev) ? prev : visible[0]?.url || ""));
  }, [visible]);

  useEffect(() => {
    setInfo(null);
  }, [activeUrl]);

  const active = visible.find((m) => m.url === activeUrl) || null;
  const activeKind = active ? classifyComponentFilename(active.name) : null;

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap gap-2">
        <Pill
          label="All parts"
          count={models.length}
          selected={filter === ALL}
          onClick={() => setFilter(ALL)}
        />
        {kindsWithModels.map((k) => (
          <Pill
            key={k}
            label={k}
            count={grouped.get(k)?.length ?? 0}
            selected={filter === k}
            onClick={() => setFilter(k)}
          />
        ))}
      </div>

      <div className="grid grid-cols-1 gap-4 lg:grid-cols-3">
        <Card className="p-3 lg:col-span-1">
          <div className="section-label mb-2 px-1">
            {visible.length} {visible.length === 1 ? "file" : "files"}
          </div>
          <div className="max-h-[520px] space-y-1 overflow-y-auto pr-1">
            {visible.map((m) => {
              const isActive = m.url === activeUrl;
              return (
                <button
                  key={m.url}
                  type="button"
                  onClick={() => setActiveUrl(m.url)}
                  className={`flex w-full items-center gap-2.5 rounded-xl border px-3 py-2 text-left transition ${
                    isActive
                      ? "border-white/25 bg-white/10"
                      : "border-transparent hover:border-white/10 hover:bg-white/5"
                  }`}
                >
                  <Box
                    className={`h-4 w-4 shrink-0 ${isActive ? "text-brand-400" : "text-ink-500"}`}
                  />
                  <span className="min-w-0 flex-1">
                    <span
                      className={`block truncate text-xs font-medium ${
                        isActive ? "text-ink-100" : "text-ink-300"
                      }`}
                    >
                      {prettyModelName(m.name)}
                    </span>
                    <span className="block text-[10px] text-ink-500">{formatKb(m.size)}</span>
                  </span>
                </button>
              );
            })}
          </div>
        </Card>

        <Card className="p-4 lg:col-span-2">
          <div className="mb-3 flex flex-wrap items-center justify-between gap-2">
            <div className="min-w-0">
              <div className="truncate text-sm font-semibold text-ink-100">
                {active ? prettyModelName(active.name) : "No model selected"}
              </div>
              <div className="mt-1 flex items-center gap-2">
                {activeKind && <Badge>{activeKind}</Badge>}
                {info && (
                  <span className="text-[11px] text-ink-500">
                    {info.x.toFixed(2)} × {info.y.toFixed(2)} × {info.z.toFixed(2)} in ·{" "}
                    {info.tris.toLocaleString()} tris
                  </span>
                )}
              </div>
            </div>
            {active && (
              <a
                href={active.url}
                download={active.name}
                className="flex items-center gap-1.5 rounded-full border border-white/10 px-3 py-1.5 text-[11px] font-semibold uppercase tracking-wider text-ink-400 transition hover:border-white/20 hover:text-ink-200"
              >
                <Download className="h-3.5 w-3.5" /> STL
              </a>
            )}
          </div>

          <div className="h-[520px]">
            {activeUrl ? (
              <Suspense fallback={<Spinner label="Loading 3D..." />}>
                <StlViewer key={activeUrl} url={activeUrl} onInfo={setInfo} />
              </Suspense>
            ) : (
              <div className="flex h-full items-center justify-center text-xs text-ink-500">
                No file selected
              </div>
            )}
          </div>
        </Card>
      </div>
    </div>
  );
}

function Pill({
  label,
  count,
  selected,
  onClick,
}: {
  label: string;
  count: number;
  selected: boolean;
  onClick: () => void;
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      className={`flex items-center gap-2 rounded-full border px-3.5 py-1.5 text-[11px] font-semibold uppercase tracking-wider transition ${
        selected
          ? "border-white/25 bg-white/10 text-ink-100"
          : "border-white/10 text-ink-400 hover:border-white/20 hover:text-ink-200"
      }`}
    >
      {label}
      <span className="rounded-full bg-white/10 px-1.5 py-0.5 text-[10px] text-ink-300">{count}</span>
    </button>
  );
}

/**
 * "J8420_A Plate.STL" -> "A Plate". Module6121 prefixes per-plate STLs with
 * the job base name (StlPlateFilePrefix) so the files stay self-describing
 * once PublishJobOutputs flattens them onto the matching share; in the UI the
 * job is already known, so the prefix is just noise.
 *
 * Split on the LAST underscore, not the first: a job base name is often the
 * customer folder leaf and carries its own underscores
 * ("856800001_88_MB_03-05-2026_A Plate"), while the plate label never does --
 * CleanFileName keeps the spaces in "Bottom Ejector Plate". If the trailing
 * segment has no letters it isn't a label (e.g. the full-assembly STL, whose
 * name is just the base name ending in a date), so show the whole stem.
 */
function prettyModelName(name: string) {
  const stem = (name || "").replace(/\.[^.]+$/, "");
  const cut = stem.lastIndexOf("_");
  const tail = cut > 0 ? stem.slice(cut + 1).trim() : "";
  if (tail && /[A-Za-z]/.test(tail)) return tail;
  return stem.replace(/_/g, " ").trim() || name;
}

function formatKb(size: number) {
  if (!size) return "";
  if (size < 1024) return `${size} B`;
  if (size < 1024 * 1024) return `${Math.round(size / 1024)} KB`;
  return `${(size / (1024 * 1024)).toFixed(1)} MB`;
}
