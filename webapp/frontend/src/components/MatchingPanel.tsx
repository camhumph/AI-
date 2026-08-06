import { Suspense, useEffect, useMemo, useState } from "react";
import { Boxes } from "lucide-react";
import StlViewer, { type StlInfo } from "./StlViewer";
import StlOverlayViewer from "./StlOverlayViewer";
import { Card, Spinner, EmptyState, Badge } from "./ui";
import {
  classifyComponentFilename,
  classifyQuoteComponentName,
  KIND_ORDER,
  type ComponentKind,
} from "../lib/componentKind";
import { matchPercent, matchTone, type GeomStats } from "../lib/stlGeometry";
import type { AssetRef, QuoteLineItem } from "../api/client";

interface Props {
  models: AssetRef[];
  quoteItems: QuoteLineItem[];
}

const PRICED_PART_KINDS: ComponentKind[] = ["TCP", "BCP", "ID HOLDER", "OD HOLDER", "ID POT", "OD POT"];

export default function MatchingPanel({ models, quoteItems }: Props) {
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

  const kindsWithModels = KIND_ORDER.filter((k) => (grouped.get(k)?.length ?? 0) > 0);

  const [activeKind, setActiveKind] = useState<ComponentKind | null>(kindsWithModels[0] ?? null);
  useEffect(() => {
    setActiveKind((prev) => (prev && (grouped.get(prev)?.length ?? 0) > 0 ? prev : kindsWithModels[0] ?? null));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [models]);

  const activeFiles = activeKind ? grouped.get(activeKind) || [] : [];

  const [leftUrl, setLeftUrl] = useState("");
  const [rightUrl, setRightUrl] = useState("");
  const [leftInfo, setLeftInfo] = useState<StlInfo | null>(null);
  const [rightInfo, setRightInfo] = useState<StlInfo | null>(null);
  const [leftGeom, setLeftGeom] = useState<GeomStats | null>(null);
  const [rightGeom, setRightGeom] = useState<GeomStats | null>(null);

  useEffect(() => {
    setLeftUrl(activeFiles[0]?.url || "");
    setRightUrl(activeFiles[1]?.url || activeFiles[0]?.url || "");
    setLeftInfo(null);
    setRightInfo(null);
    setLeftGeom(null);
    setRightGeom(null);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeKind]);

  const match = leftGeom && rightGeom ? matchPercent(leftGeom, rightGeom) : null;

  // Priced-in-quote vs. has-a-3D-file, per part — the comparison we can make
  // reliably from data this job already has.
  const coverage = useMemo(() => {
    const priced = new Set<ComponentKind>();
    for (const it of quoteItems) {
      const kind = classifyQuoteComponentName(it.component || it.role_label || "");
      if (kind !== "OTHER") priced.add(kind);
    }
    return PRICED_PART_KINDS.map((k) => ({
      kind: k,
      priced: priced.has(k),
      modeled: (grouped.get(k)?.length ?? 0) > 0,
    })).filter((r) => r.priced || r.modeled);
  }, [quoteItems, grouped]);

  if (models.length === 0) {
    return (
      <EmptyState
        icon={<Boxes className="h-8 w-8" />}
        title="No 3D files to compare"
        description="Upload STL files in the 3D tab, then come back here to compare parts side by side."
      />
    );
  }

  return (
    <div className="space-y-5">
      {coverage.length > 0 && (
        <Card className="p-5">
          <div className="section-label mb-3">Coverage — priced vs. modeled</div>
          <div className="overflow-x-auto">
            <table className="w-full text-left text-sm">
              <thead>
                <tr className="text-[11px] uppercase tracking-wider text-ink-400">
                  <th className="py-2 pr-4 font-medium">Part</th>
                  <th className="py-2 pr-4 font-medium">In Quote</th>
                  <th className="py-2 font-medium">3D File</th>
                </tr>
              </thead>
              <tbody>
                {coverage.map((row) => (
                  <tr key={row.kind} className="border-t border-ink-800/60">
                    <td className="py-2 pr-4 text-ink-100">{row.kind}</td>
                    <td className="py-2 pr-4">
                      <Badge tone={row.priced ? "success" : "danger"}>{row.priced ? "Yes" : "Missing"}</Badge>
                    </td>
                    <td className="py-2">
                      <Badge tone={row.modeled ? "success" : "danger"}>{row.modeled ? "Yes" : "Missing"}</Badge>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Card>
      )}

      <div className="flex flex-wrap gap-2">
        {kindsWithModels.map((k) => (
          <button
            key={k}
            type="button"
            onClick={() => setActiveKind(k)}
            className={`flex items-center gap-2 rounded-full border px-3.5 py-1.5 text-[11px] font-semibold uppercase tracking-wider transition ${
              activeKind === k
                ? "border-white/25 bg-white/10 text-ink-100"
                : "border-white/10 text-ink-400 hover:border-white/20 hover:text-ink-200"
            }`}
          >
            {k}
            <span className="rounded-full bg-white/10 px-1.5 py-0.5 text-[10px] text-ink-300">
              {grouped.get(k)?.length ?? 0}
            </span>
          </button>
        ))}
      </div>

      {activeKind && (
        <>
          <div className="grid grid-cols-1 gap-4 lg:grid-cols-2">
            {(
              [
                { label: "A", url: leftUrl, setUrl: setLeftUrl, info: leftInfo, setInfo: setLeftInfo },
                { label: "B", url: rightUrl, setUrl: setRightUrl, info: rightInfo, setInfo: setRightInfo },
              ] as const
            ).map((pane) => (
              <Card key={pane.label} className="p-4">
                <div className="mb-3 flex items-center justify-between gap-3">
                  <div className="section-label">File {pane.label}</div>
                  <select
                    value={pane.url}
                    onChange={(e) => pane.setUrl(e.target.value)}
                    className="glass-input max-w-[60%] rounded-full px-3 py-1.5 text-xs text-ink-200"
                  >
                    {activeFiles.map((f) => (
                      <option key={f.url} value={f.url}>
                        {f.name}
                      </option>
                    ))}
                  </select>
                </div>
                <div className="h-[380px]">
                  {pane.url ? (
                    <Suspense fallback={<Spinner label="Loading 3D..." />}>
                      <StlViewer key={pane.url} url={pane.url} onInfo={pane.setInfo} />
                    </Suspense>
                  ) : (
                    <div className="flex h-full items-center justify-center text-xs text-ink-500">
                      No file selected
                    </div>
                  )}
                </div>
                {pane.info && (
                  <div className="mt-2 text-[11px] text-ink-500">
                    {pane.info.x.toFixed(2)} × {pane.info.y.toFixed(2)} × {pane.info.z.toFixed(2)} in ·{" "}
                    {pane.info.tris.toLocaleString()} tris
                  </div>
                )}
              </Card>
            ))}
          </div>

          {leftUrl && rightUrl && (
            <Card className="p-4">
              <div className="mb-3 flex flex-wrap items-center justify-between gap-3">
                <div className="section-label">Overlay</div>
                <div className="flex items-center gap-4 text-[10px] uppercase tracking-wider text-ink-400">
                  <span className="flex items-center gap-1.5">
                    <span className="h-2.5 w-2.5 rounded-full" style={{ background: "#3b82f6" }} /> File A
                  </span>
                  <span className="flex items-center gap-1.5">
                    <span className="h-2.5 w-2.5 rounded-full" style={{ background: "#ff8c42" }} /> File B
                  </span>
                </div>
              </div>
              <div className="h-[420px]">
                <Suspense fallback={<Spinner label="Loading overlay..." />}>
                  <StlOverlayViewer
                    key={`${leftUrl}::${rightUrl}`}
                    urlA={leftUrl}
                    urlB={rightUrl}
                    onStatsA={setLeftGeom}
                    onStatsB={setRightGeom}
                  />
                </Suspense>
              </div>
              {match && (
                <div className="mt-4 flex flex-wrap items-center gap-4 border-t border-ink-800/60 pt-4">
                  <div className="text-3xl font-light tracking-tight">
                    <span
                      className={
                        matchTone(match.pct) === "success"
                          ? "text-accent-green"
                          : matchTone(match.pct) === "warning"
                          ? "text-accent-amber"
                          : "text-accent-rose"
                      }
                    >
                      {match.pct.toFixed(1)}%
                    </span>
                  </div>
                  <div className="text-xs text-ink-400">
                    geometry match
                    <div className="mt-1 text-ink-500">
                      size Δ {match.sizeDiffPct.toFixed(1)}% · mass Δ {match.massDiffPct.toFixed(1)}% · COM dist{" "}
                      {match.comDistIn.toFixed(3)} in
                    </div>
                  </div>
                </div>
              )}
            </Card>
          )}
        </>
      )}
    </div>
  );
}
