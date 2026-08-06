import { useEffect, useMemo, useState } from "react";
import { Link } from "react-router-dom";
import { Scale, RefreshCw, Search, ChevronDown, ChevronUp } from "lucide-react";
import { Card, Button, EmptyState, Badge } from "./ui";
import { classifyComponentFilename, type ComponentKind } from "../lib/componentKind";
import {
  computeGeomStats,
  matchTone,
  densityFor,
  estimateMassLb,
  loadStlGeometry,
  type GeomStats,
} from "../lib/stlGeometry";
import {
  SIGNATURE_KINDS,
  compareComponentSignature,
  compareJobSignatures,
  buildJobSignature,
  type JobSignature,
  type JobMatchResult,
} from "../lib/jobSignature";
import { api, type AssetRef, type QuoteLineItem } from "../api/client";

interface Props {
  jobId: string;
  displayName: string;
  models: AssetRef[];
  quoteItems: QuoteLineItem[];
}

interface Analyzed {
  model: AssetRef;
  kind: ComponentKind;
  stats: GeomStats;
  massLb: number;
  error?: string;
}

function isStl(name: string): boolean {
  return name.toLowerCase().endsWith(".stl");
}

const MATCH_THRESHOLD = 80;
const SCOPE_OPTIONS = [
  { label: "Most recent 10", value: 10 },
  { label: "Most recent 25", value: 25 },
  { label: "Most recent 50", value: 50 },
  { label: "All quotes", value: 0 },
];

export default function ComparisonPanel({ jobId, displayName, models, quoteItems }: Props) {
  const stlModels = useMemo(() => models.filter((m) => isStl(m.name)), [models]);

  // ---- Within-job analysis (every STL file in THIS job, compared to every other) ----
  const [analyzed, setAnalyzed] = useState<Analyzed[]>([]);
  const [loading, setLoading] = useState(false);
  const [progress, setProgress] = useState({ done: 0, total: 0 });

  const materialByKind = useMemo(() => {
    const map = new Map<ComponentKind, string>();
    for (const it of quoteItems) {
      if (!it.material) continue;
      const kind = classifyComponentFilename(it.component || it.role_label || "");
      if (kind !== "OTHER" && !map.has(kind)) map.set(kind, it.material);
    }
    return map;
  }, [quoteItems]);

  const analyze = async () => {
    setLoading(true);
    setProgress({ done: 0, total: stlModels.length });
    const results: Analyzed[] = [];
    for (const model of stlModels) {
      const kind = classifyComponentFilename(model.name);
      try {
        const geometry = await loadStlGeometry(model.url);
        const stats = computeGeomStats(geometry);
        const density = densityFor(materialByKind.get(kind));
        results.push({ model, kind, stats, massLb: estimateMassLb(stats.vol, density) });
      } catch (e) {
        results.push({
          model,
          kind,
          stats: { sx: 0, sy: 0, sz: 0, bcx: 0, bcy: 0, bcz: 0, comx: 0, comy: 0, comz: 0, vol: 0, area: 0, tris: 0 },
          massLb: 0,
          error: e instanceof Error ? e.message : "Failed to load",
        });
      }
      setProgress((p) => ({ ...p, done: p.done + 1 }));
    }
    setAnalyzed(results);
    setLoading(false);
  };

  useEffect(() => {
    if (stlModels.length > 0) analyze();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [stlModels.length]);

  const pairs = useMemo(() => {
    const ok = analyzed.filter((a) => !a.error);
    const rows: { a: Analyzed; b: Analyzed; match: ReturnType<typeof compareComponentSignature> }[] = [];
    for (let i = 0; i < ok.length; i++) {
      for (let j = i + 1; j < ok.length; j++) {
        rows.push({ a: ok[i], b: ok[j], match: compareComponentSignature(ok[i].stats, ok[j].stats) });
      }
    }
    return rows.sort((r1, r2) => r2.match.similarity - r1.match.similarity);
  }, [analyzed]);

  const likelyMatches = pairs.filter((p) => p.match.similarity >= MATCH_THRESHOLD);

  // ---- Cross-job "Similar Quotes" (checks this quote's signature against every other quote) ----
  const currentSignature: JobSignature = useMemo(() => {
    const components = new Map<ComponentKind, { model: { name: string; url: string }; stats: GeomStats }>();
    for (const kind of SIGNATURE_KINDS) {
      const row = analyzed.find((a) => a.kind === kind && !a.error);
      if (row) components.set(kind, { model: { name: row.model.name, url: row.model.url }, stats: row.stats });
    }
    return { jobId, displayName, components };
  }, [analyzed, jobId, displayName]);

  const [scope, setScope] = useState(25);
  const [jobMatches, setJobMatches] = useState<JobMatchResult[] | null>(null);
  const [matchLoading, setMatchLoading] = useState(false);
  const [matchProgress, setMatchProgress] = useState({ done: 0, total: 0 });
  const [matchError, setMatchError] = useState("");
  const [expanded, setExpanded] = useState<Set<string>>(new Set());

  const findSimilarQuotes = async () => {
    if (currentSignature.components.size === 0) {
      setMatchError("This quote has no TCP/BCP/Holder/Pot STL files to compare yet.");
      return;
    }
    setMatchError("");
    setMatchLoading(true);
    setJobMatches(null);

    try {
      const allJobs = await api.listJobs();
      let candidates = allJobs.filter((j) => j.job_id !== jobId && j.model_count > 0);
      candidates = candidates
        .slice()
        .sort((a, b) => new Date(b.updated_at).getTime() - new Date(a.updated_at).getTime());
      if (scope > 0) candidates = candidates.slice(0, scope);

      setMatchProgress({ done: 0, total: candidates.length });
      const results: JobMatchResult[] = [];

      for (const summary of candidates) {
        try {
          const detail = await api.getJob(summary.job_id);
          const candidateSignature = await buildJobSignature(summary.job_id, summary.display_name, detail.models);
          const result = compareJobSignatures(currentSignature, candidateSignature);
          if (result) results.push(result);
        } catch {
          // skip jobs that fail to load
        }
        setMatchProgress((p) => ({ ...p, done: p.done + 1 }));
      }

      results.sort((a, b) => b.score - a.score);
      setJobMatches(results.slice(0, 10));
    } catch (e) {
      setMatchError(e instanceof Error ? e.message : "Failed to compare quotes.");
    } finally {
      setMatchLoading(false);
    }
  };

  const toggleExpanded = (id: string) => {
    setExpanded((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  if (stlModels.length === 0) {
    return (
      <EmptyState
        icon={<Scale className="h-8 w-8" />}
        title="No STL files to analyze"
        description="Upload STL files in the 3D tab, then come back here to compare mass, size, and center of mass."
      />
    );
  }

  return (
    <div className="space-y-8">
      <div className="space-y-5">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <div>
            <div className="section-label">Within this quote</div>
            <p className="mt-1 max-w-2xl text-xs text-ink-500">
              Volume and center of mass are computed directly from each mesh (signed-tetrahedron method — the
              same math the shop's Match Studio overlay uses). Mass is an estimate: STL files carry no material
              data, so it uses the quote's material when available and tool steel otherwise.
            </p>
          </div>
          <Button onClick={analyze} disabled={loading}>
            <RefreshCw className={loading ? "h-4 w-4 animate-spin" : "h-4 w-4"} />
            {loading ? `Analyzing ${progress.done}/${progress.total}...` : "Re-analyze"}
          </Button>
        </div>

        <Card className="overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full text-left text-sm">
              <thead>
                <tr className="border-b border-ink-700/60 bg-ink-850/40 text-[11px] uppercase tracking-wider text-ink-400">
                  <th className="px-4 py-3 font-medium">File</th>
                  <th className="px-4 py-3 font-medium">Part</th>
                  <th className="px-4 py-3 font-medium">Size (L × W × T)</th>
                  <th className="px-4 py-3 font-medium">Volume</th>
                  <th className="px-4 py-3 font-medium">Est. Mass</th>
                  <th className="px-4 py-3 font-medium">Center of Mass</th>
                </tr>
              </thead>
              <tbody>
                {analyzed.map((row) => (
                  <tr key={row.model.url} className="border-t border-ink-800/60">
                    <td className="max-w-[220px] truncate px-4 py-3 text-ink-100">{row.model.name}</td>
                    <td className="px-4 py-3 text-ink-400">{row.kind}</td>
                    {row.error ? (
                      <td className="px-4 py-3 text-accent-rose" colSpan={4}>
                        {row.error}
                      </td>
                    ) : (
                      <>
                        <td className="px-4 py-3 text-ink-300">
                          {row.stats.sx.toFixed(2)} × {row.stats.sy.toFixed(2)} × {row.stats.sz.toFixed(2)} in
                        </td>
                        <td className="px-4 py-3 text-ink-300">{row.stats.vol.toFixed(2)} in³</td>
                        <td className="px-4 py-3 text-ink-300">{row.massLb.toFixed(2)} lb</td>
                        <td className="px-4 py-3 text-ink-500">
                          {row.stats.comx.toFixed(2)}, {row.stats.comy.toFixed(2)}, {row.stats.comz.toFixed(2)}
                        </td>
                      </>
                    )}
                  </tr>
                ))}
                {analyzed.length === 0 && loading && (
                  <tr>
                    <td colSpan={6} className="px-4 py-6 text-center text-xs text-ink-500">
                      Analyzing {progress.done}/{progress.total} files...
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </Card>

        <Card className="p-5">
          <div className="section-label mb-3">Matching files in this quote</div>
          {likelyMatches.length === 0 ? (
            <p className="text-xs text-ink-500">
              {loading ? "Still analyzing..." : "No two files in this quote came back close enough to call a match."}
            </p>
          ) : (
            <div className="space-y-2">
              {likelyMatches.map((p, i) => (
                <div
                  key={i}
                  className="flex flex-wrap items-center justify-between gap-2 border-t border-ink-800/60 pt-2 first:border-t-0 first:pt-0"
                >
                  <div className="text-xs text-ink-200">
                    <span className="text-ink-100">{p.a.model.name}</span>
                    <span className="mx-2 text-ink-600">↔</span>
                    <span className="text-ink-100">{p.b.model.name}</span>
                  </div>
                  <div className="flex items-center gap-3 text-[11px] text-ink-500">
                    <span>
                      size Δ {p.match.sizeDiffPct.toFixed(1)}% · mass Δ {p.match.massDiffPct.toFixed(1)}% · COM{" "}
                      {p.match.centerDistanceIn.toFixed(3)} in
                    </span>
                    <Badge tone={matchTone(p.match.similarity)}>{p.match.similarity.toFixed(1)}% match</Badge>
                  </div>
                </div>
              ))}
            </div>
          )}
        </Card>
      </div>

      <div className="space-y-5">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <div>
            <div className="section-label">Against other quotes</div>
            <p className="mt-1 max-w-2xl text-xs text-ink-500">
              Checks this quote's TCP / BCP / ID &amp; OD Holder / ID &amp; OD Pot signature against every other
              quote's, using the exact same 0.40 size + 0.25 mass + 0.35 center-of-mass weighting (and the same
              8% / 15% / 15% tolerance bands) as the shop's Job Matching system.
            </p>
          </div>
          <div className="flex items-center gap-2">
            <select
              value={scope}
              onChange={(e) => setScope(Number(e.target.value))}
              className="glass-input rounded-full px-3 py-2 text-xs text-ink-200"
              disabled={matchLoading}
            >
              {SCOPE_OPTIONS.map((o) => (
                <option key={o.value} value={o.value}>
                  {o.label}
                </option>
              ))}
            </select>
            <Button onClick={findSimilarQuotes} disabled={matchLoading}>
              <Search className={matchLoading ? "h-4 w-4 animate-spin" : "h-4 w-4"} />
              {matchLoading
                ? `Comparing ${matchProgress.done}/${matchProgress.total}...`
                : "Find Similar Quotes"}
            </Button>
          </div>
        </div>

        {matchError && (
          <div className="rounded-xl border border-accent-rose/40 px-4 py-2.5 text-sm text-accent-rose">
            {matchError}
          </div>
        )}

        {jobMatches && (
          <Card className="p-5">
            {jobMatches.length === 0 ? (
              <p className="text-xs text-ink-500">
                No other quotes had a matching part to compare against this one.
              </p>
            ) : (
              <div className="space-y-3">
                {jobMatches.map((m) => {
                  const isOpen = expanded.has(m.jobId);
                  return (
                    <div key={m.jobId} className="border-t border-ink-800/60 pt-3 first:border-t-0 first:pt-0">
                      <div className="flex flex-wrap items-center justify-between gap-3">
                        <div className="flex items-center gap-3">
                          <Badge tone={matchTone(m.score)}>{m.score.toFixed(1)}%</Badge>
                          <Link
                            to={`/quotes/${m.jobId}`}
                            className="text-sm font-semibold uppercase tracking-wider text-ink-100 hover:text-white"
                          >
                            {m.jobId}
                          </Link>
                          <span className="max-w-[240px] truncate text-xs text-ink-500">{m.displayName}</span>
                        </div>
                        <div className="flex items-center gap-3 text-[11px] text-ink-500">
                          <span>
                            {m.matchedComponents}/{SIGNATURE_KINDS.length} parts matched · {m.coveragePct}% coverage
                          </span>
                          <button
                            type="button"
                            onClick={() => toggleExpanded(m.jobId)}
                            className="flex items-center gap-1 rounded-full border border-white/10 px-2.5 py-1 text-ink-400 hover:border-white/20 hover:text-ink-200"
                          >
                            Details {isOpen ? <ChevronUp className="h-3 w-3" /> : <ChevronDown className="h-3 w-3" />}
                          </button>
                        </div>
                      </div>
                      {isOpen && (
                        <div className="mt-3 overflow-x-auto">
                          <table className="w-full text-left text-xs">
                            <thead>
                              <tr className="text-[10px] uppercase tracking-wider text-ink-500">
                                <th className="py-1.5 pr-4 font-medium">Part</th>
                                <th className="py-1.5 pr-4 font-medium">Similarity</th>
                                <th className="py-1.5 pr-4 font-medium">Size Δ</th>
                                <th className="py-1.5 pr-4 font-medium">Mass Δ</th>
                                <th className="py-1.5 font-medium">COM dist</th>
                              </tr>
                            </thead>
                            <tbody>
                              {m.components.map((c) => (
                                <tr key={c.kind} className="border-t border-ink-800/40">
                                  <td className="py-1.5 pr-4 text-ink-200">{c.kind}</td>
                                  <td className="py-1.5 pr-4">
                                    <Badge tone={matchTone(c.similarity)}>{c.similarity.toFixed(1)}%</Badge>
                                  </td>
                                  <td className="py-1.5 pr-4 text-ink-400">{c.sizeDiffPct.toFixed(1)}%</td>
                                  <td className="py-1.5 pr-4 text-ink-400">{c.massDiffPct.toFixed(1)}%</td>
                                  <td className="py-1.5 text-ink-400">{c.centerDistanceIn.toFixed(3)} in</td>
                                </tr>
                              ))}
                              {(m.missingCurrent.length > 0 || m.missingCandidate.length > 0) && (
                                <tr className="border-t border-ink-800/40">
                                  <td colSpan={5} className="py-1.5 text-ink-600">
                                    {m.missingCandidate.length > 0 &&
                                      `${m.jobId} has no ${m.missingCandidate.join(", ")}. `}
                                    {m.missingCurrent.length > 0 &&
                                      `This quote has no ${m.missingCurrent.join(", ")}.`}
                                  </td>
                                </tr>
                              )}
                            </tbody>
                          </table>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            )}
          </Card>
        )}
      </div>
    </div>
  );
}
