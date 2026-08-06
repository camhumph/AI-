import { useEffect, useMemo, useState } from "react";
import { Link } from "react-router-dom";
import { Radar, Search, ChevronDown, ChevronUp, X } from "lucide-react";
import Layout from "../components/Layout";
import { Card, Button, Spinner, EmptyState, Badge } from "../components/ui";
import StlOverlayViewer from "../components/StlOverlayViewer";
import { matchTone } from "../lib/stlGeometry";
import {
  SIGNATURE_KINDS,
  buildJobSignature,
  compareJobSignatures,
  type JobSignature,
  type JobMatchResult,
} from "../lib/jobSignature";
import type { ComponentKind } from "../lib/componentKind";
import { api, type JobSummary } from "../api/client";

const SCOPE_OPTIONS = [
  { label: "Most recent 10", value: 10 },
  { label: "Most recent 25", value: 25 },
  { label: "Most recent 50", value: 50 },
  { label: "All quotes", value: 0 },
];

export default function JobMatchingPage() {
  const [jobs, setJobs] = useState<JobSummary[] | null>(null);
  useEffect(() => {
    api.listJobs().then(setJobs).catch(() => setJobs([]));
  }, []);

  const [query, setQuery] = useState("");
  const [currentJobId, setCurrentJobId] = useState<string | null>(null);

  const currentSummary = useMemo(
    () => (jobs || []).find((j) => j.job_id === currentJobId) || null,
    [jobs, currentJobId]
  );

  const filteredJobs = useMemo(() => {
    if (!jobs) return [];
    const q = query.trim().toLowerCase();
    const list = q
      ? jobs.filter((j) => j.job_id.toLowerCase().includes(q) || j.display_name.toLowerCase().includes(q))
      : jobs;
    return list
      .filter((j) => j.model_count > 0)
      .slice()
      .sort((a, b) => new Date(b.updated_at).getTime() - new Date(a.updated_at).getTime())
      .slice(0, 30);
  }, [jobs, query]);

  // ---- Build the signature for whichever job is picked as "current" ----
  const [currentSignature, setCurrentSignature] = useState<JobSignature | null>(null);
  const [buildingSignature, setBuildingSignature] = useState(false);

  const [scope, setScope] = useState(25);
  const [jobMatches, setJobMatches] = useState<JobMatchResult[] | null>(null);
  const [candidateSignatures, setCandidateSignatures] = useState<Map<string, JobSignature>>(new Map());
  const [matchLoading, setMatchLoading] = useState(false);
  const [matchProgress, setMatchProgress] = useState({ done: 0, total: 0 });
  const [matchError, setMatchError] = useState("");
  const [selectedMatchId, setSelectedMatchId] = useState<string | null>(null);
  const [overlayKind, setOverlayKind] = useState<ComponentKind | null>(null);
  const [expanded, setExpanded] = useState<Set<string>>(new Set());

  useEffect(() => {
    setJobMatches(null);
    setSelectedMatchId(null);
    setMatchError("");
    setCandidateSignatures(new Map());
    if (!currentJobId || !currentSummary) {
      setCurrentSignature(null);
      return;
    }
    setBuildingSignature(true);
    setCurrentSignature(null);
    api
      .getJob(currentJobId)
      .then((detail) => buildJobSignature(currentJobId, currentSummary.display_name, detail.models))
      .then(setCurrentSignature)
      .finally(() => setBuildingSignature(false));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [currentJobId]);

  const findSimilarQuotes = async () => {
    if (!currentSignature || currentSignature.components.size === 0) {
      setMatchError("This quote has no TCP/BCP/Holder/Pot STL files to compare yet.");
      return;
    }
    setMatchError("");
    setMatchLoading(true);
    setJobMatches(null);
    setSelectedMatchId(null);

    try {
      const allJobs = await api.listJobs();
      let candidates = allJobs.filter((j) => j.job_id !== currentJobId && j.model_count > 0);
      candidates = candidates
        .slice()
        .sort((a, b) => new Date(b.updated_at).getTime() - new Date(a.updated_at).getTime());
      if (scope > 0) candidates = candidates.slice(0, scope);

      setMatchProgress({ done: 0, total: candidates.length });
      const results: JobMatchResult[] = [];
      const sigMap = new Map<string, JobSignature>();

      for (const summary of candidates) {
        try {
          const detail = await api.getJob(summary.job_id);
          const sig = await buildJobSignature(summary.job_id, summary.display_name, detail.models);
          sigMap.set(summary.job_id, sig);
          const result = compareJobSignatures(currentSignature, sig);
          if (result) results.push(result);
        } catch {
          // skip jobs that fail to load
        }
        setMatchProgress((p) => ({ ...p, done: p.done + 1 }));
      }

      results.sort((a, b) => b.score - a.score);
      setJobMatches(results.slice(0, 10));
      setCandidateSignatures(sigMap);
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

  const selectedMatch = jobMatches?.find((m) => m.jobId === selectedMatchId) || null;
  const selectedSignature = selectedMatchId ? candidateSignatures.get(selectedMatchId) || null : null;

  const chooseMatch = (jobId: string) => {
    setSelectedMatchId(jobId);
    const match = jobMatches?.find((m) => m.jobId === jobId);
    setOverlayKind(match?.components[0]?.kind ?? null);
  };

  const overlayUrlA =
    overlayKind && currentSignature ? currentSignature.components.get(overlayKind)?.model.url : undefined;
  const overlayUrlB =
    overlayKind && selectedSignature ? selectedSignature.components.get(overlayKind)?.model.url : undefined;
  const overlayComponentScore = selectedMatch?.components.find((c) => c.kind === overlayKind) || null;

  return (
    <Layout
      title="Matching"
      subtitle="Pick a quote, rank it against every other quote by size, mass, and center of mass, then overlay any match to see the geometry side by side."
    >
      <div className="space-y-8">
        <Card className="p-5">
          <div className="section-label mb-3">Current quote</div>
          {currentJobId && currentSummary ? (
            <div className="flex flex-wrap items-center justify-between gap-3">
              <div className="flex items-center gap-3">
                <span className="text-sm font-semibold uppercase tracking-wider text-ink-100">
                  {currentSummary.job_id}
                </span>
                <span className="max-w-[320px] truncate text-xs text-ink-500">{currentSummary.display_name}</span>
                {buildingSignature && <span className="text-xs text-ink-500">Building signature...</span>}
                {currentSignature && !buildingSignature && (
                  <span className="text-xs text-ink-500">
                    {currentSignature.components.size}/{SIGNATURE_KINDS.length} parts found
                  </span>
                )}
              </div>
              <button
                type="button"
                onClick={() => setCurrentJobId(null)}
                className="flex items-center gap-1 rounded-full border border-white/10 px-3 py-1.5 text-[11px] uppercase tracking-wider text-ink-400 hover:border-white/20 hover:text-ink-200"
              >
                <X className="h-3 w-3" /> Change
              </button>
            </div>
          ) : (
            <div className="space-y-3">
              <div className="relative max-w-md">
                <Search className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-ink-500" />
                <input
                  value={query}
                  onChange={(e) => setQuery(e.target.value)}
                  placeholder="Search by C-number or folder name..."
                  className="glass-input w-full rounded-full py-2.5 pl-9 pr-3 text-sm text-ink-100 placeholder:text-ink-500"
                />
              </div>
              {jobs === null ? (
                <Spinner label="Loading quotes..." />
              ) : filteredJobs.length === 0 ? (
                <EmptyState
                  icon={<Search className="h-8 w-8" />}
                  title="No quotes found"
                  description="Try a different search, or check that the quote has 3D files uploaded."
                />
              ) : (
                <div className="max-h-64 space-y-1 overflow-y-auto">
                  {filteredJobs.map((j) => (
                    <button
                      key={j.job_id}
                      type="button"
                      onClick={() => setCurrentJobId(j.job_id)}
                      className="flex w-full items-center justify-between gap-3 rounded-xl px-3 py-2 text-left text-xs transition hover:bg-white/5"
                    >
                      <span className="font-semibold uppercase tracking-wider text-ink-100">{j.job_id}</span>
                      <span className="truncate text-ink-500">{j.display_name}</span>
                    </button>
                  ))}
                </div>
              )}
            </div>
          )}
        </Card>

        {currentJobId && (
          <>
            <div className="flex flex-wrap items-center justify-between gap-3">
              <p className="max-w-2xl text-xs text-ink-500">
                Checks this quote's TCP / BCP / ID &amp; OD Holder / ID &amp; OD Pot signature against every other
                quote's, using the same 0.40 size + 0.25 mass + 0.35 center-of-mass weighting (and the same 8% /
                15% / 15% tolerance bands) as the shop's Job Matching system.
              </p>
              <div className="flex items-center gap-2">
                <select
                  value={scope}
                  onChange={(e) => setScope(Number(e.target.value))}
                  className="glass-input rounded-full px-3 py-2 text-xs text-ink-200"
                  disabled={matchLoading || buildingSignature}
                >
                  {SCOPE_OPTIONS.map((o) => (
                    <option key={o.value} value={o.value}>
                      {o.label}
                    </option>
                  ))}
                </select>
                <Button onClick={findSimilarQuotes} disabled={matchLoading || buildingSignature}>
                  <Radar className={matchLoading ? "h-4 w-4 animate-spin" : "h-4 w-4"} />
                  {matchLoading ? `Comparing ${matchProgress.done}/${matchProgress.total}...` : "Find Similar Quotes"}
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
                  <p className="text-xs text-ink-500">No other quotes had a matching part to compare against this one.</p>
                ) : (
                  <div className="space-y-3">
                    {jobMatches.map((m) => {
                      const isOpen = expanded.has(m.jobId);
                      const isChosen = selectedMatchId === m.jobId;
                      return (
                        <div
                          key={m.jobId}
                          className={`rounded-xl border p-3 transition ${
                            isChosen ? "border-white/25 bg-white/5" : "border-transparent"
                          }`}
                        >
                          <div className="flex flex-wrap items-center justify-between gap-3">
                            <div className="flex items-center gap-3">
                              <Badge tone={matchTone(m.score)}>{m.score.toFixed(1)}%</Badge>
                              <Link
                                to={`/quotes/${m.jobId}`}
                                className="text-sm font-semibold uppercase tracking-wider text-ink-100 hover:text-white"
                              >
                                {m.jobId}
                              </Link>
                              <span className="max-w-[220px] truncate text-xs text-ink-500">{m.displayName}</span>
                            </div>
                            <div className="flex items-center gap-2 text-[11px] text-ink-500">
                              <span>
                                {m.matchedComponents}/{SIGNATURE_KINDS.length} parts · {m.coveragePct}% coverage
                              </span>
                              <button
                                type="button"
                                onClick={() => toggleExpanded(m.jobId)}
                                className="flex items-center gap-1 rounded-full border border-white/10 px-2.5 py-1 text-ink-400 hover:border-white/20 hover:text-ink-200"
                              >
                                Details {isOpen ? <ChevronUp className="h-3 w-3" /> : <ChevronDown className="h-3 w-3" />}
                              </button>
                              <Button variant={isChosen ? "primary" : "secondary"} onClick={() => chooseMatch(m.jobId)}>
                                {isChosen ? "Comparing" : "Compare"}
                              </Button>
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

            {selectedMatch && selectedSignature && (
              <Card className="p-4">
                <div className="mb-3 flex flex-wrap items-center justify-between gap-3">
                  <div className="section-label">
                    Overlay — {currentSummary?.job_id} vs {selectedMatch.jobId}
                  </div>
                  <div className="flex items-center gap-4 text-[10px] uppercase tracking-wider text-ink-400">
                    <span className="flex items-center gap-1.5">
                      <span className="h-2.5 w-2.5 rounded-full" style={{ background: "#3b82f6" }} /> {currentSummary?.job_id}
                    </span>
                    <span className="flex items-center gap-1.5">
                      <span className="h-2.5 w-2.5 rounded-full" style={{ background: "#ff8c42" }} /> {selectedMatch.jobId}
                    </span>
                  </div>
                </div>

                <div className="mb-4 flex flex-wrap gap-2">
                  {selectedMatch.components.map((c) => (
                    <button
                      key={c.kind}
                      type="button"
                      onClick={() => setOverlayKind(c.kind)}
                      className={`flex items-center gap-2 rounded-full border px-3.5 py-1.5 text-[11px] font-semibold uppercase tracking-wider transition ${
                        overlayKind === c.kind
                          ? "border-white/25 bg-white/10 text-ink-100"
                          : "border-white/10 text-ink-400 hover:border-white/20 hover:text-ink-200"
                      }`}
                    >
                      {c.kind}
                      <Badge tone={matchTone(c.similarity)}>{c.similarity.toFixed(0)}%</Badge>
                    </button>
                  ))}
                </div>

                <div className="h-[420px]">
                  {overlayUrlA && overlayUrlB ? (
                    <StlOverlayViewer key={`${overlayUrlA}::${overlayUrlB}`} urlA={overlayUrlA} urlB={overlayUrlB} />
                  ) : (
                    <div className="flex h-full items-center justify-center text-xs text-ink-500">
                      Pick a matched part above
                    </div>
                  )}
                </div>

                {overlayComponentScore && (
                  <div className="mt-4 text-xs text-ink-400">
                    size Δ {overlayComponentScore.sizeDiffPct.toFixed(1)}% · mass Δ{" "}
                    {overlayComponentScore.massDiffPct.toFixed(1)}% · COM dist{" "}
                    {overlayComponentScore.centerDistanceIn.toFixed(3)} in
                  </div>
                )}
              </Card>
            )}
          </>
        )}
      </div>
    </Layout>
  );
}
