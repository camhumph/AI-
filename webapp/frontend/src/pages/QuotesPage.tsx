import { useEffect, useMemo, useState } from "react";
import { Link, useNavigate } from "react-router-dom";
import { FileStack, Image as ImageIcon, Box, Plus, Search, FolderOpen, ChevronRight, ChevronUp, ChevronDown, Trash2 } from "lucide-react";
import Layout from "../components/Layout";
import { Card, Badge, Button, EmptyState, Spinner } from "../components/ui";
import { api, type JobSummary, type WorkspaceBrowse, type WorkspaceEntry } from "../api/client";
import { useQuoteJobs } from "../context/QuoteJobsContext";

export default function QuotesPage() {
  const navigate = useNavigate();
  const [jobs, setJobs] = useState<JobSummary[] | null>(null);
  const [query, setQuery] = useState("");
  const [showCreate, setShowCreate] = useState(false);
  const [deleting, setDeleting] = useState<string | null>(null);
  const [sortKey, setSortKey] = useState<"number" | "name" | "date">("date");
  const [sortDir, setSortDir] = useState<"asc" | "desc">("desc");

  const refresh = () => api.listJobs().then(setJobs).catch(() => setJobs([]));
  useEffect(() => {
    refresh();
  }, []);

  const filtered = (jobs || []).filter(
    (j) =>
      j.display_name.toLowerCase().includes(query.toLowerCase()) ||
      j.job_id.toLowerCase().includes(query.toLowerCase())
  );

  const sorted = useMemo(() => {
    const list = [...filtered];
    list.sort((a, b) => {
      let cmp = 0;
      if (sortKey === "number") {
        cmp = a.job_id.localeCompare(b.job_id, undefined, { numeric: true });
      } else if (sortKey === "name") {
        cmp = a.display_name.localeCompare(b.display_name);
      } else {
        cmp = new Date(a.updated_at).getTime() - new Date(b.updated_at).getTime();
      }
      return sortDir === "asc" ? cmp : -cmp;
    });
    return list;
  }, [filtered, sortKey, sortDir]);

  const toggleSort = (key: "number" | "name" | "date") => {
    if (sortKey === key) {
      setSortDir((d) => (d === "asc" ? "desc" : "asc"));
    } else {
      setSortKey(key);
      setSortDir(key === "date" ? "desc" : "asc");
    }
  };

  const formatDate = (iso: string) => {
    if (!iso) return "--";
    const d = new Date(iso);
    if (Number.isNaN(d.getTime())) return "--";
    return d.toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric" });
  };

  const deleteQuote = async (jobId: string, e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (!window.confirm(`Delete quote ${jobId} from the webapp?\n\nThis removes it from the Quotes list. SolidWorks job files on disk are not deleted.`)) {
      return;
    }
    setDeleting(jobId);
    try {
      await api.deleteJob(jobId);
      setJobs((prev) => (prev || []).filter((j) => j.job_id !== jobId));
    } catch (err) {
      window.alert((err as Error).message || "Delete failed");
    } finally {
      setDeleting(null);
    }
  };

  return (
    <Layout
      title="Quotes"
      subtitle="Select one or more C-number folders from your workspace, or open an existing quote."
      actions={
        <Button onClick={() => setShowCreate(true)}>
          <Plus className="h-4 w-4" /> New Quote
        </Button>
      }
    >
      <div className="mb-5 flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
        <div className="relative max-w-md flex-1">
          <Search className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-ink-500" />
          <input
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="Search by C-number or folder name..."
            className="glass-input w-full rounded-full py-2.5 pl-9 pr-3 text-sm text-ink-100 placeholder:text-ink-500"
          />
        </div>
        <div className="flex items-center gap-2 text-[10px] font-semibold uppercase tracking-widest text-ink-400">
          <span className="hidden sm:inline">Sort</span>
          {(["number", "name", "date"] as const).map((key) => (
            <button
              key={key}
              onClick={() => toggleSort(key)}
              className={`flex items-center gap-1 rounded-full border px-3 py-1.5 transition ${
                sortKey === key
                  ? "border-white/25 bg-white/10 text-ink-100"
                  : "border-white/10 text-ink-400 hover:border-white/20 hover:text-ink-200"
              }`}
            >
              {key === "number" ? "Number" : key === "name" ? "Name" : "Date"}
              {sortKey === key &&
                (sortDir === "asc" ? <ChevronUp className="h-3 w-3" /> : <ChevronDown className="h-3 w-3" />)}
            </button>
          ))}
        </div>
      </div>

      {jobs === null ? (
        <Spinner label="Loading quotes..." />
      ) : sorted.length === 0 ? (
        <EmptyState
          icon={<FileStack className="h-8 w-8" />}
          title="No quotes yet"
          description="Click New Quote, check one or more C-number folders, then Quote selected."
          action={<Button onClick={() => setShowCreate(true)}><Plus className="h-4 w-4" /> New Quote</Button>}
        />
      ) : (
        <Card className="overflow-hidden">
          <div className="scrollbar-thin overflow-x-auto">
            <table className="w-full text-left text-sm">
              <thead>
                <tr className="border-b border-ink-700/60 bg-ink-850/40 text-[11px] uppercase tracking-wider text-ink-400">
                  <th className="px-4 py-3 font-medium">
                    <button onClick={() => toggleSort("number")} className="flex items-center gap-1 hover:text-ink-100">
                      C-Number
                      {sortKey === "number" && (sortDir === "asc" ? <ChevronUp className="h-3 w-3" /> : <ChevronDown className="h-3 w-3" />)}
                    </button>
                  </th>
                  <th className="px-4 py-3 font-medium">
                    <button onClick={() => toggleSort("name")} className="flex items-center gap-1 hover:text-ink-100">
                      Name
                      {sortKey === "name" && (sortDir === "asc" ? <ChevronUp className="h-3 w-3" /> : <ChevronDown className="h-3 w-3" />)}
                    </button>
                  </th>
                  <th className="px-4 py-3 font-medium">Parts</th>
                  <th className="px-4 py-3 font-medium">Images</th>
                  <th className="px-4 py-3 font-medium">Model</th>
                  <th className="px-4 py-3 font-medium">
                    <button onClick={() => toggleSort("date")} className="flex items-center gap-1 hover:text-ink-100">
                      Updated
                      {sortKey === "date" && (sortDir === "asc" ? <ChevronUp className="h-3 w-3" /> : <ChevronDown className="h-3 w-3" />)}
                    </button>
                  </th>
                  <th className="px-4 py-3 text-right font-medium">Actions</th>
                </tr>
              </thead>
              <tbody>
                {sorted.map((job) => (
                  <tr
                    key={job.job_id}
                    onClick={() => navigate(`/quotes/${job.job_id}`)}
                    className="cursor-pointer border-t border-ink-800/60 transition hover:bg-ink-800/30"
                  >
                    <td className="px-4 py-3 font-semibold uppercase tracking-wider text-ink-100">
                      <div className="flex items-center gap-2">
                        {job.job_id}
                        {job.base_type === "bms" && <Badge tone="warning">BMS</Badge>}
                      </div>
                    </td>
                    <td className="max-w-xs truncate px-4 py-3 text-ink-300">{job.display_name}</td>
                    <td className="px-4 py-3 text-ink-400">
                      <span className="flex items-center gap-1"><FileStack className="h-3.5 w-3.5" /> {job.part_count}</span>
                    </td>
                    <td className="px-4 py-3 text-ink-400">
                      <span className="flex items-center gap-1"><ImageIcon className="h-3.5 w-3.5" /> {job.image_count}</span>
                    </td>
                    <td className="px-4 py-3 text-ink-400">
                      <span className="flex items-center gap-1"><Box className="h-3.5 w-3.5" /> {job.model_count}</span>
                    </td>
                    <td className="px-4 py-3 text-xs text-ink-500">{formatDate(job.updated_at)}</td>
                    <td className="px-4 py-3 text-right" onClick={(e) => e.stopPropagation()}>
                      <div className="flex items-center justify-end gap-2">
                        <Link to={`/quotes/${job.job_id}`} className="text-ink-500 hover:text-ink-100" title="Open quote">
                          <ChevronRight className="h-4 w-4" />
                        </Link>
                        <button
                          type="button"
                          onClick={(e) => deleteQuote(job.job_id, e)}
                          disabled={deleting === job.job_id}
                          title="Delete quote"
                          className="rounded-full p-1.5 text-ink-500 hover:bg-accent-rose/15 hover:text-accent-rose disabled:opacity-40"
                        >
                          <Trash2 className="h-3.5 w-3.5" />
                        </button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Card>
      )}

      {showCreate && (
        <FolderPickerModal
          onClose={() => setShowCreate(false)}
          onImported={refresh}
        />
      )}
    </Layout>
  );
}

function isQuoteable(entry: WorkspaceEntry): boolean {
  // Checkbox every folder the shop can quote (C##### / BMS- / XT present).
  return Boolean(entry.is_dir && entry.quote_ready);
}

function FolderPickerModal({
  onClose,
  onImported,
}: {
  onClose: () => void;
  onImported?: () => void;
}) {
  const [browse, setBrowse] = useState<WorkspaceBrowse | null>(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [checked, setChecked] = useState<Set<string>>(new Set());
  const { startQuote } = useQuoteJobs();

  const load = (path = "") => {
    setError("");
    setChecked(new Set());
    api.browseWorkspace(path).then(setBrowse).catch((e) => setError(e.message));
  };

  useEffect(() => {
    load();
  }, []);

  const dirs = useMemo(
    () => (browse?.entries || []).filter((e) => e.is_dir),
    [browse]
  );
  const quoteable = useMemo(() => dirs.filter(isQuoteable), [dirs]);

  const toggleChecked = (path: string) => {
    setChecked((prev) => {
      const next = new Set(prev);
      if (next.has(path)) next.delete(path);
      else next.add(path);
      return next;
    });
  };

  const toggleAllQuoteable = () => {
    if (checked.size > 0 && quoteable.every((e) => checked.has(e.path))) {
      setChecked(new Set());
      return;
    }
    setChecked(new Set(quoteable.map((e) => e.path)));
  };

  const importFolder = async (folderPath: string) => {
    setBusy(true);
    setError("");
    try {
      const result = await api.importFromFolder(folderPath, true);
      const qid = result.quote_id || result.job_id;
      const label = folderPath.split(/[/\\]/).pop() || "Folder quote";
      startQuote(qid, label, {
        phase: "running",
        message: "Running in background — SolidWorks + Module6121",
        job_id: result.job_id,
      });
      onImported?.();
      onClose();
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  };

  const quoteSelected = async () => {
    const paths = Array.from(checked);
    if (paths.length === 0) return;
    if (paths.length === 1) {
      await importFolder(paths[0]);
      return;
    }
    setBusy(true);
    setError("");
    try {
      const result = await api.importFoldersBatch(paths, true);
      if (!result.launched && result.error) {
        throw new Error(result.error);
      }
      const jobs = result.jobs || [];
      const qids = result.quote_ids || jobs.map((j) => j.quote_id);
      qids.forEach((qid, i) => {
        const job = jobs[i];
        const label =
          job?.display_name ||
          paths[i]?.split(/[/\\]/).pop() ||
          qid;
        startQuote(qid, label, {
          phase: "running",
          message: `Batch ${i + 1}/${qids.length} — one job at a time (SolidWorks restarts between)`,
          job_id: job?.job_id || qid,
        });
      });
      if (result.errors && result.errors.length > 0) {
        setError(result.errors.join("; "));
      } else {
        onImported?.();
        onClose();
      }
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="overlay-backdrop fixed inset-0 z-50 flex items-center justify-center p-4">
      <Card className="flex max-h-[80vh] w-full max-w-2xl flex-col">
        <div className="flex items-center justify-between border-b border-ink-700 px-6 py-4">
          <div>
            <h3 className="text-sm font-semibold uppercase tracking-widest text-ink-100">Select Quote Folder</h3>
            <p className="mt-1 text-xs text-ink-400">
              Check one or more C-number jobs, then Quote selected. Browse under{" "}
              <code className="text-ink-300">\\Mycloudex2ultra\mexico\Downloads</code>
            </p>
          </div>
          <button onClick={onClose} className="text-ink-400 hover:text-ink-100 text-xs uppercase tracking-widest">Close</button>
        </div>

        <div className="border-b border-ink-700 px-6 py-3">
          <div className="flex flex-wrap items-center gap-2 text-xs text-ink-400">
            <FolderOpen className="h-4 w-4" />
            <span className="truncate font-mono">{browse?.path || "..."}</span>
            {browse?.roots && browse.roots.length > 0 && (
              <div className="ml-auto flex flex-wrap gap-1">
                {browse.roots.slice(0, 5).map((root) => {
                  // An offline share stays visible and says so. Hiding it would
                  // leave someone wondering where the network went.
                  const offline = browse.unreachable?.includes(root) ?? false;
                  return (
                    <button
                      key={root}
                      onClick={() => !offline && load(root)}
                      disabled={offline}
                      className={
                        offline
                          ? "cursor-not-allowed rounded-full border border-white/5 px-2 py-0.5 text-[10px] uppercase tracking-wider text-ink-600 line-through"
                          : "rounded-full border border-white/10 px-2 py-0.5 text-[10px] uppercase tracking-wider text-ink-400 hover:border-white/20 hover:text-ink-200"
                      }
                      title={offline ? `${root}\n\nNot reachable — off the company wifi?` : root}
                    >
                      {root.split(/[/\\]/).filter(Boolean).slice(-2).join("\\") || root}
                    </button>
                  );
                })}
              </div>
            )}
            {browse?.parent && (
              <button onClick={() => load(browse.parent!)} className="flex items-center gap-1 text-ink-300 hover:text-ink-100">
                <ChevronUp className="h-3.5 w-3.5" /> Up
              </button>
            )}
          </div>
          {quoteable.length > 0 && (
            <div className="mt-3 flex flex-wrap items-center gap-2">
              <label className="flex cursor-pointer items-center gap-2 rounded-full border border-white/10 px-3 py-1 text-[10px] uppercase tracking-wider text-ink-300 hover:border-white/20 hover:text-ink-100">
                <input
                  type="checkbox"
                  className="accent-brand-400"
                  checked={checked.size > 0 && quoteable.every((e) => checked.has(e.path))}
                  onChange={toggleAllQuoteable}
                />
                Select all quoteable ({quoteable.length})
              </label>
              {checked.size > 0 && (
                <Button
                  variant="primary"
                  className="px-3 py-1.5 text-[10px]"
                  disabled={busy}
                  onClick={quoteSelected}
                >
                  {busy ? "Starting..." : `Quote selected (${checked.size})`}
                </Button>
              )}
            </div>
          )}
        </div>

        <div className="scrollbar-thin flex-1 overflow-y-auto px-2 py-2">
          {!browse ? (
            <div className="p-6"><Spinner label="Loading folders..." /></div>
          ) : !browse.exists ? (
            <div className="p-6 text-xs text-ink-400">
              Workspace folder not found on this machine. Set <code className="text-ink-200">CMS_WORKSPACE_ROOT</code> in the start script.
            </div>
          ) : dirs.length === 0 ? (
            <div className="p-6 text-xs text-ink-400">No folders here.</div>
          ) : (
            dirs.map((entry) => {
              const canQuote = isQuoteable(entry);
              return (
                <div
                  key={entry.path}
                  className="flex items-center gap-3 border-b border-ink-800 px-4 py-3 hover:bg-ink-900"
                >
                  {canQuote ? (
                    <input
                      type="checkbox"
                      checked={checked.has(entry.path)}
                      onChange={() => toggleChecked(entry.path)}
                      onClick={(e) => e.stopPropagation()}
                      className="h-4 w-4 shrink-0 accent-brand-400"
                      title="Select for batch quote"
                    />
                  ) : (
                    <span className="inline-block h-4 w-4 shrink-0" />
                  )}
                  <button
                    onClick={() => load(entry.path)}
                    className="flex min-w-0 flex-1 items-center gap-2 text-left"
                  >
                    <FolderOpen className="h-4 w-4 shrink-0 text-ink-500" />
                    <div className="min-w-0">
                      <div className="truncate text-sm text-ink-100">{entry.name}</div>
                      <div className="flex gap-2 text-[10px] text-ink-500">
                        {entry.c_number && <span>{entry.c_number}</span>}
                        {entry.has_xt_csv && <span>XT</span>}
                        {entry.has_quote_sheet && <span>Quote</span>}
                        {entry.has_steel_sheet && <span>Steel</span>}
                      </div>
                    </div>
                    <ChevronRight className="ml-auto h-4 w-4 shrink-0 text-ink-600" />
                  </button>
                  {canQuote && (
                    <Button
                      variant="primary"
                      className="shrink-0 px-3 py-1.5 text-[10px]"
                      disabled={busy}
                      onClick={() => importFolder(entry.path)}
                    >
                      Quote
                    </Button>
                  )}
                </div>
              );
            })
          )}
        </div>

        {error && <p className="border-t border-ink-700 px-6 py-3 text-xs text-accent-rose">{error}</p>}
      </Card>
    </div>
  );
}
