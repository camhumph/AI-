import { useEffect, useState } from "react";
import { Link } from "react-router-dom";
import { FileStack, Image as ImageIcon, Box, Plus, Search, FolderOpen, ChevronRight, ChevronUp } from "lucide-react";
import Layout from "../components/Layout";
import { Card, Badge, Button, EmptyState, Spinner } from "../components/ui";
import { api, type JobSummary, type WorkspaceBrowse } from "../api/client";
import { useQuoteJobs } from "../context/QuoteJobsContext";

export default function QuotesPage() {
  const [jobs, setJobs] = useState<JobSummary[] | null>(null);
  const [query, setQuery] = useState("");
  const [showCreate, setShowCreate] = useState(false);

  const refresh = () => api.listJobs().then(setJobs).catch(() => setJobs([]));
  useEffect(() => {
    refresh();
  }, []);

  const filtered = (jobs || []).filter(
    (j) =>
      j.display_name.toLowerCase().includes(query.toLowerCase()) ||
      j.job_id.toLowerCase().includes(query.toLowerCase())
  );

  return (
    <Layout
      title="Quotes"
      subtitle="Select a C-number folder from your workspace or open an existing quote."
      actions={
        <Button onClick={() => setShowCreate(true)}>
          <Plus className="h-4 w-4" /> New Quote
        </Button>
      }
    >
      <div className="relative mb-5 max-w-md">
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
      ) : filtered.length === 0 ? (
        <EmptyState
          icon={<FileStack className="h-8 w-8" />}
          title="No quotes yet"
          description="Click New Quote and browse to a C-number folder on your machine."
          action={<Button onClick={() => setShowCreate(true)}><Plus className="h-4 w-4" /> New Quote</Button>}
        />
      ) : (
        <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-3">
          {filtered.map((job) => (
            <Link key={job.job_id} to={`/quotes/${job.job_id}`}>
              <Card className="h-full p-5 transition hover:border-ink-100">
                <div className="mb-3 flex items-start justify-between gap-3">
                  <div className="min-w-0">
                    <div className="text-sm font-semibold uppercase tracking-wider text-ink-100">{job.job_id}</div>
                    <div className="mt-1 truncate text-xs text-ink-400">{job.display_name}</div>
                  </div>
                  {job.base_type === "bms" && <Badge tone="warning">BMS</Badge>}
                </div>
                <div className="flex items-center gap-4 text-xs text-ink-500">
                  <span className="flex items-center gap-1"><FileStack className="h-3.5 w-3.5" /> {job.part_count}</span>
                  <span className="flex items-center gap-1"><ImageIcon className="h-3.5 w-3.5" /> {job.image_count}</span>
                  <span className="flex items-center gap-1"><Box className="h-3.5 w-3.5" /> {job.model_count}</span>
                </div>
              </Card>
            </Link>
          ))}
        </div>
      )}

      {showCreate && (
        <FolderPickerModal
          onClose={() => setShowCreate(false)}
        />
      )}
    </Layout>
  );
}

function FolderPickerModal({ onClose }: { onClose: () => void }) {
  const [browse, setBrowse] = useState<WorkspaceBrowse | null>(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const { startQuote } = useQuoteJobs();

  const load = (path = "") => {
    setError("");
    api.browseWorkspace(path).then(setBrowse).catch((e) => setError(e.message));
  };

  useEffect(() => {
    load();
  }, []);

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
      onClose();
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
            <p className="mt-1 text-xs text-ink-400">Browse to a C-number job folder on this PC</p>
          </div>
          <button onClick={onClose} className="text-ink-400 hover:text-ink-100 text-xs uppercase tracking-widest">Close</button>
        </div>

        <div className="border-b border-ink-700 px-6 py-3">
          <div className="flex items-center gap-2 text-xs text-ink-400">
            <FolderOpen className="h-4 w-4" />
            <span className="truncate font-mono">{browse?.path || "..."}</span>
            {browse?.parent && (
              <button onClick={() => load(browse.parent!)} className="ml-auto flex items-center gap-1 text-ink-300 hover:text-ink-100">
                <ChevronUp className="h-3.5 w-3.5" /> Up
              </button>
            )}
          </div>
        </div>

        <div className="scrollbar-thin flex-1 overflow-y-auto px-2 py-2">
          {!browse ? (
            <div className="p-6"><Spinner label="Loading folders..." /></div>
          ) : !browse.exists ? (
            <div className="p-6 text-xs text-ink-400">
              Workspace folder not found on this machine. Set <code className="text-ink-200">CMS_WORKSPACE_ROOT</code> in the start script.
            </div>
          ) : browse.entries.length === 0 ? (
            <div className="p-6 text-xs text-ink-400">No folders here.</div>
          ) : (
            browse.entries
              .filter((e) => e.is_dir)
              .map((entry) => (
                <div
                  key={entry.path}
                  className="flex items-center gap-3 border-b border-ink-800 px-4 py-3 hover:bg-ink-900"
                >
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
                  {(entry.has_xt_csv || entry.c_number || entry.name.match(/\.(sldasm|sldprt|x_t|step|stp|zip)$/i)) && (
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
              ))
          )}
        </div>

        {error && <p className="border-t border-ink-700 px-6 py-3 text-xs text-accent-rose">{error}</p>}
      </Card>
    </div>
  );
}
