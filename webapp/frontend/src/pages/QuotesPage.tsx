import { useEffect, useRef, useState } from "react";
import { Link, useNavigate } from "react-router-dom";
import { FileStack, Image as ImageIcon, Box, Plus, Search, Upload, X } from "lucide-react";
import Layout from "../components/Layout";
import { Card, Badge, Button, EmptyState, Spinner } from "../components/ui";
import { api, type JobSummary } from "../api/client";

export default function QuotesPage() {
  const [jobs, setJobs] = useState<JobSummary[] | null>(null);
  const [query, setQuery] = useState("");
  const [showCreate, setShowCreate] = useState(false);
  const navigate = useNavigate();

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
      subtitle="Browse jobs from C:\\CMS_Local_Workspace or upload a new CAD export to quote."
      actions={
        <Button onClick={() => setShowCreate(true)}>
          <Plus className="h-4 w-4" /> New Quote
        </Button>
      }
    >
      <div className="mb-5 relative max-w-md">
        <Search className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-ink-400" />
        <input
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          placeholder="Search jobs by name or job ID..."
          className="w-full rounded-xl border border-ink-700/60 bg-ink-850/70 py-2.5 pl-9 pr-3 text-sm text-ink-100 placeholder:text-ink-500 focus:border-brand-500/60 focus:outline-none"
        />
      </div>

      {jobs === null ? (
        <Spinner label="Loading jobs..." />
      ) : filtered.length === 0 ? (
        <EmptyState
          icon={<FileStack className="h-8 w-8" />}
          title="No jobs found"
          description="Create a new quote job and upload an XT_Export_CAD_Dimensions.csv to get started."
          action={<Button onClick={() => setShowCreate(true)}><Plus className="h-4 w-4" /> New Quote</Button>}
        />
      ) : (
        <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-3">
          {filtered.map((job) => (
            <Link key={job.job_id} to={`/quotes/${job.job_id}`}>
              <Card className="h-full p-5 transition hover:-translate-y-0.5 hover:border-brand-500/40 hover:shadow-brand-500/10">
                <div className="mb-3 flex items-start justify-between gap-3">
                  <div className="min-w-0">
                    <div className="text-sm font-semibold text-ink-100">{job.display_name}</div>
                    <div className="text-xs text-ink-400">{job.job_id}</div>
                  </div>
                  {job.base_type === "bms" ? (
                    <Badge tone="warning">BMS</Badge>
                  ) : (
                    job.sequenced_latch_lock_base && <Badge tone="brand">Latch-Lock</Badge>
                  )}
                </div>
                {job.customer && <div className="mb-3 text-xs text-ink-400">{job.customer}</div>}
                <div className="flex items-center gap-4 text-xs text-ink-400">
                  <span className="flex items-center gap-1">
                    <FileStack className="h-3.5 w-3.5" /> {job.part_count}
                  </span>
                  <span className="flex items-center gap-1">
                    <ImageIcon className="h-3.5 w-3.5" /> {job.image_count}
                  </span>
                  <span className="flex items-center gap-1">
                    <Box className="h-3.5 w-3.5" /> {job.model_count}
                  </span>
                </div>
                <div className="mt-3">
                  {job.base_type === "bms" ? (
                    <Badge>BOM-driven (Module6121) -- AI off</Badge>
                  ) : job.has_classification ? (
                    <Badge tone="success">AI Classified</Badge>
                  ) : job.has_raw_csv ? (
                    <Badge tone="warning">Ready to classify</Badge>
                  ) : (
                    <Badge>Needs CAD export</Badge>
                  )}
                </div>
              </Card>
            </Link>
          ))}
        </div>
      )}

      {showCreate && (
        <CreateQuoteModal
          onClose={() => setShowCreate(false)}
          onCreated={(jobId) => {
            setShowCreate(false);
            refresh();
            navigate(`/quotes/${jobId}`);
          }}
        />
      )}
    </Layout>
  );
}

function CreateQuoteModal({ onClose, onCreated }: { onClose: () => void; onCreated: (jobId: string) => void }) {
  const [jobId, setJobId] = useState("");
  const [displayName, setDisplayName] = useState("");
  const [customer, setCustomer] = useState("");
  const [file, setFile] = useState<File | null>(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const fileRef = useRef<HTMLInputElement>(null);

  const submit = async () => {
    if (!jobId.trim()) {
      setError("Job ID is required (e.g. J8420, T001015, C18606).");
      return;
    }
    setBusy(true);
    setError("");
    try {
      await api.createJob(jobId.trim(), displayName.trim(), customer.trim());
      if (file) {
        await api.uploadFile(jobId.trim(), "raw", file);
      }
      onCreated(jobId.trim());
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 p-4 backdrop-blur-sm">
      <Card className="w-full max-w-md p-6">
        <div className="mb-4 flex items-center justify-between">
          <h3 className="text-base font-semibold text-ink-100">New Quote Job</h3>
          <button onClick={onClose} className="text-ink-400 hover:text-ink-100">
            <X className="h-5 w-5" />
          </button>
        </div>

        <div className="space-y-3">
          <div>
            <label className="mb-1 block text-xs font-medium text-ink-400">Job ID *</label>
            <input
              value={jobId}
              onChange={(e) => setJobId(e.target.value)}
              placeholder="J8420"
              className="w-full rounded-lg border border-ink-700/60 bg-ink-850 px-3 py-2 text-sm text-ink-100 focus:border-brand-500/60 focus:outline-none"
            />
          </div>
          <div>
            <label className="mb-1 block text-xs font-medium text-ink-400">Display Name</label>
            <input
              value={displayName}
              onChange={(e) => setDisplayName(e.target.value)}
              placeholder="2-Cavity Funnel Mold"
              className="w-full rounded-lg border border-ink-700/60 bg-ink-850 px-3 py-2 text-sm text-ink-100 focus:border-brand-500/60 focus:outline-none"
            />
          </div>
          <div>
            <label className="mb-1 block text-xs font-medium text-ink-400">Customer</label>
            <input
              value={customer}
              onChange={(e) => setCustomer(e.target.value)}
              className="w-full rounded-lg border border-ink-700/60 bg-ink-850 px-3 py-2 text-sm text-ink-100 focus:border-brand-500/60 focus:outline-none"
            />
          </div>
          <div>
            <label className="mb-1 block text-xs font-medium text-ink-400">
              XT_Export_CAD_Dimensions.csv (optional -- can upload later)
            </label>
            <button
              onClick={() => fileRef.current?.click()}
              className="flex w-full items-center justify-center gap-2 rounded-lg border border-dashed border-ink-600/70 bg-ink-850/60 py-3 text-xs text-ink-300 hover:border-brand-500/50"
            >
              <Upload className="h-4 w-4" /> {file ? file.name : "Choose CSV file"}
            </button>
            <input
              ref={fileRef}
              type="file"
              accept=".csv"
              className="hidden"
              onChange={(e) => setFile(e.target.files?.[0] || null)}
            />
          </div>
          {error && <p className="text-xs text-accent-rose">{error}</p>}
        </div>

        <div className="mt-5 flex justify-end gap-2">
          <Button variant="ghost" onClick={onClose}>
            Cancel
          </Button>
          <Button onClick={submit} disabled={busy}>
            {busy ? "Creating..." : "Create Job"}
          </Button>
        </div>
      </Card>
    </div>
  );
}
