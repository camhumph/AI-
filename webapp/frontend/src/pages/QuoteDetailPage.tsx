import { lazy, Suspense, useCallback, useEffect, useRef, useState } from "react";
import { useParams, Link } from "react-router-dom";
import {
  ArrowLeft, RefreshCw, Download, Upload, Image as ImageIcon, Box,
  FileText, Layers, Sparkles, ChevronLeft,
} from "lucide-react";
import Layout from "../components/Layout";
import { Card, Badge, Button, Spinner, EmptyState } from "../components/ui";
import PartsTable from "../components/PartsTable";
import ImageGallery from "../components/ImageGallery";
import { api, type JobDetail, type QuoteSheet } from "../api/client";

// three.js is the single largest dependency in this app -- only load it when
// the "3D Model" tab is actually opened.
const StlViewer = lazy(() => import("../components/StlViewer"));

type Tab = "overview" | "parts" | "images" | "model" | "documents";

export default function QuoteDetailPage() {
  const { jobId = "" } = useParams();
  const [job, setJob] = useState<JobDetail | null>(null);
  const [quote, setQuote] = useState<QuoteSheet | null>(null);
  const [tab, setTab] = useState<Tab>("overview");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [activeModel, setActiveModel] = useState<string | null>(null);
  const rawFileRef = useRef<HTMLInputElement>(null);
  const modelFileRef = useRef<HTMLInputElement>(null);
  const docFileRef = useRef<HTMLInputElement>(null);

  const load = useCallback(() => {
    api.getJob(jobId).then((j) => {
      setJob(j);
      if (j.models.length) setActiveModel(j.models[0].url);
    }).catch((e) => setError(e.message));
    api.quoteSheet(jobId).then(setQuote).catch(() => setQuote(null));
  }, [jobId]);

  useEffect(() => {
    load();
  }, [load]);

  const classify = async (mode: "rules" | "llm") => {
    setBusy(true);
    setError("");
    try {
      await api.classifyJob(jobId, mode);
      load();
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  };

  const uploadRaw = async (file: File) => {
    setBusy(true);
    try {
      await api.uploadFile(jobId, "raw", file);
      load();
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  };

  const uploadTo = async (subfolder: "models" | "documents" | "images", file: File) => {
    setBusy(true);
    try {
      await api.uploadFile(jobId, subfolder, file);
      load();
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  };

  if (error && !job) {
    return (
      <Layout title="Quote" subtitle={jobId}>
        <EmptyState title="Job not found" description={error} action={
          <Link to="/quotes"><Button variant="secondary"><ArrowLeft className="h-4 w-4" /> Back to Quotes</Button></Link>
        } />
      </Layout>
    );
  }

  if (!job) {
    return (
      <Layout title="Loading..." subtitle={jobId}>
        <Spinner label="Loading job..." />
      </Layout>
    );
  }

  const priceByIndex: Record<string, (typeof quote extends null ? never : QuoteSheet["line_items"][number])> = {};
  quote?.line_items.forEach((li) => (priceByIndex[li.index] = li));

  const analysis = job.job_analysis || {};

  return (
    <Layout
      title={job.display_name}
      subtitle={job.customer || job.job_id}
      actions={
        <div className="flex items-center gap-2">
          <Link to="/quotes">
            <Button variant="ghost"><ChevronLeft className="h-4 w-4" /> Quotes</Button>
          </Link>
          <Button variant="secondary" onClick={() => classify("rules")} disabled={busy || !job.has_raw_csv}>
            <RefreshCw className={busy ? "h-4 w-4 animate-spin" : "h-4 w-4"} /> Re-run AI
          </Button>
        </div>
      }
    >
      {error && (
        <div className="mb-4 rounded-xl border border-accent-rose/30 bg-accent-rose/10 px-4 py-2.5 text-sm text-accent-rose">
          {error}
        </div>
      )}

      {!job.has_raw_csv && (
        <Card className="mb-5 flex items-center justify-between p-4">
          <div>
            <div className="text-sm font-medium text-ink-100">No raw CAD export yet</div>
            <p className="text-xs text-ink-400">Upload XT_Export_CAD_Dimensions.csv to run the AI classifier.</p>
          </div>
          <Button onClick={() => rawFileRef.current?.click()}>
            <Upload className="h-4 w-4" /> Upload CSV
          </Button>
          <input
            ref={rawFileRef}
            type="file"
            accept=".csv"
            className="hidden"
            onChange={(e) => e.target.files?.[0] && uploadRaw(e.target.files[0])}
          />
        </Card>
      )}

      {/* Price summary strip */}
      <div className="mb-6 grid grid-cols-1 gap-4 sm:grid-cols-3">
        <Card className="p-5 sm:col-span-1">
          <div className="text-xs font-medium uppercase tracking-wider text-ink-400">Total Quote Price</div>
          <div className="mt-1.5 text-3xl font-bold text-ink-100">
            ${(quote?.total_price ?? 0).toLocaleString(undefined, { minimumFractionDigits: 2 })}
          </div>
          <div className="mt-1 text-xs text-ink-400">
            {quote?.quoted_part_count ?? 0} of {quote?.total_part_count ?? 0} parts priced
          </div>
          <p className="mt-2 text-[11px] text-ink-500">
            Placeholder rates -- edit in Settings to match real CMS pricing.
          </p>
        </Card>
        <Card className="p-5 sm:col-span-2">
          <div className="mb-2 flex items-center gap-2">
            <Sparkles className="h-4 w-4 text-brand-400" />
            <div className="text-xs font-medium uppercase tracking-wider text-ink-400">AI Job Analysis</div>
          </div>
          <div className="flex flex-wrap gap-2 text-xs text-ink-300">
            {analysis.stack_axis && <Badge>Stack axis: {analysis.stack_axis}</Badge>}
            {analysis.sequenced_latch_lock_base && <Badge tone="brand">Plate-sequenced / latch-lock base</Badge>}
          </div>
          <p className="mt-2 text-xs leading-relaxed text-ink-400">{analysis.parting_line}</p>
        </Card>
      </div>

      {/* Tabs */}
      <div className="mb-5 flex gap-1 overflow-x-auto rounded-xl border border-ink-700/60 bg-ink-850/40 p-1">
        {[
          { id: "overview", label: "Overview", icon: Layers },
          { id: "parts", label: "Parts & Pricing", icon: FileText },
          { id: "images", label: `Images (${job.images.length})`, icon: ImageIcon },
          { id: "model", label: `3D Model (${job.models.length})`, icon: Box },
          { id: "documents", label: `Documents (${job.documents.length})`, icon: FileText },
        ].map(({ id, label, icon: Icon }) => (
          <button
            key={id}
            onClick={() => setTab(id as Tab)}
            className={`flex items-center gap-1.5 whitespace-nowrap rounded-lg px-3.5 py-2 text-sm font-medium transition ${
              tab === id ? "bg-ink-700/80 text-ink-100" : "text-ink-400 hover:text-ink-200"
            }`}
          >
            <Icon className="h-4 w-4" /> {label}
          </button>
        ))}
      </div>

      {tab === "overview" && (
        <div className="grid grid-cols-1 gap-5 lg:grid-cols-2">
          <Card className="p-5">
            <h3 className="mb-3 text-sm font-semibold text-ink-100">Rules Applied To This Job</h3>
            <ul className="space-y-2">
              {(analysis.rules_for_this_job || []).map((r, i) => (
                <li key={i} className="flex gap-2 text-xs leading-relaxed text-ink-300">
                  <span className="mt-1 h-1.5 w-1.5 shrink-0 rounded-full bg-brand-400" />
                  {r}
                </li>
              ))}
              {!analysis.rules_for_this_job?.length && (
                <p className="text-xs text-ink-500">Run the classifier to see AI reasoning here.</p>
              )}
            </ul>
          </Card>
          <Card className="p-5">
            <h3 className="mb-3 text-sm font-semibold text-ink-100">Module6121 AI Bridge</h3>
            <p className="mb-3 text-xs leading-relaxed text-ink-400">
              These exports carry AI-resolved part names/roles for your VBA macro to read and
              fill into the quoting workbook -- no re-classification needed on the macro side.
            </p>
            <div className="flex flex-wrap gap-2">
              <a href={api.bridgeCsvUrl(job.job_id)} download>
                <Button variant="secondary"><Download className="h-4 w-4" /> Download bridge CSV</Button>
              </a>
              <Button variant="ghost" onClick={() => window.open(`/api/bridge/${job.job_id}`, "_blank")}>
                View bridge JSON
              </Button>
            </div>
          </Card>
        </div>
      )}

      {tab === "parts" && <PartsTable parts={job.parts} prices={priceByIndex} />}

      {tab === "images" && (
        <div>
          {job.images.length === 0 ? (
            <EmptyState icon={<ImageIcon className="h-8 w-8" />} title="No rendered views yet"
              description="Upload JPEG/PNG assembly views for this job."
              action={
                <>
                  <Button onClick={() => docFileRef.current?.click()}><Upload className="h-4 w-4" /> Upload image</Button>
                  <input ref={docFileRef} type="file" accept="image/*" className="hidden"
                    onChange={(e) => e.target.files?.[0] && uploadTo("images", e.target.files[0])} />
                </>
              } />
          ) : (
            <ImageGallery images={job.images} />
          )}
        </div>
      )}

      {tab === "model" && (
        <div>
          {job.models.length === 0 ? (
            <EmptyState icon={<Box className="h-8 w-8" />} title="No STL model uploaded"
              description="Upload a .stl file to preview it in 3D."
              action={
                <>
                  <Button onClick={() => modelFileRef.current?.click()}><Upload className="h-4 w-4" /> Upload STL</Button>
                  <input ref={modelFileRef} type="file" accept=".stl" className="hidden"
                    onChange={(e) => e.target.files?.[0] && uploadTo("models", e.target.files[0])} />
                </>
              } />
          ) : (
            <div>
              {job.models.length > 1 && (
                <div className="mb-3 flex flex-wrap gap-2">
                  {job.models.map((m) => (
                    <button key={m.url} onClick={() => setActiveModel(m.url)}
                      className={`rounded-lg px-3 py-1.5 text-xs font-medium ${activeModel === m.url ? "bg-brand-500 text-white" : "bg-ink-800 text-ink-300"}`}>
                      {m.name}
                    </button>
                  ))}
                </div>
              )}
              <div className="h-[560px]">
                {activeModel && (
                  <Suspense fallback={<div className="flex h-full items-center justify-center"><Spinner label="Loading 3D viewer..." /></div>}>
                    <StlViewer url={activeModel} />
                  </Suspense>
                )}
              </div>
              <div className="mt-3">
                <Button variant="ghost" onClick={() => modelFileRef.current?.click()}><Upload className="h-4 w-4" /> Upload another STL</Button>
                <input ref={modelFileRef} type="file" accept=".stl" className="hidden"
                  onChange={(e) => e.target.files?.[0] && uploadTo("models", e.target.files[0])} />
              </div>
            </div>
          )}
        </div>
      )}

      {tab === "documents" && (
        <div>
          {job.documents.length === 0 ? (
            <EmptyState icon={<FileText className="h-8 w-8" />} title="No documents yet"
              description="Upload a quote sheet, steel sheet, or spec PDF for this job."
              action={
                <>
                  <Button onClick={() => docFileRef.current?.click()}><Upload className="h-4 w-4" /> Upload document</Button>
                  <input ref={docFileRef} type="file" className="hidden"
                    onChange={(e) => e.target.files?.[0] && uploadTo("documents", e.target.files[0])} />
                </>
              } />
          ) : (
            <div className="space-y-2">
              {job.documents.map((d) => (
                <a key={d.name} href={d.url} target="_blank" rel="noreferrer"
                  className="flex items-center justify-between rounded-xl border border-ink-700/50 bg-ink-850/40 px-4 py-3 hover:border-brand-500/40">
                  <span className="flex items-center gap-2 text-sm text-ink-100"><FileText className="h-4 w-4 text-ink-400" /> {d.name}</span>
                  <span className="text-xs text-ink-500">{(d.size / 1024).toFixed(1)} KB</span>
                </a>
              ))}
              <Button variant="ghost" onClick={() => docFileRef.current?.click()}><Upload className="h-4 w-4" /> Upload another</Button>
              <input ref={docFileRef} type="file" className="hidden"
                onChange={(e) => e.target.files?.[0] && uploadTo("documents", e.target.files[0])} />
            </div>
          )}
        </div>
      )}
    </Layout>
  );
}
