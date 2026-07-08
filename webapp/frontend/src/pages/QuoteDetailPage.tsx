import { lazy, Suspense, useCallback, useEffect, useRef, useState } from "react";
import { useParams, Link } from "react-router-dom";
import {
  ArrowLeft, RefreshCw, Download, Upload, Image as ImageIcon, Box,
  FileText, Layers, ChevronLeft,
} from "lucide-react";
import Layout from "../components/Layout";
import { Card, Button, Spinner, EmptyState } from "../components/ui";
import PartsTable from "../components/PartsTable";
import ImageGallery from "../components/ImageGallery";
import { api, type JobDetail, type QuoteSheet } from "../api/client";

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

  const classify = async () => {
    setBusy(true);
    setError("");
    try {
      await api.classifyJob(jobId, "rules");
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
        <EmptyState title="Not found" description={error} action={
          <Link to="/quotes"><Button variant="secondary"><ArrowLeft className="h-4 w-4" /> Back</Button></Link>
        } />
      </Layout>
    );
  }

  if (!job) {
    return (
      <Layout title="Loading" subtitle={jobId}>
        <Spinner label="Loading quote..." />
      </Layout>
    );
  }

  const priceByIndex: Record<string, QuoteSheet["line_items"][number]> = {};
  quote?.line_items.forEach((li) => (priceByIndex[li.index] = li));

  return (
    <Layout
      title={job.job_id}
      subtitle={job.display_name}
      actions={
        <div className="flex items-center gap-2">
          <Link to="/quotes">
            <Button variant="ghost"><ChevronLeft className="h-4 w-4" /> Quotes</Button>
          </Link>
          {job.base_type !== "bms" && job.has_raw_csv && (
            <Button variant="secondary" onClick={classify} disabled={busy}>
              <RefreshCw className={busy ? "h-4 w-4 animate-spin" : "h-4 w-4"} /> Refresh Parts
            </Button>
          )}
        </div>
      }
    >
      {error && (
        <div className="mb-4 border border-accent-rose/40 px-4 py-2.5 text-sm text-accent-rose">{error}</div>
      )}

      {job.base_type === "bms" && (
        <div className="mb-4 border border-accent-amber/30 px-4 py-3 text-xs text-accent-amber">
          BMS base — BOM-driven by Module6121. AI classification disabled.
        </div>
      )}

      {!job.has_raw_csv && (
        <Card className="mb-5 flex items-center justify-between p-4">
          <div>
            <div className="section-label">CAD Export</div>
            <p className="mt-1 text-xs text-ink-400">Upload XT_Export_CAD_Dimensions.csv from the job folder.</p>
          </div>
          <Button onClick={() => rawFileRef.current?.click()}>
            <Upload className="h-4 w-4" /> Upload CSV
          </Button>
          <input ref={rawFileRef} type="file" accept=".csv" className="hidden"
            onChange={(e) => e.target.files?.[0] && uploadRaw(e.target.files[0])} />
        </Card>
      )}

      <div className="mb-6 grid grid-cols-1 gap-4 sm:grid-cols-3">
        <Card className="p-5 sm:col-span-1">
          <div className="section-label">Total Price</div>
          <div className="mt-2 text-4xl font-light tracking-tight text-ink-100">
            ${(quote?.total_price ?? 0).toLocaleString(undefined, { minimumFractionDigits: 2 })}
          </div>
          <div className="mt-2 text-xs text-ink-400">
            {quote?.csv_priced_count ?? 0} priced from CSV
            {(quote?.missing_csv_price_count ?? 0) > 0 && (
              <span className="text-accent-amber"> · {quote?.missing_csv_price_count} need CSV price</span>
            )}
          </div>
          <p className="mt-3 text-[10px] leading-relaxed text-ink-500">
            {quote?.pricing_source || "Purchased Components Prices.csv"}
            {quote?.has_steel_sheet_dims && " · Steel sheet dimensions applied"}
          </p>
        </Card>
        <Card className="p-5 sm:col-span-2">
          <div className="section-label">Quote Summary</div>
          <div className="mt-3 grid grid-cols-2 gap-4 text-xs sm:grid-cols-4">
            <div>
              <div className="text-ink-500">C-Number</div>
              <div className="mt-1 text-ink-100">{job.job_id}</div>
            </div>
            <div>
              <div className="text-ink-500">Parts</div>
              <div className="mt-1 text-ink-100">{quote?.total_part_count ?? 0}</div>
            </div>
            <div>
              <div className="text-ink-500">Quoted</div>
              <div className="mt-1 text-ink-100">{quote?.quoted_part_count ?? 0}</div>
            </div>
            <div>
              <div className="text-ink-500">Base Type</div>
              <div className="mt-1 text-ink-100">{job.base_type === "bms" ? "BMS" : "Standard"}</div>
            </div>
          </div>
        </Card>
      </div>

      <div className="mb-5 flex gap-0 overflow-x-auto border border-ink-700">
        {[
          { id: "overview", label: "Overview", icon: Layers },
          { id: "parts", label: "Parts & Pricing", icon: FileText },
          { id: "images", label: `Images (${job.images.length})`, icon: ImageIcon },
          { id: "model", label: `3D (${job.models.length})`, icon: Box },
          { id: "documents", label: `Docs (${job.documents.length})`, icon: FileText },
        ].map(({ id, label, icon: Icon }) => (
          <button
            key={id}
            onClick={() => setTab(id as Tab)}
            className={`flex items-center gap-2 whitespace-nowrap border-r border-ink-700 px-4 py-3 text-[10px] font-semibold uppercase tracking-widest transition last:border-r-0 ${
              tab === id ? "bg-ink-100 text-ink-950" : "text-ink-400 hover:text-ink-100"
            }`}
          >
            <Icon className="h-3.5 w-3.5" /> {label}
          </button>
        ))}
      </div>

      {tab === "overview" && (
        <Card className="p-5">
          <div className="section-label mb-3">Module6121 Bridge</div>
          <p className="mb-4 text-xs leading-relaxed text-ink-400">
            Part names export to the macro at 127.0.0.1 after classification.
          </p>
          <div className="flex flex-wrap gap-2">
            <a href={api.bridgeCsvUrl(job.job_id)} download>
              <Button variant="secondary"><Download className="h-4 w-4" /> Bridge CSV</Button>
            </a>
          </div>
        </Card>
      )}

      {tab === "parts" && <PartsTable parts={job.parts} prices={priceByIndex} />}

      {tab === "images" && (
        job.images.length === 0 ? (
          <EmptyState icon={<ImageIcon className="h-8 w-8" />} title="No images"
            action={<Button onClick={() => docFileRef.current?.click()}><Upload className="h-4 w-4" /> Upload</Button>} />
        ) : <ImageGallery images={job.images} />
      )}

      {tab === "model" && (
        job.models.length === 0 ? (
          <EmptyState icon={<Box className="h-8 w-8" />} title="No STL"
            action={<Button onClick={() => modelFileRef.current?.click()}><Upload className="h-4 w-4" /> Upload STL</Button>} />
        ) : (
          <div>
            <div className="h-[560px]">
              {activeModel && (
                <Suspense fallback={<Spinner label="Loading 3D..." />}>
                  <StlViewer url={activeModel} />
                </Suspense>
              )}
            </div>
          </div>
        )
      )}

      {tab === "documents" && (
        job.documents.length === 0 ? (
          <EmptyState icon={<FileText className="h-8 w-8" />} title="No documents"
            action={<Button onClick={() => docFileRef.current?.click()}><Upload className="h-4 w-4" /> Upload</Button>} />
        ) : (
          <div className="space-y-2">
            {job.documents.map((d) => (
              <a key={d.name} href={d.url} target="_blank" rel="noreferrer"
                className="flex items-center justify-between border border-ink-700 px-4 py-3 hover:border-ink-100">
                <span className="flex items-center gap-2 text-sm text-ink-100"><FileText className="h-4 w-4 text-ink-500" /> {d.name}</span>
                <span className="text-xs text-ink-500">{(d.size / 1024).toFixed(1)} KB</span>
              </a>
            ))}
          </div>
        )
      )}

      <input ref={docFileRef} type="file" className="hidden"
        onChange={(e) => e.target.files?.[0] && uploadTo("documents", e.target.files[0])} />
      <input ref={modelFileRef} type="file" accept=".stl" className="hidden"
        onChange={(e) => e.target.files?.[0] && uploadTo("models", e.target.files[0])} />
    </Layout>
  );
}
