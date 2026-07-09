import { CheckCircle2, Loader2, X, AlertCircle, ExternalLink } from "lucide-react";
import { useQuoteJobs } from "../context/QuoteJobsContext";

const PHASE_LABEL: Record<string, string> = {
  queued: "Preparing",
  starting: "DME prices",
  launching: "SolidWorks",
  running: "Module6121 + AI",
  completed: "Complete",
  error: "Failed",
};

export default function BackgroundQuoteBar() {
  const { jobs, dismissQuote, openQuote } = useQuoteJobs();
  if (jobs.length === 0) return null;

  return (
    <div className="fixed bottom-5 right-5 z-50 flex w-full max-w-sm flex-col gap-2 sm:max-w-md">
      {jobs.map((job) => {
        const phase = job.status.phase;
        const done = phase === "completed";
        const failed = phase === "error";

        return (
          <div
            key={job.quoteId}
            className="glass-panel-strong flex items-start gap-3 rounded-2xl px-4 py-3"
          >
            <div className="mt-0.5 shrink-0">
              {done ? (
                <CheckCircle2 className="h-4 w-4 text-accent-green" />
              ) : failed ? (
                <AlertCircle className="h-4 w-4 text-accent-rose" />
              ) : (
                <Loader2 className="h-4 w-4 animate-spin text-brand-400" />
              )}
            </div>
            <div className="min-w-0 flex-1">
              <div className="truncate text-xs font-semibold text-ink-200">{job.label}</div>
              <div className="mt-0.5 font-mono text-[10px] text-ink-500">
                {PHASE_LABEL[phase] || phase}
                {job.status.job_id ? ` · ${job.status.job_id}` : ""}
              </div>
              {job.status.message && (
                <div className="mt-1 line-clamp-2 text-[10px] text-ink-500">{job.status.message}</div>
              )}
              {done && job.status.job_id && (
                <button
                  onClick={() => openQuote(job.status.job_id!)}
                  className="mt-2 flex items-center gap-1 text-[10px] font-semibold uppercase tracking-wider text-brand-400 hover:text-brand-400/80"
                >
                  <ExternalLink className="h-3 w-3" /> Open quote
                </button>
              )}
            </div>
            <button
              onClick={() => dismissQuote(job.quoteId)}
              className="shrink-0 rounded-full p-1.5 text-ink-500 hover:bg-white/10 hover:text-ink-200"
              title="Dismiss"
            >
              <X className="h-3.5 w-3.5" />
            </button>
          </div>
        );
      })}
    </div>
  );
}
