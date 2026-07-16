import { CheckCircle2, Loader2, X, AlertCircle, ExternalLink, Square, Trash2 } from "lucide-react";
import { useQuoteJobs } from "../context/QuoteJobsContext";
import { api } from "../api/client";

const PHASE_LABEL: Record<string, string> = {
  queued: "Preparing",
  starting: "DME prices",
  launching: "SolidWorks",
  running: "Module6121 + AI",
  completed: "Complete",
  cancelled: "Cancelled",
  error: "Failed",
};

export default function BackgroundQuoteBar() {
  const { jobs, cancelQuote, dismissQuote, openQuote } = useQuoteJobs();
  if (jobs.length === 0) return null;

  const deleteQuote = async (quoteId: string, jobId?: string) => {
    if (!window.confirm(`Delete quote ${jobId || quoteId}?\n\nRemoves it from the webapp list.`)) return;
    try {
      await api.deleteQuote(quoteId);
      if (jobId && jobId !== quoteId) {
        try {
          await api.deleteJob(jobId);
        } catch {
          /* already gone */
        }
      }
    } catch {
      /* still dismiss from bar */
    }
    dismissQuote(quoteId);
  };

  return (
    <div className="fixed bottom-5 right-5 z-50 flex w-full max-w-sm flex-col gap-2 sm:max-w-md">
      {jobs.map((job) => {
        const phase = job.status.phase;
        const done = phase === "completed";
        const failed = phase === "error";
        const cancelled = phase === "cancelled";
        const active = !done && !failed && !cancelled;
        const stuck =
          job.status.stuck_reason ||
          job.status.diagnostics?.stuck_reason ||
          "";
        const logTail =
          job.status.diagnostics?.launcher_log_tail ||
          job.status.diagnostics?.job_log_tail ||
          "";

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
              ) : cancelled ? (
                <Square className="h-4 w-4 text-ink-500" />
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
                <div className="mt-1 line-clamp-3 text-[10px] text-ink-500">{job.status.message}</div>
              )}
              {job.status.warning && (
                <div
                  className={`mt-1 line-clamp-3 text-[10px] font-semibold ${
                    job.status.cad_job_mismatch ||
                    /different job-number|differs from folder job|does not match handoff/i.test(
                      job.status.warning,
                    )
                      ? "text-accent-rose"
                      : "text-accent-amber"
                  }`}
                >
                  {job.status.warning}
                </div>
              )}
              {stuck && (
                <div className="mt-1 whitespace-pre-wrap break-words text-[10px] text-accent-rose">
                  {stuck}
                </div>
              )}
              {logTail && (active || failed) && (
                <pre className="mt-2 max-h-28 overflow-auto whitespace-pre-wrap break-words rounded-lg bg-black/30 p-2 font-mono text-[9px] leading-snug text-ink-400">
                  {logTail}
                </pre>
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
            <div className="flex shrink-0 flex-col gap-1">
              {active && (
                <button
                  onClick={() => cancelQuote(job.quoteId)}
                  className="rounded-full p-1.5 text-accent-rose/80 hover:bg-accent-rose/10 hover:text-accent-rose"
                  title="Cancel quote"
                >
                  <Square className="h-3.5 w-3.5" />
                </button>
              )}
              <button
                onClick={() => deleteQuote(job.quoteId, job.status.job_id)}
                className="rounded-full p-1.5 text-ink-500 hover:bg-accent-rose/10 hover:text-accent-rose"
                title="Delete quote"
              >
                <Trash2 className="h-3.5 w-3.5" />
              </button>
              <button
                onClick={() => dismissQuote(job.quoteId)}
                className="rounded-full p-1.5 text-ink-500 hover:bg-white/10 hover:text-ink-200"
                title="Dismiss"
              >
                <X className="h-3.5 w-3.5" />
              </button>
            </div>
          </div>
        );
      })}
    </div>
  );
}
