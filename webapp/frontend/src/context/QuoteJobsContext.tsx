import {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useMemo,
  useState,
  type ReactNode,
} from "react";
import { useNavigate } from "react-router-dom";
import { api, type QuoteRunStatus } from "../api/client";

export interface ActiveQuoteJob {
  quoteId: string;
  label: string;
  status: QuoteRunStatus;
  startedAt: number;
}

interface QuoteJobsContextValue {
  jobs: ActiveQuoteJob[];
  startQuote: (quoteId: string, label: string, initialStatus?: QuoteRunStatus) => void;
  cancelQuote: (quoteId: string) => Promise<void>;
  dismissQuote: (quoteId: string) => void;
  openQuote: (jobId: string) => void;
}

const QuoteJobsContext = createContext<QuoteJobsContextValue | null>(null);

const STORAGE_KEY = "cms_active_quotes";

function loadStored(): ActiveQuoteJob[] {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw) as ActiveQuoteJob[];
    return parsed.filter((j) => Date.now() - j.startedAt < 24 * 60 * 60 * 1000);
  } catch {
    return [];
  }
}

export function QuoteJobsProvider({ children }: { children: ReactNode }) {
  const [jobs, setJobs] = useState<ActiveQuoteJob[]>(loadStored);
  const navigate = useNavigate();

  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(jobs));
  }, [jobs]);

  const startQuote = useCallback(
    (quoteId: string, label: string, initialStatus?: QuoteRunStatus) => {
      setJobs((prev) => {
        const existing = prev.find((j) => j.quoteId === quoteId);
        if (existing) {
          return prev.map((j) =>
            j.quoteId === quoteId
              ? { ...j, label, status: initialStatus || j.status }
              : j
          );
        }
        return [
          {
            quoteId,
            label,
            status: initialStatus || { phase: "queued", message: "Starting quote pipeline..." },
            startedAt: Date.now(),
          },
          ...prev,
        ];
      });
    },
    []
  );

  const cancelQuote = useCallback(async (quoteId: string) => {
    try {
      const st = await api.cancelQuote(quoteId);
      setJobs((prev) =>
        prev.map((j) =>
          j.quoteId === quoteId
            ? { ...j, status: { ...st, phase: "cancelled", message: st.message || "Quote cancelled" } }
            : j
        )
      );
    } catch {
      setJobs((prev) =>
        prev.map((j) =>
          j.quoteId === quoteId
            ? { ...j, status: { phase: "cancelled", message: "Quote cancelled" } }
            : j
        )
      );
    }
  }, []);

  const dismissQuote = useCallback((quoteId: string) => {
    setJobs((prev) => prev.filter((j) => j.quoteId !== quoteId));
  }, []);

  const openQuote = useCallback(
    (jobId: string) => {
      navigate(`/quotes/${encodeURIComponent(jobId)}`);
    },
    [navigate]
  );

  useEffect(() => {
    const active = jobs.filter(
      (j) => !["completed", "error", "cancelled"].includes(j.status.phase)
    );
    if (active.length === 0) return;

    const iv = setInterval(async () => {
      for (const job of active) {
        try {
          const st = await api.quoteStatus(job.quoteId);
          setJobs((prev) =>
            prev.map((j) => (j.quoteId === job.quoteId ? { ...j, status: st } : j))
          );
        } catch {
          /* keep polling */
        }
      }
    }, 5000);
    return () => clearInterval(iv);
  }, [jobs]);

  const value = useMemo(
    () => ({ jobs, startQuote, cancelQuote, dismissQuote, openQuote }),
    [jobs, startQuote, cancelQuote, dismissQuote, openQuote]
  );

  return <QuoteJobsContext.Provider value={value}>{children}</QuoteJobsContext.Provider>;
}

export function useQuoteJobs() {
  const ctx = useContext(QuoteJobsContext);
  if (!ctx) throw new Error("useQuoteJobs must be used within QuoteJobsProvider");
  return ctx;
}
