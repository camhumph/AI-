const BASE = "/api";

export interface JobSummary {
  job_id: string;
  display_name: string;
  customer: string;
  notes: string;
  has_raw_csv: boolean;
  has_classification: boolean;
  part_count: number;
  image_count: number;
  model_count: number;
  sequenced_latch_lock_base: boolean;
  updated_at: string;
}

export interface PartRow {
  index: string;
  role: string;
  role_label: string;
  role_group: string;
  confidence: string;
  reason: string;
  quote: boolean;
  Component: string;
  Thickness: string;
  Width: string;
  Length: string;
  CenterX: string;
  CenterY: string;
  CenterZ: string;
}

export interface JobAnalysis {
  stack_axis?: string;
  parting_line?: string;
  sequenced_latch_lock_base?: boolean;
  rules_for_this_job?: string[];
}

export interface AssetRef {
  name: string;
  url: string;
  size: number;
}

export interface JobDetail {
  job_id: string;
  display_name: string;
  customer: string;
  notes: string;
  job_analysis: JobAnalysis;
  parts: PartRow[];
  images: AssetRef[];
  models: AssetRef[];
  documents: AssetRef[];
  has_raw_csv: boolean;
}

export interface QuoteLineItem {
  index: string;
  component: string;
  role: string;
  role_label: string;
  role_group: string;
  confidence: string;
  quote: boolean;
  price: number;
}

export interface QuoteSheet {
  job_id: string;
  line_items: QuoteLineItem[];
  total_price: number;
  quoted_part_count: number;
  total_part_count: number;
}

export interface EmailSummary {
  id: string;
  from: string;
  subject: string;
  date: string;
  job_tokens: string[];
  matched_jobs: string[];
}

export interface EmailDetail extends EmailSummary {
  to: string;
  body_text: string;
  body_html: string;
  message_id_header: string;
  attachments: { filename: string; content_type: string; size: number }[];
}

async function req<T>(path: string, options?: RequestInit): Promise<T> {
  const res = await fetch(`${BASE}${path}`, {
    headers: { "Content-Type": "application/json" },
    ...options,
  });
  if (!res.ok) {
    let detail = res.statusText;
    try {
      const body = await res.json();
      detail = body.detail || detail;
    } catch {
      /* ignore */
    }
    const err = new Error(detail) as Error & { status?: number };
    err.status = res.status;
    throw err;
  }
  return res.json();
}

export const api = {
  health: () => req<{ ok: boolean; email_configured: boolean; smtp_configured: boolean }>("/health"),

  listJobs: () => req<JobSummary[]>("/jobs"),
  getJob: (jobId: string) => req<JobDetail>(`/jobs/${encodeURIComponent(jobId)}`),
  createJob: (job_id: string, display_name: string, customer: string) =>
    req<JobDetail>("/jobs", { method: "POST", body: JSON.stringify({ job_id, display_name, customer }) }),
  classifyJob: (jobId: string, mode: "rules" | "llm" = "rules") =>
    req<JobDetail>(`/jobs/${encodeURIComponent(jobId)}/classify`, {
      method: "POST",
      body: JSON.stringify({ mode }),
    }),
  quoteSheet: (jobId: string) => req<QuoteSheet>(`/jobs/${encodeURIComponent(jobId)}/quote-sheet`),

  uploadFile: async (jobId: string, subfolder: "raw" | "images" | "models" | "documents", file: File) => {
    const form = new FormData();
    form.append("file", file);
    const res = await fetch(
      `${BASE}/jobs/${encodeURIComponent(jobId)}/upload?subfolder=${subfolder}`,
      { method: "POST", body: form }
    );
    if (!res.ok) throw new Error((await res.json()).detail || res.statusText);
    return res.json() as Promise<JobDetail>;
  },

  getPricing: () => req<Record<string, { mode: string; rate: number; minimum: number }>>("/pricing"),
  putPricing: (rates: Record<string, { mode: string; rate: number; minimum: number }>) =>
    req("/pricing", { method: "PUT", body: JSON.stringify(rates) }),

  bridgeJson: (jobId: string) => req(`/bridge/${encodeURIComponent(jobId)}`),
  bridgeCsvUrl: (jobId: string) => `${BASE}/bridge/${encodeURIComponent(jobId)}/csv`,

  emailStatus: () =>
    req<{ configured: boolean; smtp_configured: boolean; imap_host: string | null; imap_user: string | null }>(
      "/email/status"
    ),
  listEmails: () => req<EmailSummary[]>("/email/messages"),
  getEmail: (id: string) => req<EmailDetail>(`/email/messages/${encodeURIComponent(id)}`),
  replyEmail: (id: string, to: string, subject: string, body: string, in_reply_to = "") =>
    req(`/email/messages/${encodeURIComponent(id)}/reply`, {
      method: "POST",
      body: JSON.stringify({ to, subject, body, in_reply_to }),
    }),
};
