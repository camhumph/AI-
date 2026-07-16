const BASE = "/api";

export interface JobSummary {
  job_id: string;
  display_name: string;
  customer: string;
  notes: string;
  base_type: "standard" | "bms";
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
  base_type: "standard" | "bms";
  job_analysis: JobAnalysis;
  parts: PartRow[];
  images: AssetRef[];
  models: AssetRef[];
  documents: AssetRef[];
  has_raw_csv: boolean;
}

export interface QuoteLineItem {
  index: string;
  section?: "steel" | "pullcore" | "purchased" | "classified";
  component: string;
  role: string;
  role_label: string;
  role_group: string;
  confidence: string;
  quote: boolean;
  price: number;
  price_source?: string;
  thickness?: number | null;
  width?: number | null;
  length?: number | null;
  qty?: number | null;
  cu_in?: number | null;
  hours?: number | null;
  vendor?: string;
  part_number?: string;
  unit_price?: number | null;
  material?: string;
  category?: string;
}

export interface QuoteSummary {
  total_hours?: number | null;
  total_price_rough?: number | null;
  total_price_finish?: number | null;
  commission_pct?: number | null;
  commission_rough?: number | null;
  commission_finish?: number | null;
  grand_total_rough?: number | null;
  grand_total_finish?: number | null;
}

export interface QuoteSheet {
  job_id: string;
  line_items: QuoteLineItem[];
  sections?: {
    steel?: QuoteLineItem[];
    pullcore?: QuoteLineItem[];
    purchased?: QuoteLineItem[];
    classified?: QuoteLineItem[];
  };
  steel_plates?: QuoteLineItem[];
  pullcore_components?: QuoteLineItem[];
  purchased_components?: QuoteLineItem[];
  summary?: QuoteSummary;
  total_price: number;
  section_total_price?: number;
  quoted_part_count: number;
  total_part_count: number;
  csv_priced_count?: number;
  missing_csv_price_count?: number;
  pricing_source?: string;
  shop_csv?: string;
  has_steel_sheet_dims?: boolean;
}

export interface WorkspaceEntry {
  name: string;
  path: string;
  is_dir: boolean;
  c_number: string | null;
  has_xt_csv: boolean;
  has_quote_sheet: boolean;
  has_steel_sheet: boolean;
  quote_ready: boolean;
}

export interface WorkspaceBrowse {
  path: string;
  exists: boolean;
  parent: string | null;
  entries: WorkspaceEntry[];
  roots: string[];
}

export interface EmailSummary {
  id: string;
  from: string;
  from_name?: string;
  from_addr?: string;
  subject: string;
  date: string;
  snippet?: string;
  seen?: boolean;
  starred?: boolean;
  job_tokens: string[];
  matched_jobs: string[];
}

export interface EmailSettings {
  imap_host: string;
  imap_port: number;
  imap_user: string;
  imap_password_set: boolean;
  imap_folder: string;
  imap_ssl: boolean;
  smtp_host: string;
  smtp_port: number;
  smtp_user: string;
  smtp_password_set: boolean;
  smtp_from: string;
  gmail_address: string;
  configured: boolean;
  smtp_configured: boolean;
  credentials_path: string;
}

export interface QuoteRunStatus {
  phase: string;
  message?: string;
  warning?: string;
  cad_job_mismatch?: boolean;
  stuck_reason?: string;
  diagnostics?: {
    stuck_reason?: string;
    launcher_last_step?: string;
    launcher_log_tail?: string;
    macro_status?: string;
    macro_error_text?: string;
    job_log_tail?: string;
    macro_started?: boolean;
    macro_done?: boolean;
    macro_error?: boolean;
    cad_job_mismatch?: boolean;
    cad_job_mismatch_text?: string;
    warning?: string;
  };
  job_id?: string;
  c_number?: string;
  quote_id?: string;
  cad_path?: string;
  local_folder?: string;
}

export interface QuoteEmailResult {
  job_id: string;
  quote_id?: string;
  subject: string;
  cust_job: string;
  attachments_saved: number;
  attach_dir: string;
  launcher_started: boolean;
  email_handoff: string;
  poll_url?: string;
}

export interface EmailDetail extends EmailSummary {
  to: string;
  cc?: string;
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

export interface TrainingSuggestion {
  priority: string;
  role: string;
  occurrences: number;
  suggestion: string;
  examples: string;
  action: string;
}

export interface TrainingJobResult {
  job_id: string;
  folder?: string;
  base_type?: string;
  status?: string;
  rules_accuracy_pct?: number;
  accuracy_pct?: number;
  components_matched?: number;
  total_components?: number;
  qwen_accuracy_pct?: number;
  qwen_ran?: boolean;
  qwen_elapsed_sec?: number;
  xt_export?: { ok?: boolean; status?: string; reason?: string; message?: string; part_count?: string };
  macro_guidance?: string;
  detection_signals?: string[];
  reason?: string;
}

export interface TrainingReport {
  jobs_processed?: number;
  jobs_completed?: number;
  jobs_ok?: number;
  jobs_skipped?: number;
  bms_jobs?: number;
  standard_jobs?: number;
  overall_rules_accuracy_pct?: number;
  overall_qwen_accuracy_pct?: number;
  use_qwen?: boolean;
  qwen_model?: string;
  export_xt?: boolean;
  xt_exported_jobs?: number;
  running?: boolean;
  phase?: string;
  current_job?: string;
  job_index?: number;
  job_total?: number;
  message?: string;
  detail?: string;
  error?: string;
  cancelled?: boolean;
  qwen_thinking?: boolean;
  qwen_elapsed_sec?: number;
  qwen_live_output?: string;
  elapsed_sec?: number;
  started_at?: string;
  updated_at?: string;
  background?: boolean;
  started?: boolean;
  results?: TrainingJobResult[];
  suggestions?: TrainingSuggestion[];
  output_dir?: string;
  jobs_root?: string;
  disagreements_csv?: string;
  disagreements_md?: string;
}

export interface TrainingSuggestions {
  markdown: string;
  suggestions: TrainingSuggestion[];
  overall_rules_accuracy_pct: number;
  jobs_processed: number;
  bms_jobs: number;
  standard_jobs: number;
}

export const api = {
  health: () => req<{ ok: boolean; email_configured: boolean; smtp_configured: boolean }>("/health"),

  listJobs: () => req<JobSummary[]>("/jobs"),
  deleteJob: (jobId: string) =>
    req<{ deleted: boolean; job_id: string }>(`/jobs/${encodeURIComponent(jobId)}`, { method: "DELETE" }),
  browseWorkspace: (path = "") => req<WorkspaceBrowse>(`/workspace/browse?path=${encodeURIComponent(path)}`),
  importFromFolder: (folder_path: string, run_quote = true) =>
    req<JobDetail & { quote_started?: boolean; quote_id?: string; poll_url?: string }>("/jobs/import-folder", {
      method: "POST",
      body: JSON.stringify({ folder_path, run_quote }),
    }),
  importFoldersBatch: (folder_paths: string[], run_quote = true) =>
    req<{
      launched: boolean;
      batch?: boolean;
      batch_count?: number;
      quote_ids: string[];
      c_numbers?: string[];
      jobs?: { job_id: string; quote_id: string; folder_path: string; display_name: string }[];
      errors?: string[];
      error?: string | null;
    }>("/jobs/import-folders-batch", {
      method: "POST",
      body: JSON.stringify({ folder_paths, run_quote }),
    }),
  quoteStatus: (quoteId: string) => req<QuoteRunStatus>(`/quote/status/${encodeURIComponent(quoteId)}`),
  cancelQuote: (quoteId: string) =>
    req<QuoteRunStatus>(`/quote/cancel/${encodeURIComponent(quoteId)}`, { method: "POST", body: "{}" }),
  deleteQuote: (quoteId: string) =>
    req<{ deleted: boolean; quote_id: string; job_deleted?: boolean }>(
      `/quote/delete/${encodeURIComponent(quoteId)}`,
      { method: "POST", body: "{}" }
    ),
  activeQuotes: () => req<QuoteRunStatus[]>("/quote/active"),
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
  getEmailSettings: () => req<EmailSettings>("/settings/email"),
  putEmailSettings: (settings: Partial<EmailSettings> & { imap_password?: string; smtp_password?: string }) =>
    req<EmailSettings>("/settings/email", { method: "PUT", body: JSON.stringify(settings) }),
  testEmail: () => req<{ ok: boolean; message: string }>("/settings/email/test", { method: "POST" }),
  listEmails: (q = "") => req<EmailSummary[]>(`/email/messages${q ? `?q=${encodeURIComponent(q)}` : ""}`),
  getEmail: (id: string) => req<EmailDetail>(`/email/messages/${encodeURIComponent(id)}`),
  quoteEmail: (id: string, launchMacro = true) =>
    req<QuoteEmailResult>(`/email/messages/${encodeURIComponent(id)}/quote`, {
      method: "POST",
      body: JSON.stringify({ launch_macro: launchMacro }),
    }),
  quoteEmailBatch: (messageIds: string[], launchMacro = true) =>
    req<{
      launched?: boolean;
      batch?: boolean;
      batch_count?: number;
      quote_ids?: string[];
      c_numbers?: string[];
      error?: string;
      results?: QuoteEmailResult[];
      macro_started?: boolean;
    }>("/email/quote-batch", {
      method: "POST",
      body: JSON.stringify({ message_ids: messageIds, launch_macro: launchMacro }),
    }),
  replyEmail: (id: string, to: string, subject: string, body: string, in_reply_to = "") =>
    req(`/email/messages/${encodeURIComponent(id)}/reply`, {
      method: "POST",
      body: JSON.stringify({ to, subject, body, in_reply_to }),
    }),
  replyAllEmail: (
    id: string,
    to_addrs: string[],
    cc_addrs: string[],
    subject: string,
    body: string,
    in_reply_to = ""
  ) =>
    req(`/email/messages/${encodeURIComponent(id)}/reply-all`, {
      method: "POST",
      body: JSON.stringify({ to_addrs, cc_addrs, subject, body, in_reply_to }),
    }),
  forwardEmail: (id: string, to: string, body = "") =>
    req(`/email/messages/${encodeURIComponent(id)}/forward`, {
      method: "POST",
      body: JSON.stringify({ to, body }),
    }),
  composeEmail: (to_addrs: string[], subject: string, body: string, cc_addrs: string[] = []) =>
    req("/email/compose", {
      method: "POST",
      body: JSON.stringify({ to_addrs, cc_addrs, subject, body }),
    }),
  markEmailRead: (id: string, read: boolean) =>
    req(`/email/messages/${encodeURIComponent(id)}/read`, {
      method: "PATCH",
      body: JSON.stringify({ read }),
    }),
  starEmail: (id: string, starred: boolean) =>
    req(`/email/messages/${encodeURIComponent(id)}/star`, {
      method: "PATCH",
      body: JSON.stringify({ starred }),
    }),
  deleteEmail: (id: string, permanent = false) =>
    req(`/email/messages/${encodeURIComponent(id)}?permanent=${permanent}`, { method: "DELETE" }),
  archiveEmail: (id: string) =>
    req(`/email/messages/${encodeURIComponent(id)}/archive`, { method: "POST" }),

  trainingStatus: () =>
    req<TrainingReport>("/training/status"),
  qwenLive: (tail = 12000) =>
    req<{ path: string; text: string; size: number; exists: boolean; error?: string }>(
      `/training/qwen-live?tail=${tail}`
    ),
  trainingSuggestions: () => req<TrainingSuggestions>("/training/suggestions"),
  cancelTraining: () =>
    req<TrainingReport>("/training/cancel", { method: "POST", body: "{}" }),
  runTraining: (jobsRoot?: string, useQwen = true, qwenModel = "qwen3.5:9b", exportXt = true) =>
    req<TrainingReport>("/training/run", {
      method: "POST",
      body: JSON.stringify({
        jobs_root: jobsRoot || null,
        scan: true,
        use_qwen: useQwen,
        qwen_model: qwenModel,
        export_xt: exportXt,
      }),
    }),
};
