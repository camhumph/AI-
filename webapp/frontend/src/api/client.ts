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
  /** Measured mass in lb from SolidWorks (CreateMassProperty). "" when unknown. */
  MassLb: string;
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
  /** Steel grade code (A36 / 4140 / P20 / ...). */
  grade?: string;
  /** Shop label for `grade`, e.g. "#2 4140". */
  grade_label?: string;
  /** "override" when the estimator set it, "default" when the macro did. */
  grade_source?: "override" | "default";
  /** Grades the quote workbook has rows for. */
  grade_options?: string[];
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
  /** User's rename for this plate, when set. `role` is unchanged by a rename. */
  name_override?: string | null;
  /** What the role would be called without the override, for the revert hint. */
  role_label_original?: string | null;
}

/** Result of saving plate renames. */
/**
 * What the app took away from a rename — see backend/app/name_learning.py.
 *
 * `learned` true means every future job of this base type now uses the new name
 * for this role, from this one example. `reason` is why the old name was wrong,
 * and `role_suspect` marks the case that is NOT a naming problem: a new name
 * with no word in common with the old one usually means the plate was
 * classified as the wrong kind of plate, and a rename never moves the role, the
 * quote row or the price.
 */
export interface RenameLesson {
  learned: boolean;
  role?: string;
  base_type?: string;
  reason?: string;
  why?: string;
  fix?: string;
  role_suspect?: boolean;
  not_learned_because?: string;
}

export interface SheetWriteReport {
  attempted?: number;
  changed?: number;
  ok?: boolean;
  skipped_reason?: string;
  /** One entry per .xls actually saved. A job has up to three copies of each. */
  files_changed?: string[];
  cells?: Array<{
    workbook: string;
    folder?: string;
    sheet: string;
    cell?: string;
    from?: string;
    to?: string;
  }>;
  errors?: string[];
}

export interface PlateRenameResult {
  job_id: string;
  names: Record<string, string>;
  renames: Array<{ from: string; to: string }>;
  sheets: SheetWriteReport;
  learning?: RenameLesson[];
}

export interface LearnedPlateName {
  base_type: string;
  role: string;
  label: string;
  replaces: string;
  reason: string;
  why: string;
  fix: string;
  role_suspect: boolean;
  samples: number;
  jobs: string[];
  updated_at: string;
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

export interface MachiningOperation {
  op: string;
  min: number;
  /** True when the minutes come from measured CAD geometry rather than an assumed pattern. */
  measured: boolean;
  detail: string;
}

export interface MachiningPart {
  component: string;
  role_label: string;
  material: string;
  qty: number;
  skipped: boolean;
  reason?: string;
  stock?: { thickness: number; width: number; length: number };
  stock_volume_cuin?: number;
  finished_volume_cuin?: number;
  finished_basis?: string;
  removed_cuin?: number;
  removed_pct?: number;
  operations?: MachiningOperation[];
  cut_min?: number;
  efficiency_factor?: number;
  per_part_min?: number;
  total_min: number;
  total_hours: number;
  /** Full Datum cost for this part x qty: machine, setup, stock, programming,
   *  inspection, cutter wear, deburr, plus tolerance risk. Mesh path only. */
  cost_usd?: number;
  /** True when the BOM stock is smaller than the measured part -- the quote row
   *  and the CAD disagree. Mesh path only. */
  stock_short?: boolean;
  /** Confidence band and the named reasons it is that wide. Mesh path only. */
  confidence?: QuoteConfidence;
  /** Optimistic end of the band, minutes x qty. */
  low_min?: number;
  /** Pessimistic end of the band, minutes x qty. */
  high_min?: number;
  /** How the roughing cascade was chosen, when it was expanded. */
  cascade_note?: string;
}

/**
 * Mirrors ConfidenceBand in lib/quoteConfidence.ts.
 *
 * Declared structurally rather than imported so the API surface stays free of a
 * dependency on the client-side estimator -- the Python endpoint can populate
 * the same shape.
 */
export interface QuoteConfidence {
  score: number;
  lowFactor: number;
  highFactor: number;
  grade: "A" | "B" | "C" | "D";
  drivers: string[];
  /**
   * Optional machine-readable form of `drivers`, used to aggregate them across
   * plates. Optional on purpose: a band populated by the Python endpoint may
   * supply only the rendered strings, and the roll-up falls back to passing those
   * through rather than dropping them.
   */
  driverDetail?: Array<{ kind: string; text: string; value?: number }>;
  guidance: string;
}

/** One hole group found in the mesh. */
export interface GeoHole {
  dia_in: number;
  depth_in: number;
  count: number;
  blind: boolean;
  /** bore | tapped | clearance | fit — from Datum's hole purpose. */
  purpose: string;
  /** UNC spec when the diameter matches a tap drill in the crib. */
  thread?: string;
  /** Depth / diameter. Above ~8xD needs deep-hole strategy. */
  ld: number;
  /**
   * The tool that makes this hole: carousel number when it is on this
   * machine, otherwise the right tool named and marked "another machine".
   * A "reach" note means no tool of that size gets deep enough anywhere,
   * which needs a peck strategy, an extension or gun drilling.
   */
  tool: string;
}

export interface GeoPocket {
  depth_in: number;
  area_sqin: number;
  count: number;
  /** True for a counterbore seat rather than a true pocket. */
  cbore: boolean;
}

export interface GeoBore {
  dia_in: number;
  count: number;
  method: string;
}

/** Everything the mesh analysis found on one plate, for a shop-floor review. */
export interface GeometryPart {
  role_label: string;
  component: string;
  material: string;
  stock_in?: { thickness: number; width: number; length: number };
  finished_volume_cuin?: number;
  holes: GeoHole[];
  pockets: GeoPocket[];
  bores: GeoBore[];
  taps: { spec: string; count: number; tool: string }[];
  chamfer_count: number;
  chamfer_len_in: number;
  radii_count: number;
  orientations: number;
  distinct_tools: number;
  /** Engraving text the macro stamps on this plate, when known. */
  engraving?: string[];
  skipped?: boolean;
  reason?: string;
}

export interface GeometryReport {
  parts: GeometryPart[];
  totals: {
    holes: number;
    tapped: number;
    pockets: number;
    bores: number;
    chamfer_len_in: number;
  };
  notes: string[];
}

export interface MachiningEstimate {
  parts: MachiningPart[];
  summary: {
    part_count: number;
    skipped_count: number;
    total_minutes: number;
    total_hours: number;
    shop_rate_per_hour: number;
    estimated_cost: number;
    measured_minutes: number;
    assumed_minutes: number;
    /** Share of the total that rests on measured geometry. Below ~60%, treat as rough order. */
    confidence_pct: number;
    /** Job-level band, rolled up from the plates. */
    confidence?: QuoteConfidence;
    /** Optimistic and pessimistic ends of the job total, minutes. */
    low_minutes?: number;
    high_minutes?: number;
    /** Cost at each end of the band. */
    low_cost?: number;
    high_cost?: number;
  };
  rates_used: Record<string, unknown>;
  notes: string[];
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

  // A 200 that is HTML means the SPA catch-all served index.html because the API
  // route did not match -- almost always a backend running older code than the
  // frontend. Say that, instead of letting JSON.parse throw
  // "Unexpected token '<', "<!doctype "... is not valid JSON", which tells the
  // user nothing about what to do.
  const ctype = res.headers.get("content-type") || "";
  if (!ctype.includes("json")) {
    const peek = (await res.text()).slice(0, 40).replace(/\s+/g, " ");
    throw new Error(
      `${path} returned ${ctype || "no content-type"} instead of JSON ` +
        `("${peek}…"). The API route was not found, so the app shell was served ` +
        `instead — this usually means the backend is running older code than the ` +
        `frontend. Restart it (RUN.bat) and try again.`,
    );
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
  /**
   * Re-run plate naming.
   *
   * "rules" -- deterministic geometry/shop-token rules, seconds, no LLM.
   * "llm"   -- one Qwen pass over the CAD dimension export.
   * "stl"   -- two Qwen passes with the exported plate meshes measured in
   *            between: candidates from the CSV, then a final call against real
   *            geometry. Slowest, and the only mode that sees a plate's true
   *            stack thickness or which face its pockets open on.
   */
  classifyJob: (jobId: string, mode: "rules" | "llm" | "stl" = "rules") =>
    req<JobDetail>(`/jobs/${encodeURIComponent(jobId)}/classify`, {
      method: "POST",
      body: JSON.stringify({ mode }),
    }),
  quoteSheet: (jobId: string) => req<QuoteSheet>(`/jobs/${encodeURIComponent(jobId)}/quote-sheet`),

  /**
   * Rename plates. Keys are CAD indices; an empty string clears an override and
   * restores the role's default name.
   *
   * `write_sheets` also rewrites the .xls quote and steel sheets. That needs
   * Excel and pywin32 on the machine running the backend; when they are absent
   * the rename is still saved and the next macro run applies it, and
   * `sheets.skipped_reason` says so.
   */
  putPlateNames: (
    jobId: string,
    names: Record<string, string>,
    writeSheets = true,
  ) =>
    req<PlateRenameResult>(`/jobs/${encodeURIComponent(jobId)}/plate-names`, {
      method: "PUT",
      body: JSON.stringify({ names, write_sheets: writeSheets }),
    }),
  getPlateNames: (jobId: string) =>
    req<{ job_id: string; names: Record<string, string> }>(
      `/jobs/${encodeURIComponent(jobId)}/plate-names`,
    ),

  /**
   * Steel grade override, per CAD index ("4") or per role ("role:a_plate").
   *
   * Unlike a rename this DOES move money -- the grade selects which block of the
   * quote workbook a plate is priced in. An unrecognised grade is refused and
   * named back in `rejected` rather than silently dropped, because a grade with no
   * rows in the workbook would drop the plate off the sheet entirely.
   *
   * `write_sheets` also rewrites the two .xls workbooks, the same way a rename
   * does: the steel type text on the steel sheet, and the plate's row moved into
   * the new grade's block on the quote sheet. Needs Excel + pywin32; without them
   * the override is still saved, `sheets.skipped_reason` says why, and the next
   * macro run applies it.
   */
  putPlateGrades: (
    jobId: string,
    grades: Record<string, string>,
    writeSheets = true,
  ) =>
    req<{
      job_id: string;
      grades: Record<string, string>;
      rejected?: Record<string, string>;
      detail?: string;
      regraded?: { name: string; role: string; grade: string }[];
      sheets?: SheetWriteReport & {
        unresolved?: { name?: string; grade?: string; why?: string }[];
        /**
         * One per plate whose quote row was physically moved into the new
         * grade's block. `price_after` is read back AFTER Excel recalculates,
         * so it is the real repriced number, not an estimate.
         */
        moves?: {
          plate: string;
          workbook: string;
          moved_from_row: number;
          moved_to_row: number;
          grade: string;
          grade_label?: string;
          slot?: string;
          /**
           * Thick stock costs more per pound, so each grade block ends in a run
           * of premium rows (yellow in the template). "premium" means the plate
           * landed there — 5.875"+ for every grade except #1 A-36, where the cut
           * is 2.00". `band_note` explains a correction, including the case
           * where the block had no premium row left.
           */
          thickness?: number | null;
          band?: "premium" | "standard";
          band_note?: string;
          price_before?: number | null;
          price_after?: number | null;
        }[];
      };
    }>(`/jobs/${encodeURIComponent(jobId)}/plate-grades`, {
      method: "PUT",
      body: JSON.stringify({ grades, write_sheets: writeSheets }),
    }),
  getPlateGrades: (jobId: string) =>
    req<{
      job_id: string;
      grades: Record<string, string>;
      valid: string[];
      labels: Record<string, string>;
    }>(`/jobs/${encodeURIComponent(jobId)}/plate-grades`),

  /**
   * What the shop's renames have taught the app: the label now used for each
   * role, the renames behind it, and the ones that look like a misclassified
   * plate rather than a naming preference.
   */
  learnedNames: () =>
    req<{
      labels: Record<string, LearnedPlateName>;
      history: Array<Record<string, unknown>>;
      review: Array<Record<string, unknown>>;
    }>("/learning/names"),
  forgetLearnedName: (base_type: string, role: string) =>
    req<{ forgotten: boolean; role: string }>("/learning/names/forget", {
      method: "POST",
      body: JSON.stringify({ base_type, role }),
    }),

  machining: (jobId: string) =>
    req<MachiningEstimate>(`/jobs/${encodeURIComponent(jobId)}/machining`),

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
