import { useMemo, useRef, useState, type ReactNode } from "react";
import { ChevronDown, ChevronRight, Search, Layers, Wrench, Pencil } from "lucide-react";
import type { QuoteLineItem } from "../api/client";
import { Badge } from "./ui";
import { rollupParts, isConsumablePart, partIdentity, type RolledRow } from "../lib/partRollup";

const SECTION_ORDER = [
  "Steel Plates / Mold Base",
  "Pull Cores & Keys",
  "Purchased Components",
  "Mold Base Plates",
  "Rails",
  "Ejector Assembly",
  "Latch Locks / Safety",
  "Guide Hardware",
  "Core / Cavity Details",
  "Other Hardware",
  "Ignored",
] as const;

function fmtNum(v: number | null | undefined, digits = 3): string {
  if (v == null || Number.isNaN(Number(v))) return "--";
  const n = Number(v);
  if (n === 0) return "0";
  return n.toLocaleString(undefined, {
    minimumFractionDigits: 0,
    maximumFractionDigits: digits,
  });
}

function fmtMoney(v: number | null | undefined): string {
  if (v == null || Number.isNaN(Number(v)) || Number(v) === 0) return "--";
  return `$${Number(v).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

function shortComponent(name: string) {
  if (!name) return "--";
  const segs = name.split("/");
  const last = segs[segs.length - 1]?.trim();
  return last || "--";
}

function rowDescription(row: QuoteLineItem): string {
  // A rename wins over the CAD component name. The backend already resolves
  // this, but the rollup path builds its own identity string, so it is applied
  // here too rather than relying on one of the two paths.
  if (row.name_override && row.name_override.trim()) return row.name_override.trim();
  const fromComp = shortComponent(row.component || "");
  if (fromComp && fromComp !== "--") return fromComp;
  if (row.role_label && row.role_label.trim()) return row.role_label.trim();
  if (row.role && row.role.trim()) return row.role.trim();
  return "--";
}

/**
 * Click-to-edit plate name.
 *
 * Deliberately NOT a live-as-you-type save: this writes to the quote sheet and
 * the steel sheet, which are the shop's actual deliverables. Enter commits,
 * Escape cancels, and clearing the box restores the role's default name.
 */
/** Roles that identify nothing specific — mirrors GENERIC_ROLES in plate_names.py. */
const GENERIC_ROLES = new Set(["", "steel_plate", "purchased_component", "other", "ignore"]);

/**
 * Store key for renaming this row. Must ALWAYS return something usable.
 *
 * Three kinds, most specific first:
 *   "7"                 a CAD index — one solid, so two B plates rename apart.
 *   "role:b_plate"      a role — a steel row renaming "the B plate".
 *   "name:b plate"      the displayed name — the last resort.
 *
 * The name fallback is why renaming can no longer fail silently. Steel rows are
 * numbered S1, S2… (workbook positions, not identities) and every row whose name
 * the sheet parser could not map gets role "steel_plate", so neither of the first
 * two keys is always available — and a key nothing reads is a rename that
 * accepts your typing and then reverts.
 */
/**
 * Mirror of `_norm_name` in backend/app/plate_names.py. MUST stay identical.
 *
 * This is the bug that made renaming look broken. The backend looks a name key up
 * as `name:` + lowercased, non-alphanumerics collapsed to single spaces, while
 * this file used to send the raw displayed text:
 *
 *     row shown as  B Plate(17-78X35-12X3-38)
 *     UI saved      name:B Plate(17-78X35-12X3-38)
 *     API read      name:b plate 17 78x35 12x3 38
 *
 * Those never match -- not for punctuated names, not even for "A Plate", because
 * of the lowercasing. set_overrides stores whatever key it is handed, so the save
 * genuinely succeeded and the response echoed the new name back; then lookup()
 * failed to find it and the row re-rendered with the old label. Type, press
 * Enter, watch it revert.
 *
 * It only showed on the steel/pricing rows because those are the ones that reach
 * the name key: numbered S1, S2... so no CAD index, and role "steel_plate" so the
 * role key is generic.
 */
function normName(s: string): string {
  return (s || "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
}

function renameKeyFor(row: QuoteLineItem, display: string): string {
  const idx = String(row.index ?? "");
  if (/^\d+$/.test(idx)) return idx;
  const role = (row.role || "").trim();
  if (role && !GENERIC_ROLES.has(role)) return `role:${role}`;
  return `name:${normName(display)}`;
}

function EditableName({
  row,
  display,
  onSave,
  busy,
}: {
  row: QuoteLineItem;
  display: string;
  onSave: (index: string, name: string) => void;
  busy: boolean;
}) {
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState(display);
  const sent = useRef(false);

  if (!editing) {
    return (
      <span className="flex items-center gap-1.5">
        <span className="truncate">{display}</span>
        {row.name_override && (
          <span
            className="shrink-0 rounded bg-brand-500/20 px-1 py-0.5 text-[9px] font-semibold text-brand-300"
            title={
              `Renamed. Priced as "${row.role_label_original || row.role}" — ` +
              `renaming changes the label only, never the quote row.`
            }
          >
            renamed
          </span>
        )}
        {/* Always visible in edit mode. The first version was opacity-0 until
            hover, which meant the control was effectively undiscoverable --
            "the edit name thing is not working" was really "I cannot find it". */}
        <button
          type="button"
          onClick={() => {
            sent.current = false;
            setDraft(row.name_override || display);
            setEditing(true);
          }}
          className="shrink-0 rounded border border-white/10 bg-ink-800 p-1 text-ink-300 transition hover:border-brand-500/50 hover:bg-ink-700 hover:text-ink-100"
          title="Rename this plate everywhere — parts table, 3D, geometry, quote sheet and steel sheet"
          aria-label="Rename plate"
        >
          <Pencil className="h-3 w-3" />
        </button>
      </span>
    );
  }

  const commit = () => {
    // Enter fires onKeyDown AND then blurs the input, so commit ran twice and
    // saved twice. Guarded rather than dropping onBlur, because clicking away
    // should still save.
    if (sent.current) return;
    const next = draft.trim();
    setEditing(false);
    if (!next || next === (row.name_override || display)) return;
    sent.current = true;
    onSave(renameKeyFor(row, display), next);
  };

  return (
    <span className="flex items-center gap-1">
      <input
        autoFocus
        value={draft}
        disabled={busy}
        onChange={(e) => setDraft(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === "Enter") commit();
          if (e.key === "Escape") setEditing(false);
        }}
        onBlur={commit}
        maxLength={48}
        placeholder={row.role_label_original || "plate name"}
        className="w-40 rounded border border-brand-500/50 bg-ink-900 px-1.5 py-0.5 font-mono text-xs text-ink-100 outline-none focus:border-brand-400"
      />
      <span className="shrink-0 text-[9px] text-ink-500">enter to save</span>
    </span>
  );
}

/** Shop label for a grade code, matching GRADE_LABELS in app/plate_grades.py. */
const GRADE_LABELS: Record<string, string> = {
  A36: "#1 A-36",
  "4140": "#2 4140",
  P20: "#3 P20",
  "420SS": "420 SS",
  "6061": "6061 ALM",
  A2: "A-2",
  O1: "O-1",
};

const DEFAULT_GRADES = ["A36", "4140", "P20", "420SS", "6061", "A2", "O1"];

/**
 * Steel grade, editable in place.
 *
 * Keyed the same way a rename is -- CAD index when the row has one, else the role
 * -- so changing a row that stands for a role (a steel-sheet line with no single
 * CAD part behind it) moves every plate of that role, and changing a row that is
 * one CAD part moves only that part. That is the behaviour the shop asked for:
 * "apply everywhere".
 *
 * Unlike the rename this reprices, so an overridden cell is marked and carries the
 * default in its tooltip -- an estimator has to be able to see at a glance which
 * grades are the macro's and which are theirs.
 */
function GradePicker({
  row,
  display,
  onSetGrade,
}: {
  row: QuoteLineItem;
  /** Same string the rename control shows, so both build the same store key. */
  display: string;
  onSetGrade: (key: string, grade: string) => void;
}) {
  const options = row.grade_options && row.grade_options.length ? row.grade_options : DEFAULT_GRADES;
  const current = row.grade || "4140";
  const overridden = row.grade_source === "override";
  // Same three-tier key a rename uses -- see renameKeyFor. It used to fall
  // straight to `role:${row.role}`, which on a row with no CAD index and the
  // generic role "steel_plate" meant one grade change repriced every unmapped
  // steel row at once. plate_grades.grade_for now skips generic roles and reads
  // the name key instead.
  const key = renameKeyFor(row, display);

  return (
    <select
      value={current}
      onChange={(e) => onSetGrade(key, e.target.value)}
      title={
        overridden
          ? "Overridden by you. Choose the macro default to clear it."
          : "Macro default for this plate. Change to override."
      }
      className={
        "rounded border bg-ink-900 px-1 py-0.5 font-mono text-[11px] outline-none " +
        (overridden
          ? "border-accent-amber/60 text-accent-amber"
          : "border-white/15 text-ink-300 focus:border-brand-400")
      }
    >
      {options.map((g) => (
        <option key={g} value={g}>
          {GRADE_LABELS[g] || g}
        </option>
      ))}
    </select>
  );
}

function sectionHint(group: string): string {
  if (group === "Steel Plates / Mold Base") return "Quote / steel workbook";
  if (group === "Pull Cores & Keys") return "volume × $88 / in³";
  if (group === "Purchased Components") return "DME / McMaster / Jaco";
  return "";
}

/**
 * Empty scaffolding rows, hidden. NEVER hidden BY NAME.
 *
 * This used to be a name list — `["pin plate", "ejector plate"]` — described as
 * "sheet-stock lines with no price of their own". That was true of the rows it
 * was written against, and it stopped being true: those two names belong to real
 * plates on a real mold base, and the filter could not tell the difference.
 *
 * C18027 is a 7-plate base whose Ejector Plate is 0.5 x 5.25 x 11.875 at $31.14
 * on both the steel order and the quote sheet. The Parts tab showed six plates.
 * Across the registry the name list was deleting six real plates worth $720.85,
 * and catching zero of the empty rows it was meant for.
 *
 * Worse than a hidden row, because filterVisibleParts runs BEFORE the group
 * count and the group total: the header said 6, the total said $585.86 against
 * the workbook's $617.00, and nothing on the page said a line had been removed.
 * A quote silently short by a plate is the one failure this table must not have.
 *
 * So the test is now what the row IS, not what it is called. A row with no
 * price, no quantity and no dimensions is template scaffolding and carries no
 * money; anything else is a plate, whatever it happens to be named.
 */
function isHiddenSteelRow(row: QuoteLineItem): boolean {
  const isSteelGroup =
    row.role_group === "Steel Plates / Mold Base" || row.role_group === "Mold Base Plates";
  if (!isSteelGroup) return false;
  const num = (v: unknown) => Number(v) > 0;
  if (num(row.price) || num(row.qty) || num(row.hours)) return false;
  return !num(row.thickness) && !num(row.width) && !num(row.length);
}

// Exported so pages that count parts (e.g. the "Parts & Pricing (N)" tab label)
// can filter the same way PartsTable does, keeping the count accurate.
export function filterVisibleParts(items: QuoteLineItem[]): QuoteLineItem[] {
  return items.filter((p) => !isHiddenSteelRow(p));
}

function Toggle({
  active,
  onClick,
  icon,
  label,
  title,
}: {
  active: boolean;
  onClick: () => void;
  icon: ReactNode;
  label: string;
  title: string;
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      title={title}
      className={`flex items-center gap-1.5 rounded-full border px-3 py-1.5 text-[11px] font-semibold uppercase tracking-wider transition ${
        active
          ? "border-brand-500/40 bg-brand-500/15 text-brand-400"
          : "border-white/10 text-ink-400 hover:border-white/20 hover:text-ink-200"
      }`}
    >
      {icon}
      {label}
    </button>
  );
}

export default function PartsTable({
  items,
  onRenamePlate,
  onSetGrade,
  renaming,
}: {
  items: QuoteLineItem[];
  /**
   * Save a plate rename. Omit to make the table read-only -- the pencil only
   * appears when a handler is supplied, so no other caller grows an edit
   * control by accident.
   */
  onRenamePlate?: (index: string, name: string) => void;
  /** Save a steel-grade override. Omit to hide the Steel column entirely. */
  onSetGrade?: (key: string, grade: string) => void;
  renaming?: boolean;
}) {
  const [query, setQuery] = useState("");
  const [collapsed, setCollapsed] = useState<Record<string, boolean>>({});
  // Explicit mode rather than an always-on pencil on every row. Renaming writes
  // to the quote sheet and the steel sheet, so it should be something you turn on
  // deliberately, and it keeps the table clean the rest of the time.
  const [editNames, setEditNames] = useState(false);
  // Both default ON: a real mold base has hundreds of screws and plugs, and
  // showing every CAD instance of each one buries the plates that carry money.
  const [rollup, setRollup] = useState(true);
  const [hideConsumables, setHideConsumables] = useState(true);

  const consumableCount = useMemo(
    () => filterVisibleParts(items).filter((p) => isConsumablePart(p.component || "")).length,
    [items],
  );

  const grouped = useMemo(() => {
    let visible = filterVisibleParts(items);

    if (hideConsumables) {
      visible = visible.filter((p) => !isConsumablePart(p.component || ""));
    }

    const q = query.trim().toLowerCase();
    const filtered = q
      ? visible.filter(
          (p) =>
            p.component?.toLowerCase().includes(q) ||
            p.role_label?.toLowerCase().includes(q) ||
            p.role?.toLowerCase().includes(q) ||
            p.vendor?.toLowerCase().includes(q) ||
            p.part_number?.toLowerCase().includes(q)
        )
      : visible;

    const groups: Record<string, RolledRow[]> = {};
    for (const p of filtered) {
      const g = p.role_group || "Other Hardware";
      (groups[g] ||= []).push({ ...p, rollupCount: 1, identity: partIdentity(p.component || "") });
    }

    if (rollup) {
      for (const g of Object.keys(groups)) {
        groups[g] = rollupParts(groups[g]);
      }
    }

    return groups;
  }, [items, query, rollup, hideConsumables]);

  const orderedGroups = [
    ...SECTION_ORDER.filter((g) => grouped[g]?.length),
    ...Object.keys(grouped).filter((g) => !(SECTION_ORDER as readonly string[]).includes(g)),
  ];

  if (!items.length) {
    return (
      <div className="glass-panel rounded-2xl px-4 py-8 text-center text-sm text-ink-400">
        No steel, pull-core, or purchased-component lines found yet. Re-sync after Module6121 finishes
        writing the quote workbook / Pullcore Prices.csv / Purchased Components Quote.csv.
      </div>
    );
  }

  return (
    <div className="space-y-4">
      <div className="flex flex-col gap-3 sm:flex-row sm:items-center">
        <div className="relative flex-1">
          <Search className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-ink-400" />
          <input
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="Search by description, vendor, or part number..."
            className="glass-input w-full rounded-full py-2.5 pl-9 pr-3 text-sm text-ink-100 placeholder:text-ink-500"
          />
        </div>
        <div className="flex shrink-0 flex-wrap gap-2">
          <Toggle
            active={rollup}
            onClick={() => setRollup((v) => !v)}
            icon={<Layers className="h-3.5 w-3.5" />}
            label="Roll up duplicates"
            title="Merge identical parts into one row with a summed quantity, instead of one row per CAD instance."
          />
          <Toggle
            active={hideConsumables}
            onClick={() => setHideConsumables((v) => !v)}
            icon={<Wrench className="h-3.5 w-3.5" />}
            label={consumableCount > 0 ? `Hide fasteners (${consumableCount})` : "Hide fasteners"}
            title="Hide screws, dowels, plugs, baffles and other consumables that are bought by the box and never quoted per mold."
          />
          {onRenamePlate && (
            <Toggle
              active={editNames}
              onClick={() => setEditNames((v) => !v)}
              icon={<Pencil className="h-3.5 w-3.5" />}
              label={editNames ? "Done renaming" : "Rename plates"}
              title="Turn on the rename control for each plate. A rename updates the parts table, the 3D tab, Geometry, Machining, the quote sheet and the steel sheet. It changes the NAME only — the plate keeps its quote row and its price."
            />
          )}
        </div>
      </div>

      {editNames && (
        <div className="rounded-lg border border-brand-500/30 bg-brand-500/10 px-3 py-2 text-xs leading-relaxed text-ink-200">
          <strong className="text-ink-100">Renaming.</strong> Click the pencil on any plate
          row, type a name, press Enter. Clearing the box restores the original.
          <span className="text-ink-400">
            {" "}
            The name changes everywhere — including the quote sheet and steel sheet — but the
            plate keeps its quote row and price, so nothing moves in the pricing.
          </span>
        </div>
      )}

      {orderedGroups.map((group) => {
        const rows = grouped[group];
        const isCollapsed = collapsed[group];
        const groupTotal = rows.reduce((sum, r) => sum + (r.price || 0), 0);
        const isPurchased = group === "Purchased Components";
        const isPullcore = group === "Pull Cores & Keys";
        const isSteel = group === "Steel Plates / Mold Base" || group === "Mold Base Plates";
        const hint = sectionHint(group);

        return (
          <div key={group} className="glass-panel overflow-hidden rounded-2xl">
            <button
              onClick={() => setCollapsed((c) => ({ ...c, [group]: !c[group] }))}
              className="flex w-full items-center justify-between bg-ink-800/70 px-4 py-2.5 text-left"
            >
              <div className="flex flex-wrap items-center gap-2">
                {isCollapsed ? (
                  <ChevronRight className="h-4 w-4 text-ink-400" />
                ) : (
                  <ChevronDown className="h-4 w-4 text-ink-400" />
                )}
                <span className="text-sm font-semibold text-ink-100">{group}</span>
                <Badge>{rows.length}</Badge>
                {hint && <span className="text-[10px] text-ink-500">{hint}</span>}
              </div>
              {groupTotal > 0 && (
                <span className="text-xs font-medium text-ink-300">
                  {fmtMoney(groupTotal)}
                </span>
              )}
            </button>
            {!isCollapsed && (
              <div className="scrollbar-thin overflow-x-auto">
                <table className="w-full text-left text-sm">
                  <thead>
                    <tr className="border-t border-ink-700/60 bg-ink-850/40 text-[11px] uppercase tracking-wider text-ink-400">
                      <th className="px-4 py-2 font-medium">
                        {isPurchased ? "Component" : "Description"}
                      </th>
                      {isPurchased && <th className="px-4 py-2 font-medium">Vendor</th>}
                      {isPurchased && <th className="px-4 py-2 font-medium">Part #</th>}
                      <th className="px-4 py-2 font-medium">QTY</th>
                      <th className="px-4 py-2 font-medium">Thickness</th>
                      <th className="px-4 py-2 font-medium">Width</th>
                      <th className="px-4 py-2 font-medium">Length</th>
                      {isPullcore && <th className="px-4 py-2 font-medium">Cu. In.</th>}
                      {isSteel && onSetGrade && <th className="px-4 py-2 font-medium">Steel</th>}
                      {isSteel && <th className="px-4 py-2 font-medium">Hours</th>}
                      {isPurchased && <th className="px-4 py-2 text-right font-medium">Unit $</th>}
                      <th className="px-4 py-2 text-right font-medium">
                        {isPurchased ? "Ext $" : "Price"}
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    {rows.map((row) => (
                      <tr
                        key={`${row.index}-${row.identity}`}
                        className="border-t border-ink-800/60 hover:bg-ink-800/30"
                      >
                        <td
                          className="group max-w-xs px-4 py-2 font-mono text-xs text-ink-200"
                          title={row.component || row.role_label}
                        >
                          <span className="flex items-center gap-2">
                            {/* Steel rows are included, not just CAD rows: they
                                are the ones with the real plate names and the
                                prices, and after de-duplication they are usually
                                the ONLY row for a plate. Gating on
                                section === "classified" meant the pencil never
                                appeared on the rows worth renaming. */}
                            {onRenamePlate &&
                            editNames &&
                            (row.section === "classified" || row.section === "steel") ? (
                              <EditableName
                                row={row}
                                display={rollup && row.identity ? row.identity : rowDescription(row)}
                                onSave={onRenamePlate}
                                busy={Boolean(renaming)}
                              />
                            ) : (
                              <span className="truncate">
                                {rollup && row.identity ? row.identity : rowDescription(row)}
                              </span>
                            )}
                            {row.rollupCount > 1 && (
                              <span
                                className="shrink-0 rounded-full bg-white/10 px-1.5 py-0.5 text-[9px] font-semibold text-ink-400"
                                title={`${row.rollupCount} identical CAD instances merged`}
                              >
                                ×{row.rollupCount}
                              </span>
                            )}
                          </span>
                        </td>
                        {isPurchased && (
                          <td className="px-4 py-2 text-xs text-ink-300">{row.vendor || "--"}</td>
                        )}
                        {isPurchased && (
                          <td className="px-4 py-2 font-mono text-xs text-ink-300">{row.part_number || "--"}</td>
                        )}
                        <td className="px-4 py-2 text-ink-100">{fmtNum(row.qty, 0)}</td>
                        <td className="px-4 py-2 text-xs text-ink-300">{fmtNum(row.thickness)}</td>
                        <td className="px-4 py-2 text-xs text-ink-300">{fmtNum(row.width)}</td>
                        <td className="px-4 py-2 text-xs text-ink-300">{fmtNum(row.length)}</td>
                        {isPullcore && (
                          <td className="px-4 py-2 text-xs text-ink-300">{fmtNum(row.cu_in, 2)}</td>
                        )}
                        {isSteel && onSetGrade && (
                          <td className="px-4 py-2 text-xs">
                            <GradePicker
                              row={row}
                              display={rollup && row.identity ? row.identity : rowDescription(row)}
                              onSetGrade={onSetGrade}
                            />
                          </td>
                        )}
                        {isSteel && (
                          <td className="px-4 py-2 text-xs text-ink-300">{fmtNum(row.hours, 1)}</td>
                        )}
                        {isPurchased && (
                          <td className="px-4 py-2 text-right text-xs text-ink-300">
                            {fmtMoney(row.unit_price)}
                          </td>
                        )}
                        <td className="px-4 py-2 text-right font-medium text-ink-100">
                          {row.price > 0
                            ? fmtMoney(row.price)
                            : row.section === "steel"
                              ? <span className="text-[10px] text-ink-500">sheet</span>
                              : "--"}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                  {groupTotal > 0 && (
                    <tfoot>
                      <tr className="border-t border-ink-700/80 bg-ink-850/30">
                        <td
                          className="px-4 py-2 text-xs font-semibold text-ink-200"
                          colSpan={
                            isPurchased ? 8 : isPullcore || isSteel ? 6 : 5
                          }
                        >
                          Total
                        </td>
                        <td className="px-4 py-2 text-right text-sm font-semibold text-ink-100">
                          {fmtMoney(groupTotal)}
                        </td>
                      </tr>
                    </tfoot>
                  )}
                </table>
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}
