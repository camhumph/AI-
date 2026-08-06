import type { QuoteLineItem } from "../api/client";

/**
 * Collapse a CAD component name to its identity, dropping the per-instance
 * suffix SolidWorks appends.
 *
 *   "SHCS(12X2-34)-1-30"  -> "SHCS(12X2-34)"
 *   "Support Pillar-1-11"  -> "Support Pillar"
 *   "asm-1/Ejector Pin B(EX-25 .375)-1-7" -> "Ejector Pin B(EX-25 .375)"
 *
 * Without this, a job like C18599 renders 30 separate rows for one screw and
 * 288 rows of "Other Hardware".
 */
export function partIdentity(rawName: string): string {
  let s = (rawName || "").trim();
  if (!s) return "";

  // Drop the parent assembly path.
  const slash = Math.max(s.lastIndexOf("/"), s.lastIndexOf("\\"));
  if (slash >= 0) s = s.slice(slash + 1);

  // Strip trailing instance markers: "-1-30", "-1", "_1_30".
  s = s.replace(/[-_]\d+[-_]\d+\s*$/, "");
  s = s.replace(/[-_]\d+\s*$/, "");

  return s.trim() || (rawName || "").trim();
}

/**
 * Fasteners, plugs, baffles and other consumables the shop does not quote.
 *
 * These are real parts in the CAD and they belong in the export, but they are
 * bought by the box and never priced per mold, so showing 200 of them buries
 * the plates that actually carry money.
 */
const CONSUMABLE_PATTERNS: RegExp[] = [
  /\bSHCS\b/i,
  /\bFHCS\b/i,
  /\bLHCS\b/i,
  /\bBHCS\b/i,
  /\bSHSS\b/i,
  /socket[_ ]?head[_ ]?cap[_ ]?screw/i,
  /\bhexagon[_ ]?socket/i,
  /\bcap screw\b/i,
  /\bset screw\b/i,
  /\bscrew\b/i,
  /\bbolt\b/i,
  /\bwasher\b/i,
  /\bnut\b/i,
  /\bo-?ring\b/i,
  /\bdowel\b/i,
  /\bplug\b/i,          // BR37 Plug, BR25 Plug, Extension Plug
  /\bbaffle\b/i,        // SB50/SB37 Straight Baffle
  /\bnametag\b/i,
  /\bwater blade\b/i,
  /\beye ?bolt\b/i,
  /\blifting\b/i,
];

export function isConsumablePart(name: string): boolean {
  const s = partIdentity(name);
  if (!s) return false;
  return CONSUMABLE_PATTERNS.some((re) => re.test(s));
}

export interface RolledRow extends QuoteLineItem {
  /** How many CAD instances collapsed into this row. */
  rollupCount: number;
  /** The identity used to group them. */
  identity: string;
}

function dimKey(v: number | null | undefined): string {
  const n = Number(v);
  if (!Number.isFinite(n)) return "-";
  return n.toFixed(3);
}

/**
 * Group identical parts within an already-sectioned list.
 *
 * Two rows merge only when their identity AND all three dimensions match, so a
 * "SHCS(12X2-0)" at 0.75 thick stays separate from the 1.061 variant -- those
 * really are different parts, and the CAD bbox is how we can tell.
 *
 * Quantity and price are summed; everything else is taken from the first row so
 * vendor / part number / material survive.
 */
export function rollupParts(items: QuoteLineItem[]): RolledRow[] {
  const byKey = new Map<string, RolledRow>();
  const order: string[] = [];

  for (const it of items) {
    const identity = partIdentity(it.component || it.role_label || it.role || "");
    const key = [
      it.role_group || "",
      identity.toLowerCase(),
      dimKey(it.thickness),
      dimKey(it.width),
      dimKey(it.length),
    ].join("|");

    const existing = byKey.get(key);
    if (existing) {
      existing.qty = (Number(existing.qty) || 0) + (Number(it.qty) || 0);
      existing.price = (Number(existing.price) || 0) + (Number(it.price) || 0);
      if (Number.isFinite(Number(it.hours))) {
        existing.hours = (Number(existing.hours) || 0) + (Number(it.hours) || 0);
      }
      if (Number.isFinite(Number(it.cu_in))) {
        existing.cu_in = (Number(existing.cu_in) || 0) + (Number(it.cu_in) || 0);
      }
      existing.rollupCount += 1;
      continue;
    }

    const row: RolledRow = { ...it, rollupCount: 1, identity };
    byKey.set(key, row);
    order.push(key);
  }

  return order.map((k) => byKey.get(k) as RolledRow);
}
