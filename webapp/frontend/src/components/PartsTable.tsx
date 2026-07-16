import { useMemo, useState } from "react";
import { ChevronDown, ChevronRight, Search } from "lucide-react";
import type { QuoteLineItem } from "../api/client";
import { Badge } from "./ui";

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
  const fromComp = shortComponent(row.component || "");
  if (fromComp && fromComp !== "--") return fromComp;
  if (row.role_label && row.role_label.trim()) return row.role_label.trim();
  if (row.role && row.role.trim()) return row.role.trim();
  return "--";
}

function sectionHint(group: string): string {
  if (group === "Steel Plates / Mold Base") return "Quote / steel workbook";
  if (group === "Pull Cores & Keys") return "volume × $88 / in³";
  if (group === "Purchased Components") return "DME / McMaster / Jaco";
  return "";
}

export default function PartsTable({
  items,
}: {
  items: QuoteLineItem[];
}) {
  const [query, setQuery] = useState("");
  const [collapsed, setCollapsed] = useState<Record<string, boolean>>({});

  const grouped = useMemo(() => {
    const q = query.trim().toLowerCase();
    const filtered = q
      ? items.filter(
          (p) =>
            p.component?.toLowerCase().includes(q) ||
            p.role_label?.toLowerCase().includes(q) ||
            p.role?.toLowerCase().includes(q) ||
            p.vendor?.toLowerCase().includes(q) ||
            p.part_number?.toLowerCase().includes(q)
        )
      : items;
    const groups: Record<string, QuoteLineItem[]> = {};
    for (const p of filtered) {
      const g = p.role_group || "Other Hardware";
      (groups[g] ||= []).push(p);
    }
    return groups;
  }, [items, query]);

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
      <div className="relative">
        <Search className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-ink-400" />
        <input
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          placeholder="Search by description, vendor, or part number..."
          className="glass-input w-full rounded-full py-2.5 pl-9 pr-3 text-sm text-ink-100 placeholder:text-ink-500"
        />
      </div>

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
                      {isSteel && <th className="px-4 py-2 font-medium">Hours</th>}
                      {isPurchased && <th className="px-4 py-2 text-right font-medium">Unit $</th>}
                      <th className="px-4 py-2 text-right font-medium">
                        {isPurchased ? "Ext $" : "Price"}
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    {rows.map((row) => (
                      <tr key={row.index} className="border-t border-ink-800/60 hover:bg-ink-800/30">
                        <td className="max-w-xs truncate px-4 py-2 font-mono text-xs text-ink-200" title={row.component || row.role_label}>
                          {rowDescription(row)}
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
