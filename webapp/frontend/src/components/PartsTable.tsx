import { useMemo, useState } from "react";
import { ChevronDown, ChevronRight, Search } from "lucide-react";
import type { PartRow, QuoteLineItem } from "../api/client";
import { Badge } from "./ui";

const GROUP_ORDER = [
  "Mold Base Plates",
  "Rails",
  "Ejector Assembly",
  "Latch Locks / Safety",
  "Guide Hardware",
  "Core / Cavity Details",
  "Other Hardware",
  "Ignored",
];

function confidenceTone(c: string): "success" | "warning" | "danger" | "neutral" {
  const v = (c || "").toUpperCase();
  if (v === "HIGH") return "success";
  if (v === "MEDIUM") return "warning";
  if (v === "LOW") return "danger";
  return "neutral";
}

function shortComponent(name: string) {
  if (!name) return "--";
  const segs = name.split("/");
  return segs[segs.length - 1];
}

export default function PartsTable({
  parts,
  prices,
}: {
  parts: PartRow[];
  prices?: Record<string, QuoteLineItem>;
}) {
  const [query, setQuery] = useState("");
  const [collapsed, setCollapsed] = useState<Record<string, boolean>>({});

  const grouped = useMemo(() => {
    const q = query.trim().toLowerCase();
    const filtered = q
      ? parts.filter(
          (p) =>
            p.Component?.toLowerCase().includes(q) ||
            p.role_label.toLowerCase().includes(q) ||
            p.role.toLowerCase().includes(q)
        )
      : parts;
    const groups: Record<string, PartRow[]> = {};
    for (const p of filtered) {
      (groups[p.role_group] ||= []).push(p);
    }
    return groups;
  }, [parts, query]);

  const orderedGroups = GROUP_ORDER.filter((g) => grouped[g]?.length);

  return (
    <div className="space-y-4">
      <div className="relative">
        <Search className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-ink-400" />
        <input
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          placeholder="Search parts by name or role..."
          className="w-full rounded-xl border border-ink-700/60 bg-ink-850/70 py-2.5 pl-9 pr-3 text-sm text-ink-100 placeholder:text-ink-500 focus:border-brand-500/60 focus:outline-none"
        />
      </div>

      {orderedGroups.map((group) => {
        const rows = grouped[group];
        const isCollapsed = collapsed[group];
        const groupTotal = rows.reduce((sum, r) => sum + (prices?.[r.index]?.price || 0), 0);
        return (
          <div key={group} className="overflow-hidden rounded-xl border border-ink-700/60">
            <button
              onClick={() => setCollapsed((c) => ({ ...c, [group]: !c[group] }))}
              className="flex w-full items-center justify-between bg-ink-800/70 px-4 py-2.5 text-left"
            >
              <div className="flex items-center gap-2">
                {isCollapsed ? (
                  <ChevronRight className="h-4 w-4 text-ink-400" />
                ) : (
                  <ChevronDown className="h-4 w-4 text-ink-400" />
                )}
                <span className="text-sm font-semibold text-ink-100">{group}</span>
                <Badge>{rows.length}</Badge>
              </div>
              {groupTotal > 0 && (
                <span className="text-xs font-medium text-ink-300">${groupTotal.toLocaleString(undefined, { minimumFractionDigits: 2 })}</span>
              )}
            </button>
            {!isCollapsed && (
              <div className="scrollbar-thin overflow-x-auto">
                <table className="w-full text-left text-sm">
                  <thead>
                    <tr className="border-t border-ink-700/60 bg-ink-850/40 text-[11px] uppercase tracking-wider text-ink-400">
                      <th className="px-4 py-2 font-medium">Component</th>
                      <th className="px-4 py-2 font-medium">Role</th>
                      <th className="px-4 py-2 font-medium">Confidence</th>
                      <th className="px-4 py-2 font-medium">Dimensions (T x W x L)</th>
                      <th className="px-4 py-2 font-medium">Quote</th>
                      {prices && <th className="px-4 py-2 text-right font-medium">Price</th>}
                    </tr>
                  </thead>
                  <tbody>
                    {rows.map((row) => (
                      <tr key={row.index} className="border-t border-ink-800/60 hover:bg-ink-800/30">
                        <td className="max-w-xs truncate px-4 py-2 font-mono text-xs text-ink-200" title={row.Component}>
                          {shortComponent(row.Component)}
                        </td>
                        <td className="px-4 py-2 text-ink-100">{row.role_label}</td>
                        <td className="px-4 py-2">
                          <Badge tone={confidenceTone(row.confidence)}>{row.confidence}</Badge>
                        </td>
                        <td className="px-4 py-2 text-xs text-ink-400">
                          {(prices?.[row.index]?.thickness ?? row.Thickness) || "--"} x{" "}
                          {(prices?.[row.index]?.width ?? row.Width) || "--"} x{" "}
                          {(prices?.[row.index]?.length ?? row.Length) || "--"}
                        </td>
                        <td className="px-4 py-2">
                          {row.quote ? <Badge tone="brand">Quoted</Badge> : <Badge>--</Badge>}
                        </td>
                        {prices && (
                          <td className="px-4 py-2 text-right font-medium text-ink-100">
                            {prices[row.index]?.price
                              ? `$${prices[row.index].price.toLocaleString(undefined, { minimumFractionDigits: 2 })}`
                              : prices[row.index]?.price_source === "no_csv_price"
                                ? <span className="text-accent-amber text-[10px]">NO CSV</span>
                                : "--"}
                          </td>
                        )}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}
