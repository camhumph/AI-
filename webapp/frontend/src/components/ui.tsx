import type { ReactNode } from "react";
import clsx from "clsx";
import { Loader2 } from "lucide-react";

export function Card({ className, children }: { className?: string; children: ReactNode }) {
  return (
    <div className={clsx("glass-panel rounded-2xl shadow-xl shadow-black/20", className)}>
      {children}
    </div>
  );
}

export function StatCard({
  label,
  value,
  sub,
  accent = "brand",
}: {
  label: string;
  value: ReactNode;
  sub?: ReactNode;
  accent?: "brand" | "teal" | "amber" | "rose" | "green";
}) {
  const accents: Record<string, string> = {
    brand: "from-brand-500/25 to-brand-500/0 text-brand-400",
    teal: "from-accent-teal/25 to-accent-teal/0 text-accent-teal",
    amber: "from-accent-amber/25 to-accent-amber/0 text-accent-amber",
    rose: "from-accent-rose/25 to-accent-rose/0 text-accent-rose",
    green: "from-accent-green/25 to-accent-green/0 text-accent-green",
  };
  return (
    <Card className="p-5">
      <div className={clsx("mb-3 h-1 w-10 rounded-full bg-gradient-to-r", accents[accent])} />
      <div className="text-xs font-medium uppercase tracking-wider text-ink-400">{label}</div>
      <div className="mt-1.5 text-2xl font-semibold text-ink-100">{value}</div>
      {sub && <div className="mt-1 text-xs text-ink-400">{sub}</div>}
    </Card>
  );
}

export function Badge({
  children,
  tone = "neutral",
}: {
  children: ReactNode;
  tone?: "neutral" | "brand" | "success" | "warning" | "danger";
}) {
  const tones: Record<string, string> = {
    neutral: "bg-ink-700/60 text-ink-200 ring-1 ring-ink-600/60",
    brand: "bg-brand-500/15 text-brand-400 ring-1 ring-brand-500/30",
    success: "bg-accent-green/15 text-accent-green ring-1 ring-accent-green/30",
    warning: "bg-accent-amber/15 text-accent-amber ring-1 ring-accent-amber/30",
    danger: "bg-accent-rose/15 text-accent-rose ring-1 ring-accent-rose/30",
  };
  return (
    <span className={clsx("inline-flex items-center whitespace-nowrap rounded-full px-2.5 py-0.5 text-[11px] font-medium", tones[tone])}>
      {children}
    </span>
  );
}

export function EmptyState({ icon, title, description, action }: {
  icon?: ReactNode;
  title: string;
  description?: string;
  action?: ReactNode;
}) {
  return (
    <div className="flex flex-col items-center justify-center rounded-2xl border border-dashed border-ink-600/60 bg-ink-850/40 px-8 py-16 text-center">
      {icon && <div className="mb-4 text-ink-400">{icon}</div>}
      <div className="text-sm font-semibold text-ink-100">{title}</div>
      {description && <p className="mt-1.5 max-w-sm text-xs text-ink-400">{description}</p>}
      {action && <div className="mt-5">{action}</div>}
    </div>
  );
}

export function Spinner({ label }: { label?: string }) {
  return (
    <div className="flex items-center gap-2 text-sm text-ink-400">
      <Loader2 className="h-4 w-4 animate-spin" />
      {label}
    </div>
  );
}

export function Button({
  children,
  variant = "primary",
  className,
  ...props
}: React.ButtonHTMLAttributes<HTMLButtonElement> & { variant?: "primary" | "secondary" | "ghost" }) {
  const variants: Record<string, string> = {
    primary:
      "bg-gradient-to-r from-brand-500 to-brand-600 text-white shadow-lg shadow-brand-600/30 hover:brightness-110",
    secondary:
      "bg-ink-800 text-ink-100 ring-1 ring-ink-600/70 hover:bg-ink-700",
    ghost: "text-ink-300 hover:bg-ink-800/70 hover:text-ink-100",
  };
  return (
    <button
      className={clsx(
        "inline-flex items-center justify-center gap-2 rounded-xl px-4 py-2 text-sm font-medium transition disabled:cursor-not-allowed disabled:opacity-50",
        variants[variant],
        className
      )}
      {...props}
    >
      {children}
    </button>
  );
}
