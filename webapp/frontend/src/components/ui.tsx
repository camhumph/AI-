import type { ReactNode } from "react";
import clsx from "clsx";
import { Loader2 } from "lucide-react";

export function Card({ className, children }: { className?: string; children: ReactNode }) {
  return (
    <div className={clsx("glass-panel rounded-2xl", className)}>
      {children}
    </div>
  );
}

export function StatCard({
  label,
  value,
  sub,
}: {
  label: string;
  value: ReactNode;
  sub?: ReactNode;
}) {
  return (
    <Card className="p-5 transition hover:border-white/15">
      <div className="section-label mb-3">{label}</div>
      <div className="font-mono text-3xl font-light tracking-tight text-ink-100">{value}</div>
      {sub && <div className="mt-2 text-xs text-ink-500">{sub}</div>}
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
    neutral: "border border-white/10 bg-white/5 text-ink-400",
    brand: "border border-brand-500/40 bg-brand-500/15 text-brand-400",
    success: "border border-accent-green/35 bg-accent-green/10 text-accent-green",
    warning: "border border-accent-amber/35 bg-accent-amber/10 text-accent-amber",
    danger: "border border-accent-rose/35 bg-accent-rose/10 text-accent-rose",
  };
  return (
    <span
      className={clsx(
        "inline-flex items-center whitespace-nowrap rounded-full px-2.5 py-0.5 text-[10px] font-bold uppercase tracking-wider",
        tones[tone]
      )}
    >
      {children}
    </span>
  );
}

export function EmptyState({
  icon,
  title,
  description,
  action,
}: {
  icon?: ReactNode;
  title: string;
  description?: string;
  action?: ReactNode;
}) {
  return (
    <div className="flex flex-col items-center justify-center rounded-2xl border border-dashed border-white/10 bg-white/[0.02] px-8 py-16 text-center backdrop-blur-sm">
      {icon && <div className="mb-4 text-ink-500">{icon}</div>}
      <div className="text-sm font-semibold uppercase tracking-wider text-ink-200">{title}</div>
      {description && <p className="mt-2 max-w-sm text-xs text-ink-500">{description}</p>}
      {action && <div className="mt-5">{action}</div>}
    </div>
  );
}

export function Spinner({ label }: { label?: string }) {
  return (
    <div className="flex items-center gap-2 text-sm text-ink-500">
      <Loader2 className="h-4 w-4 animate-spin text-brand-400" />
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
      "rounded-full border border-white/20 bg-white text-ink-950 shadow-[0_0_24px_rgba(10,132,255,0.25)] hover:bg-ink-200 hover:shadow-[0_0_32px_rgba(10,132,255,0.35)]",
    secondary:
      "rounded-full glass-input text-ink-200 hover:border-white/20 hover:bg-white/10 hover:text-ink-100",
    ghost: "rounded-full text-ink-500 hover:bg-white/5 hover:text-ink-200",
  };
  return (
    <button
      className={clsx(
        "inline-flex items-center justify-center gap-2 px-5 py-2.5 text-[10px] font-bold uppercase tracking-[0.18em] transition disabled:cursor-not-allowed disabled:opacity-40",
        variants[variant],
        className
      )}
      {...props}
    >
      {children}
    </button>
  );
}
