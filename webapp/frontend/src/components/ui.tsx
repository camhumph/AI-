import type { ReactNode } from "react";
import clsx from "clsx";
import { Loader2 } from "lucide-react";

export function Card({ className, children }: { className?: string; children: ReactNode }) {
  return (
    <div className={clsx("glass-panel rounded-sm", className)}>
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
  void accent;
  return (
    <Card className="p-5">
      <div className="section-label mb-3">{label}</div>
      <div className="text-3xl font-light tracking-tight text-ink-100">{value}</div>
      {sub && <div className="mt-2 text-xs text-ink-400">{sub}</div>}
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
    neutral: "border border-ink-600 text-ink-300",
    brand: "border border-ink-100 text-ink-100",
    success: "border border-accent-green/60 text-accent-green",
    warning: "border border-accent-amber/60 text-accent-amber",
    danger: "border border-accent-rose/60 text-accent-rose",
  };
  return (
    <span className={clsx("inline-flex items-center whitespace-nowrap px-2 py-0.5 text-[10px] font-semibold uppercase tracking-wider", tones[tone])}>
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
    <div className="flex flex-col items-center justify-center border border-dashed border-ink-700 px-8 py-16 text-center">
      {icon && <div className="mb-4 text-ink-500">{icon}</div>}
      <div className="text-sm font-semibold uppercase tracking-wider text-ink-100">{title}</div>
      {description && <p className="mt-2 max-w-sm text-xs text-ink-400">{description}</p>}
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
      "border border-ink-100 bg-ink-100 text-ink-950 hover:bg-transparent hover:text-ink-100",
    secondary:
      "border border-ink-500 text-ink-100 hover:border-ink-100 hover:bg-ink-100 hover:text-ink-950",
    ghost: "text-ink-400 hover:text-ink-100",
  };
  return (
    <button
      className={clsx(
        "inline-flex items-center justify-center gap-2 px-5 py-2 text-xs font-semibold uppercase tracking-widest transition disabled:cursor-not-allowed disabled:opacity-40",
        variants[variant],
        className
      )}
      {...props}
    >
      {children}
    </button>
  );
}
