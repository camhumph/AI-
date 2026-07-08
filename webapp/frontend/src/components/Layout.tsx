import type { ReactNode } from "react";
import Sidebar from "./Sidebar";

export default function Layout({
  title,
  subtitle,
  actions,
  children,
}: {
  title: string;
  subtitle?: string;
  actions?: ReactNode;
  children: ReactNode;
}) {
  return (
    <div className="flex h-screen w-full overflow-hidden bg-ink-950">
      <Sidebar />
      <div className="flex min-w-0 flex-1 flex-col">
        <header className="flex shrink-0 items-center justify-between border-b border-ink-700 px-6 py-5 sm:px-10">
          <div>
            <h1 className="text-sm font-semibold uppercase tracking-[0.2em] text-ink-100 sm:text-base">{title}</h1>
            {subtitle && <p className="mt-1 text-xs text-ink-400">{subtitle}</p>}
          </div>
          {actions && <div className="flex items-center gap-3">{actions}</div>}
        </header>
        <main className="scrollbar-thin flex-1 overflow-y-auto px-6 py-8 sm:px-10">{children}</main>
      </div>
    </div>
  );
}
