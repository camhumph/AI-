import type { ReactNode } from "react";
import Sidebar, { MobileDock } from "./Sidebar";

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
    <div className="flex h-screen w-full overflow-hidden">
      <Sidebar />
      <div className="flex min-w-0 flex-1 flex-col pr-3 pt-3 pb-3">
        <header className="glass-panel mb-3 flex shrink-0 items-center justify-between rounded-2xl px-6 py-4 sm:px-8">
          <div>
            <h1 className="text-sm font-semibold uppercase tracking-[0.22em] text-ink-100 sm:text-base">
              {title}
            </h1>
            {subtitle && <p className="mt-1 text-xs text-ink-500">{subtitle}</p>}
          </div>
          {actions && <div className="flex items-center gap-3">{actions}</div>}
        </header>
        <main className="scrollbar-thin glass-panel flex-1 overflow-y-auto rounded-2xl px-6 py-8 pb-24 sm:px-10 md:pb-8 space-grid">
          {children}
        </main>
      </div>
      <MobileDock />
    </div>
  );
}
