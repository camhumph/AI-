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
    <div className="flex h-screen w-full overflow-hidden">
      <Sidebar />
      <div className="flex min-w-0 flex-1 flex-col">
        <header className="flex shrink-0 items-center justify-between border-b border-ink-700/60 bg-ink-900/40 px-4 py-4 backdrop-blur-xl sm:px-8">
          <div>
            <h1 className="text-lg font-semibold text-ink-100 sm:text-xl">{title}</h1>
            {subtitle && <p className="mt-0.5 text-sm text-ink-400">{subtitle}</p>}
          </div>
          {actions && <div className="flex items-center gap-2">{actions}</div>}
        </header>
        <main className="scrollbar-thin flex-1 overflow-y-auto px-4 py-6 sm:px-8">{children}</main>
      </div>
    </div>
  );
}
