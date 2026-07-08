import { NavLink } from "react-router-dom";
import { LayoutDashboard, Mail, FileStack, Settings, Wand2 } from "lucide-react";
import clsx from "clsx";

const NAV = [
  { to: "/", label: "Dashboard", icon: LayoutDashboard, end: true },
  { to: "/email", label: "Inbox", icon: Mail },
  { to: "/quotes", label: "Quotes", icon: FileStack },
  { to: "/settings", label: "Settings", icon: Settings },
];

export default function Sidebar() {
  return (
    <aside className="hidden md:flex md:w-64 flex-col shrink-0 border-r border-ink-700/60 bg-ink-900/60 backdrop-blur-xl">
      <div className="flex items-center gap-2.5 px-6 py-6">
        <div className="flex h-9 w-9 items-center justify-center rounded-xl bg-gradient-to-br from-brand-400 to-accent-teal shadow-lg shadow-brand-600/30">
          <Wand2 className="h-5 w-5 text-ink-950" strokeWidth={2.5} />
        </div>
        <div>
          <div className="text-sm font-semibold tracking-wide text-ink-100">CMS AI Quoting</div>
          <div className="text-[11px] text-ink-400">Mold Geometry Console</div>
        </div>
      </div>

      <nav className="flex-1 space-y-1 px-3 py-2">
        {NAV.map(({ to, label, icon: Icon, end }) => (
          <NavLink
            key={to}
            to={to}
            end={end}
            className={({ isActive }) =>
              clsx(
                "flex items-center gap-3 rounded-xl px-3 py-2.5 text-sm font-medium transition-colors",
                isActive
                  ? "bg-gradient-to-r from-brand-600/30 to-brand-500/10 text-ink-100 shadow-inner shadow-brand-500/10 ring-1 ring-brand-500/30"
                  : "text-ink-300 hover:bg-ink-800/70 hover:text-ink-100"
              )
            }
          >
            <Icon className="h-4.5 w-4.5" size={18} strokeWidth={2} />
            {label}
          </NavLink>
        ))}
      </nav>

      <div className="mx-3 mb-4 rounded-xl border border-ink-700/60 bg-ink-850/70 p-3.5">
        <div className="text-[11px] font-medium uppercase tracking-wider text-ink-400">AI Bridge</div>
        <p className="mt-1.5 text-xs leading-relaxed text-ink-300">
          Every classified job auto-exports resolved part names for
          <span className="font-semibold text-ink-100"> Module6121</span> to read.
        </p>
      </div>
    </aside>
  );
}
