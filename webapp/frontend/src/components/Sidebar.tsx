import { NavLink } from "react-router-dom";
import { LayoutDashboard, Mail, FileStack, Settings } from "lucide-react";
import clsx from "clsx";

const NAV = [
  { to: "/", label: "Dashboard", icon: LayoutDashboard, end: true },
  { to: "/email", label: "Inbox", icon: Mail },
  { to: "/quotes", label: "Quotes", icon: FileStack },
  { to: "/settings", label: "Settings", icon: Settings },
];

export default function Sidebar() {
  return (
    <aside className="hidden md:flex md:w-56 shrink-0 flex-col border-r border-ink-700/20 bg-white/50 backdrop-blur-md">
      <div className="px-6 py-8">
        <div className="text-[9px] font-bold uppercase tracking-[0.4em] text-ink-500">CMS</div>
        <div className="mt-1 text-base font-semibold uppercase tracking-[0.2em] text-ink-100">Quoting</div>
      </div>

      <nav className="flex-1 space-y-0 px-0 py-2">
        {NAV.map(({ to, label, icon: Icon, end }) => (
          <NavLink
            key={to}
            to={to}
            end={end}
            className={({ isActive }) =>
              clsx(
                "relative flex items-center gap-3 border-l-2 px-6 py-3 text-[10px] font-bold uppercase tracking-[0.16em] transition-colors",
                isActive
                  ? "border-ink-100 bg-white/80 text-ink-100 shadow-sm"
                  : "border-transparent text-ink-500 hover:border-ink-700/40 hover:text-ink-300"
              )
            }
          >
            <Icon size={15} strokeWidth={1.75} />
            {label}
          </NavLink>
        ))}
      </nav>

      <div className="border-t border-ink-700/20 px-6 py-5">
        <div className="flex items-center gap-2">
          <span className="h-1.5 w-1.5 rounded-full bg-accent-green animate-pulse" />
          <span className="text-[9px] font-bold uppercase tracking-[0.2em] text-ink-500">Systems Online</span>
        </div>
        <p className="mt-2 font-mono text-[9px] text-ink-500">127.0.0.1:8000</p>
      </div>
    </aside>
  );
}
