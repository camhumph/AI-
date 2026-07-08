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
    <aside className="hidden md:flex md:w-56 shrink-0 flex-col border-r border-ink-700 bg-ink-950">
      <div className="px-6 py-8">
        <div className="text-[10px] font-semibold uppercase tracking-[0.35em] text-ink-400">CMS</div>
        <div className="mt-1 text-lg font-light uppercase tracking-[0.15em] text-ink-100">Quoting</div>
      </div>

      <nav className="flex-1 space-y-0 px-0 py-2">
        {NAV.map(({ to, label, icon: Icon, end }) => (
          <NavLink
            key={to}
            to={to}
            end={end}
            className={({ isActive }) =>
              clsx(
                "relative flex items-center gap-3 border-l-2 px-6 py-3 text-xs font-semibold uppercase tracking-widest transition-colors",
                isActive
                  ? "border-ink-100 bg-ink-900 text-ink-100"
                  : "border-transparent text-ink-400 hover:border-ink-500 hover:text-ink-200"
              )
            }
          >
            <Icon size={16} strokeWidth={1.5} />
            {label}
          </NavLink>
        ))}
      </nav>

      <div className="border-t border-ink-700 px-6 py-5">
        <div className="flex items-center gap-2">
          <span className="h-1.5 w-1.5 rounded-full bg-accent-green" />
          <span className="text-[10px] font-semibold uppercase tracking-widest text-ink-400">Local</span>
        </div>
        <p className="mt-2 text-[10px] leading-relaxed text-ink-500">127.0.0.1:8000</p>
      </div>
    </aside>
  );
}
