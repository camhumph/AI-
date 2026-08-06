import { NavLink } from "react-router-dom";
import { LayoutDashboard, Mail, FileStack, Radar, Settings } from "lucide-react";
import clsx from "clsx";

const NAV = [
  { to: "/", label: "Dashboard", icon: LayoutDashboard, end: true },
  { to: "/email", label: "Inbox", icon: Mail },
  { to: "/quotes", label: "Quotes", icon: FileStack },
  { to: "/matching", label: "Matching", icon: Radar },
  { to: "/settings", label: "Settings", icon: Settings },
];

export function MobileDock() {
  return (
    <nav className="glass-panel-strong fixed bottom-4 left-1/2 z-40 flex -translate-x-1/2 gap-0.5 rounded-full px-2 py-2 md:hidden">
      {NAV.map(({ to, label, icon: Icon, end }) => (
        <NavLink
          key={to}
          to={to}
          end={end}
          title={label}
          className={({ isActive }) =>
            clsx(
              "flex items-center justify-center rounded-full p-3 transition-all",
              isActive
                ? "bg-white/12 text-ink-100 shadow-[0_0_16px_rgba(10,132,255,0.2)]"
                : "text-ink-500 hover:text-ink-300"
            )
          }
        >
          <Icon size={20} strokeWidth={1.75} />
        </NavLink>
      ))}
    </nav>
  );
}

export default function Sidebar() {
  return (
    <aside className="glass-panel-strong m-3 hidden w-[4.5rem] shrink-0 flex-col rounded-3xl md:flex lg:w-52">
      <div className="px-4 py-6 lg:px-6 lg:py-8">
        <div className="hidden text-[9px] font-bold uppercase tracking-[0.45em] text-ink-500 lg:block">CMS</div>
        <div className="mt-0 text-center text-[10px] font-bold uppercase tracking-[0.2em] text-ink-300 lg:mt-1 lg:text-left lg:text-base lg:text-ink-100">
          <span className="lg:hidden">C</span>
          <span className="hidden lg:inline">Quoting</span>
        </div>
      </div>

      <nav className="flex flex-1 flex-col gap-1 px-2 py-2 lg:px-3">
        {NAV.map(({ to, label, icon: Icon, end }) => (
          <NavLink
            key={to}
            to={to}
            end={end}
            title={label}
            className={({ isActive }) =>
              clsx(
                "group flex items-center justify-center gap-3 rounded-2xl px-3 py-3 text-[10px] font-bold uppercase tracking-[0.14em] transition-all lg:justify-start lg:px-4",
                isActive
                  ? "bg-white/12 text-ink-100 shadow-[inset_0_1px_0_rgba(255,255,255,0.1)]"
                  : "text-ink-500 hover:bg-white/6 hover:text-ink-300"
              )
            }
          >
            <Icon size={18} strokeWidth={1.75} className="shrink-0" />
            <span className="hidden lg:inline">{label}</span>
          </NavLink>
        ))}
      </nav>

      <div className="border-t border-white/8 px-4 py-4 lg:px-6 lg:py-5">
        <div className="flex items-center justify-center gap-2 lg:justify-start">
          <span className="h-1.5 w-1.5 rounded-full bg-accent-green shadow-[0_0_8px_rgba(48,209,88,0.8)] animate-pulse" />
          <span className="hidden text-[9px] font-bold uppercase tracking-[0.2em] text-ink-500 lg:inline">
            Online
          </span>
        </div>
        <p className="mt-2 hidden font-mono text-[9px] text-ink-600 lg:block">127.0.0.1:8000</p>
      </div>
    </aside>
  );
}
