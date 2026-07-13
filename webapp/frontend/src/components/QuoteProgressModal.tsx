import { Card, Button } from "./ui";
import type { QuoteRunStatus } from "../api/client";

export default function QuoteProgressModal({
  status,
  onClose,
}: {
  status: QuoteRunStatus;
  onClose: () => void;
}) {
  const steps = [
    { key: "queued", label: "Prepare files" },
    { key: "starting", label: "DME price lookup" },
    { key: "launching", label: "Launch SolidWorks" },
    { key: "running", label: "Module6121 + AI classify" },
    { key: "completed", label: "Done" },
  ];
  const phaseOrder = ["queued", "starting", "launching", "running", "completed"];
  const currentIdx = phaseOrder.indexOf(status.phase);

  return (
    <div className="overlay-backdrop fixed inset-0 z-50 flex items-center justify-center p-4">
      <Card className="w-full max-w-md p-6">
        <div className="section-label mb-2">Quote in progress</div>
        <p className="text-sm text-ink-300">{status.message || "Running full CMS quote pipeline..."}</p>
        {status.warning && (
          <p className="mt-2 text-xs text-accent-amber">{status.warning}</p>
        )}
        {status.stuck_reason && (
          <p className="mt-2 text-xs text-accent-rose">{status.stuck_reason}</p>
        )}
        {status.diagnostics?.launcher_log_tail && (
          <pre className="mt-2 max-h-32 overflow-auto whitespace-pre-wrap break-words rounded-lg bg-black/30 p-2 font-mono text-[9px] text-ink-400">
            {status.diagnostics.launcher_log_tail}
          </pre>
        )}
        {status.job_id && (
          <p className="mt-1 font-mono text-xs text-ink-500">C-number: {status.job_id}</p>
        )}
        <ul className="mt-5 space-y-2">
          {steps.map((s, i) => {
            const done = currentIdx >= i;
            const active = currentIdx === i;
            return (
              <li key={s.key} className={`flex items-center gap-3 text-xs ${done ? "text-ink-100" : "text-ink-500"}`}>
                <span className={`h-2 w-2 rounded-full ${active ? "bg-accent-teal animate-pulse" : done ? "bg-accent-green" : "bg-ink-700/40"}`} />
                {s.label}
              </li>
            );
          })}
        </ul>
        <p className="mt-4 text-[10px] text-ink-500">
          Opens XT/CAD, fills quote + steel sheets, looks up DME prices, runs AI for standard bases.
        </p>
        <Button variant="ghost" className="mt-4 w-full" onClick={onClose}>
          Run in background
        </Button>
      </Card>
    </div>
  );
}
