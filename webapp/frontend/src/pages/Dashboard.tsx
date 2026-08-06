import { useEffect, useState } from "react";
import { Link } from "react-router-dom";
import { FileStack, Mail, ArrowRight } from "lucide-react";
import Layout from "../components/Layout";
import { Card, Button, Badge } from "../components/ui";
import { api, type JobSummary } from "../api/client";

/** A headline figure. Skeleton while loading so the layout never jumps. */
function HeroStat({
  label, value, sub, className, small,
}: {
  label: string;
  value: string | null;
  sub: string;
  className?: string;
  small?: boolean;
}) {
  return (
    <div className={`glass-panel card-lift relative overflow-hidden rounded-2xl p-5 ${className || ""}`}>
      <div className="section-label">{label}</div>
      {value === null ? (
        <div className="skeleton mt-3 h-9 w-24" />
      ) : (
        <div className={`mt-2 ${small ? "num-lg" : "num-xl"}`}>{value}</div>
      )}
      <div className="mt-1.5 text-[11px] text-ink-500">{sub}</div>
    </div>
  );
}

export default function Dashboard() {
  const [jobs, setJobs] = useState<JobSummary[] | null>(null);
  const [emailConfigured, setEmailConfigured] = useState(false);

  useEffect(() => {
    api.listJobs().then(setJobs).catch(() => setJobs([]));
    api.emailStatus().then((s) => setEmailConfigured(s.configured)).catch(() => {});
  }, []);

  const totalParts = jobs?.reduce((s, j) => s + j.part_count, 0) ?? 0;

  const recentJobs = jobs
    ? [...jobs].sort(
        (a, b) => new Date(b.updated_at).getTime() - new Date(a.updated_at).getTime()
      )
    : null;

  return (
    <Layout
      title="Dashboard"
      subtitle="Mold quoting pipeline"
      actions={
        <Link to="/quotes">
          <Button>New Quote</Button>
        </Link>
      }
    >
      <div className="horizon-line mb-6" />

      {/* Headline figures. Tabular numerals and a real elevation step, so the
          numbers read as instrument panel rather than web page. */}
      <div className="grid grid-cols-1 gap-4 sm:grid-cols-3">
        <HeroStat
          className="rise-in rise-in-1 rail-brand"
          label="Active quotes"
          value={jobs === null ? null : String(jobs.length)}
          sub="Registered C-number jobs"
        />
        <HeroStat
          className="rise-in rise-in-2"
          label="Classified parts"
          value={jobs === null ? null : totalParts.toLocaleString()}
          sub="Across all quotes"
        />
        <HeroStat
          className={`rise-in rise-in-3 ${emailConfigured ? "rail-green" : "rail-amber"}`}
          label="Inbox"
          value={emailConfigured ? "Live" : "Setup"}
          sub={emailConfigured ? "Ready to quote from mail" : "Configure in Settings"}
          small
        />
      </div>

      <div className="mt-8 grid grid-cols-1 gap-6 lg:grid-cols-3">
        <Card className="rise-in rise-in-4 p-5 lg:col-span-2">
          <div className="mb-4 flex items-center justify-between">
            <div className="section-label">Recent Quotes</div>
            <Link to="/quotes" className="flex items-center gap-1 text-[10px] font-semibold uppercase tracking-widest text-ink-400 transition hover:text-ink-100">
              View all <ArrowRight className="h-3.5 w-3.5" />
            </Link>
          </div>
          {jobs === null ? (
            <div className="space-y-2">
              {[0, 1, 2, 3, 4].map((i) => <div key={i} className="skeleton h-12" />)}
            </div>
          ) : jobs.length === 0 ? (
            <p className="text-xs text-ink-400">No quotes yet. Browse to a C-number folder to start.</p>
          ) : (
            <div className="-mx-2">
              {(recentJobs || []).slice(0, 8).map((job) => (
                <Link
                  key={job.job_id}
                  to={`/quotes/${job.job_id}`}
                  className="group flex items-center justify-between rounded-xl px-2 py-2.5 transition hover:bg-white/5"
                >
                  <div className="min-w-0">
                    <div className="flex items-center gap-2">
                      <span className="text-sm font-semibold uppercase tracking-wider text-ink-100">
                        {job.job_id}
                      </span>
                      {job.base_type === "bms" && <Badge>BMS</Badge>}
                    </div>
                    <div className="truncate text-xs text-ink-500">{job.display_name}</div>
                  </div>
                  <div className="flex shrink-0 items-center gap-3 text-xs text-ink-500">
                    <span className="flex items-center gap-1">
                      <FileStack className="h-3.5 w-3.5" />
                      <span className="num">{job.part_count}</span>
                    </span>
                    <ArrowRight className="h-3.5 w-3.5 transition group-hover:translate-x-0.5 group-hover:text-ink-200" />
                  </div>
                </Link>
              ))}
            </div>
          )}
        </Card>

        <Card className="rise-in rise-in-5 p-5">
          <div className="section-label mb-4">Inbox</div>
          {emailConfigured ? (
            <p className="text-xs text-ink-400">Email connected. Open inbox to quote from messages.</p>
          ) : (
            <p className="text-xs text-ink-400">
              Enter your Gmail app password in Settings to enable the inbox.
            </p>
          )}
          <Link to={emailConfigured ? "/email" : "/settings"} className="mt-4 block">
            <Button variant="secondary" className="w-full">
              <Mail className="h-4 w-4" /> {emailConfigured ? "Open Inbox" : "Configure Email"}
            </Button>
          </Link>
        </Card>
      </div>
    </Layout>
  );
}
