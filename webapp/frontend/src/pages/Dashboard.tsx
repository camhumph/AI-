import { useEffect, useState } from "react";
import { Link } from "react-router-dom";
import { FileStack, Mail, Image as ImageIcon, Box, ArrowRight, Sparkles } from "lucide-react";
import Layout from "../components/Layout";
import { Card, StatCard, Badge, Spinner } from "../components/ui";
import { api, type JobSummary } from "../api/client";

export default function Dashboard() {
  const [jobs, setJobs] = useState<JobSummary[] | null>(null);
  const [emailConfigured, setEmailConfigured] = useState(false);

  useEffect(() => {
    api.listJobs().then(setJobs).catch(() => setJobs([]));
    api.emailStatus().then((s) => setEmailConfigured(s.configured)).catch(() => {});
  }, []);

  const totalParts = jobs?.reduce((s, j) => s + j.part_count, 0) ?? 0;
  const sequenced = jobs?.filter((j) => j.sequenced_latch_lock_base).length ?? 0;
  const bmsCount = jobs?.filter((j) => j.base_type === "bms").length ?? 0;

  return (
    <Layout
      title="Dashboard"
      subtitle="Your mold-base quoting pipeline at a glance."
      actions={
        <Link to="/quotes">
          <button className="inline-flex items-center gap-2 rounded-xl bg-gradient-to-r from-brand-500 to-brand-600 px-4 py-2 text-sm font-medium text-white shadow-lg shadow-brand-600/30 hover:brightness-110">
            <Sparkles className="h-4 w-4" /> New Quote
          </button>
        </Link>
      }
    >
      <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-4">
        <StatCard label="Active Jobs" value={jobs?.length ?? "--"} sub="Quote-ready folders" accent="brand" />
        <StatCard label="AI-Classified Parts" value={totalParts.toLocaleString()} sub="Across standard bases" accent="teal" />
        <StatCard
          label="Latch-Lock Bases"
          value={sequenced}
          sub="Plate-sequenced / secondary parting lines"
          accent="rose"
        />
        <StatCard
          label="BMS Bases"
          value={bmsCount}
          sub="BOM-driven by Module6121 (AI off)"
          accent="amber"
        />
      </div>

      <div className="mt-6 grid grid-cols-1 gap-6 lg:grid-cols-3">
        <Card className="lg:col-span-2 p-5">
          <div className="mb-4 flex items-center justify-between">
            <h2 className="text-sm font-semibold text-ink-100">Recent Jobs</h2>
            <Link to="/quotes" className="flex items-center gap-1 text-xs font-medium text-brand-400 hover:text-brand-300">
              View all <ArrowRight className="h-3.5 w-3.5" />
            </Link>
          </div>
          {jobs === null ? (
            <Spinner label="Loading jobs..." />
          ) : jobs.length === 0 ? (
            <p className="text-sm text-ink-400">No jobs yet. Create one from the Quotes page.</p>
          ) : (
            <div className="space-y-2">
              {jobs.slice(0, 6).map((job) => (
                <Link
                  key={job.job_id}
                  to={`/quotes/${job.job_id}`}
                  className="flex items-center justify-between rounded-xl border border-ink-700/50 bg-ink-850/40 px-4 py-3 transition hover:border-brand-500/40 hover:bg-ink-800/50"
                >
                  <div>
                    <div className="flex items-center gap-2">
                      <span className="text-sm font-medium text-ink-100">{job.display_name}</span>
                      {job.base_type === "bms" ? (
                        <Badge tone="warning">BMS</Badge>
                      ) : (
                        job.sequenced_latch_lock_base && <Badge tone="brand">Latch-Lock Base</Badge>
                      )}
                    </div>
                    <div className="mt-1 flex items-center gap-3 text-xs text-ink-400">
                      <span className="flex items-center gap-1">
                        <FileStack className="h-3.5 w-3.5" /> {job.part_count} parts
                      </span>
                      <span className="flex items-center gap-1">
                        <ImageIcon className="h-3.5 w-3.5" /> {job.image_count} views
                      </span>
                      <span className="flex items-center gap-1">
                        <Box className="h-3.5 w-3.5" /> {job.model_count} STL
                      </span>
                    </div>
                  </div>
                  <ArrowRight className="h-4 w-4 text-ink-500" />
                </Link>
              ))}
            </div>
          )}
        </Card>

        <Card className="p-5">
          <div className="mb-4 flex items-center gap-2">
            <Mail className="h-4 w-4 text-ink-300" />
            <h2 className="text-sm font-semibold text-ink-100">Inbox</h2>
          </div>
          {emailConfigured ? (
            <p className="text-sm text-ink-300">
              Email is connected. Head to the Inbox to read messages and quote directly from them.
            </p>
          ) : (
            <div className="space-y-3">
              <p className="text-sm text-ink-400">
                Connect an inbox to read customer emails and quote them with one click.
              </p>
              <Link
                to="/settings"
                className="inline-flex items-center gap-1 text-xs font-medium text-brand-400 hover:text-brand-300"
              >
                Configure email <ArrowRight className="h-3.5 w-3.5" />
              </Link>
            </div>
          )}
          <Link to="/email">
            <button className="mt-4 w-full rounded-xl bg-ink-800 py-2 text-sm font-medium text-ink-100 ring-1 ring-ink-600/60 hover:bg-ink-700">
              Open Inbox
            </button>
          </Link>
        </Card>
      </div>
    </Layout>
  );
}
