import { useEffect, useState } from "react";
import { Link } from "react-router-dom";
import { FileStack, Mail, ArrowRight } from "lucide-react";
import Layout from "../components/Layout";
import { Card, StatCard, Button, Spinner } from "../components/ui";
import { api, type JobSummary } from "../api/client";

export default function Dashboard() {
  const [jobs, setJobs] = useState<JobSummary[] | null>(null);
  const [emailConfigured, setEmailConfigured] = useState(false);

  useEffect(() => {
    api.listJobs().then(setJobs).catch(() => setJobs([]));
    api.emailStatus().then((s) => setEmailConfigured(s.configured)).catch(() => {});
  }, []);

  const totalParts = jobs?.reduce((s, j) => s + j.part_count, 0) ?? 0;

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
      <div className="grid grid-cols-1 gap-4 sm:grid-cols-3">
        <StatCard label="Active Quotes" value={jobs?.length ?? "--"} sub="Registered C-number jobs" />
        <StatCard label="Classified Parts" value={totalParts.toLocaleString()} sub="Across all quotes" />
        <StatCard label="Inbox" value={emailConfigured ? "Connected" : "Setup"} sub={emailConfigured ? "Ready" : "Configure in Settings"} />
      </div>

      <div className="mt-8 grid grid-cols-1 gap-6 lg:grid-cols-3">
        <Card className="lg:col-span-2 p-5">
          <div className="mb-4 flex items-center justify-between">
            <div className="section-label">Recent Quotes</div>
            <Link to="/quotes" className="flex items-center gap-1 text-[10px] font-semibold uppercase tracking-widest text-ink-400 hover:text-ink-100">
              View all <ArrowRight className="h-3.5 w-3.5" />
            </Link>
          </div>
          {jobs === null ? (
            <Spinner label="Loading..." />
          ) : jobs.length === 0 ? (
            <p className="text-xs text-ink-400">No quotes yet. Browse to a C-number folder to start.</p>
          ) : (
            <div className="divide-y divide-ink-800">
              {jobs.slice(0, 8).map((job) => (
                <Link key={job.job_id} to={`/quotes/${job.job_id}`}
                  className="flex items-center justify-between py-3 transition hover:text-ink-100">
                  <div>
                    <div className="text-sm font-semibold uppercase tracking-wider text-ink-100">{job.job_id}</div>
                    <div className="text-xs text-ink-500">{job.display_name}</div>
                  </div>
                  <div className="flex items-center gap-1 text-xs text-ink-500">
                    <FileStack className="h-3.5 w-3.5" /> {job.part_count}
                    <ArrowRight className="ml-2 h-3.5 w-3.5" />
                  </div>
                </Link>
              ))}
            </div>
          )}
        </Card>

        <Card className="p-5">
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
