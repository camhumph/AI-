import { useEffect, useState } from "react";
import { Save, Mail, Send, FolderCog, CheckCircle2, XCircle, Brain, Play } from "lucide-react";
import Layout from "../components/Layout";
import { Card, Badge, Button, Spinner } from "../components/ui";
import { api } from "../api/client";

type Rate = { mode: string; rate: number; minimum: number };

type EmailSettings = {
  imap_host: string;
  imap_port: number;
  imap_user: string;
  imap_password_set: boolean;
  imap_folder: string;
  imap_ssl: boolean;
  smtp_host: string;
  smtp_port: number;
  smtp_user: string;
  smtp_password_set: boolean;
  smtp_from: string;
  gmail_address: string;
  configured: boolean;
  smtp_configured: boolean;
  credentials_path: string;
};

export default function SettingsPage() {
  const [rates, setRates] = useState<Record<string, Rate> | null>(null);
  const [emailSettings, setEmailSettings] = useState<EmailSettings | null>(null);
  const [imapPassword, setImapPassword] = useState("");
  const [smtpPassword, setSmtpPassword] = useState("");
  const [emailSaving, setEmailSaving] = useState(false);
  const [emailSaved, setEmailSaved] = useState(false);
  const [trainingStatus, setTrainingStatus] = useState<{ jobs_processed?: number; jobs_ok?: number } | null>(null);
  const [trainingRunning, setTrainingRunning] = useState(false);
  const [jobsRoot, setJobsRoot] = useState("");

  useEffect(() => {
    api.getPricing().then(setRates);
    api.getEmailSettings().then(setEmailSettings);
    api.trainingStatus().then(setTrainingStatus).catch(() => setTrainingStatus(null));
  }, []);

  const saveEmail = async () => {
    if (!emailSettings) return;
    setEmailSaving(true);
    try {
      const updated = await api.putEmailSettings({
        imap_host: emailSettings.imap_host,
        imap_port: emailSettings.imap_port,
        imap_user: emailSettings.imap_user,
        imap_password: imapPassword,
        imap_folder: emailSettings.imap_folder,
        imap_ssl: emailSettings.imap_ssl,
        smtp_host: emailSettings.smtp_host,
        smtp_port: emailSettings.smtp_port,
        smtp_user: emailSettings.smtp_user,
        smtp_password: smtpPassword,
        smtp_from: emailSettings.smtp_from,
        gmail_address: emailSettings.gmail_address,
      });
      setEmailSettings(updated);
      setImapPassword("");
      setSmtpPassword("");
      setEmailSaved(true);
      setTimeout(() => setEmailSaved(false), 2000);
    } finally {
      setEmailSaving(false);
    }
  };

  const [emailTestMsg, setEmailTestMsg] = useState("");

  const testEmail = async () => {
    setEmailTestMsg("");
    try {
      const r = await api.testEmail();
      setEmailTestMsg(r.message);
    } catch (e) {
      setEmailTestMsg((e as Error).message);
    }
  };

  const runTraining = async () => {
    setTrainingRunning(true);
    try {
      const result = await api.runTraining(jobsRoot || undefined);
      setTrainingStatus(result);
    } finally {
      setTrainingRunning(false);
    }
  };

  return (
    <Layout title="Settings" subtitle="Email credentials, AI training, pricing, and Module6121 bridge.">
      <div className="grid grid-cols-1 gap-6 lg:grid-cols-2">
        <Card className="p-5 lg:col-span-2">
          <div className="mb-4 flex items-center justify-between">
            <div className="flex items-center gap-2">
              <Mail className="h-4 w-4 text-ink-300" />
              <h3 className="text-sm font-semibold text-ink-100">Gmail Credentials</h3>
            </div>
            <div className="flex items-center gap-2">
              <StatusRow ok={!!emailSettings?.configured} label={emailSettings?.configured ? "Connected" : "Not connected"} />
              <Button variant="secondary" onClick={testEmail} disabled={!emailSettings?.configured}>
                Test
              </Button>
              <Button onClick={saveEmail} disabled={emailSaving || !emailSettings}>
                <Save className="h-4 w-4" /> {emailSaving ? "Saving..." : emailSaved ? "Saved" : "Save"}
              </Button>
            </div>
          </div>
          <p className="mb-4 text-xs text-ink-400">
            Enter your Gmail app password here only — not in <code className="text-ink-200">gmail_app_password.txt</code>.
            Module6121 and the inbox both read from{" "}
            <code className="text-ink-200">{emailSettings?.credentials_path || "cms_data/email_credentials.json"}</code> on this PC.
          </p>
          {emailTestMsg && <p className="mb-3 text-xs text-ink-300">{emailTestMsg}</p>}
          {!emailSettings ? (
            <Spinner label="Loading email settings..." />
          ) : (
            <div className="grid grid-cols-1 gap-3 md:grid-cols-2">
              <Field label="Gmail address" value={emailSettings.gmail_address} onChange={(v) => setEmailSettings({ ...emailSettings, gmail_address: v })} />
              <Field label="IMAP user (usually same as Gmail)" value={emailSettings.imap_user} onChange={(v) => setEmailSettings({ ...emailSettings, imap_user: v, smtp_user: v })} />
              <Field label="IMAP app password" type="password" placeholder={emailSettings.imap_password_set ? "•••••••• (saved — leave blank to keep)" : "16-character app password"} value={imapPassword} onChange={setImapPassword} />
              <Field label="SMTP app password" type="password" placeholder={emailSettings.smtp_password_set ? "•••••••• (saved — leave blank to keep)" : "same app password"} value={smtpPassword} onChange={setSmtpPassword} />
            </div>
          )}
        </Card>

        <Card className="p-5">
          <div className="mb-3 flex items-center gap-2">
            <Send className="h-4 w-4 text-ink-300" />
            <h3 className="text-sm font-semibold text-ink-100">SMTP Replies</h3>
          </div>
          <StatusRow ok={!!emailSettings?.smtp_configured} label={emailSettings?.smtp_configured ? "Ready to send" : "Save credentials above"} />
        </Card>

        <Card className="p-5">
          <div className="mb-3 flex items-center gap-2">
            <FolderCog className="h-4 w-4 text-ink-300" />
            <h3 className="text-sm font-semibold text-ink-100">Module6121 Bridge</h3>
          </div>
          <StatusRow ok label="Active at http://127.0.0.1:8000" />
        </Card>

        <Card className="p-5 lg:col-span-2">
          <div className="mb-4 flex items-center justify-between">
            <div className="flex items-center gap-2">
              <Brain className="h-4 w-4 text-ink-300" />
              <h3 className="text-sm font-semibold text-ink-100">AI Training from Quote + Steel Sheets</h3>
            </div>
            <Button onClick={runTraining} disabled={trainingRunning}>
              <Play className="h-4 w-4" /> {trainingRunning ? "Training..." : "Run Training Scan"}
            </Button>
          </div>
          <p className="mb-3 text-xs text-ink-400">
            Point at a folder of completed jobs (each with XT export + quote/steel Excel). The AI works backwards:
            reads plate names from your steel sheets, matches them to CAD dimensions, and writes CORRECT_ME training files.
          </p>
          <Field
            label="Jobs root folder (optional)"
            value={jobsRoot}
            onChange={setJobsRoot}
            placeholder="C:\CMS_Local_Workspace\AI_Jobs"
          />
          {trainingStatus && (
            <p className="mt-3 text-xs text-ink-300">
              Last run: {trainingStatus.jobs_ok ?? 0} / {trainingStatus.jobs_processed ?? 0} jobs trained successfully.
            </p>
          )}
        </Card>
      </div>

      <Card className="mt-6 p-5">
        <div className="section-label mb-4">Purchased Component Prices (CSV)</div>
        <p className="mb-4 text-xs text-ink-400">
          Quote totals pull directly from Purchased Components Prices.csv and per-job Purchased Components Quote.csv.
        </p>
        {!rates ? (
          <Spinner label="Loading CSV prices..." />
        ) : Object.keys(rates).length === 0 ? (
          <p className="text-xs text-ink-400">No prices in CSV yet.</p>
        ) : (
          <div className="scrollbar-thin overflow-x-auto">
            <table className="w-full text-left text-sm">
              <thead>
                <tr className="text-[10px] uppercase tracking-widest text-ink-500">
                  <th className="px-3 py-2">Component</th>
                  <th className="px-3 py-2">Unit Price</th>
                </tr>
              </thead>
              <tbody>
                {Object.entries(rates).map(([comp, spec]) => (
                  <tr key={comp} className="border-t border-ink-800">
                    <td className="px-3 py-2 text-ink-200">{comp}</td>
                    <td className="px-3 py-2 text-ink-100">${spec.rate.toFixed(2)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </Card>
    </Layout>
  );
}

function Field({
  label,
  value,
  onChange,
  type = "text",
  placeholder = "",
}: {
  label: string;
  value: string;
  onChange: (v: string) => void;
  type?: string;
  placeholder?: string;
}) {
  return (
    <label className="block text-xs text-ink-400">
      {label}
      <input
        type={type}
        value={value}
        placeholder={placeholder}
        onChange={(e) => onChange(e.target.value)}
        className="mt-1 w-full rounded-lg border border-ink-700/60 bg-ink-850 px-3 py-2 text-sm text-ink-100 placeholder:text-ink-500"
      />
    </label>
  );
}

function StatusRow({ ok, label }: { ok: boolean; label: string }) {
  return (
    <div className="flex items-center gap-2">
      {ok ? <CheckCircle2 className="h-4 w-4 text-accent-green" /> : <XCircle className="h-4 w-4 text-ink-500" />}
      <Badge tone={ok ? "success" : "neutral"}>{label}</Badge>
    </div>
  );
}
