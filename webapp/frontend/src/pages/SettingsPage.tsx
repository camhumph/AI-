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
  const [saving, setSaving] = useState(false);
  const [saved, setSaved] = useState(false);
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

  const updateRate = (role: string, field: keyof Rate, value: string) => {
    setRates((prev) => {
      if (!prev) return prev;
      const numeric = field === "mode" ? value : parseFloat(value) || 0;
      return { ...prev, [role]: { ...prev[role], [field]: numeric } };
    });
  };

  const save = async () => {
    if (!rates) return;
    setSaving(true);
    try {
      await api.putPricing(rates);
      setSaved(true);
      setTimeout(() => setSaved(false), 2000);
    } finally {
      setSaving(false);
    }
  };

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
              <StatusRow ok={!!emailSettings?.configured} label={emailSettings?.configured ? "Inbox connected" : "Not connected"} />
              <Button onClick={saveEmail} disabled={emailSaving || !emailSettings}>
                <Save className="h-4 w-4" /> {emailSaving ? "Saving..." : emailSaved ? "Saved!" : "Save Email"}
              </Button>
            </div>
          </div>
          <p className="mb-4 text-xs text-ink-400">
            Enter your Gmail app password here only — not in <code className="text-ink-200">gmail_app_password.txt</code>.
            Module6121 and the inbox both read from{" "}
            <code className="text-ink-200">{emailSettings?.credentials_path || "cms_data/email_credentials.json"}</code> on this PC.
          </p>
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
        <div className="mb-4 flex items-center justify-between">
          <div>
            <h3 className="text-sm font-semibold text-ink-100">Pricing Rules</h3>
            <p className="text-xs text-ink-400">Rates drive the Total Price on every quote.</p>
          </div>
          <Button onClick={save} disabled={saving || !rates}>
            <Save className="h-4 w-4" /> {saving ? "Saving..." : saved ? "Saved!" : "Save Rates"}
          </Button>
        </div>

        {!rates ? (
          <Spinner label="Loading pricing config..." />
        ) : (
          <div className="scrollbar-thin overflow-x-auto">
            <table className="w-full text-left text-sm">
              <thead>
                <tr className="text-[11px] uppercase tracking-wider text-ink-400">
                  <th className="px-3 py-2 font-medium">Role</th>
                  <th className="px-3 py-2 font-medium">Pricing Mode</th>
                  <th className="px-3 py-2 font-medium">Rate</th>
                  <th className="px-3 py-2 font-medium">Minimum</th>
                </tr>
              </thead>
              <tbody>
                {Object.entries(rates).map(([role, spec]) => (
                  <tr key={role} className="border-t border-ink-800/60">
                    <td className="px-3 py-2 text-ink-200">{role.replace(/_/g, " ")}</td>
                    <td className="px-3 py-2">
                      <select
                        value={spec.mode}
                        onChange={(e) => updateRate(role, "mode", e.target.value)}
                        className="rounded-lg border border-ink-700/60 bg-ink-850 px-2 py-1 text-xs text-ink-100"
                      >
                        <option value="flat">flat</option>
                        <option value="per_cuin">per cubic inch</option>
                        <option value="per_inch">per inch (length)</option>
                      </select>
                    </td>
                    <td className="px-3 py-2">
                      <input
                        type="number"
                        step="0.01"
                        value={spec.rate}
                        onChange={(e) => updateRate(role, "rate", e.target.value)}
                        className="w-24 rounded-lg border border-ink-700/60 bg-ink-850 px-2 py-1 text-xs text-ink-100"
                      />
                    </td>
                    <td className="px-3 py-2">
                      <input
                        type="number"
                        step="0.01"
                        value={spec.minimum}
                        onChange={(e) => updateRate(role, "minimum", e.target.value)}
                        className="w-24 rounded-lg border border-ink-700/60 bg-ink-850 px-2 py-1 text-xs text-ink-100"
                      />
                    </td>
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
