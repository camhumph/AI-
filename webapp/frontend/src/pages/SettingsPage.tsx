import { useEffect, useState } from "react";
import { Save, Mail, Send, FolderCog, CheckCircle2, XCircle } from "lucide-react";
import Layout from "../components/Layout";
import { Card, Badge, Button, Spinner } from "../components/ui";
import { api } from "../api/client";

type Rate = { mode: string; rate: number; minimum: number };

export default function SettingsPage() {
  const [rates, setRates] = useState<Record<string, Rate> | null>(null);
  const [saving, setSaving] = useState(false);
  const [saved, setSaved] = useState(false);
  const [emailStatus, setEmailStatus] = useState<{ configured: boolean; smtp_configured: boolean; imap_host: string | null; imap_user: string | null } | null>(null);

  useEffect(() => {
    api.getPricing().then(setRates);
    api.emailStatus().then(setEmailStatus);
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

  return (
    <Layout title="Settings" subtitle="Pricing rules, email connection, and Module6121 bridge status.">
      <div className="grid grid-cols-1 gap-6 lg:grid-cols-3">
        <Card className="p-5">
          <div className="mb-3 flex items-center gap-2">
            <Mail className="h-4 w-4 text-ink-300" />
            <h3 className="text-sm font-semibold text-ink-100">Email (IMAP)</h3>
          </div>
          <StatusRow ok={!!emailStatus?.configured} label={emailStatus?.configured ? `Connected as ${emailStatus.imap_user}` : "Not configured"} />
          <p className="mt-3 text-xs leading-relaxed text-ink-400">
            Set <code className="text-ink-200">CMS_IMAP_HOST</code>, <code className="text-ink-200">CMS_IMAP_USER</code>,{" "}
            <code className="text-ink-200">CMS_IMAP_PASSWORD</code> as secrets to read your inbox.
          </p>
        </Card>

        <Card className="p-5">
          <div className="mb-3 flex items-center gap-2">
            <Send className="h-4 w-4 text-ink-300" />
            <h3 className="text-sm font-semibold text-ink-100">Email (SMTP reply)</h3>
          </div>
          <StatusRow ok={!!emailStatus?.smtp_configured} label={emailStatus?.smtp_configured ? "Connected" : "Not configured"} />
          <p className="mt-3 text-xs leading-relaxed text-ink-400">
            Set <code className="text-ink-200">CMS_SMTP_HOST</code>, <code className="text-ink-200">CMS_SMTP_USER</code>,{" "}
            <code className="text-ink-200">CMS_SMTP_PASSWORD</code> as secrets to send replies.
          </p>
        </Card>

        <Card className="p-5">
          <div className="mb-3 flex items-center gap-2">
            <FolderCog className="h-4 w-4 text-ink-300" />
            <h3 className="text-sm font-semibold text-ink-100">Module6121 Bridge</h3>
          </div>
          <StatusRow ok label="Active -- exports on every quote view" />
          <p className="mt-3 text-xs leading-relaxed text-ink-400">
            Files land in <code className="text-ink-200">CMS_VBA_BRIDGE_DIR</code> (defaults to{" "}
            <code className="text-ink-200">backend/data/vba_bridge</code>). Point the macro at that
            folder, or use the per-job <code className="text-ink-200">/api/bridge/&#123;job&#125;</code> endpoint.
          </p>
        </Card>
      </div>

      <Card className="mt-6 p-5">
        <div className="mb-4 flex items-center justify-between">
          <div>
            <h3 className="text-sm font-semibold text-ink-100">Pricing Rules</h3>
            <p className="text-xs text-ink-400">
              Placeholder rates. Replace with your real CMS price book -- these drive the Total Price on every quote.
            </p>
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

function StatusRow({ ok, label }: { ok: boolean; label: string }) {
  return (
    <div className="flex items-center gap-2">
      {ok ? <CheckCircle2 className="h-4 w-4 text-accent-green" /> : <XCircle className="h-4 w-4 text-ink-500" />}
      <Badge tone={ok ? "success" : "neutral"}>{label}</Badge>
    </div>
  );
}
