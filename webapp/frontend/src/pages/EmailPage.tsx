import { useEffect, useState } from "react";
import { useNavigate } from "react-router-dom";
import { Mail, Paperclip, Reply, Sparkles, RefreshCw, Send, AlertCircle } from "lucide-react";
import Layout from "../components/Layout";
import { Card, Badge, Button, Spinner, EmptyState } from "../components/ui";
import { api, type EmailSummary, type EmailDetail } from "../api/client";

export default function EmailPage() {
  const [status, setStatus] = useState<{ configured: boolean; smtp_configured: boolean } | null>(null);
  const [messages, setMessages] = useState<EmailSummary[] | null>(null);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [detail, setDetail] = useState<EmailDetail | null>(null);
  const [loadError, setLoadError] = useState("");
  const [replying, setReplying] = useState(false);
  const [replyBody, setReplyBody] = useState("");
  const [sendState, setSendState] = useState<"idle" | "sending" | "sent" | "error">("idle");
  const navigate = useNavigate();

  const refresh = () => {
    setLoadError("");
    api.listEmails().then(setMessages).catch((e) => setLoadError(e.message));
  };

  useEffect(() => {
    api.emailStatus().then(setStatus);
    refresh();
  }, []);

  useEffect(() => {
    if (!selectedId) return;
    setDetail(null);
    api.getEmail(selectedId).then(setDetail).catch((e) => setLoadError(e.message));
    setReplying(false);
    setSendState("idle");
  }, [selectedId]);

  const quoteThis = (jobTokenOrNone?: string) => {
    if (jobTokenOrNone) {
      navigate(`/quotes/${jobTokenOrNone}`);
    } else {
      navigate("/quotes");
    }
  };

  const sendReply = async () => {
    if (!detail) return;
    setSendState("sending");
    try {
      const fromAddr = detail.from.match(/<(.+)>/)?.[1] || detail.from;
      await api.replyEmail(detail.id, fromAddr, detail.subject, replyBody, detail.message_id_header || "");
      setSendState("sent");
      setReplying(false);
    } catch {
      setSendState("error");
    }
  };

  if (status && !status.configured) {
    return (
      <Layout title="Inbox" subtitle="Read and reply to customer emails, and quote them in one click.">
        <EmptyState
          icon={<Mail className="h-8 w-8" />}
          title="Connect your email to get started"
          description="Set CMS_IMAP_HOST, CMS_IMAP_USER, and CMS_IMAP_PASSWORD (an app password works for Gmail/Outlook/Office365) as secrets in the Cursor Dashboard, then reload this page. Add CMS_SMTP_* variables too if you want to send replies from here."
          action={
            <Button variant="secondary" onClick={() => window.location.reload()}>
              <RefreshCw className="h-4 w-4" /> Reload
            </Button>
          }
        />
      </Layout>
    );
  }

  return (
    <Layout title="Inbox" subtitle="Read, reply, and quote customer emails.">
      <div className="grid h-[calc(100vh-160px)] grid-cols-1 gap-4 lg:grid-cols-[340px_1fr]">
        <Card className="flex flex-col overflow-hidden">
          <div className="flex items-center justify-between border-b border-ink-700/60 px-4 py-3">
            <span className="text-sm font-semibold text-ink-100">Messages</span>
            <button onClick={refresh} className="text-ink-400 hover:text-ink-100">
              <RefreshCw className="h-4 w-4" />
            </button>
          </div>
          <div className="scrollbar-thin flex-1 overflow-y-auto">
            {messages === null ? (
              <div className="p-4"><Spinner label="Loading inbox..." /></div>
            ) : loadError ? (
              <div className="flex items-start gap-2 p-4 text-xs text-accent-rose"><AlertCircle className="h-4 w-4 shrink-0" /> {loadError}</div>
            ) : messages.length === 0 ? (
              <p className="p-4 text-sm text-ink-400">No messages found.</p>
            ) : (
              messages.map((m) => (
                <button
                  key={m.id}
                  onClick={() => setSelectedId(m.id)}
                  className={`block w-full border-b border-ink-800/60 px-4 py-3 text-left transition ${
                    selectedId === m.id ? "bg-ink-800/70" : "hover:bg-ink-850/60"
                  }`}
                >
                  <div className="flex items-center justify-between">
                    <span className="truncate text-xs font-medium text-ink-200">{m.from}</span>
                    <span className="shrink-0 text-[10px] text-ink-500">{formatDate(m.date)}</span>
                  </div>
                  <div className="mt-0.5 truncate text-sm text-ink-100">{m.subject || "(no subject)"}</div>
                  {m.matched_jobs.length > 0 && (
                    <div className="mt-1.5 flex gap-1">
                      {m.matched_jobs.map((j) => (
                        <Badge key={j} tone="brand">{j}</Badge>
                      ))}
                    </div>
                  )}
                </button>
              ))
            )}
          </div>
        </Card>

        <Card className="flex flex-col overflow-hidden">
          {!selectedId ? (
            <div className="flex flex-1 items-center justify-center text-sm text-ink-500">
              Select a message to read it
            </div>
          ) : !detail ? (
            <div className="p-6"><Spinner label="Loading message..." /></div>
          ) : (
            <div className="flex flex-1 flex-col overflow-hidden">
              <div className="border-b border-ink-700/60 px-6 py-4">
                <div className="flex items-start justify-between gap-4">
                  <div className="min-w-0">
                    <h2 className="truncate text-base font-semibold text-ink-100">{detail.subject || "(no subject)"}</h2>
                    <p className="mt-1 text-xs text-ink-400">
                      From <span className="text-ink-200">{detail.from}</span> &middot; {formatDate(detail.date)}
                    </p>
                  </div>
                  <div className="flex shrink-0 gap-2">
                    <Button variant="secondary" onClick={() => setReplying((r) => !r)}>
                      <Reply className="h-4 w-4" /> Reply
                    </Button>
                    <Button onClick={() => quoteThis(detail.matched_jobs[0])}>
                      <Sparkles className="h-4 w-4" /> Quote This
                    </Button>
                  </div>
                </div>
                {detail.matched_jobs.length > 0 && (
                  <p className="mt-2 text-xs text-ink-400">
                    Detected job reference{detail.matched_jobs.length > 1 ? "s" : ""}:{" "}
                    {detail.matched_jobs.map((j) => (
                      <Badge key={j} tone="brand">{j}</Badge>
                    ))}
                  </p>
                )}
              </div>

              <div className="scrollbar-thin flex-1 overflow-y-auto px-6 py-4">
                {detail.body_html ? (
                  <div className="prose prose-invert prose-sm max-w-none text-ink-200" dangerouslySetInnerHTML={{ __html: detail.body_html }} />
                ) : (
                  <pre className="whitespace-pre-wrap font-sans text-sm text-ink-200">{detail.body_text}</pre>
                )}
                {detail.attachments.length > 0 && (
                  <div className="mt-4 space-y-1.5 border-t border-ink-800/60 pt-3">
                    <div className="text-xs font-medium text-ink-400">Attachments</div>
                    {detail.attachments.map((a, i) => (
                      <div key={i} className="flex items-center gap-2 text-xs text-ink-300">
                        <Paperclip className="h-3.5 w-3.5" /> {a.filename} <span className="text-ink-500">({(a.size / 1024).toFixed(1)} KB)</span>
                      </div>
                    ))}
                  </div>
                )}
              </div>

              {replying && (
                <div className="border-t border-ink-700/60 px-6 py-4">
                  <textarea
                    value={replyBody}
                    onChange={(e) => setReplyBody(e.target.value)}
                    rows={5}
                    placeholder="Write your reply..."
                    className="w-full rounded-xl border border-ink-700/60 bg-ink-850/70 p-3 text-sm text-ink-100 placeholder:text-ink-500 focus:border-brand-500/60 focus:outline-none"
                  />
                  <div className="mt-2 flex items-center justify-between">
                    <span className="text-xs text-ink-500">
                      {sendState === "error" && <span className="text-accent-rose">Failed to send. Check SMTP settings.</span>}
                      {sendState === "sent" && <span className="text-accent-green">Reply sent.</span>}
                    </span>
                    <Button onClick={sendReply} disabled={sendState === "sending" || !replyBody.trim()}>
                      <Send className="h-4 w-4" /> {sendState === "sending" ? "Sending..." : "Send Reply"}
                    </Button>
                  </div>
                </div>
              )}
            </div>
          )}
        </Card>
      </div>
    </Layout>
  );
}

function formatDate(iso: string) {
  if (!iso) return "";
  const d = new Date(iso);
  if (isNaN(d.getTime())) return iso;
  return d.toLocaleString(undefined, { month: "short", day: "numeric", hour: "2-digit", minute: "2-digit" });
}
