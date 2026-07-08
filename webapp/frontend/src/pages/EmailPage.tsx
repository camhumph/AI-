import { useEffect, useState } from "react";
import { useNavigate } from "react-router-dom";
import { Mail, Paperclip, Reply, RefreshCw, Send, AlertCircle, Trash2 } from "lucide-react";
import Layout from "../components/Layout";
import { Card, Button, Spinner, EmptyState } from "../components/ui";
import { api, type EmailSummary, type EmailDetail, type QuoteRunStatus } from "../api/client";
import QuoteProgressModal from "../components/QuoteProgressModal";

export default function EmailPage() {
  const [status, setStatus] = useState<{ configured: boolean; smtp_configured: boolean } | null>(null);
  const [messages, setMessages] = useState<EmailSummary[] | null>(null);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [detail, setDetail] = useState<EmailDetail | null>(null);
  const [loadError, setLoadError] = useState("");
  const [replying, setReplying] = useState(false);
  const [replyBody, setReplyBody] = useState("");
  const [sendState, setSendState] = useState<"idle" | "sending" | "sent" | "error">("idle");
  const [quoting, setQuoting] = useState(false);
  const [quoteError, setQuoteError] = useState("");
  const [quoteProgress, setQuoteProgress] = useState<QuoteRunStatus | null>(null);
  const [activeQuoteId, setActiveQuoteId] = useState<string | null>(null);
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
    setQuoteError("");
    api.getEmail(selectedId).then(setDetail).catch((e) => setLoadError(e.message));
    setReplying(false);
    setSendState("idle");
  }, [selectedId]);

  useEffect(() => {
    if (!activeQuoteId) return;
    const iv = setInterval(async () => {
      try {
        const st = await api.quoteStatus(activeQuoteId);
        setQuoteProgress(st);
        if (st.phase === "completed" && st.job_id) {
          clearInterval(iv);
          setActiveQuoteId(null);
          navigate(`/quotes/${encodeURIComponent(st.job_id)}`);
        }
        if (st.phase === "error") {
          clearInterval(iv);
          setQuoteError(st.message || "Quote failed");
          setActiveQuoteId(null);
          setQuoting(false);
        }
      } catch {
        /* keep polling */
      }
    }, 3000);
    return () => clearInterval(iv);
  }, [activeQuoteId, navigate]);

  const quoteThis = async () => {
    if (!detail) return;
    setQuoting(true);
    setQuoteError("");
    setQuoteProgress({ phase: "queued", message: "Starting full quote pipeline..." });
    try {
      const result = await api.quoteEmail(detail.id, true);
      const qid = result.quote_id || result.job_id;
      setActiveQuoteId(qid);
      setQuoteProgress({ phase: "running", message: "DME price lookup → SolidWorks → Module6121 → AI...", job_id: result.job_id });
    } catch (e) {
      setQuoteError(e instanceof Error ? e.message : "Could not start quote");
      setQuoting(false);
      setQuoteProgress(null);
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
          title="Connect your email"
          description="Open Settings, enter your Gmail app password, and save."
          action={<Button variant="secondary" onClick={() => navigate("/settings")}>Settings</Button>}
        />
      </Layout>
    );
  }

  return (
    <Layout title="Inbox" subtitle="Select an email and press Quote — the full pipeline runs automatically.">
      {quoteProgress && activeQuoteId && (
        <QuoteProgressModal status={quoteProgress} onClose={() => { setActiveQuoteId(null); setQuoting(false); setQuoteProgress(null); }} />
      )}

      <div className="grid h-[calc(100vh-160px)] grid-cols-1 gap-4 lg:grid-cols-[340px_1fr]">
        <Card className="flex flex-col overflow-hidden">
          <div className="flex items-center justify-between border-b border-ink-700/20 px-4 py-3">
            <span className="section-label">Messages</span>
            <button onClick={refresh} className="text-ink-500 hover:text-ink-100">
              <RefreshCw className="h-4 w-4" />
            </button>
          </div>
          <div className="scrollbar-thin flex-1 overflow-y-auto">
            {messages === null ? (
              <div className="p-4"><Spinner label="Loading inbox..." /></div>
            ) : loadError ? (
              <div className="flex items-start gap-2 p-4 text-xs text-accent-rose"><AlertCircle className="h-4 w-4 shrink-0" /> {loadError}</div>
            ) : messages.length === 0 ? (
              <p className="p-4 text-sm text-ink-500">No messages found.</p>
            ) : (
              messages.map((m) => (
                <button
                  key={m.id}
                  onClick={() => setSelectedId(m.id)}
                  className={`block w-full border-b border-ink-700/15 px-4 py-3 text-left transition ${
                    selectedId === m.id ? "bg-white/90" : "hover:bg-white/50"
                  }`}
                >
                  <div className="flex items-center justify-between">
                    <span className="truncate text-xs font-medium text-ink-300">{m.from}</span>
                    <span className="shrink-0 text-[10px] text-ink-500">{formatDate(m.date)}</span>
                  </div>
                  <div className="mt-0.5 truncate text-sm text-ink-100">{m.subject || "(no subject)"}</div>
                </button>
              ))
            )}
          </div>
        </Card>

        <Card className="flex flex-col overflow-hidden">
          {!selectedId ? (
            <div className="flex flex-1 items-center justify-center text-sm text-ink-500">Select a message</div>
          ) : !detail ? (
            <div className="p-6"><Spinner label="Loading message..." /></div>
          ) : (
            <div className="flex flex-1 flex-col overflow-hidden">
              <div className="flex items-center gap-1 border-b border-ink-700/20 bg-white/40 px-3 py-2">
                <button onClick={() => setReplying((r) => !r)} className="rounded p-2 text-ink-500 hover:bg-white/80 hover:text-ink-100" title="Reply">
                  <Reply className="h-4 w-4" />
                </button>
                <button className="rounded p-2 text-ink-400" disabled title="Delete in Gmail"><Trash2 className="h-4 w-4" /></button>
                <div className="flex-1" />
                <button
                  onClick={quoteThis}
                  disabled={quoting}
                  className="border border-ink-100 bg-ink-100 px-8 py-2.5 text-[10px] font-bold uppercase tracking-[0.2em] text-white transition hover:bg-transparent hover:text-ink-100 disabled:opacity-50"
                >
                  {quoting ? "Quoting..." : "Quote"}
                </button>
              </div>

              <div className="border-b border-ink-700/20 px-6 py-4">
                <h2 className="truncate text-sm font-semibold uppercase tracking-wider text-ink-100">{detail.subject || "(no subject)"}</h2>
                <p className="mt-1 text-xs text-ink-500">From {detail.from} · {formatDate(detail.date)}</p>
                {quoteError && <p className="mt-2 text-xs text-accent-rose">{quoteError}</p>}
              </div>

              <div className="scrollbar-thin flex-1 overflow-y-auto px-6 py-4">
                {detail.body_html ? (
                  <div className="prose prose-sm max-w-none text-ink-300" dangerouslySetInnerHTML={{ __html: detail.body_html }} />
                ) : (
                  <pre className="whitespace-pre-wrap font-sans text-sm text-ink-300">{detail.body_text}</pre>
                )}
                {detail.attachments.length > 0 && (
                  <div className="mt-4 border-t border-ink-700/20 pt-3">
                    <div className="section-label mb-2">Attachments (auto-used by Quote)</div>
                    {detail.attachments.map((a, i) => (
                      <div key={i} className="flex items-center gap-2 text-xs text-ink-400">
                        <Paperclip className="h-3.5 w-3.5" /> {a.filename}
                      </div>
                    ))}
                  </div>
                )}
              </div>

              {replying && (
                <div className="border-t border-ink-700/20 px-6 py-4">
                  <textarea
                    value={replyBody}
                    onChange={(e) => setReplyBody(e.target.value)}
                    rows={4}
                    className="w-full border border-ink-700/30 bg-white/80 p-3 text-sm text-ink-100 focus:border-ink-100 focus:outline-none"
                  />
                  <div className="mt-2 flex justify-end">
                    <Button onClick={sendReply} disabled={sendState === "sending" || !replyBody.trim()}>
                      <Send className="h-4 w-4" /> Send
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
