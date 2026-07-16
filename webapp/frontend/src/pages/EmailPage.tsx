import { useCallback, useEffect, useState } from "react";
import { useNavigate } from "react-router-dom";
import {
  Mail,
  Paperclip,
  Reply,
  ReplyAll,
  Forward,
  RefreshCw,
  Send,
  AlertCircle,
  Trash2,
  Archive,
  Star,
  Search,
  PenSquare,
  MailOpen,
  MailX,
  X,
  Inbox,
} from "lucide-react";
import Layout from "../components/Layout";
import { Button, Spinner, EmptyState } from "../components/ui";
import { api, type EmailSummary, type EmailDetail } from "../api/client";
import { useQuoteJobs } from "../context/QuoteJobsContext";

type ComposeMode = "new" | "reply" | "replyAll" | "forward" | null;

export default function EmailPage() {
  const [status, setStatus] = useState<{ configured: boolean; smtp_configured: boolean } | null>(null);
  const [messages, setMessages] = useState<EmailSummary[] | null>(null);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [detail, setDetail] = useState<EmailDetail | null>(null);
  const [loadError, setLoadError] = useState("");
  const [search, setSearch] = useState("");
  const [composeMode, setComposeMode] = useState<ComposeMode>(null);
  const [composeTo, setComposeTo] = useState("");
  const [composeCc, setComposeCc] = useState("");
  const [composeSubject, setComposeSubject] = useState("");
  const [composeBody, setComposeBody] = useState("");
  const [sendState, setSendState] = useState<"idle" | "sending" | "sent" | "error">("idle");
  const [quoting, setQuoting] = useState(false);
  const [quoteError, setQuoteError] = useState("");
  const [busyAction, setBusyAction] = useState("");
  const [checkedIds, setCheckedIds] = useState<Set<string>>(new Set());
  const navigate = useNavigate();
  const { startQuote } = useQuoteJobs();

  const refresh = useCallback(() => {
    setLoadError("");
    api.listEmails(search.trim()).then(setMessages).catch((e) => setLoadError(e.message));
  }, [search]);

  useEffect(() => {
    api.emailStatus().then(setStatus);
  }, []);

  useEffect(() => {
    const t = setTimeout(refresh, search ? 300 : 0);
    return () => clearTimeout(t);
  }, [refresh, search]);

  useEffect(() => {
    if (!selectedId) {
      setDetail(null);
      return;
    }
    setDetail(null);
    setQuoteError("");
    api
      .getEmail(selectedId)
      .then(async (d) => {
        setDetail(d);
        if (!d.seen) {
          await api.markEmailRead(selectedId, true).catch(() => {});
          setMessages((prev) =>
            prev?.map((m) => (m.id === selectedId ? { ...m, seen: true } : m)) ?? prev
          );
        }
      })
      .catch((e) => setLoadError(e.message));
    setComposeMode(null);
    setSendState("idle");
  }, [selectedId]);

  const quoteThis = async () => {
    if (!detail) return;
    setQuoting(true);
    setQuoteError("");
    try {
      const result = await api.quoteEmail(detail.id, true);
      const qid = result.quote_id || result.job_id;
      startQuote(qid, detail.subject || "Email quote", {
        phase: "running",
        message: "Running in background — DME → SolidWorks → Module6121",
        job_id: result.job_id,
      });
    } catch (e) {
      setQuoteError(e instanceof Error ? e.message : "Could not start quote");
    } finally {
      setQuoting(false);
    }
  };

  const toggleChecked = (id: string) => {
    setCheckedIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const quoteSelected = async () => {
    const ids = Array.from(checkedIds);
    if (ids.length === 0) return;
    setQuoting(true);
    setQuoteError("");
    try {
      if (ids.length === 1) {
        const result = await api.quoteEmail(ids[0], true);
        const qid = result.quote_id || result.job_id;
        const subj = messages?.find((m) => m.id === ids[0])?.subject || "Email quote";
        startQuote(qid, subj, {
          phase: "running",
          message: "Running in background — DME → SolidWorks → Module6121",
          job_id: result.job_id,
        });
      } else {
        const result = await api.quoteEmailBatch(ids, true);
        if (result.error && !result.launched) {
          throw new Error(result.error);
        }
        const qids = result.quote_ids || [];
        qids.forEach((qid, i) => {
          const mid = ids[i];
          const subj = messages?.find((m) => m.id === mid)?.subject || qid;
          startQuote(qid, subj, {
            phase: "running",
            message: `Batch ${i + 1}/${qids.length} — one job at a time (SolidWorks restarts between)`,
            job_id: result.c_numbers?.[i] || qid,
          });
        });
        setCheckedIds(new Set());
      }
    } catch (e) {
      setQuoteError(e instanceof Error ? e.message : "Could not start batch quote");
    } finally {
      setQuoting(false);
    }
  };

  const runAction = async (action: string, fn: () => Promise<unknown>) => {
    setBusyAction(action);
    try {
      await fn();
      if (["delete", "archive"].includes(action)) {
        setMessages((prev) => prev?.filter((m) => m.id !== selectedId) ?? prev);
        setSelectedId(null);
        setDetail(null);
      } else {
        refresh();
        if (selectedId) {
          const d = await api.getEmail(selectedId);
          setDetail(d);
        }
      }
    } catch (e) {
      setLoadError(e instanceof Error ? e.message : "Action failed");
    } finally {
      setBusyAction("");
    }
  };

  const openCompose = (mode: ComposeMode) => {
    if (!detail && mode !== "new") return;
    setComposeMode(mode);
    setSendState("idle");
    if (mode === "new") {
      setComposeTo("");
      setComposeCc("");
      setComposeSubject("");
      setComposeBody("");
    } else if (detail) {
      const fromAddr = parseAddr(detail.from);
      if (mode === "reply") {
        setComposeTo(fromAddr);
        setComposeCc("");
        setComposeSubject(detail.subject.startsWith("Re:") ? detail.subject : `Re: ${detail.subject}`);
        setComposeBody("");
      } else if (mode === "replyAll") {
        const toList = [fromAddr, ...parseAddrList(detail.to)].filter(Boolean);
        setComposeTo([...new Set(toList)].join(", "));
        setComposeCc(parseAddrList(detail.cc || "").join(", "));
        setComposeSubject(detail.subject.startsWith("Re:") ? detail.subject : `Re: ${detail.subject}`);
        setComposeBody("");
      } else if (mode === "forward") {
        setComposeTo("");
        setComposeCc("");
        setComposeSubject(detail.subject.startsWith("Fwd:") ? detail.subject : `Fwd: ${detail.subject}`);
        setComposeBody("");
      }
    }
  };

  const sendCompose = async () => {
    setSendState("sending");
    try {
      const toList = parseAddrList(composeTo);
      const ccList = parseAddrList(composeCc);
      if (composeMode === "new") {
        await api.composeEmail(toList, composeSubject, composeBody, ccList);
      } else if (detail) {
        if (composeMode === "forward") {
          await api.forwardEmail(detail.id, composeTo, composeBody);
        } else if (composeMode === "replyAll") {
          await api.replyAllEmail(
            detail.id,
            toList,
            ccList,
            composeSubject,
            composeBody,
            detail.message_id_header || ""
          );
        } else {
          await api.replyEmail(
            detail.id,
            toList[0] || composeTo,
            composeSubject,
            composeBody,
            detail.message_id_header || ""
          );
        }
      }
      setSendState("sent");
      setComposeMode(null);
    } catch {
      setSendState("error");
    }
  };

  const toggleStar = async (id: string, starred: boolean) => {
    await api.starEmail(id, starred);
    setMessages((prev) => prev?.map((m) => (m.id === id ? { ...m, starred } : m)) ?? prev);
    if (detail?.id === id) setDetail({ ...detail, starred });
  };

  if (status && !status.configured) {
    return (
      <Layout title="Inbox" subtitle="Read, reply, and quote customer emails.">
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
    <Layout title="Inbox" subtitle="Gmail-style inbox — quote runs in the background.">
      <div className="glass-panel-strong flex h-[calc(100vh-148px)] flex-col gap-0 overflow-hidden rounded-2xl">
        {/* Toolbar */}
        <div className="flex shrink-0 items-center gap-2 border-b border-white/8 px-3 py-2.5">
          <Button onClick={() => openCompose("new")} className="px-4 py-2">
            <PenSquare className="h-3.5 w-3.5" /> Compose
          </Button>
          <div className="relative ml-2 flex-1 max-w-xl">
            <Search className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-ink-500" />
            <input
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="Search mail"
              className="glass-input w-full rounded-full py-2 pl-9 pr-4 text-sm text-ink-200 placeholder:text-ink-500"
            />
          </div>
          <button onClick={refresh} className="rounded-full p-2 text-ink-500 hover:bg-ink-800/40 hover:text-ink-200" title="Refresh">
            <RefreshCw className={`h-4 w-4 ${busyAction === "refresh" ? "animate-spin" : ""}`} />
          </button>
        </div>

        <div className="grid min-h-0 flex-1 grid-cols-1 lg:grid-cols-[minmax(280px,360px)_1fr]">
          {/* Message list */}
          <div className="flex min-h-0 flex-col border-r border-ink-700/25 bg-ink-900/40">
            <div className="flex items-center gap-2 border-b border-ink-700/20 px-4 py-2 text-[10px] font-bold uppercase tracking-wider text-ink-500">
              <Inbox className="h-3.5 w-3.5" /> Inbox
              {messages && <span className="text-ink-600">({messages.length})</span>}
              {checkedIds.size > 0 && (
                <button
                  type="button"
                  disabled={quoting}
                  onClick={quoteSelected}
                  className="ml-auto rounded bg-sky-700/80 px-2 py-0.5 text-[10px] font-semibold normal-case tracking-normal text-white hover:bg-sky-600 disabled:opacity-50"
                >
                  {quoting ? "Starting…" : `Quote selected (${checkedIds.size})`}
                </button>
              )}
            </div>
            <div className="scrollbar-thin flex-1 overflow-y-auto">
              {messages === null ? (
                <div className="p-6"><Spinner label="Loading inbox..." /></div>
              ) : loadError && messages.length === 0 ? (
                <div className="flex items-start gap-2 p-4 text-xs text-accent-rose">
                  <AlertCircle className="h-4 w-4 shrink-0" /> {loadError}
                </div>
              ) : messages.length === 0 ? (
                <p className="p-6 text-sm text-ink-500">No messages found.</p>
              ) : (
                messages.map((m) => (
                  <div
                    key={m.id}
                    className={`inbox-row flex cursor-pointer items-start gap-2 border-b border-ink-700/15 px-3 py-2.5 ${
                      selectedId === m.id ? "inbox-row-selected" : ""
                    } ${!m.seen ? "inbox-unread" : ""}`}
                    onClick={() => setSelectedId(m.id)}
                  >
                    <input
                      type="checkbox"
                      className="mt-1 shrink-0"
                      checked={checkedIds.has(m.id)}
                      onClick={(e) => e.stopPropagation()}
                      onChange={() => toggleChecked(m.id)}
                      title="Select for batch quote"
                    />
                    <button
                      onClick={(e) => {
                        e.stopPropagation();
                        toggleStar(m.id, !m.starred);
                      }}
                      className="mt-0.5 shrink-0 p-0.5 text-ink-600 hover:text-accent-amber"
                    >
                      <Star className={`h-3.5 w-3.5 ${m.starred ? "fill-accent-amber text-accent-amber" : ""}`} />
                    </button>
                    <div className="min-w-0 flex-1">
                      <div className="flex items-baseline justify-between gap-2">
                        <span className="inbox-sender truncate text-xs text-ink-400">
                          {m.from_name || m.from}
                        </span>
                        <span className="shrink-0 text-[10px] text-ink-600">{formatListDate(m.date)}</span>
                      </div>
                      <div className="inbox-subject truncate text-sm text-ink-300">{m.subject || "(no subject)"}</div>
                      {m.snippet && (
                        <div className="truncate text-[11px] text-ink-600">{m.snippet}</div>
                      )}
                    </div>
                  </div>
                ))
              )}
            </div>
          </div>

          {/* Reading pane */}
          <div className="flex min-h-0 flex-col bg-ink-950/30">
            {!selectedId ? (
              <div className="flex flex-1 flex-col items-center justify-center gap-3 text-ink-500">
                <Mail className="h-10 w-10 opacity-40" />
                <p className="text-sm">Select a message to read</p>
              </div>
            ) : !detail ? (
              <div className="p-8"><Spinner label="Loading message..." /></div>
            ) : (
              <>
                <div className="flex shrink-0 flex-wrap items-center gap-0.5 border-b border-ink-700/20 bg-ink-850/50 px-2 py-1.5">
                  <ToolbarBtn
                    icon={<Archive className="h-4 w-4" />}
                    title="Archive"
                    disabled={!!busyAction}
                    onClick={() => runAction("archive", () => api.archiveEmail(detail.id))}
                  />
                  <ToolbarBtn
                    icon={<Trash2 className="h-4 w-4" />}
                    title="Delete"
                    disabled={!!busyAction}
                    onClick={() => runAction("delete", () => api.deleteEmail(detail.id))}
                  />
                  <ToolbarBtn
                    icon={detail.seen ? <MailX className="h-4 w-4" /> : <MailOpen className="h-4 w-4" />}
                    title={detail.seen ? "Mark unread" : "Mark read"}
                    disabled={!!busyAction}
                    onClick={() =>
                      runAction("read", () => api.markEmailRead(detail.id, !detail.seen))
                    }
                  />
                  <ToolbarBtn
                    icon={<Star className={`h-4 w-4 ${detail.starred ? "fill-accent-amber text-accent-amber" : ""}`} />}
                    title={detail.starred ? "Unstar" : "Star"}
                    onClick={() => toggleStar(detail.id, !detail.starred)}
                  />
                  <div className="mx-1 h-5 w-px bg-ink-700/30" />
                  <ToolbarBtn icon={<Reply className="h-4 w-4" />} title="Reply" onClick={() => openCompose("reply")} />
                  <ToolbarBtn icon={<ReplyAll className="h-4 w-4" />} title="Reply all" onClick={() => openCompose("replyAll")} />
                  <ToolbarBtn icon={<Forward className="h-4 w-4" />} title="Forward" onClick={() => openCompose("forward")} />
                  <div className="flex-1" />
                  <Button onClick={quoteThis} disabled={quoting} className="mx-2 px-6 py-2">
                    {quoting ? "Starting…" : "Quote"}
                  </Button>
                </div>

                <div className="border-b border-ink-700/15 px-6 py-4">
                  <h2 className="text-base font-semibold text-ink-100">{detail.subject || "(no subject)"}</h2>
                  <div className="mt-2 flex flex-wrap items-center gap-x-4 gap-y-1 text-xs text-ink-500">
                    <span><span className="text-ink-600">From</span> {detail.from}</span>
                    <span><span className="text-ink-600">To</span> {detail.to}</span>
                    <span>{formatDetailDate(detail.date)}</span>
                  </div>
                  {quoteError && <p className="mt-2 text-xs text-accent-rose">{quoteError}</p>}
                </div>

                <div className="scrollbar-thin flex-1 overflow-y-auto px-6 py-5">
                  {detail.body_html ? (
                    <div
                      className="prose prose-sm max-w-none text-ink-300 [&_a]:text-accent-teal"
                      dangerouslySetInnerHTML={{ __html: detail.body_html }}
                    />
                  ) : (
                    <pre className="whitespace-pre-wrap font-sans text-sm leading-relaxed text-ink-300">
                      {detail.body_text}
                    </pre>
                  )}
                  {detail.attachments.length > 0 && (
                    <div className="mt-6 rounded-lg border border-ink-700/25 bg-ink-900/50 p-4">
                      <div className="mb-2 text-[10px] font-bold uppercase tracking-wider text-ink-500">
                        {detail.attachments.length} attachment{detail.attachments.length > 1 ? "s" : ""}
                      </div>
                      <div className="flex flex-wrap gap-2">
                        {detail.attachments.map((a, i) => (
                          <div
                            key={i}
                            className="flex items-center gap-2 rounded border border-ink-700/30 bg-ink-850/60 px-3 py-2 text-xs text-ink-300"
                          >
                            <Paperclip className="h-3.5 w-3.5 text-ink-500" />
                            <span>{a.filename}</span>
                            <span className="text-ink-600">({formatSize(a.size)})</span>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}
                </div>
              </>
            )}
          </div>
        </div>
      </div>

      {composeMode && (
        <ComposePanel
          mode={composeMode}
          to={composeTo}
          cc={composeCc}
          subject={composeSubject}
          body={composeBody}
          sendState={sendState}
          onToChange={setComposeTo}
          onCcChange={setComposeCc}
          onSubjectChange={setComposeSubject}
          onBodyChange={setComposeBody}
          onSend={sendCompose}
          onClose={() => setComposeMode(null)}
        />
      )}
    </Layout>
  );
}

function ToolbarBtn({
  icon,
  title,
  onClick,
  disabled,
}: {
  icon: React.ReactNode;
  title: string;
  onClick: () => void;
  disabled?: boolean;
}) {
  return (
    <button
      onClick={onClick}
      disabled={disabled}
      title={title}
      className="rounded p-2 text-ink-500 transition hover:bg-ink-800/40 hover:text-ink-200 disabled:opacity-40"
    >
      {icon}
    </button>
  );
}

function ComposePanel({
  mode,
  to,
  cc,
  subject,
  body,
  sendState,
  onToChange,
  onCcChange,
  onSubjectChange,
  onBodyChange,
  onSend,
  onClose,
}: {
  mode: ComposeMode;
  to: string;
  cc: string;
  subject: string;
  body: string;
  sendState: string;
  onToChange: (v: string) => void;
  onCcChange: (v: string) => void;
  onSubjectChange: (v: string) => void;
  onBodyChange: (v: string) => void;
  onSend: () => void;
  onClose: () => void;
}) {
  const title =
    mode === "new" ? "New message" : mode === "reply" ? "Reply" : mode === "replyAll" ? "Reply all" : "Forward";

  return (
    <div className="compose-shadow glass-panel-strong fixed bottom-20 right-4 z-40 flex w-full max-w-xl flex-col overflow-hidden rounded-2xl md:bottom-4">
      <div className="flex items-center justify-between border-b border-white/8 px-4 py-3">
        <span className="text-xs font-semibold uppercase tracking-wider text-ink-200">{title}</span>
        <button onClick={onClose} className="rounded-full p-1 text-ink-500 hover:bg-white/10 hover:text-ink-200">
          <X className="h-4 w-4" />
        </button>
      </div>
      <div className="space-y-0 border-b border-ink-700/20">
        <ComposeField label="To" value={to} onChange={onToChange} />
        {(mode === "new" || mode === "replyAll") && (
          <ComposeField label="Cc" value={cc} onChange={onCcChange} />
        )}
        <ComposeField label="Subject" value={subject} onChange={onSubjectChange} />
      </div>
      <textarea
        value={body}
        onChange={(e) => onBodyChange(e.target.value)}
        rows={10}
        placeholder="Write your message..."
        className="resize-none bg-ink-950/40 px-4 py-3 text-sm text-ink-200 placeholder:text-ink-600 focus:outline-none"
      />
      <div className="flex items-center justify-between border-t border-ink-700/20 px-4 py-3">
        {sendState === "error" && <span className="text-xs text-accent-rose">Send failed</span>}
        {sendState === "sent" && <span className="text-xs text-accent-green">Sent</span>}
        <div className="ml-auto flex gap-2">
          <Button variant="ghost" onClick={onClose}>Discard</Button>
          <Button onClick={onSend} disabled={sendState === "sending" || !to.trim() || !body.trim()}>
            <Send className="h-4 w-4" /> {sendState === "sending" ? "Sending…" : "Send"}
          </Button>
        </div>
      </div>
    </div>
  );
}

function ComposeField({
  label,
  value,
  onChange,
}: {
  label: string;
  value: string;
  onChange: (v: string) => void;
}) {
  return (
    <div className="flex items-center border-b border-ink-700/15">
      <span className="w-16 shrink-0 px-4 text-[10px] font-bold uppercase tracking-wider text-ink-500">{label}</span>
      <input
        value={value}
        onChange={(e) => onChange(e.target.value)}
        className="flex-1 bg-transparent py-2.5 pr-4 text-sm text-ink-200 focus:outline-none"
      />
    </div>
  );
}

function parseAddr(from: string): string {
  const m = from.match(/<(.+)>/);
  return (m?.[1] || from).trim();
}

function parseAddrList(raw: string): string[] {
  return raw
    .split(/[,;]/)
    .map((s) => parseAddr(s.trim()))
    .filter(Boolean);
}

function formatListDate(iso: string) {
  if (!iso) return "";
  const d = new Date(iso);
  if (isNaN(d.getTime())) return iso.slice(0, 10);
  const now = new Date();
  if (d.toDateString() === now.toDateString()) {
    return d.toLocaleTimeString(undefined, { hour: "numeric", minute: "2-digit" });
  }
  return d.toLocaleDateString(undefined, { month: "short", day: "numeric" });
}

function formatDetailDate(iso: string) {
  if (!iso) return "";
  const d = new Date(iso);
  if (isNaN(d.getTime())) return iso;
  return d.toLocaleString(undefined, {
    weekday: "short",
    month: "short",
    day: "numeric",
    year: "numeric",
    hour: "numeric",
    minute: "2-digit",
  });
}

function formatSize(bytes: number) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}
