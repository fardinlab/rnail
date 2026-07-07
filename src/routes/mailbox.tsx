import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { useCallback, useEffect, useMemo, useState, useSyncExternalStore } from "react";
import {
  Inbox,
  Star,
  Send,
  FileText,
  ShieldAlert,
  Archive,
  Trash2,
  Search,
  RefreshCw,
  ArrowDownWideNarrow,
  ArrowUpWideNarrow,
  Copy,
  LogOut,
  Loader2,
  Mail as MailIcon,
  Menu,
} from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Badge } from "@/components/ui/badge";
import { ScrollArea } from "@/components/ui/scroll-area";
import { Sheet, SheetContent, SheetTrigger } from "@/components/ui/sheet";
import { toast } from "sonner";
import {
  extractOtp,
  getMessage,
  getSessionSnapshot,
  listMessages,
  signOut,
  subscribeSession,
  type GraphMessage,
  type MailFolder,
} from "@/lib/graph";

export const Route = createFileRoute("/mailbox")({
  component: MailboxPage,
});

const FOLDERS: { key: MailFolder; label: string; icon: React.ComponentType<{ className?: string }> }[] = [
  { key: "inbox", label: "Inbox", icon: Inbox },
  { key: "starred", label: "Starred", icon: Star },
  { key: "sent", label: "Sent", icon: Send },
  { key: "drafts", label: "Drafts", icon: FileText },
  { key: "junk", label: "Junk", icon: ShieldAlert },
  { key: "archive", label: "Archive", icon: Archive },
  { key: "trash", label: "Trash", icon: Trash2 },
];

function useSession() {
  return useSyncExternalStore(
    (cb) => subscribeSession(cb),
    () => JSON.stringify(getSessionSnapshot()),
    () => JSON.stringify(getSessionSnapshot()),
  );
}

function formatDate(iso: string) {
  const d = new Date(iso);
  const now = new Date();
  const sameDay = d.toDateString() === now.toDateString();
  return sameDay
    ? d.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })
    : d.toLocaleDateString([], { month: "short", day: "numeric" });
}

function MailboxPage() {
  const sessionKey = useSession();
  const session = useMemo(() => JSON.parse(sessionKey) as ReturnType<typeof getSessionSnapshot>, [sessionKey]);
  const navigate = useNavigate();

  const [folder, setFolder] = useState<MailFolder>("inbox");
  const [messages, setMessages] = useState<GraphMessage[]>([]);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [selected, setSelected] = useState<GraphMessage | null>(null);
  const [search, setSearch] = useState("");
  const [searchInput, setSearchInput] = useState("");
  const [loading, setLoading] = useState(false);
  const [newestFirst, setNewestFirst] = useState(true);

  useEffect(() => {
    if (!session.connected) navigate({ to: "/" });
  }, [session.connected, navigate]);

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const data = await listMessages(folder, { search: search || undefined });
      setMessages(data);
      setSelectedId((prev) => (data.find((m) => m.id === prev) ? prev : data[0]?.id ?? null));
    } catch (e) {
      toast.error("Failed to load mail", {
        description: e instanceof Error ? e.message : String(e),
      });
    } finally {
      setLoading(false);
    }
  }, [folder, search]);

  useEffect(() => {
    void load();
  }, [load]);

  // Auto-refresh every 10s
  useEffect(() => {
    const id = window.setInterval(() => {
      void load();
    }, 10_000);
    return () => window.clearInterval(id);
  }, [load]);


  useEffect(() => {
    if (!selectedId) {
      setSelected(null);
      return;
    }
    let cancelled = false;
    void (async () => {
      try {
        const msg = await getMessage(selectedId);
        if (!cancelled) setSelected(msg);
      } catch (e) {
        if (!cancelled) toast.error("Failed to open message", { description: String(e) });
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [selectedId]);

  const sorted = useMemo(() => {
    const copy = [...messages];
    copy.sort((a, b) => {
      const t = new Date(b.receivedDateTime).getTime() - new Date(a.receivedDateTime).getTime();
      return newestFirst ? t : -t;
    });
    return copy;
  }, [messages, newestFirst]);

  const otp = useMemo(() => {
    if (!selected) return null;
    const text =
      (selected.body?.content ?? "") + " " + (selected.subject ?? "") + " " + selected.bodyPreview;
    return extractOtp(text);
  }, [selected]);

  const handleCopyOtp = async () => {
    if (!otp) return;
    await navigator.clipboard.writeText(otp);
    toast.success("OTP copied", { description: otp });
  };

  const sidebar = (
    <nav className="flex flex-col gap-1 p-3">
      <div className="px-2 py-2 mb-2">
        <div className="flex items-center gap-2">
          <div className="h-8 w-8 rounded-md bg-primary text-primary-foreground grid place-items-center">
            <MailIcon className="h-4 w-4" />
          </div>
          <div className="min-w-0">
            <div className="text-sm font-semibold truncate">
              {session.guest ? "Guest mode" : session.email ?? "Mailbox"}
            </div>
            <div className="text-xs text-muted-foreground">
              {session.guest ? "Demo data" : "Connected"}
            </div>
          </div>
        </div>
      </div>
      {FOLDERS.map(({ key, label, icon: Icon }) => (
        <button
          key={key}
          onClick={() => {
            setFolder(key);
            setSelectedId(null);
          }}
          className={`flex items-center gap-3 px-3 py-2 rounded-md text-sm transition-colors text-left ${
            folder === key
              ? "bg-primary text-primary-foreground"
              : "hover:bg-muted text-foreground/80"
          }`}
        >
          <Icon className="h-4 w-4" />
          <span>{label}</span>
        </button>
      ))}
      <div className="mt-auto pt-4">
        <Button
          variant="ghost"
          size="sm"
          className="w-full justify-start"
          onClick={async () => {
            const { msalSignOut } = await import("@/lib/msal");
            await msalSignOut().catch(() => null);
            signOut();
            navigate({ to: "/" });
          }}
        >
          <LogOut className="h-4 w-4 mr-2" />
          Sign out
        </Button>
      </div>
    </nav>
  );

  return (
    <div className="h-screen w-full flex bg-background text-foreground overflow-hidden">
      {/* Desktop sidebar */}
      <aside className="hidden md:flex w-60 shrink-0 border-r border-border flex-col bg-card">
        {sidebar}
      </aside>

      <div className="flex-1 flex flex-col min-w-0">
        {/* Toolbar */}
        <header className="h-14 border-b border-border flex items-center gap-2 px-3 md:px-4 bg-card">
          <Sheet>
            <SheetTrigger asChild>
              <Button variant="ghost" size="icon" className="md:hidden">
                <Menu className="h-5 w-5" />
              </Button>
            </SheetTrigger>
            <SheetContent side="left" className="w-64 p-0">
              {sidebar}
            </SheetContent>
          </Sheet>

          <form
            className="flex-1 max-w-xl relative"
            onSubmit={(e) => {
              e.preventDefault();
              setSearch(searchInput);
            }}
          >
            <Search className="h-4 w-4 absolute left-3 top-1/2 -translate-y-1/2 text-muted-foreground" />
            <Input
              placeholder="Search mail"
              className="pl-9 h-9"
              value={searchInput}
              onChange={(e) => setSearchInput(e.target.value)}
            />
          </form>

          <Button
            variant="outline"
            size="sm"
            onClick={() => setNewestFirst((v) => !v)}
            title={newestFirst ? "Newest first" : "Oldest first"}
          >
            {newestFirst ? (
              <ArrowDownWideNarrow className="h-4 w-4 mr-1.5" />
            ) : (
              <ArrowUpWideNarrow className="h-4 w-4 mr-1.5" />
            )}
            <span className="hidden sm:inline">
              {newestFirst ? "Newest" : "Oldest"}
            </span>
          </Button>
          <Button variant="outline" size="sm" onClick={() => void load()} disabled={loading}>
            {loading ? (
              <Loader2 className="h-4 w-4 animate-spin" />
            ) : (
              <RefreshCw className="h-4 w-4" />
            )}
            <span className="hidden sm:inline ml-1.5">Refresh</span>
          </Button>
        </header>

        <div className="flex-1 min-h-0 grid grid-cols-1 lg:grid-cols-[minmax(280px,380px)_1fr]">
          {/* Message list */}
          <section className="border-r border-border min-h-0 flex flex-col bg-card/40">
            <div className="px-4 py-3 border-b border-border flex items-center justify-between">
              <h2 className="text-sm font-semibold capitalize">{folder}</h2>
              <Badge variant="secondary" className="text-xs">{sorted.length}</Badge>
            </div>
            <ScrollArea className="flex-1">
              {sorted.length === 0 && !loading && (
                <div className="p-8 text-center text-sm text-muted-foreground">
                  No messages
                </div>
              )}
              <ul>
                {sorted.map((m) => {
                  const active = m.id === selectedId;
                  return (
                    <li key={m.id}>
                      <button
                        onClick={() => setSelectedId(m.id)}
                        className={`w-full text-left px-4 py-3 border-b border-border/60 transition-colors ${
                          active ? "bg-accent" : "hover:bg-muted/60"
                        }`}
                      >
                        <div className="flex items-center gap-2 mb-1">
                          {!m.isRead && (
                            <span className="h-2 w-2 rounded-full bg-primary shrink-0" />
                          )}
                          <span className={`text-sm truncate flex-1 ${!m.isRead ? "font-semibold" : ""}`}>
                            {m.from?.emailAddress.name ?? m.from?.emailAddress.address ?? "Unknown"}
                          </span>
                          <span className="text-xs text-muted-foreground shrink-0">
                            {formatDate(m.receivedDateTime)}
                          </span>
                        </div>
                        <div className="text-sm truncate mb-0.5">{m.subject || "(no subject)"}</div>
                        <div className="text-xs text-muted-foreground line-clamp-2">
                          {m.bodyPreview}
                        </div>
                      </button>
                    </li>
                  );
                })}
              </ul>
            </ScrollArea>
          </section>

          {/* Reading pane */}
          <section className="min-h-0 flex flex-col bg-background">
            {selected ? (
              <>
                <div className="px-6 py-4 border-b border-border">
                  <h1 className="text-xl font-semibold mb-2">
                    {selected.subject || "(no subject)"}
                  </h1>
                  <div className="flex items-center justify-between gap-3">
                    <div className="min-w-0">
                      <div className="text-sm font-medium truncate">
                        {selected.from?.emailAddress.name ?? selected.from?.emailAddress.address}
                      </div>
                      <div className="text-xs text-muted-foreground truncate">
                        {selected.from?.emailAddress.address}
                      </div>
                    </div>
                    <div className="text-xs text-muted-foreground shrink-0">
                      {new Date(selected.receivedDateTime).toLocaleString()}
                    </div>
                  </div>

                  {otp && (
                    <div className="mt-4 flex items-center justify-between gap-3 rounded-lg border border-primary/30 bg-primary/5 px-4 py-3">
                      <div className="min-w-0">
                        <div className="text-xs uppercase tracking-wider text-primary font-semibold">
                          Verification code
                        </div>
                        <div className="text-2xl font-mono font-bold tracking-widest mt-0.5">
                          {otp}
                        </div>
                      </div>
                      <Button size="sm" onClick={handleCopyOtp}>
                        <Copy className="h-4 w-4 mr-1.5" />
                        Copy
                      </Button>
                    </div>
                  )}
                </div>
                <ScrollArea className="flex-1">
                  <div className="px-6 py-5">
                    {selected.body?.contentType === "text/html" ? (
                      <div
                        className="prose prose-sm max-w-none dark:prose-invert"
                        dangerouslySetInnerHTML={{ __html: selected.body.content }}
                      />
                    ) : (
                      <pre className="whitespace-pre-wrap font-sans text-sm leading-relaxed">
                        {selected.body?.content ?? selected.bodyPreview}
                      </pre>
                    )}
                  </div>
                </ScrollArea>
              </>
            ) : (
              <div className="flex-1 grid place-items-center text-sm text-muted-foreground">
                Select a message to read
              </div>
            )}
          </section>
        </div>
      </div>
    </div>
  );
}
