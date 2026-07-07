import { exchangeRefreshToken } from "./token.functions";

// Microsoft Graph API client + in-memory credential store.
// Credentials are NEVER persisted to localStorage or a database — only kept
// in module-scope memory for the lifetime of the tab.

export type MailFolder =
  | "inbox"
  | "starred"
  | "sent"
  | "drafts"
  | "junk"
  | "archive"
  | "trash";

export const FOLDER_TO_GRAPH: Record<MailFolder, string> = {
  inbox: "inbox",
  starred: "inbox", // filtered client-side via flag.flagStatus == 'flagged'
  sent: "sentitems",
  drafts: "drafts",
  junk: "junkemail",
  archive: "archive",
  trash: "deleteditems",
};

export interface Credentials {
  email: string;
  password: string;
  refreshToken: string;
  clientId: string;
  tenantId?: string;
}

export interface GraphMessage {
  id: string;
  subject: string | null;
  bodyPreview: string;
  body?: { contentType: string; content: string };
  from?: { emailAddress: { name?: string; address: string } };
  toRecipients?: { emailAddress: { name?: string; address: string } }[];
  receivedDateTime: string;
  isRead: boolean;
  flag?: { flagStatus: "notFlagged" | "flagged" | "complete" };
}

interface Account {
  creds: Credentials;
  accessToken: string | null;
  refreshToken: string;
  expiresAt: number;
}

interface Session {
  accounts: Account[];
  activeIndex: number;
  guest: boolean;
  msal: boolean;
  msalEmail: string | null;
}

const session: Session = {
  accounts: [],
  activeIndex: -1,
  guest: false,
  msal: false,
  msalEmail: null,
};

const listeners = new Set<() => void>();
const emit = () => listeners.forEach((l) => l());

export function subscribeSession(l: () => void) {
  listeners.add(l);
  return () => listeners.delete(l);
}

function activeAccount(): Account | null {
  return session.accounts[session.activeIndex] ?? null;
}

export function getSessionSnapshot() {
  const active = activeAccount();
  return {
    connected: !!active || session.guest || session.msal,
    guest: session.guest,
    msal: session.msal,
    email: session.msal ? session.msalEmail : active?.creds.email ?? null,
    accounts: session.accounts.map((a) => a.creds.email),
    activeIndex: session.activeIndex,
  };
}

export function signOut() {
  session.accounts = [];
  session.activeIndex = -1;
  session.guest = false;
  session.msal = false;
  session.msalEmail = null;
  emit();
}

export function switchAccount(index: number) {
  if (index >= 0 && index < session.accounts.length) {
    session.activeIndex = index;
    emit();
  }
}

export function removeAccount(index: number) {
  session.accounts.splice(index, 1);
  if (session.accounts.length === 0) {
    session.activeIndex = -1;
  } else if (session.activeIndex >= session.accounts.length) {
    session.activeIndex = session.accounts.length - 1;
  }
  emit();
}

export function enterGuestMode() {
  signOut();
  session.guest = true;
  emit();
}

export async function signInMicrosoft() {
  const { loginPopup } = await import("./msal");
  const account = await loginPopup();
  signOut();
  session.msal = true;
  session.msalEmail = account.username ?? null;
  emit();
}

function parseOneLine(line: string): Credentials {
  const parts = line.split("|").map((p) => p.trim());
  if (parts.length < 4) {
    throw new Error("Expected format: email|password|refresh_token|client_id");
  }
  const [email, password, refreshToken, clientId, tenantId] = parts;
  if (!email || !password || !refreshToken || !clientId) {
    throw new Error("All 4 fields are required.");
  }
  return { email, password, refreshToken, clientId, tenantId: tenantId || undefined };
}

export function parseCredentialsLine(input: string): Credentials {
  const line = input.trim().split(/\r?\n/)[0]?.trim() ?? "";
  return parseOneLine(line);
}

export function parseCredentialsLines(input: string): Credentials[] {
  const lines = input
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter((l) => l.length > 0 && l.includes("|"));
  if (lines.length === 0) throw new Error("Paste at least one credentials line.");
  return lines.map(parseOneLine);
}

async function refreshAccessToken(account: Account): Promise<string> {
  const data = await exchangeRefreshToken({
    data: {
      clientId: account.creds.clientId,
      refreshToken: account.refreshToken,
      tenantId: account.creds.tenantId || import.meta.env.VITE_TENANT_ID || undefined,
    },
  });
  account.accessToken = data.access_token;
  if (data.refresh_token) account.refreshToken = data.refresh_token;
  account.expiresAt = Date.now() + (data.expires_in - 60) * 1000;
  return data.access_token;
}

export async function connect(creds: Credentials) {
  signOut();
  const account: Account = {
    creds,
    accessToken: null,
    refreshToken: creds.refreshToken,
    expiresAt: 0,
  };
  await refreshAccessToken(account);
  session.accounts = [account];
  session.activeIndex = 0;
  emit();
}

export async function connectMany(list: Credentials[]): Promise<{
  successes: string[];
  failures: { email: string; error: string }[];
}> {
  signOut();
  const successes: string[] = [];
  const failures: { email: string; error: string }[] = [];
  for (const creds of list) {
    const account: Account = {
      creds,
      accessToken: null,
      refreshToken: creds.refreshToken,
      expiresAt: 0,
    };
    try {
      await refreshAccessToken(account);
      session.accounts.push(account);
      successes.push(creds.email);
    } catch (e) {
      failures.push({ email: creds.email, error: e instanceof Error ? e.message : String(e) });
    }
  }
  if (session.accounts.length > 0) session.activeIndex = 0;
  emit();
  return { successes, failures };
}

async function ensureToken(): Promise<string> {
  if (session.guest) throw new Error("Guest mode: no live Microsoft Graph access.");
  if (session.msal) {
    const { acquireGraphToken } = await import("./msal");
    return acquireGraphToken();
  }
  const account = activeAccount();
  if (!account) throw new Error("Not connected.");
  if (!account.accessToken || Date.now() >= account.expiresAt) {
    await refreshAccessToken(account);
  }
  return account.accessToken!;
}


async function graphFetch<T>(path: string, init: RequestInit = {}): Promise<T> {
  const token = await ensureToken();
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    ...init,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...(init.headers ?? {}),
    },
  });
  if (!res.ok) {
    throw new Error(`Graph ${res.status}: ${await res.text()}`);
  }
  return (await res.json()) as T;
}

// --- Demo data for guest mode -------------------------------------------------
const DEMO_MESSAGES: GraphMessage[] = [
  {
    id: "demo-1",
    subject: "Your Microsoft verification code",
    bodyPreview:
      "Use code 483920 to complete sign-in. This code expires in 10 minutes.",
    body: {
      contentType: "text/html",
      content:
        "<p>Hi there,</p><p>Your verification code is <b>483920</b>. It expires in 10 minutes.</p><p>— Microsoft account team</p>",
    },
    from: { emailAddress: { name: "Microsoft", address: "account-security-noreply@microsoft.com" } },
    receivedDateTime: new Date(Date.now() - 3 * 60_000).toISOString(),
    isRead: false,
    flag: { flagStatus: "notFlagged" },
  },
  {
    id: "demo-2",
    subject: "GitHub sign-in code",
    bodyPreview: "Your one-time code is 902-114. Do not share it with anyone.",
    body: {
      contentType: "text/plain",
      content: "Your GitHub OTP is 902114. Do not share this with anyone.",
    },
    from: { emailAddress: { name: "GitHub", address: "noreply@github.com" } },
    receivedDateTime: new Date(Date.now() - 42 * 60_000).toISOString(),
    isRead: false,
    flag: { flagStatus: "flagged" },
  },
  {
    id: "demo-3",
    subject: "Weekly product digest",
    bodyPreview: "Here's what shipped this week across the platform...",
    body: { contentType: "text/plain", content: "Highlights from the week..." },
    from: { emailAddress: { name: "Product Team", address: "digest@example.com" } },
    receivedDateTime: new Date(Date.now() - 26 * 3600_000).toISOString(),
    isRead: true,
    flag: { flagStatus: "notFlagged" },
  },
];

// --- Public API hooks --------------------------------------------------------
export async function listMessages(
  folder: MailFolder,
  opts: { search?: string; top?: number } = {},
): Promise<GraphMessage[]> {
  if (session.guest) {
    let list = [...DEMO_MESSAGES];
    if (folder === "starred") list = list.filter((m) => m.flag?.flagStatus === "flagged");
    if (opts.search) {
      const q = opts.search.toLowerCase();
      list = list.filter(
        (m) =>
          (m.subject ?? "").toLowerCase().includes(q) ||
          m.bodyPreview.toLowerCase().includes(q),
      );
    }
    return list;
  }

  const params = new URLSearchParams({
    $top: String(opts.top ?? 25),
    $orderby: "receivedDateTime desc",
    $select: "id,subject,bodyPreview,from,receivedDateTime,isRead,flag",
  });
  if (folder === "starred") {
    params.set("$filter", "flag/flagStatus eq 'flagged'");
  }
  if (opts.search) params.set("$search", `"${opts.search.replace(/"/g, "'")}"`);

  const graphFolder = FOLDER_TO_GRAPH[folder];
  const data = await graphFetch<{ value: GraphMessage[] }>(
    `/me/mailFolders/${graphFolder}/messages?${params.toString()}`,
  );
  return data.value;
}

export async function getMessage(id: string): Promise<GraphMessage> {
  if (session.guest) {
    const m = DEMO_MESSAGES.find((x) => x.id === id);
    if (!m) throw new Error("Not found");
    return m;
  }
  return graphFetch<GraphMessage>(`/me/messages/${id}`);
}

// OTP detection: find 4-8 digit codes, allowing an optional dash/space in the middle.
export function extractOtp(text: string): string | null {
  if (!text) return null;
  const cleaned = text.replace(/<[^>]+>/g, " ");
  // Match 4-8 digit sequences, optionally split by one dash/space
  const patterns = [
    /\b(\d{3})[-\s](\d{3})\b/, // 483-920
    /\b(\d{4,8})\b/,
  ];
  const contextRe = /(code|otp|verification|passcode|one[-\s]?time|pin)/i;
  if (!contextRe.test(cleaned)) {
    // still try, but only if a lone 6-digit number exists
    const m = cleaned.match(/\b(\d{6})\b/);
    return m ? m[1] : null;
  }
  for (const re of patterns) {
    const m = cleaned.match(re);
    if (m) return m.slice(1).filter(Boolean).join("");
  }
  return null;
}
