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

interface Session {
  creds: Credentials | null;
  accessToken: string | null;
  refreshToken: string | null;
  expiresAt: number;
  guest: boolean;
  msal: boolean;
  email: string | null;
}

const session: Session = {
  creds: null,
  accessToken: null,
  refreshToken: null,
  expiresAt: 0,
  guest: false,
  msal: false,
  email: null,
};

const listeners = new Set<() => void>();
const emit = () => listeners.forEach((l) => l());

export function subscribeSession(l: () => void) {
  listeners.add(l);
  return () => listeners.delete(l);
}

export function getSessionSnapshot() {
  return {
    connected: !!session.accessToken || session.guest || session.msal,
    guest: session.guest,
    msal: session.msal,
    email: session.email,
  };
}

export function signOut() {
  session.creds = null;
  session.accessToken = null;
  session.refreshToken = null;
  session.expiresAt = 0;
  session.guest = false;
  session.msal = false;
  session.email = null;
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
  session.email = account.username ?? null;
  emit();
}

export function parseCredentialsLine(input: string): Credentials {
  const line = input.trim().split(/\r?\n/)[0]?.trim() ?? "";
  const parts = line.split("|").map((p) => p.trim());
  if (parts.length < 4) {
    throw new Error("Expected format: email|password|refresh_token|client_id");
  }
  const [email, password, refreshToken, clientId] = parts;
  if (!email || !password || !refreshToken || !clientId) {
    throw new Error("All 4 fields are required.");
  }
  return { email, password, refreshToken, clientId };
}

async function refreshAccessToken(creds: Credentials): Promise<string> {
  const body = new URLSearchParams({
    client_id: creds.clientId,
    grant_type: "refresh_token",
    refresh_token: session.refreshToken || creds.refreshToken,
    scope:
      "offline_access openid profile Mail.Read Mail.ReadWrite Mail.Send User.Read",
  });

  const res = await fetch(
    "https://login.microsoftonline.com/common/oauth2/v2.0/token",
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body,
    },
  );
  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Token refresh failed: ${res.status} ${err}`);
  }
  const data = (await res.json()) as {
    access_token: string;
    refresh_token?: string;
    expires_in: number;
  };
  session.accessToken = data.access_token;
  if (data.refresh_token) session.refreshToken = data.refresh_token;
  session.expiresAt = Date.now() + (data.expires_in - 60) * 1000;
  return data.access_token;
}

export async function connect(creds: Credentials) {
  signOut();
  session.creds = creds;
  session.refreshToken = creds.refreshToken;
  session.email = creds.email;
  await refreshAccessToken(creds);
  emit();
}

async function ensureToken(): Promise<string> {
  if (session.guest) throw new Error("Guest mode: no live Microsoft Graph access.");
  if (session.msal) {
    const { acquireGraphToken } = await import("./msal");
    return acquireGraphToken();
  }
  if (!session.creds) throw new Error("Not connected.");
  if (!session.accessToken || Date.now() >= session.expiresAt) {
    await refreshAccessToken(session.creds);
  }
  return session.accessToken!;
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
