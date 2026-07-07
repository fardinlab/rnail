// MSAL Browser setup — PKCE, no client secret.
// Tokens are cached in sessionStorage by MSAL (cleared on tab close).

import {
  PublicClientApplication,
  type AccountInfo,
  type Configuration,
  InteractionRequiredAuthError,
} from "@azure/msal-browser";

const clientId = import.meta.env.VITE_CLIENT_ID as string | undefined;
const tenantId = (import.meta.env.VITE_TENANT_ID as string | undefined) ?? "common";

export const MSAL_SCOPES = ["Mail.Read", "offline_access", "openid", "profile", "email"];

let instance: PublicClientApplication | null = null;
let initPromise: Promise<PublicClientApplication> | null = null;

export function isMsalConfigured() {
  return !!clientId;
}

export async function getMsal(): Promise<PublicClientApplication> {
  if (!clientId) {
    throw new Error(
      "VITE_CLIENT_ID is not set. Add it to your environment to enable Microsoft sign-in.",
    );
  }
  if (instance) return instance;
  if (initPromise) return initPromise;

  const config: Configuration = {
    auth: {
      clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
      redirectUri: window.location.origin,
    },
    cache: {
      cacheLocation: "sessionStorage",
      storeAuthStateInCookie: false,
    },
  };

  initPromise = (async () => {
    const pca = new PublicClientApplication(config);
    await pca.initialize();
    // Handle redirect return, if any.
    await pca.handleRedirectPromise().catch(() => null);
    instance = pca;
    return pca;
  })();
  return initPromise;
}

export async function loginPopup(): Promise<AccountInfo> {
  const pca = await getMsal();
  const result = await pca.loginPopup({ scopes: MSAL_SCOPES, prompt: "select_account" });
  pca.setActiveAccount(result.account);
  return result.account;
}

export async function acquireGraphToken(): Promise<string> {
  const pca = await getMsal();
  const account = pca.getActiveAccount() ?? pca.getAllAccounts()[0];
  if (!account) throw new Error("Not signed in");
  try {
    const res = await pca.acquireTokenSilent({ account, scopes: ["Mail.Read"] });
    return res.accessToken;
  } catch (e) {
    if (e instanceof InteractionRequiredAuthError) {
      const res = await pca.acquireTokenPopup({ scopes: ["Mail.Read"] });
      return res.accessToken;
    }
    throw e;
  }
}

export async function msalSignOut() {
  const pca = await getMsal();
  const account = pca.getActiveAccount() ?? pca.getAllAccounts()[0];
  pca.setActiveAccount(null);
  if (account) {
    try {
      await pca.clearCache({ account });
    } catch {
      /* ignore */
    }
  }
}

export async function getActiveAccount(): Promise<AccountInfo | null> {
  const pca = await getMsal();
  return pca.getActiveAccount() ?? pca.getAllAccounts()[0] ?? null;
}
