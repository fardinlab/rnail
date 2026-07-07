import { createServerFn } from "@tanstack/react-start";
import { z } from "zod";

const InputSchema = z.object({
  clientId: z.string().min(1),
  refreshToken: z.string().min(1),
  tenantId: z.string().optional(),
});

export const exchangeRefreshToken = createServerFn({ method: "POST" })
  .inputValidator((data) => InputSchema.parse(data))
  .handler(async ({ data }) => {
    const tenant = data.tenantId || "common";
    const tokenUrl = `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`;
    const attempts: Array<string | undefined> = [
      "https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Mail.ReadWrite offline_access",
      "https://graph.microsoft.com/.default offline_access",
      "https://graph.microsoft.com/Mail.Read",
      undefined,
    ];

    let lastErrorText = "";
    for (const scope of attempts) {
      const body = new URLSearchParams({
        client_id: data.clientId,
        grant_type: "refresh_token",
        refresh_token: data.refreshToken,
      });
      if (scope) body.set("scope", scope);

      const res = await fetch(tokenUrl, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body,
      });

      const text = await res.text();
      if (res.ok) {
        return JSON.parse(text) as {
          access_token: string;
          refresh_token?: string;
          expires_in: number;
        };
      }
      lastErrorText = `${res.status} ${text}`;
    }

    if (lastErrorText.includes("AADSTS70000")) {
      throw new Error(
        "This refresh token is expired or was not granted Mail.Read for this client. Sign in with Microsoft once to grant access, then use the new session.",
      );
    }
    throw new Error(`Token refresh failed: ${lastErrorText}`);
  });
