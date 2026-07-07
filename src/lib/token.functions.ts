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
    const body = new URLSearchParams({
      client_id: data.clientId,
      grant_type: "refresh_token",
      refresh_token: data.refreshToken,
      scope:
        "offline_access openid profile Mail.Read Mail.ReadWrite Mail.Send User.Read",
    });

    const res = await fetch(
      `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body,
      },
    );

    const text = await res.text();
    if (!res.ok) {
      throw new Error(`Token refresh failed: ${res.status} ${text}`);
    }
    return JSON.parse(text) as {
      access_token: string;
      refresh_token?: string;
      expires_in: number;
    };
  });
