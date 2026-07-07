import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { useState } from "react";
import { Mail, Loader2, UserRound } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Textarea } from "@/components/ui/textarea";
import { Card } from "@/components/ui/card";
import { toast } from "sonner";
import {
  connect,
  enterGuestMode,
  parseCredentialsLine,
  signInMicrosoft,
} from "@/lib/graph";
import { isMsalConfigured } from "@/lib/msal";

export const Route = createFileRoute("/")({
  component: LoginPage,
});

function LoginPage() {
  const navigate = useNavigate();
  const [value, setValue] = useState("");
  const [loading, setLoading] = useState(false);
  const [msalLoading, setMsalLoading] = useState(false);
  const msalReady = isMsalConfigured();

  const handleConnect = async () => {
    setLoading(true);
    try {
      const creds = parseCredentialsLine(value);
      await connect(creds);
      toast.success("Connected", { description: creds.email });
      navigate({ to: "/mailbox" });
    } catch (e) {
      toast.error("Connection failed", {
        description: e instanceof Error ? e.message : String(e),
      });
    } finally {
      setLoading(false);
    }
  };

  const handleMsal = async () => {
    setMsalLoading(true);
    try {
      await signInMicrosoft();
      toast.success("Signed in with Microsoft");
      navigate({ to: "/mailbox" });
    } catch (e) {
      toast.error("Microsoft sign-in failed", {
        description: e instanceof Error ? e.message : String(e),
      });
    } finally {
      setMsalLoading(false);
    }
  };

  const handleGuest = () => {
    enterGuestMode();
    navigate({ to: "/mailbox" });
  };

  return (
    <div className="min-h-screen w-full bg-gradient-to-br from-background via-background to-muted flex items-center justify-center px-4 py-10">
      <Card className="w-full max-w-2xl p-8 shadow-2xl border-border/60">
        <div className="flex items-center gap-3 mb-6">
          <div className="h-11 w-11 rounded-lg bg-primary text-primary-foreground grid place-items-center">
            <Mail className="h-6 w-6" />
          </div>
          <div>
            <h1 className="text-2xl font-semibold tracking-tight">Mailbox</h1>
            <p className="text-sm text-muted-foreground">
              Sign in with Microsoft, paste a refresh token, or explore in guest mode.
            </p>
          </div>
        </div>

        <Button
          onClick={handleMsal}
          disabled={!msalReady || msalLoading}
          className="w-full h-11 mb-4"
          size="lg"
        >
          {msalLoading ? (
            <Loader2 className="h-4 w-4 animate-spin" />
          ) : (
            <>
              <svg viewBox="0 0 23 23" className="h-4 w-4 mr-2" aria-hidden>
                <path fill="#f25022" d="M1 1h10v10H1z" />
                <path fill="#7fba00" d="M12 1h10v10H12z" />
                <path fill="#00a4ef" d="M1 12h10v10H1z" />
                <path fill="#ffb900" d="M12 12h10v10H12z" />
              </svg>
              Sign in with Microsoft
            </>
          )}
        </Button>
        {!msalReady && (
          <p className="text-xs text-muted-foreground mb-4">
            Set <code className="font-mono">VITE_CLIENT_ID</code> (and optional{" "}
            <code className="font-mono">VITE_TENANT_ID</code>) to enable Microsoft sign-in.
          </p>
        )}

        <div className="relative my-2">
          <div className="absolute inset-0 flex items-center">
            <div className="w-full border-t border-border" />
          </div>
          <div className="relative flex justify-center text-xs uppercase">
            <span className="bg-card px-2 text-muted-foreground">or</span>
          </div>
        </div>

        <label className="text-sm font-medium mb-2 mt-4 block" htmlFor="creds">
          Credentials
        </label>
        <Textarea
          id="creds"
          value={value}
          onChange={(e) => setValue(e.target.value)}
          placeholder="email|password|refresh_token|client_id|tenant_id (optional)"
          className="min-h-32 font-mono text-sm resize-y"
          spellCheck={false}
          autoComplete="off"
        />
        <p className="mt-2 text-xs text-muted-foreground">
          Format: <code className="font-mono">email|password|refresh_token|client_id</code> or add
          optional <code className="font-mono">|tenant_id</code>.
          Credentials stay in memory only — never saved to storage.
        </p>

        <div className="mt-6 flex flex-col sm:flex-row gap-3">
          <Button
            onClick={handleConnect}
            disabled={loading || !value.trim()}
            variant="outline"
            className="flex-1 h-11"
            size="lg"
          >
            {loading ? <Loader2 className="h-4 w-4 animate-spin" /> : "Connect"}
          </Button>
          <Button variant="ghost" onClick={handleGuest} className="h-11" size="lg">
            <UserRound className="h-4 w-4 mr-2" />
            Continue as guest
          </Button>
        </div>
      </Card>
    </div>
  );
}

