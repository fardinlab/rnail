import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { useState } from "react";
import { Mail, Loader2, UserRound } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Textarea } from "@/components/ui/textarea";
import { Card } from "@/components/ui/card";
import { toast } from "sonner";
import { connect, enterGuestMode, parseCredentialsLine } from "@/lib/graph";

export const Route = createFileRoute("/")({
  component: LoginPage,
});

function LoginPage() {
  const navigate = useNavigate();
  const [value, setValue] = useState("");
  const [loading, setLoading] = useState(false);

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
              Sign in with a Microsoft refresh token, or explore in guest mode.
            </p>
          </div>
        </div>

        <label className="text-sm font-medium mb-2 block" htmlFor="creds">
          Credentials
        </label>
        <Textarea
          id="creds"
          value={value}
          onChange={(e) => setValue(e.target.value)}
          placeholder="email|password|refresh_token|client_id"
          className="min-h-40 font-mono text-sm resize-y"
          spellCheck={false}
          autoComplete="off"
        />
        <p className="mt-2 text-xs text-muted-foreground">
          Format: <code className="font-mono">email|password|refresh_token|client_id</code>.
          Credentials stay in memory only — never saved to storage.
        </p>

        <div className="mt-6 flex flex-col sm:flex-row gap-3">
          <Button
            onClick={handleConnect}
            disabled={loading || !value.trim()}
            className="flex-1 h-11"
            size="lg"
          >
            {loading ? (
              <Loader2 className="h-4 w-4 animate-spin" />
            ) : (
              "Connect"
            )}
          </Button>
          <Button
            variant="outline"
            onClick={handleGuest}
            className="h-11"
            size="lg"
          >
            <UserRound className="h-4 w-4 mr-2" />
            Continue as guest
          </Button>
        </div>
      </Card>
    </div>
  );
}
