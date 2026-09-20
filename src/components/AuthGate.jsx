import { useEffect, useState } from "react";
import { initializeVfsStorage } from "../services/vfsStorage";
export default function AuthGate({ children }) {
  const [ready, setReady] = useState(false); const [session, setSession] = useState(null); const [error, setError] = useState("");
  useEffect(() => { initializeVfsStorage().then(async () => {
    let current = window.VfsAuth.getSession();
    if (!current) { try { await window.VfsAuth.getAccessToken(); current = window.VfsAuth.getSession(); } catch { current = null; } }
    setSession(current); setReady(true);
  }).catch((err) => { setError(err.message); setReady(true); }); }, []);
  if (!session) return <main className="auth-gate"><section><h1>Catalogo Opere</h1><p>{ready ? "Accedi per gestire i cataloghi salvati su VFS2." : "Verifica sessione OAuth2…"}</p>{error && <p className="error">{error}</p>}<button disabled={!ready} onClick={() => window.VfsAuth.login()}>Accedi con Google</button></section></main>;
  return children;
}
