import React, { useEffect, useState } from "react";
import ReactDOM from "react-dom/client";
import App from "./App.jsx";
import AuthGate from "./components/AuthGate.jsx";
import GdprWidget from "./components/GdprWidget.jsx";
import ProjectPicker from "./components/ProjectPicker.jsx";
import { getProjectIdFromUrl } from "./models/projectRoute";
import "./styles.css";

function CatalogBootstrap() {
  const projectId = getProjectIdFromUrl();
  return projectId ? <App initialProjectId={projectId} /> : <ProjectPicker />;
}

function PublicCatalog({ slug }) {
  const [value, setValue] = useState(null);
  const [error, setError] = useState("");
  useEffect(() => {
    fetch(`/vfs/public/catalogo-opere/${encodeURIComponent(slug)}/project.json`)
      .then((response) => {
        if (!response.ok) throw new Error(`Pubblicazione non disponibile (${response.status})`);
        return response.json();
      })
      .then(setValue)
      .catch((reason) => setError(reason.message));
  }, [slug]);
  if (error) return <main className="auth-gate"><section><h1>Catalogo non disponibile</h1><p>{error}</p></section></main>;
  if (!value) return <main className="auth-gate"><section><p>Caricamento catalogo…</p></section></main>;
  return <App initialPublicState={value} readOnly />;
}

const publicMatch = window.location.pathname.match(/\/pub\/([^/]+)\/?$/i);

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    {publicMatch?.[1] ? <PublicCatalog slug={decodeURIComponent(publicMatch[1])} /> : <AuthGate><CatalogBootstrap /></AuthGate>}
    <GdprWidget />
  </React.StrictMode>,
);
