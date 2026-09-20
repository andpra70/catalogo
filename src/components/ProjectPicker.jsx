import { useEffect, useState } from "react";
import { EMPTY_CATALOG_PROJECT, normalizeSavedItems, VFS_PROJECTS_INDEX } from "../models/storageModels";
import { projectUrl } from "../models/projectRoute";
import { readVfsJson } from "../services/vfsStorage";

function formatUpdatedAt(value) {
  if (!value) return "";
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? "" : date.toLocaleString("it-IT");
}

export default function ProjectPicker() {
  const [projects, setProjects] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  useEffect(() => {
    let active = true;
    readVfsJson(VFS_PROJECTS_INDEX, [])
      .then((items) => {
        if (active) setProjects(normalizeSavedItems(items));
      })
      .catch((loadError) => {
        if (active) setError(loadError?.message || "Errore caricamento progetti");
      })
      .finally(() => {
        if (active) setLoading(false);
      });
    return () => { active = false; };
  }, []);

  return (
    <main className="project-picker-page">
      <section className="project-picker-card">
        <div className="project-picker-kicker">Catalogo Opere</div>
        <h1>Seleziona un progetto</h1>
        <p>Nessun progetto specificato nell’URL. Scegli un catalogo salvato su VFS2.</p>

        {loading && <div className="project-picker-status">Caricamento progetti…</div>}
        {!loading && error && <div className="project-picker-status error">{error}</div>}
        {!loading && !error && projects.length === 0 && (
          <div className="project-picker-status project-picker-empty">
            <div>
              <strong>Nessun progetto disponibile.</strong>
              <span>Puoi creare ora il progetto iniziale “{EMPTY_CATALOG_PROJECT.name}”.</span>
            </div>
            <a href={projectUrl(EMPTY_CATALOG_PROJECT.id)} className="project-picker-create">
              Crea e apri “{EMPTY_CATALOG_PROJECT.name}”
            </a>
          </div>
        )}
        {!loading && !error && projects.length > 0 && (
          <div className="project-picker-list">
            {projects.map((project) => (
              <a key={project.id} href={projectUrl(project.id)} className="project-picker-item">
                <strong>{project.name}</strong>
                <span>
                  {project.id}
                  {formatUpdatedAt(project.updatedAt) ? ` · ${formatUpdatedAt(project.updatedAt)}` : ""}
                </span>
              </a>
            ))}
          </div>
        )}
      </section>
    </main>
  );
}
