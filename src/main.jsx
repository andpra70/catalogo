import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App.jsx";
import AuthGate from "./components/AuthGate.jsx";
import ProjectPicker from "./components/ProjectPicker.jsx";
import { getProjectIdFromUrl } from "./models/projectRoute";
import "./styles.css";

function CatalogBootstrap() {
  const projectId = getProjectIdFromUrl();
  return projectId ? <App initialProjectId={projectId} /> : <ProjectPicker />;
}

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode><AuthGate><CatalogBootstrap /></AuthGate></React.StrictMode>,
);
