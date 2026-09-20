import { normalizeSavedItems, VFS_CATALOG_ROOT } from "../models/storageModels";
function loadScript(src, globalName) {
  if (window[globalName]) return Promise.resolve();
  return new Promise((resolve, reject) => {
    const script = document.createElement("script"); script.src = src; script.onload = resolve;
    script.onerror = () => reject(new Error(`Impossibile caricare ${src}`)); document.head.appendChild(script);
  });
}
export async function initializeVfsStorage() {
  await loadScript("/auth/widget.js", "VfsAuth"); await loadScript("/vfs/widget.js", "VfsWidget");
}
export async function ensureCatalogDirectories() {
  for (const path of [VFS_CATALOG_ROOT, `${VFS_CATALOG_ROOT}/projects`, `${VFS_CATALOG_ROOT}/themes`, `${VFS_CATALOG_ROOT}/images`]) {
    try { await window.VfsWidget.mkdir(path); } catch { /* directory marker may already exist */ }
  }
}
export async function readVfsJson(path, fallback = null) {
  const token = await window.VfsAuth.getAccessToken();
  const response = await fetch(window.VfsWidget.downloadUrl(path, false), { headers: { Authorization: `Bearer ${token}` } });
  if (response.status === 404) return fallback;
  if (!response.ok) throw new Error(`Lettura VFS fallita (${response.status})`);
  return response.json();
}
export const writeVfsJson = (path, value) => window.VfsWidget.salvaFileTesto(path, JSON.stringify(value, null, 2), "application/json");
export const uploadVfsFile = (directory, file, name) => window.VfsWidget.upload(directory, file, { name });
export async function readVfsDataUrl(path) {
  const token = await window.VfsAuth.getAccessToken();
  const response = await fetch(window.VfsWidget.downloadUrl(path, false), { headers: { Authorization: `Bearer ${token}` } });
  if (!response.ok) return "";
  const blob = await response.blob();
  return new Promise((resolve, reject) => {
    const reader = new FileReader(); reader.onload = () => resolve(String(reader.result || "")); reader.onerror = reject; reader.readAsDataURL(blob);
  });
}
export async function deleteVfsFile(path) {
  const token = await window.VfsAuth.getAccessToken();
  const response = await fetch("/vfs/api/file", {
    method: "DELETE", headers: { Authorization: `Bearer ${token}`, "content-type": "application/json" }, body: JSON.stringify({ path }),
  });
  if (!response.ok) throw new Error(`Eliminazione VFS fallita (${response.status})`);
}
export async function loadSavedIndexes(projectIndexPath, themeIndexPath) {
  const [projects, themes] = await Promise.all([readVfsJson(projectIndexPath, []), readVfsJson(themeIndexPath, [])]);
  return { projects: normalizeSavedItems(projects), themes: normalizeSavedItems(themes) };
}
