export const VFS_CATALOG_ROOT = "catalogo-opere";
export const VFS_PROJECTS_INDEX = `${VFS_CATALOG_ROOT}/projects-index.json`;
export const VFS_THEMES_INDEX = `${VFS_CATALOG_ROOT}/themes-index.json`;
export const EMPTY_CATALOG_PROJECT = Object.freeze({ id: "progetto", name: "progetto" });
export const projectVfsPath = (id) => `${VFS_CATALOG_ROOT}/projects/${id}.json`;
export const themeVfsPath = (id) => `${VFS_CATALOG_ROOT}/themes/${id}.json`;
export const VFS_PROJECTS_DIRECTORY = `${VFS_CATALOG_ROOT}/projects`;

export function normalizeProjectFileSelection(value) {
  if (value?.volume !== "private" || value?.type !== "file") return null;
  const path = String(value.path || "").replace(/^\/+|\/+$/g, "");
  if (!path.toLowerCase().endsWith(".json")) return null;
  return { volume: "private", type: "file", path };
}

export function isCatalogProjectPayload(value) {
  const catalog = value?.catalog || value;
  return Boolean(
    catalog
      && typeof catalog === "object"
      && !Array.isArray(catalog)
      && Array.isArray(catalog.pages)
      && Array.isArray(catalog.works)
      && catalog.theme
      && typeof catalog.theme === "object"
      && !Array.isArray(catalog.theme),
  );
}

export function normalizeSavedItems(value) {
  return Array.isArray(value) ? value.filter((item) => item?.id && item?.name).map((item) => ({ id: String(item.id), name: String(item.name), updatedAt: item.updatedAt || null })) : [];
}
