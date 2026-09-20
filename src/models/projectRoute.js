export const PROJECT_QUERY_PARAM = "project";

export function normalizeProjectId(value) {
  const decoded = String(value || "").trim();
  return decoded && !decoded.includes("/") && decoded !== "." && decoded !== ".." ? decoded : null;
}

export function getProjectIdFromUrl(locationLike = window.location) {
  const queryId = normalizeProjectId(new URLSearchParams(locationLike.search).get(PROJECT_QUERY_PARAM));
  if (queryId) return queryId;

  const match = String(locationLike.pathname || "").match(/\/project\/([^/]+)\/?$/i);
  if (!match) return null;
  try {
    return normalizeProjectId(decodeURIComponent(match[1]));
  } catch {
    return null;
  }
}

export function projectUrl(projectId) {
  const url = new URL(window.location.href);
  url.searchParams.set(PROJECT_QUERY_PARAM, normalizeProjectId(projectId) || "");
  return `${url.pathname}${url.search}${url.hash}`;
}
