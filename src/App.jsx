import { useEffect, useMemo, useRef, useState } from "react";

const STORAGE_KEY = "catalogo-opere-state-v1";
const PROJECTS_INDEX_KEY = "catalogo-opere-projects-index-v1";
const PROJECT_STORAGE_PREFIX = "catalogo-opere-project:";
const THEMES_INDEX_KEY = "catalogo-opere-themes-index-v1";
const THEME_STORAGE_PREFIX = "catalogo-opere-theme:";
const IMAGE_DB_NAME = "catalogo-opere-assets";
const IMAGE_STORE_NAME = "images";
const DEFAULT_WORK_DPI = 75;
const MM_PER_INCH = 25.4;
const DEFAULT_PREFACE_TITLE_MD = "# Prefazione";
const DEFAULT_PREFACE_BODY_MD =
  "Questo catalogo raccoglie una selezione di opere con una sequenza pensata per accompagnare la lettura tra immagini, ritmo di pagina e apparati testuali.\n\n## Linea curatoriale\n- relazione tra *materia* e luce\n- dialogo tra archivio e presente\n- attenzione al ritmo visivo di pagina\n\n> Questo testo e un template: sostituiscilo con la prefazione definitiva.";
const DEFAULT_BACK_SUMMARY_MD =
  "## Sintesi\nUna selezione di opere tra fotografia, pittura e immagini d'autore che esplora materia, luce e memoria in forma di catalogo editoriale.\n\n### Focus\n- sequenza narrativa per nuclei tematici\n- alternanza tra immagini e testi\n- attenzione al progetto grafico";
const DEFAULT_BACK_BIO_MD =
  "# Biografia Autore\n## Profilo\n**Andrea Rossi** (1987) vive e lavora a Milano.\nE un artista visivo che unisce *fotografia*, **pittura** e pratiche editoriali.\n\n### Ricerca\n- memoria e archivio\n- paesaggio contemporaneo\n- rapporto tra immagine e narrazione\n\n> Nota curatoriale: il suo lavoro alterna rigore documentario e visione poetica.\n\n## Percorso\n1. Formazione in arti visive e fotografia\n2. Prime mostre collettive in spazi indipendenti\n3. Sviluppo di progetti ibridi tra stampa e installazione";
const DEFAULT_INTRO_CURATORIAL_MD =
  "# Introduzione\nQuesto catalogo nasce come strumento di lettura e di lavoro: non solo una raccolta di immagini, ma un percorso tra opere, materiali e relazioni.\n\n## Intento editoriale\nLa sequenza delle pagine costruisce una progressione che alterna visione ravvicinata e visione d'insieme, con l'obiettivo di valorizzare ritmo, pause e contrasti.\n\n### Obiettivi\n- restituire il contesto di produzione delle opere\n- evidenziare continuita e differenze tra i cicli\n- offrire una consultazione chiara per studio, archivio e presentazione\n\n## Testo curatoriale\nLa selezione propone un attraversamento tematico tra **materia**, *luce* e memoria visiva. Ogni nucleo mette in dialogo immagini con scale differenti, affinita formali e scarti narrativi.\n\n### Chiavi di lettura\n1. rapporto tra superficie e profondita\n2. tensione tra documento e interpretazione\n3. costruzione di una grammatica visiva coerente\n\n> Nota: questo testo e un template di base. Personalizzalo con riferimenti puntuali a mostra, opere e cronologia.\n\nPer approfondimenti critici: [scheda progetto](https://example.com).";

const FONT_OPTIONS = [
  { label: "Cormorant Garamond", value: "'Cormorant Garamond', serif" },
  { label: "IBM Plex Serif", value: "'IBM Plex Serif', serif" },
  { label: "Space Grotesk", value: "'Space Grotesk', sans-serif" },
  { label: "Playfair Display", value: "'Playfair Display', serif" },
  { label: "Libre Baskerville", value: "'Libre Baskerville', serif" },
  { label: "Lora", value: "'Lora', serif" },
  { label: "Merriweather", value: "'Merriweather', serif" },
  { label: "Bodoni Moda", value: "'Bodoni Moda', serif" },
  { label: "DM Serif Display", value: "'DM Serif Display', serif" },
  { label: "Spectral", value: "'Spectral', serif" },
  { label: "Source Serif 4", value: "'Source Serif 4', serif" },
  { label: "Crimson Pro", value: "'Crimson Pro', serif" },
  { label: "Manrope", value: "'Manrope', sans-serif" },
  { label: "Plus Jakarta Sans", value: "'Plus Jakarta Sans', sans-serif" },
  { label: "Work Sans", value: "'Work Sans', sans-serif" },
  { label: "Archivo", value: "'Archivo', sans-serif" },
  { label: "Sora", value: "'Sora', sans-serif" },
  { label: "Outfit", value: "'Outfit', sans-serif" },
  { label: "Inter Tight", value: "'Inter Tight', sans-serif" },
  { label: "Fraunces", value: "'Fraunces', serif" },
  { label: "EB Garamond", value: "'EB Garamond', serif" },
];

const WORK_FIELDS = [
  ["title", "Titolo"],
  ["author", "Autore"],
  ["year", "Anno"],
  ["type", "Tipo"],
  ["technique", "Tecnica"],
  ["dimensions", "Dimensioni"],
  ["dpi", "DPI"],
  ["inventory", "Inventario"],
  ["location", "Collocazione"],
  ["notes", "Note"],
];

const PAGE_FORMATS = [
  { id: "a4-portrait", label: "A4 Verticale", width: 210, height: 297 },
  { id: "square", label: "Quadrato", width: 240, height: 240 },
  { id: "landscape", label: "Orizzontale", width: 297, height: 210 },
  { id: "a5-portrait", label: "A5 Verticale", width: 148, height: 210 },
  { id: "a5-landscape", label: "A5 Orizzontale", width: 210, height: 148 },
  { id: "a6-portrait", label: "A6 Verticale", width: 105, height: 148 },
  { id: "a6-landscape", label: "A6 Orizzontale", width: 148, height: 105 },
  { id: "a3-portrait", label: "A3 Verticale", width: 297, height: 420 },
  { id: "a3-landscape", label: "A3 Orizzontale", width: 420, height: 297 },
  { id: "letter-portrait", label: "US Letter Verticale", width: 216, height: 279 },
  { id: "letter-landscape", label: "US Letter Orizzontale", width: 279, height: 216 },
  { id: "legal-portrait", label: "US Legal Verticale", width: 216, height: 356 },
  { id: "b5-portrait", label: "B5 Verticale", width: 176, height: 250 },
  { id: "b5-landscape", label: "B5 Orizzontale", width: 250, height: 176 },
  { id: "square-small", label: "Quadrato 210", width: 210, height: 210 },
];

function getPageFormat(formatId) {
  return PAGE_FORMATS.find((f) => f.id === formatId) || PAGE_FORMATS[0];
}

function uid(prefix = "id") {
  return `${prefix}_${Math.random().toString(36).slice(2, 10)}`;
}

function mmToCssPx(mm) {
  return (Number(mm) * 96) / 25.4;
}

function slugifyFileBaseName(value, fallback = "catalogo-book") {
  const raw = String(value || "").trim();
  const base = (raw || fallback)
    .replace(/[^\w\-]+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "")
    .toLowerCase();
  return base || fallback;
}

function clamp01(value, fallback = 1) {
  const n = Number(value);
  if (!Number.isFinite(n)) return fallback;
  return Math.max(0, Math.min(1, n));
}

function parseColorToRgbaParts(input) {
  const raw = String(input || "").trim();
  if (!raw) return null;
  const hexMatch = raw.match(/^#([0-9a-fA-F]{3,4}|[0-9a-fA-F]{6}|[0-9a-fA-F]{8})$/);
  if (hexMatch) {
    const hex = hexMatch[1];
    if (hex.length === 3 || hex.length === 4) {
      const r = parseInt(hex[0] + hex[0], 16);
      const g = parseInt(hex[1] + hex[1], 16);
      const b = parseInt(hex[2] + hex[2], 16);
      const a = hex.length === 4 ? parseInt(hex[3] + hex[3], 16) / 255 : 1;
      return { r, g, b, a };
    }
    if (hex.length === 6 || hex.length === 8) {
      const r = parseInt(hex.slice(0, 2), 16);
      const g = parseInt(hex.slice(2, 4), 16);
      const b = parseInt(hex.slice(4, 6), 16);
      const a = hex.length === 8 ? parseInt(hex.slice(6, 8), 16) / 255 : 1;
      return { r, g, b, a };
    }
  }
  const rgbMatch = raw.match(/^rgba?\(\s*([0-9.]+)\s*[, ]\s*([0-9.]+)\s*[, ]\s*([0-9.]+)(?:\s*[,/]\s*([0-9.]+))?\s*\)$/i);
  if (rgbMatch) {
    const r = clampNum(Math.round(Number(rgbMatch[1]) || 0), 0, 255);
    const g = clampNum(Math.round(Number(rgbMatch[2]) || 0), 0, 255);
    const b = clampNum(Math.round(Number(rgbMatch[3]) || 0), 0, 255);
    const a = clamp01(rgbMatch[4] == null ? 1 : Number(rgbMatch[4]), 1);
    return { r, g, b, a };
  }
  return null;
}

function rgbaPartsToCss({ r, g, b, a }) {
  const alpha = Math.round(clamp01(a, 1) * 1000) / 1000;
  return `rgba(${clampNum(Math.round(r), 0, 255)}, ${clampNum(Math.round(g), 0, 255)}, ${clampNum(Math.round(b), 0, 255)}, ${alpha})`;
}

function toHex2(value) {
  return clampNum(Math.round(value), 0, 255).toString(16).padStart(2, "0");
}

function colorToHexAlpha(color, fallback = "rgba(255, 255, 255, 0.42)") {
  const parsed = parseColorToRgbaParts(color) || parseColorToRgbaParts(fallback) || { r: 255, g: 255, b: 255, a: 0.42 };
  return {
    hex: `#${toHex2(parsed.r)}${toHex2(parsed.g)}${toHex2(parsed.b)}`,
    alpha: clamp01(parsed.a, 0.42),
  };
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function isSafeLinkUrl(url) {
  const raw = String(url || "").trim();
  if (!raw) return false;
  return /^(https?:\/\/|mailto:|tel:|\/(?!\/)|#)/i.test(raw);
}

function renderMarkdownInline(md) {
  const codeTokens = [];
  let text = escapeHtml(String(md || "")).replace(/\r\n/g, "\n");

  text = text.replace(/`([^`\n]+)`/g, (_, code) => {
    const token = `@@CODE_${codeTokens.length}@@`;
    codeTokens.push(`<code>${escapeHtml(code)}</code>`);
    return token;
  });

  text = text.replace(/\[([^\]]+)\]\(([^)\s]+)\)/g, (_, label, url) => {
    if (!isSafeLinkUrl(url)) return escapeHtml(label);
    return `<a href="${escapeHtml(url)}" target="_blank" rel="noopener noreferrer">${escapeHtml(label)}</a>`;
  });
  text = text.replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>");
  text = text.replace(/\*([^*]+)\*/g, "<em>$1</em>");
  text = text.replace(/\n/g, "<br/>");

  codeTokens.forEach((html, idx) => {
    text = text.replace(`@@CODE_${idx}@@`, html);
  });
  return text;
}

function renderMarkdownToHtml(md, { inline = false } = {}) {
  const source = String(md || "").replace(/\r\n/g, "\n").trim();
  if (!source) return "";
  if (inline) return renderMarkdownInline(source);
  const lines = source.split("\n");
  const out = [];
  let i = 0;
  const isHeading = (line) => /^\s*#{1,7}(?:\s+|$)/.test(line);
  const isQuote = (line) => /^\s*>\s?/.test(line);
  const isUl = (line) => /^\s*[-*+]\s+/.test(line);
  const isOl = (line) => /^\s*\d+[.)]\s+/.test(line);

  while (i < lines.length) {
    const line = lines[i];
    if (!line.trim()) {
      i += 1;
      continue;
    }

    const heading = line.match(/^\s*(#{1,7})(?:\s+|$)(.*?)\s*#*\s*$/);
    if (heading) {
      const level = heading[1].length;
      const rawHeadingText = String(heading[2] || "").trim();
      const content = rawHeadingText ? renderMarkdownInline(rawHeadingText) : "&nbsp;";
      out.push(level <= 6 ? `<h${level}>${content}</h${level}>` : `<p class="md-h7">${content}</p>`);
      i += 1;
      continue;
    }

    if (isQuote(line)) {
      const quoteLines = [];
      while (i < lines.length && isQuote(lines[i])) {
        quoteLines.push(lines[i].replace(/^\s*>\s?/, ""));
        i += 1;
      }
      out.push(`<blockquote>${renderMarkdownInline(quoteLines.join("\n"))}</blockquote>`);
      continue;
    }

    if (isUl(line)) {
      const items = [];
      while (i < lines.length && isUl(lines[i])) {
        const item = lines[i].replace(/^\s*[-*+]\s+/, "");
        const task = item.match(/^\[( |x|X)\]\s+(.*)$/);
        if (!task) items.push(`<li>${renderMarkdownInline(item)}</li>`);
        else {
          const checked = task[1].toLowerCase() === "x" ? " checked" : "";
          items.push(`<li class="task-item"><input type="checkbox" disabled${checked} /> <span>${renderMarkdownInline(task[2])}</span></li>`);
        }
        i += 1;
      }
      out.push(`<ul>${items.join("")}</ul>`);
      continue;
    }

    if (isOl(line)) {
      const items = [];
      while (i < lines.length && isOl(lines[i])) {
        items.push(`<li>${renderMarkdownInline(lines[i].replace(/^\s*\d+[.)]\s+/, ""))}</li>`);
        i += 1;
      }
      out.push(`<ol>${items.join("")}</ol>`);
      continue;
    }

    const para = [];
    while (i < lines.length && lines[i].trim() && !isHeading(lines[i]) && !isQuote(lines[i]) && !isUl(lines[i]) && !isOl(lines[i])) {
      para.push(lines[i]);
      i += 1;
    }
    out.push(`<p>${renderMarkdownInline(para.join("\n"))}</p>`);
  }

  return out.join("");
}

function normalizeWorkDpi(value) {
  const dpi = Number(value);
  if (Number.isFinite(dpi) && dpi > 0) return Math.round(dpi);
  return DEFAULT_WORK_DPI;
}

function normalizeWorkData(work) {
  return { ...(work || {}), dpi: normalizeWorkDpi(work?.dpi) };
}

function readImagePhysicalSizeMm(imageDimensions, dpiValue) {
  const widthPx = Number(imageDimensions?.width);
  const heightPx = Number(imageDimensions?.height);
  const dpi = normalizeWorkDpi(dpiValue);
  if (!(widthPx > 0 && heightPx > 0)) return null;
  return {
    width: (widthPx * MM_PER_INCH) / dpi,
    height: (heightPx * MM_PER_INCH) / dpi,
    aspect: widthPx / heightPx,
  };
}

function createTextBlock(text = "Nuovo testo") {
  return {
    id: uid("txt"),
    kind: "text",
    x: 40,
    y: 48,
    w: 220,
    h: 100,
    text,
    fontSize: 18,
    fontWeight: 500,
    color: "#1f1f1f",
    bgColor: "rgba(255, 255, 255, 0.42)",
    align: "left",
    borderWidthPct: 5,
    borderColor: "#ffffff",
  };
}

function createInsideFrontCoverPage(marginsOverride) {
  const page = createPage("inside-front-cover", "Seconda di copertina", marginsOverride);
  page.showPageNumber = false;
  page.bgColor = "#ffffff";
  page.textBlocks = [
    {
      ...createTextBlock(
        "Credits\nCuratela: Nome Curatore\nTesti: Nome Autore\nProgetto grafico: Studio / Designer\nFotografie: Autore / Archivio\nStampa: Tipografia\nAnno: 2026",
      ),
      x: 28,
      y: 36,
      w: 300,
      h: 150,
      fontSize: 12,
      fontWeight: 400,
      align: "left",
    },
  ];
  return page;
}

function createInsideBackCoverPage(marginsOverride) {
  const page = createPage("inside-back-cover", "Terza di copertina", marginsOverride);
  page.showPageNumber = false;
  page.bgColor = "#ffffff";
  page.textBlocks = [];
  return page;
}

function createPrefacePage(marginsOverride, bgColor = "#ffffff", borderPct = 3, pageFormatId = "a4-portrait", themeCfg = null) {
  const page = createPage("page", "Prefazione", marginsOverride);
  page.bgColor = bgColor;
  const titleMd = themeCfg?.defaultPrefaceTitleMd || DEFAULT_PREFACE_TITLE_MD;
  const bodyMd = themeCfg?.defaultPrefaceBodyMd || DEFAULT_PREFACE_BODY_MD;
  const area = getPageContentBounds(page, pageFormatId);
  const padX = Math.max(12, Math.round(area.width * 0.04));
  const topPad = Math.max(14, Math.round(area.height * 0.06));
  const interGap = Math.max(10, Math.round(area.height * 0.035));
  const bottomPad = Math.max(12, Math.round(area.height * 0.05));
  const titleW = Math.max(120, area.width - padX * 2);
  const titleFont = Math.max(18, Math.round(area.height * 0.08));
  const titleBorderPx = borderPxFromPercent(titleW, borderPct, 0, 64);
  const titleH = Math.max(36, Math.round(titleFont * 1.25 + titleBorderPx * 2 + 10));
  const bodyW = titleW;
  const bodyFont = Math.max(12, Math.round(area.height * 0.048));
  const bodyBorderPx = borderPxFromPercent(bodyW, borderPct, 0, 64);
  const bodyY = topPad + titleH + interGap;
  const bodyH = Math.max(80, area.height - bodyY - bottomPad);
  page.textBlocks = [
    {
      ...createTextBlock(titleMd),
      x: padX,
      y: topPad,
      w: titleW,
      h: titleH,
      fontSize: titleFont,
      fontWeight: 700,
      align: "left",
      borderWidthPct: borderPct,
    },
    {
      ...createTextBlock(bodyMd),
      x: padX,
      y: bodyY,
      w: bodyW,
      h: Math.max(Math.round(bodyFont * 1.35 + bodyBorderPx * 2 + 12), bodyH),
      fontSize: bodyFont,
      fontWeight: 400,
      align: "left",
      borderWidthPct: borderPct,
    },
  ];
  return page;
}

function createPage(type = "page", title = "Pagina", marginsOverride) {
  return {
    id: uid("page"),
    type,
    title,
    bgColor: "#ffffff",
    pageNumber: 1,
    showPageNumber: true,
    pageNumberColor: "#6b614f",
    margins: marginsOverride || { top: 28, right: 28, bottom: 38, left: 28 },
    textBlocks: [],
    placements: [],
  };
}

function createRenderSpacerPage(id, title, marginsOverride) {
  return {
    id,
    type: "spacer",
    title,
    bgColor: "#ffffff",
    showPageNumber: false,
    pageNumberColor: "#6b614f",
    margins: marginsOverride || { top: 28, right: 28, bottom: 38, left: 28 },
    textBlocks: [],
    placements: [],
  };
}

function createDefaultState(themeSeed = null) {
  const defaultMargins = { top: 28, right: 28, bottom: 38, left: 28 };
  const seed = themeSeed || {};
  const theme = {
    fontFamily: FONT_OPTIONS[2].value,
    bodyFontSize: 15,
    titleFontSize: 26,
    fontWeight: 400,
    textColor: "#111111",
    accentColor: "#111111",
    uiTint: "#f5f5f5",
    paperShadow: "rgba(0, 0, 0, 0.10)",
    autoShowCaptionDefault: true,
    defaultPageBgColor: "#ffffff",
    defaultElementBorderPct: 3,
    defaultElementBorderColor: "#ffffff",
    defaultPageNumberColor: "#6b614f",
    defaultTextBgColor: "rgba(255, 255, 255, 0.42)",
    defaultPrefaceTitleMd: DEFAULT_PREFACE_TITLE_MD,
    defaultPrefaceBodyMd: DEFAULT_PREFACE_BODY_MD,
    defaultBackSummaryMd: DEFAULT_BACK_SUMMARY_MD,
    defaultBackBioMd: DEFAULT_BACK_BIO_MD,
    defaultIntroCuratorialMd: DEFAULT_INTRO_CURATORIAL_MD,
    ...seed,
    pageMargins: { ...defaultMargins, ...(seed.pageMargins || {}) },
  };
  const projectMargins = theme.pageMargins || defaultMargins;
  const coverFront = createPage("cover-front", "Copertina", projectMargins);
  coverFront.showPageNumber = false;
  coverFront.bgColor = "#ffffff";
  coverFront.textBlocks = [
    {
      ...createTextBlock("CATALOGO OPERE"),
      x: 46,
      y: 72,
      w: 280,
      h: 90,
      fontSize: 24,
      fontWeight: 700,
      align: "center",
      color: "#121212",
    },
    {
      ...createTextBlock("Mostra / Collezione"),
      x: 58,
      y: 176,
      w: 250,
      h: 60,
      fontSize: 14,
      fontWeight: 500,
      align: "center",
      color: "#555555",
    },
  ];

  const innerPage = createPage("page", "Pagina 1", projectMargins);
  innerPage.pageNumber = 1;
  innerPage.textBlocks = [{ ...createTextBlock(theme.defaultIntroCuratorialMd || DEFAULT_INTRO_CURATORIAL_MD), x: 36, y: 36, w: 290, h: 120 }];

  const coverBack = createPage("cover-back", "Retro copertina", projectMargins);
  coverBack.showPageNumber = false;
  coverBack.bgColor = "#ffffff";
  coverBack.textBlocks = [
    {
      ...createTextBlock(
        theme.defaultBackSummaryMd || DEFAULT_BACK_SUMMARY_MD,
      ),
      x: 30,
      y: 28,
      w: 300,
      h: 96,
      fontSize: 13,
      fontWeight: 500,
      align: "left",
    },
    {
      ...createTextBlock(
        theme.defaultBackBioMd || DEFAULT_BACK_BIO_MD,
      ),
      x: 30,
      y: 132,
      w: 300,
      h: 116,
      fontSize: 12,
      fontWeight: 400,
      align: "left",
    },
    {
      ...createTextBlock("Edizione 1/200\nCodice edizione: CAT-2026-001"),
      x: 30,
      y: 258,
      w: 300,
      h: 42,
      fontSize: 11,
      fontWeight: 500,
      align: "left",
    },
  ];

  return {
    projectTitle: "Progetto",
    works: [],
    selectedWorkId: null,
    pages: [coverFront, innerPage, coverBack],
    currentSpread: 0,
    activePageId: innerPage.id,
    selectedElement: null,
    theme,
    pageFormat: "a4-portrait",
    summaryPageEdits: {},
    specialPages: {
      insideFront: createInsideFrontCoverPage(projectMargins),
      insideBack: createInsideBackCoverPage(projectMargins),
    },
    layoutAssist: {
      snapToGrid: true,
      showGuides: true,
      gridSize: 12,
      snapThreshold: 8,
      boundMode: "margins",
      distributeRespectMargins: true,
    },
    bookViewMode: "real",
  };
}

function sanitizeWorkForStorage(work) {
  const { imageUrl, _imageFile, ...rest } = normalizeWorkData(work);
  return { ...rest };
}

function sanitizeStateForStorage(state) {
  return {
    ...state,
    works: (state.works || []).map(sanitizeWorkForStorage),
  };
}

function sanitizeStateForExport(state) {
  return {
    ...state,
    works: (state.works || []).map((work) => {
      const { _imageFile, ...rest } = normalizeWorkData(work);
      return rest;
    }),
  };
}

function loadProjectsIndex() {
  try {
    const raw = localStorage.getItem(PROJECTS_INDEX_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function saveProjectsIndex(index) {
  localStorage.setItem(PROJECTS_INDEX_KEY, JSON.stringify(index));
}

function projectStorageKey(projectId) {
  return `${PROJECT_STORAGE_PREFIX}${projectId}`;
}

function sanitizeThemeForStorage(theme) {
  const baseTheme = createDefaultState().theme;
  return {
    ...baseTheme,
    ...(theme || {}),
    pageMargins: {
      ...baseTheme.pageMargins,
      ...((theme || {}).pageMargins || {}),
    },
  };
}

function loadThemesIndex() {
  try {
    const raw = localStorage.getItem(THEMES_INDEX_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function saveThemesIndex(index) {
  localStorage.setItem(THEMES_INDEX_KEY, JSON.stringify(index));
}

function themeStorageKey(themeId) {
  return `${THEME_STORAGE_PREFIX}${themeId}`;
}

function openImageDb() {
  return new Promise((resolve, reject) => {
    if (typeof indexedDB === "undefined") {
      reject(new Error("IndexedDB non disponibile"));
      return;
    }
    const request = indexedDB.open(IMAGE_DB_NAME, 1);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(IMAGE_STORE_NAME)) {
        db.createObjectStore(IMAGE_STORE_NAME);
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error("Errore apertura IndexedDB"));
  });
}

async function idbPutImage(imageId, dataUrl) {
  if (!imageId || !dataUrl) return;
  const db = await openImageDb();
  await new Promise((resolve, reject) => {
    const tx = db.transaction(IMAGE_STORE_NAME, "readwrite");
    tx.objectStore(IMAGE_STORE_NAME).put(dataUrl, imageId);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error || new Error("Errore salvataggio immagine"));
  });
  db.close();
}

async function idbGetImage(imageId) {
  if (!imageId) return "";
  const db = await openImageDb();
  const result = await new Promise((resolve, reject) => {
    const tx = db.transaction(IMAGE_STORE_NAME, "readonly");
    const req = tx.objectStore(IMAGE_STORE_NAME).get(imageId);
    req.onsuccess = () => resolve(req.result || "");
    req.onerror = () => reject(req.error || new Error("Errore lettura immagine"));
  });
  db.close();
  return typeof result === "string" ? result : "";
}

async function idbDeleteImage(imageId) {
  if (!imageId) return;
  const db = await openImageDb();
  await new Promise((resolve, reject) => {
    const tx = db.transaction(IMAGE_STORE_NAME, "readwrite");
    tx.objectStore(IMAGE_STORE_NAME).delete(imageId);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error || new Error("Errore cancellazione immagine"));
  });
  db.close();
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return createDefaultState();
    const parsed = JSON.parse(raw);
    const base = createDefaultState();
    const theme = {
      ...base.theme,
      ...(parsed.theme || {}),
      pageMargins: {
        ...base.theme.pageMargins,
        ...(parsed.theme?.pageMargins || {}),
      },
    };
    const baseSpecial = base.specialPages || {};
    const parsedSpecial = parsed.specialPages || {};
    return {
      ...base,
      ...parsed,
      theme,
      summaryPageEdits: parsed.summaryPageEdits || {},
      specialPages: {
        insideFront: {
          ...baseSpecial.insideFront,
          ...(parsedSpecial.insideFront || {}),
          margins: { ...theme.pageMargins, ...(parsedSpecial.insideFront?.margins || baseSpecial.insideFront?.margins || {}) },
        },
        insideBack: {
          ...baseSpecial.insideBack,
          ...(parsedSpecial.insideBack || {}),
          margins: { ...theme.pageMargins, ...(parsedSpecial.insideBack?.margins || baseSpecial.insideBack?.margins || {}) },
        },
      },
      pages: (parsed.pages || base.pages).map((p) => ({
        ...p,
        margins: { ...theme.pageMargins, ...(p?.margins || {}) },
      })),
      works: (parsed.works || []).map((w) => ({ ...normalizeWorkData(w), imageUrl: "" })),
    };
  } catch {
    return createDefaultState();
  }
}

function workLabel(work) {
  return work.title?.trim() || work.inventory?.trim() || "Opera senza titolo";
}

function defaultPlacementCaption(placement, work) {
  if (placement?.captionOverride?.trim()) return placement.captionOverride;
  return [workLabel(work), [work?.author, work?.year].filter(Boolean).join(", ")].filter(Boolean).join("\n");
}

function summaryPagesFromWorks(
  works,
  marginsOverride,
  workPageMap = new Map(),
  summaryPageEdits = {},
  defaultBgColor = "#ffffff",
  defaultBorderPct = 3,
  defaultPageNumberColor = "#6b614f",
  pageFormatId = "a4-portrait",
) {
  const chunkSize = 14;
  const chunks = [];
  for (let i = 0; i < works.length; i += chunkSize) chunks.push(works.slice(i, i + chunkSize));
  return chunks.map((chunk, idx) => {
    const area = getPageContentBounds({ margins: marginsOverride || { top: 28, right: 26, bottom: 38, left: 26 } }, pageFormatId);
    const padX = Math.max(14, Math.round(area.width * 0.06));
    const titleY = Math.max(12, Math.round(area.height * 0.05));
    const titleH = Math.max(34, Math.round(area.height * 0.12));
    const listY = titleY + titleH + Math.max(8, Math.round(area.height * 0.03));
    const listH = Math.max(80, area.height - listY - Math.max(8, Math.round(area.height * 0.03)));
    const titleBlock = {
      ...createTextBlock(idx === 0 ? "Elenco opere" : `Elenco opere (${idx + 1})`),
      x: padX,
      y: titleY,
      w: Math.max(100, area.width - padX * 2),
      h: titleH,
      fontSize: Math.max(16, Math.round(area.height * 0.075)),
      fontWeight: 700,
      align: "left",
      borderWidthPct: defaultBorderPct,
    };
    const listText =
      chunk
        .map((work, rowIdx) => {
          const pageNo = workPageMap.get(work.id);
          const title = workLabel(work);
          const author = work.author?.trim() || "n/d";
          const year = work.year?.trim() || "n/d";
          const pageLabel = pageNo ?? "-";
          return `${idx * chunkSize + rowIdx + 1}. **${title}** — *${author}* — anno: ${year} — **pag. ${pageLabel}**`;
        })
        .join("\n") || "Nessuna opera inserita.";
    const listBlock = {
      ...createTextBlock(listText),
      x: padX,
      y: listY,
      w: Math.max(100, area.width - padX * 2),
      h: listH,
      fontSize: Math.max(11, Math.round(area.height * 0.045)),
      fontWeight: 400,
      align: "left",
      borderWidthPct: defaultBorderPct,
    };
    const generated = {
      id: `summary_${idx}`,
      type: "summary-page",
      title: `Elenco opere ${idx + 1}`,
      bgColor: defaultBgColor || "#ffffff",
      pageNumber: 1,
      showPageNumber: true,
      pageNumberColor: defaultPageNumberColor || "#6b614f",
      margins: marginsOverride || { top: 28, right: 26, bottom: 38, left: 26 },
      textBlocks: [titleBlock, listBlock],
      placements: [],
    };
    const edit = summaryPageEdits?.[generated.id];
    if (!edit) return generated;
    const editedBlocks = Array.isArray(edit.textBlocks) ? edit.textBlocks : generated.textBlocks;
    const mergedTextBlocks = [
      editedBlocks[0] ? { ...editedBlocks[0] } : generated.textBlocks[0],
      {
        ...(editedBlocks[1] ? { ...editedBlocks[1] } : generated.textBlocks[1]),
        text: generated.textBlocks[1]?.text || "",
      },
      ...editedBlocks.slice(2),
    ];
    return {
      ...generated,
      ...edit,
      margins: { ...(generated.margins || {}), ...(edit.margins || {}) },
      textBlocks: mergedTextBlocks,
      placements: edit.placements || generated.placements,
    };
  });
}

function scalePageLayoutForFormat(page, fromFormatId, toFormatId) {
  if (!page || fromFormatId === toFormatId) return page;
  const fromFmt = getPageFormat(fromFormatId);
  const toFmt = getPageFormat(toFormatId);
  const sx = (toFmt.width || 1) / Math.max(1, fromFmt.width || 1);
  const sy = (toFmt.height || 1) / Math.max(1, fromFmt.height || 1);
  const sf = Math.min(sx, sy);
  const scalePosX = (v) => Math.round((v || 0) * sx);
  const scalePosY = (v) => Math.round((v || 0) * sy);
  const scaleSize = (v, minVal) => Math.max(minVal, Math.round((v || 0) * sf));
  const scaleText = (t) => ({
    ...t,
    x: scalePosX(t.x),
    y: scalePosY(t.y),
    w: scaleSize(t.w, 40),
    h: scaleSize(t.h, 24),
    fontSize: Math.max(10, Math.round((t.fontSize || 16) * sf)),
  });
  const scalePlacement = (p) => ({
    ...p,
    x: scalePosX(p.x),
    y: scalePosY(p.y),
    w: scaleSize(p.w, 40),
    h: scaleSize(p.h, 40),
    captionX: scalePosX(p.captionX ?? p.x ?? 0),
    captionY: scalePosY(p.captionY ?? 0),
    captionW: scaleSize(p.captionW, 40),
    captionH: scaleSize(p.captionH, 24),
  });
  return {
    ...page,
    textBlocks: (page.textBlocks || []).map(scaleText),
    placements: (page.placements || []).map(scalePlacement),
  };
}

function estimateRenderedInnerBoundsForFormat(page, fromFormatId, toFormatId, canonicalMetric) {
  if (!canonicalMetric) return null;
  const fromFmt = getPageFormat(fromFormatId);
  const toFmt = getPageFormat(toFormatId);
  const m = page.margins || { top: 0, right: 0, bottom: 0, left: 0 };
  const estPageW = Math.max(40, (canonicalMetric.pageWidth || 0) * ((toFmt.width || 1) / Math.max(1, fromFmt.width || 1)));
  const estPageH = Math.max(40, (canonicalMetric.pageHeight || 0) * ((toFmt.height || 1) / Math.max(1, fromFmt.height || 1)));
  return {
    innerWidth: Math.max(40, estPageW - (m.left || 0) - (m.right || 0)),
    innerHeight: Math.max(40, estPageH - (m.top || 0) - (m.bottom || 0)),
  };
}

function fitPageElementsToEstimatedRenderedBoundsSoft(page, fromFormatId, toFormatId, canonicalMetric, boundMode = "margins") {
  if (!page || page.type === "summary") return page;
  const est = estimateRenderedInnerBoundsForFormat(page, fromFormatId, toFormatId, canonicalMetric);
  if (!est) return page;
  const box = getRenderedConstraintBox(page, est.innerWidth, est.innerHeight, boundMode);
  return {
    ...page,
    placements: (page.placements || []).map((p) => fitPlacementToBoundsSoft(p, box)),
    textBlocks: (page.textBlocks || []).map((t) => fitTextBlockToBoundsSoft(t, box)),
  };
}

function normalizePageElementsToRenderedBounds(page, innerWidth, innerHeight, boundMode = "margins") {
  if (!page || page.type === "summary") return page;
  const box = getRenderedConstraintBox(page, innerWidth, innerHeight, boundMode);
  return {
    ...page,
    placements: (page.placements || []).map((p) => normalizePlacementToBounds(p, box)),
    textBlocks: (page.textBlocks || []).map((t) => normalizeTextBlockToBounds(t, box)),
  };
}

function buildWorkFirstPageMapForCatalog(frontCover, insideFront, innerPages) {
  const map = new Map();
  const pages = [frontCover, insideFront, ...(innerPages || [])];
  let pageNo = 0;
  for (const page of pages) {
    if (!page) continue;
    const isCover = page.type === "cover-front" || page.type === "cover-back";
    const isSpacer = page.type === "spacer";
    const isInsideCover = page.type === "inside-front-cover" || page.type === "inside-back-cover";
    if (!isCover && !isSpacer && !isInsideCover) {
      pageNo += 1;
    }
    if (pageNo <= 0) continue;
    for (const pl of page.placements || []) {
      if (!map.has(pl.workId)) map.set(pl.workId, pageNo);
    }
  }
  return map;
}

function buildBookSpreads(pages) {
  if (!pages.length) return [];
  if (pages.length === 1) return [[null, pages[0]]];
  const frontCover = pages[0];
  const backCover = pages[pages.length - 1];
  const internals = [...pages.slice(1, -1)];
  if (internals.length % 2 !== 0) {
    internals.push(createRenderSpacerPage(`auto_spacer_${internals.length}`, "Pagina vuota", frontCover?.margins));
  }
  const spreads = [[null, frontCover]];
  for (let i = 0; i < internals.length; i += 2) {
    spreads.push([internals[i] || null, internals[i + 1] || null]);
  }
  spreads.push([backCover, null]);
  return spreads;
}

function buildTechnicalSpreads(pages) {
  const arr = [];
  for (let i = 0; i < pages.length; i += 2) {
    arr.push([pages[i] || null, pages[i + 1] || null]);
  }
  return arr;
}

function cloneWorkDraft(work) {
  const base = {
    id: uid("work"),
    title: "",
    author: "",
    year: "",
    type: "",
    technique: "",
    dimensions: "",
    dpi: DEFAULT_WORK_DPI,
    inventory: "",
    location: "",
    notes: "",
    imageUrl: "",
  };
  return { ...base, ...(work || {}) };
}

function createPlacementForWork(workId, x = 42, y = 42) {
  return {
    id: uid("plc"),
    workId,
    x,
    y,
    w: 150,
    h: 190,
    zIndex: 100,
    borderWidthPct: 5,
    borderColor: "#ffffff",
    imageOffsetX: 0,
    imageOffsetY: 0,
    imageScale: 1,
    showCaption: true,
    captionX: x,
    captionY: y + 196,
    captionW: 220,
    captionH: 52,
    captionBorderWidthPct: 5,
    captionBorderColor: "#ffffff",
  };
}

function nextPlacementLayerZ(page) {
  const values = (page?.placements || []).map((p) => Number(p?.zIndex)).filter((v) => Number.isFinite(v));
  return (values.length ? Math.max(...values) : 90) + 10;
}

function createAutoWorkPage(work, pageFormatId, imageDimensions) {
  const page = createPage("page", workLabel(work));
  page.bgColor = "#fbfbfb";
  const format = getPageFormat(pageFormatId);
  const margins = page.margins;
  const innerW = Math.max(220, format.width - (margins.left + margins.right));
  const innerH = Math.max(260, format.height - (margins.top + margins.bottom));
  const captionH = 50;
  const gap = 8;
  const maxImageW = Math.max(60, innerW);
  const maxImageH = Math.max(60, innerH - captionH - gap);
  const hasImageSize = Number(imageDimensions?.width) > 0 && Number(imageDimensions?.height) > 0;
  const fitted = hasImageSize
    ? fitImageMmToBox(work, imageDimensions, maxImageW, maxImageH, Number(imageDimensions.width) / Math.max(1, Number(imageDimensions.height)))
    : null;
  const placementW = fitted?.w ?? Math.min(220, Math.max(140, Math.round(innerW * 0.68)));
  const placementH = fitted?.h ?? Math.round(placementW * 1.12);
  const x = Math.max(0, Math.round((innerW - placementW) / 2));
  const y = Math.max(0, Math.round((innerH - placementH - captionH - gap) / 2));
  const placement = createPlacementForWork(work.id, x, y);
  placement.w = placementW;
  placement.h = placementH;
  placement.captionX = Math.max(0, Math.round((innerW - Math.max(220, placementW)) / 2));
  placement.captionY = y + placementH + gap;
  placement.captionW = Math.min(innerW, Math.max(220, placementW));
  placement.captionH = captionH;
  page.placements = [placement];
  return page;
}

function createAutoWorkPageFromImageAspect(work, pageFormatId, aspectRatio, boundsOverride) {
  const page = createPage("page", workLabel(work));
  page.bgColor = "#fbfbfb";
  const bounds = boundsOverride || getPageContentBounds(page, pageFormatId);
  const areaW = bounds.width;
  const areaH = bounds.height;
  const captionGap = 8;
  const safeAspect = Number.isFinite(aspectRatio) && aspectRatio > 0 ? aspectRatio : 1;
  const title = workLabel(work);
  const meta = [work.author, work.year].filter(Boolean).join(", ");
  const oneLineCaption = [title, meta].filter(Boolean).join(" — ");

  // Stima semplice per scegliere una didascalia più compatta quando possibile.
  const approxCharsPerLine = Math.max(18, Math.floor(areaW / 7.2));
  const preferOneLine = oneLineCaption.length <= approxCharsPerLine && safeAspect >= 1;
  const captionOverride = preferOneLine ? oneLineCaption : [title, meta].filter(Boolean).join("\n");
  const captionH = preferOneLine ? 36 : oneLineCaption.length > approxCharsPerLine * 1.6 ? 62 : 52;
  const maxImageW = Math.max(60, areaW);
  const maxImageH = Math.max(60, areaH - captionH - captionGap);

  const fitted = fitImageMmToBox(work, null, maxImageW, maxImageH, safeAspect);
  const w = clampNum(Math.round(fitted.w), 60, maxImageW);
  const h = clampNum(Math.round(fitted.h), 60, maxImageH);

  const x = Math.round((areaW - w) / 2);
  const y = Math.max(6, Math.round((areaH - h - captionH - captionGap) / 2));
  const placement = createPlacementForWork(work.id, x, y);
  placement.w = w;
  placement.h = h;
  placement.captionW = Math.min(areaW, Math.max(160, Math.round(w)));
  placement.captionH = captionH;
  placement.captionX = Math.round((areaW - placement.captionW) / 2);
  placement.captionY = clampNum(y + h + captionGap, 0, Math.max(0, areaH - captionH));
  placement.captionOverride = captionOverride;
  page.placements = [placement];
  return page;
}

function clampNum(value, min, max) {
  return Math.max(min, Math.min(max, Number.isFinite(value) ? value : min));
}

function fitImageMmToBox(work, imageDimensions, maxW, maxH, fallbackAspect = 1) {
  const safeMaxW = Math.max(1, Number(maxW) || 1);
  const safeMaxH = Math.max(1, Number(maxH) || 1);
  const physical = readImagePhysicalSizeMm(imageDimensions, work?.dpi);
  const safeAspect = Number.isFinite(physical?.aspect) && physical.aspect > 0 ? physical.aspect : Math.max(0.01, Number(fallbackAspect) || 1);

  let w = Number.isFinite(physical?.width) && physical.width > 0 ? physical.width : safeMaxW;
  let h = Number.isFinite(physical?.height) && physical.height > 0 ? physical.height : w / safeAspect;
  if (!(w > 0) || !(h > 0)) {
    w = safeMaxW;
    h = w / safeAspect;
  }

  const downScale = Math.min(1, safeMaxW / w, safeMaxH / h);
  w *= downScale;
  h *= downScale;
  return {
    w: clampNum(Math.round(w), 1, safeMaxW),
    h: clampNum(Math.round(h), 1, safeMaxH),
  };
}

function borderPxFromPercent(sizePx, pct, minPx = 0, maxPx = 64) {
  const p = Number.isFinite(Number(pct)) ? Number(pct) : 5;
  return clampNum(Math.round((Math.max(0, sizePx || 0) * p) / 100), minPx, maxPx);
}

function applyThemeDefaultsToPlacement(placement, theme) {
  if (!placement) return placement;
  const borderPctDefault = theme?.defaultElementBorderPct;
  const showCaptionDefault = theme?.autoShowCaptionDefault;
  const borderColorDefault = theme?.defaultElementBorderColor || theme?.accentColor;
  const captionBorderColorDefault = theme?.defaultCaptionBorderColor || theme?.accentColor || borderColorDefault;
  return {
    ...placement,
    borderWidthPct: Number.isFinite(Number(borderPctDefault)) ? Number(borderPctDefault) : placement.borderWidthPct,
    borderColor: borderColorDefault || placement.borderColor,
    showCaption: typeof showCaptionDefault === "boolean" ? showCaptionDefault : placement.showCaption,
    captionBorderWidthPct: Number.isFinite(Number(borderPctDefault)) ? Number(borderPctDefault) : placement.captionBorderWidthPct,
    captionBorderColor: captionBorderColorDefault || placement.captionBorderColor,
  };
}

function applyThemeAutoDefaultsToPage(page, theme) {
  if (!page) return page;
  const showCaptionDefault = theme?.autoShowCaptionDefault ?? true;
  const bgColorDefault = theme?.defaultPageBgColor || "#ffffff";
  const borderPctDefault = theme?.defaultElementBorderPct ?? 3;
  const borderColorDefault = theme?.defaultElementBorderColor || theme?.accentColor;
  const captionBorderColorDefault = theme?.defaultCaptionBorderColor || theme?.accentColor || borderColorDefault;
  const textBgDefault = theme?.defaultTextBgColor || "rgba(255, 255, 255, 0.42)";
  return {
    ...page,
    bgColor: bgColorDefault,
    placements: (page.placements || []).map((p) => ({
      ...p,
      showCaption: showCaptionDefault,
      borderWidthPct: borderPctDefault,
      borderColor: borderColorDefault || p.borderColor,
      captionBorderWidthPct: borderPctDefault,
      captionBorderColor: captionBorderColorDefault || p.captionBorderColor,
    })),
    textBlocks: (page.textBlocks || []).map((t) => ({
      ...t,
      borderWidthPct: t.borderWidthPct ?? borderPctDefault,
      bgColor: t.bgColor ?? textBgDefault,
    })),
  };
}

function getPageContentBounds(page, pageFormatId) {
  const format = getPageFormat(pageFormatId);
  const margins = page.margins || { top: 0, right: 0, bottom: 0, left: 0 };
  return {
    width: Math.max(40, format.width - (margins.left || 0) - (margins.right || 0)),
    height: Math.max(40, format.height - (margins.top || 0) - (margins.bottom || 0)),
  };
}

function getPageConstraintBox(page, pageFormatId, boundMode = "margins") {
  const format = getPageFormat(pageFormatId);
  const m = page.margins || { top: 0, right: 0, bottom: 0, left: 0 };
  const content = getPageContentBounds(page, pageFormatId);
  if (boundMode === "page") {
    return {
      minX: -(m.left || 0),
      minY: -(m.top || 0),
      maxWidth: Math.max(40, format.width),
      maxHeight: Math.max(40, format.height),
      maxXForWidth: (w) => content.width + (m.right || 0) - w,
      maxYForHeight: (h) => content.height + (m.bottom || 0) - h,
    };
  }
  return {
    minX: 0,
    minY: 0,
    maxWidth: content.width,
    maxHeight: content.height,
    maxXForWidth: (w) => content.width - w,
    maxYForHeight: (h) => content.height - h,
  };
}

function getRenderedConstraintBox(page, innerWidth, innerHeight, boundMode = "margins") {
  const m = page.margins || { top: 0, right: 0, bottom: 0, left: 0 };
  if (boundMode === "page") {
    return {
      minX: -(m.left || 0),
      minY: -(m.top || 0),
      maxWidth: Math.max(40, innerWidth + (m.left || 0) + (m.right || 0)),
      maxHeight: Math.max(40, innerHeight + (m.top || 0) + (m.bottom || 0)),
      maxXForWidth: (w) => innerWidth + (m.right || 0) - w,
      maxYForHeight: (h) => innerHeight + (m.bottom || 0) - h,
    };
  }
  return {
    minX: 0,
    minY: 0,
    maxWidth: Math.max(40, innerWidth),
    maxHeight: Math.max(40, innerHeight),
    maxXForWidth: (w) => innerWidth - w,
    maxYForHeight: (h) => innerHeight - h,
  };
}

function normalizeTextBlockToBounds(txt, box) {
  const w = clampNum(txt.w ?? 220, 40, box.maxWidth);
  const h = clampNum(txt.h ?? 100, 28, box.maxHeight);
  const x = clampNum(txt.x ?? 0, box.minX, Math.max(box.minX, box.maxXForWidth(w)));
  const y = clampNum(txt.y ?? 0, box.minY, Math.max(box.minY, box.maxYForHeight(h)));
  return { ...txt, x, y, w, h };
}

function normalizePlacementToBounds(p, box) {
  const w = clampNum(p.w ?? 150, 40, box.maxWidth);
  const h = clampNum(p.h ?? 150, 40, box.maxHeight);
  const x = clampNum(p.x ?? 0, box.minX, Math.max(box.minX, box.maxXForWidth(w)));
  const y = clampNum(p.y ?? 0, box.minY, Math.max(box.minY, box.maxYForHeight(h)));

  const captionW = clampNum(p.captionW ?? 220, 40, box.maxWidth);
  const captionH = clampNum(p.captionH ?? 52, 24, box.maxHeight);
  const captionX = clampNum(p.captionX ?? x, box.minX, Math.max(box.minX, box.maxXForWidth(captionW)));
  const defaultCaptionY = y + h + 8;
  const captionY = clampNum(p.captionY ?? defaultCaptionY, box.minY, Math.max(box.minY, box.maxYForHeight(captionH)));

  return { ...p, x, y, w, h, captionX, captionY, captionW, captionH };
}

function normalizePageElementsToBounds(page, pageFormatId, boundMode = "margins") {
  if (!page || page.type === "summary") return page;
  const box = getPageConstraintBox(page, pageFormatId, boundMode);
  return {
    ...page,
    placements: (page.placements || []).map((p) => normalizePlacementToBounds(p, box)),
    textBlocks: (page.textBlocks || []).map((t) => normalizeTextBlockToBounds(t, box)),
  };
}

function fitTextBlockToBoundsSoft(txt, box) {
  const rawW = Math.max(40, Number(txt.w ?? 220));
  const rawH = Math.max(28, Number(txt.h ?? 100));
  const w = rawW > box.maxWidth ? Math.max(40, box.maxWidth) : rawW;
  const h = rawH > box.maxHeight ? Math.max(28, box.maxHeight) : rawH;
  const x = clampNum(txt.x ?? 0, box.minX, Math.max(box.minX, box.maxXForWidth(w)));
  const y = clampNum(txt.y ?? 0, box.minY, Math.max(box.minY, box.maxYForHeight(h)));
  return { ...txt, x, y, w: Math.round(w), h: Math.round(h) };
}

function fitPlacementToBoundsSoft(p, box) {
  const rawW = Math.max(40, Number(p.w ?? 150));
  const rawH = Math.max(40, Number(p.h ?? 190));
  const w = rawW > box.maxWidth ? Math.max(40, box.maxWidth) : rawW;
  const h = rawH > box.maxHeight ? Math.max(40, box.maxHeight) : rawH;
  const x = clampNum(p.x ?? 0, box.minX, Math.max(box.minX, box.maxXForWidth(w)));
  const y = clampNum(p.y ?? 0, box.minY, Math.max(box.minY, box.maxYForHeight(h)));

  const rawCaptionW = Math.max(40, Number(p.captionW ?? 220));
  const rawCaptionH = Math.max(24, Number(p.captionH ?? 52));
  const captionW = rawCaptionW > box.maxWidth ? Math.max(40, box.maxWidth) : rawCaptionW;
  const captionH = rawCaptionH > box.maxHeight ? Math.max(24, box.maxHeight) : rawCaptionH;
  const defaultCaptionY = y + h + 8;
  const captionX = clampNum(p.captionX ?? x, box.minX, Math.max(box.minX, box.maxXForWidth(captionW)));
  const captionY = clampNum(p.captionY ?? defaultCaptionY, box.minY, Math.max(box.minY, box.maxYForHeight(captionH)));

  return {
    ...p,
    x: Math.round(x),
    y: Math.round(y),
    w: Math.round(w),
    h: Math.round(h),
    captionX: Math.round(captionX),
    captionY: Math.round(captionY),
    captionW: Math.round(captionW),
    captionH: Math.round(captionH),
  };
}

function fitPageElementsToBoundsSoft(page, pageFormatId, boundMode = "margins") {
  if (!page || page.type === "summary") return page;
  const box = getPageConstraintBox(page, pageFormatId, boundMode);
  return {
    ...page,
    placements: (page.placements || []).map((p) => fitPlacementToBoundsSoft(p, box)),
    textBlocks: (page.textBlocks || []).map((t) => fitTextBlockToBoundsSoft(t, box)),
  };
}

function readImageFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function readImageDimensions(src) {
  return new Promise((resolve) => {
    if (!src) {
      resolve(null);
      return;
    }
    const img = new Image();
    img.onload = () => resolve({ width: img.naturalWidth || img.width, height: img.naturalHeight || img.height });
    img.onerror = () => resolve(null);
    img.src = src;
  });
}

export default function App() {
  const [state, setState] = useState(loadState);
  const prevPageFormatRef = useRef(state.pageFormat);
  const skipNextPageFormatAdjustRef = useRef(false);
  const [workEditor, setWorkEditor] = useState({ open: false, draft: null, mode: "create" });
  const [themeOpen, setThemeOpen] = useState(false);
  const [dragDropActive, setDragDropActive] = useState(false);
  const [persistWarning, setPersistWarning] = useState("");
  const importJsonRef = useRef(null);
  const [bookView, setBookView] = useState({ zoom: 1, panX: 0, panY: 0 });
  const [topbarMenuOpen, setTopbarMenuOpen] = useState(false);
  const [helpOpen, setHelpOpen] = useState(false);
  const topbarActionsRef = useRef(null);
  const [pageMetrics, setPageMetrics] = useState({});
  const [autoLayoutMode, setAutoLayoutMode] = useState("hero");
  const autoLayoutModeRef = useRef("hero");
  const [savedProjects, setSavedProjects] = useState(() => loadProjectsIndex());
  const [currentProjectId, setCurrentProjectId] = useState(null);
  const [projectDialog, setProjectDialog] = useState({ open: false, mode: "save", name: "" });
  const [savedThemes, setSavedThemes] = useState(() => loadThemesIndex());
  const [currentThemeId, setCurrentThemeId] = useState(null);
  const [themeDialog, setThemeDialog] = useState({ open: false, mode: "save", name: "" });
  const [printProgress, setPrintProgress] = useState({ active: false, current: 0, total: 0, message: "" });

  useEffect(() => {
    autoLayoutModeRef.current = autoLayoutMode;
  }, [autoLayoutMode]);

  useEffect(() => {
    function onDocPointerDown(e) {
      if (!topbarActionsRef.current?.contains(e.target)) {
        setTopbarMenuOpen(false);
        setHelpOpen(false);
      }
    }
    document.addEventListener("pointerdown", onDocPointerDown);
    return () => document.removeEventListener("pointerdown", onDocPointerDown);
  }, []);

  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(sanitizeStateForStorage(state)));
      if (persistWarning) setPersistWarning("");
    } catch (err) {
      const isQuota = err instanceof DOMException && err.name === "QuotaExceededError";
      setPersistWarning(
        isQuota
          ? "Salvataggio locale pieno: layout salvato parzialmente. Riduci testi/stato o esporta JSON."
          : "Errore nel salvataggio locale.",
      );
    }
  }, [state]);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      const worksWithImages = await Promise.all(
        (state.works || []).map(async (work) => {
          if (work.imageUrl || !work.imageId) return work;
          try {
            const imageUrl = await idbGetImage(work.imageId);
            return imageUrl ? { ...work, imageUrl } : work;
          } catch {
            return work;
          }
        }),
      );
      if (cancelled) return;
      const changed = worksWithImages.some((w, i) => w.imageUrl !== state.works[i]?.imageUrl);
      if (changed) {
        setState((prev) => ({ ...prev, works: worksWithImages }));
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [state.works]);

  useEffect(() => {
    if (state.specialPages?.insideFront && state.specialPages?.insideBack) return;
    setState((prev) => {
      if (prev.specialPages?.insideFront && prev.specialPages?.insideBack) return prev;
      const margins = prev.theme?.pageMargins || createDefaultState().theme.pageMargins;
      return {
        ...prev,
        specialPages: {
          insideFront: prev.specialPages?.insideFront || createInsideFrontCoverPage(margins),
          insideBack: prev.specialPages?.insideBack || createInsideBackCoverPage(margins),
        },
      };
    });
  }, [state.specialPages]);

  useEffect(() => {
    const backCover = state.pages?.find((p) => p.type === "cover-back");
    if (!backCover) return;
    const hasEditionCode = (backCover.textBlocks || []).some((t) => String(t.text || "").toLowerCase().includes("codice edizione"));
    if (hasEditionCode) return;
    setState((prev) => ({
      ...prev,
      pages: (prev.pages || []).map((p) =>
        p.type !== "cover-back"
          ? p
          : {
              ...p,
              textBlocks: [
                ...(p.textBlocks || []),
                {
                  ...createTextBlock("Edizione 1/200\nCodice edizione: CAT-2026-001"),
                  x: 30,
                  y: 204,
                  w: 300,
                  h: 42,
                  fontSize: 11,
                  fontWeight: 500,
                  align: "left",
                },
              ],
            },
      ),
    }));
  }, [state.pages]);

  const editablePages = state.pages;
  const editableSpecialPages = [state.specialPages?.insideFront, state.specialPages?.insideBack].filter(Boolean);
  const currentFormat = getPageFormat(state.pageFormat);
  const frontCover = editablePages[0];
  const backCover = editablePages[editablePages.length - 1];
  const innerPages = editablePages.slice(1, -1).map((page) => ({
    ...page,
    title: page.title || "Pagina",
  }));
  const insideFrontBlank = state.specialPages?.insideFront || createInsideFrontCoverPage(state.theme.pageMargins);
  const insideBackBlank = state.specialPages?.insideBack || createInsideBackCoverPage(state.theme.pageMargins);
  const workPageMap = buildWorkFirstPageMapForCatalog(frontCover, insideFrontBlank, innerPages);
  const summaryPages = summaryPagesFromWorks(
    state.works,
    state.theme.pageMargins,
    workPageMap,
    state.summaryPageEdits,
    state.theme.defaultPageBgColor,
    state.theme.defaultElementBorderPct,
    state.theme.defaultPageNumberColor,
    state.pageFormat,
  );
  const allEditablePages = [...editablePages, ...editableSpecialPages, ...summaryPages];
  const renderPagesBase = [frontCover, insideFrontBlank, ...innerPages, ...summaryPages, insideBackBlank, backCover];
  let autoPageCounter = 0;
  const renderPages = renderPagesBase.map((page, idx) => {
    if (!page) return page;
    const isCover = idx === 0 || idx === renderPagesBase.length - 1;
    const isSpacer = page.type === "spacer";
    const isInsideCover = page.type === "inside-front-cover" || page.type === "inside-back-cover";
    if (!isCover && !isSpacer && !isInsideCover) autoPageCounter += 1;
    return {
      ...page,
      computedPageNumber: isCover || isSpacer || isInsideCover ? null : autoPageCounter,
    };
  });

  const spreads = useMemo(
    () => (state.bookViewMode === "technical" ? buildTechnicalSpreads(renderPages) : buildBookSpreads(renderPages)),
    [renderPages, state.bookViewMode],
  );

  const currentSpreadIndex = Math.min(state.currentSpread, Math.max(spreads.length - 1, 0));
  const currentSpread = spreads[currentSpreadIndex] || [null, null];

  useEffect(() => {
    if (state.currentSpread !== currentSpreadIndex) {
      setState((prev) => ({ ...prev, currentSpread: currentSpreadIndex }));
    }
  }, [currentSpreadIndex, state.currentSpread]);

  useEffect(() => {
    const prevFormat = prevPageFormatRef.current;
    if (!prevFormat || prevFormat === state.pageFormat) return;
    if (skipNextPageFormatAdjustRef.current) {
      skipNextPageFormatAdjustRef.current = false;
      prevPageFormatRef.current = state.pageFormat;
      return;
    }
    setState((prev) => {
      const fromFmt = prevPageFormatRef.current || prev.pageFormat;
      const toFmt = prev.pageFormat;
      const boundMode = prev.layoutAssist?.boundMode || "margins";
      const canonicalMetric = Object.values(pageMetrics || {}).find((pm) => pm?.pageWidth && pm?.pageHeight) || null;
      const fitPageToNewFormat = (p) => {
        const scaled = scalePageLayoutForFormat(p, fromFmt, toFmt);
        return canonicalMetric
          ? fitPageElementsToEstimatedRenderedBoundsSoft(scaled, fromFmt, toFmt, canonicalMetric, boundMode)
          : scaled;
      };
      return {
        ...prev,
        pages: (prev.pages || []).map((p) => fitPageToNewFormat(p)),
        specialPages: {
          insideFront: prev.specialPages?.insideFront
            ? fitPageToNewFormat(prev.specialPages.insideFront)
            : prev.specialPages?.insideFront,
          insideBack: prev.specialPages?.insideBack
            ? fitPageToNewFormat(prev.specialPages.insideBack)
            : prev.specialPages?.insideBack,
        },
        summaryPageEdits: Object.fromEntries(
          Object.entries(prev.summaryPageEdits || {}).map(([id, p]) => [id, fitPageToNewFormat(p)]),
        ),
      };
    });
    prevPageFormatRef.current = state.pageFormat;
  }, [state.pageFormat]);

  function patchState(updater) {
    setState((prev) => (typeof updater === "function" ? updater(prev) : { ...prev, ...updater }));
  }

  function patchTheme(patch) {
    patchState((prev) => {
      const nextTheme = { ...prev.theme, ...patch };
      const hasBgPatch = Object.prototype.hasOwnProperty.call(patch, "defaultPageBgColor");
      const hasBorderPatch = Object.prototype.hasOwnProperty.call(patch, "defaultElementBorderPct");
      const hasBorderColorPatch = Object.prototype.hasOwnProperty.call(patch, "defaultElementBorderColor");
      const hasPageNumberColorPatch = Object.prototype.hasOwnProperty.call(patch, "defaultPageNumberColor");
      const hasTextBgPatch = Object.prototype.hasOwnProperty.call(patch, "defaultTextBgColor");
      if (!hasBgPatch && !hasBorderPatch && !hasBorderColorPatch && !hasPageNumberColorPatch && !hasTextBgPatch) {
        return { ...prev, theme: nextTheme };
      }
      const nextBg = patch.defaultPageBgColor || prev.theme?.defaultPageBgColor || "#ffffff";
      const nextBorderPct = Number.isFinite(Number(patch.defaultElementBorderPct))
        ? Number(patch.defaultElementBorderPct)
        : prev.theme?.defaultElementBorderPct ?? 3;
      const nextBorderColor = patch.defaultElementBorderColor || prev.theme?.defaultElementBorderColor || "#ffffff";
      const nextPageNumberColor = patch.defaultPageNumberColor || prev.theme?.defaultPageNumberColor || "#6b614f";
      const nextTextBgColor = patch.defaultTextBgColor || prev.theme?.defaultTextBgColor || "rgba(255, 255, 255, 0.42)";
      const patchPageElementsBorders = (page) => ({
        ...page,
        bgColor: hasBgPatch ? nextBg : page.bgColor,
        pageNumberColor: hasPageNumberColorPatch ? nextPageNumberColor : page.pageNumberColor,
        placements: (page.placements || []).map((pl) => ({
          ...pl,
          borderWidthPct: hasBorderPatch ? nextBorderPct : pl.borderWidthPct,
          borderColor: hasBorderColorPatch ? nextBorderColor : pl.borderColor,
          captionBorderWidthPct: hasBorderPatch ? nextBorderPct : pl.captionBorderWidthPct,
          captionBorderColor: hasBorderColorPatch ? nextBorderColor : pl.captionBorderColor,
        })),
        textBlocks: (page.textBlocks || []).map((tx) => ({
          ...tx,
          borderWidthPct: hasBorderPatch ? nextBorderPct : tx.borderWidthPct,
          borderColor: hasBorderColorPatch ? nextBorderColor : tx.borderColor,
          bgColor: hasTextBgPatch ? nextTextBgColor : tx.bgColor,
        })),
      });
      return {
        ...prev,
        theme: nextTheme,
        pages: (prev.pages || []).map(patchPageElementsBorders),
        specialPages: {
          insideFront: prev.specialPages?.insideFront ? patchPageElementsBorders(prev.specialPages.insideFront) : prev.specialPages?.insideFront,
          insideBack: prev.specialPages?.insideBack ? patchPageElementsBorders(prev.specialPages.insideBack) : prev.specialPages?.insideBack,
        },
        summaryPageEdits: Object.fromEntries(
          Object.entries(prev.summaryPageEdits || {}).map(([id, page]) => [id, patchPageElementsBorders(page)]),
        ),
      };
    });
  }

  function updateGlobalMargins(patch) {
    patchState((prev) => {
      const nextMargins = { ...(prev.theme?.pageMargins || {}), ...patch };
      return {
        ...prev,
        theme: { ...prev.theme, pageMargins: nextMargins },
        pages: (prev.pages || []).map((page) => ({ ...page, margins: { ...nextMargins } })),
        specialPages: {
          insideFront: prev.specialPages?.insideFront ? { ...prev.specialPages.insideFront, margins: { ...nextMargins } } : prev.specialPages?.insideFront,
          insideBack: prev.specialPages?.insideBack ? { ...prev.specialPages.insideBack, margins: { ...nextMargins } } : prev.specialPages?.insideBack,
        },
      };
    });
  }

  function patchPage(pageId, updater) {
    patchState((prev) => ({
      ...prev,
      pages: prev.pages.map((page) => {
        if (page.id !== pageId) return page;
        const next = typeof updater === "function" ? updater(page) : { ...page, ...updater };
        return next;
      }),
      summaryPageEdits:
        String(pageId).startsWith("summary_")
          ? {
              ...(prev.summaryPageEdits || {}),
              [pageId]: (() => {
                const generatedPage =
                  summaryPages.find((p) => p.id === pageId) || { id: pageId, textBlocks: [], placements: [], margins: prev.theme?.pageMargins };
                return typeof updater === "function" ? updater(generatedPage) : { ...generatedPage, ...updater };
              })(),
            }
          : prev.summaryPageEdits,
      specialPages: {
        insideFront:
          prev.specialPages?.insideFront?.id === pageId
            ? (typeof updater === "function" ? updater(prev.specialPages.insideFront) : { ...prev.specialPages.insideFront, ...updater })
            : prev.specialPages?.insideFront,
        insideBack:
          prev.specialPages?.insideBack?.id === pageId
            ? (typeof updater === "function" ? updater(prev.specialPages.insideBack) : { ...prev.specialPages.insideBack, ...updater })
            : prev.specialPages?.insideBack,
      },
    }));
  }

  function openCreateWork(prefill) {
    setWorkEditor({ open: true, draft: cloneWorkDraft(prefill), mode: "create" });
  }

  function openEditWork(work) {
    setWorkEditor({ open: true, draft: cloneWorkDraft(work), mode: "edit" });
  }

  async function saveWork(draft) {
    const incoming = normalizeWorkData({ ...draft });
    const imageDimensions = await readImageDimensions(incoming.imageUrl);
    const existing = state.works.find((w) => w.id === incoming.id);
    if (incoming._imageFile && incoming.imageUrl) {
      const nextImageId = incoming.imageId || uid("img");
      await idbPutImage(nextImageId, incoming.imageUrl);
      if (existing?.imageId && existing.imageId !== nextImageId) {
        await idbDeleteImage(existing.imageId);
      }
      incoming.imageId = nextImageId;
    }
    delete incoming._imageFile;

    patchState((prev) => {
      const exists = prev.works.some((w) => w.id === incoming.id);
      const works = exists ? prev.works.map((w) => (w.id === incoming.id ? incoming : w)) : [...prev.works, incoming];
      if (exists) {
        return { ...prev, works, selectedWorkId: incoming.id };
      }
      let autoPage = createAutoWorkPage(incoming, prev.pageFormat, imageDimensions || null);
      autoPage = applyThemeAutoDefaultsToPage(autoPage, prev.theme);
      autoPage.margins = { ...(prev.theme?.pageMargins || autoPage.margins) };
      const nextPages = [...prev.pages];
      nextPages.splice(nextPages.length - 1, 0, autoPage);
      return {
        ...prev,
        works,
        pages: nextPages,
        selectedWorkId: incoming.id,
        activePageId: autoPage.id,
        selectedElement: { pageId: autoPage.id, kind: "placement", elementId: autoPage.placements[0].id },
      };
    });
    setWorkEditor({ open: false, draft: null, mode: "create" });
  }

  async function deleteWork(workId) {
    const work = state.works.find((w) => w.id === workId);
    if (work?.imageId) {
      try {
        await idbDeleteImage(work.imageId);
      } catch {
        // ignore
      }
    }
    patchState((prev) => ({
      ...prev,
      works: prev.works.filter((w) => w.id !== workId),
      selectedWorkId: prev.selectedWorkId === workId ? null : prev.selectedWorkId,
      pages: prev.pages.map((page) => ({
        ...page,
        placements: (page.placements || []).filter((p) => p.workId !== workId),
      })),
      selectedElement: null,
    }));
  }

  async function handleDropFiles(files) {
    const imageFiles = [...files].filter((f) => f.type.startsWith("image/"));
    if (!imageFiles.length) return;

    if (imageFiles.length === 1) {
      const imageFile = imageFiles[0];
      const imageUrl = await readImageFile(imageFile);
      const titleFromName = imageFile.name.replace(/\.[^.]+$/, "");
      openCreateWork({ title: titleFromName, imageUrl, _imageFile: imageFile, dpi: DEFAULT_WORK_DPI });
      return;
    }

    const importedWorks = [];
    for (const file of imageFiles) {
      const imageUrl = await readImageFile(file);
      const imageDimensions = await readImageDimensions(imageUrl);
      const work = {
        ...normalizeWorkData(cloneWorkDraft()),
        title: file.name.replace(/\.[^.]+$/, ""),
        imageUrl,
        imageId: uid("img"),
      };
      try {
        await idbPutImage(work.imageId, imageUrl);
      } catch {
        // fallback: keep in-memory imageUrl even if IndexedDB fails
      }
      importedWorks.push({ work, imageDimensions });
    }
    if (!importedWorks.length) return;

    patchState((prev) => {
      const autoPages = importedWorks.map(({ work, imageDimensions }) => {
        let page = createAutoWorkPage(work, prev.pageFormat, imageDimensions);
        page = applyThemeAutoDefaultsToPage(page, prev.theme);
        page.margins = { ...(prev.theme?.pageMargins || page.margins) };
        return page;
      });
      const nextPages = [...prev.pages];
      nextPages.splice(nextPages.length - 1, 0, ...autoPages);
      const importedWorkItems = importedWorks.map((it) => it.work);
      const lastWork = importedWorkItems[importedWorkItems.length - 1];
      const firstPage = autoPages[0];
      return {
        ...prev,
        works: [...prev.works, ...importedWorkItems],
        pages: nextPages,
        selectedWorkId: lastWork.id,
        activePageId: firstPage?.id || prev.activePageId,
        selectedElement: firstPage?.placements?.[0]
          ? { pageId: firstPage.id, kind: "placement", elementId: firstPage.placements[0].id }
          : prev.selectedElement,
      };
    });
  }

  async function onFrameDrop(e) {
    e.preventDefault();
    setDragDropActive(false);
    if (e.dataTransfer?.files?.length) {
      await handleDropFiles(e.dataTransfer.files);
    }
  }

  function onFrameDragOver(e) {
    e.preventDefault();
    if (!dragDropActive) setDragDropActive(true);
  }

  function onFrameDragLeave(e) {
    if (e.currentTarget.contains(e.relatedTarget)) return;
    setDragDropActive(false);
  }

  function selectPage(page) {
    if (!page) return;
    const isEditable = allEditablePages.some((p) => p.id === page.id);
    patchState((prev) => ({
      ...prev,
      activePageId: isEditable ? page.id : prev.activePageId,
      selectedElement: null,
    }));
  }

  function addInnerPage() {
    const newPage = createPage("page", `Pagina ${Math.max(state.pages.length - 1, 1)}`, state.theme.pageMargins);
    patchState((prev) => {
      const next = [...prev.pages];
      next.splice(next.length - 1, 0, newPage);
      return { ...prev, pages: next, activePageId: newPage.id };
    });
  }

  function removeInnerPage(pageId) {
    patchState((prev) => {
      if (prev.pages.length <= 3) return prev;
      const target = prev.pages.find((p) => p.id === pageId);
      if (!target || target.type !== "page") return prev;
      const pages = prev.pages.filter((p) => p.id !== pageId);
      return {
        ...prev,
        pages,
        activePageId: pages[1]?.id || pages[0]?.id || null,
        selectedElement: null,
      };
    });
  }

  const activeEditablePage = allEditablePages.find((p) => p.id === state.activePageId) || state.pages[1];
  const selectedWork = state.works.find((w) => w.id === state.selectedWorkId) || null;

  function addTextToActivePage() {
    if (!activeEditablePage || activeEditablePage.type === "summary") return;
    const bounds = getActivePageLayoutBounds(activeEditablePage);
    const block = createTextBlock("Testo editabile");
    const targetW = Math.max(120, Math.round(bounds.width * 0.9));
    const borderPx = borderPxFromPercent(targetW, block.borderWidthPct ?? 5, 0, Math.max(48, targetW));
    const minH = Math.ceil((block.fontSize || state.theme.bodyFontSize || 16) * 1.25 + borderPx * 2 + 8);
    block.bgColor = state.theme?.defaultTextBgColor || block.bgColor || "rgba(255, 255, 255, 0.42)";
    block.w = targetW;
    block.h = Math.max(block.h || 0, minH);
    block.x = Math.max(0, Math.round((bounds.width - block.w) / 2));
    block.y = Math.max(0, Math.round(bounds.height * 0.08));
    patchPage(activeEditablePage.id, (page) => ({ ...page, textBlocks: [...page.textBlocks, block] }));
    patchState((prev) => ({ ...prev, selectedElement: { pageId: activeEditablePage.id, kind: "text", elementId: block.id } }));
  }

  async function addSelectedWorkToActivePage() {
    if (!activeEditablePage || activeEditablePage.type === "summary" || !selectedWork) return;
    const bounds = getActivePageLayoutBounds(activeEditablePage);
    const imageDimensions = await readImageDimensions(selectedWork.imageUrl);
    const captionH = (state.theme?.autoShowCaptionDefault ?? true) ? 46 : 28;
    const captionGap = 8;
    const maxW = Math.max(80, Math.round(bounds.width * 0.7));
    const maxH = Math.max(80, Math.round(bounds.height * 0.7) - captionH - captionGap);
    const fallbackAspect =
      Number(imageDimensions?.width) > 0 && Number(imageDimensions?.height) > 0
        ? Number(imageDimensions.width) / Math.max(1, Number(imageDimensions.height))
        : 1;
    const fitted = fitImageMmToBox(selectedWork, imageDimensions, maxW, maxH, fallbackAspect);
    const x = Math.max(0, Math.round((bounds.width - fitted.w) / 2));
    const y = Math.max(0, Math.round((bounds.height - fitted.h - captionH - captionGap) / 2));
    const placement = applyThemeDefaultsToPlacement(createPlacementForWork(selectedWork.id, x, y), state.theme);
    placement.w = fitted.w;
    placement.h = fitted.h;
    placement.captionX = Math.max(0, Math.round((bounds.width - Math.max(160, fitted.w)) / 2));
    placement.captionY = y + fitted.h + captionGap;
    placement.captionW = Math.min(bounds.width, Math.max(160, fitted.w));
    placement.captionH = captionH;
    patchPage(activeEditablePage.id, (page) => ({
      ...page,
      placements: [...page.placements, { ...placement, zIndex: nextPlacementLayerZ(page) }],
    }));
    patchState((prev) => ({ ...prev, selectedElement: { pageId: activeEditablePage.id, kind: "placement", elementId: placement.id } }));
  }

  function getAutoLayoutChunkSize(mode) {
    if (mode === "two-cols") return 2;
    if (mode === "grid4") return 4;
    return 1;
  }

  function buildAutoLayoutPageForWorks(worksChunk, pageFormat, margins, dimMap, measuredBounds, mode, layoutAssistCfg, themeCfg) {
    const page = createPage("page", worksChunk.length === 1 ? workLabel(worksChunk[0]) : "Pagina opere", margins);
    page.bgColor = "#fbfbfb";
    const showCaption = themeCfg?.autoShowCaptionDefault ?? true;
    const area = getPageContentBounds(page, pageFormat);
    const base = { x: 0, y: 0, w: Math.max(80, area.width), h: Math.max(100, area.height) };
    const gap = 8;
    const count = Math.max(1, worksChunk.length);

    let slots = [];
    if (mode === "hero" || count === 1) {
      slots = [{ x: base.x, y: base.y, w: base.w, h: base.h }];
    } else if (mode === "two-cols") {
      const portrait = base.h >= base.w;
      if (portrait) {
        const h1 = Math.floor((base.h - gap) / 2);
        const h2 = base.h - h1 - gap;
        slots = [
          { x: base.x, y: base.y, w: base.w, h: h1 },
          { x: base.x, y: base.y + h1 + gap, w: base.w, h: h2 },
        ];
      } else {
        const w1 = Math.floor((base.w - gap) / 2);
        const w2 = base.w - w1 - gap;
        slots = [
          { x: base.x, y: base.y, w: w1, h: base.h },
          { x: base.x + w1 + gap, y: base.y, w: w2, h: base.h },
        ];
      }
    } else {
      const w1 = Math.floor((base.w - gap) / 2);
      const w2 = base.w - w1 - gap;
      const h1 = Math.floor((base.h - gap) / 2);
      const h2 = base.h - h1 - gap;
      slots = [
        { x: base.x, y: base.y, w: w1, h: h1 },
        { x: base.x + w1 + gap, y: base.y, w: w2, h: h1 },
        { x: base.x, y: base.y + h1 + gap, w: w1, h: h2 },
        { x: base.x + w1 + gap, y: base.y + h1 + gap, w: w2, h: h2 },
      ];
    }

    page.placements = worksChunk.map((work, idx) => {
      const slot = slots[idx] || slots[slots.length - 1] || { x: 0, y: 0, w: base.w, h: base.h };
      const captionH = showCaption ? (mode === "grid4" ? 34 : 46) : 0;
      const captionGap = showCaption ? 6 : 0;
      const slotW = Math.max(40, slot.w);
      const imgH = Math.max(40, slot.h - captionH - captionGap);
      const imageDimensions = dimMap?.get(work.id) || null;
      const fallbackAspect =
        Number(imageDimensions?.width) > 0 && Number(imageDimensions?.height) > 0
          ? Number(imageDimensions.width) / Math.max(1, Number(imageDimensions.height))
          : slotW / Math.max(1, imgH);
      const fitted = fitImageMmToBox(work, imageDimensions, slotW, imgH, fallbackAspect);
      const imageX = slot.x + Math.round((slotW - fitted.w) / 2);
      const imageY = slot.y + Math.round((imgH - fitted.h) / 2);
      const p = createPlacementForWork(work.id, slot.x, slot.y);
      p.x = imageX;
      p.y = imageY;
      p.w = fitted.w;
      p.h = fitted.h;
      p.showCaption = showCaption;
      p.captionX = slot.x;
      p.captionY = slot.y + imgH + captionGap;
      p.captionW = slotW;
      p.captionH = showCaption ? captionH : 28;
      p.captionOverride = [workLabel(work), [work.author, work.year].filter(Boolean).join(", ")].filter(Boolean).join("\n");
      p.zIndex = 100 + idx * 10;
      return p;
    });

    return normalizePageElementsToBounds(applyThemeAutoDefaultsToPage(page, themeCfg || state.theme), pageFormat, "margins");
  }

  async function autoGeneratePagesFromCatalog() {
    const works = state.works || [];
    if (!works.length) return;
    const effectiveAutoLayoutMode = autoLayoutModeRef.current || autoLayoutMode || "hero";
    const dimensions = await Promise.all(
      works.map(async (w) => {
        const dims = await readImageDimensions(w.imageUrl);
        return [w.id, dims];
      }),
    );
    const dimMap = new Map(dimensions);
    const metricValues = Object.values(pageMetrics || {});
    const measuredBounds = metricValues.find((m) => m?.width && m?.height) || null;

    patchState((prev) => {
      const boundMode = prev.layoutAssist?.boundMode || "margins";
      const defaultBg = prev.theme?.defaultPageBgColor || "#ffffff";
      const frontBase = prev.pages[0] || createDefaultState().pages[0];
      const backBase = prev.pages[prev.pages.length - 1] || createDefaultState().pages.at(-1);
      const front = { ...frontBase, placements: [], bgColor: defaultBg };
      const back = { ...backBase, placements: [], bgColor: defaultBg };
      const existingPreface = (prev.pages || []).find((p) => p?.title === "Prefazione" && p.type === "page");
      let prefacePage = createPrefacePage(
        prev.theme?.pageMargins,
        defaultBg,
        prev.theme?.defaultElementBorderPct ?? 3,
        prev.pageFormat,
        prev.theme,
      );
      if (existingPreface) {
        prefacePage = {
          ...prefacePage,
          ...existingPreface,
          id: existingPreface.id,
          margins: { ...(prev.theme?.pageMargins || prefacePage.margins) },
          placements: [], // la prefazione resta testuale nel flusso auto
          textBlocks: existingPreface.textBlocks || prefacePage.textBlocks,
        };
      }
      const chunkSize = getAutoLayoutChunkSize(effectiveAutoLayoutMode);
      const chunks = [];
      for (let i = 0; i < prev.works.length; i += chunkSize) chunks.push(prev.works.slice(i, i + chunkSize));
      const autoPages = chunks.map((chunk) =>
        buildAutoLayoutPageForWorks(
          chunk,
          prev.pageFormat,
          prev.theme?.pageMargins,
          dimMap,
          undefined,
          effectiveAutoLayoutMode,
          prev.layoutAssist,
          prev.theme,
        ),
      );
      const frontPage = { ...front, margins: { ...(prev.theme?.pageMargins || front.margins) } };
      const backPage = { ...back, margins: { ...(prev.theme?.pageMargins || back.margins) } };
      const normalizedPreface = existingPreface
        ? { ...prefacePage, margins: { ...(prev.theme?.pageMargins || prefacePage.margins) } }
        : normalizePageElementsToBounds(prefacePage, prev.pageFormat, "margins");
      const normalizedAutoPages = autoPages.length ? autoPages : [createPage("page", "Pagina 1", prev.theme?.pageMargins)];
      const pages = [frontPage, normalizedPreface, ...normalizedAutoPages, backPage];

      const nextSpecialPages = {
        insideFront: prev.specialPages?.insideFront
          ? { ...prev.specialPages.insideFront, placements: [] }
          : prev.specialPages?.insideFront,
        insideBack: prev.specialPages?.insideBack
          ? { ...prev.specialPages.insideBack, placements: [] }
          : prev.specialPages?.insideBack,
      };
      return {
        ...prev,
        pages,
        specialPages: nextSpecialPages,
        activePageId: prefacePage.id,
        currentSpread: 0,
        selectedElement: prefacePage.textBlocks?.[0]
          ? { pageId: prefacePage.id, kind: "text", elementId: prefacePage.textBlocks[0].id }
          : null,
      };
    });
  }

  function getActivePageLayoutBounds(page) {
    const measured = pageMetrics?.[page.id];
    if (measured?.width && measured?.height) {
      return { width: measured.width, height: measured.height };
    }
    const fallback = getPageContentBounds(page, state.pageFormat);
    return { width: fallback.width, height: fallback.height };
  }

  async function addWorkToPageAtPosition(pageId, workId, x, y) {
    const page = allEditablePages.find((p) => p.id === pageId);
    if (!page || page.type === "summary") return;
    const work = state.works.find((w) => w.id === workId);
    const bounds = getActivePageLayoutBounds(page);
    const safeX = clampNum(Math.round(x), 0, Math.max(0, bounds.width - 1));
    const safeY = clampNum(Math.round(y), 0, Math.max(0, bounds.height - 1));
    const imageDimensions = await readImageDimensions(work?.imageUrl);
    const maxW = Math.max(40, bounds.width - safeX);
    const maxH = Math.max(40, bounds.height - safeY);
    const fallbackAspect =
      Number(imageDimensions?.width) > 0 && Number(imageDimensions?.height) > 0
        ? Number(imageDimensions.width) / Math.max(1, Number(imageDimensions.height))
        : 1;
    const fitted = fitImageMmToBox(work, imageDimensions, maxW, maxH, fallbackAspect);
    const placement = applyThemeDefaultsToPlacement(createPlacementForWork(workId, safeX, safeY), state.theme);
    placement.w = fitted.w;
    placement.h = fitted.h;
    placement.captionX = safeX;
    placement.captionY = safeY + fitted.h + 8;
    placement.captionW = Math.min(Math.max(40, bounds.width - safeX), Math.max(160, fitted.w));
    placement.captionH = 46;
    patchPage(pageId, (p) => ({ ...p, placements: [...p.placements, { ...placement, zIndex: nextPlacementLayerZ(p) }] }));
    patchState((prev) => ({
      ...prev,
      activePageId: pageId,
      selectedWorkId: workId,
      selectedElement: { pageId, kind: "placement", elementId: placement.id },
    }));
  }

  function movePlacementToPage(sourcePageId, placementId, targetPageId, x, y) {
    patchState((prev) => {
      const findEditablePageById = (id) =>
        prev.pages.find((p) => p.id === id) ||
        (prev.specialPages?.insideFront?.id === id ? prev.specialPages.insideFront : null) ||
        (prev.specialPages?.insideBack?.id === id ? prev.specialPages.insideBack : null) ||
        null;
      const sourcePage = findEditablePageById(sourcePageId);
      const targetPage = findEditablePageById(targetPageId);
      if (!sourcePage || !targetPage || targetPage.type === "summary") return prev;
      const placement = (sourcePage.placements || []).find((p) => p.id === placementId);
      if (!placement) return prev;

      const dx = x - (placement.x ?? 0);
      const dy = y - (placement.y ?? 0);
      const movedPlacement = {
        ...placement,
        x,
        y,
        captionX: (placement.captionX ?? placement.x ?? 0) + dx,
        captionY: (placement.captionY ?? ((placement.y ?? 0) + (placement.h ?? 0) + 8)) + dy,
      };

      const patchSinglePage = (page) => {
        if (!page) return page;
        if (page.id === sourcePageId && page.id === targetPageId) {
          return {
            ...page,
            placements: (page.placements || []).map((p) => (p.id === placementId ? movedPlacement : p)),
          };
        }
        if (page.id === sourcePageId) {
          return { ...page, placements: (page.placements || []).filter((p) => p.id !== placementId) };
        }
        if (page.id === targetPageId) {
          return { ...page, placements: [...(page.placements || []), movedPlacement] };
        }
        return page;
      };

      return {
        ...prev,
        pages: (prev.pages || []).map(patchSinglePage),
        specialPages: {
          insideFront: patchSinglePage(prev.specialPages?.insideFront),
          insideBack: patchSinglePage(prev.specialPages?.insideBack),
        },
        activePageId: targetPageId,
        selectedElement: { pageId: targetPageId, kind: "placement", elementId: placementId },
      };
    });
  }

  function ensurePlacementsForLayout(page, targetCount) {
    const current = [...(page.placements || [])];
    if (current.length >= targetCount) return current.slice(0, targetCount);
    const pool = state.works.filter((w) => !current.some((p) => p.workId === w.id));
    for (const work of pool) {
      current.push({
        ...applyThemeDefaultsToPlacement(createPlacementForWork(work.id, 20, 20), state.theme),
        zIndex: nextPlacementLayerZ({ placements: current }),
      });
      if (current.length >= targetCount) break;
    }
    return current;
  }

  function applyLayoutPreset(preset) {
    const page = activeEditablePage;
    if (!page || page.type === "summary") return;
    patchPage(page.id, (prevPage) => {
      const measured = getActivePageLayoutBounds(prevPage);
      const areaW = Math.max(140, measured.width);
      const areaH = Math.max(180, measured.height);
      let placements = [...(prevPage.placements || [])];
      const fitPlacementInSlot = (placement, slot, defaultCaptionH, captionGap) => {
        const showCaption = placement.showCaption !== false;
        const capH = showCaption ? defaultCaptionH : 0;
        const gap = showCaption ? captionGap : 0;
        const imageAreaH = Math.max(40, slot.h - capH - gap);
        const ratio = Math.max(0.01, Number(placement.w || 1) / Math.max(1, Number(placement.h || 1)));
        let w = Math.max(40, slot.w);
        let h = w / ratio;
        if (h > imageAreaH) {
          h = imageAreaH;
          w = h * ratio;
        }
        w = clampNum(Math.round(w), 40, Math.max(40, slot.w));
        h = clampNum(Math.round(h), 40, Math.max(40, imageAreaH));
        const x = Math.round(slot.x + (slot.w - w) / 2);
        const y = Math.round(slot.y + (imageAreaH - h) / 2);
        return {
          ...placement,
          x,
          y,
          w,
          h,
          captionX: slot.x,
          captionY: slot.y + imageAreaH + gap,
          captionW: Math.max(40, slot.w),
          captionH: showCaption ? defaultCaptionH : 28,
        };
      };

      if (preset === "hero") {
        placements = ensurePlacementsForLayout(prevPage, 1);
        if (!placements.length) return prevPage;
        const p = { ...placements[0] };
        const slot = { x: 0, y: 0, w: areaW, h: areaH };
        placements = [fitPlacementInSlot(p, slot, 50, 8)];
      }

      if (preset === "two-cols") {
        placements = ensurePlacementsForLayout(prevPage, 2);
        if (!placements.length) return prevPage;
        const gap = 12;
        const portrait = areaH >= areaW;
        if (portrait) {
          const h1 = Math.floor((areaH - gap) / 2);
          const h2 = areaH - h1 - gap;
          const slots = [
            { x: 0, y: 0, w: areaW, h: h1 },
            { x: 0, y: h1 + gap, w: areaW, h: h2 },
          ];
          placements = placements.slice(0, 2).map((p, i) => fitPlacementInSlot(p, slots[i], 48, 8));
        } else {
          const w1 = Math.floor((areaW - gap) / 2);
          const w2 = areaW - w1 - gap;
          const slots = [
            { x: 0, y: 0, w: w1, h: areaH },
            { x: w1 + gap, y: 0, w: w2, h: areaH },
          ];
          placements = placements.slice(0, 2).map((p, i) => fitPlacementInSlot(p, slots[i], 48, 8));
        }
      }

      if (preset === "grid4") {
        placements = ensurePlacementsForLayout(prevPage, 4);
        if (!placements.length) return prevPage;
        const cols = 2;
        const rows = 2;
        const gapX = 12;
        const gapY = 12;
        const w = Math.round((areaW - gapX) / cols);
        const h = Math.round((areaH - gapY) / rows);
        placements = placements.slice(0, 4).map((p, i) => {
          const col = i % cols;
          const row = Math.floor(i / cols);
          const slot = {
            x: col * (w + gapX),
            y: row * (h + gapY),
            w,
            h,
          };
          return fitPlacementInSlot(p, slot, 36, 6);
        });
      }

      return { ...prevPage, placements };
    });
  }

  function distributeActivePageElements(axis) {
    const page = activeEditablePage;
    if (!page || page.type === "summary") return;
    patchPage(page.id, (prevPage) => {
      const items = [
        ...(prevPage.placements || []).map((p) => ({ kind: "placement", id: p.id, x: p.x, y: p.y, w: p.w, h: p.h })),
        ...(prevPage.textBlocks || []).map((t) => ({ kind: "text", id: t.id, x: t.x, y: t.y, w: t.w, h: t.h })),
      ];
      if (items.length < 3) return prevPage;

      const key = axis === "x" ? "x" : "y";
      const sizeKey = axis === "x" ? "w" : "h";
      const sorted = [...items].sort((a, b) => a[key] - b[key]);
      const totalSize = sorted.reduce((sum, it) => sum + it[sizeKey], 0);
      const respectMargins = !!state.layoutAssist?.distributeRespectMargins;

      let startPos;
      let span;
      if (respectMargins) {
        const bounds = getActivePageLayoutBounds(prevPage);
        startPos = 0;
        span = axis === "x" ? bounds.width : bounds.height;
      } else {
        const first = sorted[0];
        const last = sorted[sorted.length - 1];
        startPos = first[key];
        span = last[key] + last[sizeKey] - first[key];
      }

      const gap = (span - totalSize) / (sorted.length - 1);
      if (!Number.isFinite(gap)) return prevPage;

      let cursor = startPos;
      const posMap = new Map();
      for (const it of sorted) {
        posMap.set(`${it.kind}:${it.id}`, Math.round(cursor));
        cursor += it[sizeKey] + gap;
      }

      const placements = prevPage.placements.map((p) => {
        const next = posMap.get(`placement:${p.id}`);
        if (next == null) return p;
        if (axis === "x") {
          const dx = next - p.x;
          return { ...p, x: next, captionX: (p.captionX ?? p.x) + dx };
        }
        const dy = next - p.y;
        return { ...p, y: next, captionY: (p.captionY ?? p.y + p.h + 8) + dy };
      });
      const textBlocks = prevPage.textBlocks.map((t) => {
        const next = posMap.get(`text:${t.id}`);
        if (next == null) return t;
        return axis === "x" ? { ...t, x: next } : { ...t, y: next };
      });
      return { ...prevPage, placements, textBlocks };
    });
  }

  function updateElementOnPage(pageId, kind, elementId, patch) {
    patchPage(pageId, (page) => {
      if (kind === "text") {
        return {
          ...page,
          textBlocks: page.textBlocks.map((el) => (el.id === elementId ? { ...el, ...patch } : el)),
        };
      }
      return {
        ...page,
        placements: page.placements.map((el) => {
          if (el.id !== elementId) return el;
          const next = { ...el, ...patch };
          const movesMainX = Object.prototype.hasOwnProperty.call(patch, "x") && !Object.prototype.hasOwnProperty.call(patch, "captionX");
          const movesMainY = Object.prototype.hasOwnProperty.call(patch, "y") && !Object.prototype.hasOwnProperty.call(patch, "captionY");
          if (movesMainX) {
            const dx = (patch.x ?? el.x ?? 0) - (el.x ?? 0);
            next.captionX = (el.captionX ?? el.x ?? 0) + dx;
          }
          if (movesMainY) {
            const dy = (patch.y ?? el.y ?? 0) - (el.y ?? 0);
            next.captionY = (el.captionY ?? ((el.y ?? 0) + (el.h ?? 0) + 8)) + dy;
          }
          return next;
        }),
      };
    });
  }

  function deleteSelectedElement() {
    const sel = state.selectedElement;
    if (!sel) return;
    patchPage(sel.pageId, (page) => ({
      ...page,
      textBlocks: sel.kind === "text" ? page.textBlocks.filter((el) => el.id !== sel.elementId) : page.textBlocks,
      placements: sel.kind === "placement" ? page.placements.filter((el) => el.id !== sel.elementId) : page.placements,
    }));
    patchState((prev) => ({ ...prev, selectedElement: null }));
  }

  function deleteElementDirect(sel) {
    patchPage(sel.pageId, (page) => ({
      ...page,
      textBlocks: sel.kind === "text" ? page.textBlocks.filter((el) => el.id !== sel.elementId) : page.textBlocks,
      placements: sel.kind === "placement" ? page.placements.filter((el) => el.id !== sel.elementId) : page.placements,
    }));
    patchState((prev) => {
      const isSame =
        prev.selectedElement?.pageId === sel.pageId &&
        prev.selectedElement?.kind === sel.kind &&
        prev.selectedElement?.elementId === sel.elementId;
      return isSame ? { ...prev, selectedElement: null } : prev;
    });
  }

  function normalizeActivePageNow() {
    if (!activeEditablePage) return;
    const measured = pageMetrics?.[activeEditablePage.id];
    if (measured?.width && measured?.height) {
      const box = getRenderedConstraintBox(
        activeEditablePage,
        measured.width,
        measured.height,
        state.layoutAssist?.boundMode || "margins",
      );
      patchState((prev) => ({
        ...prev,
        pages: prev.pages.map((p) =>
          p.id === activeEditablePage.id
            ? {
                ...p,
                placements: (p.placements || []).map((pl) => normalizePlacementToBounds(pl, box)),
                textBlocks: (p.textBlocks || []).map((t) => normalizeTextBlockToBounds(t, box)),
              }
            : p,
        ),
      }));
      return;
    }
    patchState((prev) => ({
      ...prev,
      pages: prev.pages.map((p) =>
        p.id === activeEditablePage.id
          ? normalizePageElementsToBounds(p, prev.pageFormat, prev.layoutAssist?.boundMode || "margins")
          : p,
      ),
    }));
  }

  const selectedElementData = (() => {
    const sel = state.selectedElement;
    if (!sel) return null;
    const page = allEditablePages.find((p) => p.id === sel.pageId);
    if (!page) return null;
    return sel.kind === "text"
      ? page.textBlocks.find((t) => t.id === sel.elementId)
      : page.placements.find((p) => p.id === sel.elementId);
  })();

  const selectedPlacementWork = (() => {
    const sel = state.selectedElement;
    if (!sel || sel.kind !== "placement" || !selectedElementData) return null;
    return state.works.find((w) => w.id === selectedElementData.workId) || null;
  })();

  const effectiveSelectedWork = selectedPlacementWork || selectedWork;

  const activePageWorks = (() => {
    if (!activeEditablePage) return [];
    const seen = new Set();
    return (activeEditablePage.placements || [])
      .map((pl) => {
        const work = state.works.find((w) => w.id === pl.workId);
        if (!work || seen.has(work.id)) return null;
        seen.add(work.id);
        return { placementId: pl.id, work };
      })
      .filter(Boolean);
  })();

  function exportCatalogJson() {
    const baseName = slugifyFileBaseName(state.projectTitle, "catalogo-opere");
    const payload = {
      exportedAt: new Date().toISOString(),
      version: 1,
      catalog: sanitizeStateForExport(state),
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${baseName}-${new Date().toISOString().slice(0, 10)}.json`;
    a.click();
    URL.revokeObjectURL(url);
  }

  async function importCatalogJson(file) {
    const text = await file.text();
    const parsed = JSON.parse(text);
    const incoming = parsed.catalog || parsed;
    const importedWorks = await Promise.all(
      (incoming.works || []).map(async (work) => {
        const normalized = normalizeWorkData(work);
        if (!normalized.imageUrl) return normalized;
        const imageId = normalized.imageId || uid("img");
        try {
          await idbPutImage(imageId, normalized.imageUrl);
        } catch {
          // keep imageUrl in memory even if IDB fails
        }
        return { ...normalized, imageId };
      }),
    );
    const base = createDefaultState();
    skipNextPageFormatAdjustRef.current = true;
    setState({
      ...base,
      ...incoming,
      works: importedWorks,
      selectedElement: null,
      currentSpread: 0,
    });
  }

  async function onImportJsonInput(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      await importCatalogJson(file);
    } catch (err) {
      window.alert(`Import JSON fallito: ${err?.message || "file non valido"}`);
    } finally {
      e.target.value = "";
    }
  }

  function persistProjectByName(name, projectIdOverride = null) {
    const trimmed = String(name || "").trim();
    if (!trimmed) return;
    const existing = savedProjects.find((p) => p.id === (projectIdOverride || currentProjectId));
    const now = new Date().toISOString();
    const projectId = projectIdOverride || existing?.id || currentProjectId || uid("prj");
    try {
      localStorage.setItem(projectStorageKey(projectId), JSON.stringify(sanitizeStateForStorage(state)));
      const nextList = [
        ...savedProjects.filter((p) => p.id !== projectId && p.name !== trimmed),
        { id: projectId, name: trimmed, updatedAt: now },
      ].sort((a, b) => String(b.updatedAt).localeCompare(String(a.updatedAt)));
      saveProjectsIndex(nextList);
      setSavedProjects(nextList);
      setCurrentProjectId(projectId);
      setProjectDialog({ open: false, mode: "save", name: "" });
    } catch (err) {
      window.alert(`Salvataggio progetto fallito: ${err?.message || "errore locale"}`);
    }
  }

  function openSaveProjectDialog(mode = "saveAs") {
    const existing = savedProjects.find((p) => p.id === currentProjectId);
    setProjectDialog({
      open: true,
      mode,
      name: existing?.name || `Progetto ${new Date().toLocaleString("it-IT")}`,
    });
  }

  function saveProjectQuick() {
    const existing = savedProjects.find((p) => p.id === currentProjectId);
    if (existing) {
      persistProjectByName(existing.name, existing.id);
      return;
    }
    openSaveProjectDialog("save");
  }

  function renameCurrentProject() {
    const existing = savedProjects.find((p) => p.id === currentProjectId);
    if (!existing) return;
    setProjectDialog({ open: true, mode: "rename", name: existing.name });
  }

  function persistThemeByName(name, themeIdOverride = null) {
    const trimmed = String(name || "").trim();
    if (!trimmed) return;
    const existing = savedThemes.find((t) => t.id === (themeIdOverride || currentThemeId));
    const now = new Date().toISOString();
    const themeId = themeIdOverride || existing?.id || currentThemeId || uid("thm");
    try {
      localStorage.setItem(themeStorageKey(themeId), JSON.stringify(sanitizeThemeForStorage(state.theme)));
      const nextList = [
        ...savedThemes.filter((t) => t.id !== themeId && t.name !== trimmed),
        { id: themeId, name: trimmed, updatedAt: now },
      ].sort((a, b) => String(b.updatedAt).localeCompare(String(a.updatedAt)));
      saveThemesIndex(nextList);
      setSavedThemes(nextList);
      setCurrentThemeId(themeId);
      setThemeDialog({ open: false, mode: "save", name: "" });
    } catch (err) {
      window.alert(`Salvataggio tema fallito: ${err?.message || "errore locale"}`);
    }
  }

  function openSaveThemeDialog(mode = "saveAs") {
    const existing = savedThemes.find((t) => t.id === currentThemeId);
    setThemeDialog({
      open: true,
      mode,
      name: existing?.name || `Tema ${new Date().toLocaleString("it-IT")}`,
    });
  }

  function saveThemeQuick() {
    const existing = savedThemes.find((t) => t.id === currentThemeId);
    if (existing) {
      persistThemeByName(existing.name, existing.id);
      return;
    }
    openSaveThemeDialog("save");
  }

  function renameCurrentTheme() {
    const existing = savedThemes.find((t) => t.id === currentThemeId);
    if (!existing) return;
    setThemeDialog({ open: true, mode: "rename", name: existing.name });
  }

  function applySavedTheme(themePayload) {
    const nextTheme = sanitizeThemeForStorage(themePayload);
    patchTheme(nextTheme);
    if (nextTheme.pageMargins) updateGlobalMargins(nextTheme.pageMargins);
  }

  function loadThemeFromList(themeId) {
    try {
      const raw = localStorage.getItem(themeStorageKey(themeId));
      if (!raw) {
        window.alert("Tema non trovato.");
        return;
      }
      const incoming = JSON.parse(raw);
      applySavedTheme(incoming);
      setCurrentThemeId(themeId);
      setTopbarMenuOpen(false);
    } catch (err) {
      window.alert(`Caricamento tema fallito: ${err?.message || "errore locale"}`);
    }
  }

  function deleteThemeFromList(themeId) {
    if (!window.confirm("Eliminare il tema salvato?")) return;
    try {
      localStorage.removeItem(themeStorageKey(themeId));
      const nextList = savedThemes.filter((t) => t.id !== themeId);
      saveThemesIndex(nextList);
      setSavedThemes(nextList);
      if (currentThemeId === themeId) setCurrentThemeId(null);
    } catch (err) {
      window.alert(`Eliminazione tema fallita: ${err?.message || "errore locale"}`);
    }
  }

  async function loadProjectFromList(projectId) {
    try {
      const raw = localStorage.getItem(projectStorageKey(projectId));
      if (!raw) {
        window.alert("Progetto non trovato.");
        return;
      }
      const incoming = JSON.parse(raw);
      const base = createDefaultState();
      skipNextPageFormatAdjustRef.current = true;
      setState({
        ...base,
        ...incoming,
        works: (incoming.works || []).map((w) => ({ ...normalizeWorkData(w), imageUrl: "" })),
        selectedElement: null,
        currentSpread: 0,
      });
      setCurrentProjectId(projectId);
      setTopbarMenuOpen(false);
    } catch (err) {
      window.alert(`Caricamento progetto fallito: ${err?.message || "errore locale"}`);
    }
  }

  function deleteProjectFromList(projectId) {
    if (!window.confirm("Eliminare il progetto salvato?")) return;
    try {
      localStorage.removeItem(projectStorageKey(projectId));
      const nextList = savedProjects.filter((p) => p.id !== projectId);
      saveProjectsIndex(nextList);
      setSavedProjects(nextList);
      if (currentProjectId === projectId) setCurrentProjectId(null);
    } catch (err) {
      window.alert(`Eliminazione progetto fallita: ${err?.message || "errore locale"}`);
    }
  }

  async function printCatalogPdf() {
    setTopbarMenuOpen(false);
    const printCatalogEl = document.querySelector(".print-catalog");
    const pageEls = Array.from(document.querySelectorAll(".print-page-sheet"));
    const pageSourceEls = Array.from(document.querySelectorAll(".print-page-source"));
    const exportPages = (renderPages || []).filter(Boolean);
    if (!printCatalogEl || !pageEls.length) {
      window.alert("Nessuna pagina disponibile per l'esportazione PDF.");
      return;
    }

    const previousCatalogStyle = printCatalogEl.getAttribute("style");
    const previousPageStyles = pageEls.map((el) => el.getAttribute("style"));
    const previousPageSourceStyles = pageSourceEls.map((el) => el.getAttribute("style"));
    try {
      setPrintProgress({ active: true, current: 0, total: pageEls.length, message: "Preparazione export PDF..." });
      const [{ default: html2canvas }, { jsPDF }] = await Promise.all([import("html2canvas"), import("jspdf")]);
      const fmt = getPageFormat(state.pageFormat);
      const orientation = fmt.width > fmt.height ? "landscape" : "portrait";
      const targetWpx = Math.round(mmToCssPx(fmt.width));
      const targetHpx = Math.round(mmToCssPx(fmt.height));
      const pdf = new jsPDF({
        orientation,
        unit: "mm",
        format: [fmt.width, fmt.height],
        compress: true,
      });
      const scale = Math.min(3, Math.max(2, window.devicePixelRatio || 1));

      printCatalogEl.setAttribute(
        "style",
        `${previousCatalogStyle || ""};display:block;position:fixed;left:-100000px;top:0;opacity:1;pointer-events:none;z-index:-1;`,
      );
      pageEls.forEach((el, pageIdx) => {
        const pageBg = exportPages[pageIdx]?.bgColor || "#ffffff";
        el.setAttribute(
          "style",
          `${el.getAttribute("style") || ""};display:block;position:relative;overflow:hidden;background:${pageBg};width:${targetWpx}px;height:${targetHpx}px;`,
        );
      });
      pageSourceEls.forEach((el) => {
        el.setAttribute(
          "style",
          `${el.getAttribute("style") || ""};position:absolute;left:0;top:0;transform-origin:top left;`,
        );
      });

      for (let i = 0; i < pageEls.length; i += 1) {
        setPrintProgress({ active: true, current: i + 1, total: pageEls.length, message: `Rendering pagina ${i + 1}/${pageEls.length}...` });
        await new Promise((resolve) => requestAnimationFrame(resolve));
        const canvas = await html2canvas(pageEls[i], {
          backgroundColor: exportPages[i]?.bgColor || "#ffffff",
          scale,
          useCORS: true,
          logging: false,
        });
        const imgData = canvas.toDataURL("image/png");
        if (i > 0) pdf.addPage([fmt.width, fmt.height], orientation);
        pdf.addImage(imgData, "PNG", 0, 0, fmt.width, fmt.height);
      }

      const baseName = slugifyFileBaseName(state.projectTitle, "catalogo-book");
      setPrintProgress({ active: true, current: pageEls.length, total: pageEls.length, message: "Salvataggio PDF..." });
      pdf.save(`${baseName || "catalogo-book"}.pdf`);
    } catch (err) {
      window.alert(`Esportazione PDF fallita: ${err?.message || "errore sconosciuto"}`);
    } finally {
      if (previousCatalogStyle == null) printCatalogEl.removeAttribute("style");
      else printCatalogEl.setAttribute("style", previousCatalogStyle);
      pageEls.forEach((el, idx) => {
        const prev = previousPageStyles[idx];
        if (prev == null) el.removeAttribute("style");
        else el.setAttribute("style", prev);
      });
      pageSourceEls.forEach((el, idx) => {
        const prev = previousPageSourceStyles[idx];
        if (prev == null) el.removeAttribute("style");
        else el.setAttribute("style", prev);
      });
      setPrintProgress({ active: false, current: 0, total: 0, message: "" });
    }
  }

  function createNewProject() {
    if (!window.confirm("Creare un nuovo progetto vuoto? Le modifiche non salvate del progetto corrente andranno perse.")) return;
    skipNextPageFormatAdjustRef.current = true;
    setState(createDefaultState(state.theme));
    setCurrentProjectId(null);
    setBookView({ zoom: 1, panX: 0, panY: 0 });
    setTopbarMenuOpen(false);
    setThemeOpen(false);
    setHelpOpen(false);
  }

  return (
    <div
      className={`app-shell ${dragDropActive ? "drag-active" : ""}`}
      style={{
        "--theme-font": state.theme.fontFamily,
        "--theme-color": state.theme.textColor,
        "--accent-color": state.theme.accentColor,
        "--ui-tint": state.theme.uiTint,
        "--paper-shadow": state.theme.paperShadow,
        "--theme-body-size": `${state.theme.bodyFontSize}px`,
        "--theme-title-size": `${state.theme.titleFontSize}px`,
        "--print-page-w-mm": `${currentFormat.width}mm`,
        "--print-page-h-mm": `${currentFormat.height}mm`,
        fontFamily: state.theme.fontFamily,
        color: state.theme.textColor,
      }}
      onDrop={onFrameDrop}
      onDragOver={onFrameDragOver}
      onDragLeave={onFrameDragLeave}
    >
      <header className="topbar">
        <div className="topbar-main">
          <div>
            <div className="topbar-title-row">
              <h1>Catalogo Opere</h1>
              <input
                className="page-title-chip"
                aria-label="Titolo progetto"
                value={state.projectTitle || ""}
                onChange={(e) => patchState({ projectTitle: e.target.value })}
                placeholder="Titolo progetto"
              />
            </div>
            {persistWarning && <p className="warn-text">{persistWarning}</p>}
          </div>
          <div ref={topbarActionsRef} className="topbar-actions">
            <IconButton icon="help" onClick={() => setHelpOpen((v) => !v)} title="Guida icone" ariaLabel="Guida icone" />
            <IconButton icon="theme" onClick={() => setThemeOpen((v) => !v)} title="Tema" ariaLabel="Tema" />
            <IconButton icon="menu" onClick={() => setTopbarMenuOpen((v) => !v)} title="Menu" ariaLabel="Menu" />
            {topbarMenuOpen && (
              <div className="topbar-overflow">
                <button onClick={createNewProject}>Nuovo progetto</button>
                <button onClick={saveProjectQuick}>Salva progetto</button>
                <button onClick={() => openSaveProjectDialog("saveAs")}>Salva come...</button>
                <button onClick={renameCurrentProject} disabled={!currentProjectId}>Rinomina progetto</button>
                <button onClick={saveThemeQuick}>Salva tema</button>
                <button onClick={() => openSaveThemeDialog("saveAs")}>Salva tema come...</button>
                <button onClick={renameCurrentTheme} disabled={!currentThemeId}>Rinomina tema</button>
                <button onClick={exportCatalogJson}>Esporta JSON</button>
                <button onClick={printCatalogPdf}>Esporta PDF Book</button>
                <button onClick={() => importJsonRef.current?.click()}>Importa JSON</button>
                <div className="saved-projects-list">
                  <small>Progetti salvati</small>
                  {!savedProjects.length && <div className="saved-project-row empty">Nessun progetto</div>}
                  {savedProjects.map((project) => (
                    <div key={project.id} className="saved-project-row">
                      <button
                        className={`saved-project-load ${currentProjectId === project.id ? "active" : ""}`}
                        onClick={() => loadProjectFromList(project.id)}
                        title={project.updatedAt ? new Date(project.updatedAt).toLocaleString("it-IT") : undefined}
                      >
                        {project.name}
                      </button>
                      <button className="saved-project-delete" onClick={() => deleteProjectFromList(project.id)} title="Elimina progetto">
                        ×
                      </button>
                    </div>
                  ))}
                </div>
                <div className="saved-projects-list">
                  <small>Temi salvati</small>
                  {!savedThemes.length && <div className="saved-project-row empty">Nessun tema</div>}
                  {savedThemes.map((themeItem) => (
                    <div key={themeItem.id} className="saved-project-row">
                      <button
                        className={`saved-project-load ${currentThemeId === themeItem.id ? "active" : ""}`}
                        onClick={() => loadThemeFromList(themeItem.id)}
                        title={themeItem.updatedAt ? new Date(themeItem.updatedAt).toLocaleString("it-IT") : undefined}
                      >
                        {themeItem.name}
                      </button>
                      <button className="saved-project-delete" onClick={() => deleteThemeFromList(themeItem.id)} title="Elimina tema">
                        ×
                      </button>
                    </div>
                  ))}
                </div>
              </div>
            )}
            {helpOpen && <HelpOverlay onClose={() => setHelpOpen(false)} />}
            <input ref={importJsonRef} type="file" accept="application/json,.json" hidden onChange={onImportJsonInput} />
          </div>
        </div>
        <div className="book-toolbar top-tools">
          <div className="nav-group">
            <IconButton
              icon="chevronLeft"
              title="Spread precedente"
              ariaLabel="Spread precedente"
              onClick={() => patchState((p) => ({ ...p, currentSpread: Math.max(0, p.currentSpread - 1) }))}
            />
            <span className="toolbar-badge" title={`Spread ${currentSpreadIndex + 1} di ${Math.max(spreads.length, 1)}`}>
              {currentSpreadIndex + 1}/{Math.max(spreads.length, 1)}
            </span>
            <IconButton icon="plus" title="Aggiungi pagina" ariaLabel="Aggiungi pagina" onClick={addInnerPage} />
            <IconButton
              icon="chevronRight"
              title="Spread successivo"
              ariaLabel="Spread successivo"
              onClick={() =>
                patchState((p) => ({ ...p, currentSpread: Math.min(spreads.length - 1, p.currentSpread + 1) }))
              }
            />
            <button
              className={`icon-btn ${autoLayoutMode === "hero" ? "active-toggle" : ""}`}
              onClick={() => {
                autoLayoutModeRef.current = "hero";
                setAutoLayoutMode("hero");
                applyLayoutPreset("hero");
              }}
              title="Layout 1 (hero) / preset auto"
              aria-label="Layout 1"
            >
              1
            </button>
            <button
              className={`icon-btn ${autoLayoutMode === "two-cols" ? "active-toggle" : ""}`}
              onClick={() => {
                autoLayoutModeRef.current = "two-cols";
                setAutoLayoutMode("two-cols");
                applyLayoutPreset("two-cols");
              }}
              title="Layout 2 colonne / preset auto"
              aria-label="Layout 2"
            >
              2
            </button>
            <button
              className={`icon-btn ${autoLayoutMode === "grid4" ? "active-toggle" : ""}`}
              onClick={() => {
                autoLayoutModeRef.current = "grid4";
                setAutoLayoutMode("grid4");
                applyLayoutPreset("grid4");
              }}
              title="Layout griglia 4 / preset auto"
              aria-label="Layout 4"
            >
              4
            </button>
            <button
              className="icon-btn tool-auto-btn"
              onClick={autoGeneratePagesFromCatalog}
              title="Auto impagina con preset 1/2/4 selezionato"
              aria-label="Auto impagina"
            >
              [auto]
            </button>
            <span className="toolbar-badge" title="Preset auto-layout attivo">
              Auto: {autoLayoutMode === "hero" ? "1" : autoLayoutMode === "two-cols" ? "2" : "4"}
            </span>
            <button
              className={`icon-btn ${state.layoutAssist?.distributeRespectMargins ? "active-toggle" : ""}`}
              onClick={() =>
                patchState((p) => ({
                  ...p,
                  layoutAssist: {
                    ...p.layoutAssist,
                    distributeRespectMargins: !p.layoutAssist?.distributeRespectMargins,
                  },
                }))
              }
              title={state.layoutAssist?.distributeRespectMargins ? "Distribuzione su margini" : "Distribuzione tra estremi elementi"}
              aria-label="Modalità distribuzione"
            >
              <Icon name="distributeBounds" />
            </button>
            <button
              className={`icon-btn ${state.bookViewMode === "real" ? "active-toggle" : ""}`}
              onClick={() =>
                patchState((p) => ({
                  ...p,
                  bookViewMode: p.bookViewMode === "real" ? "technical" : "real",
                  currentSpread: 0,
                }))
              }
              title={state.bookViewMode === "real" ? "Vista libro reale" : "Vista impaginazione tecnica"}
              aria-label="Modalità spread"
            >
              <Icon name="spreadMode" />
            </button>
            <IconButton icon="minus" onClick={() => setBookView((v) => ({ ...v, zoom: Math.max(0.4, +(v.zoom - 0.1).toFixed(2)) }))} title="Zoom out" ariaLabel="Zoom out" />
            <span className="zoom-readout">{Math.round(bookView.zoom * 100)}%</span>
            <IconButton icon="plus" onClick={() => setBookView((v) => ({ ...v, zoom: Math.min(3, +(v.zoom + 0.1).toFixed(2)) }))} title="Zoom in" ariaLabel="Zoom in" />
            <IconButton icon="reset" onClick={() => setBookView({ zoom: 1, panX: 0, panY: 0 })} title="Reset vista" ariaLabel="Reset vista" />
            <label className="inline-mini">
              F
              <select
                value={state.pageFormat || "a4-portrait"}
                onChange={(e) => patchState((p) => ({ ...p, pageFormat: e.target.value }))}
              >
                {PAGE_FORMATS.map((fmt) => (
                  <option key={fmt.id} value={fmt.id}>{fmt.label}</option>
                ))}
              </select>
            </label>
            <button
              className={`icon-btn ${state.layoutAssist?.snapToGrid ? "active-toggle" : ""}`}
              onClick={() =>
                patchState((p) => ({
                  ...p,
                  layoutAssist: { ...p.layoutAssist, snapToGrid: !p.layoutAssist?.snapToGrid },
                }))
              }
              title="Aggancia a griglia"
              aria-label="Snap griglia"
            >
              <Icon name="grid" />
            </button>
            <button
              className={`icon-btn ${state.layoutAssist?.showGuides ? "active-toggle" : ""}`}
              onClick={() =>
                patchState((p) => ({
                  ...p,
                  layoutAssist: { ...p.layoutAssist, showGuides: !p.layoutAssist?.showGuides },
                }))
              }
              title="Mostra guide magnetiche"
              aria-label="Guide"
            >
              <Icon name="guides" />
            </button>
            <button
              className={`icon-btn ${state.layoutAssist?.boundMode === "page" ? "active-toggle" : ""}`}
              onClick={() =>
                patchState((p) => ({
                  ...p,
                  layoutAssist: {
                    ...p.layoutAssist,
                    boundMode: p.layoutAssist?.boundMode === "page" ? "margins" : "page",
                  },
                }))
              }
              title={state.layoutAssist?.boundMode === "page" ? "Vincolo: pagina intera" : "Vincolo: margini"}
              aria-label="Modalità vincolo"
            >
              <Icon name="bounds" />
            </button>
            <label className="inline-mini">
              G
              <input
                type="number"
                min="4"
                max="48"
                value={state.layoutAssist?.gridSize ?? 12}
                onChange={(e) =>
                  patchState((p) => ({
                    ...p,
                    layoutAssist: { ...p.layoutAssist, gridSize: Math.max(4, Number(e.target.value) || 12) },
                  }))
                }
              />
            </label>
            <IconButton icon="text" onClick={addTextToActivePage} title="Aggiungi testo" ariaLabel="Aggiungi testo" />
            <IconButton icon="image" onClick={addSelectedWorkToActivePage} disabled={!selectedWork} title="Aggiungi opera selezionata" ariaLabel="Aggiungi opera selezionata" />
            <IconButton icon="fit" onClick={normalizeActivePageNow} title="Riporta elementi dentro vincoli" ariaLabel="Riporta dentro" />
            <IconButton icon="trash" onClick={deleteSelectedElement} disabled={!state.selectedElement} title="Elimina elemento selezionato" ariaLabel="Elimina elemento" />
          </div>
        </div>
      </header>
      {printProgress.active && (
        <div className="export-progress" role="status" aria-live="polite">
          <strong>{printProgress.message || "Esportazione PDF..."}</strong>
          <div className="export-progress-bar">
            <span
              style={{
                width: `${Math.round(((printProgress.current || 0) * 100) / Math.max(1, printProgress.total || 1))}%`,
              }}
            />
          </div>
          <small>
            {printProgress.current}/{printProgress.total}
          </small>
        </div>
      )}
      <main className="workspace">
        <Filmstrip
          works={state.works}
          selectedWorkId={state.selectedWorkId}
          onSelect={(id) => patchState((prev) => ({ ...prev, selectedWorkId: id }))}
          onEdit={(work) => openEditWork(work)}
          onDelete={deleteWork}
          onAdd={() => openCreateWork()}
          onImportFiles={handleDropFiles}
          vertical
        />

        <section className="book-panel">
          <BookSpread
            spread={currentSpread}
            works={state.works}
            theme={state.theme}
            activePageId={state.activePageId}
            selectedElement={state.selectedElement}
            onSelectPage={selectPage}
            onSelectElement={(sel) => patchState((prev) => ({ ...prev, selectedElement: sel }))}
            onMoveElement={(pageId, kind, elementId, patch) => updateElementOnPage(pageId, kind, elementId, patch)}
            onDropWorkToPage={addWorkToPageAtPosition}
            onMovePlacementToPage={movePlacementToPage}
            layoutAssist={state.layoutAssist}
            pageFormat={state.pageFormat}
            onDeleteElementDirect={deleteElementDirect}
            view={bookView}
            onViewChange={setBookView}
            onPageMetrics={(pageId, metrics) =>
              setPageMetrics((prev) => {
                const curr = prev[pageId];
                if (
                  curr &&
                  curr.width === metrics.width &&
                  curr.height === metrics.height &&
                  curr.pageWidth === metrics.pageWidth &&
                  curr.pageHeight === metrics.pageHeight
                ) {
                  return prev;
                }
                return { ...prev, [pageId]: metrics };
              })
            }
          />

          {dragDropActive && (
            <div className="drop-overlay">
              <div>
                <strong>Rilascia immagine qui</strong>
                <p>Si aprirà l'editor dell'opera</p>
              </div>
            </div>
          )}
        </section>

        <aside className="side-panel">
          <PanelSection title="Pagina attiva">
            {activeEditablePage ? (
              <>
                <PageInspector
                  page={activeEditablePage}
                  onChange={(patch) => patchPage(activeEditablePage.id, (p) => ({ ...p, ...patch }))}
                  onDeletePage={() => removeInnerPage(activeEditablePage.id)}
                />
                <div className="stack-fields">
                  <label>Opere nella pagina</label>
                  {activePageWorks.length ? (
                    <div className="page-works-list">
                      {activePageWorks.map(({ placementId, work }) => {
                        const isPlacementSelected =
                          state.selectedElement?.kind === "placement" && state.selectedElement?.elementId === placementId;
                        return (
                          <div key={placementId} className="page-work-row">
                            <button
                              className={`page-work-select ${isPlacementSelected ? "active" : ""}`}
                              onClick={() =>
                                patchState((prev) => ({
                                  ...prev,
                                  selectedWorkId: work.id,
                                  selectedElement: { pageId: activeEditablePage.id, kind: "placement", elementId: placementId },
                                }))
                              }
                            >
                              {workLabel(work)}
                            </button>
                            <button className="small-btn" onClick={() => openEditWork(work)}>Modifica</button>
                          </div>
                        );
                      })}
                    </div>
                  ) : (
                    <p className="muted">Nessuna opera in questa pagina.</p>
                  )}
                </div>
              </>
            ) : (
              <p className="muted">Seleziona una pagina</p>
            )}
          </PanelSection>

          <PanelSection title="Elemento selezionato">
            {state.selectedElement && selectedElementData ? (
              <ElementInspector
                kind={state.selectedElement.kind}
                data={selectedElementData}
                onChange={(patch) =>
                  updateElementOnPage(state.selectedElement.pageId, state.selectedElement.kind, state.selectedElement.elementId, patch)
                }
              />
            ) : (
              <p className="muted">Seleziona un testo o una opera impaginata.</p>
            )}
          </PanelSection>

          <PanelSection title="Opera selezionata">
            {effectiveSelectedWork ? (
              <div className="selected-work-card">
                {effectiveSelectedWork.imageUrl ? (
                  <img src={effectiveSelectedWork.imageUrl} alt={workLabel(effectiveSelectedWork)} />
                ) : (
                  <div className="img-ph">Nessuna immagine</div>
                )}
                <div>
                  <strong>{workLabel(effectiveSelectedWork)}</strong>
                  <p>{effectiveSelectedWork.author || "Autore non indicato"}</p>
                  <button className="small-btn" onClick={() => openEditWork(effectiveSelectedWork)}>Modifica opera</button>
                </div>
              </div>
            ) : (
              <p className="muted">Seleziona un'opera dal filmstrip.</p>
            )}
          </PanelSection>
        </aside>
      </main>

      <section className="print-catalog" aria-hidden="true">
        {renderPages.map((page, idx) => {
          if (!page) return null;
          const m = page.margins || { top: 0, right: 0, bottom: 0, left: 0 };
          const fallbackInner = getPageContentBounds(page, state.pageFormat);
          const measuredPage = pageMetrics?.[page.id];
          const canonicalMeasuredPage = Object.values(pageMetrics || {}).find((pm) => pm?.pageWidth && pm?.pageHeight);
          const measuredInner = measuredPage || Object.values(pageMetrics || {}).find((pm) => pm?.width && pm?.height) || fallbackInner;
          const sourceW = Math.max(
            1,
            Math.round(measuredPage?.pageWidth || canonicalMeasuredPage?.pageWidth || ((measuredInner?.width || 0) + (m.left || 0) + (m.right || 0))),
          );
          const sourceH = Math.max(
            1,
            Math.round(measuredPage?.pageHeight || canonicalMeasuredPage?.pageHeight || ((measuredInner?.height || 0) + (m.top || 0) + (m.bottom || 0))),
          );
          const targetWpx = mmToCssPx(currentFormat.width);
          const targetHpx = mmToCssPx(currentFormat.height);
          const scale = Math.min(targetWpx / sourceW, targetHpx / sourceH);
          return (
            <div key={`print_page_${page.id}_${idx}`} className="print-page-sheet">
              <div
                className="print-page-source"
                style={{
                  width: `${sourceW}px`,
                  height: `${sourceH}px`,
                  transform: `scale(${scale})`,
                  "--print-scale": String(scale),
                }}
              >
                <PageCanvas
                  page={page}
                  side={idx % 2 ? "left" : "right"}
                  works={state.works}
                  theme={state.theme}
                  active={false}
                  selectedElement={null}
                  onSelectPage={() => {}}
                  onSelectElement={() => {}}
              onMoveElement={() => {}}
              onDropWorkToPage={() => {}}
              onMovePlacementToPage={() => {}}
                  layoutAssist={{ snapToGrid: false, showGuides: false, gridSize: 12, snapThreshold: 0, boundMode: "margins" }}
                  pageFormat={state.pageFormat}
                  onDeleteElementDirect={() => {}}
                  zoomScale={1}
                  onPageMetrics={() => {}}
                />
              </div>
            </div>
          );
        })}
      </section>

      {themeOpen && (
        <ThemePanel
          theme={state.theme}
          onClose={() => setThemeOpen(false)}
          onChange={patchTheme}
          onMarginsChange={updateGlobalMargins}
        />
      )}

      {workEditor.open && (
        <WorkEditorModal
          draft={workEditor.draft}
          onCancel={() => setWorkEditor({ open: false, draft: null, mode: "create" })}
          onSave={saveWork}
        />
      )}

      {projectDialog.open && (
        <ProjectNameModal
          mode={projectDialog.mode}
          name={projectDialog.name}
          onCancel={() => setProjectDialog({ open: false, mode: "save", name: "" })}
          onChangeName={(name) => setProjectDialog((prev) => ({ ...prev, name }))}
          onConfirm={() => {
            if (projectDialog.mode === "rename") {
              const existing = savedProjects.find((p) => p.id === currentProjectId);
              persistProjectByName(projectDialog.name, existing?.id || currentProjectId);
            } else if (projectDialog.mode === "saveAs") {
              persistProjectByName(projectDialog.name, uid("prj"));
            } else {
              persistProjectByName(projectDialog.name, currentProjectId || null);
            }
          }}
        />
      )}

      {themeDialog.open && (
        <ThemeNameModal
          mode={themeDialog.mode}
          name={themeDialog.name}
          onCancel={() => setThemeDialog({ open: false, mode: "save", name: "" })}
          onChangeName={(name) => setThemeDialog((prev) => ({ ...prev, name }))}
          onConfirm={() => {
            if (themeDialog.mode === "rename") {
              const existing = savedThemes.find((t) => t.id === currentThemeId);
              persistThemeByName(themeDialog.name, existing?.id || currentThemeId);
            } else if (themeDialog.mode === "saveAs") {
              persistThemeByName(themeDialog.name, uid("thm"));
            } else {
              persistThemeByName(themeDialog.name, currentThemeId || null);
            }
          }}
        />
      )}
    </div>
  );
}

function ProjectNameModal({ mode, name, onChangeName, onCancel, onConfirm }) {
  const title = mode === "rename" ? "Rinomina progetto" : mode === "saveAs" ? "Salva progetto come" : "Salva progetto";
  return (
    <div className="modal-backdrop" onClick={onCancel}>
      <div className="modal modal-sm" onClick={(e) => e.stopPropagation()}>
        <div className="modal-head">
          <h2>{title}</h2>
          <button onClick={onCancel}>✕</button>
        </div>
        <label>
          Nome progetto
          <input
            autoFocus
            value={name}
            onChange={(e) => onChangeName(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter") onConfirm();
              if (e.key === "Escape") onCancel();
            }}
          />
        </label>
        <div className="modal-foot">
          <button className="ghost-btn" onClick={onCancel}>Annulla</button>
          <button className="primary-btn" onClick={onConfirm} disabled={!String(name || "").trim()}>Conferma</button>
        </div>
      </div>
    </div>
  );
}

function ThemeNameModal({ mode, name, onChangeName, onCancel, onConfirm }) {
  const title = mode === "rename" ? "Rinomina tema" : mode === "saveAs" ? "Salva tema come" : "Salva tema";
  return (
    <div className="modal-backdrop" onClick={onCancel}>
      <div className="modal modal-sm" onClick={(e) => e.stopPropagation()}>
        <div className="modal-head">
          <h2>{title}</h2>
          <button onClick={onCancel}>✕</button>
        </div>
        <label>
          Nome tema
          <input
            autoFocus
            value={name}
            onChange={(e) => onChangeName(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter") onConfirm();
              if (e.key === "Escape") onCancel();
            }}
          />
        </label>
        <div className="modal-foot">
          <button className="ghost-btn" onClick={onCancel}>Annulla</button>
          <button className="primary-btn" onClick={onConfirm} disabled={!String(name || "").trim()}>Conferma</button>
        </div>
      </div>
    </div>
  );
}

function PanelSection({ title, children }) {
  return (
    <section className="panel-section">
      <h3>{title}</h3>
      {children}
    </section>
  );
}

function Icon({ name }) {
  const common = { width: 16, height: 16, viewBox: "0 0 24 24", fill: "none", stroke: "currentColor", strokeWidth: 1.8, strokeLinecap: "round", strokeLinejoin: "round", "aria-hidden": true };
  const icons = {
    theme: <svg {...common}><circle cx="12" cy="12" r="3.2" /><path d="M12 2.5v2.3M12 19.2v2.3M2.5 12h2.3M19.2 12h2.3M5.3 5.3l1.6 1.6M17.1 17.1l1.6 1.6M18.7 5.3l-1.6 1.6M6.9 17.1l-1.6 1.6"/></svg>,
    menu: <svg {...common}><path d="M5 7h14M5 12h14M5 17h14" /></svg>,
    help: <svg {...common}><circle cx="12" cy="12" r="9" /><path d="M9.9 9.3a2.4 2.4 0 1 1 3.8 2c-.9.7-1.7 1.2-1.7 2.4" /><path d="M12 16.9h.01" /></svg>,
    chevronLeft: <svg {...common}><path d="M14.5 5.5 8 12l6.5 6.5" /></svg>,
    chevronRight: <svg {...common}><path d="M9.5 5.5 16 12l-6.5 6.5" /></svg>,
    plus: <svg {...common}><path d="M12 5v14M5 12h14" /></svg>,
    minus: <svg {...common}><path d="M5 12h14" /></svg>,
    reset: <svg {...common}><path d="M20 12a8 8 0 1 1-2.3-5.7" /><path d="M20 4.5v5h-5" /></svg>,
    grid: <svg {...common}><path d="M4 4h16v16H4zM4 10h16M10 4v16" /></svg>,
    guides: <svg {...common}><path d="M12 3v18M3 12h18" /><circle cx="12" cy="12" r="2" /></svg>,
    bounds: <svg {...common}><path d="M4 4h16v16H4z" /><path d="M8 8h8v8H8z" /></svg>,
    distributeBounds: <svg {...common}><path d="M4 6h16M4 18h16" /><path d="M7 9v6M12 7v10M17 9v6" /></svg>,
    spreadMode: <svg {...common}><path d="M3.5 6h8v12h-8zM12.5 6h8v12h-8z" /><path d="M12 6v12" /></svg>,
    text: <svg {...common}><path d="M5 6h14M12 6v12M8.5 18h7" /></svg>,
    image: <svg {...common}><rect x="4" y="5" width="16" height="14" rx="1.5" /><path d="m7 15 3.2-3.3 2.7 2.7 2.1-2.2L18 15" /><circle cx="9" cy="9" r="1.2" /></svg>,
    fit: <svg {...common}><path d="M4 9V4h5M20 9V4h-5M4 15v5h5M20 15v5h-5" /></svg>,
    trash: <svg {...common}><path d="M4 7h16M9 7V5h6v2M8 7l.7 12h6.6L16 7" /></svg>,
    upload: <svg {...common}><path d="M12 16V6M8.5 9.5 12 6l3.5 3.5" /><path d="M5 18.5h14" /></svg>,
    close: <svg {...common}><path d="M6 6l12 12M18 6 6 18" /></svg>,
  };
  return icons[name] || null;
}

function IconButton({ icon, title, ariaLabel, ...props }) {
  return (
    <button className="icon-btn" title={title} data-tip={title} aria-label={ariaLabel || title} {...props}>
      <Icon name={icon} />
    </button>
  );
}

function HelpOverlay({ onClose }) {
  const items = [
    ["Frecce", "Sfoglia spread"],
    ["+", "Aggiungi pagina"],
    ["1/2/4", "Preset layout pagina"],
    ["[auto]", "Una pagina per opera dal catalogo"],
    ["Distrib.", "Distribuzione su margini on/off"],
    ["Spread", "Vista libro reale / tecnica"],
    ["Zoom", "Zoom vista e reset"],
    ["Griglia", "Snap griglia"],
    ["Guide", "Guide magnetiche"],
    ["Vincoli", "Margini / pagina intera"],
    ["T / Immagine", "Aggiungi testo / opera selezionata"],
    ["Fit", "Riporta dentro i vincoli"],
    ["Cestino", "Elimina elemento selezionato"],
  ];
  return (
    <div className="topbar-overflow help-panel">
      <div className="help-head">
        <strong>Legenda strumenti</strong>
        <button className="icon-btn" onClick={onClose} aria-label="Chiudi guida" title="Chiudi guida">
          <Icon name="close" />
        </button>
      </div>
      <div className="help-list">
        {items.map(([k, v]) => (
          <div key={k} className="help-row">
            <span>{k}</span>
            <small>{v}</small>
          </div>
        ))}
      </div>
    </div>
  );
}

function ThemePanel({ theme, onChange, onMarginsChange, onClose }) {
  return (
    <div className="floating-panel theme-panel">
      <div className="floating-head">
        <strong>Tema catalogo</strong>
        <button onClick={onClose}>✕</button>
      </div>
      <label>
        Font
        <select
          value={theme.fontFamily}
          onChange={(e) => onChange({ fontFamily: e.target.value })}
          style={{ fontFamily: theme.fontFamily }}
        >
          {FONT_OPTIONS.map((opt) => (
            <option key={opt.value} value={opt.value} style={{ fontFamily: opt.value }}>
              {opt.label}
            </option>
          ))}
        </select>
      </label>
      <RangeField label="Font base" min={12} max={24} value={theme.bodyFontSize} onChange={(v) => onChange({ bodyFontSize: v })} />
      <RangeField label="Titoli" min={18} max={42} value={theme.titleFontSize} onChange={(v) => onChange({ titleFontSize: v })} />
      <RangeField label="Peso" min={300} max={800} step={100} value={theme.fontWeight} onChange={(v) => onChange({ fontWeight: v })} />
      <label>
        Colore testo
        <input type="color" value={theme.textColor} onChange={(e) => onChange({ textColor: e.target.value })} />
      </label>
      <label>
        Colore accento
        <input type="color" value={theme.accentColor} onChange={(e) => onChange({ accentColor: e.target.value })} />
      </label>
      <label>
        Tinta UI
        <input type="color" value={theme.uiTint} onChange={(e) => onChange({ uiTint: e.target.value })} />
      </label>
      <label>
        Sfondo pagina default
        <input
          type="color"
          value={theme.defaultPageBgColor || "#ffffff"}
          onChange={(e) => onChange({ defaultPageBgColor: e.target.value })}
        />
      </label>
      <ColorAlphaField
        label="Sfondo rettangoli testo default"
        value={theme.defaultTextBgColor || "rgba(255, 255, 255, 0.42)"}
        fallback="rgba(255, 255, 255, 0.42)"
        onChange={(color) => onChange({ defaultTextBgColor: color })}
      />
      <label>
        Colore numero pagina default
        <input
          type="color"
          value={theme.defaultPageNumberColor || "#6b614f"}
          onChange={(e) => onChange({ defaultPageNumberColor: e.target.value })}
        />
      </label>
      <label className="inline-check">
        <input
          type="checkbox"
          checked={theme.autoShowCaptionDefault ?? true}
          onChange={(e) => onChange({ autoShowCaptionDefault: e.target.checked })}
        />
        Default mostra didascalia (auto)
      </label>
      <RangeField
        label="Bordo elementi default %"
        min={0}
        max={20}
        value={theme.defaultElementBorderPct ?? 3}
        onChange={(v) => onChange({ defaultElementBorderPct: v })}
      />
      <label>
        Colore bordo elementi default
        <input
          type="color"
          value={theme.defaultElementBorderColor || "#ffffff"}
          onChange={(e) => onChange({ defaultElementBorderColor: e.target.value })}
        />
      </label>
      <label>
        Template Introduzione/Curatoriale (MD)
        <textarea
          rows={12}
          value={theme.defaultIntroCuratorialMd || DEFAULT_INTRO_CURATORIAL_MD}
          onChange={(e) => onChange({ defaultIntroCuratorialMd: e.target.value })}
        />
      </label>
      <label>
        Template Prefazione Titolo (MD)
        <textarea
          rows={3}
          value={theme.defaultPrefaceTitleMd || DEFAULT_PREFACE_TITLE_MD}
          onChange={(e) => onChange({ defaultPrefaceTitleMd: e.target.value })}
        />
      </label>
      <label>
        Template Prefazione Testo (MD)
        <textarea
          rows={8}
          value={theme.defaultPrefaceBodyMd || DEFAULT_PREFACE_BODY_MD}
          onChange={(e) => onChange({ defaultPrefaceBodyMd: e.target.value })}
        />
      </label>
      <label>
        Template Sintesi Retro (MD)
        <textarea
          rows={6}
          value={theme.defaultBackSummaryMd || DEFAULT_BACK_SUMMARY_MD}
          onChange={(e) => onChange({ defaultBackSummaryMd: e.target.value })}
        />
      </label>
      <label>
        Template Biografia Retro (MD)
        <textarea
          rows={10}
          value={theme.defaultBackBioMd || DEFAULT_BACK_BIO_MD}
          onChange={(e) => onChange({ defaultBackBioMd: e.target.value })}
        />
      </label>
      <div className="grid-2">
        <RangeField
          label="Margine top"
          min={0}
          max={80}
          value={theme.pageMargins?.top ?? 28}
          onChange={(v) => onMarginsChange?.({ top: v })}
        />
        <RangeField
          label="Margine right"
          min={0}
          max={80}
          value={theme.pageMargins?.right ?? 28}
          onChange={(v) => onMarginsChange?.({ right: v })}
        />
        <RangeField
          label="Margine bottom"
          min={0}
          max={80}
          value={theme.pageMargins?.bottom ?? 38}
          onChange={(v) => onMarginsChange?.({ bottom: v })}
        />
        <RangeField
          label="Margine left"
          min={0}
          max={80}
          value={theme.pageMargins?.left ?? 28}
          onChange={(v) => onMarginsChange?.({ left: v })}
        />
      </div>
    </div>
  );
}

function PageInspector({ page, onChange, onDeletePage }) {
  return (
    <div className="stack-fields">
      <label>
        Sfondo pagina
        <input type="color" value={page.bgColor || "#ffffff"} onChange={(e) => onChange({ bgColor: e.target.value })} />
      </label>
      <label className="inline-check">
        <input type="checkbox" checked={!!page.showPageNumber} onChange={(e) => onChange({ showPageNumber: e.target.checked })} />
        Mostra numero pagina
      </label>
      <label>
        Colore numero pagina
        <input type="color" value={page.pageNumberColor || "#666666"} onChange={(e) => onChange({ pageNumberColor: e.target.value })} />
      </label>
      {page.type === "page" && (
        <button className="danger-btn" onClick={onDeletePage}>
          Elimina pagina
        </button>
      )}
    </div>
  );
}

function ElementInspector({ kind, data, onChange }) {
  if (!data) return null;
  if (kind === "text") {
    return (
      <div className="stack-fields">
        <label>
          Testo (Markdown)
          <textarea rows={5} value={data.text || ""} onChange={(e) => onChange({ text: e.target.value })} />
        </label>
        <RangeField label="Font size" min={10} max={42} value={data.fontSize || 16} onChange={(v) => onChange({ fontSize: v })} />
        <RangeField label="Peso" min={300} max={800} step={100} value={data.fontWeight || 500} onChange={(v) => onChange({ fontWeight: v })} />
        <label>
          Colore
          <input type="color" value={data.color || "#222222"} onChange={(e) => onChange({ color: e.target.value })} />
        </label>
        <ColorAlphaField
          label="Sfondo testo"
          value={data.bgColor || "rgba(255, 255, 255, 0.42)"}
          fallback="rgba(255, 255, 255, 0.42)"
          onChange={(color) => onChange({ bgColor: color })}
        />
        <label>
          Allineamento
          <select value={data.align || "left"} onChange={(e) => onChange({ align: e.target.value })}>
            <option value="left">Sinistra</option>
            <option value="center">Centro</option>
            <option value="right">Destra</option>
          </select>
        </label>
        <RangeField label="Bordo %" min={0} max={20} value={data.borderWidthPct ?? 5} onChange={(v) => onChange({ borderWidthPct: v })} />
        <label>
          Colore bordo
          <input type="color" value={data.borderColor || "#ffffff"} onChange={(e) => onChange({ borderColor: e.target.value })} />
        </label>
      </div>
    );
  }
  return (
    <div className="stack-fields">
      <label>
        Didascalia (Markdown)
        <textarea
          rows={4}
          value={data.captionOverride || ""}
          onChange={(e) => onChange({ captionOverride: e.target.value })}
          placeholder="Se vuoto usa didascalia automatica (titolo, autore, anno)"
        />
      </label>
      <RangeField label="Bordo immagine %" min={0} max={20} value={data.borderWidthPct ?? 5} onChange={(v) => onChange({ borderWidthPct: v })} />
      <label>
        Colore bordo immagine
        <input type="color" value={data.borderColor || "#ffffff"} onChange={(e) => onChange({ borderColor: e.target.value })} />
      </label>
      <label className="inline-check">
        <input type="checkbox" checked={!!data.showCaption} onChange={(e) => onChange({ showCaption: e.target.checked })} />
        Mostra didascalia
      </label>
      <RangeField
        label="Bordo didascalia %"
        min={0}
        max={20}
        value={data.captionBorderWidthPct ?? 5}
        onChange={(v) => onChange({ captionBorderWidthPct: v })}
      />
      <label>
        Colore bordo didascalia
        <input
          type="color"
          value={data.captionBorderColor || "#ffffff"}
          onChange={(e) => onChange({ captionBorderColor: e.target.value })}
        />
      </label>
    </div>
  );
}

function RangeField({ label, min, max, value, onChange, step = 1 }) {
  return (
    <label className="range-field">
      <span>{label}: {value}</span>
      <input type="range" min={min} max={max} step={step} value={value} onChange={(e) => onChange(Number(e.target.value))} />
    </label>
  );
}

function ColorAlphaField({ label, value, fallback = "rgba(255, 255, 255, 0.42)", onChange }) {
  const parsed = colorToHexAlpha(value, fallback);
  const setHex = (hex) => {
    const rgba = parseColorToRgbaParts(hex) || parseColorToRgbaParts(parsed.hex) || { r: 255, g: 255, b: 255, a: parsed.alpha };
    onChange?.(rgbaPartsToCss({ ...rgba, a: parsed.alpha }));
  };
  const setAlphaPct = (pct) => {
    const alpha = clamp01((Number(pct) || 0) / 100, parsed.alpha);
    const rgba = parseColorToRgbaParts(parsed.hex) || { r: 255, g: 255, b: 255, a: alpha };
    onChange?.(rgbaPartsToCss({ ...rgba, a: alpha }));
  };
  return (
    <label className="color-alpha-field">
      <span>{label}</span>
      <div className="color-alpha-row">
        <input type="color" value={parsed.hex} onChange={(e) => setHex(e.target.value)} />
        <input type="range" min={0} max={100} step={1} value={Math.round(parsed.alpha * 100)} onChange={(e) => setAlphaPct(e.target.value)} />
        <small>{Math.round(parsed.alpha * 100)}%</small>
      </div>
    </label>
  );
}

function NumberField({ label, value, onChange }) {
  return (
    <label>
      {label}
      <input type="number" value={Math.round(value || 0)} onChange={(e) => onChange(Number(e.target.value) || 0)} />
    </label>
  );
}

function Filmstrip({ works, selectedWorkId, onSelect, onEdit, onDelete, onAdd, onImportFiles, vertical = false }) {
  const fileRef = useRef(null);

  async function onFileInput(e) {
    if (e.target.files?.length) {
      await onImportFiles(e.target.files);
      e.target.value = "";
    }
  }

  return (
    <footer className={`filmstrip ${vertical ? "vertical" : ""}`}>
      <div className="filmstrip-head">
        <div className="filmstrip-actions">
          <button className={vertical ? "icon-btn" : ""} onClick={onAdd} title="Nuova opera" aria-label="Nuova opera">
            {vertical ? <Icon name="plus" /> : "+ Nuova opera"}
          </button>
          <button className={vertical ? "icon-btn" : ""} onClick={() => fileRef.current?.click()} title="Importa immagini" aria-label="Importa immagini">
            {vertical ? <Icon name="upload" /> : "Importa immagine"}
          </button>
          <input ref={fileRef} type="file" accept="image/*" multiple hidden onChange={onFileInput} />
        </div>
      </div>
      <div className="filmstrip-row">
        {works.map((work) => (
          <button
            key={work.id}
            className={`film-card ${selectedWorkId === work.id ? "selected" : ""}`}
            draggable
            onDragStart={(e) => {
              e.dataTransfer.setData("application/x-catalog-work-id", work.id);
              e.dataTransfer.effectAllowed = "copyMove";
            }}
            onClick={() => {
              if (selectedWorkId === work.id) onEdit(work);
              else onSelect(work.id);
            }}
            onDoubleClick={() => onEdit(work)}
            title="Click per selezionare, click su selezionata o doppio click per modificare"
          >
            <span
              className="film-delete"
              onClick={(e) => {
                e.stopPropagation();
                onDelete(work.id);
              }}
            >
              ×
            </span>
            {work.imageUrl ? <img src={work.imageUrl} alt={workLabel(work)} /> : <div className="thumb-placeholder">No img</div>}
            {!vertical && <small>{workLabel(work)}</small>}
          </button>
        ))}
        {!works.length && <div className="film-empty">Nessuna opera. Aggiungi dal pulsante o trascina una immagine sul frame.</div>}
      </div>
    </footer>
  );
}

function WorkEditorModal({ draft, onCancel, onSave }) {
  const [form, setForm] = useState(draft);
  const [dropOver, setDropOver] = useState(false);

  async function onFileChange(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    const imageUrl = await readImageFile(file);
    setForm((prev) => ({ ...prev, imageUrl, _imageFile: file }));
  }

  async function onImageDrop(e) {
    e.preventDefault();
    setDropOver(false);
    const file = [...(e.dataTransfer?.files || [])].find((f) => f.type.startsWith("image/"));
    if (!file) return;
    const imageUrl = await readImageFile(file);
    setForm((prev) => ({ ...prev, imageUrl, _imageFile: file }));
  }

  return (
    <div className="modal-backdrop" onClick={onCancel}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <div className="modal-head">
          <h2>Editor Opera</h2>
          <button onClick={onCancel}>✕</button>
        </div>
        <div className="work-editor-grid">
          <div
            className={`image-upload ${dropOver ? "drop-over" : ""}`}
            onDragOver={(e) => {
              const hasImage = [...(e.dataTransfer?.items || [])].some((it) => it.type?.startsWith("image/"));
              if (!hasImage) return;
              e.preventDefault();
              setDropOver(true);
            }}
            onDragLeave={(e) => {
              if (e.currentTarget.contains(e.relatedTarget)) return;
              setDropOver(false);
            }}
            onDrop={onImageDrop}
          >
            {form.imageUrl ? <img src={form.imageUrl} alt={form.title || "Anteprima"} /> : <div className="img-ph lg">Trascina o scegli immagine</div>}
            <label className="small-btn">
              Trascina immagine qui o scegli file
              <input type="file" hidden accept="image/*" onChange={onFileChange} />
            </label>
          </div>
          <div className="work-fields">
            {WORK_FIELDS.map(([key, label]) => (
              <label key={key} className={key === "notes" ? "field-span-2" : ""}>
                {label}
                {key === "notes" ? (
                  <textarea rows={3} value={form[key] || ""} onChange={(e) => setForm((prev) => ({ ...prev, [key]: e.target.value }))} />
                ) : key === "dpi" ? (
                  <input
                    type="number"
                    min={1}
                    step={1}
                    value={form[key] ?? DEFAULT_WORK_DPI}
                    onChange={(e) => setForm((prev) => ({ ...prev, [key]: e.target.value }))}
                  />
                ) : (
                  <input value={form[key] || ""} onChange={(e) => setForm((prev) => ({ ...prev, [key]: e.target.value }))} />
                )}
              </label>
            ))}
          </div>
        </div>
        <div className="modal-foot">
          <button className="ghost-btn" onClick={onCancel}>Annulla</button>
          <button className="primary-btn" onClick={() => onSave(form)}>Salva opera</button>
        </div>
      </div>
    </div>
  );
}

function PageCanvasFrame({ pageFormat, children }) {
  const format = getPageFormat(pageFormat);
  const rulerMarksX = Array.from({ length: Math.floor(format.width / 10) + 1 }, (_, i) => i);
  const rulerMarksY = Array.from({ length: Math.floor(format.height / 10) + 1 }, (_, i) => i);
  return (
    <div className="page-frame">
      <div className="ruler-corner" aria-hidden="true">cm</div>
      <div className="page-ruler page-ruler-top" aria-hidden="true">
        {rulerMarksX.map((cm) => {
          const leftPct = (cm * 10 * 100) / format.width;
          return (
            <div key={`rx_${cm}`} className="ruler-mark major" style={{ left: `${leftPct}%` }}>
              <span className="ruler-label">{cm}</span>
            </div>
          );
        })}
      </div>
      <div className="page-ruler page-ruler-left" aria-hidden="true">
        {rulerMarksY.map((cm) => {
          const topPct = (cm * 10 * 100) / format.height;
          return (
            <div key={`ry_${cm}`} className="ruler-mark major" style={{ top: `${topPct}%` }}>
              <span className="ruler-label">{cm}</span>
            </div>
          );
        })}
      </div>
      <div className="page-frame-canvas">{children}</div>
    </div>
  );
}

function BookSpread({
  spread,
  works,
  theme,
  activePageId,
  selectedElement,
  onSelectPage,
  onSelectElement,
  onMoveElement,
  onDropWorkToPage,
  onMovePlacementToPage,
  layoutAssist,
  pageFormat,
  onDeleteElementDirect,
  view,
  onViewChange,
  interactive = true,
  onPageMetrics,
}) {
  const viewportRef = useRef(null);
  const panRef = useRef(null);

  function onWheelZoom(e) {
    if (!interactive) return;
    if (!(e.ctrlKey || e.metaKey || e.altKey)) return;
    e.preventDefault();
    const delta = e.deltaY > 0 ? -0.08 : 0.08;
    onViewChange((prev) => {
      const nextZoom = Math.max(0.35, Math.min(3, +(prev.zoom + delta).toFixed(2)));
      return { ...prev, zoom: nextZoom };
    });
  }

  function onViewportPointerDown(e) {
    if (!interactive) return;
    const wantsPan = e.shiftKey || e.button === 1 || e.target === viewportRef.current;
    if (!wantsPan) return;
    e.preventDefault();
    panRef.current = {
      x: e.clientX,
      y: e.clientY,
      startPanX: view?.panX || 0,
      startPanY: view?.panY || 0,
    };
    window.addEventListener("pointermove", onViewportPointerMove);
    window.addEventListener("pointerup", onViewportPointerUp);
  }

  function onViewportPointerMove(e) {
    const pan = panRef.current;
    if (!pan) return;
    const dx = e.clientX - pan.x;
    const dy = e.clientY - pan.y;
    onViewChange((prev) => ({ ...prev, panX: pan.startPanX + dx, panY: pan.startPanY + dy }));
  }

  function onViewportPointerUp() {
    panRef.current = null;
    window.removeEventListener("pointermove", onViewportPointerMove);
    window.removeEventListener("pointerup", onViewportPointerUp);
  }

  useEffect(() => () => onViewportPointerUp(), []);

  const leftPage = spread[0] || null;
  const rightPage = spread[1] || null;

  return (
    <div
      ref={viewportRef}
      className={`book-viewport ${interactive ? "interactive" : "print-mode"}`}
      onWheel={onWheelZoom}
      onPointerDown={onViewportPointerDown}
      title={interactive ? "Alt/Ctrl+rotella per zoom, Shift+drag sullo sfondo per pan, Alt durante drag per snap off" : undefined}
    >
      <div
        className="book-transform"
        style={{
          transform: interactive
            ? `translate(${Math.round(view?.panX || 0)}px, ${Math.round(view?.panY || 0)}px) scale(${view?.zoom || 1})`
            : "none",
        }}
      >
        <div className="book-shell">
          {leftPage ? (
            <PageCanvasFrame pageFormat={pageFormat}>
              <PageCanvas
                key={leftPage.id}
                page={leftPage}
                side="left"
                works={works}
                theme={theme}
                active={activePageId === leftPage.id}
                selectedElement={selectedElement}
                onSelectPage={() => onSelectPage(leftPage)}
                onSelectElement={onSelectElement}
                onMoveElement={onMoveElement}
                onDropWorkToPage={onDropWorkToPage}
                onMovePlacementToPage={onMovePlacementToPage}
                layoutAssist={layoutAssist}
                pageFormat={pageFormat}
                onDeleteElementDirect={onDeleteElementDirect}
                zoomScale={interactive ? view?.zoom || 1 : 1}
                onPageMetrics={onPageMetrics}
              />
            </PageCanvasFrame>
          ) : (
            <div className="page-canvas empty" />
          )}
          <div className="book-spine" />
          {rightPage ? (
            <PageCanvasFrame pageFormat={pageFormat}>
              <PageCanvas
                key={rightPage.id}
                page={rightPage}
                side="right"
                works={works}
                theme={theme}
                active={activePageId === rightPage.id}
                selectedElement={selectedElement}
                onSelectPage={() => onSelectPage(rightPage)}
                onSelectElement={onSelectElement}
                onMoveElement={onMoveElement}
                onDropWorkToPage={onDropWorkToPage}
                onMovePlacementToPage={onMovePlacementToPage}
                layoutAssist={layoutAssist}
                pageFormat={pageFormat}
                onDeleteElementDirect={onDeleteElementDirect}
                zoomScale={interactive ? view?.zoom || 1 : 1}
                onPageMetrics={onPageMetrics}
              />
            </PageCanvasFrame>
          ) : (
            <div className="page-canvas empty" />
          )}
        </div>
      </div>
    </div>
  );
}

function PageCanvas({
  page,
  side,
  works,
  theme,
  active,
  selectedElement,
  onSelectPage,
  onSelectElement,
  onMoveElement,
  onDropWorkToPage,
  onMovePlacementToPage,
  layoutAssist,
  pageFormat,
  onDeleteElementDirect,
  zoomScale = 1,
  onPageMetrics,
}) {
  const pageRef = useRef(null);
  const innerRef = useRef(null);
  const dragRef = useRef(null);
  const [dragOver, setDragOver] = useState(false);
  const [guides, setGuides] = useState([]);
  const [inlineEdit, setInlineEdit] = useState(null);
  const [resizeLock, setResizeLock] = useState(null);

  function startDrag(e, kind, elementId, coords, handle = "main", size = null) {
    e.stopPropagation();
    if (kind === "placement" && handle === "main" && (e.altKey || e.shiftKey)) {
      dragRef.current = {
        mode: e.altKey ? "image-pan" : "image-zoom",
        kind,
        elementId,
        handle,
        originX: e.clientX,
        originY: e.clientY,
        start: coords,
        size,
        lifted: false,
      };
      onSelectPage();
      onSelectElement({ pageId: page.id, kind, elementId });
      window.addEventListener("pointermove", onPointerMove);
      window.addEventListener("pointerup", onPointerUp);
      return;
    }
    const rect = pageRef.current?.getBoundingClientRect();
    if (!rect) return;
    dragRef.current = {
      mode: "move",
      kind,
      elementId,
      handle,
      originX: e.clientX,
      originY: e.clientY,
      start: coords,
      size,
      lifted: false,
    };
    onSelectPage();
    onSelectElement({ pageId: page.id, kind, elementId });
    window.addEventListener("pointermove", onPointerMove);
    window.addEventListener("pointerup", onPointerUp);
  }

  function startResize(e, kind, elementId, size, handle = "main", anchor = { x: 0, y: 0 }) {
    e.preventDefault();
    e.stopPropagation();
    setResizeLock({ kind, elementId, handle });
    dragRef.current = {
      mode: "resize",
      kind,
      elementId,
      handle,
      originX: e.clientX,
      originY: e.clientY,
      start: size,
      anchor,
      lifted: false,
    };
    onSelectPage();
    onSelectElement({ pageId: page.id, kind, elementId });
    window.addEventListener("pointermove", onPointerMove);
    window.addEventListener("pointerup", onPointerUp);
  }

  function onPointerMove(e) {
    const drag = dragRef.current;
    if (!drag) return;
    const innerRect = innerRef.current?.getBoundingClientRect();
    const scale = zoomScale || 1;
    const width = (innerRect?.width || 0) / scale;
    const height = (innerRect?.height || 0) / scale;
    const constraintBox = getRenderedConstraintBox(page, width, height, layoutAssist?.boundMode || "margins");
    const altSnapOff = e.altKey;
    const isPlacementFreeDrag = drag.kind === "placement" && drag.handle === "main" && drag.mode === "move" && !e.shiftKey;
    const isCaptionFreeDrag = drag.kind === "placement" && drag.handle === "caption" && drag.mode === "move" && !e.shiftKey;
    const gridSize = Math.max(4, layoutAssist?.gridSize || 12);
    const baseThreshold = layoutAssist?.snapThreshold ?? 8;
    const threshold = isPlacementFreeDrag || isCaptionFreeDrag || altSnapOff ? 0 : Math.max(1, baseThreshold / scale);
    const snapToGrid = isPlacementFreeDrag || isCaptionFreeDrag || altSnapOff ? false : !!layoutAssist?.snapToGrid;
    const showGuides = isPlacementFreeDrag || isCaptionFreeDrag || altSnapOff ? false : !!layoutAssist?.showGuides;

    const guideTargetsX = [
      0,
      width / 2,
      width,
      page.margins?.left ?? 0,
      width - (page.margins?.right ?? 0),
    ];
    const guideTargetsY = [
      0,
      height / 2,
      height,
      page.margins?.top ?? 0,
      height - (page.margins?.bottom ?? 0),
    ];
    const nextGuides = [];

    const snapAxis = (value, targets, axis) => {
      let snapped = value;
      if (snapToGrid) {
        snapped = Math.round(snapped / gridSize) * gridSize;
      }
      let nearest = null;
      for (const t of targets) {
        const dist = Math.abs(snapped - t);
        if (dist <= threshold && (!nearest || dist < nearest.dist)) nearest = { t, dist };
      }
      if (nearest) {
        snapped = Math.round(nearest.t);
        if (showGuides) nextGuides.push({ axis, value: Math.round(nearest.t) });
      }
      return Math.round(snapped);
    };

    const dx = Math.round((e.clientX - drag.originX) / scale);
    const dy = Math.round((e.clientY - drag.originY) / scale);
    if (!drag.lifted && (Math.abs(dx) > 1 || Math.abs(dy) > 1)) drag.lifted = true;
    if (drag.mode === "image-pan") {
      onMoveElement(page.id, "placement", drag.elementId, {
        imageOffsetX: Math.round((drag.start?.imageOffsetX ?? 0) + dx),
        imageOffsetY: Math.round((drag.start?.imageOffsetY ?? 0) + dy),
      });
      return;
    }
    if (drag.mode === "image-zoom") {
      const delta = (drag.originY - e.clientY) / 180;
      const baseScale = drag.start?.imageScale ?? 1;
      const nextScale = clampNum(Number((baseScale + delta).toFixed(3)), 1, 6);
      onMoveElement(page.id, "placement", drag.elementId, { imageScale: nextScale });
      return;
    }
    if (drag.mode === "resize") {
      const maxW = Math.max(40, constraintBox.maxWidth - Math.max(0, (drag.anchor?.x ?? 0) - constraintBox.minX));
      const maxH = Math.max(28, constraintBox.maxHeight - Math.max(0, (drag.anchor?.y ?? 0) - constraintBox.minY));
      if (drag.kind === "placement" && drag.handle === "caption") {
        const nextW = clampNum(snapAxis(drag.start.w + dx, guideTargetsX, "x"), 40, maxW);
        const nextH = clampNum(snapAxis(drag.start.h + dy, guideTargetsY, "y"), 28, maxH);
        onMoveElement(page.id, drag.kind, drag.elementId, { captionW: nextW, captionH: nextH });
      } else if (drag.kind === "placement") {
        const startW = Math.max(40, Number(drag.start?.w) || 40);
        const startH = Math.max(40, Number(drag.start?.h) || 40);
        const ratio = Math.max(0.01, startW / startH);
        const useWidthDriver = Math.abs(dx / startW) >= Math.abs(dy / startH);

        let nextW;
        let nextH;
        if (useWidthDriver) {
          nextW = clampNum(snapAxis(startW + dx, guideTargetsX, "x"), 40, maxW);
          nextH = nextW / ratio;
        } else {
          nextH = clampNum(snapAxis(startH + dy, guideTargetsY, "y"), 40, maxH);
          nextW = nextH * ratio;
        }

        if (nextH > maxH) {
          nextH = maxH;
          nextW = nextH * ratio;
        }
        if (nextW > maxW) {
          nextW = maxW;
          nextH = nextW / ratio;
        }
        if (nextH < 40) {
          nextH = 40;
          nextW = nextH * ratio;
        }
        if (nextW < 40) {
          nextW = 40;
          nextH = nextW / ratio;
        }

        onMoveElement(page.id, drag.kind, drag.elementId, { w: Math.round(nextW), h: Math.round(nextH) });
      } else {
        const nextW = clampNum(snapAxis(drag.start.w + dx, guideTargetsX, "x"), 40, maxW);
        const nextH = clampNum(snapAxis(drag.start.h + dy, guideTargetsY, "y"), 28, maxH);
        onMoveElement(page.id, drag.kind, drag.elementId, { w: nextW, h: nextH });
      }
    } else {
      const draggedW = drag.size?.w ?? 0;
      const draggedH = drag.size?.h ?? 0;

      const enableElementSnap = drag.kind === "placement" && drag.handle === "main" ? !!e.shiftKey : true;
      const elementTargetsX = [];
      const elementTargetsY = [];
      if (enableElementSnap) {
        for (const plc of page.placements || []) {
          if (drag.kind === "placement" && plc.id === drag.elementId) continue;
          elementTargetsX.push(plc.x, plc.x + plc.w / 2, plc.x + plc.w);
          elementTargetsY.push(plc.y, plc.y + plc.h / 2, plc.y + plc.h);
          if (plc.showCaption) {
            const cx = plc.captionX ?? plc.x;
            const cy = plc.captionY ?? plc.y + plc.h + 8;
            const cw = plc.captionW ?? 220;
            const ch = plc.captionH ?? 70;
            elementTargetsX.push(cx, cx + cw / 2, cx + cw);
            elementTargetsY.push(cy, cy + ch / 2, cy + ch);
          }
        }
        for (const txt of page.textBlocks || []) {
          if (drag.kind === "text" && txt.id === drag.elementId) continue;
          elementTargetsX.push(txt.x, txt.x + txt.w / 2, txt.x + txt.w);
          elementTargetsY.push(txt.y, txt.y + txt.h / 2, txt.y + txt.h);
        }
      }

      const snapMoveAxis = (base, size, targets, axis, max) => {
        let candidateBase = snapToGrid ? Math.round(base / gridSize) * gridSize : base;
        let best = null;
        const offsets = [0, size / 2, size];
        for (const t of targets) {
          for (const off of offsets) {
            const cand = t - off;
            const dist = Math.abs(cand - candidateBase);
            if (dist <= threshold && (!best || dist < best.dist)) {
              best = { cand, dist, t };
            }
          }
        }
        if (best) {
          candidateBase = best.cand;
          if (showGuides) nextGuides.push({ axis, value: Math.round(best.t) });
        }
        const min = axis === "x" ? constraintBox.minX : constraintBox.minY;
        const maxPos = axis === "x" ? constraintBox.maxXForWidth(size) : constraintBox.maxYForHeight(size);
        const clamped = clampNum(Math.round(candidateBase), min, Math.max(min, Math.round(maxPos)));
        return clamped;
      };

      const nextX = snapMoveAxis(drag.start.x + dx, draggedW, [...guideTargetsX, ...elementTargetsX], "x", width || Infinity);
      const nextY = snapMoveAxis(drag.start.y + dy, draggedH, [...guideTargetsY, ...elementTargetsY], "y", height || Infinity);
      if (drag.kind === "placement" && drag.handle === "caption") {
        onMoveElement(page.id, drag.kind, drag.elementId, { captionX: nextX, captionY: nextY });
      } else {
        onMoveElement(page.id, drag.kind, drag.elementId, { x: nextX, y: nextY });
      }
    }
    setGuides(nextGuides);
  }

  function onPointerUp() {
    dragRef.current = null;
    setResizeLock(null);
    setGuides([]);
    window.removeEventListener("pointermove", onPointerMove);
    window.removeEventListener("pointerup", onPointerUp);
  }

  useEffect(() => () => onPointerUp(), []);
  useEffect(() => {
    setInlineEdit(null);
  }, [page.id]);
  useEffect(() => {
    const el = innerRef.current;
    const pageEl = pageRef.current;
    if (!el || !pageEl || !onPageMetrics) return;
    const emit = () => {
      const rect = el.getBoundingClientRect();
      const pageRect = pageEl.getBoundingClientRect();
      const scale = zoomScale || 1;
      onPageMetrics(page.id, {
        width: Math.round(rect.width / scale),
        height: Math.round(rect.height / scale),
        pageWidth: Math.round(pageRect.width / scale),
        pageHeight: Math.round(pageRect.height / scale),
      });
    };
    emit();
    if (typeof ResizeObserver === "undefined") return;
    const ro = new ResizeObserver(() => emit());
    ro.observe(el);
    return () => ro.disconnect();
  }, [page.id, onPageMetrics, zoomScale, page.margins?.top, page.margins?.right, page.margins?.bottom, page.margins?.left]);

  const marginStyle = {
    inset: `${page.margins?.top ?? 20}px ${page.margins?.right ?? 20}px ${page.margins?.bottom ?? 30}px ${page.margins?.left ?? 20}px`,
  };
  const format = getPageFormat(pageFormat);
  const topPlacementZ =
    (page.placements || []).reduce((maxZ, p) => {
      const z = Number(p?.zIndex);
      return Number.isFinite(z) ? Math.max(maxZ, z) : maxZ;
    }, 0) + 2000;

  function handleWorkDragOver(e) {
    if (page.type === "summary") return;
    const types = e.dataTransfer?.types || [];
    if (!types.includes("application/x-catalog-work-id") && !types.includes("application/x-catalog-placement")) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = types.includes("application/x-catalog-placement") ? "move" : "copy";
    if (!dragOver) setDragOver(true);
  }

  function handleWorkDragLeave() {
    if (dragOver) setDragOver(false);
  }

  function handleWorkDrop(e) {
    if (page.type === "summary") return;
    const placementPayloadRaw = e.dataTransfer?.getData("application/x-catalog-placement");
    const workId = e.dataTransfer?.getData("application/x-catalog-work-id");
    if (!placementPayloadRaw && !workId) return;
    e.preventDefault();
    setDragOver(false);
    const rect = innerRef.current?.getBoundingClientRect();
    if (!rect) return;
    const innerW = rect.width / (zoomScale || 1);
    const innerH = rect.height / (zoomScale || 1);
    const box = getRenderedConstraintBox(page, innerW, innerH, layoutAssist?.boundMode || "margins");
    const defaultW = 150;
    const defaultH = 190;
    const x = clampNum(
      Math.round((e.clientX - rect.left) / (zoomScale || 1) - 75),
      box.minX,
      Math.max(box.minX, box.maxXForWidth(defaultW)),
    );
    const y = clampNum(
      Math.round((e.clientY - rect.top) / (zoomScale || 1) - 95),
      box.minY,
      Math.max(box.minY, box.maxYForHeight(defaultH)),
    );
    if (placementPayloadRaw) {
      try {
        const payload = JSON.parse(placementPayloadRaw);
        if (payload?.sourcePageId && payload?.placementId) {
          onMovePlacementToPage?.(payload.sourcePageId, payload.placementId, page.id, x, y);
          return;
        }
      } catch {
        // ignore malformed payload
      }
    }
    onDropWorkToPage?.(page.id, workId, x, y);
  }

  function startInlineTextEdit(txt) {
    setInlineEdit({ kind: "text", id: txt.id, value: txt.text || "" });
    onSelectPage();
    onSelectElement({ pageId: page.id, kind: "text", elementId: txt.id });
  }

  function startInlineCaptionEdit(placement, work) {
    setInlineEdit({ kind: "caption", id: placement.id, value: defaultPlacementCaption(placement, work) });
    onSelectPage();
    onSelectElement({ pageId: page.id, kind: "placement", elementId: placement.id });
  }

  function commitInlineTextEdit() {
    if (!inlineEdit) return;
    if (inlineEdit.kind === "caption") {
      onMoveElement(page.id, "placement", inlineEdit.id, { captionOverride: inlineEdit.value });
    } else {
      onMoveElement(page.id, "text", inlineEdit.id, { text: inlineEdit.value });
    }
    setInlineEdit(null);
  }

  function cancelInlineTextEdit() {
    setInlineEdit(null);
  }

  async function maximizePlacementToBounds(placement, work) {
    const innerRect = innerRef.current?.getBoundingClientRect();
    const scale = zoomScale || 1;
    const innerW = (innerRect?.width || 0) / scale;
    const innerH = (innerRect?.height || 0) / scale;
    const box = getRenderedConstraintBox(page, innerW, innerH, layoutAssist?.boundMode || "margins");
    const imageDimensions = await readImageDimensions(work?.imageUrl);
    const fallbackAspect =
      Number(imageDimensions?.width) > 0 && Number(imageDimensions?.height) > 0
        ? Number(imageDimensions.width) / Math.max(1, Number(imageDimensions.height))
        : Math.max(0.01, Number(placement?.w || 1) / Math.max(1, Number(placement?.h || 1)));
    const fitted = fitImageMmToBox(work, imageDimensions, box.maxWidth, box.maxHeight, fallbackAspect);
    onMoveElement(page.id, "placement", placement.id, {
      x: Math.round(box.minX),
      y: Math.round(box.minY),
      w: Math.round(fitted.w),
      h: Math.round(fitted.h),
      imageOffsetX: 0,
      imageOffsetY: 0,
      imageScale: 1,
    });
  }

  return (
    <article
      ref={pageRef}
      className={`page-canvas ${active ? "active" : ""} ${page.type} ${dragOver ? "drag-over-work" : ""}`}
      onClick={onSelectPage}
      onDragOver={handleWorkDragOver}
      onDragLeave={handleWorkDragLeave}
      onDrop={handleWorkDrop}
      style={{
        background: page.bgColor || "#f7f3ea",
        color: theme.textColor,
        fontFamily: theme.fontFamily,
        fontSize: `${theme.bodyFontSize}px`,
        fontWeight: theme.fontWeight,
        aspectRatio: `${format.width} / ${format.height}`,
      }}
    >
      <div ref={innerRef} className={`page-inner ${side}`} style={marginStyle}>
        {!!guides.length &&
          guides.map((g, idx) => (
            <div
              key={`${g.axis}_${g.value}_${idx}`}
              className={`snap-guide ${g.axis}`}
              style={g.axis === "x" ? { left: g.value } : { top: g.value }}
            />
          ))}
        {page.type === "summary" ? (
          <SummaryPage page={page} theme={theme} />
        ) : (
          <>
            {[...(page.placements || [])]
              .map((placement, orderIdx) => ({ placement, orderIdx }))
              .sort(
                (a, b) =>
                  (Number(a.placement.zIndex) || 0) - (Number(b.placement.zIndex) || 0) || a.orderIdx - b.orderIdx,
              )
              .map(({ placement }) => {
              const work = works.find((w) => w.id === placement.workId);
              const isSelected =
                selectedElement?.pageId === page.id &&
                selectedElement?.kind === "placement" &&
                selectedElement?.elementId === placement.id;
              if (!work) return null;
              return (
                <div key={placement.id} style={{ position: "absolute", inset: 0, pointerEvents: "none" }}>
                  {(() => {
                    const layerZ = Number.isFinite(Number(placement.zIndex)) ? Number(placement.zIndex) : 100;
                    const lockActive = !!resizeLock;
                    const placementLockedOut =
                      lockActive && !(resizeLock.kind === "placement" && resizeLock.elementId === placement.id && resizeLock.handle === "main");
                    const captionLockedOut =
                      lockActive && !(resizeLock.kind === "placement" && resizeLock.elementId === placement.id && resizeLock.handle === "caption");
                    const artBorderPx = borderPxFromPercent(placement.w, placement.borderWidthPct ?? 5, 0, Math.max(64, placement.w));
                    const capW = placement.captionW ?? 220;
                    const capBorderPx = borderPxFromPercent(capW, placement.captionBorderWidthPct ?? 5, 0, Math.max(48, capW));
                    return (
                      <>
                  <div
                    className={`draggable-box artwork-box ${isSelected ? "selected" : ""}`}
                    style={{
                      left: placement.x,
                      top: placement.y,
                      width: placement.w,
                      height: placement.h,
                      zIndex: isSelected ? layerZ + 1000 : layerZ,
                      pointerEvents: placementLockedOut ? "none" : "auto",
                      border: `${artBorderPx}px solid ${placement.borderColor || "#ffffff"}`,
                    }}
                    onPointerDown={(e) =>
                      startDrag(
                        e,
                        "placement",
                        placement.id,
                        {
                          x: placement.x,
                          y: placement.y,
                          imageOffsetX: placement.imageOffsetX ?? 0,
                          imageOffsetY: placement.imageOffsetY ?? 0,
                          imageScale: placement.imageScale ?? 1,
                        },
                        "main",
                        {
                          w: placement.w,
                          h: placement.h,
                        },
                      )
                    }
                    onClick={(e) => {
                      e.stopPropagation();
                      onSelectPage();
                      onSelectElement({ pageId: page.id, kind: "placement", elementId: placement.id });
                    }}
                    onDoubleClick={(e) => {
                      e.stopPropagation();
                      onSelectPage();
                      onSelectElement({ pageId: page.id, kind: "placement", elementId: placement.id });
                      maximizePlacementToBounds(placement, work);
                    }}
                  >
                    {work.imageUrl ? (
                      <img
                        src={work.imageUrl}
                        alt={workLabel(work)}
                        draggable={false}
                        style={{
                          position: "absolute",
                          left: "50%",
                          top: "50%",
                          width: `${Math.max(1, placement.imageScale ?? 1) * 100}%`,
                          height: `${Math.max(1, placement.imageScale ?? 1) * 100}%`,
                          objectPosition: `calc(50% + ${Math.round(placement.imageOffsetX ?? 0)}px) calc(50% + ${Math.round(
                            placement.imageOffsetY ?? 0,
                          )}px)`,
                          transform: "translate(-50%, -50%)",
                          transformOrigin: "center center",
                        }}
                      />
                    ) : (
                      <div className="img-ph">Nessuna immagine</div>
                    )}
                    {isSelected && (
                      <button
                        className="element-delete-btn"
                        onClick={(e) => {
                          e.stopPropagation();
                          onDeleteElementDirect?.({ pageId: page.id, kind: "placement", elementId: placement.id });
                        }}
                        title="Rimuovi elemento"
                      >
                        ×
                      </button>
                    )}
                    {isSelected && (
                      <span
                        className="element-transfer-btn"
                        draggable
                        title="Trascina su un'altra pagina"
                        onPointerDown={(e) => e.stopPropagation()}
                        onDragStart={(e) => {
                          e.stopPropagation();
                          e.dataTransfer.setData(
                            "application/x-catalog-placement",
                            JSON.stringify({ sourcePageId: page.id, placementId: placement.id }),
                          );
                          e.dataTransfer.effectAllowed = "move";
                        }}
                      >
                        ⇄
                      </span>
                    )}
                    <span
                      className="resize-handle"
                      onPointerDown={(e) =>
                        startResize(
                          e,
                          "placement",
                          placement.id,
                          { w: placement.w, h: placement.h },
                          "main",
                          { x: placement.x, y: placement.y },
                        )
                      }
                    />
                  </div>
                  {placement.showCaption && (
                    <div
                      className={`draggable-box caption-box ${isSelected ? "selected" : ""}`}
                      style={{
                        left: placement.captionX ?? placement.x,
                        top: placement.captionY ?? placement.y + placement.h + 8,
                        width: placement.captionW ?? 220,
                        height: placement.captionH ?? 70,
                        zIndex: isSelected ? layerZ + 1001 : layerZ + 1,
                        pointerEvents: captionLockedOut ? "none" : "auto",
                        border: `${capBorderPx}px solid ${placement.captionBorderColor || "#ffffff"}`,
                      }}
                      onPointerDown={(e) =>
                        (e.target.closest("a")
                          ? null
                          : startDrag(
                              e,
                              "placement",
                              placement.id,
                              {
                                x: placement.captionX ?? placement.x,
                                y: placement.captionY ?? placement.y + placement.h + 8,
                              },
                              "caption",
                              { w: placement.captionW ?? 220, h: placement.captionH ?? 70 },
                            ))
                      }
                      onDoubleClick={(e) => {
                        e.stopPropagation();
                        startInlineCaptionEdit(placement, work);
                      }}
                      onClick={(e) => {
                        e.stopPropagation();
                        onSelectPage();
                        onSelectElement({ pageId: page.id, kind: "placement", elementId: placement.id });
                      }}
                    >
                      {isSelected && (
                        <button
                          className="element-delete-btn"
                          onClick={(e) => {
                            e.stopPropagation();
                            onMoveElement?.(page.id, "placement", placement.id, { showCaption: false });
                          }}
                          title="Nascondi didascalia"
                        >
                          ×
                        </button>
                      )}
                      {inlineEdit?.kind === "caption" && inlineEdit?.id === placement.id ? (
                        <textarea
                          className="inline-text-editor caption-inline-editor"
                          value={inlineEdit.value}
                          autoFocus
                          onPointerDown={(e) => e.stopPropagation()}
                          onClick={(e) => e.stopPropagation()}
                          onChange={(e) => setInlineEdit((prev) => (prev ? { ...prev, value: e.target.value } : prev))}
                          onBlur={commitInlineTextEdit}
                          onKeyDown={(e) => {
                            if (e.key === "Escape") {
                              e.preventDefault();
                              cancelInlineTextEdit();
                            }
                            if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
                              e.preventDefault();
                              commitInlineTextEdit();
                            }
                          }}
                        />
                      ) : (
                        <div
                          className="markdown-content caption-markdown"
                          dangerouslySetInnerHTML={{ __html: renderMarkdownToHtml(defaultPlacementCaption(placement, work), { inline: false }) }}
                        />
                      )}
                      <span
                        className="resize-handle"
                        onPointerDown={(e) =>
                          startResize(
                            e,
                            "placement",
                            placement.id,
                            { w: placement.captionW ?? 220, h: placement.captionH ?? 70 },
                            "caption",
                            { x: placement.captionX ?? placement.x, y: placement.captionY ?? placement.y + placement.h + 8 },
                          )
                        }
                      />
                    </div>
                  )}
                      </>
                    );
                  })()}
                </div>
              );
            })}

            {(page.textBlocks || []).map((txt) => {
              const isSelected =
                selectedElement?.pageId === page.id && selectedElement?.kind === "text" && selectedElement?.elementId === txt.id;
              const isEditing = inlineEdit?.kind === "text" && inlineEdit?.id === txt.id;
              return (
                (() => {
                  const textLockedOut =
                    !!resizeLock && !(resizeLock.kind === "text" && resizeLock.elementId === txt.id && resizeLock.handle === "main");
                  const txtBorderPx = borderPxFromPercent(txt.w, txt.borderWidthPct ?? 5, 0, Math.max(48, txt.w));
                  return (
                <div
                  key={txt.id}
                  className={`draggable-box text-box ${isSelected ? "selected" : ""}`}
                  style={{
                    left: txt.x,
                    top: txt.y,
                    width: txt.w,
                    height: txt.h,
                    zIndex: isSelected ? topPlacementZ + 1 : topPlacementZ,
                    pointerEvents: textLockedOut ? "none" : "auto",
                    color: txt.color || theme.textColor,
                    background: txt.bgColor || theme.defaultTextBgColor || "rgba(255, 255, 255, 0.42)",
                    fontSize: txt.fontSize || theme.bodyFontSize,
                    fontWeight: txt.fontWeight || theme.fontWeight,
                    textAlign: txt.align || "left",
                    border: `${txtBorderPx}px solid ${txt.borderColor || "#ffffff"}`,
                  }}
                  onPointerDown={(e) => {
                    if (e.target.closest("a")) return;
                    if (isEditing) {
                      e.stopPropagation();
                      return;
                    }
                    startDrag(e, "text", txt.id, { x: txt.x, y: txt.y }, "main", { w: txt.w, h: txt.h });
                  }}
                  onClick={(e) => {
                    e.stopPropagation();
                    onSelectPage();
                    onSelectElement({ pageId: page.id, kind: "text", elementId: txt.id });
                  }}
                  onDoubleClick={(e) => {
                    e.stopPropagation();
                    startInlineTextEdit(txt);
                  }}
                >
                  {isSelected && (
                    <button
                      className="element-delete-btn"
                      onClick={(e) => {
                        e.stopPropagation();
                        onDeleteElementDirect?.({ pageId: page.id, kind: "text", elementId: txt.id });
                      }}
                      title="Rimuovi testo"
                    >
                      ×
                    </button>
                  )}
                  {isEditing ? (
                    <textarea
                      className="inline-text-editor"
                      value={inlineEdit.value}
                      autoFocus
                      onPointerDown={(e) => e.stopPropagation()}
                      onClick={(e) => e.stopPropagation()}
                      onChange={(e) => setInlineEdit((prev) => (prev ? { ...prev, value: e.target.value } : prev))}
                      onBlur={commitInlineTextEdit}
                      onKeyDown={(e) => {
                        if (e.key === "Escape") {
                          e.preventDefault();
                          cancelInlineTextEdit();
                        }
                        if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
                          e.preventDefault();
                          commitInlineTextEdit();
                        }
                      }}
                    />
                  ) : (
                    <div className="markdown-content" dangerouslySetInnerHTML={{ __html: renderMarkdownToHtml(txt.text, { inline: false }) }} />
                  )}
                  <span
                    className="resize-handle"
                    onPointerDown={(e) => {
                      if (isEditing) {
                        e.stopPropagation();
                        return;
                      }
                      startResize(e, "text", txt.id, { w: txt.w, h: txt.h }, "main", { x: txt.x, y: txt.y });
                    }}
                  />
                </div>
                  );
                })()
              );
            })}
          </>
        )}
      </div>

      {page.showPageNumber && (
        <div className={`page-number ${side}`} style={{ color: page.pageNumberColor || "#6b614f" }}>
          {page.computedPageNumber ?? ""}
        </div>
      )}
    </article>
  );
}

function SummaryPage({ page, theme }) {
  return (
    <div className="summary-page">
      <h4 style={{ fontSize: theme.titleFontSize * 0.7 }}>Elenco Opere</h4>
      <div className="summary-list">
        {(page.summaryItems || []).map((work, idx) => (
          <div key={work.id} className="summary-row">
            <span>{idx + 1}.</span>
            <span>{workLabel(work)}</span>
            <span>{work.author || "-"}</span>
            <span>{work.year || "-"}</span>
            <span>{work.inventory || "-"}</span>
          </div>
        ))}
        {!page.summaryItems?.length && <p className="muted">Nessuna opera inserita.</p>}
      </div>
    </div>
  );
}
