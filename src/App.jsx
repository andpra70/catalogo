import { useEffect, useMemo, useRef, useState } from "react";

const STORAGE_KEY = "catalogo-opere-state-v1";
const IMAGE_DB_NAME = "catalogo-opere-assets";
const IMAGE_STORE_NAME = "images";

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
  ["inventory", "Inventario"],
  ["location", "Collocazione"],
  ["notes", "Note"],
];

const PAGE_FORMATS = [
  { id: "a4-portrait", label: "A4 Verticale", width: 210, height: 297 },
  { id: "square", label: "Quadrato", width: 240, height: 240 },
  { id: "landscape", label: "Orizzontale", width: 297, height: 210 },
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

function createPrefacePage(marginsOverride, bgColor = "#ffffff", borderPct = 3) {
  const page = createPage("page", "Prefazione", marginsOverride);
  page.bgColor = bgColor;
  page.textBlocks = [
    {
      ...createTextBlock("Prefazione"),
      x: 26,
      y: 28,
      w: 300,
      h: 44,
      fontSize: 22,
      fontWeight: 700,
      align: "left",
      borderWidthPct: borderPct,
    },
    {
      ...createTextBlock(
        "Questo catalogo raccoglie una selezione di opere con una sequenza pensata per accompagnare la lettura tra immagini, ritmo di pagina e apparati testuali. La prefazione può essere modificata, ampliata o sostituita con un testo curatoriale completo.",
      ),
      x: 26,
      y: 84,
      w: 300,
      h: 132,
      fontSize: 13,
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

function createDefaultState() {
  const defaultMargins = { top: 28, right: 28, bottom: 38, left: 28 };
  const coverFront = createPage("cover-front", "Copertina", defaultMargins);
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

  const innerPage = createPage("page", "Pagina 1", defaultMargins);
  innerPage.pageNumber = 1;
  innerPage.textBlocks = [{ ...createTextBlock("Introduzione o testo curatoriale"), x: 36, y: 36, w: 290, h: 120 }];

  const coverBack = createPage("cover-back", "Retro copertina", defaultMargins);
  coverBack.showPageNumber = false;
  coverBack.bgColor = "#ffffff";
  coverBack.textBlocks = [
    {
      ...createTextBlock(
        "Sintesi\nUna selezione di opere tra fotografia, pittura e immagini d'autore che esplora materia, luce e memoria in forma di catalogo editoriale.",
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
        "Biografia autore (esempio)\nAndrea Rossi (1987) vive e lavora a Milano. La sua ricerca attraversa fotografia e pratiche miste, con attenzione al rapporto tra archivio, paesaggio e costruzione della memoria visiva. Ha esposto in mostre collettive e progetti indipendenti in Italia e all'estero.",
      ),
      x: 30,
      y: 188,
      w: 300,
      h: 118,
      fontSize: 12,
      fontWeight: 400,
      align: "left",
    },
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
  ];

  return {
    works: [],
    selectedWorkId: null,
    pages: [coverFront, innerPage, coverBack],
    currentSpread: 0,
    activePageId: innerPage.id,
    selectedElement: null,
    theme: {
      fontFamily: FONT_OPTIONS[2].value,
      bodyFontSize: 15,
      titleFontSize: 26,
      fontWeight: 400,
      textColor: "#111111",
      accentColor: "#111111",
      uiTint: "#f5f5f5",
      paperShadow: "rgba(0, 0, 0, 0.10)",
      pageMargins: defaultMargins,
      autoShowCaptionDefault: true,
      defaultPageBgColor: "#ffffff",
      defaultElementBorderPct: 3,
    },
    pageFormat: "a4-portrait",
    summaryPageEdits: {},
    specialPages: {
      insideFront: createInsideFrontCoverPage(defaultMargins),
      insideBack: createInsideBackCoverPage(defaultMargins),
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
  const { imageUrl, _imageFile, ...rest } = work || {};
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
      const { _imageFile, ...rest } = work || {};
      return rest;
    }),
  };
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
      works: (parsed.works || []).map((w) => ({ ...w, imageUrl: "" })),
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

function captionPartsFromText(text) {
  const lines = String(text || "").split("\n");
  return {
    title: lines[0] || "",
    meta: lines.slice(1).join("\n"),
  };
}

function summaryPagesFromWorks(
  works,
  marginsOverride,
  workPageMap = new Map(),
  summaryPageEdits = {},
  defaultBgColor = "#ffffff",
  defaultBorderPct = 3,
) {
  const chunkSize = 14;
  const chunks = [];
  for (let i = 0; i < works.length; i += chunkSize) chunks.push(works.slice(i, i + chunkSize));
  return chunks.map((chunk, idx) => {
    const titleBlock = {
      ...createTextBlock(idx === 0 ? "Elenco opere" : `Elenco opere (${idx + 1})`),
      x: 24,
      y: 18,
      w: 310,
      h: 38,
      fontSize: 18,
      fontWeight: 700,
      align: "left",
      borderWidthPct: defaultBorderPct,
    };
    const listText =
      chunk
        .map((work) => {
          const pageNo = workPageMap.get(work.id);
          const parts = [workLabel(work), work.author || "-", work.year || "-"].filter(Boolean);
          return `${parts[0]} - ${parts[1]} - ${parts[2]}  |  pag. ${pageNo ?? "-"}`;
        })
        .join("\n") || "Nessuna opera inserita.";
    const listBlock = {
      ...createTextBlock(listText),
      x: 24,
      y: 64,
      w: 310,
      h: 208,
      fontSize: 12,
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
      pageNumberColor: "#6b614f",
      margins: marginsOverride || { top: 28, right: 26, bottom: 38, left: 26 },
      textBlocks: [titleBlock, listBlock],
      placements: [],
    };
    const edit = summaryPageEdits?.[generated.id];
    if (!edit) return generated;
    return {
      ...generated,
      ...edit,
      margins: { ...(generated.margins || {}), ...(edit.margins || {}) },
      textBlocks: edit.textBlocks || generated.textBlocks,
      placements: edit.placements || generated.placements,
    };
  });
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
    borderWidthPct: 5,
    borderColor: "#ffffff",
    showCaption: true,
    captionX: x,
    captionY: y + 196,
    captionW: 220,
    captionH: 52,
    captionBorderWidthPct: 5,
    captionBorderColor: "#ffffff",
  };
}

function createAutoWorkPage(work, pageFormatId) {
  const page = createPage("page", workLabel(work));
  page.bgColor = "#fbfbfb";
  const format = getPageFormat(pageFormatId);
  const margins = page.margins;
  const innerW = Math.max(220, format.width - (margins.left + margins.right));
  const innerH = Math.max(260, format.height - (margins.top + margins.bottom));
  const placementW = Math.min(220, Math.max(140, Math.round(innerW * 0.68)));
  const placementH = Math.round(placementW * 1.12);
  const captionH = 50;
  const gap = 8;
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

  // Fit massimo preservando aspect ratio dentro l'area disponibile sopra la didascalia.
  let w = maxImageW;
  let h = w / safeAspect;
  if (h > maxImageH) {
    h = maxImageH;
    w = h * safeAspect;
  }

  w = clampNum(Math.round(w), 60, maxImageW);
  h = clampNum(Math.round(h), 60, maxImageH);

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

function borderPxFromPercent(sizePx, pct, minPx = 0, maxPx = 64) {
  const p = Number.isFinite(Number(pct)) ? Number(pct) : 5;
  return clampNum(Math.round((Math.max(0, sizePx || 0) * p) / 100), minPx, maxPx);
}

function applyThemeAutoDefaultsToPage(page, theme) {
  if (!page) return page;
  const showCaptionDefault = theme?.autoShowCaptionDefault ?? true;
  const bgColorDefault = theme?.defaultPageBgColor || "#ffffff";
  const borderPctDefault = theme?.defaultElementBorderPct ?? 3;
  return {
    ...page,
    bgColor: bgColorDefault,
    placements: (page.placements || []).map((p) => ({
      ...p,
      showCaption: showCaptionDefault,
      borderWidthPct: borderPctDefault,
      captionBorderWidthPct: borderPctDefault,
    })),
    textBlocks: (page.textBlocks || []).map((t) => ({
      ...t,
      borderWidthPct: t.borderWidthPct ?? borderPctDefault,
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

  function patchState(updater) {
    setState((prev) => (typeof updater === "function" ? updater(prev) : { ...prev, ...updater }));
  }

  function patchTheme(patch) {
    patchState((prev) => {
      const nextTheme = { ...prev.theme, ...patch };
      const hasBgPatch = Object.prototype.hasOwnProperty.call(patch, "defaultPageBgColor");
      const hasBorderPatch = Object.prototype.hasOwnProperty.call(patch, "defaultElementBorderPct");
      if (!hasBgPatch && !hasBorderPatch) {
        return { ...prev, theme: nextTheme };
      }
      const nextBg = patch.defaultPageBgColor || prev.theme?.defaultPageBgColor || "#ffffff";
      const nextBorderPct = Number.isFinite(Number(patch.defaultElementBorderPct))
        ? Number(patch.defaultElementBorderPct)
        : prev.theme?.defaultElementBorderPct ?? 3;
      const patchPageElementsBorders = (page) => ({
        ...page,
        bgColor: hasBgPatch ? nextBg : page.bgColor,
        placements: (page.placements || []).map((pl) => ({
          ...pl,
          borderWidthPct: hasBorderPatch ? nextBorderPct : pl.borderWidthPct,
          captionBorderWidthPct: hasBorderPatch ? nextBorderPct : pl.captionBorderWidthPct,
        })),
        textBlocks: (page.textBlocks || []).map((tx) => ({
          ...tx,
          borderWidthPct: hasBorderPatch ? nextBorderPct : tx.borderWidthPct,
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
    const incoming = { ...draft };
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
      let autoPage = createAutoWorkPage(incoming, prev.pageFormat);
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
      openCreateWork({ title: titleFromName, imageUrl, _imageFile: imageFile });
      return;
    }

    const importedWorks = [];
    for (const file of imageFiles) {
      const imageUrl = await readImageFile(file);
      const work = {
        ...cloneWorkDraft(),
        title: file.name.replace(/\.[^.]+$/, ""),
        imageUrl,
        imageId: uid("img"),
      };
      try {
        await idbPutImage(work.imageId, imageUrl);
      } catch {
        // fallback: keep in-memory imageUrl even if IndexedDB fails
      }
      importedWorks.push(work);
    }
    if (!importedWorks.length) return;

    patchState((prev) => {
      const autoPages = importedWorks.map((work) => {
        let page = createAutoWorkPage(work, prev.pageFormat);
        page = applyThemeAutoDefaultsToPage(page, prev.theme);
        page.margins = { ...(prev.theme?.pageMargins || page.margins) };
        return page;
      });
      const nextPages = [...prev.pages];
      nextPages.splice(nextPages.length - 1, 0, ...autoPages);
      const lastWork = importedWorks[importedWorks.length - 1];
      const firstPage = autoPages[0];
      return {
        ...prev,
        works: [...prev.works, ...importedWorks],
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
    block.w = targetW;
    block.h = Math.max(block.h || 0, minH);
    block.x = Math.max(0, Math.round((bounds.width - block.w) / 2));
    block.y = Math.max(0, Math.round(bounds.height * 0.08));
    patchPage(activeEditablePage.id, (page) => ({ ...page, textBlocks: [...page.textBlocks, block] }));
    patchState((prev) => ({ ...prev, selectedElement: { pageId: activeEditablePage.id, kind: "text", elementId: block.id } }));
  }

  function addSelectedWorkToActivePage() {
    if (!activeEditablePage || activeEditablePage.type === "summary" || !selectedWork) return;
    const placement = createPlacementForWork(selectedWork.id, 42, 42);
    patchPage(activeEditablePage.id, (page) => ({ ...page, placements: [...page.placements, placement] }));
    patchState((prev) => ({ ...prev, selectedElement: { pageId: activeEditablePage.id, kind: "placement", elementId: placement.id } }));
  }

  function getAutoLayoutChunkSize(mode) {
    if (mode === "two-cols") return 2;
    if (mode === "grid4") return 4;
    return 1;
  }

  function buildAutoLayoutPageForWorks(worksChunk, pageFormat, margins, dimMap, measuredBounds, mode) {
    const page = createPage("page", worksChunk.length === 1 ? workLabel(worksChunk[0]) : "Pagina opere", margins);
    page.bgColor = "#fbfbfb";
    const bounds = measuredBounds || getPageContentBounds(page, pageFormat);
    const areaW = Math.max(140, bounds.width);
    const areaH = Math.max(180, bounds.height);

    const fitInSlot = (work, slot) => {
      const dims = dimMap.get(work.id);
      const aspect = dims?.width && dims?.height ? dims.width / dims.height : 1;
      const safeAspect = Number.isFinite(aspect) && aspect > 0 ? aspect : 1;
      const title = workLabel(work);
      const meta = [work.author, work.year].filter(Boolean).join(", ");
      const oneLine = [title, meta].filter(Boolean).join(" — ");
      const approxCharsPerLine = Math.max(18, Math.floor(slot.w / 7.2));
      const preferOneLine = oneLine.length <= approxCharsPerLine && safeAspect >= 1;
      const captionOverride = preferOneLine ? oneLine : [title, meta].filter(Boolean).join("\n");
      const captionH = preferOneLine ? 36 : oneLine.length > approxCharsPerLine * 1.6 ? 62 : 52;
      const gap = 8;
      const maxW = Math.max(60, slot.w);
      const maxH = Math.max(60, slot.h - captionH - gap);
      let w = maxW;
      let h = w / safeAspect;
      if (h > maxH) {
        h = maxH;
        w = h * safeAspect;
      }
      w = clampNum(Math.round(w), 60, maxW);
      h = clampNum(Math.round(h), 60, maxH);
      const x = slot.x + Math.round((slot.w - w) / 2);
      const y = slot.y + Math.max(0, Math.round((slot.h - h - captionH - gap) / 2));
      const p = createPlacementForWork(work.id, x, y);
      p.w = w;
      p.h = h;
      p.captionW = Math.min(slot.w, Math.max(160, w));
      p.captionH = captionH;
      p.captionX = slot.x + Math.round((slot.w - p.captionW) / 2);
      p.captionY = clampNum(y + h + gap, slot.y, slot.y + Math.max(0, slot.h - captionH));
      p.captionOverride = captionOverride;
      return p;
    };

    let slots = [];
    if (mode === "two-cols") {
      const gap = 12;
      const w = Math.round((areaW - gap) / 2);
      const h = Math.round(areaH * 0.92);
      const y = Math.max(0, Math.round((areaH - h) / 2));
      slots = [
        { x: 0, y, w, h },
        { x: w + gap, y, w, h },
      ];
    } else if (mode === "grid4") {
      const cols = 2;
      const rows = 2;
      const gapX = 12;
      const gapY = 12;
      const w = Math.round((areaW - gapX) / cols);
      const h = Math.round((areaH - gapY) / rows);
      for (let i = 0; i < 4; i += 1) {
        const col = i % cols;
        const row = Math.floor(i / cols);
        slots.push({ x: col * (w + gapX), y: row * (h + gapY), w, h });
      }
    } else {
      slots = [{ x: 0, y: 0, w: areaW, h: areaH }];
    }

    page.placements = worksChunk.map((work, idx) => fitInSlot(work, slots[idx] || slots[0]));
    return applyThemeAutoDefaultsToPage(page, state.theme);
  }

  async function autoGeneratePagesFromCatalog() {
    const works = state.works || [];
    if (!works.length) return;
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
      const defaultBg = prev.theme?.defaultPageBgColor || "#ffffff";
      const frontBase = prev.pages[0] || createDefaultState().pages[0];
      const backBase = prev.pages[prev.pages.length - 1] || createDefaultState().pages.at(-1);
      const front = { ...frontBase, placements: [], bgColor: defaultBg };
      const back = { ...backBase, placements: [], bgColor: defaultBg };
      const prefacePage = createPrefacePage(prev.theme?.pageMargins, defaultBg, prev.theme?.defaultElementBorderPct ?? 3);
      const chunkSize = getAutoLayoutChunkSize(autoLayoutMode);
      const chunks = [];
      for (let i = 0; i < prev.works.length; i += chunkSize) chunks.push(prev.works.slice(i, i + chunkSize));
      const autoPages = chunks.map((chunk) =>
        buildAutoLayoutPageForWorks(chunk, prev.pageFormat, prev.theme?.pageMargins, dimMap, measuredBounds || undefined, autoLayoutMode),
      );
      const pages = [
        { ...front, margins: { ...(prev.theme?.pageMargins || front.margins) } },
        prefacePage,
        ...(autoPages.length ? autoPages : [createPage("page", "Pagina 1", prev.theme?.pageMargins)]),
        { ...back, margins: { ...(prev.theme?.pageMargins || back.margins) } },
      ];
      return {
        ...prev,
        pages,
        specialPages: {
          insideFront: prev.specialPages?.insideFront
            ? { ...prev.specialPages.insideFront, placements: [] }
            : prev.specialPages?.insideFront,
          insideBack: prev.specialPages?.insideBack
            ? { ...prev.specialPages.insideBack, placements: [] }
            : prev.specialPages?.insideBack,
        },
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

  function addWorkToPageAtPosition(pageId, workId, x, y) {
    const page = allEditablePages.find((p) => p.id === pageId);
    if (!page || page.type === "summary") return;
    const placement = createPlacementForWork(workId, Math.max(0, x), Math.max(0, y));
    patchPage(pageId, (p) => ({ ...p, placements: [...p.placements, placement] }));
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
      current.push(createPlacementForWork(work.id, 20, 20));
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

      if (preset === "hero") {
        placements = ensurePlacementsForLayout(prevPage, 1);
        if (!placements.length) return prevPage;
        const p = { ...placements[0] };
        p.w = Math.round(areaW * 0.82);
        p.h = Math.round(areaH * 0.62);
        p.x = Math.round((areaW - p.w) / 2);
        p.y = Math.max(10, Math.round((areaH - p.h - 60) / 2));
        p.captionW = Math.min(areaW, p.w);
        p.captionH = 50;
        p.captionX = Math.round((areaW - p.captionW) / 2);
        p.captionY = p.y + p.h + 8;
        placements = [p];
      }

      if (preset === "two-cols") {
        placements = ensurePlacementsForLayout(prevPage, 2);
        if (!placements.length) return prevPage;
        const gap = 12;
        const w = Math.round((areaW - gap) / 2);
        const h = Math.round(areaH * 0.55);
        placements = placements.slice(0, 2).map((p, i) => {
          const x = i * (w + gap);
          const y = Math.max(10, Math.round((areaH - h - 56) / 2));
          return {
            ...p,
            x,
            y,
            w,
            h,
            captionX: x,
            captionY: y + h + 8,
            captionW: w,
            captionH: 48,
          };
        });
      }

      if (preset === "grid4") {
        placements = ensurePlacementsForLayout(prevPage, 4);
        if (!placements.length) return prevPage;
        const cols = 2;
        const rows = 2;
        const gapX = 12;
        const gapY = 12;
        const w = Math.round((areaW - gapX) / cols);
        const h = Math.round((areaH - gapY - 2 * 46) / rows);
        placements = placements.slice(0, 4).map((p, i) => {
          const col = i % cols;
          const row = Math.floor(i / cols);
          const x = col * (w + gapX);
          const y = row * (h + gapY + 38);
          return {
            ...p,
            x,
            y,
            w,
            h,
            captionX: x,
            captionY: y + h + 6,
            captionW: w,
            captionH: 36,
          };
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
        placements: page.placements.map((el) => (el.id === elementId ? { ...el, ...patch } : el)),
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

  function exportCatalogJson() {
    const payload = {
      exportedAt: new Date().toISOString(),
      version: 1,
      catalog: sanitizeStateForExport(state),
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `catalogo-opere-${new Date().toISOString().slice(0, 10)}.json`;
    a.click();
    URL.revokeObjectURL(url);
  }

  async function importCatalogJson(file) {
    const text = await file.text();
    const parsed = JSON.parse(text);
    const incoming = parsed.catalog || parsed;
    const importedWorks = await Promise.all(
      (incoming.works || []).map(async (work) => {
        if (!work.imageUrl) return work;
        const imageId = work.imageId || uid("img");
        try {
          await idbPutImage(imageId, work.imageUrl);
        } catch {
          // keep imageUrl in memory even if IDB fails
        }
        return { ...work, imageId };
      }),
    );
    const base = createDefaultState();
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

  function printCatalogPdf() {
    const fmt = getPageFormat(state.pageFormat);
    const styleId = "catalog-print-page-size-style";
    let styleEl = document.getElementById(styleId);
    if (!styleEl) {
      styleEl = document.createElement("style");
      styleEl.id = styleId;
      document.head.appendChild(styleEl);
    }
    styleEl.textContent = `@page { size: ${fmt.width}mm ${fmt.height}mm; margin: 0; }`;
    const cleanup = () => {
      window.removeEventListener("afterprint", cleanup);
      styleEl?.remove();
    };
    window.addEventListener("afterprint", cleanup);
    window.print();
  }

  function resetSummaryPagesOverrides() {
    patchState((prev) => ({ ...prev, summaryPageEdits: {} }));
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
            <h1>Catalogo Opere</h1>
            {persistWarning && <p className="warn-text">{persistWarning}</p>}
          </div>
          <div ref={topbarActionsRef} className="topbar-actions">
            <IconButton icon="help" onClick={() => setHelpOpen((v) => !v)} title="Guida icone" ariaLabel="Guida icone" />
            <IconButton icon="theme" onClick={() => setThemeOpen((v) => !v)} title="Tema" ariaLabel="Tema" />
            <IconButton icon="menu" onClick={() => setTopbarMenuOpen((v) => !v)} title="Menu" ariaLabel="Menu" />
            {topbarMenuOpen && (
              <div className="topbar-overflow">
                <button onClick={exportCatalogJson}>Esporta JSON</button>
                <button onClick={() => importJsonRef.current?.click()}>Importa JSON</button>
                <button onClick={resetSummaryPagesOverrides}>Rigenera elenco opere</button>
                <button onClick={printCatalogPdf}>Stampa / PDF Catalogo</button>
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
                if (curr && curr.width === metrics.width && curr.height === metrics.height) return prev;
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
              <PageInspector
                page={activeEditablePage}
                onChange={(patch) => patchPage(activeEditablePage.id, (p) => ({ ...p, ...patch }))}
                onDeletePage={() => removeInnerPage(activeEditablePage.id)}
              />
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
            {selectedWork ? (
              <div className="selected-work-card">
                {selectedWork.imageUrl ? <img src={selectedWork.imageUrl} alt={workLabel(selectedWork)} /> : <div className="img-ph">Nessuna immagine</div>}
                <div>
                  <strong>{workLabel(selectedWork)}</strong>
                  <p>{selectedWork.author || "Autore non indicato"}</p>
                  <button className="small-btn" onClick={() => openEditWork(selectedWork)}>Modifica opera</button>
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
          const measuredInner = pageMetrics?.[page.id] || getPageContentBounds(page, state.pageFormat);
          const sourceW = Math.max(1, Math.round((measuredInner?.width || 0) + (m.left || 0) + (m.right || 0)));
          const sourceH = Math.max(1, Math.round((measuredInner?.height || 0) + (m.top || 0) + (m.bottom || 0)));
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
          Testo
          <textarea rows={5} value={data.text || ""} onChange={(e) => onChange({ text: e.target.value })} />
        </label>
        <RangeField label="Font size" min={10} max={42} value={data.fontSize || 16} onChange={(v) => onChange({ fontSize: v })} />
        <RangeField label="Peso" min={300} max={800} step={100} value={data.fontWeight || 500} onChange={(v) => onChange({ fontWeight: v })} />
        <label>
          Colore
          <input type="color" value={data.color || "#222222"} onChange={(e) => onChange({ color: e.target.value })} />
        </label>
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
          ) : (
            <div className="page-canvas empty" />
          )}
          <div className="book-spine" />
          {rightPage ? (
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

  function startDrag(e, kind, elementId, coords, handle = "main", size = null) {
    e.stopPropagation();
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
    };
    onSelectPage();
    onSelectElement({ pageId: page.id, kind, elementId });
    window.addEventListener("pointermove", onPointerMove);
    window.addEventListener("pointerup", onPointerUp);
  }

  function startResize(e, kind, elementId, size, handle = "main", anchor = { x: 0, y: 0 }) {
    e.stopPropagation();
    dragRef.current = {
      mode: "resize",
      kind,
      elementId,
      handle,
      originX: e.clientX,
      originY: e.clientY,
      start: size,
      anchor,
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
    const isCaptionFreeDrag = drag.kind === "placement" && drag.handle === "caption" && drag.mode === "move" && !e.shiftKey;
    const gridSize = Math.max(4, layoutAssist?.gridSize || 12);
    const baseThreshold = layoutAssist?.snapThreshold ?? 8;
    const threshold = isCaptionFreeDrag || altSnapOff ? 0 : Math.max(1, baseThreshold / scale);
    const snapToGrid = isCaptionFreeDrag || altSnapOff ? false : !!layoutAssist?.snapToGrid;
    const showGuides = isCaptionFreeDrag || altSnapOff ? false : !!layoutAssist?.showGuides;

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
    if (drag.mode === "resize") {
      const maxW = Math.max(40, constraintBox.maxWidth - Math.max(0, (drag.anchor?.x ?? 0) - constraintBox.minX));
      const maxH = Math.max(28, constraintBox.maxHeight - Math.max(0, (drag.anchor?.y ?? 0) - constraintBox.minY));
      const nextW = clampNum(snapAxis(drag.start.w + dx, guideTargetsX, "x"), 40, maxW);
      const nextH = clampNum(snapAxis(drag.start.h + dy, guideTargetsY, "y"), 28, maxH);
      if (drag.kind === "placement" && drag.handle === "caption") {
        onMoveElement(page.id, drag.kind, drag.elementId, { captionW: nextW, captionH: nextH });
      } else {
        onMoveElement(page.id, drag.kind, drag.elementId, { w: nextW, h: nextH });
      }
    } else {
      const draggedW = drag.size?.w ?? 0;
      const draggedH = drag.size?.h ?? 0;

      const elementTargetsX = [];
      const elementTargetsY = [];
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
    if (!el || !onPageMetrics) return;
    const emit = () => {
      const rect = el.getBoundingClientRect();
      const scale = zoomScale || 1;
      onPageMetrics(page.id, {
        width: Math.round(rect.width / scale),
        height: Math.round(rect.height / scale),
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
            {(page.placements || []).map((placement) => {
              const work = works.find((w) => w.id === placement.workId);
              const isSelected =
                selectedElement?.pageId === page.id &&
                selectedElement?.kind === "placement" &&
                selectedElement?.elementId === placement.id;
              if (!work) return null;
              return (
                <div key={placement.id}>
                  {(() => {
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
                      border: `${artBorderPx}px solid ${placement.borderColor || "#ffffff"}`,
                    }}
                    onPointerDown={(e) =>
                      startDrag(e, "placement", placement.id, { x: placement.x, y: placement.y }, "main", {
                        w: placement.w,
                        h: placement.h,
                      })
                    }
                    onClick={(e) => {
                      e.stopPropagation();
                      onSelectPage();
                      onSelectElement({ pageId: page.id, kind: "placement", elementId: placement.id });
                    }}
                  >
                    {work.imageUrl ? <img src={work.imageUrl} alt={workLabel(work)} draggable={false} /> : <div className="img-ph">Nessuna immagine</div>}
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
                        border: `${capBorderPx}px solid ${placement.captionBorderColor || "#ffffff"}`,
                      }}
                      onPointerDown={(e) =>
                        startDrag(
                          e,
                          "placement",
                          placement.id,
                          {
                            x: placement.captionX ?? placement.x,
                            y: placement.captionY ?? placement.y + placement.h + 8,
                          },
                          "caption",
                          { w: placement.captionW ?? 220, h: placement.captionH ?? 70 },
                        )
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
                        (() => {
                          const parts = captionPartsFromText(defaultPlacementCaption(placement, work));
                          return (
                            <>
                              <strong>{parts.title}</strong>
                              {parts.meta ? <small>{parts.meta}</small> : null}
                            </>
                          );
                        })()
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
                    color: txt.color || theme.textColor,
                    fontSize: txt.fontSize || theme.bodyFontSize,
                    fontWeight: txt.fontWeight || theme.fontWeight,
                    textAlign: txt.align || "left",
                    border: `${txtBorderPx}px solid ${txt.borderColor || "#ffffff"}`,
                  }}
                  onPointerDown={(e) => {
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
                    txt.text
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
