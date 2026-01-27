import { FORMULA_SUGGESTIONS } from "./constants.js";

const defaultCallbacks = {
  isDarkMode: () => false,
  toggleTheme: () => {},
  clearSpreadsheet: () => {},
  recalculateFormulas: () => {},
  updateURL: async () => {},
  showToast: () => {},
  openCsvImport: () => {},
  exportCSV: () => {},
  exportExcel: () => {},
  openJSONModal: () => {},
  openHashTool: () => {},
  copyLink: () => {},
  showQRModal: () => {},
  showEmbedModal: () => {},
  openTemplateGallery: () => {},
  startPresentation: () => {},
  toggleReadOnly: () => {},
  isReadOnly: () => false,
  toggleDependencies: () => {},
  openP2PModal: () => {},
  startP2PHost: () => {},
  focusP2PJoin: () => {},
  copyP2PId: () => {},
  getP2PId: () => "",
  openPasswordModal: () => {},
  openGitHub: () => {},
  applyFormat: () => {},
  toggleZen: () => {},
  isZen: () => false,
};

const VISUAL_FORMULAS = [
  { signature: "PROGRESS(value, total)", description: "Progress bar visual" },
  { signature: "TAG(label)", description: "Colored tag badge" },
  { signature: "RATING(value, max)", description: "Star rating visual" },
];

const SHORTCUTS = [
  { label: "Open command palette", keys: "Ctrl+K / Cmd+K" },
  { label: "Zen mode", keys: "Alt+Z" },
  { label: "Bold", keys: "Ctrl+B" },
  { label: "Italic", keys: "Ctrl+I" },
  { label: "Underline", keys: "Ctrl+U" },
  { label: "Close dialogs", keys: "Esc" },
];

const ACTIONS = [
  {
    id: "theme-toggle",
    category: "System",
    order: 100,
    getLabel: (ctx) => (ctx.isDarkMode && ctx.isDarkMode() ? "Switch to Light Mode" : "Switch to Dark Mode"),
    getIcon: (ctx) => (ctx.isDarkMode && ctx.isDarkMode() ? "fa-sun" : "fa-moon"),
    keywords: ["theme", "dark", "light", "appearance"],
    run: (ctx) => ctx.toggleTheme && ctx.toggleTheme(),
  },
  {
    id: "save-state",
    category: "System",
    order: 110,
    label: "Save Spreadsheet State",
    icon: "fa-floppy-disk",
    keywords: ["save", "state", "url", "hash"],
    run: async (ctx) => {
      if (ctx.recalculateFormulas) ctx.recalculateFormulas();
      if (ctx.updateURL) await ctx.updateURL();
      if (ctx.showToast) ctx.showToast("State saved", "success");
    },
  },
  {
    id: "clear-spreadsheet",
    category: "System",
    order: 120,
    label: "Clear Spreadsheet",
    icon: "fa-eraser",
    keywords: ["clear", "reset", "erase", "wipe"],
    run: (ctx) => ctx.clearSpreadsheet && ctx.clearSpreadsheet(),
  },
  {
    id: "toggle-readonly",
    category: "System",
    order: 130,
    getLabel: (ctx) => (ctx.isReadOnly && ctx.isReadOnly() ? "Disable Read-only Mode" : "Enable Read-only Mode"),
    getIcon: (ctx) => (ctx.isReadOnly && ctx.isReadOnly() ? "fa-pen-to-square" : "fa-eye"),
    keywords: ["read", "readonly", "view", "edit", "lock"],
    run: (ctx) => ctx.toggleReadOnly && ctx.toggleReadOnly(),
    isVisible: (ctx) => typeof ctx.toggleReadOnly === "function",
  },
  {
    id: "toggle-zen",
    category: "System",
    order: 135,
    getLabel: (ctx) => (ctx.isZen && ctx.isZen() ? "Disable Zen Mode" : "Enable Zen Mode"),
    getIcon: (ctx) => (ctx.isZen && ctx.isZen() ? "fa-eye-slash" : "fa-eye"),
    keywords: ["zen", "focus", "distraction", "mode"],
    run: (ctx) => ctx.toggleZen && ctx.toggleZen(),
    isVisible: (ctx) => typeof ctx.toggleZen === "function",
  },
  {
    id: "password-protect",
    category: "System",
    order: 140,
    label: "Password Protection",
    icon: "fa-lock",
    keywords: ["password", "lock", "encrypt", "security"],
    restoreFocus: false,
    run: (ctx) => ctx.openPasswordModal && ctx.openPasswordModal(),
  },
  {
    id: "import-csv",
    category: "Files",
    order: 200,
    label: "Import CSV",
    icon: "fa-file-import",
    keywords: ["csv", "import", "upload"],
    run: (ctx) => ctx.openCsvImport && ctx.openCsvImport(),
  },
  {
    id: "export-csv",
    category: "Files",
    order: 210,
    label: "Export CSV",
    icon: "fa-file-csv",
    keywords: ["csv", "export", "download"],
    run: (ctx) => ctx.exportCSV && ctx.exportCSV(),
  },
  {
    id: "export-excel",
    category: "Files",
    order: 220,
    label: "Export Excel",
    icon: "fa-file-excel",
    keywords: ["excel", "xlsx", "export", "download"],
    run: (ctx) => ctx.exportExcel && ctx.exportExcel(),
  },
  {
    id: "import-json",
    category: "Files",
    order: 230,
    label: "Import JSON",
    icon: "fa-file-code",
    keywords: ["json", "import", "hash", "decode"],
    restoreFocus: false,
    run: (ctx) => ctx.openHashTool && ctx.openHashTool(),
  },
  {
    id: "copy-json",
    category: "Files",
    order: 240,
    label: "Copy JSON",
    icon: "fa-file-code",
    keywords: ["json", "copy", "export", "ai"],
    restoreFocus: false,
    run: (ctx) => ctx.openJSONModal && ctx.openJSONModal(),
  },
  {
    id: "copy-link",
    category: "Share",
    order: 300,
    label: "Generate Shareable Link",
    icon: "fa-link",
    keywords: ["link", "share", "copy", "url"],
    run: (ctx) => {
      if (ctx.recalculateFormulas) ctx.recalculateFormulas();
      if (ctx.updateURL) {
        const updatePromise = ctx.updateURL();
        if (updatePromise && typeof updatePromise.catch === "function") {
          updatePromise.catch(() => {});
        }
      }
      if (ctx.copyLink) ctx.copyLink();
    },
  },
  {
    id: "qr-code",
    category: "Share",
    order: 310,
    label: "Show QR Code",
    icon: "fa-qrcode",
    keywords: ["qr", "code", "share", "mobile"],
    restoreFocus: false,
    run: (ctx) => ctx.showQRModal && ctx.showQRModal(),
  },
  {
    id: "embed-code",
    category: "Share",
    order: 320,
    label: "Generate Embed Code",
    icon: "fa-code",
    keywords: ["embed", "iframe", "share"],
    restoreFocus: false,
    run: (ctx) => ctx.showEmbedModal && ctx.showEmbedModal(),
  },
  {
    id: "p2p-host",
    category: "Collaboration",
    order: 400,
    label: "Host Live Session (P2P)",
    icon: "fa-wifi",
    keywords: ["p2p", "host", "collaboration", "session"],
    restoreFocus: false,
    run: (ctx) => ctx.startP2PHost && ctx.startP2PHost(),
  },
  {
    id: "p2p-join",
    category: "Collaboration",
    order: 410,
    label: "Join Live Session (P2P)",
    icon: "fa-plug",
    keywords: ["p2p", "join", "collaboration", "session"],
    restoreFocus: false,
    run: (ctx) => ctx.focusP2PJoin && ctx.focusP2PJoin(),
  },
  {
    id: "p2p-copy-id",
    category: "Collaboration",
    order: 420,
    label: "Copy Peer ID",
    icon: "fa-copy",
    keywords: ["p2p", "copy", "peer", "id"],
    restoreFocus: false,
    run: (ctx) => {
      const id = ctx.getP2PId ? ctx.getP2PId() : "";
      if (!id) {
        if (ctx.showToast) ctx.showToast("Start hosting to get a peer ID", "warning");
        return;
      }
      if (ctx.copyP2PId) ctx.copyP2PId();
    },
  },
  {
    id: "template-gallery",
    category: "Tools",
    order: 500,
    label: "Open Template Gallery",
    icon: "fa-table-cells-large",
    keywords: ["template", "gallery", "starter"],
    restoreFocus: false,
    run: (ctx) => ctx.openTemplateGallery && ctx.openTemplateGallery(),
  },
  {
    id: "presentation-mode",
    category: "Tools",
    order: 510,
    label: "Presentation Mode",
    icon: "fa-tv",
    keywords: ["presentation", "slides", "present"],
    restoreFocus: false,
    run: (ctx) => ctx.startPresentation && ctx.startPresentation(),
  },
  {
    id: "trace-dependencies",
    category: "Tools",
    order: 520,
    label: "Trace Dependencies",
    icon: "fa-diagram-project",
    keywords: ["trace", "dependencies", "logic", "visualize"],
    run: (ctx) => ctx.toggleDependencies && ctx.toggleDependencies(),
  },
  {
    id: "help-formulas",
    category: "Tools",
    order: 530,
    label: "Help / Formula Cheat Sheet",
    icon: "fa-book",
    keywords: ["help", "formula", "cheat", "guide"],
    restoreFocus: false,
    run: () => openHelpModal(),
  },
  {
    id: "open-source",
    category: "Tools",
    order: 540,
    label: "Open Source (GitHub)",
    icon: "fa-github",
    keywords: ["github", "source", "open", "repo"],
    run: (ctx) => ctx.openGitHub && ctx.openGitHub(),
  },
  {
    id: "format-bold",
    category: "Formatting",
    order: 600,
    label: "Bold",
    icon: "fa-bold",
    keywords: ["bold", "format", "style"],
    shortcut: "Ctrl+B",
    run: (ctx) => ctx.applyFormat && ctx.applyFormat("bold"),
  },
  {
    id: "format-italic",
    category: "Formatting",
    order: 610,
    label: "Italic",
    icon: "fa-italic",
    keywords: ["italic", "format", "style"],
    shortcut: "Ctrl+I",
    run: (ctx) => ctx.applyFormat && ctx.applyFormat("italic"),
  },
  {
    id: "format-underline",
    category: "Formatting",
    order: 620,
    label: "Underline",
    icon: "fa-underline",
    keywords: ["underline", "format", "style"],
    shortcut: "Ctrl+U",
    run: (ctx) => ctx.applyFormat && ctx.applyFormat("underline"),
  },
];

let callbacks = { ...defaultCallbacks };
let overlayEl = null;
let paletteEl = null;
let inputEl = null;
let resultsEl = null;
let activeIndex = 0;
let visibleActions = [];
let lastFocused = null;
let isOpen = false;
let initialized = false;
let helpModalEl = null;

const MAX_RESULTS = 24;

function normalize(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function getActionLabel(action) {
  if (action.getLabel) {
    return action.getLabel(callbacks);
  }
  return action.label || "";
}

function getActionIcon(action) {
  if (action.getIcon) {
    return action.getIcon(callbacks);
  }
  return action.icon || "fa-bolt";
}

function getSearchText(action) {
  const label = getActionLabel(action);
  const keywords = Array.isArray(action.keywords) ? action.keywords.join(" ") : "";
  const category = action.category || "";
  const description = action.description || "";
  return [label, keywords, category, description].filter(Boolean).join(" ");
}

function isVisible(action) {
  if (action.isVisible) {
    return action.isVisible(callbacks);
  }
  return true;
}

function scoreToken(token, text) {
  if (!token) return 0;
  const directIndex = text.indexOf(token);
  if (directIndex !== -1) {
    return token.length * 12 - directIndex;
  }
  let matchIndex = 0;
  let start = -1;
  for (let i = 0; i < text.length; i++) {
    if (text[i] !== token[matchIndex]) continue;
    if (start === -1) start = i;
    matchIndex += 1;
    if (matchIndex === token.length) {
      const spread = i - start;
      return token.length * 5 - spread - start * 0.1;
    }
  }
  return null;
}

function scoreMatch(query, action) {
  const normalizedQuery = normalize(query);
  if (!normalizedQuery) return 1;

  const haystack = normalize(getSearchText(action));
  if (!haystack) return -1;

  const tokens = normalizedQuery.split(" ").filter(Boolean);
  let score = 0;
  for (const token of tokens) {
    const tokenScore = scoreToken(token, haystack);
    if (tokenScore === null) return -1;
    score += tokenScore;
  }

  const label = normalize(getActionLabel(action));
  if (label.startsWith(normalizedQuery)) {
    score += 20;
  }

  return score;
}

function buildPalette() {
  overlayEl = document.createElement("div");
  overlayEl.className = "command-palette-overlay";
  overlayEl.setAttribute("aria-hidden", "true");

  paletteEl = document.createElement("div");
  paletteEl.className = "command-palette";
  paletteEl.setAttribute("role", "dialog");
  paletteEl.setAttribute("aria-modal", "true");
  paletteEl.setAttribute("aria-label", "Command palette");

  const searchRow = document.createElement("div");
  searchRow.className = "command-palette-search";

  const searchIcon = document.createElement("i");
  searchIcon.className = "fa-solid fa-magnifying-glass";

  inputEl = document.createElement("input");
  inputEl.className = "command-palette-input";
  inputEl.type = "text";
  inputEl.placeholder = "Type a command...";
  inputEl.setAttribute("autocomplete", "off");
  inputEl.setAttribute("autocapitalize", "off");
  inputEl.setAttribute("autocorrect", "off");
  inputEl.setAttribute("spellcheck", "false");
  inputEl.setAttribute("aria-controls", "command-palette-results");

  const hint = document.createElement("div");
  hint.className = "command-palette-hint";
  hint.textContent = "Esc to close";

  searchRow.append(searchIcon, inputEl, hint);

  resultsEl = document.createElement("ul");
  resultsEl.id = "command-palette-results";
  resultsEl.className = "command-palette-results";
  resultsEl.setAttribute("role", "listbox");

  paletteEl.append(searchRow, resultsEl);
  overlayEl.appendChild(paletteEl);
  document.body.appendChild(overlayEl);

  overlayEl.addEventListener("click", (event) => {
    if (event.target === overlayEl) {
      closePalette();
    }
  });

  paletteEl.addEventListener("click", (event) => {
    event.stopPropagation();
  });

  inputEl.addEventListener("input", () => {
    filterActions(inputEl.value);
  });

  inputEl.addEventListener("keydown", (event) => {
    if (event.key === "ArrowDown") {
      event.preventDefault();
      moveActive(1);
    } else if (event.key === "ArrowUp") {
      event.preventDefault();
      moveActive(-1);
    } else if (event.key === "Enter") {
      event.preventDefault();
      const action = visibleActions[activeIndex];
      if (action) executeAction(action);
    } else if (event.key === "Escape") {
      event.preventDefault();
      closePalette();
    }
  });

  resultsEl.addEventListener("mousemove", (event) => {
    const item = event.target.closest(".command-palette-item");
    if (!item) return;
    const index = Number(item.dataset.index);
    if (!Number.isNaN(index)) {
      setActiveIndex(index);
    }
  });

  resultsEl.addEventListener("click", (event) => {
    const item = event.target.closest(".command-palette-item");
    if (!item) return;
    const index = Number(item.dataset.index);
    const action = visibleActions[index];
    if (action) executeAction(action);
  });
}

function buildHelpModal() {
  if (helpModalEl) return;

  helpModalEl = document.createElement("div");
  helpModalEl.className = "modal hidden";
  helpModalEl.setAttribute("role", "dialog");
  helpModalEl.setAttribute("aria-modal", "true");
  helpModalEl.setAttribute("aria-label", "Formula cheat sheet");

  const backdrop = document.createElement("div");
  backdrop.className = "modal-backdrop";

  const content = document.createElement("div");
  content.className = "modal-content help-modal-content";

  const header = document.createElement("div");
  header.className = "help-modal-header";

  const title = document.createElement("h3");
  title.textContent = "Formula Cheat Sheet";

  const closeBtn = document.createElement("button");
  closeBtn.type = "button";
  closeBtn.className = "modal-icon-btn";
  closeBtn.setAttribute("aria-label", "Close help");
  closeBtn.innerHTML = '<i class="fa-solid fa-xmark"></i>';

  header.append(title, closeBtn);

  const subtitle = document.createElement("p");
  subtitle.className = "help-modal-subtitle";
  subtitle.textContent = "Quick syntax and shortcuts for common tasks.";

  const grid = document.createElement("div");
  grid.className = "help-modal-grid";

  const formulaSection = document.createElement("div");
  formulaSection.className = "help-section";
  const formulaTitle = document.createElement("h4");
  formulaTitle.textContent = "Formulas";
  const formulaList = document.createElement("ul");
  formulaList.className = "help-list";

  const allFormulas = [...FORMULA_SUGGESTIONS.map((item) => ({
    signature: item.signature,
    description: item.description,
  })), ...VISUAL_FORMULAS];

  allFormulas.forEach((item) => {
    const row = document.createElement("li");
    row.className = "help-row";
    const code = document.createElement("span");
    code.className = "help-code";
    code.textContent = `=${item.signature}`;
    const desc = document.createElement("span");
    desc.className = "help-note";
    desc.textContent = item.description;
    row.append(code, desc);
    formulaList.appendChild(row);
  });

  formulaSection.append(formulaTitle, formulaList);

  const shortcutSection = document.createElement("div");
  shortcutSection.className = "help-section";
  const shortcutTitle = document.createElement("h4");
  shortcutTitle.textContent = "Shortcuts";
  const shortcutList = document.createElement("ul");
  shortcutList.className = "help-list";

  SHORTCUTS.forEach((item) => {
    const row = document.createElement("li");
    row.className = "help-row";
    const label = document.createElement("span");
    label.className = "help-note";
    label.textContent = item.label;
    const keys = document.createElement("span");
    keys.className = "help-kbd";
    keys.textContent = item.keys;
    row.append(label, keys);
    shortcutList.appendChild(row);
  });

  shortcutSection.append(shortcutTitle, shortcutList);
  grid.append(formulaSection, shortcutSection);

  content.append(header, subtitle, grid);
  helpModalEl.append(backdrop, content);
  document.body.appendChild(helpModalEl);

  closeBtn.addEventListener("click", () => closeHelpModal());
  helpModalEl.addEventListener("click", (event) => {
    if (event.target === helpModalEl || event.target === backdrop) {
      closeHelpModal();
    }
  });
}

function openHelpModal() {
  if (!helpModalEl) buildHelpModal();
  if (helpModalEl) {
    helpModalEl.classList.remove("hidden");
    const closeBtn = helpModalEl.querySelector(".modal-icon-btn");
    if (closeBtn && typeof closeBtn.focus === "function") {
      requestAnimationFrame(() => closeBtn.focus());
    }
  }
}

function closeHelpModal() {
  if (helpModalEl) {
    helpModalEl.classList.add("hidden");
  }
}

function filterActions(query) {
  const available = ACTIONS.filter((action) => isVisible(action));
  if (!query) {
    visibleActions = available.sort((a, b) => (a.order || 999) - (b.order || 999));
  } else {
    visibleActions = available
      .map((action) => ({ action, score: scoreMatch(query, action) }))
      .filter((entry) => entry.score >= 0)
      .sort((a, b) => b.score - a.score || (a.action.order || 999) - (b.action.order || 999))
      .map((entry) => entry.action);
  }

  if (MAX_RESULTS && visibleActions.length > MAX_RESULTS) {
    visibleActions = visibleActions.slice(0, MAX_RESULTS);
  }

  activeIndex = 0;
  renderResults();
}

function renderResults() {
  if (!resultsEl) return;
  resultsEl.innerHTML = "";

  if (!visibleActions.length) {
    const empty = document.createElement("li");
    empty.className = "command-palette-empty";
    empty.textContent = "No commands found.";
    resultsEl.appendChild(empty);
    if (inputEl) inputEl.removeAttribute("aria-activedescendant");
    return;
  }

  visibleActions.forEach((action, index) => {
    const item = document.createElement("li");
    item.className = "command-palette-item";
    item.dataset.index = String(index);
    item.id = `command-palette-item-${index}`;
    item.setAttribute("role", "option");

    const iconWrap = document.createElement("div");
    iconWrap.className = "command-palette-item-icon";
    const icon = document.createElement("i");
    icon.className = `fa-solid ${getActionIcon(action)}`;
    iconWrap.appendChild(icon);

    const body = document.createElement("div");
    body.className = "command-palette-item-body";

    const label = document.createElement("div");
    label.className = "command-palette-item-label";
    label.textContent = getActionLabel(action);
    body.appendChild(label);

    if (action.category) {
      const meta = document.createElement("div");
      meta.className = "command-palette-item-meta";
      meta.textContent = action.category;
      body.appendChild(meta);
    }

    item.append(iconWrap, body);

    if (action.shortcut) {
      const shortcut = document.createElement("div");
      shortcut.className = "command-palette-item-shortcut";
      shortcut.textContent = action.shortcut;
      item.appendChild(shortcut);
    }

    resultsEl.appendChild(item);
  });

  updateActiveItem();
}

function updateActiveItem() {
  if (!resultsEl) return;
  const items = resultsEl.querySelectorAll(".command-palette-item");
  items.forEach((item, index) => {
    const isActive = index === activeIndex;
    item.classList.toggle("active", isActive);
    item.setAttribute("aria-selected", isActive ? "true" : "false");
  });

  if (inputEl) {
    if (items[activeIndex]) {
      inputEl.setAttribute("aria-activedescendant", items[activeIndex].id);
      items[activeIndex].scrollIntoView({ block: "nearest" });
    } else {
      inputEl.removeAttribute("aria-activedescendant");
    }
  }
}

function setActiveIndex(index) {
  if (!visibleActions.length) return;
  activeIndex = Math.max(0, Math.min(index, visibleActions.length - 1));
  updateActiveItem();
}

function moveActive(delta) {
  if (!visibleActions.length) return;
  activeIndex = (activeIndex + delta + visibleActions.length) % visibleActions.length;
  updateActiveItem();
}

async function executeAction(action) {
  try {
    if (typeof action.run === "function") {
      const result = action.run(callbacks);
      if (result && typeof result.then === "function") {
        await result;
      }
    }
  } catch (error) {
    console.error("Command palette action failed:", error);
    if (callbacks.showToast) {
      callbacks.showToast("Command failed", "error");
    }
  }

  if (!action.keepOpen) {
    closePalette({ restoreFocus: action.restoreFocus !== false });
  }
}

function openPalette() {
  if (isOpen || !overlayEl || !inputEl) return;
  isOpen = true;
  lastFocused = document.activeElement;
  overlayEl.classList.add("is-open");
  overlayEl.setAttribute("aria-hidden", "false");
  inputEl.value = "";
  filterActions("");
  requestAnimationFrame(() => {
    inputEl.focus();
  });
}

function closePalette(options = {}) {
  const { restoreFocus = true } = options;
  if (!isOpen || !overlayEl) return;
  isOpen = false;
  overlayEl.classList.remove("is-open");
  overlayEl.setAttribute("aria-hidden", "true");
  if (inputEl && typeof inputEl.blur === "function") {
    inputEl.blur();
  }
  if (restoreFocus && lastFocused && typeof lastFocused.focus === "function") {
    lastFocused.focus();
  }
}

function togglePalette() {
  if (isOpen) {
    closePalette();
  } else {
    openPalette();
  }
}

function handleGlobalKeydown(event) {
  if (event.defaultPrevented) return;
  const key = event.key ? event.key.toLowerCase() : "";
  const isOpenShortcut = (event.ctrlKey || event.metaKey) && key === "k";

  if (isOpenShortcut) {
    event.preventDefault();
    togglePalette();
    return;
  }

  if (isOpen && event.key === "Escape") {
    event.preventDefault();
    closePalette();
  }
}

export const CommandPaletteManager = {
  init(callbackOverrides = {}) {
    if (initialized) return;
    initialized = true;
    callbacks = { ...defaultCallbacks, ...callbackOverrides };
    buildPalette();
    buildHelpModal();
    document.addEventListener("keydown", handleGlobalKeydown);
    filterActions("");
  },
  open() {
    openPalette();
  },
  close() {
    closePalette();
  },
};
