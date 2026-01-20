// Dynamic Spreadsheet Web App
// Data persists in URL hash for easy sharing

// Import constants from ES6 module
import { DEBOUNCE_DELAY, DEFAULT_COLS, DEFAULT_ROWS, MAX_COLS, MAX_ROWS } from "./modules/constants.js";
import { buildRangeRef, FormulaDropdownManager, FormulaEvaluator, isValidFormula } from "./modules/formulaManager.js";
import { DependencyTracer } from "./modules/dependencyTracer.js";
import { PasswordManager } from "./modules/passwordManager.js";
import { CSVManager } from "./modules/csvManager.js";
import { JSONManager } from "./modules/jsonManager.js";
import { HashToolManager } from "./modules/hashToolManager.js";
import { P2PManager } from "./modules/p2pManager.js";
import { PresentationManager } from "./modules/presentationManager.js";
import { TemplateManager } from "./modules/templateManager.js";
import { ThemeManager } from "./modules/themeManager.js";
import { UIModeManager } from "./modules/uiModeManager.js";
import { QRCodeManager } from "./modules/qrCodeManager.js";
import { CellFormattingManager } from "./modules/cellFormattingManager.js";
import {
  addColumn,
  addRow,
  clearActiveHeaders,
  clearSelectedCells,
  clearSelection,
  clearSpreadsheet,
  focusCellAt,
  getCellContentElement,
  getCellElement,
  getSelectionBounds,
  getState,
  handleMouseDown,
  handleMouseLeave,
  handleMouseMove,
  handleMouseUp,
  handleResizeStart,
  handleTouchEnd,
  handleTouchMove,
  handleTouchStart,
  hasMultiSelection,
  renderGrid,
  setActiveHeaders,
  setCallbacks,
  setState,
  updateSelectionVisuals,
  insertRowAt,
  insertColumnAt,
} from "./modules/rowColManager.js";
import { escapeHTML, isValidCSSColor, sanitizeHTML } from "./modules/security.js";
import { showToast } from "./modules/toastManager.js";
import {
  createDefaultColumnWidths,
  createDefaultRowHeights,
  createEmptyCellStyle,
  createEmptyCellStyles,
  createEmptyData,
  isCellStylesDefault,
  isColWidthsDefault,
  isDataEmpty,
  isFormulasEmpty,
  isRowHeightsDefault,
  normalizeAlignment,
  normalizeCellStyles,
  normalizeColumnWidths,
  normalizeFontSize,
  normalizeRowHeights,
  minifyStateForExport,
  URLManager,
  validateAndNormalizeState,
} from "./modules/urlManager.js";

(function () {
  "use strict";
  // Data model - dynamic 2D array (rows/cols managed by rowColManager)
  let data = createEmptyData(DEFAULT_ROWS, DEFAULT_COLS);

  // Formula storage - parallel array to data
  let formulas = createEmptyData(DEFAULT_ROWS, DEFAULT_COLS);

  // Cell styles - alignment, colors, and font size
  let cellStyles = createEmptyCellStyles(DEFAULT_ROWS, DEFAULT_COLS);

  // Read-only mode flag
  let isReadOnly = false;

  // Embed mode flag
  let isEmbedMode = false;

  // Debounce timer
  let debounceTimer = null;
  let dependencyDrawQueued = false;

  // Formula range selection mode (for clicking to select ranges like Google Sheets)
  let formulaEditMode = false; // true when typing a formula
  let formulaEditCell = null; // { row, col, element } of cell being edited
  let formulaRangeStart = null; // Start of range being selected
  let formulaRangeEnd = null; // End of range being selected
  let editingCell = null; // { row, col } when editing a cell's text

  // P2P collaboration state
  let isRemoteUpdate = false;
  let hasInitialSync = false;
  let fullSyncTimer = null;
  const FULL_SYNC_DELAY = 300;
  const REMOTE_CURSOR_CLASS = "remote-active";

  const p2pUI = {
    modal: null,
    startHostBtn: null,
    joinBtn: null,
    copyIdBtn: null,
    idDisplay: null,
    statusEl: null,
    myIdInput: null,
    remoteIdInput: null,
    startHostLabel: "",
    joinLabel: "",
  };

  // Encryption state
  // Encryption state handled by PasswordManager

  // Toast functions moved to modules/toastManager.js

  // Create empty data array with specified dimensions
  // Factory functions moved to modules/urlManager.js

  // ========== Security Functions ==========

  // Security Functions moved to modules/security.js

  // Insert text at current cursor position in contentEditable
  function insertTextAtCursor(text) {
    const selection = window.getSelection();
    if (!selection.rangeCount) return;

    const range = selection.getRangeAt(0);
    range.deleteContents();

    const textNode = document.createTextNode(text);
    range.insertNode(textNode);

    // Move cursor after inserted text
    range.setStartAfter(textNode);
    range.collapse(true);
    selection.removeAllRanges();
    selection.addRange(range);
  }

  // Move caret to end of contentEditable element
  function setCaretToEnd(element) {
    const range = document.createRange();
    range.selectNodeContents(element);
    range.collapse(false);
    const selection = window.getSelection();
    selection.removeAllRanges();
    selection.addRange(range);
  }

  function applyFormulaSuggestion(formulaName) {
    const target = formulaEditCell ? formulaEditCell.element : document.activeElement;
    if (!target || !target.classList.contains("cell-content")) return;

    const row = parseInt(target.dataset.row, 10);
    const col = parseInt(target.dataset.col, 10);
    if (isNaN(row) || isNaN(col)) return;

    const newValue = `=${formulaName}(`;
    target.innerText = newValue;
    setCaretToEnd(target);

    formulaEditMode = true;
    formulaEditCell = { row, col, element: target };
    formulas[row][col] = newValue;
    data[row][col] = newValue;

    FormulaDropdownManager.hide();
    debouncedUpdateURL();
    scheduleDependencyDraw();
  }

  // ========== Formula Evaluation Functions ==========

  // Get numeric value from cell (returns 0 for empty/non-numeric)
  function getCellValue(row, col) {
    const { rows, cols } = getState();
    if (row < 0 || row >= rows || col < 0 || col >= cols) return 0;
    const val = data[row][col];
    if (!val || val === "") return 0;
    // Strip HTML tags and parse
    const stripped = String(val)
      .replace(/<[^>]*>/g, "")
      .trim();
    const normalized = stripped.replace(/,/g, "");
    const num = parseFloat(normalized);
    return isNaN(num) ? 0 : num;
  }

  // Recalculate all formula cells
  function recalculateFormulas() {
    const { rows, cols } = getState();
    const container = document.getElementById("spreadsheet");
    const activeElement = document.activeElement;
    const maxPasses = rows * cols;
    let needsUpdate = false;

    // Multiple passes to propagate formulas that depend on other formulas.
    for (let pass = 0; pass < maxPasses; pass++) {
      let changed = false;
      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
          const formula = formulas[r][c];
          if (formula && formula.startsWith("=")) {
            const result = String(FormulaEvaluator.evaluate(formula, { getCellValue, data, rows, cols }));
            if (data[r][c] !== result) {
              data[r][c] = result;
              changed = true;
              needsUpdate = true;
            }
          }
        }
      }
      if (!changed) break;
    }

    if (!needsUpdate || !container) return;

    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        const formula = formulas[r][c];
        if (formula && formula.startsWith("=")) {
          const cellContent = container.querySelector(`.cell-content[data-row="${r}"][data-col="${c}"]`);
          if (!cellContent) continue;

          const isEditingFormula = cellContent === activeElement && cellContent.innerText.trim().startsWith("=");
          if (!isEditingFormula) {
            cellContent.innerText = data[r][c];
          }
        }
      }
    }
  }

  // Normalization functions moved to modules/urlManager.js

  function extractPlainText(value) {
    if (value === null || value === undefined) return "";
    // Use DOMParser for safe HTML parsing (doesn't execute scripts)
    const parser = new DOMParser();
    const doc = parser.parseFromString("<body>" + String(value) + "</body>", "text/html");
    const text = doc.body.textContent || "";
    return text.replace(/\u00a0/g, " ");
  }

  function getDataArray() {
    return data;
  }

  function setDataArray(newData) {
    data = newData;
  }

  function getFormulasArray() {
    return formulas;
  }

  function setFormulasArray(newFormulas) {
    formulas = newFormulas;
  }

  function getCellStylesArray() {
    return cellStyles;
  }

  function setCellStylesArray(newCellStyles) {
    cellStyles = newCellStyles;
  }

  function formatStatNumber(value) {
    return Number.isInteger(value) ? value : value.toFixed(2);
  }

  // Calculate and render selection summary stats
  function updateSelectionStatusBar(boundsOverride = null) {
    const bar = document.getElementById("selection-status");
    const countEl = document.getElementById("stat-count");
    const sumEl = document.getElementById("stat-sum");
    const avgEl = document.getElementById("stat-avg");
    const sumWrapper = sumEl ? sumEl.parentElement : null;
    const avgWrapper = avgEl ? avgEl.parentElement : null;

    if (!bar || !countEl || !sumEl || !avgEl) return;
    if (isEmbedMode) {
      bar.classList.remove("active");
      return;
    }

    const bounds = boundsOverride || getSelectionBounds();
    if (!bounds) {
      bar.classList.remove("active");
      return;
    }

    const selectedCellCount = (bounds.maxRow - bounds.minRow + 1) * (bounds.maxCol - bounds.minCol + 1);
    if (selectedCellCount < 2) {
      bar.classList.remove("active");
      return;
    }

    let sum = 0;
    let countNumeric = 0;
    let countTotal = 0;

    for (let r = bounds.minRow; r <= bounds.maxRow; r++) {
      for (let c = bounds.minCol; c <= bounds.maxCol; c++) {
        const raw = data[r] && data[r][c] !== undefined ? data[r][c] : "";
        const text = extractPlainText(raw).trim();
        if (!text) continue;

        countTotal++;
        const cleaned = text.replace(/[^0-9.-]+/g, "");
        const numeric = cleaned && cleaned !== "-" && cleaned !== "." && cleaned !== "-." ? parseFloat(cleaned) : NaN;

        if (!isNaN(numeric)) {
          sum += numeric;
          countNumeric++;
        }
      }
    }

    if (countTotal === 0) {
      bar.classList.remove("active");
      return;
    }

    countEl.innerText = countTotal;

    if (countNumeric > 0) {
      sumEl.innerText = formatStatNumber(sum);
      avgEl.innerText = (sum / countNumeric).toFixed(2);
      if (sumWrapper) sumWrapper.style.display = "inline";
      if (avgWrapper) avgWrapper.style.display = "inline";
    } else {
      if (sumWrapper) sumWrapper.style.display = "none";
      if (avgWrapper) avgWrapper.style.display = "none";
    }

    bar.classList.add("active");
  }

  // Helper functions moved to modules/urlManager.js

  // Minify state object keys for smaller URL payload
  // State helpers moved to modules/urlManager.js

  // ========== Serialization Codec (Wrapper for minify/expand + compression) ==========
  // Handles state serialization without encryption concerns
  // Codec, validation, and decode functions moved to modules/urlManager.js

  // Build current state object (only includes non-empty/non-default values)
  function buildCurrentState() {
    const { rows, cols, colWidths, rowHeights } = getState();
    const stateObj = {
      rows,
      cols,
      theme: ThemeManager.isDarkMode() ? "dark" : "light",
    };

    // Only include readOnly if true (saves URL bytes)
    if (isReadOnly || isEmbedMode) {
      stateObj.readOnly = 1;
    }

    // Only include embed if true (saves URL bytes)
    if (isEmbedMode) {
      stateObj.embed = 1;
    }

    // Only include data if not all empty
    if (!isDataEmpty(data)) {
      stateObj.data = data;
    }

    // Only include formulas if any exist
    if (!isFormulasEmpty(formulas)) {
      stateObj.formulas = formulas;
    }

    // Only include cell styles if any are non-default
    if (!isCellStylesDefault(cellStyles)) {
      stateObj.cellStyles = cellStyles;
    }

    // Only include colWidths if not all default
    if (!isColWidthsDefault(colWidths, cols)) {
      stateObj.colWidths = colWidths;
    }

    // Only include rowHeights if not all default
    if (!isRowHeightsDefault(rowHeights, rows)) {
      stateObj.rowHeights = rowHeights;
    }

    return stateObj;
  }

  // Update URL hash without page jump
  async function updateURL() {
    const state = buildCurrentState();

    await URLManager.updateURL(state, PasswordManager.getPassword());
  }

  // Debounced URL update
  function debouncedUpdateURL() {
    if (debounceTimer) {
      clearTimeout(debounceTimer);
    }
    debounceTimer = setTimeout(updateURL, DEBOUNCE_DELAY);
  }

  function scheduleDependencyDraw() {
    if (!DependencyTracer.isActive) return;
    if (dependencyDrawQueued) return;
    dependencyDrawQueued = true;
    requestAnimationFrame(() => {
      dependencyDrawQueued = false;
      DependencyTracer.draw(getFormulasArray());
    });
  }

  function refreshDependencyLayer() {
    DependencyTracer.init();
    scheduleDependencyDraw();
  }
  // Handle input changes
  function handleInput(event) {
    const target = event.target;
    if (!target.classList.contains("cell-content")) return;

    // GUARD: Block in read-only mode
    if (isReadOnly) {
      event.preventDefault();
      target.blur();
      return;
    }

    const row = parseInt(target.dataset.row, 10);
    const col = parseInt(target.dataset.col, 10);

    // DON'T clear selection if in formula mode (user may be selecting range)
    if (hasMultiSelection() && !formulaEditMode) {
      clearSelection();
      setActiveHeaders(row, col);
    }

    const { rows, cols } = getState();
    if (!isNaN(row) && !isNaN(col) && row < rows && col < cols) {
      setEditingCell(row, col);
      const previousFormula = formulas[row] ? formulas[row][col] : "";
      const rawValue = target.innerText.trim();

      if (rawValue.startsWith("=")) {
        // Enter formula edit mode
        formulaEditMode = true;
        formulaEditCell = { row, col, element: target };

        // Store formula but DON'T evaluate during typing
        formulas[row][col] = rawValue;
        data[row][col] = rawValue;

        FormulaDropdownManager.update(target, rawValue);
      } else {
        // Exit formula edit mode
        formulaEditMode = false;
        formulaEditCell = null;

        // Regular value - clear any existing formula
        formulas[row][col] = "";
        data[row][col] = sanitizeHTML(target.innerHTML);

        FormulaDropdownManager.hide();

        // Recalculate dependent formulas when regular values change
        recalculateFormulas();
      }

      debouncedUpdateURL();
      if (canBroadcastP2P()) {
        P2PManager.broadcastCellUpdate(row, col, data[row][col], formulas[row][col]);
      }
      updateSelectionStatusBar();
      if (previousFormula !== formulas[row][col]) {
        scheduleDependencyDraw();
      }
    }
  }

  // Selection and hover functions moved to modules/rowColManager.js

  // Cell positioning functions moved to modules/rowColManager.js

  function setEditingCell(row, col) {
    editingCell = { row, col };
  }

  function clearEditingCell() {
    editingCell = null;
  }

  function isEditingCell(row, col) {
    return !!(editingCell && editingCell.row === row && editingCell.col === col);
  }

  function handleFocusIn(event) {
    const target = event.target;
    if (!target.classList.contains("cell-content")) return;

    const row = parseInt(target.dataset.row, 10);
    const col = parseInt(target.dataset.col, 10);

    if (isNaN(row) || isNaN(col)) return;

    // If we have a multi-selection and focus moves to a cell outside it, clear selection
    if (hasMultiSelection()) {
      const bounds = getSelectionBounds();
      if (bounds && (row < bounds.minRow || row > bounds.maxRow || col < bounds.minCol || col > bounds.maxCol)) {
        clearSelection();
      }
    }

    // Update header highlighting for single cell focus (if no multi-selection)
    if (!hasMultiSelection()) {
      setActiveHeaders(row, col);
    }

    // Show formula text when focused (for editing)
    if (formulas[row][col] && formulas[row][col].startsWith("=")) {
      target.innerText = formulas[row][col];
    }

    if (canBroadcastP2P()) {
      P2PManager.broadcastCursor(row, col, "#ff0055");
    }
  }

  function handleFocusOut(event) {
    const target = event.target;
    if (!target.classList.contains("cell-content")) return;

    FormulaDropdownManager.hide();

    const row = parseInt(target.dataset.row, 10);
    const col = parseInt(target.dataset.col, 10);

    // If we're in formula edit mode and currently selecting a range, don't process blur
    const { isSelecting, rows, cols } = getState();
    if (formulaEditMode && isSelecting) {
      return;
    }

    // Evaluate formula when blurred
    if (!isNaN(row) && !isNaN(col)) {
      const rawValue = target.innerText.trim();
      const previousFormula = formulas[row] ? formulas[row][col] : "";

      if (rawValue.startsWith("=")) {
        // NOW evaluate the formula
        formulas[row][col] = rawValue;
        const result = FormulaEvaluator.evaluate(rawValue, { getCellValue, data, rows, cols });
        data[row][col] = String(result);
        target.innerText = String(result);

        // Recalculate all dependent formulas
        recalculateFormulas();
        debouncedUpdateURL();
        updateSelectionStatusBar();
        if (canBroadcastP2P()) {
          P2PManager.broadcastCellUpdate(row, col, data[row][col], formulas[row][col]);
        }
        if (previousFormula !== formulas[row][col]) {
          scheduleDependencyDraw();
        }
      }
    }

    // Exit formula edit mode when focus truly leaves
    clearEditingCell();
    formulaEditMode = false;
    formulaEditCell = null;
    formulaRangeStart = null;
    formulaRangeEnd = null;

    const container = document.getElementById("spreadsheet");
    if (!container) return;

    const next = event.relatedTarget;
    if (next && container.contains(next)) {
      return;
    }

    clearActiveHeaders();
  }

  // ========== P2P Collaboration ==========

  function setP2PStatus(message) {
    if (p2pUI.statusEl) {
      p2pUI.statusEl.textContent = message;
    }
  }

  function resetP2PControls() {
    if (p2pUI.startHostBtn) {
      p2pUI.startHostBtn.disabled = false;
      if (p2pUI.startHostLabel) {
        p2pUI.startHostBtn.innerHTML = p2pUI.startHostLabel;
      }
    }
    if (p2pUI.joinBtn) {
      p2pUI.joinBtn.disabled = false;
      if (p2pUI.joinLabel) {
        p2pUI.joinBtn.innerHTML = p2pUI.joinLabel;
      }
    }
    if (p2pUI.idDisplay) {
      p2pUI.idDisplay.classList.add("hidden");
    }
  }

  function canBroadcastP2P() {
    if (isRemoteUpdate) return false;
    if (!P2PManager.canSend()) return false;
    if (!P2PManager.isHost && !hasInitialSync) return false;
    return true;
  }

  function sendFullSyncNow() {
    if (!canBroadcastP2P()) return;
    recalculateFormulas();
    const fullState = buildCurrentState();
    P2PManager.sendFullSync(fullState);
  }

  function scheduleFullSync() {
    if (!canBroadcastP2P()) return;
    if (fullSyncTimer) {
      clearTimeout(fullSyncTimer);
    }
    fullSyncTimer = setTimeout(() => {
      sendFullSyncNow();
    }, FULL_SYNC_DELAY);
  }

  function handleGridResize() {
    scheduleFullSync();
    scheduleDependencyDraw();
  }

  function clearRemoteCursor() {
    document.querySelectorAll(`.${REMOTE_CURSOR_CLASS}`).forEach((el) => el.classList.remove(REMOTE_CURSOR_CLASS));
  }

  function handleP2PConnection(amIHost) {
    hasInitialSync = amIHost;
    if (amIHost) {
      recalculateFormulas();
      const fullState = buildCurrentState();
      P2PManager.sendInitialSync(fullState);
      setP2PStatus("Connected. You can now edit together.");
      if (p2pUI.modal) {
        p2pUI.modal.classList.add("hidden");
      }
    } else {
      setP2PStatus("Connected. Waiting for host data...");
      showToast("Waiting for host data...", "info");
    }
  }

  function handleInitialSync(remoteState) {
    if (!remoteState) return;
    isRemoteUpdate = true;
    applyLoadedState(remoteState);
    renderGrid();
    DependencyTracer.init();
    recalculateFormulas();
    debouncedUpdateURL();
    updateSelectionStatusBar();
    scheduleDependencyDraw();
    clearRemoteCursor();
    isRemoteUpdate = false;
    hasInitialSync = true;
    setP2PStatus("Synced with host.");
    if (p2pUI.modal) {
      p2pUI.modal.classList.add("hidden");
    }
    showToast("Synced with host!", "success");
  }

  function handleRemoteCellUpdate({ row, col, value, formula }) {
    const rowIndex = parseInt(row, 10);
    const colIndex = parseInt(col, 10);
    if (isNaN(rowIndex) || isNaN(colIndex)) return;

    const { rows, cols } = getState();
    if (rowIndex < 0 || colIndex < 0 || rowIndex >= rows || colIndex >= cols) {
      P2PManager.requestFullSync();
      return;
    }

    const rawValue = value === undefined || value === null ? "" : String(value);
    const rawFormula = formula === undefined || formula === null ? "" : String(formula);
    const previousFormula = formulas[rowIndex] ? formulas[rowIndex][colIndex] : "";
    const safeValue = sanitizeHTML(rawValue);
    const isFormulaCell = rawFormula.startsWith("=");
    const isFormulaEdit = isFormulaCell && rawValue.trim().startsWith("=");

    isRemoteUpdate = true;
    data[rowIndex][colIndex] = isFormulaEdit ? rawFormula : safeValue;
    formulas[rowIndex][colIndex] = rawFormula;

    const cellContent = getCellContentElement(rowIndex, colIndex);
    if (cellContent && document.activeElement !== cellContent) {
      if (isFormulaEdit) {
        cellContent.innerText = rawFormula;
      } else {
        cellContent.innerHTML = safeValue;
      }
    }

    if (!isFormulaEdit) {
      recalculateFormulas();
    }

    debouncedUpdateURL();
    if (previousFormula !== formulas[rowIndex][colIndex]) {
      scheduleDependencyDraw();
    }
    isRemoteUpdate = false;
  }

  function handleRemoteCursor({ row, col }) {
    const rowIndex = parseInt(row, 10);
    const colIndex = parseInt(col, 10);
    if (isNaN(rowIndex) || isNaN(colIndex)) return;

    const { rows, cols } = getState();
    if (rowIndex < 0 || colIndex < 0 || rowIndex >= rows || colIndex >= cols) return;

    clearRemoteCursor();
    const cell = getCellElement(rowIndex, colIndex);
    if (cell) {
      cell.classList.add(REMOTE_CURSOR_CLASS);
    }
  }

  function handleSyncRequest() {
    if (!P2PManager.canSend()) return;
    if (!hasInitialSync && !P2PManager.isHost) return;
    sendFullSyncNow();
  }

  function handleP2PDisconnect() {
    hasInitialSync = false;
    clearRemoteCursor();
    resetP2PControls();
    setP2PStatus("Disconnected.");
    showToast("Peer disconnected", "warning");
  }

  // ========== Mouse Selection Handlers ==========

  function handleCellDoubleClick(event) {
    if (!(event.target instanceof Element)) return;

    let cellContent = null;
    if (event.target.classList.contains("cell-content")) {
      cellContent = event.target;
    } else if (event.target.classList.contains("cell")) {
      cellContent = event.target.querySelector(".cell-content");
    } else {
      cellContent = event.target.closest(".cell-content");
    }

    if (!cellContent || !cellContent.classList.contains("cell-content")) return;

    const row = parseInt(cellContent.dataset.row, 10);
    const col = parseInt(cellContent.dataset.col, 10);

    if (isNaN(row) || isNaN(col)) return;
    setEditingCell(row, col);
  }

  // Mouse handlers moved to modules/rowColManager.js

  function handleSelectionKeyDown(event) {
    if (FormulaDropdownManager.isOpen()) {
      if (event.key === "ArrowDown") {
        FormulaDropdownManager.moveSelection(1);
        event.preventDefault();
        return;
      }
      if (event.key === "ArrowUp") {
        FormulaDropdownManager.moveSelection(-1);
        event.preventDefault();
        return;
      }
      if (event.key === "Enter" || event.key === "Tab") {
        const formulaName = FormulaDropdownManager.getActiveFormulaName();
        if (formulaName) {
          applyFormulaSuggestion(formulaName);
        }
        event.preventDefault();
        return;
      }
      if (event.key === "Escape") {
        FormulaDropdownManager.hide();
        event.preventDefault();
        return;
      }
    }

    const state = getState();
    const { rows, cols, selectionStart } = state;

    if (event.key === "ArrowUp" || event.key === "ArrowDown" || event.key === "ArrowLeft" || event.key === "ArrowRight") {
      const target = event.target;
      if (target.classList.contains("cell-content") && !event.altKey && !event.ctrlKey && !event.metaKey) {
        const row = parseInt(target.dataset.row, 10);
        const col = parseInt(target.dataset.col, 10);

        if (!isNaN(row) && !isNaN(col) && !formulaEditMode && !isEditingCell(row, col)) {
          let nextRow = row;
          let nextCol = col;

          if (event.key === "ArrowUp") nextRow -= 1;
          if (event.key === "ArrowDown") nextRow += 1;
          if (event.key === "ArrowLeft") nextCol -= 1;
          if (event.key === "ArrowRight") nextCol += 1;

          nextRow = Math.max(0, Math.min(rows - 1, nextRow));
          nextCol = Math.max(0, Math.min(cols - 1, nextCol));

          event.preventDefault();

          if (event.shiftKey) {
            if (!selectionStart) {
              setState("selectionStart", { row, col });
            }
            setState("selectionEnd", { row: nextRow, col: nextCol });
          } else {
            setState("selectionStart", { row: nextRow, col: nextCol });
            setState("selectionEnd", { row: nextRow, col: nextCol });
          }

          updateSelectionVisuals();
          focusCellAt(nextRow, nextCol);
          return;
        }
      }
    }

    // Escape key clears selection
    if (event.key === "Escape" && hasMultiSelection()) {
      clearSelection();
      event.preventDefault();
      return;
    }

    // Ctrl+Shift+Arrow keys to insert rows/columns
    if (event.ctrlKey && event.shiftKey) {
      const target = event.target;
      if (target.classList.contains("cell-content")) {
        const row = parseInt(target.dataset.row, 10);
        const col = parseInt(target.dataset.col, 10);

        if (!isNaN(row) && !isNaN(col)) {
          if (event.key === "ArrowDown") {
            // Insert row below current
            event.preventDefault();
            insertRowAt(row + 1);
            return;
          }
          if (event.key === "ArrowUp") {
            // Insert row at current (shifting down)
            event.preventDefault();
            insertRowAt(row);
            return;
          }
          if (event.key === "ArrowRight") {
            // Insert column right
            event.preventDefault();
            insertColumnAt(col + 1);
            return;
          }
          if (event.key === "ArrowLeft") {
            // Insert column left
            event.preventDefault();
            insertColumnAt(col);
            return;
          }
        }
      }
    }

    // Delete/Backspace key clears selected cells
    if (event.key === "Delete" || event.key === "Backspace") {
      const activeElement = document.activeElement;
      const isEditingContent = activeElement && activeElement.classList.contains("cell-content") && activeElement.innerText.length > 0;

      // Clear all cells if multi-selection, or clear single cell if not actively editing
      if (hasMultiSelection() || !isEditingContent) {
        event.preventDefault();
        clearSelectedCells();
        return;
      }
    }

    // Enter key: evaluate formula / move to cell below
    if (event.key === "Enter") {
      const target = event.target;
      if (!target.classList.contains("cell-content")) return;

      const row = parseInt(target.dataset.row, 10);
      const col = parseInt(target.dataset.col, 10);

      if (isNaN(row) || isNaN(col)) return;

      // Prevent default newline behavior
      event.preventDefault();
      clearEditingCell();

      // Check if this is a formula cell - evaluate it
      const rawValue = target.innerText.trim();
      if (rawValue.startsWith("=")) {
        const previousFormula = formulas[row] ? formulas[row][col] : "";
        formulas[row][col] = rawValue;
        const result = FormulaEvaluator.evaluate(rawValue, { getCellValue, data, rows, cols });
        data[row][col] = String(result);
        target.innerText = String(result);
        recalculateFormulas();
        debouncedUpdateURL();
        if (canBroadcastP2P()) {
          P2PManager.broadcastCellUpdate(row, col, data[row][col], formulas[row][col]);
        }
        if (previousFormula !== formulas[row][col]) {
          scheduleDependencyDraw();
        }

        // Exit formula edit mode
        formulaEditMode = false;
        formulaEditCell = null;
      }

      // Try to move to cell below
      const nextRow = row + 1;
      if (nextRow < rows) {
        // Focus cell below
        const nextCell = document.querySelector(`.cell-content[data-row="${nextRow}"][data-col="${col}"]`);
        if (nextCell) {
          nextCell.focus();
        }
      } else {
        // No row below - just blur current cell
        target.blur();
      }
    }
  }

  // Resize and Touch handlers moved to modules/rowColManager.js

  // Handle paste to strip unwanted HTML
  function handlePaste(event) {
    const target = event.target;
    if (!target.classList.contains("cell-content")) return;

    // GUARD: Block in read-only mode (silent)
    if (isReadOnly) {
      event.preventDefault();
      return;
    }

    event.preventDefault();
    const text = event.clipboardData.getData("text/plain");
    document.execCommand("insertText", false, text);
  }

  // clearSpreadsheet, addRow, addColumn moved to modules/rowColManager.js

  // Load state from URL on page load
  // Returns true if data loaded successfully, false if waiting for password
  function loadStateFromURL() {
    const hash = window.location.hash.slice(1); // Remove #

    if (hash) {
      const loadedState = URLManager.decodeState(hash);
      if (loadedState) {
        // Check if data is encrypted
        if (loadedState.encrypted) {
          // delegate to PasswordManager
          PasswordManager.handleEncryptedData(loadedState.data);

          // Initialize with default state while waiting for password
          setState("rows", DEFAULT_ROWS);
          setState("cols", DEFAULT_COLS);
          setState("colWidths", createDefaultColumnWidths(DEFAULT_COLS));
          setState("rowHeights", createDefaultRowHeights(DEFAULT_ROWS));
          data = createEmptyData(DEFAULT_ROWS, DEFAULT_COLS);
          cellStyles = createEmptyCellStyles(DEFAULT_ROWS, DEFAULT_COLS);
          formulas = createEmptyData(DEFAULT_ROWS, DEFAULT_COLS);
          isReadOnly = false; // Default while waiting for decrypt
          return false;
        }

        const loadedRows = loadedState.rows;
        const loadedCols = loadedState.cols;
        setState("rows", loadedRows);
        setState("cols", loadedCols);
        setState("colWidths", loadedState.colWidths || createDefaultColumnWidths(loadedCols));
        setState("rowHeights", loadedState.rowHeights || createDefaultRowHeights(loadedRows));
        data = loadedState.data;
        formulas = loadedState.formulas || createEmptyData(loadedRows, loadedCols);
        cellStyles = loadedState.cellStyles || createEmptyCellStyles(loadedRows, loadedCols);

        // Apply theme from URL if present
        if (loadedState.theme) {
          ThemeManager.applyTheme(loadedState.theme);
        }

        // Load read-only mode
        UIModeManager.applyReadOnlyState(loadedState.readOnly);

        // Load embed mode
        isEmbedMode = loadedState.embed || false;
        if (isEmbedMode) {
          UIModeManager.applyEmbedMode();
        }

        return true;
      }
    }

    // Default state
    setState("rows", DEFAULT_ROWS);
    setState("cols", DEFAULT_COLS);
    setState("colWidths", createDefaultColumnWidths(DEFAULT_COLS));
    setState("rowHeights", createDefaultRowHeights(DEFAULT_ROWS));
    data = createEmptyData(DEFAULT_ROWS, DEFAULT_COLS);
    cellStyles = createEmptyCellStyles(DEFAULT_ROWS, DEFAULT_COLS);
    formulas = createEmptyData(DEFAULT_ROWS, DEFAULT_COLS);
    isReadOnly = false;
    isEmbedMode = false;
    UIModeManager.clearReadOnlyMode();
    UIModeManager.clearEmbedMode();
    return true;
  }

  // ========== Password/Encryption Modal Functions ==========
  // Handled by PasswordManager module

  // Apply loaded state to variables
  function applyLoadedState(loadedState) {
    if (!loadedState) return;

    const r = Math.min(Math.max(1, loadedState.rows || DEFAULT_ROWS), MAX_ROWS);
    const c = Math.min(Math.max(1, loadedState.cols || DEFAULT_COLS), MAX_COLS);

    setState("rows", r);
    setState("cols", c);

    // Process data array
    let d = loadedState.data;
    if (Array.isArray(d)) {
      d = d.slice(0, r).map((row) => {
        if (Array.isArray(row)) {
          return row.slice(0, c).map((cell) => sanitizeHTML(String(cell || "")));
        }
        return Array(c).fill("");
      });
      while (d.length < r) d.push(Array(c).fill(""));
      d = d.map((row) => {
        while (row.length < c) row.push("");
        return row;
      });
    } else {
      d = createEmptyData(r, c);
    }
    data = d;

    // Process formulas
    let f = loadedState.formulas;
    if (Array.isArray(f)) {
      f = f.slice(0, r).map((row, rowIdx) => {
        if (Array.isArray(row)) {
          return row.slice(0, c).map((cell, colIdx) => {
            const formula = String(cell || "");
            if (formula.startsWith("=")) {
              if (isValidFormula(formula)) {
                return formula;
              } else {
                data[rowIdx][colIdx] = escapeHTML(formula);
                return "";
              }
            }
            return formula;
          });
        }
        return Array(c).fill("");
      });
      while (f.length < r) f.push(Array(c).fill(""));
      f = f.map((row) => {
        while (row.length < c) row.push("");
        return row;
      });
    } else {
      f = createEmptyData(r, c);
    }
    formulas = f;

    cellStyles = normalizeCellStyles(loadedState.cellStyles, r, c);
    setState("colWidths", normalizeColumnWidths(loadedState.colWidths, c));
    setState("rowHeights", normalizeRowHeights(loadedState.rowHeights, r));

    if (loadedState.theme) {
      ThemeManager.applyTheme(loadedState.theme);
    }

    isEmbedMode = Boolean(loadedState.embed);
    if (isEmbedMode) {
      UIModeManager.applyEmbedMode();
    } else {
      UIModeManager.clearEmbedMode();
    }

    // Apply read-only mode if present (embed mode forces read-only)
    UIModeManager.applyReadOnlyState(loadedState.readOnly);
  }

  // Initialize the app
  function init() {
    // Set up callbacks for rowColManager module
    setCallbacks({
      debouncedUpdateURL,
      recalculateFormulas,
      getDataArray,
      setDataArray,
      getFormulasArray,
      setFormulasArray,
      getCellStylesArray,
      setCellStylesArray,
      PasswordManager,
      getReadOnlyFlag: () => isReadOnly,
      // Formula mode callbacks
      getFormulaEditMode: () => formulaEditMode,
      getFormulaEditCell: () => formulaEditCell,
      setFormulaRangeStart: (val) => {
        formulaRangeStart = val;
      },
      setFormulaRangeEnd: (val) => {
        formulaRangeEnd = val;
      },
      getFormulaRangeStart: () => formulaRangeStart,
      getFormulaRangeEnd: () => formulaRangeEnd,
      buildRangeRef,
      insertTextAtCursor,
      FormulaDropdownManager,
      onSelectionChange: updateSelectionStatusBar,
      onGridResize: handleGridResize,
      onFormulaChange: scheduleDependencyDraw,
    });

    CSVManager.init({
      getState,
      getDataArray,
      setDataArray,
      setFormulasArray,
      setCellStylesArray,
      setState,
      renderGrid,
      recalculateFormulas,
      debouncedUpdateURL,
      showToast,
      extractPlainText,
      onImport: () => {
        scheduleFullSync();
        refreshDependencyLayer();
      },
    });

    JSONManager.init({
      buildCurrentState: () => minifyStateForExport(buildCurrentState()),
      recalculateFormulas,
      showToast,
    });
    HashToolManager.init({
      showToast,
    });

    P2PManager.init({
      onHostReady: (id) => {
        if (p2pUI.myIdInput) {
          p2pUI.myIdInput.value = id;
        }
        if (p2pUI.idDisplay) {
          p2pUI.idDisplay.classList.remove("hidden");
        }
        setP2PStatus("Waiting for connection...");
      },
      onConnectionOpened: handleP2PConnection,
      onInitialSync: handleInitialSync,
      onRemoteCellUpdate: handleRemoteCellUpdate,
      onRemoteCursorMove: handleRemoteCursor,
      onSyncRequest: handleSyncRequest,
      onConnectionClosed: handleP2PDisconnect,
      onConnectionError: () => {
        resetP2PControls();
        setP2PStatus("Connection error.");
      },
    });

    // Initialize Theme Manager
    ThemeManager.init({
      debouncedUpdateURL,
      scheduleFullSync,
    });

    // Initialize UI Mode Manager
    UIModeManager.init({
      buildCurrentState,
      updateURL,
      scheduleFullSync,
      showToast,
      URLManager,
      PasswordManager,
      getReadOnlyFlag: () => isReadOnly,
      setReadOnlyFlag: (val) => {
        isReadOnly = val;
      },
      getEmbedModeFlag: () => isEmbedMode,
      setEmbedModeFlag: (val) => {
        isEmbedMode = val;
      },
    });

    // Initialize QR Code Manager
    QRCodeManager.init({
      recalculateFormulas,
      updateURL,
      showToast,
    });

    // Initialize Cell Formatting Manager
    CellFormattingManager.init({
      getSelectionBounds,
      getCellContentElement,
      getCellElement,
      getCellStylesArray,
      getDataArray,
      getFormulasArray,
      getState,
      debouncedUpdateURL,
      scheduleFullSync,
      normalizeAlignment,
      normalizeFontSize,
      isValidCSSColor,
      sanitizeHTML,
      createEmptyCellStyle,
      canBroadcastP2P,
      P2PManager,
    });

    // Load theme preference first (before any rendering)
    ThemeManager.loadTheme();

    // Load any existing state from URL
    loadStateFromURL();

    // Recalculate after loading initial state
    recalculateFormulas();

    // Render the grid
    renderGrid();

    // Initialize Dependency Tracer layer
    DependencyTracer.init();

    // Initialize Formula Dropdown
    FormulaDropdownManager.init(applyFormulaSuggestion);

    // Initialize Password Manager
    PasswordManager.init({
      decryptAndDecode: URLManager.decryptAndDecode,
      onDecryptSuccess: (state) => {
        applyLoadedState(state);
        recalculateFormulas(); // Recalculate after decryption
        renderGrid();
        DependencyTracer.init();
        scheduleDependencyDraw();
      },
      updateURL: updateURL,
      showToast: showToast,
      validateState: validateAndNormalizeState,
    });

    PresentationManager.init();

    // Initialize URL length indicator with current hash length
    const currentHash = window.location.hash.slice(1);
    URLManager.updateURLLengthIndicator(currentHash.length);

    // Set up event delegation for input handling
    const container = document.getElementById("spreadsheet");
    if (container) {
      container.addEventListener("input", handleInput);
      container.addEventListener("focusin", handleFocusIn);
      container.addEventListener("focusout", handleFocusOut);
      container.addEventListener("paste", handlePaste);

      // Selection mouse events
      container.addEventListener("mousedown", handleResizeStart);
      container.addEventListener("mousedown", handleMouseDown);
      container.addEventListener("dblclick", handleCellDoubleClick);
      container.addEventListener("mousemove", handleMouseMove);
      container.addEventListener("mouseleave", handleMouseLeave);
      container.addEventListener("mouseup", handleMouseUp);
      container.addEventListener("keydown", handleSelectionKeyDown);
      container.addEventListener("keyup", () => updateSelectionStatusBar());

      // Touch events for mobile selection
      container.addEventListener("touchstart", handleTouchStart, { passive: false });
      container.addEventListener("touchmove", handleTouchMove, { passive: false });
      container.addEventListener("touchend", handleTouchEnd);
    }

    const gridWrapper = document.querySelector(".grid-wrapper");
    if (gridWrapper) {
      gridWrapper.addEventListener("scroll", function () {
        if (FormulaDropdownManager.isOpen() && FormulaDropdownManager.anchor) {
          FormulaDropdownManager.position(FormulaDropdownManager.anchor);
        }
        scheduleDependencyDraw();
      });
    }

    window.addEventListener("resize", function () {
      if (FormulaDropdownManager.isOpen() && FormulaDropdownManager.anchor) {
        FormulaDropdownManager.position(FormulaDropdownManager.anchor);
      }
      scheduleDependencyDraw();
    });

    // Global mouseup to catch drag ending outside container
    document.addEventListener("mouseup", handleMouseUp);

    // Format button event listeners
    const boldBtn = document.getElementById("format-bold");
    const italicBtn = document.getElementById("format-italic");
    const underlineBtn = document.getElementById("format-underline");
    const alignLeftBtn = document.getElementById("align-left");
    const alignCenterBtn = document.getElementById("align-center");
    const alignRightBtn = document.getElementById("align-right");
    const cellBgPicker = document.getElementById("cell-bg-color");
    const cellTextColorPicker = document.getElementById("cell-text-color");
    const fontSizeList = document.getElementById("font-size-list");

    if (boldBtn) {
      boldBtn.addEventListener("mousedown", function (e) {
        e.preventDefault(); // Prevent focus loss
        CellFormattingManager.applyFormat("bold");
      });
    }
    if (italicBtn) {
      italicBtn.addEventListener("mousedown", function (e) {
        e.preventDefault();
        CellFormattingManager.applyFormat("italic");
      });
    }
    if (underlineBtn) {
      underlineBtn.addEventListener("mousedown", function (e) {
        e.preventDefault();
        CellFormattingManager.applyFormat("underline");
      });
    }
    if (alignLeftBtn) {
      alignLeftBtn.addEventListener("mousedown", function (e) {
        e.preventDefault();
        CellFormattingManager.applyAlignment("left");
      });
    }
    if (alignCenterBtn) {
      alignCenterBtn.addEventListener("mousedown", function (e) {
        e.preventDefault();
        CellFormattingManager.applyAlignment("center");
      });
    }
    if (alignRightBtn) {
      alignRightBtn.addEventListener("mousedown", function (e) {
        e.preventDefault();
        CellFormattingManager.applyAlignment("right");
      });
    }
    if (cellBgPicker) {
      cellBgPicker.addEventListener("input", function (e) {
        CellFormattingManager.applyCellBackground(e.target.value);
      });
    }
    if (cellTextColorPicker) {
      cellTextColorPicker.addEventListener("input", function (e) {
        CellFormattingManager.applyCellTextColor(e.target.value);
      });
    }
    if (fontSizeList) {
      fontSizeList.addEventListener("mousedown", function (e) {
        if (e.target.closest("button")) {
          e.preventDefault();
        }
      });
      fontSizeList.addEventListener("click", function (e) {
        const button = e.target.closest("button[data-size]");
        if (!button) return;
        CellFormattingManager.applyFontSize(button.dataset.size);
      });
    }

    // Button event listeners
    const addRowBtn = document.getElementById("add-row");
    const addColBtn = document.getElementById("add-col");
    const clearBtn = document.getElementById("clear-spreadsheet");
    const themeToggleBtn = document.getElementById("theme-toggle");
    const copyUrlBtn = document.getElementById("copy-url");
    const importCsvBtn = document.getElementById("import-csv");
    const importCsvInput = document.getElementById("import-csv-file");
    const exportCsvBtn = document.getElementById("export-csv");
    const presentBtn = document.getElementById("present-btn");
    const qrBtn = document.getElementById("qr-btn");
    const qrCloseBtn = document.getElementById("qr-close-btn");
    const qrBackdrop = document.querySelector("#qr-modal .modal-backdrop");
    const traceDepsBtn = document.getElementById("trace-deps-btn");

    if (addRowBtn) {
      addRowBtn.addEventListener("click", () => {
        addRow();
        scheduleFullSync();
        refreshDependencyLayer();
      });
    }
    if (addColBtn) {
      addColBtn.addEventListener("click", () => {
        addColumn();
        scheduleFullSync();
        refreshDependencyLayer();
      });
    }
    if (clearBtn) {
      clearBtn.addEventListener("click", () => {
        clearSpreadsheet();
        scheduleFullSync();
        refreshDependencyLayer();
      });
    }
    if (themeToggleBtn) {
      themeToggleBtn.addEventListener("click", () => ThemeManager.toggleTheme());
    }
    if (copyUrlBtn) {
      copyUrlBtn.addEventListener("click", () => QRCodeManager.copyURL());
    }
    if (importCsvBtn && importCsvInput) {
      importCsvBtn.addEventListener("click", function () {
        importCsvInput.click();
      });
    }
    if (importCsvInput) {
      importCsvInput.addEventListener("change", function (e) {
        const file = e.target.files && e.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = function () {
          CSVManager.importCSVText(String(reader.result || ""));
        };
        reader.onerror = function () {
          alert("Failed to read the CSV file.");
        };
        reader.readAsText(file);
        e.target.value = "";
      });
    }
    if (exportCsvBtn) {
      exportCsvBtn.addEventListener("click", () => CSVManager.downloadCSV());
    }
    if (presentBtn) {
      presentBtn.addEventListener("click", () => {
        recalculateFormulas();
        const { rows, cols } = getState();
        PresentationManager.start(data, formulas, {
          getCellValue,
          data,
          rows,
          cols,
        });
      });
    }
    if (qrBtn) {
      qrBtn.addEventListener("click", () => QRCodeManager.showQRModal());
    }
    if (qrCloseBtn) {
      qrCloseBtn.addEventListener("click", () => QRCodeManager.hideQRModal());
    }
    if (qrBackdrop) {
      qrBackdrop.addEventListener("click", () => QRCodeManager.hideQRModal());
    }
    if (traceDepsBtn) {
      traceDepsBtn.addEventListener("click", () => {
        DependencyTracer.init();
        const isActive = DependencyTracer.toggle();
        traceDepsBtn.classList.toggle("active", isActive);
        if (isActive) {
          DependencyTracer.draw(getFormulasArray());
          showToast("Visualizing formula dependencies", "info");
        } else {
          showToast("Logic visualization hidden", "info");
        }
      });
    }

    // Embed mode event listeners
    const generateEmbedBtn = document.getElementById("generate-embed");
    if (generateEmbedBtn) {
      generateEmbedBtn.addEventListener("click", () => UIModeManager.showEmbedModal());
    }

    const embedCopyBtn = document.getElementById("embed-copy-btn");
    if (embedCopyBtn) {
      embedCopyBtn.addEventListener("click", async () => {
        const textarea = document.getElementById("embed-code-textarea");
        try {
          await navigator.clipboard.writeText(textarea.value);
          showToast("Embed code copied to clipboard!", "success");
          embedCopyBtn.innerHTML = '<i class="fa-solid fa-check"></i> Copied!';
          setTimeout(() => {
            embedCopyBtn.innerHTML = '<i class="fa-solid fa-copy"></i> Copy to Clipboard';
          }, 2000);
        } catch (err) {
          textarea.select();
          document.execCommand("copy");
          showToast("Embed code copied!", "success");
        }
      });
    }

    const embedCloseBtn = document.getElementById("embed-close-btn");
    if (embedCloseBtn) {
      embedCloseBtn.addEventListener("click", () => UIModeManager.hideEmbedModal());
    }

    const embedBackdrop = document.querySelector("#embed-modal .modal-backdrop");
    if (embedBackdrop) {
      embedBackdrop.addEventListener("click", () => UIModeManager.hideEmbedModal());
    }

    // JSON editor modal event listeners
    const openJSONBtn = document.getElementById("open-json-btn");
    const jsonCloseBtn = document.getElementById("json-close-btn");
    const jsonCopyBtn = document.getElementById("json-copy-btn");
    const jsonBackdrop = document.querySelector("#json-modal .modal-backdrop");

    if (openJSONBtn) {
      openJSONBtn.addEventListener("click", () => JSONManager.openModal());
    }
    if (jsonCloseBtn) {
      jsonCloseBtn.addEventListener("click", () => JSONManager.closeModal());
    }
    if (jsonBackdrop) {
      jsonBackdrop.addEventListener("click", () => JSONManager.closeModal());
    }
    if (jsonCopyBtn) {
      jsonCopyBtn.addEventListener("click", () => JSONManager.copyJSONToClipboard());
    }

    // JSON hash tool modal event listeners
    const openHashToolBtn = document.getElementById("open-hash-tool-btn");
    const hashToolCloseBtn = document.getElementById("hash-tool-close-btn");
    const hashToolBackdrop = document.querySelector("#hash-tool-modal .modal-backdrop");

    if (openHashToolBtn) {
      openHashToolBtn.addEventListener("click", () => HashToolManager.openModal());
    }
    if (hashToolCloseBtn) {
      hashToolCloseBtn.addEventListener("click", () => HashToolManager.closeModal());
    }
    if (hashToolBackdrop) {
      hashToolBackdrop.addEventListener("click", () => HashToolManager.closeModal());
    }

    // P2P collaboration event listeners
    const p2pBtn = document.getElementById("p2p-btn");
    p2pUI.modal = document.getElementById("p2p-modal");
    p2pUI.startHostBtn = document.getElementById("p2p-start-host");
    p2pUI.joinBtn = document.getElementById("p2p-join-btn");
    p2pUI.copyIdBtn = document.getElementById("p2p-copy-id");
    p2pUI.idDisplay = document.getElementById("p2p-id-display");
    p2pUI.statusEl = document.getElementById("p2p-status");
    p2pUI.myIdInput = document.getElementById("p2p-my-id");
    p2pUI.remoteIdInput = document.getElementById("p2p-remote-id");
    const p2pCloseBtn = document.getElementById("p2p-close-btn");
    const p2pBackdrop = document.querySelector("#p2p-modal .modal-backdrop");

    if (p2pUI.startHostBtn) {
      p2pUI.startHostLabel = p2pUI.startHostBtn.innerHTML;
    }
    if (p2pUI.joinBtn) {
      p2pUI.joinLabel = p2pUI.joinBtn.innerHTML;
    }

    if (p2pBtn && p2pUI.modal) {
      p2pBtn.addEventListener("click", () => p2pUI.modal.classList.remove("hidden"));
    }
    if (p2pCloseBtn && p2pUI.modal) {
      p2pCloseBtn.addEventListener("click", () => p2pUI.modal.classList.add("hidden"));
    }
    if (p2pBackdrop && p2pUI.modal) {
      p2pBackdrop.addEventListener("click", () => p2pUI.modal.classList.add("hidden"));
    }

    if (p2pUI.startHostBtn) {
      p2pUI.startHostBtn.addEventListener("click", async () => {
        p2pUI.startHostBtn.disabled = true;
        p2pUI.startHostBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Connecting...';

        const started = await P2PManager.startHosting();
        if (!started) {
          p2pUI.startHostBtn.disabled = false;
          p2pUI.startHostBtn.innerHTML = p2pUI.startHostLabel;
          return;
        }
        p2pUI.startHostBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Hosting...';
      });
    }

    if (p2pUI.joinBtn) {
      p2pUI.joinBtn.addEventListener("click", async () => {
        const id = p2pUI.remoteIdInput ? p2pUI.remoteIdInput.value.trim() : "";
        if (!id) {
          showToast("Enter host ID to join", "warning");
          return;
        }

        p2pUI.joinBtn.disabled = true;
        p2pUI.joinBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Connecting...';

        const started = await P2PManager.joinSession(id);
        if (!started) {
          p2pUI.joinBtn.disabled = false;
          p2pUI.joinBtn.innerHTML = p2pUI.joinLabel;
        }
      });
    }

    if (p2pUI.copyIdBtn) {
      p2pUI.copyIdBtn.addEventListener("click", async () => {
        if (!p2pUI.myIdInput) return;
        try {
          await navigator.clipboard.writeText(p2pUI.myIdInput.value);
          showToast("ID copied!", "success");
        } catch (err) {
          p2pUI.myIdInput.select();
          document.execCommand("copy");
          showToast("ID copied!", "success");
        }
      });
    }

    // Read-only mode event listeners
    const toggleReadOnlyBtn = document.getElementById("toggle-readonly");
    if (toggleReadOnlyBtn) {
      toggleReadOnlyBtn.addEventListener("click", () => UIModeManager.toggleReadOnlyMode());
    }

    const enableEditingBtn = document.getElementById("enable-editing");
    if (enableEditingBtn) {
      enableEditingBtn.addEventListener("click", function () {
        if (isReadOnly) {
          UIModeManager.toggleReadOnlyMode();
        }
      });
    }

    // Handle browser back/forward
    window.addEventListener("hashchange", function () {
      loadStateFromURL();
      renderGrid();
      DependencyTracer.init();
      scheduleDependencyDraw();
    });
  }

  // Start when DOM is ready
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }

  // ========== Toolbar Scroll Logic ==========
  function initToolbarScroll() {
    const toolbar = document.querySelector(".toolbar");
    const scrollLeftBtn = document.getElementById("scroll-left");
    const scrollRightBtn = document.getElementById("scroll-right");

    if (!toolbar || !scrollLeftBtn || !scrollRightBtn) return;

    function updateScrollButtons() {
      // Check if content overflows
      const isOverflowing = toolbar.scrollWidth > toolbar.clientWidth;
      const scrollLeft = toolbar.scrollLeft;
      const maxScroll = toolbar.scrollWidth - toolbar.clientWidth;

      // Tolerance (fixes weird browser sub-pixel issues)
      const tolerance = 2;

      if (!isOverflowing) {
        scrollLeftBtn.classList.add("hidden");
        scrollRightBtn.classList.add("hidden");
        return;
      }

      // Show/Hide Left Button
      if (scrollLeft > tolerance) {
        scrollLeftBtn.classList.remove("hidden");
      } else {
        scrollLeftBtn.classList.add("hidden");
      }

      // Show/Hide Right Button
      if (scrollLeft < maxScroll - tolerance) {
        scrollRightBtn.classList.remove("hidden");
      } else {
        scrollRightBtn.classList.add("hidden");
      }
    }

    // Scroll amount for button clicks
    const scrollAmount = 200;

    scrollLeftBtn.addEventListener("click", () => {
      toolbar.scrollBy({ left: -scrollAmount, behavior: "smooth" });
    });

    scrollRightBtn.addEventListener("click", () => {
      toolbar.scrollBy({ left: scrollAmount, behavior: "smooth" });
    });

    // Listen for scroll events
    toolbar.addEventListener("scroll", () => {
      // Debounce the UI update slightly for performance
      requestAnimationFrame(updateScrollButtons);
    });

    // Update on resize
    window.addEventListener("resize", updateScrollButtons);

    // Initial check
    updateScrollButtons();
  }

  // Tools Modal Logic
  const toolsMenuBtn = document.getElementById("tools-menu-btn");
  const toolsModal = document.getElementById("tools-modal");
  const toolsCloseBtn = document.getElementById("tools-close-btn");

  if (toolsMenuBtn && toolsModal && toolsCloseBtn) {
    toolsMenuBtn.addEventListener("click", () => {
      toolsModal.classList.remove("hidden");
    });

    toolsCloseBtn.addEventListener("click", () => {
      toolsModal.classList.add("hidden");
    });

    toolsModal.addEventListener("click", (e) => {
      if (e.target === toolsModal || e.target.classList.contains("modal-backdrop")) {
        toolsModal.classList.add("hidden");
      }
    });

    // Close modal when a tool button is clicked (improved UX)
    toolsModal.querySelectorAll(".tool-item").forEach((btn) => {
      btn.addEventListener("click", () => {
        // Optional: Don't close for toggles if we want to toggle multiple times
        // But for now, let's close it to emulate a menu
        // Exception: maybe theme toggle?
        if (!btn.id.includes("toggle")) {
          toolsModal.classList.add("hidden");
        }
      });
    });
  }

  // Initialize all modules
  initToolbarScroll();
  TemplateManager.init();
})();
