// Dynamic Spreadsheet Web App
// Data persists in URL hash for easy sharing

// Import constants from ES6 module
import { DEBOUNCE_DELAY, DEFAULT_COLS, DEFAULT_ROWS, MAX_COLS, MAX_ROWS } from "./modules/constants.js";
import { buildRangeRef, FormulaDropdownManager, FormulaEvaluator, isValidFormula } from "./modules/formulaManager.js";
import { PasswordManager } from "./modules/passwordManager.js";
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
  normalizeCellStyles,
  normalizeColumnWidths,
  normalizeRowHeights,
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

  // Formula range selection mode (for clicking to select ranges like Google Sheets)
  let formulaEditMode = false; // true when typing a formula
  let formulaEditCell = null; // { row, col, element } of cell being edited
  let formulaRangeStart = null; // Start of range being selected
  let formulaRangeEnd = null; // End of range being selected
  let editingCell = null; // { row, col } when editing a cell's text

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
    const num = parseFloat(stripped);
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

  // Get current theme
  function isDarkMode() {
    return document.body.classList.contains("dark-mode");
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

  function csvEscape(value) {
    const text = String(value);
    const needsQuotes = /[",\r\n]/.test(text) || /^\s|\s$/.test(text);
    const escaped = text.replace(/"/g, '""');
    return needsQuotes ? `"${escaped}"` : escaped;
  }

  function buildCSV() {
    const { rows, cols } = getState();
    const lines = [];
    for (let r = 0; r < rows; r++) {
      const rowValues = [];
      for (let c = 0; c < cols; c++) {
        const raw = data[r][c];
        const text = extractPlainText(raw);
        rowValues.push(csvEscape(text));
      }
      lines.push(rowValues.join(","));
    }
    return lines.join("\r\n");
  }

  function downloadCSV() {
    recalculateFormulas();
    const csv = buildCSV();
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "spreadsheet.csv";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    showToast("CSV downloaded", "success");
  }

  function parseCSV(text) {
    if (!text) return [];

    const rows = [];
    let row = [];
    let field = "";
    let inQuotes = false;

    for (let i = 0; i < text.length; i++) {
      const char = text[i];

      if (inQuotes) {
        if (char === '"') {
          if (text[i + 1] === '"') {
            field += '"';
            i++;
          } else {
            inQuotes = false;
          }
        } else {
          field += char;
        }
        continue;
      }

      if (char === '"') {
        inQuotes = true;
      } else if (char === ",") {
        row.push(field);
        field = "";
      } else if (char === "\r") {
        if (text[i + 1] === "\n") {
          i++;
        }
        row.push(field);
        rows.push(row);
        row = [];
        field = "";
      } else if (char === "\n") {
        row.push(field);
        rows.push(row);
        row = [];
        field = "";
      } else {
        field += char;
      }
    }

    row.push(field);
    rows.push(row);

    if (rows.length && rows[0].length && rows[0][0]) {
      rows[0][0] = rows[0][0].replace(/^\uFEFF/, "");
    }

    if (/\r?\n$/.test(text)) {
      const lastRow = rows[rows.length - 1];
      if (lastRow && lastRow.length === 1 && lastRow[0] === "") {
        rows.pop();
      }
    }

    return rows;
  }

  function importCSVText(text) {
    const parsedRows = parseCSV(text);
    if (!parsedRows.length) {
      showToast("CSV file is empty", "error");
      return;
    }

    const maxCols = parsedRows.reduce((max, row) => Math.max(max, row.length), 0);
    const nextRows = Math.min(Math.max(parsedRows.length, 1), MAX_ROWS);
    const nextCols = Math.min(Math.max(maxCols, 1), MAX_COLS);

    const truncated = parsedRows.length > MAX_ROWS || maxCols > MAX_COLS;

    // Update state in rowColManager
    setState("rows", nextRows);
    setState("cols", nextCols);
    setState("colWidths", createDefaultColumnWidths(nextCols));
    setState("rowHeights", createDefaultRowHeights(nextRows));

    data = createEmptyData(nextRows, nextCols);
    formulas = createEmptyData(nextRows, nextCols);
    cellStyles = createEmptyCellStyles(nextRows, nextCols);

    for (let r = 0; r < nextRows; r++) {
      const sourceRow = Array.isArray(parsedRows[r]) ? parsedRows[r] : [];
      for (let c = 0; c < nextCols; c++) {
        const raw = sourceRow[c] !== undefined ? String(sourceRow[c]) : "";
        if (raw.startsWith("=")) {
          // Security: Validate formula against whitelist before storing
          if (isValidFormula(raw)) {
            formulas[r][c] = raw;
            data[r][c] = raw;
          } else {
            // Invalid formula pattern - treat as escaped text
            formulas[r][c] = "";
            data[r][c] = escapeHTML(raw);
          }
        } else {
          // Escape HTML in CSV values to prevent XSS
          data[r][c] = escapeHTML(raw);
        }
      }
    }

    renderGrid();
    recalculateFormulas();
    debouncedUpdateURL();

    if (truncated) {
      showToast("CSV imported (some data truncated due to size limits)", "warning");
    } else {
      showToast("CSV imported successfully", "success");
    }
  }

  // Helper functions moved to modules/urlManager.js

  // Minify state object keys for smaller URL payload
  // State helpers moved to modules/urlManager.js

  // ========== Serialization Codec (Wrapper for minify/expand + compression) ==========
  // Handles state serialization without encryption concerns
  // Codec, validation, and decode functions moved to modules/urlManager.js

  // Apply theme to body
  function applyTheme(theme) {
    if (theme === "dark") {
      document.body.classList.add("dark-mode");
    } else if (theme === "light") {
      document.body.classList.remove("dark-mode");
    }
    // Save to localStorage as well
    try {
      localStorage.setItem("spreadsheet-theme", theme);
    } catch (e) {}
  }

  // Build current state object (only includes non-empty/non-default values)
  function buildCurrentState() {
    const { rows, cols, colWidths, rowHeights } = getState();
    const stateObj = {
      rows,
      cols,
      theme: isDarkMode() ? "dark" : "light",
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

  // ========== Read-Only Mode Helper Functions ==========

  // Apply read-only mode UI state
  function applyReadOnlyMode() {
    document.body.classList.add("readonly-mode");

    // Set contentEditable on all cells
    const container = document.getElementById("spreadsheet");
    if (container) {
      container.querySelectorAll(".cell-content").forEach(cell => {
        cell.contentEditable = "false";
      });
    }

    // Update toggle button
    const toggleBtn = document.getElementById("toggle-readonly");
    if (toggleBtn) {
      toggleBtn.classList.add("active");
      const icon = toggleBtn.querySelector("i");
      if (icon) icon.className = "fa-solid fa-eye";
    }

    // Show banner
    const banner = document.getElementById("readonly-banner");
    if (banner) banner.classList.remove("hidden");
  }

  // Clear read-only mode UI state
  function clearReadOnlyMode() {
    document.body.classList.remove("readonly-mode");

    // Reset contentEditable
    const container = document.getElementById("spreadsheet");
    if (container) {
      container.querySelectorAll(".cell-content").forEach(cell => {
        cell.contentEditable = "true";
      });
    }

    // Update toggle button
    const toggleBtn = document.getElementById("toggle-readonly");
    if (toggleBtn) {
      toggleBtn.classList.remove("active");
      const icon = toggleBtn.querySelector("i");
      if (icon) icon.className = "fa-solid fa-pen-to-square";
    }

    // Hide banner
    const banner = document.getElementById("readonly-banner");
    if (banner) banner.classList.add("hidden");
  }

  // Apply embed mode UI state
  function applyEmbedMode() {
    document.body.classList.add("embed-mode");
    isReadOnly = true;
    applyReadOnlyMode();
  }

  // Clear embed mode UI state
  function clearEmbedMode() {
    document.body.classList.remove("embed-mode");
  }

  // Toggle read-only mode
  function toggleReadOnlyMode() {
    isReadOnly = !isReadOnly;
    applyReadOnlyState(isReadOnly);
    showToast(
      isReadOnly
        ? "Read-only mode enabled - Share this link for view-only access"
        : "Edit mode enabled",
      "success"
    );

    // Update URL immediately (no debounce)
    updateURL();
  }

  // Generate embed code
  async function generateEmbedCode() {
    if (isEmbedMode) {
      showToast("Already in embed mode. Share current URL instead.", "warning");
      return null;
    }

    const currentState = buildCurrentState();
    currentState.embed = 1;
    currentState.readOnly = 1;

    const encoded = await URLManager.encodeState(currentState, PasswordManager.getPassword());
    const embedURL = window.location.origin + window.location.pathname + "#" + encoded;

    if (embedURL.length > 2000) {
      showToast("Warning: Embed URL is very long", "warning");
    }

    return `<iframe
    src="${embedURL}"
    width="800"
    height="600"
    frameborder="0"
    style="border: 1px solid #e0e0e0; border-radius: 8px;"
    title="Embedded Spreadsheet">
</iframe>`;
  }

  // Show embed modal with generated code
  async function showEmbedModal() {
    const embedCode = await generateEmbedCode();
    if (!embedCode) return;

    const modal = document.getElementById("embed-modal");
    const textarea = document.getElementById("embed-code-textarea");

    textarea.value = embedCode;
    modal.classList.remove("hidden");
    textarea.select();
  }

  // Hide embed modal
  function hideEmbedModal() {
    const modal = document.getElementById("embed-modal");
    modal.classList.add("hidden");
  }

  // Handle input changes
  function handleInput(event) {
    const target = event.target;
    if (!target.classList.contains("cell-content")) return;

    // GUARD: Block in read-only mode
    if (isReadOnly) {
      event.preventDefault();
      target.blur();
      showToast("Cannot edit in read-only mode", "warning");
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

      if (rawValue.startsWith("=")) {
        // NOW evaluate the formula
        formulas[row][col] = rawValue;
        const result = FormulaEvaluator.evaluate(rawValue, { getCellValue, data, rows, cols });
        data[row][col] = String(result);
        target.innerText = String(result);

        // Recalculate all dependent formulas
        recalculateFormulas();
        debouncedUpdateURL();
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
        formulas[row][col] = rawValue;
        const result = FormulaEvaluator.evaluate(rawValue, { getCellValue, data, rows, cols });
        data[row][col] = String(result);
        target.innerText = String(result);
        recalculateFormulas();
        debouncedUpdateURL();

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

  function forEachTargetCell(callback) {
    const bounds = getSelectionBounds();
    if (bounds) {
      for (let r = bounds.minRow; r <= bounds.maxRow; r++) {
        for (let c = bounds.minCol; c <= bounds.maxCol; c++) {
          callback(r, c);
        }
      }
      return true;
    }

    const activeElement = document.activeElement;
    if (activeElement && activeElement.classList.contains("cell-content")) {
      const row = parseInt(activeElement.dataset.row, 10);
      const col = parseInt(activeElement.dataset.col, 10);
      if (!isNaN(row) && !isNaN(col)) {
        callback(row, col);
        return true;
      }
    }

    const { activeRow, activeCol } = getState();
    if (activeRow !== null && activeCol !== null) {
      callback(activeRow, activeCol);
      return true;
    }

    return false;
  }

  function applyAlignment(align) {
    const normalized = normalizeAlignment(align);
    if (!normalized) return;

    const updated = forEachTargetCell(function (row, col) {
      if (!cellStyles[row]) cellStyles[row] = [];
      if (!cellStyles[row][col]) cellStyles[row][col] = createEmptyCellStyle();
      cellStyles[row][col].align = normalized;

      const cellContent = getCellContentElement(row, col);
      if (cellContent) {
        cellContent.style.textAlign = normalized;
      }
    });

    if (updated) {
      debouncedUpdateURL();
    }
  }

  function applyCellBackground(color) {
    if (!isValidCSSColor(color)) return;

    const updated = forEachTargetCell(function (row, col) {
      if (!cellStyles[row]) cellStyles[row] = [];
      if (!cellStyles[row][col]) cellStyles[row][col] = createEmptyCellStyle();
      cellStyles[row][col].bg = color;

      const cell = getCellElement(row, col);
      if (cell) {
        if (color) {
          cell.style.setProperty("--cell-bg", color);
        } else {
          cell.style.removeProperty("--cell-bg");
        }
      }
    });

    if (updated) {
      debouncedUpdateURL();
    }
  }

  function applyCellTextColor(color) {
    if (!isValidCSSColor(color)) return;

    const updated = forEachTargetCell(function (row, col) {
      if (!cellStyles[row]) cellStyles[row] = [];
      if (!cellStyles[row][col]) cellStyles[row][col] = createEmptyCellStyle();
      cellStyles[row][col].color = color;

      const cellContent = getCellContentElement(row, col);
      if (cellContent) {
        cellContent.style.color = color;
      }
    });

    if (updated) {
      debouncedUpdateURL();
    }
  }

  function applyFontSize(size) {
    const normalized = normalizeFontSize(size);

    const updated = forEachTargetCell(function (row, col) {
      if (!cellStyles[row]) cellStyles[row] = [];
      if (!cellStyles[row][col]) cellStyles[row][col] = createEmptyCellStyle();
      cellStyles[row][col].fontSize = normalized;

      const cellContent = getCellContentElement(row, col);
      if (cellContent) {
        cellContent.style.fontSize = normalized ? `${normalized}px` : "";
      }
    });

    if (updated) {
      debouncedUpdateURL();
    }
  }

  // Apply text formatting using modern Selection/Range API (replaces deprecated execCommand)
  function applyFormat(command) {
    const selection = window.getSelection();
    if (!selection || !selection.rangeCount) return;

    const range = selection.getRangeAt(0);
    const selectedText = range.toString();
    if (!selectedText) return;

    // Create wrapper element based on command
    let wrapper;
    switch (command) {
      case "bold":
        wrapper = document.createElement("b");
        break;
      case "italic":
        wrapper = document.createElement("i");
        break;
      case "underline":
        wrapper = document.createElement("u");
        break;
      default:
        return;
    }

    // Only proceed if selection is within a cell-content element
    const activeElement = document.activeElement;
    if (!activeElement || !activeElement.classList.contains("cell-content")) return;

    try {
      // Wrap the selected content
      range.surroundContents(wrapper);
    } catch (e) {
      // surroundContents fails if selection crosses element boundaries
      // Fall back to extracting and re-inserting
      const fragment = range.extractContents();
      wrapper.appendChild(fragment);
      range.insertNode(wrapper);
    }

    // Update data after formatting
    const row = parseInt(activeElement.dataset.row, 10);
    const col = parseInt(activeElement.dataset.col, 10);
    const { rows, cols } = getState();
    if (!isNaN(row) && !isNaN(col) && row < rows && col < cols) {
      data[row][col] = sanitizeHTML(activeElement.innerHTML);
      debouncedUpdateURL();
    }
  }

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
          applyTheme(loadedState.theme);
        }

        // Load read-only mode
        applyReadOnlyState(loadedState.readOnly);

        // Load embed mode
        isEmbedMode = loadedState.embed || false;
        if (isEmbedMode) {
          applyEmbedMode();
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
    clearReadOnlyMode();
    clearEmbedMode();
    return true;
  }

  // Toggle dark/light mode
  function toggleTheme() {
    const body = document.body;
    const isDark = body.classList.toggle("dark-mode");
    const theme = isDark ? "dark" : "light";

    // Save preference to localStorage
    try {
      localStorage.setItem("spreadsheet-theme", theme);
    } catch (e) {
      // localStorage not available
    }

    // Update URL with new theme
    debouncedUpdateURL();
  }

  // Load saved theme preference
  function loadTheme() {
    try {
      const savedTheme = localStorage.getItem("spreadsheet-theme");
      if (savedTheme === "dark") {
        document.body.classList.add("dark-mode");
      } else if (savedTheme === "light") {
        document.body.classList.remove("dark-mode");
      } else {
        // Check system preference if no saved preference
        if (window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches) {
          document.body.classList.add("dark-mode");
        }
      }
    } catch (e) {
      // localStorage not available
    }
  }

  // Copy URL to clipboard
  function copyURL() {
    const url = window.location.href;
    const copyBtn = document.getElementById("copy-url");

    navigator.clipboard
      .writeText(url)
      .then(function () {
        // Show success feedback
        showToast("Link copied to clipboard!", "success");
        if (copyBtn) {
          copyBtn.classList.add("copied");
          const icon = copyBtn.querySelector("i");
          if (icon) {
            icon.className = "fa-solid fa-check";
          }

          // Reset after 2 seconds
          setTimeout(function () {
            copyBtn.classList.remove("copied");
            if (icon) {
              icon.className = "fa-solid fa-copy";
            }
          }, 2000);
        }
      })
      .catch(function (err) {
        console.error("Failed to copy URL:", err);
        // Fallback for older browsers
        const textArea = document.createElement("textarea");
        textArea.value = url;
        textArea.style.position = "fixed";
        textArea.style.opacity = "0";
        document.body.appendChild(textArea);
        textArea.select();
        try {
          document.execCommand("copy");
          if (copyBtn) {
            copyBtn.classList.add("copied");
            setTimeout(function () {
              copyBtn.classList.remove("copied");
            }, 2000);
          }
        } catch (e) {
          console.error("Fallback copy failed:", e);
        }
        document.body.removeChild(textArea);
      });
  }

  // ========== Password/Encryption Modal Functions ==========
  // Handled by PasswordManager module

  /**
   * Apply read-only state to the spreadsheet
   * @param {boolean} readOnlyFlag - Whether to enable read-only mode
   */
  function applyReadOnlyState(readOnlyFlag) {
    // If in embed mode, always stay read-only
    if (isEmbedMode) {
      isReadOnly = true;
      applyReadOnlyMode();
      return;
    }

    isReadOnly = readOnlyFlag || false;
    if (isReadOnly) {
      applyReadOnlyMode();
    } else {
      clearReadOnlyMode();
    }
  }

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
      applyTheme(loadedState.theme);
    }

    // Apply read-only mode if present
    applyReadOnlyState(loadedState.readOnly);
  }

  // Initialize the app
  function init() {
    // Set up callbacks for rowColManager module
    setCallbacks({
      debouncedUpdateURL,
      recalculateFormulas,
      getDataArray: () => data,
      setDataArray: (newData) => {
        data = newData;
      },
      getFormulasArray: () => formulas,
      setFormulasArray: (newFormulas) => {
        formulas = newFormulas;
      },
      getCellStylesArray: () => cellStyles,
      setCellStylesArray: (newCellStyles) => {
        cellStyles = newCellStyles;
      },
      PasswordManager,
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
    });

    // Load theme preference first (before any rendering)
    loadTheme();

    // Load any existing state from URL
    loadStateFromURL();

    // Render the grid
    renderGrid();

    // Initialize Formula Dropdown
    FormulaDropdownManager.init(applyFormulaSuggestion);

    // Initialize Password Manager
    PasswordManager.init({
      decryptAndDecode: URLManager.decryptAndDecode,
      onDecryptSuccess: (state) => {
        applyLoadedState(state);
        renderGrid();
      },
      updateURL: updateURL,
      showToast: showToast,
      validateState: validateAndNormalizeState,
    });

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
      });
    }

    window.addEventListener("resize", function () {
      if (FormulaDropdownManager.isOpen() && FormulaDropdownManager.anchor) {
        FormulaDropdownManager.position(FormulaDropdownManager.anchor);
      }
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
        applyFormat("bold");
      });
    }
    if (italicBtn) {
      italicBtn.addEventListener("mousedown", function (e) {
        e.preventDefault();
        applyFormat("italic");
      });
    }
    if (underlineBtn) {
      underlineBtn.addEventListener("mousedown", function (e) {
        e.preventDefault();
        applyFormat("underline");
      });
    }
    if (alignLeftBtn) {
      alignLeftBtn.addEventListener("mousedown", function (e) {
        e.preventDefault();
        applyAlignment("left");
      });
    }
    if (alignCenterBtn) {
      alignCenterBtn.addEventListener("mousedown", function (e) {
        e.preventDefault();
        applyAlignment("center");
      });
    }
    if (alignRightBtn) {
      alignRightBtn.addEventListener("mousedown", function (e) {
        e.preventDefault();
        applyAlignment("right");
      });
    }
    if (cellBgPicker) {
      cellBgPicker.addEventListener("input", function (e) {
        applyCellBackground(e.target.value);
      });
    }
    if (cellTextColorPicker) {
      cellTextColorPicker.addEventListener("input", function (e) {
        applyCellTextColor(e.target.value);
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
        applyFontSize(button.dataset.size);
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

    if (addRowBtn) {
      addRowBtn.addEventListener("click", addRow);
    }
    if (addColBtn) {
      addColBtn.addEventListener("click", addColumn);
    }
    if (clearBtn) {
      clearBtn.addEventListener("click", clearSpreadsheet);
    }
    if (themeToggleBtn) {
      themeToggleBtn.addEventListener("click", toggleTheme);
    }
    if (copyUrlBtn) {
      copyUrlBtn.addEventListener("click", copyURL);
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
          importCSVText(String(reader.result || ""));
        };
        reader.onerror = function () {
          alert("Failed to read the CSV file.");
        };
        reader.readAsText(file);
        e.target.value = "";
      });
    }
    if (exportCsvBtn) {
      exportCsvBtn.addEventListener("click", downloadCSV);
    }

    // Embed mode event listeners
    const generateEmbedBtn = document.getElementById("generate-embed");
    if (generateEmbedBtn) {
      generateEmbedBtn.addEventListener("click", showEmbedModal);
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
      embedCloseBtn.addEventListener("click", hideEmbedModal);
    }

    const embedBackdrop = document.querySelector("#embed-modal .modal-backdrop");
    if (embedBackdrop) {
      embedBackdrop.addEventListener("click", hideEmbedModal);
    }

    // Read-only mode event listeners
    const toggleReadOnlyBtn = document.getElementById("toggle-readonly");
    if (toggleReadOnlyBtn) {
      toggleReadOnlyBtn.addEventListener("click", toggleReadOnlyMode);
    }

    const enableEditingBtn = document.getElementById("enable-editing");
    if (enableEditingBtn) {
      enableEditingBtn.addEventListener("click", function() {
        if (isReadOnly) {
          toggleReadOnlyMode();
        }
      });
    }

    // Password/Encryption event listeners
    // Password/Encryption event listeners handled by PasswordManager
    /*
    const lockBtn = document.getElementById("lock-btn");
    const modalCancel = document.getElementById("modal-cancel");
    const modalSubmit = document.getElementById("modal-submit");
    const modalBackdrop = document.querySelector(".modal-backdrop");
    const passwordInput = document.getElementById("password-input");

    if (lockBtn) {
      lockBtn.addEventListener("click", handleLockButtonClick);
    }
    if (modalCancel) {
      modalCancel.addEventListener("click", hidePasswordModal);
    }
    if (modalSubmit) {
      modalSubmit.addEventListener("click", handleModalSubmit);
    }
    if (modalBackdrop) {
      modalBackdrop.addEventListener("click", function () {
        // Only allow closing if not in decrypt mode (user must enter password)
        if (modalMode !== "decrypt") {
          hidePasswordModal();
        }
      });
    }
    if (passwordInput) {
      passwordInput.addEventListener("keydown", function (e) {
        if (e.key === "Enter") {
          e.preventDefault();
          const confirmInput = document.getElementById("password-confirm");
          // If confirm is visible and empty, focus it; otherwise submit
          if (confirmInput && confirmInput.style.display !== "none" && !confirmInput.value) {
            confirmInput.focus();
          } else {
            handleModalSubmit();
          }
        }
      });
    }
    const passwordConfirm = document.getElementById("password-confirm");
    if (passwordConfirm) {
      passwordConfirm.addEventListener("keydown", function (e) {
        if (e.key === "Enter") {
          e.preventDefault();
          handleModalSubmit();
        }
      });
    }
    */

    // Handle browser back/forward
    window.addEventListener("hashchange", function () {
      loadStateFromURL();
      renderGrid();
    });
  }

  // Start when DOM is ready
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }

  // Expose audit function globally for testing URL sizes
  window.auditURLSize = async function () {
    const scenarios = [
      { name: "Empty 30x15", rows: 30, cols: 15, fill: null },
      { name: "Short text (A1, B1...)", rows: 30, cols: 15, fill: "coords" },
      { name: "Medium text (10 chars)", rows: 30, cols: 15, fill: "medium" },
      { name: "Numbers only", rows: 30, cols: 15, fill: "numbers" },
      { name: "With formulas", rows: 30, cols: 15, fill: "formulas" },
    ];

    console.log("=== URL Size Audit ===");
    console.log("Max grid: 30 rows x 15 cols = 450 cells\n");

    for (const scenario of scenarios) {
      const testData = [];
      for (let r = 0; r < scenario.rows; r++) {
        const row = [];
        for (let c = 0; c < scenario.cols; c++) {
          const colLetter = String.fromCharCode(65 + c);
          const cellRef = colLetter + (r + 1);
          switch (scenario.fill) {
            case "coords":
              row.push(cellRef);
              break;
            case "medium":
              row.push("Text_" + cellRef.padEnd(5, "X"));
              break;
            case "numbers":
              row.push(String(Math.floor(Math.random() * 10000)));
              break;
            case "formulas":
              // Only put formulas in some cells to simulate realistic usage
              if (r > 0 && c === 0) {
                row.push("=SUM(B" + (r + 1) + ":O" + (r + 1) + ")");
              } else {
                row.push(String(Math.floor(Math.random() * 100)));
              }
              break;
            default:
              row.push("");
          }
        }
        testData.push(row);
      }

      const state = {
        rows: scenario.rows,
        cols: scenario.cols,
        theme: "light",
        data: testData,
      };

      const json = JSON.stringify(state);
      const compressed = LZString.compressToEncodedURIComponent(json);

      console.log(`${scenario.name}:`);
      console.log(`  JSON size: ${json.length.toLocaleString()} chars`);
      console.log(`  Compressed: ${compressed.length.toLocaleString()} chars`);
      console.log(`  Compression ratio: ${((1 - compressed.length / json.length) * 100).toFixed(1)}%`);
      console.log("");
    }

    console.log("=== Thresholds ===");
    console.log("< 2,000 chars: OK (safe for all browsers)");
    console.log("2,000-4,000: Warning (some older browsers may truncate)");
    console.log("4,000-8,000: Caution (URL shorteners may fail)");
    console.log("> 8,000: Critical (some browsers may fail)");
  };
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

  // Initialize all modules
  initToolbarScroll();
})();
