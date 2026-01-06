// Row and Column Manager Module
// Handles grid rendering, selection, resizing, and row/column operations

import {
  ACTIVE_HEADER_CLASS,
  DEFAULT_COL_WIDTH,
  DEFAULT_COLS,
  DEFAULT_ROW_HEIGHT,
  DEFAULT_ROWS,
  HEADER_ROW_HEIGHT,
  MAX_COLS,
  MAX_ROWS,
  MIN_COL_WIDTH,
  MIN_ROW_HEIGHT,
  ROW_HEADER_WIDTH
} from "./constants.js";

import { colToLetter } from "./formulaManager.js";
import { sanitizeHTML } from "./security.js";
import { showToast } from "./toastManager.js";
import {
  createEmptyData,
  createEmptyCellStyle,
  createEmptyCellStyles,
  createDefaultColumnWidths,
  createDefaultRowHeights
} from "./urlManager.js";

// ========== State ==========
const state = {
  rows: DEFAULT_ROWS,
  cols: DEFAULT_COLS,
  colWidths: createDefaultColumnWidths(DEFAULT_COLS),
  rowHeights: createDefaultRowHeights(DEFAULT_ROWS),
  selectionStart: null,
  selectionEnd: null,
  isSelecting: false,
  hoverRow: null,
  hoverCol: null,
  activeRow: null,
  activeCol: null,
  resizeState: null
};

// ========== Callbacks ==========
let callbacks = {
  debouncedUpdateURL: null,
  recalculateFormulas: null,
  getDataArray: null,
  setDataArray: null,
  getFormulasArray: null,
  setFormulasArray: null,
  getCellStylesArray: null,
  setCellStylesArray: null,
  PasswordManager: null,
  // Formula mode callbacks
  getFormulaEditMode: null,
  getFormulaEditCell: null,
  setFormulaRangeStart: null,
  setFormulaRangeEnd: null,
  getFormulaRangeStart: null,
  getFormulaRangeEnd: null,
  buildRangeRef: null,
  insertTextAtCursor: null,
  FormulaDropdownManager: null
};

// ========== State Accessors ==========
export function getState() {
  return state;
}

export function setState(key, value) {
  if (key in state) {
    state[key] = value;
  }
}

export function setCallbacks(cbs) {
  callbacks = { ...callbacks, ...cbs };
}

// ========== UI Update ==========
export function updateUI() {
  const addRowBtn = document.getElementById("add-row");
  const addColBtn = document.getElementById("add-col");
  const gridSizeEl = document.getElementById("grid-size");

  if (addRowBtn) {
    addRowBtn.disabled = state.rows >= MAX_ROWS;
  }
  if (addColBtn) {
    addColBtn.disabled = state.cols >= MAX_COLS;
  }
  if (gridSizeEl) {
    gridSizeEl.textContent = `${state.rows} × ${state.cols}`;
  }
}

// ========== Grid Template ==========
export function applyGridTemplate() {
  const container = document.getElementById("spreadsheet");
  if (!container) return;

  const columnSizes = state.colWidths.map((width) => `${width}px`).join(" ");
  const rowSizes = state.rowHeights.map((height) => `${height}px`).join(" ");

  container.style.gridTemplateColumns = `${ROW_HEADER_WIDTH}px ${columnSizes}`;
  container.style.gridTemplateRows = `${HEADER_ROW_HEIGHT}px ${rowSizes}`;
}

// ========== Render Grid ==========
export function renderGrid() {
  const container = document.getElementById("spreadsheet");
  if (!container) return;

  const data = callbacks.getDataArray ? callbacks.getDataArray() : [];
  const cellStyles = callbacks.getCellStylesArray ? callbacks.getCellStylesArray() : [];

  // Clear selection when grid is re-rendered
  state.selectionStart = null;
  state.selectionEnd = null;
  state.isSelecting = false;
  state.hoverRow = null;
  state.hoverCol = null;

  container.innerHTML = "";
  applyGridTemplate();

  // Corner cell (empty)
  const corner = document.createElement("div");
  corner.className = "corner-cell";
  container.appendChild(corner);

  // Column headers (A-Z) - sticky to top
  for (let col = 0; col < state.cols; col++) {
    const header = document.createElement("div");
    header.className = "header-cell col-header";
    header.textContent = colToLetter(col);
    header.dataset.col = col;
    const colResize = document.createElement("div");
    colResize.className = "resize-handle col-resize";
    colResize.dataset.col = col;
    colResize.setAttribute("aria-hidden", "true");
    header.appendChild(colResize);
    container.appendChild(header);
  }

  // Rows
  for (let row = 0; row < state.rows; row++) {
    // Row header (1, 2, 3...) - sticky to left
    const rowHeader = document.createElement("div");
    rowHeader.className = "header-cell row-header";
    rowHeader.textContent = row + 1;
    rowHeader.dataset.row = row;
    const rowResize = document.createElement("div");
    rowResize.className = "resize-handle row-resize";
    rowResize.dataset.row = row;
    rowResize.setAttribute("aria-hidden", "true");
    rowHeader.appendChild(rowResize);
    container.appendChild(rowHeader);

    // Data cells
    for (let col = 0; col < state.cols; col++) {
      const cell = document.createElement("div");
      cell.className = "cell";

      const contentDiv = document.createElement("div");
      contentDiv.className = "cell-content";
      contentDiv.contentEditable = "true";
      contentDiv.dataset.row = row;
      contentDiv.dataset.col = col;
      contentDiv.innerHTML = sanitizeHTML(data[row] ? data[row][col] || "" : "");
      contentDiv.setAttribute("aria-label", `Cell ${colToLetter(col)}${row + 1}`);

      const style = cellStyles[row] ? cellStyles[row][col] : null;
      if (style) {
        contentDiv.style.textAlign = style.align || "";
        contentDiv.style.color = style.color || "";
        contentDiv.style.fontSize = style.fontSize ? `${style.fontSize}px` : "";
        if (style.bg) {
          cell.style.setProperty("--cell-bg", style.bg);
        } else {
          cell.style.removeProperty("--cell-bg");
        }
      } else {
        contentDiv.style.textAlign = "";
        contentDiv.style.color = "";
        contentDiv.style.fontSize = "";
        cell.style.removeProperty("--cell-bg");
      }

      cell.appendChild(contentDiv);
      container.appendChild(cell);
    }
  }

  updateUI();
}

// ========== Header Highlighting ==========
export function clearActiveHeaders() {
  if (state.activeRow !== null) {
    const rowHeader = document.querySelector(`.row-header[data-row="${state.activeRow}"]`);
    if (rowHeader) rowHeader.classList.remove(ACTIVE_HEADER_CLASS);
  }
  if (state.activeCol !== null) {
    const colHeader = document.querySelector(`.col-header[data-col="${state.activeCol}"]`);
    if (colHeader) colHeader.classList.remove(ACTIVE_HEADER_CLASS);
  }
  state.activeRow = null;
  state.activeCol = null;
}

export function setActiveHeaders(row, col) {
  if (state.activeRow === row && state.activeCol === col) return;
  clearActiveHeaders();
  state.activeRow = row;
  state.activeCol = col;

  const rowHeader = document.querySelector(`.row-header[data-row="${row}"]`);
  if (rowHeader) rowHeader.classList.add(ACTIVE_HEADER_CLASS);

  const colHeader = document.querySelector(`.col-header[data-col="${col}"]`);
  if (colHeader) colHeader.classList.add(ACTIVE_HEADER_CLASS);
}

export function setActiveHeadersForRange(minRow, maxRow, minCol, maxCol) {
  // Clear existing header highlights
  document.querySelectorAll(`.${ACTIVE_HEADER_CLASS}`).forEach((el) => {
    el.classList.remove(ACTIVE_HEADER_CLASS);
  });

  // Highlight all row headers in range
  for (let r = minRow; r <= maxRow; r++) {
    const rowHeader = document.querySelector(`.row-header[data-row="${r}"]`);
    if (rowHeader) rowHeader.classList.add(ACTIVE_HEADER_CLASS);
  }

  // Highlight all column headers in range
  for (let c = minCol; c <= maxCol; c++) {
    const colHeader = document.querySelector(`.col-header[data-col="${c}"]`);
    if (colHeader) colHeader.classList.add(ACTIVE_HEADER_CLASS);
  }

  // Update active row/col tracking
  state.activeRow = minRow;
  state.activeCol = minCol;
}

// ========== Selection Functions ==========
export function getSelectionBounds() {
  if (!state.selectionStart || !state.selectionEnd) return null;
  return {
    minRow: Math.min(state.selectionStart.row, state.selectionEnd.row),
    maxRow: Math.max(state.selectionStart.row, state.selectionEnd.row),
    minCol: Math.min(state.selectionStart.col, state.selectionEnd.col),
    maxCol: Math.max(state.selectionStart.col, state.selectionEnd.col),
  };
}

export function hasMultiSelection() {
  if (!state.selectionStart || !state.selectionEnd) return false;
  return state.selectionStart.row !== state.selectionEnd.row || state.selectionStart.col !== state.selectionEnd.col;
}

export function clearSelection() {
  state.selectionStart = null;
  state.selectionEnd = null;
  state.isSelecting = false;

  const container = document.getElementById("spreadsheet");
  if (!container) return;

  // Remove selection classes from all cells
  container.querySelectorAll(".cell-selected").forEach((cell) => {
    cell.classList.remove("cell-selected", "selection-top", "selection-bottom", "selection-left", "selection-right");
  });

  // Remove selecting mode from container
  container.classList.remove("selecting");
}

export function updateSelectionVisuals() {
  const container = document.getElementById("spreadsheet");
  if (!container) return;

  const bounds = getSelectionBounds();
  if (!bounds) {
    clearSelection();
    return;
  }

  // Clear previous selection classes
  container.querySelectorAll(".cell-selected").forEach((cell) => {
    cell.classList.remove("cell-selected", "selection-top", "selection-bottom", "selection-left", "selection-right");
  });

  const { minRow, maxRow, minCol, maxCol } = bounds;

  // Apply selection classes to cells in range
  for (let r = minRow; r <= maxRow; r++) {
    for (let c = minCol; c <= maxCol; c++) {
      const cellContent = container.querySelector(`.cell-content[data-row="${r}"][data-col="${c}"]`);
      if (cellContent && cellContent.parentElement) {
        const cell = cellContent.parentElement;
        cell.classList.add("cell-selected");

        // Add border classes for outer edges
        if (r === minRow) cell.classList.add("selection-top");
        if (r === maxRow) cell.classList.add("selection-bottom");
        if (c === minCol) cell.classList.add("selection-left");
        if (c === maxCol) cell.classList.add("selection-right");
      }
    }
  }

  // Highlight headers for the entire range
  setActiveHeadersForRange(minRow, maxRow, minCol, maxCol);
}

export function clearSelectedCells() {
  if (!state.selectionStart || !state.selectionEnd) return;

  const bounds = getSelectionBounds();
  const container = document.getElementById("spreadsheet");
  const data = callbacks.getDataArray ? callbacks.getDataArray() : [];
  const formulas = callbacks.getFormulasArray ? callbacks.getFormulasArray() : [];

  for (let r = bounds.minRow; r <= bounds.maxRow; r++) {
    for (let c = bounds.minCol; c <= bounds.maxCol; c++) {
      // Clear data and formula
      if (data[r]) data[r][c] = "";
      if (formulas[r]) formulas[r][c] = "";

      // Update DOM
      const cell = container.querySelector(`.cell-content[data-row="${r}"][data-col="${c}"]`);
      if (cell) {
        cell.innerHTML = "";
      }
    }
  }

  if (callbacks.recalculateFormulas) callbacks.recalculateFormulas();
  if (callbacks.debouncedUpdateURL) callbacks.debouncedUpdateURL();
}

// ========== Hover Functions ==========
export function getCellContentFromTarget(target) {
  if (!(target instanceof Element)) return null;

  if (target.classList.contains("cell-content")) {
    return target;
  }
  if (target.classList.contains("cell")) {
    return target.querySelector(".cell-content");
  }
  return target.closest(".cell-content");
}

export function addHoverRow(row) {
  const container = document.getElementById("spreadsheet");
  if (!container) return;

  container.querySelectorAll(`.cell-content[data-row="${row}"]`).forEach((cellContent) => {
    if (cellContent.parentElement) {
      cellContent.parentElement.classList.add("hover-row");
    }
  });

  const rowHeader = container.querySelector(`.row-header[data-row="${row}"]`);
  if (rowHeader) {
    rowHeader.classList.add("header-hover");
  }
}

export function addHoverCol(col) {
  const container = document.getElementById("spreadsheet");
  if (!container) return;

  container.querySelectorAll(`.cell-content[data-col="${col}"]`).forEach((cellContent) => {
    if (cellContent.parentElement) {
      cellContent.parentElement.classList.add("hover-col");
    }
  });

  const colHeader = container.querySelector(`.col-header[data-col="${col}"]`);
  if (colHeader) {
    colHeader.classList.add("header-hover");
  }
}

export function removeHoverRow(row) {
  const container = document.getElementById("spreadsheet");
  if (!container) return;

  container.querySelectorAll(`.cell-content[data-row="${row}"]`).forEach((cellContent) => {
    if (cellContent.parentElement) {
      cellContent.parentElement.classList.remove("hover-row");
    }
  });

  const rowHeader = container.querySelector(`.row-header[data-row="${row}"]`);
  if (rowHeader) {
    rowHeader.classList.remove("header-hover");
  }
}

export function removeHoverCol(col) {
  const container = document.getElementById("spreadsheet");
  if (!container) return;

  container.querySelectorAll(`.cell-content[data-col="${col}"]`).forEach((cellContent) => {
    if (cellContent.parentElement) {
      cellContent.parentElement.classList.remove("hover-col");
    }
  });

  const colHeader = container.querySelector(`.col-header[data-col="${col}"]`);
  if (colHeader) {
    colHeader.classList.remove("header-hover");
  }
}

export function clearHoverHighlights() {
  if (state.hoverRow !== null) {
    removeHoverRow(state.hoverRow);
  }
  if (state.hoverCol !== null) {
    removeHoverCol(state.hoverCol);
  }
  state.hoverRow = null;
  state.hoverCol = null;
}

export function setHoverHighlight(row, col) {
  if (row === state.hoverRow && col === state.hoverCol) return;

  if (state.hoverRow !== null && state.hoverRow !== row) {
    removeHoverRow(state.hoverRow);
  }
  if (state.hoverCol !== null && state.hoverCol !== col) {
    removeHoverCol(state.hoverCol);
  }

  state.hoverRow = row;
  state.hoverCol = col;

  if (state.hoverRow !== null) {
    addHoverRow(state.hoverRow);
  }
  if (state.hoverCol !== null) {
    addHoverCol(state.hoverCol);
  }
}

export function updateHoverFromTarget(target) {
  if (state.isSelecting || hasMultiSelection()) {
    clearHoverHighlights();
    return;
  }

  const cellContent = getCellContentFromTarget(target);
  if (!cellContent || !cellContent.classList.contains("cell-content")) {
    clearHoverHighlights();
    return;
  }

  const row = parseInt(cellContent.dataset.row, 10);
  const col = parseInt(cellContent.dataset.col, 10);
  if (isNaN(row) || isNaN(col)) {
    clearHoverHighlights();
    return;
  }

  setHoverHighlight(row, col);
}

// ========== Cell Positioning ==========
export function getCellFromPoint(x, y) {
  const element = document.elementFromPoint(x, y);
  if (!element) return null;

  // Check if it's a cell-content or its parent cell
  let cellContent = element;
  if (element.classList.contains("cell")) {
    cellContent = element.querySelector(".cell-content");
  }

  if (!cellContent || !cellContent.classList.contains("cell-content")) return null;

  const row = parseInt(cellContent.dataset.row, 10);
  const col = parseInt(cellContent.dataset.col, 10);

  if (isNaN(row) || isNaN(col)) return null;
  return { row, col };
}

export function getCellContentElement(row, col) {
  return document.querySelector(`.cell-content[data-row="${row}"][data-col="${col}"]`);
}

export function getCellElement(row, col) {
  const cellContent = getCellContentElement(row, col);
  return cellContent ? cellContent.parentElement : null;
}

export function focusCellAt(row, col) {
  const cellContent = getCellContentElement(row, col);
  if (!cellContent) return null;
  cellContent.focus();
  cellContent.scrollIntoView({ block: "nearest", inline: "nearest" });
  return cellContent;
}

// ========== Resize Handlers ==========
export function handleResizeStart(event) {
  if (event.button !== 0) return;
  if (!(event.target instanceof Element)) return;

  const handle = event.target.closest(".resize-handle");
  if (!handle) return;

  const isColResize = handle.classList.contains("col-resize");
  const indexValue = isColResize ? handle.dataset.col : handle.dataset.row;
  const index = parseInt(indexValue, 10);
  if (isNaN(index)) return;

  event.preventDefault();
  event.stopImmediatePropagation();

  state.resizeState = {
    type: isColResize ? "col" : "row",
    index,
    startX: event.clientX,
    startY: event.clientY,
    startSize: isColResize ? state.colWidths[index] || DEFAULT_COL_WIDTH : state.rowHeights[index] || DEFAULT_ROW_HEIGHT,
  };

  state.isSelecting = false;
  document.body.classList.add("resizing");
  document.body.style.cursor = isColResize ? "col-resize" : "row-resize";

  document.addEventListener("mousemove", handleResizeMove);
  document.addEventListener("mouseup", handleResizeEnd);
}

export function handleResizeMove(event) {
  if (!state.resizeState) return;

  if (state.resizeState.type === "col") {
    const delta = event.clientX - state.resizeState.startX;
    const nextWidth = Math.max(MIN_COL_WIDTH, state.resizeState.startSize + delta);
    state.colWidths[state.resizeState.index] = nextWidth;
  } else {
    const delta = event.clientY - state.resizeState.startY;
    const nextHeight = Math.max(MIN_ROW_HEIGHT, state.resizeState.startSize + delta);
    state.rowHeights[state.resizeState.index] = nextHeight;
  }

  applyGridTemplate();
}

export function handleResizeEnd() {
  if (!state.resizeState) return;

  document.removeEventListener("mousemove", handleResizeMove);
  document.removeEventListener("mouseup", handleResizeEnd);
  document.body.classList.remove("resizing");
  document.body.style.cursor = "";
  state.resizeState = null;
  if (callbacks.debouncedUpdateURL) callbacks.debouncedUpdateURL();
}

// ========== Mouse Event Handlers ==========
export function handleMouseDown(event) {
  // Only handle left mouse button
  if (event.button !== 0) return;

  const target = event.target;

  // Check if clicking on a cell
  let cellContent = target;
  if (target.classList.contains("cell")) {
    cellContent = target.querySelector(".cell-content");
  }

  if (!cellContent || !cellContent.classList.contains("cell-content")) {
    // Clicked outside cells - clear selection
    clearSelection();
    return;
  }

  const row = parseInt(cellContent.dataset.row, 10);
  const col = parseInt(cellContent.dataset.col, 10);

  if (isNaN(row) || isNaN(col)) return;

  const container = document.getElementById("spreadsheet");

  // If in formula edit mode and clicking on a different cell
  const formulaEditMode = callbacks.getFormulaEditMode ? callbacks.getFormulaEditMode() : false;
  const formulaEditCell = callbacks.getFormulaEditCell ? callbacks.getFormulaEditCell() : null;

  if (formulaEditMode && formulaEditCell) {
    // Don't process clicks on the formula cell itself
    if (row !== formulaEditCell.row || col !== formulaEditCell.col) {
      event.preventDefault();
      event.stopPropagation();

      if (callbacks.FormulaDropdownManager) {
        callbacks.FormulaDropdownManager.hide();
      }

      // Start range selection for formula
      if (callbacks.setFormulaRangeStart) callbacks.setFormulaRangeStart({ row, col });
      if (callbacks.setFormulaRangeEnd) callbacks.setFormulaRangeEnd({ row, col });
      state.isSelecting = true;
      clearHoverHighlights();

      // Show visual selection
      state.selectionStart = { row, col };
      state.selectionEnd = { row, col };
      updateSelectionVisuals();

      if (container) {
        container.classList.add("selecting");
      }

      return;
    }
  }

  // Shift+click: extend selection from anchor
  if (event.shiftKey && state.selectionStart) {
    state.selectionEnd = { row, col };
    updateSelectionVisuals();
    event.preventDefault();
    return;
  }

  // Start new selection
  state.selectionStart = { row, col };
  state.selectionEnd = { row, col };
  state.isSelecting = true;
  clearHoverHighlights();

  if (container) {
    container.classList.add("selecting");
  }

  updateSelectionVisuals();
}

export function handleMouseMove(event) {
  if (!state.isSelecting) {
    updateHoverFromTarget(event.target);
    return;
  }

  const cellCoords = getCellFromPoint(event.clientX, event.clientY);
  if (!cellCoords) return;

  // Only update if position changed
  if (state.selectionEnd && cellCoords.row === state.selectionEnd.row && cellCoords.col === state.selectionEnd.col) {
    return;
  }

  // If in formula edit mode, update formula range
  const formulaEditMode = callbacks.getFormulaEditMode ? callbacks.getFormulaEditMode() : false;
  const formulaRangeStart = callbacks.getFormulaRangeStart ? callbacks.getFormulaRangeStart() : null;

  if (formulaEditMode && formulaRangeStart) {
    if (callbacks.setFormulaRangeEnd) callbacks.setFormulaRangeEnd(cellCoords);
    state.selectionEnd = cellCoords;
    updateSelectionVisuals();
    event.preventDefault();
    return;
  }

  state.selectionEnd = cellCoords;
  updateSelectionVisuals();

  // Prevent text selection during drag
  event.preventDefault();
}

export function handleMouseLeave() {
  clearHoverHighlights();
}

export function handleMouseUp(event) {
  if (!state.isSelecting) return;

  state.isSelecting = false;

  const container = document.getElementById("spreadsheet");
  if (container) {
    container.classList.remove("selecting");
  }

  // If in formula edit mode, insert the range reference
  const formulaEditMode = callbacks.getFormulaEditMode ? callbacks.getFormulaEditMode() : false;
  const formulaEditCell = callbacks.getFormulaEditCell ? callbacks.getFormulaEditCell() : null;
  const formulaRangeStart = callbacks.getFormulaRangeStart ? callbacks.getFormulaRangeStart() : null;
  const formulaRangeEnd = callbacks.getFormulaRangeEnd ? callbacks.getFormulaRangeEnd() : null;

  if (formulaEditMode && formulaEditCell && formulaRangeStart) {
    const rangeRef = callbacks.buildRangeRef
      ? callbacks.buildRangeRef(formulaRangeStart.row, formulaRangeStart.col, formulaRangeEnd.row, formulaRangeEnd.col)
      : "";

    // Focus back on formula cell and insert range
    formulaEditCell.element.focus();

    const data = callbacks.getDataArray ? callbacks.getDataArray() : [];
    const formulas = callbacks.getFormulasArray ? callbacks.getFormulasArray() : [];

    // Use setTimeout to ensure focus is established before inserting
    setTimeout(function () {
      if (callbacks.insertTextAtCursor) callbacks.insertTextAtCursor(rangeRef);

      // Update stored formula
      if (formulas[formulaEditCell.row]) {
        formulas[formulaEditCell.row][formulaEditCell.col] = formulaEditCell.element.innerText;
      }
      if (data[formulaEditCell.row]) {
        data[formulaEditCell.row][formulaEditCell.col] = formulaEditCell.element.innerText;
      }

      // Clear formula range selection but stay in formula edit mode
      if (callbacks.setFormulaRangeStart) callbacks.setFormulaRangeStart(null);
      if (callbacks.setFormulaRangeEnd) callbacks.setFormulaRangeEnd(null);
      clearSelection();

      if (callbacks.debouncedUpdateURL) callbacks.debouncedUpdateURL();
    }, 0);

    return;
  }

  // If single cell selected, allow normal focus behavior
  if (!hasMultiSelection()) {
    // Let the cell receive focus for editing
    const cellContent = document.querySelector(`.cell-content[data-row="${state.selectionStart.row}"][data-col="${state.selectionStart.col}"]`);
    if (cellContent) {
      cellContent.focus();
    }
  }
}

// ========== Touch Event Handlers ==========
export function handleTouchStart(event) {
  // Only handle single touch
  if (event.touches.length !== 1) return;

  const touch = event.touches[0];
  const cellCoords = getCellFromPoint(touch.clientX, touch.clientY);

  if (!cellCoords) {
    clearSelection();
    return;
  }

  // Start new selection
  state.selectionStart = cellCoords;
  state.selectionEnd = cellCoords;
  state.isSelecting = true;

  const container = document.getElementById("spreadsheet");
  if (container) {
    container.classList.add("selecting");
  }

  updateSelectionVisuals();
}

export function handleTouchMove(event) {
  if (!state.isSelecting) return;
  if (event.touches.length !== 1) return;

  const touch = event.touches[0];
  const cellCoords = getCellFromPoint(touch.clientX, touch.clientY);

  if (!cellCoords) return;

  // Only update if position changed
  if (state.selectionEnd && cellCoords.row === state.selectionEnd.row && cellCoords.col === state.selectionEnd.col) {
    return;
  }

  state.selectionEnd = cellCoords;
  updateSelectionVisuals();

  // Prevent scrolling during selection
  event.preventDefault();
}

export function handleTouchEnd(event) {
  if (!state.isSelecting) return;

  state.isSelecting = false;

  const container = document.getElementById("spreadsheet");
  if (container) {
    container.classList.remove("selecting");
  }

  // If single cell selected, allow focus for editing
  if (!hasMultiSelection() && state.selectionStart) {
    const cellContent = document.querySelector(`.cell-content[data-row="${state.selectionStart.row}"][data-col="${state.selectionStart.col}"]`);
    if (cellContent) {
      cellContent.focus();
    }
  }
}

// ========== Add/Remove Row/Column ==========
export function addRow() {
  if (state.rows >= MAX_ROWS) {
    showToast(`Maximum ${MAX_ROWS} rows allowed`, "warning");
    return;
  }

  const data = callbacks.getDataArray ? callbacks.getDataArray() : [];
  const formulas = callbacks.getFormulasArray ? callbacks.getFormulasArray() : [];
  const cellStyles = callbacks.getCellStylesArray ? callbacks.getCellStylesArray() : [];

  state.rows++;
  data.push(Array(state.cols).fill(""));
  formulas.push(Array(state.cols).fill(""));
  cellStyles.push(
    Array(state.cols)
      .fill(null)
      .map(() => createEmptyCellStyle())
  );
  state.rowHeights.push(DEFAULT_ROW_HEIGHT);

  renderGrid();
  if (callbacks.debouncedUpdateURL) callbacks.debouncedUpdateURL();
  showToast(`Row ${state.rows} added`, "success");
}

export function addColumn() {
  if (state.cols >= MAX_COLS) {
    showToast(`Maximum ${MAX_COLS} columns allowed`, "warning");
    return;
  }

  const data = callbacks.getDataArray ? callbacks.getDataArray() : [];
  const formulas = callbacks.getFormulasArray ? callbacks.getFormulasArray() : [];
  const cellStyles = callbacks.getCellStylesArray ? callbacks.getCellStylesArray() : [];

  state.cols++;
  data.forEach((row) => row.push(""));
  formulas.forEach((row) => row.push(""));
  cellStyles.forEach((row) => row.push(createEmptyCellStyle()));
  state.colWidths.push(DEFAULT_COL_WIDTH);

  renderGrid();
  if (callbacks.debouncedUpdateURL) callbacks.debouncedUpdateURL();
  showToast(`Column ${colToLetter(state.cols - 1)} added`, "success");
}

// ========== Clear Spreadsheet ==========
export function clearSpreadsheet() {
  if (!confirm("Clear all data and reset to 10×10 grid?")) {
    return;
  }

  // Reset to default dimensions
  state.rows = DEFAULT_ROWS;
  state.cols = DEFAULT_COLS;
  state.colWidths = createDefaultColumnWidths(state.cols);
  state.rowHeights = createDefaultRowHeights(state.rows);

  // Reset data arrays via callbacks
  if (callbacks.setDataArray) callbacks.setDataArray(createEmptyData(state.rows, state.cols));
  if (callbacks.setFormulasArray) callbacks.setFormulasArray(createEmptyData(state.rows, state.cols));
  if (callbacks.setCellStylesArray) callbacks.setCellStylesArray(createEmptyCellStyles(state.rows, state.cols));

  // Clear password
  if (callbacks.PasswordManager) callbacks.PasswordManager.setPassword(null);

  // Clear any selection
  clearSelection();

  // Re-render and update URL
  renderGrid();
  if (callbacks.debouncedUpdateURL) callbacks.debouncedUpdateURL();

  showToast("Spreadsheet cleared", "success");
}
