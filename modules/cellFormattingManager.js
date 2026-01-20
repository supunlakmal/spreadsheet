/**
 * cellFormattingManager.js
 * Manages cell formatting operations (alignment, colors, font size, text formatting)
 */

let getSelectionBounds;
let getCellContentElement;
let getCellElement;
let getCellStylesArray;
let getDataArray;
let getFormulasArray;
let getState;
let debouncedUpdateURL;
let scheduleFullSync;
let normalizeAlignment;
let normalizeFontSize;
let isValidCSSColor;
let sanitizeHTML;
let createEmptyCellStyle;
let canBroadcastP2P;
let P2PManager;

export const CellFormattingManager = {
  /**
   * Initialize the cell formatting manager with required callbacks
   * @param {Object} callbacks - Required callbacks from script.js
   */
  init(callbacks) {
    getSelectionBounds = callbacks.getSelectionBounds;
    getCellContentElement = callbacks.getCellContentElement;
    getCellElement = callbacks.getCellElement;
    getCellStylesArray = callbacks.getCellStylesArray;
    getDataArray = callbacks.getDataArray;
    getFormulasArray = callbacks.getFormulasArray;
    getState = callbacks.getState;
    debouncedUpdateURL = callbacks.debouncedUpdateURL;
    scheduleFullSync = callbacks.scheduleFullSync;
    normalizeAlignment = callbacks.normalizeAlignment;
    normalizeFontSize = callbacks.normalizeFontSize;
    isValidCSSColor = callbacks.isValidCSSColor;
    sanitizeHTML = callbacks.sanitizeHTML;
    createEmptyCellStyle = callbacks.createEmptyCellStyle;
    canBroadcastP2P = callbacks.canBroadcastP2P;
    P2PManager = callbacks.P2PManager;
  },

  /**
   * Apply a callback to all target cells (selection or active cell)
   * @param {Function} callback - Function to call with (row, col) for each cell
   * @returns {boolean} True if any cells were processed
   * @private
   */
  forEachTargetCell(callback) {
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
  },

  /**
   * Apply alignment to selected cells
   * @param {string} align - Alignment value (left, center, right)
   */
  applyAlignment(align) {
    const normalized = normalizeAlignment(align);
    if (!normalized) return;

    const cellStyles = getCellStylesArray();
    const updated = this.forEachTargetCell(function (row, col) {
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
      scheduleFullSync();
    }
  },

  /**
   * Apply background color to selected cells
   * @param {string} color - CSS color value
   */
  applyCellBackground(color) {
    if (!isValidCSSColor(color)) return;

    const cellStyles = getCellStylesArray();
    const updated = this.forEachTargetCell(function (row, col) {
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
      scheduleFullSync();
    }
  },

  /**
   * Apply text color to selected cells
   * @param {string} color - CSS color value
   */
  applyCellTextColor(color) {
    if (!isValidCSSColor(color)) return;

    const cellStyles = getCellStylesArray();
    const updated = this.forEachTargetCell(function (row, col) {
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
      scheduleFullSync();
    }
  },

  /**
   * Apply font size to selected cells
   * @param {number} size - Font size in pixels
   */
  applyFontSize(size) {
    const normalized = normalizeFontSize(size);

    const cellStyles = getCellStylesArray();
    const updated = this.forEachTargetCell(function (row, col) {
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
      scheduleFullSync();
    }
  },

  /**
   * Apply text formatting (bold, italic, underline) to selected text
   * Uses modern Selection/Range API (replaces deprecated execCommand)
   * @param {string} command - Format command (bold, italic, underline)
   */
  applyFormat(command) {
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
    const data = getDataArray();
    const formulas = getFormulasArray();
    if (!isNaN(row) && !isNaN(col) && row < rows && col < cols) {
      data[row][col] = sanitizeHTML(activeElement.innerHTML);
      debouncedUpdateURL();
      if (canBroadcastP2P()) {
        P2PManager.broadcastCellUpdate(row, col, data[row][col], formulas[row][col]);
      }
    }
  },
};
