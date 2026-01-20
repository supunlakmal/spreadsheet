/**
 * selectionStatusManager.js
 * Manages the selection status bar that shows count, sum, and average of selected cells
 */

let getSelectionBounds;
let getDataArray;
let getEmbedModeFlag;

export const SelectionStatusManager = {
  /**
   * Initialize the selection status manager with required callbacks
   * @param {Object} callbacks - Required callbacks from script.js
   * @param {Function} callbacks.getSelectionBounds - Get current selection bounds
   * @param {Function} callbacks.getDataArray - Get data array
   * @param {Function} callbacks.getEmbedModeFlag - Check if in embed mode
   */
  init(callbacks) {
    getSelectionBounds = callbacks.getSelectionBounds;
    getDataArray = callbacks.getDataArray;
    getEmbedModeFlag = callbacks.getEmbedModeFlag;
  },

  /**
   * Extract plain text from HTML value
   * @param {*} value - Cell value (may contain HTML)
   * @returns {string} Plain text without HTML tags
   */
  extractPlainText(value) {
    if (value === null || value === undefined) return "";
    // Use DOMParser for safe HTML parsing (doesn't execute scripts)
    const parser = new DOMParser();
    const doc = parser.parseFromString("<body>" + String(value) + "</body>", "text/html");
    const text = doc.body.textContent || "";
    return text.replace(/\u00a0/g, " ");
  },

  /**
   * Format number for display in status bar
   * @param {number} value - Number to format
   * @returns {string|number} Formatted number (integers stay as-is, decimals to 2 places)
   */
  formatStatNumber(value) {
    return Number.isInteger(value) ? value : value.toFixed(2);
  },

  /**
   * Update the selection status bar with count, sum, and average
   * @param {Object|null} boundsOverride - Optional bounds object to use instead of current selection
   */
  updateStatusBar(boundsOverride = null) {
    const bar = document.getElementById("selection-status");
    const countEl = document.getElementById("stat-count");
    const sumEl = document.getElementById("stat-sum");
    const avgEl = document.getElementById("stat-avg");
    const sumWrapper = sumEl ? sumEl.parentElement : null;
    const avgWrapper = avgEl ? avgEl.parentElement : null;

    if (!bar || !countEl || !sumEl || !avgEl) return;
    if (getEmbedModeFlag()) {
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

    const data = getDataArray();
    for (let r = bounds.minRow; r <= bounds.maxRow; r++) {
      for (let c = bounds.minCol; c <= bounds.maxCol; c++) {
        const raw = data[r] && data[r][c] !== undefined ? data[r][c] : "";
        const text = this.extractPlainText(raw).trim();
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
      sumEl.innerText = this.formatStatNumber(sum);
      avgEl.innerText = (sum / countNumeric).toFixed(2);
      if (sumWrapper) sumWrapper.style.display = "inline";
      if (avgWrapper) avgWrapper.style.display = "inline";
    } else {
      if (sumWrapper) sumWrapper.style.display = "none";
      if (avgWrapper) avgWrapper.style.display = "none";
    }

    bar.classList.add("active");
  },
};
