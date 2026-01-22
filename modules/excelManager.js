/* global ExcelJS */

import { createEmptyData, createEmptyCellStyles } from "./urlManager.js";

export const ExcelManager = {
  callbacks: {
    getState: () => ({ rows: 0, cols: 0 }),
    getDataArray: () => [],
    getFormulasArray: () => [],
    getCellStylesArray: () => [],
    recalculateFormulas: () => {},
    showToast: () => {},
    extractPlainText: (value) => (value === null || value === undefined ? "" : String(value)),
  },

  init(callbacks = {}) {
    ExcelManager.callbacks = { ...ExcelManager.callbacks, ...callbacks };
  },

  /**
   * Expand a minified state object (like from templates.json) into full arrays.
   */
  expandMinifiedState(minState) {
    if (!minState) return null;

    const rows = minState.r || 0;
    const cols = minState.c || 0;
    
    // Create empty structures
    const data = createEmptyData(rows, cols);
    const formulas = createEmptyData(rows, cols);
    const cellStyles = createEmptyCellStyles(rows, cols);

    // Expand data - minState.d is [[r, c, val], ...]
    if (minState.d) {
      minState.d.forEach(([r, c, val]) => {
        if (r < rows && c < cols) data[r][c] = val;
      });
    }

    // Expand formulas - minState.f is [[r, c, formula], ...]
    if (minState.f) {
      minState.f.forEach(([r, c, f]) => {
        if (r < rows && c < cols) formulas[r][c] = f;
      });
    }

    // Expand styles - minState.s is [[r, c, styleObj], ...]
    // styleObj uses keys: a (align), b (bgColor), c (color), z (fontSize), bo (bold), i (italic), u (underline)
    if (minState.s) {
      minState.s.forEach(([r, c, s]) => {
        if (r < rows && c < cols) {
          cellStyles[r][c] = {
            align: s.a || "left",
            bgColor: s.b || "#ffffff",
            color: s.c || "#000000",
            fontSize: s.z || "12",
            bold: !!s.bo,
            italic: !!s.i,
            underline: !!s.u
          };
        }
      });
    }

    return { rows, cols, data, formulas, cellStyles };
  },

  /**
   * Convert hex color to ARGB format for ExcelJS
   * @param {string} hex - Hex color like "#ff0000"
   * @returns {string} ARGB format like "FFFF0000"
   */
  hexToARGB(hex) {
    if (!hex || hex === "transparent") return null;
    // Remove # if present
    const cleanHex = hex.replace("#", "").toUpperCase();
    // If it's a 3-char hex, expand it
    if (cleanHex.length === 3) {
      const r = cleanHex[0];
      const g = cleanHex[1];
      const b = cleanHex[2];
      return "FF" + r + r + g + g + b + b;
    }
    // Already 6 chars, add FF for full opacity
    if (cleanHex.length === 6) {
      return "FF" + cleanHex;
    }
    return null;
  },

  /**
   * Export the current spreadsheet as an Excel (.xlsx) file.
   */
  async downloadExcel() {
    const { 
      getState, 
      getDataArray, 
      getFormulasArray,
      getCellStylesArray,
      recalculateFormulas,
      showToast
    } = ExcelManager.callbacks;

    if (!getState || !getDataArray) return;

    try {
      if (recalculateFormulas) {
        recalculateFormulas();
      }

      const { rows, cols } = getState();
      const state = {
        rows,
        cols,
        data: getDataArray(),
        formulas: getFormulasArray ? getFormulasArray() : [],
        cellStyles: getCellStylesArray ? getCellStylesArray() : []
      };

      await ExcelManager.exportFromState(state, "spreadsheet.xlsx");
    } catch (error) {
      console.error("Excel export failed:", error);
      if (showToast) {
        showToast("Failed to export Excel file", "error");
      }
    }
  },

  /**
   * Core export logic that takes a full state object.
   */
  async exportFromState(state, filename = "spreadsheet.xlsx") {
    const { rows, cols, data, formulas, cellStyles } = state;
    const { extractPlainText, showToast } = ExcelManager.callbacks;

    // Check if ExcelJS is available
    if (typeof window.ExcelJS === "undefined") {
      console.error("ExcelJS library not loaded");
      if (showToast) showToast("Excel library not loaded", "error");
      return;
    }

    try {
      // Create workbook and worksheet
      const workbook = new window.ExcelJS.Workbook();
      workbook.creator = "Spreadsheet App";
      workbook.created = new Date();
      
      const worksheet = workbook.addWorksheet("Spreadsheet");

      // Process each cell
      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
          const formula = formulas[r] && formulas[r][c] ? formulas[r][c] : "";
          const rawValue = data[r] && data[r][c] !== undefined ? data[r][c] : "";
          const style = cellStyles[r] && cellStyles[r][c] ? cellStyles[r][c] : null;
          
          if (!rawValue && !formula && !style) continue;
          
          const cell = worksheet.getCell(r + 1, c + 1);
          
          if (formula && formula.startsWith("=")) {
            cell.value = { formula: formula.substring(1) };
          } else {
            const plainValue = extractPlainText(rawValue);
            const numValue = parseFloat(plainValue.replace(/,/g, ""));
            if (!isNaN(numValue) && plainValue.trim() !== "" && !isNaN(plainValue)) {
                cell.value = numValue;
            } else {
                cell.value = plainValue;
            }
          }
          
          if (style) {
            const font = {};
            if (style.bold) font.bold = true;
            if (style.italic) font.italic = true;
            if (style.underline) font.underline = true;
            if (style.fontSize) font.size = parseInt(style.fontSize, 10);
            if (style.color && style.color !== "#000000") {
              const argb = ExcelManager.hexToARGB(style.color);
              if (argb) font.color = { argb };
            }
            if (Object.keys(font).length > 0) cell.font = font;
            
            if (style.bgColor && style.bgColor !== "#ffffff" && style.bgColor !== "transparent") {
              const argb = ExcelManager.hexToARGB(style.bgColor);
              if (argb) {
                cell.fill = {
                  type: "pattern",
                  pattern: "solid",
                  fgColor: { argb }
                };
              }
            }
            
            if (style.align && style.align !== "left") {
              cell.alignment = { horizontal: style.align };
            }
          }
        }
      }

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { 
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" 
      });
      
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);

      if (showToast) {
        showToast(`Excel file "${filename}" downloaded!`, "success");
      }
    } catch (error) {
      console.error("Export From State failed:", error);
      throw error;
    }
  }
};
