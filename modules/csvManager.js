import { MAX_COLS, MAX_ROWS } from "./constants.js";
import { isValidFormula } from "./formulaManager.js";
import { escapeHTML } from "./security.js";
import { createDefaultColumnWidths, createDefaultRowHeights, createEmptyCellStyles, createEmptyData } from "./urlManager.js";

function csvEscape(value) {
  const text = String(value);
  const needsQuotes = /[",\r\n]/.test(text) || /^\s|\s$/.test(text);
  const escaped = text.replace(/"/g, '""');
  return needsQuotes ? `"${escaped}"` : escaped;
}

export function parseCSV(text) {
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

export const CSVManager = {
  callbacks: {
    getState: () => ({ rows: 0, cols: 0 }),
    getDataArray: () => [],
    setDataArray: () => {},
    setFormulasArray: () => {},
    setCellStylesArray: () => {},
    setState: () => {},
    renderGrid: () => {},
    recalculateFormulas: () => {},
    debouncedUpdateURL: () => {},
    showToast: () => {},
    extractPlainText: (value) => (value === null || value === undefined ? "" : String(value)),
    onImport: null,
  },

  init(callbacks = {}) {
    CSVManager.callbacks = { ...CSVManager.callbacks, ...callbacks };
  },

  buildCSV() {
    const { getState, getDataArray, extractPlainText } = CSVManager.callbacks;
    if (!getState || !getDataArray || !extractPlainText) return "";

    const { rows, cols } = getState();
    const data = getDataArray();
    const lines = [];

    for (let r = 0; r < rows; r++) {
      const rowValues = [];
      for (let c = 0; c < cols; c++) {
        const raw = data[r] && data[r][c] !== undefined ? data[r][c] : "";
        const text = extractPlainText(raw);
        rowValues.push(csvEscape(text));
      }
      lines.push(rowValues.join(","));
    }

    return lines.join("\r\n");
  },

  downloadCSV() {
    const { recalculateFormulas, showToast } = CSVManager.callbacks;

    if (recalculateFormulas) {
      recalculateFormulas();
    }
    const csv = CSVManager.buildCSV();

    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "spreadsheet.csv";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);

    if (showToast) {
      showToast("CSV downloaded", "success");
    }
  },

  importCSVText(text) {
    const parsedRows = parseCSV(text);
    if (!parsedRows.length) {
      if (CSVManager.callbacks.showToast) {
        CSVManager.callbacks.showToast("CSV file is empty", "error");
      }
      return;
    }

    const maxColsInFile = parsedRows.reduce((max, row) => Math.max(max, row.length), 0);
    const nextRows = Math.min(Math.max(parsedRows.length, 1), MAX_ROWS);
    const nextCols = Math.min(Math.max(maxColsInFile, 1), MAX_COLS);
    const truncated = parsedRows.length > MAX_ROWS || maxColsInFile > MAX_COLS;

    CSVManager.callbacks.setState("rows", nextRows);
    CSVManager.callbacks.setState("cols", nextCols);
    CSVManager.callbacks.setState("colWidths", createDefaultColumnWidths(nextCols));
    CSVManager.callbacks.setState("rowHeights", createDefaultRowHeights(nextRows));

    const data = createEmptyData(nextRows, nextCols);
    const formulas = createEmptyData(nextRows, nextCols);
    const cellStyles = createEmptyCellStyles(nextRows, nextCols);

    for (let r = 0; r < nextRows; r++) {
      const sourceRow = Array.isArray(parsedRows[r]) ? parsedRows[r] : [];
      for (let c = 0; c < nextCols; c++) {
        const raw = sourceRow[c] !== undefined ? String(sourceRow[c]) : "";
        if (raw.startsWith("=")) {
          if (isValidFormula(raw)) {
            formulas[r][c] = raw;
            data[r][c] = raw;
          } else {
            formulas[r][c] = "";
            data[r][c] = escapeHTML(raw);
          }
        } else {
          data[r][c] = escapeHTML(raw);
        }
      }
    }

    CSVManager.callbacks.setDataArray(data);
    CSVManager.callbacks.setFormulasArray(formulas);
    CSVManager.callbacks.setCellStylesArray(cellStyles);

    if (CSVManager.callbacks.renderGrid) {
      CSVManager.callbacks.renderGrid();
    }
    if (CSVManager.callbacks.recalculateFormulas) {
      CSVManager.callbacks.recalculateFormulas();
    }
    if (CSVManager.callbacks.debouncedUpdateURL) {
      CSVManager.callbacks.debouncedUpdateURL();
    }
    if (CSVManager.callbacks.onImport) {
      CSVManager.callbacks.onImport({
        truncated,
        rows: nextRows,
        cols: nextCols,
      });
    }

    if (CSVManager.callbacks.showToast) {
      if (truncated) {
        CSVManager.callbacks.showToast("CSV imported (some data truncated due to size limits)", "warning");
      } else {
        CSVManager.callbacks.showToast("CSV imported successfully", "success");
      }
    }
  },
};
