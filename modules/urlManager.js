import {
  DEFAULT_COL_WIDTH,
  DEFAULT_COLS,
  DEFAULT_ROW_HEIGHT,
  DEFAULT_ROWS,
  FONT_SIZE_OPTIONS,
  KEY_MAP,
  KEY_MAP_REVERSE,
  MAX_COLS,
  MAX_ROWS,
  MIN_COL_WIDTH,
  MIN_ROW_HEIGHT,
  STYLE_KEY_MAP,
  URL_LENGTH_CAUTION,
  URL_LENGTH_CRITICAL,
  URL_LENGTH_MAX_DISPLAY,
  URL_LENGTH_WARNING,
} from "./constants.js";
import { CryptoUtils, EncryptionCodec } from "./encryption.js";
import { isValidFormula } from "./formulaManager.js";
import { escapeHTML, isValidCSSColor, sanitizeHTML } from "./security.js";

// Factory functions
export function createEmptyData(r, c) {
  return Array(r)
    .fill(null)
    .map(() => Array(c).fill(""));
}

export function createEmptyCellStyle() {
  return { align: "", bg: "", color: "", fontSize: "" };
}

export function createEmptyCellStyles(r, c) {
  return Array(r)
    .fill(null)
    .map(() =>
      Array(c)
        .fill(null)
        .map(() => createEmptyCellStyle())
    );
}

export function createDefaultColumnWidths(count) {
  return Array(count).fill(DEFAULT_COL_WIDTH);
}

export function createDefaultRowHeights(count) {
  return Array(count).fill(DEFAULT_ROW_HEIGHT);
}

// Normalizers
export function normalizeColumnWidths(widths, count) {
  const normalized = [];
  for (let i = 0; i < count; i++) {
    const raw = widths && widths[i];
    const value = parseInt(raw, 10);
    if (Number.isFinite(value) && value > 0) {
      normalized.push(Math.max(MIN_COL_WIDTH, value));
    } else {
      normalized.push(DEFAULT_COL_WIDTH);
    }
  }
  return normalized;
}

export function normalizeRowHeights(heights, count) {
  const normalized = [];
  for (let i = 0; i < count; i++) {
    const raw = heights && heights[i];
    const value = parseInt(raw, 10);
    if (Number.isFinite(value) && value > 0) {
      normalized.push(Math.max(MIN_ROW_HEIGHT, value));
    } else {
      normalized.push(DEFAULT_ROW_HEIGHT);
    }
  }
  return normalized;
}

export function normalizeAlignment(value) {
  if (value === "left" || value === "center" || value === "right") {
    return value;
  }
  return "";
}

export function normalizeFontSize(value) {
  if (value === null || value === undefined) return "";
  let raw = String(value).trim();
  if (raw === "") return "";
  if (raw.endsWith("px")) {
    raw = raw.slice(0, -2).trim();
  }
  const size = parseInt(raw, 10);
  if (isNaN(size)) return "";
  if (!FONT_SIZE_OPTIONS.includes(size)) return "";
  return String(size);
}

export function normalizeCellStyles(styles, r, c) {
  const normalized = createEmptyCellStyles(r, c);
  if (!Array.isArray(styles)) return normalized;

  for (let row = 0; row < r; row++) {
    const sourceRow = Array.isArray(styles[row]) ? styles[row] : [];
    for (let col = 0; col < c; col++) {
      const cellStyle = sourceRow[col];
      if (cellStyle && typeof cellStyle === "object") {
        normalized[row][col] = {
          align: normalizeAlignment(cellStyle.align),
          bg: isValidCSSColor(cellStyle.bg) ? cellStyle.bg : "",
          color: isValidCSSColor(cellStyle.color) ? cellStyle.color : "",
          fontSize: normalizeFontSize(cellStyle.fontSize),
        };
      }
    }
  }
  return normalized;
}

// Check if all cells in data array are empty
export function isDataEmpty(d) {
  return d.every((row) => row.every((cell) => cell === ""));
}

// Check if all cells in formulas array are empty
export function isFormulasEmpty(f) {
  return f.every((row) => row.every((cell) => cell === ""));
}

// Check if a cell style is default (all empty)
export function isCellStyleDefault(style) {
  return !style || (style.align === "" && style.bg === "" && style.color === "" && style.fontSize === "");
}

// Check if all cell styles are default
export function isCellStylesDefault(styles) {
  return styles.every((row) => row.every((cell) => isCellStyleDefault(cell)));
}

// Check if column widths are all default
export function isColWidthsDefault(widths, count) {
  if (widths.length !== count) return false;
  return widths.every((w) => w === DEFAULT_COL_WIDTH);
}

// Check if row heights are all default
export function isRowHeightsDefault(heights, count) {
  if (heights.length !== count) return false;
  return heights.every((h) => h === DEFAULT_ROW_HEIGHT);
}

// Safe JSON parse with prototype pollution protection
function safeJSONParse(jsonString) {
  const parsed = JSON.parse(jsonString);

  function createSafeCopy(obj) {
    if (obj === null || typeof obj !== "object") return obj;
    if (Array.isArray(obj)) {
      return obj.map(createSafeCopy);
    }

    const safe = Object.create(null);
    for (const key of Object.keys(obj)) {
      // Block dangerous keys that could pollute prototypes
      if (key === "__proto__" || key === "constructor" || key === "prototype") {
        console.warn("Blocked prototype pollution attempt via key:", key);
        continue;
      }
      // Block keys containing prototype chain accessor patterns
      if (key.includes("__proto__") || key.includes("constructor.prototype")) {
        console.warn("Blocked prototype pollution attempt via key pattern:", key);
        continue;
      }
      safe[key] = createSafeCopy(obj[key]);
    }
    return safe;
  }

  return createSafeCopy(parsed);
}

// ========== Sparse Array Conversion Helpers ==========

// Detect whether an array is in sparse format [[row,col,val],...] or dense format [[val1,val2],...]
function detectArrayFormat(array) {
  if (!array || array.length === 0) return "empty";

  const first = array[0];
  if (!Array.isArray(first)) return "invalid";

  // Sparse format signature: first element is [number, number, anything]
  // Dense format signature: first element is array of values (strings/objects)
  if (first.length === 3 &&
      typeof first[0] === "number" &&
      typeof first[1] === "number" &&
      Number.isInteger(first[0]) &&
      Number.isInteger(first[1])) {
    return "sparse";
  }

  return "dense";
}

// Convert dense 2D array to sparse triplet array [[row, col, value], ...]
function convertToSparse(array2D) {
  if (!array2D || !Array.isArray(array2D)) return undefined;

  const sparse = [];
  for (let r = 0; r < array2D.length; r++) {
    const row = array2D[r];
    if (!Array.isArray(row)) continue;

    for (let c = 0; c < row.length; c++) {
      const val = row[c];
      // Only include non-empty values
      if (val !== "") {
        sparse.push([r, c, val]);
      }
    }
  }

  return sparse.length > 0 ? sparse : undefined;
}

// Convert sparse triplet array to dense 2D array
function expandFromSparse(sparse, rows, cols, emptyValue = "") {
  // Initialize dense array with empty values
  const dense = Array(rows).fill(null).map(() => Array(cols).fill(emptyValue));

  if (!sparse || sparse.length === 0) return dense;

  for (const entry of sparse) {
    // Validate triplet format
    if (!Array.isArray(entry) || entry.length !== 3) {
      console.warn("Invalid sparse triplet format:", entry);
      continue;
    }

    const [r, c, val] = entry;

    // Validate indices (type check and bounds check)
    if (!Number.isInteger(r) || !Number.isInteger(c)) {
      console.warn("Non-integer indices in sparse entry:", r, c);
      continue;
    }

    if (r < 0 || r >= rows || c < 0 || c >= cols) {
      console.warn("Out of bounds sparse entry (ignoring):", r, c, "grid size:", rows, "x", cols);
      continue;
    }

    // Set value in dense array
    dense[r][c] = val;
  }

  return dense;
}

// Convert cell styles dense 2D array to sparse format
function convertCellStylesToSparse(styles2D) {
  if (!styles2D || !Array.isArray(styles2D)) return undefined;

  const sparse = [];

  for (let r = 0; r < styles2D.length; r++) {
    const row = styles2D[r];
    if (!Array.isArray(row)) continue;

    for (let c = 0; c < row.length; c++) {
      const style = row[c];

      // Skip default/empty styles
      if (isCellStyleDefault(style)) continue;

      // Minify non-default style (remove empty properties)
      const minified = {};
      for (const [key, val] of Object.entries(style)) {
        if (val !== "") {
          minified[STYLE_KEY_MAP[key] || key] = val;
        }
      }

      // Only add if there are non-empty properties
      if (Object.keys(minified).length > 0) {
        sparse.push([r, c, minified]);
      }
    }
  }

  return sparse.length > 0 ? sparse : undefined;
}

// Expand cell styles from sparse format to dense 2D array
function expandCellStylesFromSparse(sparse, rows, cols) {
  // Initialize with default empty styles
  const dense = createEmptyCellStyles(rows, cols);

  if (!sparse || sparse.length === 0) return dense;

  for (const entry of sparse) {
    // Validate triplet format
    if (!Array.isArray(entry) || entry.length !== 3) {
      console.warn("Invalid sparse style triplet:", entry);
      continue;
    }

    const [r, c, minStyle] = entry;

    // Validate indices
    if (!Number.isInteger(r) || !Number.isInteger(c)) {
      console.warn("Non-integer indices in sparse style entry:", r, c);
      continue;
    }

    if (r < 0 || r >= rows || c < 0 || c >= cols) {
      console.warn("Out of bounds sparse style entry:", r, c);
      continue;
    }

    // Expand minified style keys back to full names
    if (minStyle && typeof minStyle === "object") {
      dense[r][c] = {
        align: minStyle.a || minStyle.align || "",
        bg: minStyle.b || minStyle.bg || "",
        color: minStyle.c || minStyle.color || "",
        fontSize: minStyle.z || minStyle.fontSize || "",
      };
    }
  }

  return dense;
}

// ========== End Sparse Array Helpers ==========

// Minify state object keys for smaller URL payload
function minifyState(state) {
  const minified = {};
  for (const [key, value] of Object.entries(state)) {
    const shortKey = KEY_MAP[key] || key;

    // Convert data and formulas to sparse format
    if ((key === "data" || key === "formulas") && Array.isArray(value)) {
      const sparse = convertToSparse(value);
      if (sparse !== undefined) {
        minified[shortKey] = sparse;
      }
      // If sparse is undefined (no data), don't include the key at all
    }
    // Convert cellStyles to sparse format
    else if (key === "cellStyles" && Array.isArray(value)) {
      const sparse = convertCellStylesToSparse(value);
      if (sparse !== undefined) {
        minified[shortKey] = sparse;
      }
      // If sparse is undefined (no styles), don't include the key at all
    }
    // Pass through other keys unchanged
    else {
      minified[shortKey] = value;
    }
  }
  return minified;
}

// Expand minified state keys back to full names (backward compatible)
function expandState(minified) {
  // Handle null, non-objects, and arrays (legacy format) - return as-is
  if (!minified || typeof minified !== "object" || Array.isArray(minified)) {
    return minified;
  }
  const expanded = {};
  for (const [key, value] of Object.entries(minified)) {
    const fullKey = KEY_MAP_REVERSE[key] || key;

    // Handle data: detect sparse vs dense format
    if ((fullKey === "data" || key === "d") && Array.isArray(value)) {
      const format = detectArrayFormat(value);
      if (format === "sparse") {
        // Mark as sparse for validateAndNormalizeState to expand later
        expanded.data = { _format: "sparse", _data: value };
      } else {
        // Dense format (legacy) - pass through as-is
        expanded.data = value;
      }
    }
    // Handle formulas: detect sparse vs dense format
    else if ((fullKey === "formulas" || key === "f") && Array.isArray(value)) {
      const format = detectArrayFormat(value);
      if (format === "sparse") {
        expanded.formulas = { _format: "sparse", _data: value };
      } else {
        // Dense format (legacy) - pass through as-is
        expanded.formulas = value;
      }
    }
    // Handle cellStyles: detect sparse vs dense format
    else if ((fullKey === "cellStyles" || key === "s") && Array.isArray(value)) {
      const format = detectArrayFormat(value);
      if (format === "sparse") {
        // Mark as sparse for validateAndNormalizeState to expand later
        expanded.cellStyles = { _format: "sparse", _data: value };
      } else {
        // Dense format (legacy) - expand minified style keys
        expanded.cellStyles = value.map((row) =>
          row.map((cell) => {
            if (!cell || typeof cell !== "object") {
              return { align: "", bg: "", color: "", fontSize: "" };
            }
            return {
              align: cell.a || cell.align || "",
              bg: cell.b || cell.bg || "",
              color: cell.c || cell.color || "",
              fontSize: cell.z || cell.fontSize || "",
            };
          })
        );
      }
    }
    // Pass through other keys unchanged
    else {
      expanded[fullKey] = value;
    }
  }
  return expanded;
}

// ========== Serialization Codec (Wrapper for minify/expand + compression) ==========
const SerializationCodec = {
  // Serialize a state object to a compressed string
  serialize(state) {
    const json = JSON.stringify(minifyState(state));
    return LZString.compressToEncodedURIComponent(json);
  },

  // Deserialize a compressed string to a state object
  // Returns null on failure
  deserialize(compressed) {
    try {
      const json = LZString.decompressFromEncodedURIComponent(compressed);
      if (!json || json.length === 0) return null;
      return expandState(safeJSONParse(json));
    } catch (e) {
      console.error("Deserialization failed:", e);
      return null;
    }
  },
};

// ========== URL Codec (Main Wrapper - combines serialization and encryption) ==========
const URLCodec = {
  // Encode state to URL-ready string
  async encode(state, options = {}) {
    const serialized = SerializationCodec.serialize(state);

    // Optionally encrypt
    if (options.password) {
      try {
        return await EncryptionCodec.encrypt(serialized, options.password);
      } catch (e) {
        console.error("Encryption failed, falling back to unencrypted:", e);
        return serialized;
      }
    }

    return serialized;
  },

  // Decode URL hash - returns { state } or { encrypted: true, data }
  decode(hash) {
    // Security check
    if (!hash || hash.length > 100000) {
      if (hash && hash.length > 100000) {
        console.warn("Hash too long, rejecting");
      }
      return null;
    }

    // Check for encrypted data
    if (EncryptionCodec.isEncrypted(hash)) {
      return {
        encrypted: true,
        data: EncryptionCodec.unwrap(hash),
      };
    }

    // Try to deserialize
    const state = SerializationCodec.deserialize(hash);
    if (state) {
      return { state };
    }

    // Try legacy format
    return this._decodeLegacy(hash);
  },

  // Decrypt and decode an encrypted payload
  async decryptAndDecode(encryptedData, password) {
    const decrypted = await CryptoUtils.decrypt(encryptedData, password);
    const state = SerializationCodec.deserialize(decrypted);
    if (!state) {
      throw new Error("Invalid decrypted data");
    }
    return state;
  },

  // Handle legacy uncompressed format
  _decodeLegacy(hash) {
    try {
      // Only attempt legacy decode if it looks like valid URL-encoded JSON
      if (hash.startsWith("%7B") || hash.startsWith("%5B") || hash.startsWith("{") || hash.startsWith("[")) {
        const decoded = decodeURIComponent(hash);
        return { state: expandState(safeJSONParse(decoded)) };
      }
    } catch (e) {
      console.warn("Legacy decode failed:", e);
    }
    return null;
  },
};

// Validate and normalize a parsed state object
export function validateAndNormalizeState(parsed) {
  try {
    // Handle new format (object with rows, cols, data)
    if (parsed && typeof parsed === "object" && !Array.isArray(parsed)) {
      const r = Math.min(Math.max(1, parsed.rows || DEFAULT_ROWS), MAX_ROWS);
      const c = Math.min(Math.max(1, parsed.cols || DEFAULT_COLS), MAX_COLS);

      // Validate and normalize data array
      let d = parsed.data;
      // Check if data is marked as sparse format by expandState
      if (d && typeof d === "object" && d._format === "sparse") {
        // Sparse format: expand to dense
        d = expandFromSparse(d._data, r, c, "");
        // Sanitize all cells
        d = d.map((row) => row.map((cell) => sanitizeHTML(String(cell))));
      }
      // Handle direct sparse array (backward compatibility - no metadata wrapper)
      else if (Array.isArray(d)) {
        const format = detectArrayFormat(d);
        if (format === "sparse") {
          // Direct sparse array without metadata wrapper
          d = expandFromSparse(d, r, c, "");
          d = d.map((row) => row.map((cell) => sanitizeHTML(String(cell))));
        } else {
          // Dense format (legacy): existing logic
          // Ensure correct dimensions
          d = d.slice(0, r).map((row) => {
            if (Array.isArray(row)) {
              return row.slice(0, c).map((cell) => sanitizeHTML(String(cell || "")));
            }
            return Array(c).fill("");
          });
          // Pad rows if needed
          while (d.length < r) {
            d.push(Array(c).fill(""));
          }
          // Pad columns if needed
          d = d.map((row) => {
            while (row.length < c) row.push("");
            return row;
          });
        }
      } else {
        d = createEmptyData(r, c);
      }

      // Load formulas (backward compatible - create empty if not present)
      // Security: Validate formulas against whitelist to prevent injection
      let f = parsed.formulas;
      // Check if formulas is marked as sparse format by expandState
      if (f && typeof f === "object" && f._format === "sparse") {
        // Sparse format: expand to dense
        f = expandFromSparse(f._data, r, c, "");
        // Validate all formulas
        f = f.map((row, rowIdx) => row.map((cell, colIdx) => {
          const formula = String(cell);
          if (formula.startsWith("=")) {
            if (isValidFormula(formula)) {
              return formula;
            } else {
              // Invalid formula - convert to escaped text in data
              d[rowIdx][colIdx] = escapeHTML(formula);
              return "";
            }
          }
          return formula;
        }));
      }
      // Handle direct sparse array (backward compatibility)
      else if (Array.isArray(f)) {
        const format = detectArrayFormat(f);
        if (format === "sparse") {
          // Direct sparse array without metadata wrapper
          f = expandFromSparse(f, r, c, "");
          // Validate all formulas
          f = f.map((row, rowIdx) => row.map((cell, colIdx) => {
            const formula = String(cell);
            if (formula.startsWith("=")) {
              if (isValidFormula(formula)) {
                return formula;
              } else {
                // Invalid formula - convert to escaped text in data
                d[rowIdx][colIdx] = escapeHTML(formula);
                return "";
              }
            }
            return formula;
          }));
        } else {
          // Dense format (legacy): existing logic
          f = f.slice(0, r).map((row, rowIdx) => {
            if (Array.isArray(row)) {
              return row.slice(0, c).map((cell, colIdx) => {
                const formula = String(cell || "");
                if (formula.startsWith("=")) {
                  // Validate formula against whitelist
                  if (isValidFormula(formula)) {
                    return formula;
                  } else {
                    // Invalid formula - convert to escaped text in data
                    d[rowIdx][colIdx] = escapeHTML(formula);
                    return "";
                  }
                }
                return formula;
              });
            }
            return Array(c).fill("");
          });
          while (f.length < r) {
            f.push(Array(c).fill(""));
          }
          f = f.map((row) => {
            while (row.length < c) row.push("");
            return row;
          });
        }
      } else {
        f = createEmptyData(r, c);
      }

      // Handle cellStyles: detect sparse vs dense format
      let s;
      const parsedStyles = parsed.cellStyles;
      if (parsedStyles && typeof parsedStyles === "object" && parsedStyles._format === "sparse") {
        // Sparse format: expand to dense first, then normalize/validate
        const expanded = expandCellStylesFromSparse(parsedStyles._data, r, c);
        s = normalizeCellStyles(expanded, r, c);
      } else if (Array.isArray(parsedStyles)) {
        const format = detectArrayFormat(parsedStyles);
        if (format === "sparse") {
          // Direct sparse array without metadata wrapper
          const expanded = expandCellStylesFromSparse(parsedStyles, r, c);
          s = normalizeCellStyles(expanded, r, c);
        } else {
          // Dense format (legacy): use normalizeCellStyles directly
          s = normalizeCellStyles(parsedStyles, r, c);
        }
      } else {
        // No styles or invalid format
        s = normalizeCellStyles(parsedStyles, r, c);
      }
      const w = normalizeColumnWidths(parsed.colWidths, c);
      const h = normalizeRowHeights(parsed.rowHeights, r);
      return {
        rows: r,
        cols: c,
        data: d,
        formulas: f,
        cellStyles: s,
        colWidths: w,
        rowHeights: h,
        theme: parsed.theme || null,
        readOnly: Boolean(parsed.readOnly),
        embed: Boolean(parsed.embed),
      };
    }

    // Handle legacy format (just array, assume 10x10)
    if (Array.isArray(parsed)) {
      const d = parsed.slice(0, DEFAULT_ROWS).map((row) => {
        if (Array.isArray(row)) {
          return row.slice(0, DEFAULT_COLS).map((cell) => sanitizeHTML(String(cell || "")));
        }
        return Array(DEFAULT_COLS).fill("");
      });
      while (d.length < DEFAULT_ROWS) {
        d.push(Array(DEFAULT_COLS).fill(""));
      }
      const f = createEmptyData(DEFAULT_ROWS, DEFAULT_COLS);
      const s = createEmptyCellStyles(DEFAULT_ROWS, DEFAULT_COLS);
      const w = createDefaultColumnWidths(DEFAULT_COLS);
      const h = createDefaultRowHeights(DEFAULT_ROWS);
      return {
        rows: DEFAULT_ROWS,
        cols: DEFAULT_COLS,
        data: d,
        formulas: f,
        cellStyles: s,
        colWidths: w,
        rowHeights: h,
        theme: null,
        readOnly: false,
      };
    }
  } catch (e) {
    console.warn("Failed to validate state:", e);
  }
  return null;
}

// Update URL length indicator in status bar
function updateURLLengthIndicator(length) {
  const valueEl = document.getElementById("url-length-value");
  const barEl = document.getElementById("url-progress-bar");
  const msgEl = document.getElementById("url-length-message");
  if (!valueEl || !barEl) return;

  valueEl.textContent = length.toLocaleString();

  // Calculate progress percentage (capped at 100%)
  const percent = Math.min((length / URL_LENGTH_MAX_DISPLAY) * 100, 100);
  barEl.style.width = percent + "%";

  // Update color and message based on thresholds
  barEl.classList.remove("warning", "caution", "critical");
  if (msgEl) {
    msgEl.classList.remove("warning", "caution", "critical");
    msgEl.textContent = "";
  }

  if (length >= URL_LENGTH_CRITICAL) {
    barEl.classList.add("critical");
    if (msgEl) {
      msgEl.classList.add("critical");
      msgEl.textContent = "Some browsers may fail";
    }
  } else if (length >= URL_LENGTH_CAUTION) {
    barEl.classList.add("caution");
    if (msgEl) {
      msgEl.classList.add("caution");
      msgEl.textContent = "URL shorteners may fail";
    }
  } else if (length >= URL_LENGTH_WARNING) {
    barEl.classList.add("warning");
    if (msgEl) {
      msgEl.classList.add("warning");
      msgEl.textContent = "Some older browsers may truncate";
    }
  }
}

export const URLManager = {
  // Decode URL hash to state object
  // Returns { encrypted: true, data: base64String } if data is encrypted
  // Returns { rows, ... } if success
  // Returns null if failure
  decodeState(hash) {
    try {
      const result = URLCodec.decode(hash);
      if (!result) return null;

      // If encrypted, return the encrypted marker
      if (result.encrypted) {
        return result;
      }

      // Validate and normalize the decoded state
      return validateAndNormalizeState(result.state);
    } catch (e) {
      console.warn("Failed to decode state from URL:", e);
    }
    return null;
  },

  async encodeState(state, password) {
    return await URLCodec.encode(state, { password });
  },

  async decryptAndDecode(encryptedData, password) {
    return await URLCodec.decryptAndDecode(encryptedData, password);
  },

  // Update URL hash without page jump
  async updateURL(state, password) {
    const encoded = await this.encodeState(state, password);
    const newHash = "#" + encoded;

    // Update URL length indicator
    updateURLLengthIndicator(encoded.length);

    if (history.replaceState) {
      history.replaceState(null, null, newHash);
    } else {
      location.hash = newHash;
    }
  },

  updateURLLengthIndicator,
};
