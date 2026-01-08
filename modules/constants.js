/**
 * Constants Module for Spreadsheet Application
 * Contains all configuration constants used throughout the app
 */

// ========== Layout & Sizing Limits ==========
export const MAX_ROWS = 30;
export const MAX_COLS = 15;
export const DEBOUNCE_DELAY = 200;
export const ACTIVE_HEADER_CLASS = "header-active";
export const ROW_HEADER_WIDTH = 40;
export const HEADER_ROW_HEIGHT = 32;
export const DEFAULT_COL_WIDTH = 100;
export const MIN_COL_WIDTH = 80;

// Mobile-responsive layout detection
export const IS_MOBILE_LAYOUT = typeof window !== "undefined" && window.matchMedia && window.matchMedia("(max-width: 768px)").matches;
export const DEFAULT_ROW_HEIGHT = IS_MOBILE_LAYOUT ? 44 : 32;
export const MIN_ROW_HEIGHT = DEFAULT_ROW_HEIGHT;

// Default starting grid size
export const DEFAULT_ROWS = 10;
export const DEFAULT_COLS = 10;

// ========== URL Length Thresholds ==========
export const URL_LENGTH_WARNING = 2000; // Yellow - some older browsers may truncate
export const URL_LENGTH_CAUTION = 4000; // Orange - URL shorteners may fail
export const URL_LENGTH_CRITICAL = 8000; // Red - some browsers may fail
export const URL_LENGTH_MAX_DISPLAY = 10000; // For progress bar scaling

// ========== Key Minification Mapping for URL Compression ==========
export const KEY_MAP = {
  rows: "r",
  cols: "c",
  theme: "t",
  data: "d",
  formulas: "f",
  cellStyles: "s",
  colWidths: "w",
  rowHeights: "h",
  readOnly: "ro",
  embed: "e",
};
export const KEY_MAP_REVERSE = Object.fromEntries(Object.entries(KEY_MAP).map(([k, v]) => [v, k]));

export const STYLE_KEY_MAP = {
  align: "a",
  bg: "b",
  color: "c",
  fontSize: "z",
};
export const STYLE_KEY_MAP_REVERSE = Object.fromEntries(Object.entries(STYLE_KEY_MAP).map(([k, v]) => [v, k]));

// ========== Formula Configuration ==========
export const FORMULA_SUGGESTIONS = [
  { name: "SUM", signature: "SUM(range)", description: "Adds numbers in a range" },
  { name: "AVG", signature: "AVG(range)", description: "Average of numbers in a range" },
];

// Valid formula patterns (security whitelist)
export const VALID_FORMULA_PATTERNS = [/^=\s*SUM\s*\(\s*[A-Z]+\d+\s*:\s*[A-Z]+\d+\s*\)\s*$/i, /^=\s*AVG\s*\(\s*[A-Z]+\d+\s*:\s*[A-Z]+\d+\s*\)\s*$/i];

export const FONT_SIZE_OPTIONS = [10, 12, 14, 16, 18, 24];

// ========== HTML Sanitization ==========
// Allowed HTML tags for sanitization (preserves basic formatting)
export const ALLOWED_TAGS = ["B", "I", "U", "STRONG", "EM", "SPAN", "BR"];
export const ALLOWED_SPAN_STYLES = ["font-weight", "font-style", "text-decoration", "color", "background-color"];

// ========== Toast Notification System ==========
export const TOAST_DURATION = 3000; // Default duration in ms
export const TOAST_ICONS = {
  success: "fa-check",
  error: "fa-xmark",
  warning: "fa-exclamation",
  info: "fa-info",
};
