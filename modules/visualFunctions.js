import { isValidCSSColor } from "./security.js";
import { parseCellRef } from "./formulaManager.js";

const COLOR_ALIASES = {
  green: "#22c55e",
  red: "#ef4444",
  blue: "#3b82f6",
  orange: "#f97316",
  yellow: "#facc15",
  gray: "#9ca3af",
  black: "#111111",
  white: "#ffffff",
};

function splitArgs(raw) {
  if (!raw) return [];
  if (!raw.trim()) return [];

  const args = [];
  let current = "";
  let quote = null;

  for (let i = 0; i < raw.length; i++) {
    const ch = raw[i];
    if (quote) {
      if (ch === "\\" && i + 1 < raw.length) {
        current += raw[i + 1];
        i++;
        continue;
      }
      if (ch === quote) {
        quote = null;
        continue;
      }
      current += ch;
      continue;
    }

    if (ch === "'" || ch === '"') {
      quote = ch;
      continue;
    }

    if (ch === ",") {
      args.push(current.trim());
      current = "";
      continue;
    }

    current += ch;
  }

  if (current.length || raw.trim().length) {
    args.push(current.trim());
  }

  return args;
}

function parseFormula(formula) {
  if (!formula || typeof formula !== "string") return null;
  const trimmed = formula.trim();
  if (!trimmed.startsWith("=")) return null;

  const body = trimmed.slice(1).trim();
  const openIndex = body.indexOf("(");
  const closeIndex = body.lastIndexOf(")");
  if (openIndex <= 0 || closeIndex <= openIndex) return null;

  const name = body.slice(0, openIndex).trim().toUpperCase();
  const trailing = body.slice(closeIndex + 1).trim();
  if (trailing) return null;

  const argsRaw = body.slice(openIndex + 1, closeIndex);
  return { name, args: splitArgs(argsRaw) };
}

function extractPlainText(value) {
  if (value === null || value === undefined) return "";
  const parser = new DOMParser();
  const doc = parser.parseFromString("<body>" + String(value) + "</body>", "text/html");
  const text = doc.body.textContent || "";
  return text.replace(/\u00a0/g, " ");
}

function isCellRef(value) {
  return /^[A-Z]+[1-9]\d*$/i.test(value);
}

function resolveNumberArg(arg, context) {
  if (arg === null || arg === undefined) return null;
  const trimmed = String(arg).trim();
  if (!trimmed) return null;

  if (isCellRef(trimmed)) {
    const ref = parseCellRef(trimmed);
    if (ref && context && typeof context.getCellValue === "function") {
      const value = context.getCellValue(ref.row, ref.col);
      return Number.isFinite(value) ? value : 0;
    }
  }

  const percentMatch = trimmed.match(/^(-?\d+(?:\.\d+)?)%$/);
  const numericText = percentMatch ? percentMatch[1] : trimmed.replace(/,/g, "");
  const num = parseFloat(numericText);
  return Number.isNaN(num) ? null : num;
}

function resolveTextArg(arg, context) {
  if (arg === null || arg === undefined) return "";
  const trimmed = String(arg).trim();
  if (!trimmed) return "";

  if (isCellRef(trimmed)) {
    const ref = parseCellRef(trimmed);
    if (ref && context && Array.isArray(context.data)) {
      const value = context.data[ref.row] && context.data[ref.row][ref.col];
      return extractPlainText(value);
    }
  }

  return trimmed;
}

function resolveColor(value) {
  if (!value) return "";
  const trimmed = String(value).trim();
  if (!trimmed) return "";
  const lower = trimmed.toLowerCase();
  if (COLOR_ALIASES[lower]) return COLOR_ALIASES[lower];
  if (isValidCSSColor(trimmed)) return trimmed;
  return "";
}

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function getContrastColor(color) {
  if (!color) return "#ffffff";

  let r = 0;
  let g = 0;
  let b = 0;
  const hexMatch = color.match(/^#([0-9a-f]{3}|[0-9a-f]{6}|[0-9a-f]{8})$/i);
  const rgbMatch = color.match(/^rgba?\((\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})/i);

  if (hexMatch) {
    const hex = hexMatch[1];
    if (hex.length === 3) {
      r = parseInt(hex[0] + hex[0], 16);
      g = parseInt(hex[1] + hex[1], 16);
      b = parseInt(hex[2] + hex[2], 16);
    } else {
      r = parseInt(hex.slice(0, 2), 16);
      g = parseInt(hex.slice(2, 4), 16);
      b = parseInt(hex.slice(4, 6), 16);
    }
  } else if (rgbMatch) {
    r = parseInt(rgbMatch[1], 10);
    g = parseInt(rgbMatch[2], 10);
    b = parseInt(rgbMatch[3], 10);
  } else {
    return "#ffffff";
  }

  const luminance = (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255;
  return luminance > 0.6 ? "#111111" : "#ffffff";
}

function buildProgress(percent, label, color) {
  const normalized = clamp(Number.isFinite(percent) ? percent : 0, 0, 100);
  const wrapper = document.createElement("div");
  wrapper.className = "visual-progress";

  const track = document.createElement("div");
  track.className = "visual-progress-track";
  track.setAttribute("role", "progressbar");
  track.setAttribute("aria-valuemin", "0");
  track.setAttribute("aria-valuemax", "100");
  track.setAttribute("aria-valuenow", String(Math.round(normalized)));

  const fill = document.createElement("div");
  fill.className = "visual-progress-fill";
  fill.style.width = `${normalized}%`;
  if (color) {
    fill.style.backgroundColor = color;
    fill.style.color = getContrastColor(color);
  }

  if (label) {
    fill.textContent = label;
  } else {
    fill.textContent = `${Math.round(normalized)}%`;
  }

  track.appendChild(fill);
  wrapper.appendChild(track);

  if (label) {
    const labelEl = document.createElement("div");
    labelEl.className = "visual-progress-label";
    labelEl.textContent = `${Math.round(normalized)}%`;
    wrapper.appendChild(labelEl);
  }

  return wrapper;
}

function buildTag(label, color) {
  if (!label) return null;
  const tag = document.createElement("span");
  tag.className = "visual-tag";
  tag.textContent = label;

  if (color) {
    tag.style.backgroundColor = color;
    tag.style.borderColor = color;
    tag.style.color = getContrastColor(color);
  }

  return tag;
}

function buildRating(value, max) {
  if (!Number.isFinite(value)) return null;
  const total = clamp(Math.round(Number.isFinite(max) ? max : 5), 1, 10);
  const score = clamp(Math.round(value), 0, total);

  const rating = document.createElement("div");
  rating.className = "visual-rating";
  rating.setAttribute("role", "img");
  rating.setAttribute("aria-label", `${score} of ${total}`);

  for (let i = 0; i < total; i++) {
    const star = document.createElement("i");
    star.className = i < score ? "fa-solid fa-star" : "fa-regular fa-star";
    rating.appendChild(star);
  }

  return rating;
}

export const VisualFunctions = {
  process(formula, context = {}) {
    const parsed = parseFormula(formula);
    if (!parsed) return null;

    const { name, args } = parsed;
    if (name === "PROGRESS") {
      const value = resolveNumberArg(args[0], context);
      if (value === null) return null;

      const secondNumeric = resolveNumberArg(args[1], context);
      let percent = value;
      let label = "";
      let color = "";

      if (secondNumeric !== null) {
        percent = secondNumeric === 0 ? 0 : (value / secondNumeric) * 100;
        label = resolveTextArg(args[2], context);
        color = resolveColor(args[3]);
      } else {
        if (args.length === 1 && value >= 0 && value <= 1) {
          percent = value * 100;
        }

        const maybeColor = resolveColor(args[1]);
        if (maybeColor && args.length === 2) {
          color = maybeColor;
        } else {
          label = resolveTextArg(args[1], context);
          color = resolveColor(args[2]);
        }
      }

      return buildProgress(percent, label, color);
    }

    if (name === "TAG") {
      const label = resolveTextArg(args[0], context);
      if (!label) return null;
      const color = resolveColor(args[1]);
      return buildTag(label, color);
    }

    if (name === "RATING") {
      const value = resolveNumberArg(args[0], context);
      if (value === null) return null;
      const max = resolveNumberArg(args[1], context);
      return buildRating(value, max);
    }

    return null;
  },
};
