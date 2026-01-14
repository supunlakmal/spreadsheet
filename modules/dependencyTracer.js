/**
 * Visual Formula Dependency Tracer
 * Draws Bezier curves between formula cells and their data sources.
 */
import { parseRange, parseCellRef } from "./formulaManager.js";
import { getState, getCellElement } from "./rowColManager.js";

const SVG_NS = "http://www.w3.org/2000/svg";
const ARROW_ID = "dependency-arrowhead";
const LINE_COLOR = "#ff0055";

function normalizeRef(ref) {
  return ref.replace(/\$/g, "").toUpperCase();
}

function extractRefs(formula) {
  if (!formula || typeof formula !== "string") return [];
  if (!formula.startsWith("=")) return [];
  const matches = formula.match(/(\$?[A-Z]+\$?\d+)(?::(\$?[A-Z]+\$?\d+))?/gi);
  return matches || [];
}

function isInBounds(row, col, rows, cols) {
  return row >= 0 && col >= 0 && row < rows && col < cols;
}

export const DependencyTracer = {
  isActive: false,
  container: null,
  svg: null,
  resizeObserver: null,
  lastFormulas: null,

  init() {
    const container = document.getElementById("spreadsheet");
    if (!container) return;

    this.container = container;

    if (!this.svg || !container.contains(this.svg)) {
      this.svg = this.createLayer();
      this.container.appendChild(this.svg);
    }

    if (this.isActive) {
      this.svg.classList.remove("hidden");
    } else {
      this.svg.classList.add("hidden");
    }

    this.updateSvgSize();

    if (this.resizeObserver) {
      this.resizeObserver.disconnect();
      this.resizeObserver = null;
    }

    if (typeof ResizeObserver !== "undefined") {
      this.resizeObserver = new ResizeObserver(() => {
        if (this.isActive) {
          this.draw(this.lastFormulas);
        }
      });
      this.resizeObserver.observe(this.container);
    }
  },

  createLayer() {
    const svg = document.createElementNS(SVG_NS, "svg");
    svg.classList.add("dependency-layer", "hidden");
    svg.setAttribute("aria-hidden", "true");

    const defs = document.createElementNS(SVG_NS, "defs");
    const marker = document.createElementNS(SVG_NS, "marker");
    marker.setAttribute("id", ARROW_ID);
    marker.setAttribute("markerWidth", "10");
    marker.setAttribute("markerHeight", "7");
    marker.setAttribute("refX", "9");
    marker.setAttribute("refY", "3.5");
    marker.setAttribute("orient", "auto");

    const polygon = document.createElementNS(SVG_NS, "polygon");
    polygon.setAttribute("points", "0 0, 10 3.5, 0 7");
    polygon.setAttribute("fill", LINE_COLOR);

    marker.appendChild(polygon);
    defs.appendChild(marker);
    svg.appendChild(defs);

    return svg;
  },

  ensureLayer() {
    const container = document.getElementById("spreadsheet");
    if (!container) return false;

    if (this.container !== container) {
      this.container = container;
    }

    if (!this.svg || !container.contains(this.svg)) {
      this.svg = this.createLayer();
      this.container.appendChild(this.svg);
    }

    if (this.isActive) {
      this.svg.classList.remove("hidden");
    } else {
      this.svg.classList.add("hidden");
    }

    this.updateSvgSize();
    return true;
  },

  updateSvgSize() {
    if (!this.container || !this.svg) return;
    const width = Math.max(this.container.scrollWidth, this.container.clientWidth);
    const height = Math.max(this.container.scrollHeight, this.container.clientHeight);
    this.svg.setAttribute("width", width);
    this.svg.setAttribute("height", height);
    this.svg.setAttribute("viewBox", `0 0 ${width} ${height}`);
  },

  toggle() {
    this.isActive = !this.isActive;
    if (!this.ensureLayer()) return this.isActive;

    if (this.isActive) {
      this.svg.classList.remove("hidden");
      this.draw(this.lastFormulas);
    } else {
      this.svg.classList.add("hidden");
      this.clear();
    }
    return this.isActive;
  },

  clear() {
    if (!this.svg) return;
    this.svg.querySelectorAll("path").forEach((path) => path.remove());
  },

  getCellCenter(row, col) {
    const cell = getCellElement(row, col);
    if (!cell || !this.container) return null;

    const rect = cell.getBoundingClientRect();
    const containerRect = this.container.getBoundingClientRect();

    return {
      x: rect.left - containerRect.left + rect.width / 2,
      y: rect.top - containerRect.top + rect.height / 2,
    };
  },

  getRangeCenter(startRow, startCol, endRow, endCol) {
    const topLeft = this.getCellCenter(startRow, startCol);
    const bottomRight = this.getCellCenter(endRow, endCol);
    if (!topLeft || !bottomRight) return null;

    return {
      x: (topLeft.x + bottomRight.x) / 2,
      y: (topLeft.y + bottomRight.y) / 2,
    };
  },

  draw(formulas) {
    if (!this.isActive) return;
    if (!this.ensureLayer()) return;

    const formulaData = Array.isArray(formulas) ? formulas : this.lastFormulas;
    if (Array.isArray(formulas)) {
      this.lastFormulas = formulas;
    }
    this.clear();

    const { rows, cols } = getState();
    if (!Array.isArray(formulaData) || rows <= 0 || cols <= 0) return;

    const drawPath = (source, target) => {
      if (!this.svg) return;
      if (!source || !target) return;
      if (source.x === target.x && source.y === target.y) return;

      const path = document.createElementNS(SVG_NS, "path");
      const deltaX = target.x - source.x;
      const c1x = source.x + deltaX * 0.5;
      const c1y = source.y;
      const c2x = target.x - deltaX * 0.5;
      const c2y = target.y;
      const d = `M ${source.x} ${source.y} C ${c1x} ${c1y}, ${c2x} ${c2y}, ${target.x} ${target.y}`;

      path.setAttribute("d", d);
      path.setAttribute("stroke", LINE_COLOR);
      path.setAttribute("stroke-width", "2");
      path.setAttribute("fill", "none");
      path.setAttribute("opacity", "0.6");
      path.setAttribute("marker-end", `url(#${ARROW_ID})`);
      path.classList.add("dependency-line");

      this.svg.appendChild(path);

      const length = path.getTotalLength();
      path.style.strokeDasharray = `${length}`;
      path.style.strokeDashoffset = `${length}`;
      path.style.animation = "dash 1.5s ease-out forwards";
    };

    for (let r = 0; r < rows; r++) {
      const rowFormulas = formulaData[r];
      if (!Array.isArray(rowFormulas)) continue;

      for (let c = 0; c < cols; c++) {
        const formula = rowFormulas[c];
        if (!formula || typeof formula !== "string" || !formula.startsWith("=")) continue;

        const targetCenter = this.getCellCenter(r, c);
        if (!targetCenter) continue;

        const refs = extractRefs(formula);
        if (!refs.length) continue;

        refs.forEach((refStr) => {
          const cleaned = normalizeRef(refStr);

          if (cleaned.includes(":")) {
            const range = parseRange(cleaned);
            if (!range) return;
            if (!isInBounds(range.startRow, range.startCol, rows, cols)) return;
            if (!isInBounds(range.endRow, range.endCol, rows, cols)) return;
            const sourceCenter = this.getRangeCenter(range.startRow, range.startCol, range.endRow, range.endCol);
            drawPath(sourceCenter, targetCenter);
            return;
          }

          const cellRef = parseCellRef(cleaned);
          if (!cellRef) return;
          if (!isInBounds(cellRef.row, cellRef.col, rows, cols)) return;
          const sourceCenter = this.getCellCenter(cellRef.row, cellRef.col);
          drawPath(sourceCenter, targetCenter);
        });
      }
    }
  },
};
