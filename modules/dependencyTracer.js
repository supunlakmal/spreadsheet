/**
 * Visual Formula Dependency Tracer
 * Draws Bezier curves between formula cells and their data sources.
 */
import { parseRange, parseCellRef } from "./formulaManager.js";
import { getState, getCellElement } from "./rowColManager.js";

const SVG_NS = "http://www.w3.org/2000/svg";
const ARROW_ID = "dependency-arrowhead";
const DOT_ID = "dependency-dot";
const LINE_COLOR = "#2196F3"; // Professional Blue

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

    // Arrow Marker (End) - Smaller Size
    const arrowMarker = document.createElementNS(SVG_NS, "marker");
    arrowMarker.setAttribute("id", ARROW_ID);
    arrowMarker.setAttribute("markerWidth", "6");
    arrowMarker.setAttribute("markerHeight", "6");
    arrowMarker.setAttribute("refX", "5"); // Tip
    arrowMarker.setAttribute("refY", "3");
    arrowMarker.setAttribute("orient", "auto");

    const arrowPath = document.createElementNS(SVG_NS, "path");
    arrowPath.setAttribute("d", "M0,0 L6,3 L0,6 L1.5,3 z"); // Sharper, smaller arrow
    arrowPath.setAttribute("fill", LINE_COLOR);
    arrowMarker.appendChild(arrowPath);

    // Dot Marker (Start)
    const dotMarker = document.createElementNS(SVG_NS, "marker");
    dotMarker.setAttribute("id", DOT_ID);
    dotMarker.setAttribute("markerWidth", "8");
    dotMarker.setAttribute("markerHeight", "8");
    dotMarker.setAttribute("refX", "4"); // Center of dot
    dotMarker.setAttribute("refY", "4");
    dotMarker.setAttribute("orient", "auto");

    const dotCircle = document.createElementNS(SVG_NS, "circle");
    dotCircle.setAttribute("cx", "4");
    dotCircle.setAttribute("cy", "4");
    dotCircle.setAttribute("r", "3");
    dotCircle.setAttribute("fill", LINE_COLOR);
    dotMarker.appendChild(dotCircle);

    defs.appendChild(arrowMarker);
    defs.appendChild(dotMarker);
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
    this.clearHighlights();
  },

  clearHighlights() {
    if (!this.container) return;
    const sources = this.container.querySelectorAll(".dependency-source");
    const targets = this.container.querySelectorAll(".dependency-target");
    sources.forEach((el) => el.classList.remove("dependency-source"));
    targets.forEach((el) => el.classList.remove("dependency-target"));
  },

  getCellCenter(row, col) {
    return this.getCellAnchor(row, col, 'center');
  },
  
  getCellAnchor(row, col, side = 'center') {
    const cell = getCellElement(row, col);
    if (!cell || !this.container) return null;

    const rect = cell.getBoundingClientRect();
    const containerRect = this.container.getBoundingClientRect();

    const topOffset = rect.top - containerRect.top + 8; // Near top (corner-like)
    
    if (side === 'left') {
        return {
            x: rect.left - containerRect.left + 2, // Left edge
            y: topOffset,
        };
    }
    
    if (side === 'right') {
        return {
            x: rect.right - containerRect.left - 2, // Right edge
            y: topOffset,
        };
    }

    // Default center
    return {
      x: rect.left - containerRect.left + rect.width / 2,
      y: rect.top - containerRect.top + rect.height / 2,
    };
  },

  getRangeCenter(startRow, startCol, endRow, endCol) {
     // For ranges, we'll just use the visual center for now to avoid complexity
    const topLeft = this.getCellCenter(startRow, startCol);
    const bottomRight = this.getCellCenter(endRow, endCol);
    if (!topLeft || !bottomRight) return null;

    return {
      x: (topLeft.x + bottomRight.x) / 2,
      y: (topLeft.y + bottomRight.y) / 2,
    };
  },

  highlightCell(row, col, type) {
    const cell = getCellElement(row, col);
    if (cell) {
        cell.classList.add(type === "source" ? "dependency-source" : "dependency-target");
    }
  },

  draw(formulas) {
    if (!this.isActive) return;
    if (!this.ensureLayer()) return;

    const formulaData = Array.isArray(formulas) ? formulas : this.lastFormulas;
    if (Array.isArray(formulas)) {
      this.lastFormulas = formulas;
    }
    this.clear(); // This now calls clearHighlights too

    const { rows, cols } = getState();
    if (!Array.isArray(formulaData) || rows <= 0 || cols <= 0) return;

    const drawPath = (source, target, direction) => {
      if (!this.svg) return;
      if (!source || !target) return;
      if (source.x === target.x && source.y === target.y) return;

      const path = document.createElementNS(SVG_NS, "path");
      
      const deltaX = target.x - source.x;
      const deltaY = target.y - source.y;
      
      let c1x, c1y, c2x, c2y;

      if (direction === 'right-to-left') {
          // R -> L: Curve out to left from source, enter from right to target
          c1x = source.x - 30;
          c1y = source.y;
          c2x = target.x + 30;
          c2y = target.y;
      } else if (direction === 'vertical') {
           // Same column: Curve out to right and back in
          c1x = source.x + 40;
          c1y = source.y + deltaY * 0.2;
          c2x = target.x + 40;
          c2y = target.y - deltaY * 0.2;
      } else {
          // L -> R (Standard): Curve right from source, enter left to target
          c1x = source.x + 30;
          c1y = source.y;
          c2x = target.x - 30;
          c2y = target.y;
      }

      const d = `M ${source.x} ${source.y} C ${c1x} ${c1y}, ${c2x} ${c2y}, ${target.x} ${target.y}`;

      path.setAttribute("d", d);
      path.setAttribute("stroke", LINE_COLOR);
      path.setAttribute("stroke-width", "2");
      path.setAttribute("fill", "none");
      path.setAttribute("opacity", "0.8");
      path.setAttribute("marker-start", `url(#${DOT_ID})`);
      path.setAttribute("marker-end", `url(#${ARROW_ID})`); 
      path.classList.add("dependency-line");

      this.svg.appendChild(path);
    };

    for (let r = 0; r < rows; r++) {
      const rowFormulas = formulaData[r];
      if (!Array.isArray(rowFormulas)) continue;

      for (let c = 0; c < cols; c++) {
        const formula = rowFormulas[c];
        if (!formula || typeof formula !== "string" || !formula.startsWith("=")) continue;

        const refs = extractRefs(formula);
        if (!refs.length) continue;

        // Highlight target (the cell containing the formula)
        this.highlightCell(r, c, "target");

        refs.forEach((refStr) => {
          const cleaned = normalizeRef(refStr);

          // Handle Range refs (keeping simple center logic or default L->R for now to avoid complexity explosion)
          if (cleaned.includes(":")) {
            const range = parseRange(cleaned);
            if (!range) return;
            if (!isInBounds(range.startRow, range.startCol, rows, cols)) return;
            if (!isInBounds(range.endRow, range.endCol, rows, cols)) return;
            
            for (let rr = range.startRow; rr <= range.endRow; rr++) {
                for (let cc = range.startCol; cc <= range.endCol; cc++) {
                    this.highlightCell(rr, cc, "source");
                }
            }

            // For ranges, we just draw from center of range to target 'left' default
            const sourceCenter = this.getRangeCenter(range.startRow, range.startCol, range.endRow, range.endCol);
            const targetCenter = this.getCellAnchor(r, c, 'left');
            drawPath(sourceCenter, targetCenter, 'left-to-right');
            return;
          }

          const cellRef = parseCellRef(cleaned);
          if (!cellRef) return;
          if (!isInBounds(cellRef.row, cellRef.col, rows, cols)) return;
          
          // Highlight source cell
          this.highlightCell(cellRef.row, cellRef.col, "source");
          
          // Determine Direction
          let sourceSide, targetSide, direction;
          
          if (c > cellRef.col) {
              // Target is to the RIGHT of Source (Standard Reading Order)
              // Source -> [Right Edge] ..... [Left Edge] -> Target
              sourceSide = 'right';
              targetSide = 'left';
              direction = 'left-to-right';
          } else if (c < cellRef.col) {
              // Target is to the LEFT of Source (Reverse Flow)
              // Source -> [Left Edge] ..... [Right Edge] -> Target
              sourceSide = 'left';
              targetSide = 'right';
              direction = 'right-to-left';
          } else {
              // Same Column (Vertical)
              // Use Right-to-Right loop to avoid text
              sourceSide = 'right';
              targetSide = 'right';
              direction = 'vertical';
          }

          const sourcePoint = this.getCellAnchor(cellRef.row, cellRef.col, sourceSide);
          const targetPoint = this.getCellAnchor(r, c, targetSide);
          
          drawPath(sourcePoint, targetPoint, direction);
        });
      }
    }
  },
};
