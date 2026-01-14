import { getState } from "./rowColManager.js";
import { sanitizeHTML } from "./security.js";
import { showToast } from "./toastManager.js";
import { VisualFunctions } from "./visualFunctions.js";

function extractPlainText(value) {
  if (value === null || value === undefined) return "";
  const parser = new DOMParser();
  const doc = parser.parseFromString("<body>" + String(value) + "</body>", "text/html");
  const text = doc.body.textContent || "";
  return text.replace(/\u00a0/g, " ");
}

export const PresentationManager = {
  overlay: null,
  isActive: false,
  slideCount: 0,
  slides: [],
  currentIndex: 0,
  lastFocus: null,
  _initialized: false,
  _scrollHandler: null,
  _resizeHandler: null,
  _controlsEl: null,
  _scrollRaf: null,

  init() {
    if (this._initialized) return;
    this._initialized = true;

    document.addEventListener("keydown", (event) => {
      if (!this.isActive) return;

      if (event.key === "Escape") {
        event.preventDefault();
        this.stop();
        return;
      }

      if (this._handleNavigationKey(event)) {
        return;
      }
    });
  },

  start(data, formulas, context = {}) {
    if (this.isActive) return;
    if (this.overlay) {
      this.stop();
    }

    const { rows, cols } = getState();
    const rowCount = Number.isFinite(rows) && rows > 0 ? rows : 0;
    const colCount = Number.isFinite(cols) && cols > 0 ? cols : 0;
    const safeData = Array.isArray(data) ? data : [];
    const safeFormulas = Array.isArray(formulas) ? formulas : [];

    const overlay = document.createElement("div");
    overlay.className = "presentation-overlay";
    overlay.setAttribute("tabindex", "0");
    overlay.setAttribute("role", "dialog");
    overlay.setAttribute("aria-modal", "true");
    overlay.setAttribute("aria-label", "Presentation mode");

    const slides = [];
    let slideCount = 0;

    for (let r = 0; r < rowCount; r++) {
      const rowData = Array.isArray(safeData[r]) ? safeData[r] : [];
      const rowFormulas = Array.isArray(safeFormulas[r]) ? safeFormulas[r] : [];

      if (!this._rowHasContent(rowData, rowFormulas, colCount)) {
        continue;
      }

      const slide = document.createElement("section");
      slide.className = "slide";
      slide.setAttribute("data-slide-index", String(slideCount));

      const contentDiv = document.createElement("div");
      contentDiv.className = "slide-content";

      let titleFound = false;

      for (let c = 0; c < colCount; c++) {
        const rawValue = rowData[c];
        const formula = rowFormulas[c];
        const hasFormula = typeof formula === "string" && formula.trim().startsWith("=");
        const hasValue = this._cellHasText(rawValue);

        if (!hasFormula && !hasValue) continue;

        const cellWrapper = document.createElement("div");
        let isVisual = false;

        if (hasFormula) {
          try {
            const visualEl = VisualFunctions.process(String(formula), context);
            if (visualEl) {
              cellWrapper.appendChild(visualEl);
              cellWrapper.classList.add("slide-visual");
              isVisual = true;
            }
          } catch (err) {
            console.warn("Presentation visual error:", err);
          }
        }

        if (!isVisual) {
          const displayValue = this._getDisplayValue(rawValue, formula);
          cellWrapper.innerHTML = sanitizeHTML(displayValue);
        }

        if (!titleFound) {
          cellWrapper.classList.add("slide-title");
          titleFound = true;
        } else {
          cellWrapper.classList.add("slide-text");
        }

        contentDiv.appendChild(cellWrapper);
      }

      slide.appendChild(contentDiv);
      overlay.appendChild(slide);
      slides.push(slide);
      slideCount++;
    }

    if (slideCount === 0) {
      showToast("Spreadsheet is empty. Add data to present.", "warning");
      return;
    }

    const controls = document.createElement("div");
    controls.className = "presentation-controls";
    controls.innerHTML = `
      <span class="slide-counter">1 / ${slideCount}</span>
      <button id="exit-pres-btn" type="button"><i class="fa-solid fa-xmark"></i> Exit</button>
    `;
    overlay.appendChild(controls);

    document.body.appendChild(overlay);
    document.body.classList.add("presentation-mode-active");

    this.lastFocus = document.activeElement;
    overlay.focus();

    this.overlay = overlay;
    this.isActive = true;
    this.slides = slides;
    this.slideCount = slideCount;
    this._controlsEl = controls;
    this.currentIndex = 0;

    const exitBtn = overlay.querySelector("#exit-pres-btn");
    if (exitBtn) {
      exitBtn.addEventListener("click", () => this.stop());
    }

    this._scrollHandler = () => {
      if (this._scrollRaf) return;
      this._scrollRaf = requestAnimationFrame(() => {
        this._scrollRaf = null;
        this._updateCounter();
      });
    };

    this._resizeHandler = () => this._updateCounter();
    overlay.addEventListener("scroll", this._scrollHandler);
    window.addEventListener("resize", this._resizeHandler);
    this._updateCounter();
  },

  stop() {
    if (!this.overlay) return;

    if (this._scrollHandler) {
      this.overlay.removeEventListener("scroll", this._scrollHandler);
    }
    if (this._resizeHandler) {
      window.removeEventListener("resize", this._resizeHandler);
    }
    if (this._scrollRaf) {
      cancelAnimationFrame(this._scrollRaf);
      this._scrollRaf = null;
    }

    this.overlay.remove();
    this.overlay = null;
    this.isActive = false;
    this.slideCount = 0;
    this.slides = [];
    this.currentIndex = 0;
    this._controlsEl = null;
    document.body.classList.remove("presentation-mode-active");

    if (this.lastFocus && typeof this.lastFocus.focus === "function" && document.contains(this.lastFocus)) {
      this.lastFocus.focus();
    }
    this.lastFocus = null;
  },

  _updateCounter() {
    if (!this.overlay || !this._controlsEl || !this.slideCount) return;
    const slideHeight = this.overlay.clientHeight || window.innerHeight || 1;
    const index = clampIndex(Math.round(this.overlay.scrollTop / slideHeight), this.slideCount);
    this.currentIndex = index;

    const counter = this._controlsEl.querySelector(".slide-counter");
    if (counter) {
      counter.textContent = `${index + 1} / ${this.slideCount}`;
    }
  },

  _scrollToSlide(index) {
    if (!this.overlay || !this.slideCount) return;
    const target = clampIndex(index, this.slideCount);
    const slideHeight = this.overlay.clientHeight || window.innerHeight || 1;
    this.overlay.scrollTo({ top: target * slideHeight, behavior: "smooth" });
  },

  _handleNavigationKey(event) {
    if (!this.overlay || !this.slideCount) return false;

    let nextIndex = null;
    if (event.key === "ArrowDown" || event.key === "PageDown" || event.key === " ") {
      nextIndex = this.currentIndex + 1;
    } else if (event.key === "ArrowUp" || event.key === "PageUp") {
      nextIndex = this.currentIndex - 1;
    } else if (event.key === "Home") {
      nextIndex = 0;
    } else if (event.key === "End") {
      nextIndex = this.slideCount - 1;
    }

    if (nextIndex === null) return false;

    event.preventDefault();
    this._scrollToSlide(nextIndex);
    return true;
  },

  _rowHasContent(rowData, rowFormulas, colCount) {
    for (let c = 0; c < colCount; c++) {
      const formula = rowFormulas && rowFormulas[c];
      if (typeof formula === "string" && formula.trim() !== "") return true;
      if (this._cellHasText(rowData && rowData[c])) return true;
    }
    return false;
  },

  _cellHasText(value) {
    const text = extractPlainText(value);
    return text.trim() !== "";
  },

  _getDisplayValue(rawValue, formula) {
    const rawText = rawValue === null || rawValue === undefined ? "" : String(rawValue);
    if (rawText.trim() !== "") return rawText;
    if (typeof formula === "string" && formula.trim() !== "") return String(formula);
    return "";
  },
};

function clampIndex(index, count) {
  if (count <= 0) return 0;
  if (index < 0) return 0;
  if (index >= count) return count - 1;
  return index;
}
