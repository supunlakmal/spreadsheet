import { URLManager } from "./urlManager.js";
import { ExcelManager } from "./excelManager.js";

export const TemplateManager = {
  templates: [],

  init() {
    this.cacheDOM();
    this.bindEvents();
  },

  cacheDOM() {
    this.modal = document.getElementById("templates-modal");
    this.closeBtn = document.getElementById("templates-close-btn");
    this.galleryContainer = document.getElementById("templates-gallery");
    this.openBtn = document.getElementById("open-templates-btn"); // Button in Tools modal
    this.searchInput = document.getElementById("template-search-input");
  },

  bindEvents() {
    if (this.openBtn) {
      this.openBtn.addEventListener("click", () => this.openGallery());
    }
    if (this.closeBtn) {
      this.closeBtn.addEventListener("click", () => this.closeGallery());
    }
    if (this.modal) {
      this.modal.addEventListener("click", (e) => {
        if (e.target === this.modal) this.closeGallery();
      });
    }
    
    // Check initial hash
    this.checkHash();

    // Listen for hash changes
    window.addEventListener("hashchange", () => this.checkHash());

    // Listen for search input
    if (this.searchInput) {
      this.searchInput.addEventListener("input", (e) => {
        this.filterTemplates(e.target.value);
      });
    }
  },

  checkHash() {
    if (window.location.hash === "#gallery") {
      this.openGallery();
    }
  },

  async openGallery() {
    if (this.modal) {
      this.modal.classList.remove("hidden");
      if (this.templates.length === 0) {
        await this.loadTemplates();
      }
    }
  },

  closeGallery() {
    if (this.modal) {
      this.modal.classList.add("hidden");
      
      // If closing and hash is #gallery, clear it
      if (window.location.hash === "#gallery") {
        history.pushState("", document.title, window.location.pathname + window.location.search);
      }
    }
  },

  async loadTemplates() {
    try {
      this.galleryContainer.innerHTML = '<div class="loading-spinner"><i class="fa-solid fa-circle-notch fa-spin"></i> Loading templates...</div>';
      const response = await fetch("templates.json");
      if (!response.ok) throw new Error("Failed to load templates");
      this.templates = await response.json();
      this.renderTemplates();
    } catch (error) {
      console.error("Error loading templates:", error);
      this.galleryContainer.innerHTML = '<div class="error-message">Failed to load templates. Please try again later.</div>';
    }
  },

  renderTemplates(templatesToRender = null) {
    if (!this.galleryContainer) return;
    
    const templates = templatesToRender || this.templates;
    this.galleryContainer.innerHTML = "";

    if (templates.length === 0) {
      this.galleryContainer.innerHTML = '<div class="no-results">No templates found matching your search.</div>';
      return;
    }

    const grid = document.createElement("div");
    grid.className = "templates-grid";

    templates.forEach(template => {
      const card = document.createElement("div");
      card.className = "template-card";
      
      const icon = template.icon || "fa-table";
      
      card.innerHTML = `
        <div class="template-header">
           <div class="template-icon">
              <i class="fa-solid ${icon}"></i>
           </div>
        </div>
        <div class="template-info">
          <h3>${template.name}</h3>
          <p>${template.description}</p>
        </div>
        <div class="template-actions">
           <button class="use-template-btn" data-id="${template.id}">Use Template</button>
           <button class="template-download-btn" title="Download Excel (.xlsx)" data-id="${template.id}">
              <i class="fa-solid fa-file-excel"></i>
              <span>Excel</span>
           </button>
        </div>
      `;

      // Use Template button
      card.querySelector(".use-template-btn").addEventListener("click", () => {
        this.useTemplate(template);
      });

      // Download Excel button
      card.querySelector(".template-download-btn").addEventListener("click", (e) => {
        e.stopPropagation();
        this.downloadTemplateAsExcel(template);
      });

      grid.appendChild(card);
    });

    this.galleryContainer.appendChild(grid);
  },

  async downloadTemplateAsExcel(template) {
    if (!template || !template.data) return;

    try {
      // 1. Expand the minified state
      const state = ExcelManager.expandMinifiedState(template.data);
      if (!state) return;

      // 2. Export
      const filename = `${template.name.toLowerCase().replace(/\s+/g, "_")}_template.xlsx`;
      await ExcelManager.exportFromState(state, filename);
    } catch (error) {
      console.error("Failed to download template as Excel:", error);
    }
  },

  async useTemplate(template) {
    if (!template) return;

    let hash = "";

    // 1. If 'data' object exists, encode it on the fly
    if (template.data && typeof template.data === "object") {
        try {
            // Encode state to hash (no password for templates)
            hash = await URLManager.encodeState(template.data);
        } catch (e) {
            console.error("Failed to encode template data:", e);
            return;
        }
    } 
    // 2. Legacy 'link' support
    else if (template.link) {
         const link = template.link;
         if (link.startsWith("/#") || link.startsWith("#")) {
            hash = link.includes("#") ? link.split("#")[1] : link;
         }
    }

    if (hash) {
       const url = window.location.origin + window.location.pathname + "#" + hash;
       window.open(url, "_blank");
    } else {
       console.warn("Invalid template format:", template);
    }
  },

  filterTemplates(query) {
    if (!query) {
      this.renderTemplates(this.templates);
      return;
    }

    const lowerQuery = query.toLowerCase();
    const filtered = this.templates.filter(template => {
      const nameMatch = template.name.toLowerCase().includes(lowerQuery);
      const descMatch = template.description && template.description.toLowerCase().includes(lowerQuery);
      return nameMatch || descMatch;
    });

    this.renderTemplates(filtered);
  }
};
