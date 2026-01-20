import { URLManager } from "./urlManager.js";

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

  renderTemplates() {
    if (!this.galleryContainer) return;
    
    this.galleryContainer.innerHTML = "";
    const grid = document.createElement("div");
    grid.className = "templates-grid";

    this.templates.forEach(template => {
      const card = document.createElement("div");
      card.className = "template-card";
      
      const icon = template.icon || "fa-table";
      
      card.innerHTML = `
        <div class="template-icon">
          <i class="fa-solid ${icon}"></i>
        </div>
        <div class="template-info">
          <h3>${template.name}</h3>
          <p>${template.description}</p>
        </div>
        <div class="template-actions">
           <button class="use-template-btn" data-id="${template.id}">Use Template</button>
        </div>
      `;

      // Handle click
      const btn = card.querySelector(".use-template-btn");
      btn.addEventListener("click", () => {
        this.useTemplate(template);
      });

      grid.appendChild(card);
    });

    this.galleryContainer.appendChild(grid);
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
    
    // this.closeGallery(); // Kept open as requested
  }
};
