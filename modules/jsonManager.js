const defaultCallbacks = {
  buildCurrentState: () => ({}),
  recalculateFormulas: () => {},
  showToast: () => {},
};

const elements = {
  modal: null,
  textarea: null,
  errorEl: null,
  copyBtn: null,
  docs: null,
  docsToggle: null,
};

function clearJSONError() {
  if (elements.errorEl) {
    elements.errorEl.classList.add("hidden");
    elements.errorEl.textContent = "";
  }
}

export const JSONManager = {
  callbacks: { ...defaultCallbacks },

  init(callbacks = {}) {
    this.callbacks = { ...defaultCallbacks, ...callbacks };
    elements.modal = document.getElementById("json-modal");
    elements.textarea = document.getElementById("json-editor");
    elements.errorEl = document.getElementById("json-error");
    elements.copyBtn = document.getElementById("json-copy-btn");
    elements.docs = document.getElementById("json-docs");
    elements.docsToggle = document.getElementById("json-docs-toggle");

    if (elements.docsToggle) {
      elements.docsToggle.addEventListener("click", () => {
        this.toggleDocs();
      });
    }
  },

  openModal() {
    if (!elements.modal || !elements.textarea) return;

    if (this.callbacks.recalculateFormulas) {
      this.callbacks.recalculateFormulas();
    }
    const exportState = this.callbacks.buildCurrentState ? this.callbacks.buildCurrentState() : {};
    elements.textarea.value = JSON.stringify(exportState, null, 2);
    elements.textarea.scrollTop = 0;
    clearJSONError();

    elements.modal.classList.remove("hidden");
    elements.textarea.focus();
  },

  closeModal() {
    if (elements.modal) {
      elements.modal.classList.add("hidden");
    }
    if (elements.docs) {
      elements.docs.classList.remove("visible");
    }
    if (elements.docsToggle) {
      elements.docsToggle.textContent = "Show Schema Reference";
    }
    clearJSONError();
  },

  toggleDocs() {
    if (!elements.docs || !elements.docsToggle) return;
    const isVisible = elements.docs.classList.contains("visible");
    if (isVisible) {
      elements.docs.classList.remove("visible");
      elements.docsToggle.textContent = "Show Schema Reference";
    } else {
      elements.docs.classList.add("visible");
      elements.docsToggle.textContent = "Hide Schema Reference";
    }
  },

  async copyJSONToClipboard() {
    if (!elements.textarea) return;

    try {
      await navigator.clipboard.writeText(elements.textarea.value);
      if (elements.copyBtn) {
        const original = elements.copyBtn.innerHTML;
        elements.copyBtn.innerHTML = '<i class="fa-solid fa-check"></i> Copied!';
        setTimeout(() => {
          if (elements.copyBtn) {
            elements.copyBtn.innerHTML = original;
          }
        }, 1800);
      }
      if (this.callbacks.showToast) {
        this.callbacks.showToast("JSON copied to clipboard", "success");
      }
    } catch (err) {
      elements.textarea.select();
      try {
        document.execCommand("copy");
        if (this.callbacks.showToast) {
          this.callbacks.showToast("JSON copied to clipboard", "success");
        }
      } catch (e) {
        if (this.callbacks.showToast) {
          this.callbacks.showToast("Press Ctrl+C to copy the JSON", "warning");
        }
      }
    }
  },

};
