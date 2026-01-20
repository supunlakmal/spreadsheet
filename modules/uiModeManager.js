/**
 * uiModeManager.js
 * Manages read-only mode, embed mode, and related UI states
 */

let buildCurrentState;
let updateURL;
let scheduleFullSync;
let showToast;
let URLManager;
let PasswordManager;
let getReadOnlyFlag;
let setReadOnlyFlag;
let getEmbedModeFlag;
let setEmbedModeFlag;

export const UIModeManager = {
  /**
   * Initialize the UI mode manager with required callbacks
   * @param {Object} callbacks - Required callbacks from script.js
   */
  init(callbacks) {
    buildCurrentState = callbacks.buildCurrentState;
    updateURL = callbacks.updateURL;
    scheduleFullSync = callbacks.scheduleFullSync;
    showToast = callbacks.showToast;
    URLManager = callbacks.URLManager;
    PasswordManager = callbacks.PasswordManager;
    getReadOnlyFlag = callbacks.getReadOnlyFlag;
    setReadOnlyFlag = callbacks.setReadOnlyFlag;
    getEmbedModeFlag = callbacks.getEmbedModeFlag;
    setEmbedModeFlag = callbacks.setEmbedModeFlag;
  },

  /**
   * Apply read-only mode UI state
   * Disables editing and updates visual indicators
   */
  applyReadOnlyMode() {
    document.body.classList.add("readonly-mode");

    // Set contentEditable on all cells
    const container = document.getElementById("spreadsheet");
    if (container) {
      container.querySelectorAll(".cell-content").forEach((cell) => {
        cell.contentEditable = "false";
      });
    }

    // Update toggle button
    const toggleBtn = document.getElementById("toggle-readonly");
    if (toggleBtn) {
      toggleBtn.classList.add("active");
      const icon = toggleBtn.querySelector("i");
      if (icon) icon.className = "fa-solid fa-eye";
    }

    // Show banner
    const banner = document.getElementById("readonly-banner");
    if (banner) banner.classList.remove("hidden");
  },

  /**
   * Clear read-only mode UI state
   * Re-enables editing and updates visual indicators
   */
  clearReadOnlyMode() {
    document.body.classList.remove("readonly-mode");

    // Reset contentEditable
    const container = document.getElementById("spreadsheet");
    if (container) {
      container.querySelectorAll(".cell-content").forEach((cell) => {
        cell.contentEditable = "true";
      });
    }

    // Update toggle button
    const toggleBtn = document.getElementById("toggle-readonly");
    if (toggleBtn) {
      toggleBtn.classList.remove("active");
      const icon = toggleBtn.querySelector("i");
      if (icon) icon.className = "fa-solid fa-pen-to-square";
    }

    // Hide banner
    const banner = document.getElementById("readonly-banner");
    if (banner) banner.classList.add("hidden");
  },

  /**
   * Apply embed mode UI state
   * Sets read-only mode and adds embed-specific styling
   */
  applyEmbedMode() {
    document.body.classList.add("embed-mode");
    setReadOnlyFlag(true);
    this.applyReadOnlyMode();
  },

  /**
   * Clear embed mode UI state
   */
  clearEmbedMode() {
    document.body.classList.remove("embed-mode");
  },

  /**
   * Toggle read-only mode
   * Updates URL and syncs with P2P peers
   */
  toggleReadOnlyMode() {
    const currentReadOnly = getReadOnlyFlag();
    setReadOnlyFlag(!currentReadOnly);
    this.applyReadOnlyState(!currentReadOnly);
    showToast(
      !currentReadOnly ? "Read-only mode enabled - Share this link for view-only access" : "Edit mode enabled",
      "success"
    );

    // Update URL immediately (no debounce)
    updateURL();
    scheduleFullSync();
  },

  /**
   * Apply read-only state based on flag
   * @param {boolean} readOnlyFlag - Whether to enable read-only mode
   */
  applyReadOnlyState(readOnlyFlag) {
    // If in embed mode, always stay read-only
    if (getEmbedModeFlag()) {
      setReadOnlyFlag(true);
      this.applyReadOnlyMode();
      return;
    }

    setReadOnlyFlag(readOnlyFlag || false);
    if (readOnlyFlag) {
      this.applyReadOnlyMode();
    } else {
      this.clearReadOnlyMode();
    }
  },

  /**
   * Generate embed code for the current spreadsheet
   * @returns {Promise<string|null>} HTML iframe embed code or null if already in embed mode
   */
  async generateEmbedCode() {
    if (getEmbedModeFlag()) {
      showToast("Already in embed mode. Share current URL instead.", "warning");
      return null;
    }

    const currentState = buildCurrentState();
    currentState.embed = 1;
    currentState.readOnly = 1;

    const encoded = await URLManager.encodeState(currentState, PasswordManager.getPassword());
    const embedURL = window.location.origin + window.location.pathname + "#" + encoded;

    if (embedURL.length > 2000) {
      showToast("Warning: Embed URL is very long", "warning");
    }

    return `<iframe
    src="${embedURL}"
    width="800"
    height="600"
    frameborder="0"
    style="border: 1px solid #e0e0e0; border-radius: 8px;"
    title="Embedded Spreadsheet">
</iframe>`;
  },

  /**
   * Show embed modal with generated code
   */
  async showEmbedModal() {
    const embedCode = await this.generateEmbedCode();
    if (!embedCode) return;

    const modal = document.getElementById("embed-modal");
    const textarea = document.getElementById("embed-code-textarea");

    textarea.value = embedCode;
    modal.classList.remove("hidden");
    textarea.select();
  },

  /**
   * Hide embed modal
   */
  hideEmbedModal() {
    const modal = document.getElementById("embed-modal");
    modal.classList.add("hidden");
  },
};
