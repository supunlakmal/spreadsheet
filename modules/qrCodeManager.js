/**
 * qrCodeManager.js
 * Manages QR code generation and URL copying functionality
 */

const MAX_QR_URL_LENGTH = 2000;

let recalculateFormulas;
let updateURL;
let showToast;

export const QRCodeManager = {
  /**
   * Initialize the QR code manager with required callbacks
   * @param {Object} callbacks - Required callbacks from script.js
   * @param {Function} callbacks.recalculateFormulas - Recalculate all formulas
   * @param {Function} callbacks.updateURL - Update the URL with current state
   * @param {Function} callbacks.showToast - Show toast notification
   */
  init(callbacks) {
    recalculateFormulas = callbacks.recalculateFormulas;
    updateURL = callbacks.updateURL;
    showToast = callbacks.showToast;
  },

  /**
   * Copy current URL to clipboard
   * Shows visual feedback on success
   */
  copyURL() {
    const url = window.location.href;
    const copyBtn = document.getElementById("copy-url");

    navigator.clipboard
      .writeText(url)
      .then(function () {
        // Show success feedback
        showToast("Link copied to clipboard!", "success");
        if (copyBtn) {
          copyBtn.classList.add("copied");
          const icon = copyBtn.querySelector("i");
          if (icon) {
            icon.className = "fa-solid fa-check";
          }

          // Reset after 2 seconds
          setTimeout(function () {
            copyBtn.classList.remove("copied");
            if (icon) {
              icon.className = "fa-solid fa-copy";
            }
          }, 2000);
        }
      })
      .catch(function (err) {
        console.error("Failed to copy URL:", err);
        // Fallback for older browsers
        const textArea = document.createElement("textarea");
        textArea.value = url;
        textArea.style.position = "fixed";
        textArea.style.opacity = "0";
        document.body.appendChild(textArea);
        textArea.select();
        try {
          document.execCommand("copy");
          if (copyBtn) {
            copyBtn.classList.add("copied");
            setTimeout(function () {
              copyBtn.classList.remove("copied");
            }, 2000);
          }
        } catch (e) {
          console.error("Fallback copy failed:", e);
        }
        document.body.removeChild(textArea);
      });
  },

  /**
   * Show QR modal and generate QR code for current URL
   * Recalculates formulas and updates URL before generating
   */
  async showQRModal() {
    const modal = document.getElementById("qr-modal");
    const container = document.getElementById("qrcode-container");
    const warningText = document.getElementById("qr-warning");

    if (!modal || !container) return;

    if (typeof QRCode !== "function") {
      alert("QR Generator module not loaded.");
      return;
    }

    // Ensure formulas are fresh and the latest state is encoded into the URL
    recalculateFormulas();
    try {
      await updateURL();
    } catch (err) {
      console.error("Failed to sync data before generating QR", err);
      alert("Could not prepare data for QR code.");
      return;
    }

    const currentUrl = window.location.href;

    container.innerHTML = "";
    if (warningText) {
      warningText.textContent = "";
      warningText.classList.add("hidden");
    }

    if (currentUrl.length > MAX_QR_URL_LENGTH) {
      if (warningText) {
        warningText.textContent = `Spreadsheet too large for QR transfer (${currentUrl.length} characters). Remove some data and try again.`;
        warningText.classList.remove("hidden");
      }
      modal.classList.remove("hidden");
      return;
    }

    try {
      new QRCode(container, {
        text: currentUrl,
        width: 240,
        height: 240,
        colorDark: "#000000",
        colorLight: "#ffffff",
        correctLevel: QRCode.CorrectLevel.M,
      });
      modal.classList.remove("hidden");
    } catch (e) {
      console.error("QR Library missing or error", e);
      alert("Could not generate QR code.");
    }
  },

  /**
   * Hide QR modal and clear QR code
   */
  hideQRModal() {
    const modal = document.getElementById("qr-modal");
    const container = document.getElementById("qrcode-container");
    const warningText = document.getElementById("qr-warning");

    if (modal) {
      modal.classList.add("hidden");
    }
    if (container) {
      container.innerHTML = "";
    }
    if (warningText) {
      warningText.textContent = "";
      warningText.classList.add("hidden");
    }
  },
};
