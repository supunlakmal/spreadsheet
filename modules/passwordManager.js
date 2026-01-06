/**
 * Password Manager Module
 * Handles all password-related state, UI, and interactions.
 */
export const PasswordManager = {
  // State
  currentPassword: null,
  pendingEncryptedData: null,
  modalMode: "set",

  // Callbacks provided by main script
  callbacks: {
    decryptAndDecode: async () => null,
    onDecryptSuccess: () => {},
    updateURL: () => {},
    showToast: () => {},
    validateState: () => {},
  },

  /**
   * Initialize the password manager
   * @param {Object} callbacks - Functions required for integration
   */
  init(callbacks) {
    this.callbacks = { ...this.callbacks, ...callbacks };
    this.attachEventListeners();
    this.updateLockButtonUI();
  },

  /**
   * Get the current password
   * @returns {string|null}
   */
  getPassword() {
    return this.currentPassword;
  },

  /**
   * Set the current password (e.g. after successful decryption or setting new)
   * @param {string|null} password 
   */
  setPassword(password) {
    this.currentPassword = password;
    this.updateLockButtonUI();
  },

  /**
   * Handle encrypted data found during load
   * @param {string} encryptedData 
   */
  handleEncryptedData(encryptedData) {
    this.pendingEncryptedData = encryptedData;
    // Show password modal after a short delay to let UI render
    setTimeout(() => this.showPasswordModal("decrypt"), 100);
  },

  // ========== UI Functions ==========

  showPasswordModal(mode) {
    this.modalMode = mode;
    const modal = document.getElementById("password-modal");
    const title = document.getElementById("modal-title");
    const description = document.getElementById("modal-description");
    const confirmInput = document.getElementById("password-confirm");
    const submitBtn = document.getElementById("modal-submit");
    const passwordInput = document.getElementById("password-input");
    const errorEl = document.getElementById("modal-error");

    if (!modal) return;

    // Reset form
    if (passwordInput) passwordInput.value = "";
    if (confirmInput) confirmInput.value = "";
    if (errorEl) {
      errorEl.classList.add("hidden");
      errorEl.textContent = "";
    }

    if (mode === "set") {
      if (title) title.textContent = "Set Password";
      if (description) description.textContent = "Enter a password to encrypt this spreadsheet. Anyone with the link will need this password to view it.";
      if (confirmInput) {
        confirmInput.style.display = "";
        confirmInput.placeholder = "Confirm password";
      }
      if (submitBtn) submitBtn.textContent = "Set Password";
    } else if (mode === "decrypt") {
      if (title) title.textContent = "Enter Password";
      if (description) description.textContent = "This spreadsheet is password-protected. Enter the password to view it.";
      if (confirmInput) confirmInput.style.display = "none";
      if (submitBtn) submitBtn.textContent = "Unlock";
    } else if (mode === "remove") {
      if (title) title.textContent = "Remove Password";
      if (description) description.textContent = "Enter the current password to remove encryption from this spreadsheet.";
      if (confirmInput) confirmInput.style.display = "none";
      if (submitBtn) submitBtn.textContent = "Remove Password";
    }

    modal.classList.remove("hidden");
    if (passwordInput) passwordInput.focus();
  },

  hidePasswordModal() {
    const modal = document.getElementById("password-modal");
    if (modal) modal.classList.add("hidden");
    // Clear pending data if we cancelled decryption (though usually we'd want to keep it if they just closed modal?)
    // Original logic cleared it:
    // this.pendingEncryptedData = null; 
    // But if we cancel decryption, we probably can't view the file anyway.
  },

  showModalError(message) {
    const errorEl = document.getElementById("modal-error");
    if (errorEl) {
      errorEl.textContent = message;
      errorEl.classList.remove("hidden");
    }
  },

  updateLockButtonUI() {
    const lockBtn = document.getElementById("lock-btn");
    if (!lockBtn) return;

    const icon = lockBtn.querySelector("i");
    if (this.currentPassword) {
      lockBtn.classList.add("locked");
      lockBtn.title = "Remove Password Protection";
      if (icon) icon.className = "fa-solid fa-lock";
    } else {
      lockBtn.classList.remove("locked");
      lockBtn.title = "Password Protection";
      if (icon) icon.className = "fa-solid fa-lock-open";
    }
  },

  handleLockButtonClick() {
    if (this.currentPassword) {
      // Already encrypted - offer to remove password
      this.showPasswordModal("remove");
    } else {
      // Not encrypted - set password
      this.showPasswordModal("set");
    }
  },

  async handleModalSubmit() {
    const passwordInput = document.getElementById("password-input");
    const confirmInput = document.getElementById("password-confirm");
    
    if (!passwordInput) return;
    const password = passwordInput.value;

    if (!password) {
      this.showModalError("Please enter a password.");
      return;
    }

    if (this.modalMode === "set") {
      // Setting new password
      const confirm = confirmInput ? confirmInput.value : "";
      if (password !== confirm) {
        this.showModalError("Passwords do not match.");
        return;
      }
      if (password.length < 4) {
        this.showModalError("Password must be at least 4 characters.");
        return;
      }

      this.currentPassword = password;
      this.updateLockButtonUI();
      this.hidePasswordModal();
      // Re-encode state with encryption
      this.callbacks.updateURL();
      this.callbacks.showToast("Password protection enabled", "success");

    } else if (this.modalMode === "decrypt") {
      // Decrypting loaded data
      if (!this.pendingEncryptedData) {
        this.showModalError("No encrypted data to decrypt.");
        return;
      }

      try {
        const rawState = await this.callbacks.decryptAndDecode(this.pendingEncryptedData, password);
        if (!rawState) {
          throw new Error("Invalid decrypted data");
        }

        // Validate state
        const validatedState = this.callbacks.validateState(rawState);
        if (!validatedState) {
          throw new Error("Validation failed");
        }

        // Success
        this.currentPassword = password;
        this.pendingEncryptedData = null; // Clear it
        this.callbacks.onDecryptSuccess(validatedState);
        
        this.hidePasswordModal();
        this.updateLockButtonUI();
        this.callbacks.showToast("Spreadsheet unlocked", "success");

      } catch (e) {
        console.error("Decryption failed:", e);
        this.showModalError("Incorrect password.");
      }

    } else if (this.modalMode === "remove") {
      // For now, we trust they have the password if they are here (or we could verify it matches current)
      // Note: Original code didn't strictly verify against currentPassword (it just checked if you entered *something*?)
      // Wait, original: `if (modalMode === "remove") { ... currentPassword = null ... }`
      // It didn't verify the typed password matches the active `currentPassword`. 
      // It prompted "Enter the current password", but didn't actually check it against `currentPassword`.
      // The assumption is if you are viewing the file, you know the password (or it was just unlocked).
      // However, for better security UX, we might want to check, but let's stick to original behavior to minimize regression risk.
      // But wait, if I type "wrong" password to remove, it shouldn't work? 
      // Actually, since `currentPassword` is in memory, we CAN check it.
      // If `currentPassword` is set, we SHOULD check.
      
      if (this.currentPassword && password !== this.currentPassword) {
         this.showModalError("Incorrect password.");
         return;
      }

      this.currentPassword = null;
      this.updateLockButtonUI();
      this.hidePasswordModal();
      // Re-encode state without encryption
      this.callbacks.updateURL();
      this.callbacks.showToast("Password protection removed", "success");
    }
  },

  attachEventListeners() {
    const lockBtn = document.getElementById("lock-btn");
    const modalCancel = document.getElementById("modal-cancel");
    const modalSubmit = document.getElementById("modal-submit");
    const modalBackdrop = document.querySelector(".modal-backdrop");
    const passwordInput = document.getElementById("password-input");
    const passwordConfirm = document.getElementById("password-confirm");

    if (lockBtn) {
      // Use logical wrapper
      lockBtn.addEventListener("click", () => this.handleLockButtonClick());
    }
    if (modalCancel) {
      modalCancel.addEventListener("click", () => this.hidePasswordModal());
    }
    if (modalSubmit) {
      modalSubmit.addEventListener("click", () => this.handleModalSubmit());
    }
    if (modalBackdrop) {
      modalBackdrop.addEventListener("click", () => {
        // Only allow closing if not in decrypt mode
        if (this.modalMode !== "decrypt") {
          this.hidePasswordModal();
        }
      });
    }

    if (passwordInput) {
      passwordInput.addEventListener("keydown", (e) => {
        if (e.key === "Enter") {
          e.preventDefault();
          // If confirm is visible and empty, focus it; otherwise submit
          if (passwordConfirm && passwordConfirm.style.display !== "none" && !passwordConfirm.value) {
            passwordConfirm.focus();
          } else {
            this.handleModalSubmit();
          }
        }
      });
    }

    if (passwordConfirm) {
      passwordConfirm.addEventListener("keydown", (e) => {
        if (e.key === "Enter") {
          e.preventDefault();
          this.handleModalSubmit();
        }
      });
    }
  }
};
