/**
 * themeManager.js
 * Manages theme (dark/light mode) functionality
 */

let debouncedUpdateURL;
let scheduleFullSync;

export const ThemeManager = {
  /**
   * Initialize the theme manager with required callbacks
   * @param {Object} callbacks - Required callbacks from script.js
   * @param {Function} callbacks.debouncedUpdateURL - Debounced URL update function
   * @param {Function} callbacks.scheduleFullSync - Schedule full P2P sync function
   */
  init(callbacks) {
    debouncedUpdateURL = callbacks.debouncedUpdateURL;
    scheduleFullSync = callbacks.scheduleFullSync;
  },

  /**
   * Check if dark mode is currently active
   * @returns {boolean} True if dark mode is active
   */
  isDarkMode() {
    return document.body.classList.contains("dark-mode");
  },

  /**
   * Apply a specific theme to the document body
   * @param {string} theme - Theme to apply ("dark" or "light")
   */
  applyTheme(theme) {
    if (theme === "dark") {
      document.body.classList.add("dark-mode");
    } else if (theme === "light") {
      document.body.classList.remove("dark-mode");
    }
    // Save to localStorage as well
    try {
      localStorage.setItem("spreadsheet-theme", theme);
    } catch (e) {
      // localStorage not available
    }
  },

  /**
   * Toggle between dark and light mode
   * Updates localStorage, URL state, and syncs with P2P peers
   */
  toggleTheme() {
    const body = document.body;
    const isDark = body.classList.toggle("dark-mode");
    const theme = isDark ? "dark" : "light";

    // Save preference to localStorage
    try {
      localStorage.setItem("spreadsheet-theme", theme);
    } catch (e) {
      // localStorage not available
    }

    // Update URL with new theme
    debouncedUpdateURL();
    scheduleFullSync();
  },

  /**
   * Load saved theme preference from localStorage or system preference
   * Called on app initialization
   */
  loadTheme() {
    try {
      const savedTheme = localStorage.getItem("spreadsheet-theme");
      if (savedTheme === "dark") {
        document.body.classList.add("dark-mode");
      } else if (savedTheme === "light") {
        document.body.classList.remove("dark-mode");
      } else {
        // Check system preference if no saved preference
        if (window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches) {
          document.body.classList.add("dark-mode");
        }
      }
    } catch (e) {
      // localStorage not available
    }
  },
};
