import { TOAST_DURATION, TOAST_ICONS } from "./constants.js";

/**
 * Dismiss a toast with exit animation
 * @param {HTMLElement} toast - The toast element to dismiss
 */
export function dismissToast(toast) {
  if (!toast || toast.classList.contains("toast-exit")) return;

  toast.classList.add("toast-exit");
  toast.addEventListener(
    "animationend",
    () => {
      toast.remove();
    },
    { once: true }
  );
}

/**
 * Show a toast notification
 * @param {string} message - The message to display
 * @param {string} type - Type: 'success', 'error', 'warning', 'info'
 * @param {number} duration - Duration in ms (0 for no auto-dismiss)
 */
export function showToast(message, type = "info", duration = TOAST_DURATION) {
  const container = document.getElementById("toast-container");
  if (!container) return;

  // Create toast element
  const toast = document.createElement("div");
  toast.className = `toast toast-${type}`;
  toast.setAttribute("role", "alert");

  // Icon
  const icon = document.createElement("div");
  icon.className = "toast-icon";
  icon.innerHTML = `<i class="fa-solid ${TOAST_ICONS[type] || TOAST_ICONS.info}"></i>`;
  toast.appendChild(icon);

  // Message
  const msg = document.createElement("div");
  msg.className = "toast-message";
  msg.textContent = message;
  toast.appendChild(msg);

  // Close button
  const closeBtn = document.createElement("button");
  closeBtn.className = "toast-close";
  closeBtn.innerHTML = '<i class="fa-solid fa-xmark"></i>';
  closeBtn.setAttribute("aria-label", "Dismiss");
  closeBtn.addEventListener("click", () => dismissToast(toast));
  toast.appendChild(closeBtn);

  // Progress bar (if auto-dismiss)
  if (duration > 0) {
    const progress = document.createElement("div");
    progress.className = "toast-progress";
    progress.style.animationDuration = `${duration}ms`;
    toast.appendChild(progress);

    // Auto-dismiss after duration
    setTimeout(() => dismissToast(toast), duration);
  }

  // Add to container
  container.appendChild(toast);

  // Limit max toasts
  const toasts = container.querySelectorAll(".toast:not(.toast-exit)");
  if (toasts.length > 5) {
    dismissToast(toasts[0]);
  }

  return toast;
}
