import { ALLOWED_SPAN_STYLES, ALLOWED_TAGS } from "./constants.js";

// Validate CSS color values to prevent CSS injection
export function isValidCSSColor(color) {
  if (color === null || color === undefined) return false;
  if (typeof color !== "string") return false;

  // Allow empty string (to clear color)
  if (color === "") return true;

  // Validate hex colors (#RGB, #RRGGBB, #RRGGBBAA)
  if (/^#([0-9A-Fa-f]{3}|[0-9A-Fa-f]{6}|[0-9A-Fa-f]{8})$/.test(color)) {
    return true;
  }

  // Validate rgb/rgba with proper bounds
  if (/^rgba?\(\s*\d{1,3}\s*,\s*\d{1,3}\s*,\s*\d{1,3}\s*(,\s*(0|1|0?\.\d+)\s*)?\)$/.test(color)) {
    return true;
  }

  // Reject anything else (prevents CSS injection)
  return false;
}

// Escape HTML entities for safe display (converts HTML to plain text display)
export function escapeHTML(str) {
  if (!str || typeof str !== "string") return "";
  return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#x27;");
}

// Filter safe CSS styles for span elements
export function filterSafeStyles(styleString) {
  if (!styleString) return "";
  const safeStyles = [];
  const parts = styleString.split(";");
  for (const part of parts) {
    const colonIndex = part.indexOf(":");
    if (colonIndex === -1) continue;
    const prop = part.substring(0, colonIndex).trim().toLowerCase();
    if (ALLOWED_SPAN_STYLES.includes(prop)) {
      safeStyles.push(part.trim());
    }
  }
  return safeStyles.join("; ");
}

// Defense-in-depth: Check for dangerous patterns that might bypass sanitization
// Returns true if content appears safe, false if dangerous patterns detected
export function isContentSafe(html) {
  if (!html || typeof html !== "string") return true;

  // Dangerous patterns to reject (case-insensitive)
  const dangerousPatterns = [
    /<script/i, // Script tags
    /javascript:/i, // JavaScript protocol
    /on\w+\s*=/i, // Event handlers (onclick, onerror, etc.)
    /data:\s*text\/html/i, // Data URLs with HTML
    /<iframe/i, // Iframes
    /<object/i, // Object embeds
    /<embed/i, // Embed tags
    /<link/i, // Link tags (can load external resources)
    /<meta/i, // Meta tags (can redirect)
    /<base/i, // Base tag (can change URL resolution)
    /expression\s*\(/i, // CSS expressions (IE)
    /url\s*\(\s*["']?\s*javascript:/i, // JavaScript in CSS url()
  ];

  for (const pattern of dangerousPatterns) {
    if (pattern.test(html)) {
      console.warn("Blocked dangerous content pattern:", pattern.toString());
      return false;
    }
  }
  return true;
}

// Sanitize HTML using DOMParser (does NOT execute scripts/event handlers)
export function sanitizeHTML(html) {
  if (!html || typeof html !== "string") return "";

  // Defense-in-depth: Pre-check for obviously dangerous patterns
  if (!isContentSafe(html)) {
    // Return escaped version instead of potentially dangerous content
    return escapeHTML(html);
  }

  // Use DOMParser - it does NOT execute scripts or event handlers
  const parser = new DOMParser();
  const doc = parser.parseFromString("<body>" + html + "</body>", "text/html");

  function sanitizeNode(node) {
    const childNodes = Array.from(node.childNodes);

    for (const child of childNodes) {
      if (child.nodeType === Node.TEXT_NODE) {
        continue;
      }

      if (child.nodeType === Node.ELEMENT_NODE) {
        const tagName = child.tagName.toUpperCase();

        if (!ALLOWED_TAGS.includes(tagName)) {
          // Replace disallowed tags with their text content
          const textNode = document.createTextNode(child.textContent || "");
          node.replaceChild(textNode, child);
        } else {
          // Remove all attributes except safe styles on SPAN
          const attrs = Array.from(child.attributes);
          for (const attr of attrs) {
            if (tagName === "SPAN" && attr.name === "style") {
              const safeStyle = filterSafeStyles(attr.value);
              if (safeStyle) {
                child.setAttribute("style", safeStyle);
              } else {
                child.removeAttribute("style");
              }
            } else {
              child.removeAttribute(attr.name);
            }
          }
          sanitizeNode(child);
        }
      } else {
        // Remove comments and other node types
        node.removeChild(child);
      }
    }
  }

  sanitizeNode(doc.body);
  const result = doc.body.innerHTML;

  // Defense-in-depth: Final verification of sanitized output
  if (!isContentSafe(result)) {
    console.warn("Sanitized output still contains dangerous patterns, escaping");
    return escapeHTML(html);
  }

  return result;
}
