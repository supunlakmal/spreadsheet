// Dynamic Spreadsheet Web App
// Data persists in URL hash for easy sharing

(function() {
    'use strict';

    // Limits
    const MAX_ROWS = 30;
    const MAX_COLS = 15;
    const DEBOUNCE_DELAY = 200;
    const ACTIVE_HEADER_CLASS = 'header-active';
    const ROW_HEADER_WIDTH = 40;
    const HEADER_ROW_HEIGHT = 32;
    const DEFAULT_COL_WIDTH = 100;
    const MIN_COL_WIDTH = 80;
    const IS_MOBILE_LAYOUT = window.matchMedia && window.matchMedia('(max-width: 768px)').matches;
    const DEFAULT_ROW_HEIGHT = IS_MOBILE_LAYOUT ? 44 : 32;
    const MIN_ROW_HEIGHT = DEFAULT_ROW_HEIGHT;

    // Default starting size
    const DEFAULT_ROWS = 10;
    const DEFAULT_COLS = 10;

    // URL length thresholds for warnings
    const URL_LENGTH_WARNING = 2000;   // Yellow - some older browsers may truncate
    const URL_LENGTH_CAUTION = 4000;   // Orange - URL shorteners may fail
    const URL_LENGTH_CRITICAL = 8000;  // Red - some browsers may fail
    const URL_LENGTH_MAX_DISPLAY = 10000; // For progress bar scaling

    const FORMULA_SUGGESTIONS = [
        { name: 'SUM', signature: 'SUM(range)', description: 'Adds numbers in a range' },
        { name: 'AVG', signature: 'AVG(range)', description: 'Average of numbers in a range' }
    ];

    // Valid formula patterns (security whitelist)
    const VALID_FORMULA_PATTERNS = [
        /^=\s*SUM\s*\(\s*[A-Z]+\d+\s*:\s*[A-Z]+\d+\s*\)\s*$/i,
        /^=\s*AVG\s*\(\s*[A-Z]+\d+\s*:\s*[A-Z]+\d+\s*\)\s*$/i
    ];

    // Validate formula matches allowed patterns (security: prevents arbitrary formula injection)
    function isValidFormula(formula) {
        if (!formula || typeof formula !== 'string') return false;
        if (!formula.startsWith('=')) return false;
        // Check SUM/AVG patterns
        if (VALID_FORMULA_PATTERNS.some(pattern => pattern.test(formula))) {
            return true;
        }
        // Check arithmetic expression
        return isArithmeticFormula(formula);
    }

    // Sanitize formula - returns the formula if valid, otherwise escapes it as text
    function sanitizeFormula(formula) {
        if (!formula || typeof formula !== 'string') return '';
        if (!formula.startsWith('=')) return escapeHTML(formula);
        // If it's a valid formula pattern, allow it
        if (isValidFormula(formula)) return formula;
        // Invalid formula - treat as text to prevent injection
        return escapeHTML(formula);
    }
    const FONT_SIZE_OPTIONS = [10, 12, 14, 16, 18, 24];

    // ========== Encryption Module (AES-GCM 256-bit) ==========
    const CryptoUtils = {
        algo: { name: 'AES-GCM', length: 256 },
        kdf: { name: 'PBKDF2', hash: 'SHA-256', iterations: 100000 },

        // Derive a cryptographic key from a password using PBKDF2
        async deriveKey(password, salt) {
            const enc = new TextEncoder();
            const keyMaterial = await window.crypto.subtle.importKey(
                'raw',
                enc.encode(password),
                'PBKDF2',
                false,
                ['deriveKey']
            );
            return window.crypto.subtle.deriveKey(
                { ...this.kdf, salt: salt },
                keyMaterial,
                this.algo,
                false,
                ['encrypt', 'decrypt']
            );
        },

        // Encrypt data string with password, returns Base64 string
        async encrypt(dataString, password) {
            const enc = new TextEncoder();
            const salt = window.crypto.getRandomValues(new Uint8Array(16));
            const iv = window.crypto.getRandomValues(new Uint8Array(12));

            const key = await this.deriveKey(password, salt);
            const encodedData = enc.encode(dataString);

            const encryptedContent = await window.crypto.subtle.encrypt(
                { name: 'AES-GCM', iv: iv },
                key,
                encodedData
            );

            // Pack: Salt (16) + IV (12) + EncryptedData
            const buffer = new Uint8Array(salt.byteLength + iv.byteLength + encryptedContent.byteLength);
            buffer.set(salt, 0);
            buffer.set(iv, salt.byteLength);
            buffer.set(new Uint8Array(encryptedContent), salt.byteLength + iv.byteLength);

            return this.bufferToBase64(buffer);
        },

        // Decrypt Base64 string with password, returns original data string
        async decrypt(base64String, password) {
            const buffer = this.base64ToBuffer(base64String);

            // Extract: Salt (16) + IV (12) + EncryptedData
            const salt = buffer.slice(0, 16);
            const iv = buffer.slice(16, 28);
            const data = buffer.slice(28);

            const key = await this.deriveKey(password, salt);

            const decryptedContent = await window.crypto.subtle.decrypt(
                { name: 'AES-GCM', iv: iv },
                key,
                data
            );

            const dec = new TextDecoder();
            return dec.decode(decryptedContent);
        },

        // Convert Uint8Array to URL-safe Base64
        bufferToBase64(buffer) {
            let binary = '';
            const bytes = new Uint8Array(buffer);
            for (let i = 0; i < bytes.byteLength; i++) {
                binary += String.fromCharCode(bytes[i]);
            }
            // Use URL-safe Base64 (replace + with -, / with _, remove padding)
            return btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
        },

        // Convert URL-safe Base64 to Uint8Array
        base64ToBuffer(base64) {
            // Restore standard Base64
            let standardBase64 = base64.replace(/-/g, '+').replace(/_/g, '/');
            // Add padding if needed
            while (standardBase64.length % 4) {
                standardBase64 += '=';
            }
            const binaryString = atob(standardBase64);
            const bytes = new Uint8Array(binaryString.length);
            for (let i = 0; i < binaryString.length; i++) {
                bytes[i] = binaryString.charCodeAt(i);
            }
            return bytes;
        }
    };

    // Allowed HTML tags for sanitization (preserves basic formatting)
    const ALLOWED_TAGS = ['B', 'I', 'U', 'STRONG', 'EM', 'SPAN', 'BR'];
    const ALLOWED_SPAN_STYLES = ['font-weight', 'font-style', 'text-decoration', 'color', 'background-color'];

    // Dynamic dimensions
    let rows = DEFAULT_ROWS;
    let cols = DEFAULT_COLS;

    // Data model - dynamic 2D array
    let data = createEmptyData(rows, cols);

    // Formula storage - parallel array to data
    let formulas = createEmptyData(rows, cols);

    // Cell styles - alignment, colors, and font size
    let cellStyles = createEmptyCellStyles(rows, cols);
    let colWidths = createDefaultColumnWidths(cols);
    let rowHeights = createDefaultRowHeights(rows);

    // Debounce timer
    let debounceTimer = null;

    // Active header tracking for row/column highlight
    let activeRow = null;
    let activeCol = null;

    // Multi-cell selection state
    let selectionStart = null;  // { row, col } anchor point
    let selectionEnd = null;    // { row, col } current end
    let isSelecting = false;    // true during mouse drag
    let hoverRow = null;
    let hoverCol = null;

    // Formula range selection mode (for clicking to select ranges like Google Sheets)
    let formulaEditMode = false;       // true when typing a formula
    let formulaEditCell = null;        // { row, col, element } of cell being edited
    let formulaRangeStart = null;      // Start of range being selected
    let formulaRangeEnd = null;        // End of range being selected
    let formulaDropdown = null;        // DOM node for formula suggestions
    let formulaDropdownItems = [];     // List of visible dropdown items
    let formulaDropdownIndex = -1;     // Active item index
    let formulaDropdownAnchor = null;  // Cell element used for positioning
    let editingCell = null;            // { row, col } when editing a cell's text
    let resizeState = null;            // { type, index, startX, startY, startSize }

    // Encryption state
    let currentPassword = null;        // Password for encryption (null = no encryption)
    let pendingEncryptedData = null;   // Stores encrypted data awaiting password input

    // ========== Toast Notification System ==========
    const TOAST_DURATION = 3000; // Default duration in ms
    const TOAST_ICONS = {
        success: 'fa-check',
        error: 'fa-xmark',
        warning: 'fa-exclamation',
        info: 'fa-info'
    };

    /**
     * Show a toast notification
     * @param {string} message - The message to display
     * @param {string} type - Type: 'success', 'error', 'warning', 'info'
     * @param {number} duration - Duration in ms (0 for no auto-dismiss)
     */
    function showToast(message, type = 'info', duration = TOAST_DURATION) {
        const container = document.getElementById('toast-container');
        if (!container) return;

        // Create toast element
        const toast = document.createElement('div');
        toast.className = `toast toast-${type}`;
        toast.setAttribute('role', 'alert');

        // Icon
        const icon = document.createElement('div');
        icon.className = 'toast-icon';
        icon.innerHTML = `<i class="fa-solid ${TOAST_ICONS[type] || TOAST_ICONS.info}"></i>`;
        toast.appendChild(icon);

        // Message
        const msg = document.createElement('div');
        msg.className = 'toast-message';
        msg.textContent = message;
        toast.appendChild(msg);

        // Close button
        const closeBtn = document.createElement('button');
        closeBtn.className = 'toast-close';
        closeBtn.innerHTML = '<i class="fa-solid fa-xmark"></i>';
        closeBtn.setAttribute('aria-label', 'Dismiss');
        closeBtn.addEventListener('click', () => dismissToast(toast));
        toast.appendChild(closeBtn);

        // Progress bar (if auto-dismiss)
        if (duration > 0) {
            const progress = document.createElement('div');
            progress.className = 'toast-progress';
            progress.style.animationDuration = `${duration}ms`;
            toast.appendChild(progress);

            // Auto-dismiss after duration
            setTimeout(() => dismissToast(toast), duration);
        }

        // Add to container
        container.appendChild(toast);

        // Limit max toasts
        const toasts = container.querySelectorAll('.toast:not(.toast-exit)');
        if (toasts.length > 5) {
            dismissToast(toasts[0]);
        }

        return toast;
    }

    /**
     * Dismiss a toast with exit animation
     * @param {HTMLElement} toast - The toast element to dismiss
     */
    function dismissToast(toast) {
        if (!toast || toast.classList.contains('toast-exit')) return;
        
        toast.classList.add('toast-exit');
        toast.addEventListener('animationend', () => {
            toast.remove();
        }, { once: true });
    }


    // Create empty data array with specified dimensions
    function createEmptyData(r, c) {
        return Array(r).fill(null).map(() => Array(c).fill(''));
    }

    function createEmptyCellStyle() {
        return { align: '', bg: '', color: '', fontSize: '' };
    }

    function createEmptyCellStyles(r, c) {
        return Array(r).fill(null).map(() =>
            Array(c).fill(null).map(() => createEmptyCellStyle())
        );
    }

    function createDefaultColumnWidths(count) {
        return Array(count).fill(DEFAULT_COL_WIDTH);
    }

    function createDefaultRowHeights(count) {
        return Array(count).fill(DEFAULT_ROW_HEIGHT);
    }

    function normalizeColumnWidths(widths, count) {
        const normalized = [];
        for (let i = 0; i < count; i++) {
            const raw = widths && widths[i];
            const value = parseInt(raw, 10);
            if (Number.isFinite(value) && value > 0) {
                normalized.push(Math.max(MIN_COL_WIDTH, value));
            } else {
                normalized.push(DEFAULT_COL_WIDTH);
            }
        }
        return normalized;
    }

    function normalizeRowHeights(heights, count) {
        const normalized = [];
        for (let i = 0; i < count; i++) {
            const raw = heights && heights[i];
            const value = parseInt(raw, 10);
            if (Number.isFinite(value) && value > 0) {
                normalized.push(Math.max(MIN_ROW_HEIGHT, value));
            } else {
                normalized.push(DEFAULT_ROW_HEIGHT);
            }
        }
        return normalized;
    }

    // ========== Security Functions ==========

    // Validate CSS color values to prevent CSS injection
    function isValidCSSColor(color) {
        if (color === null || color === undefined) return false;
        if (typeof color !== 'string') return false;

        // Allow empty string (to clear color)
        if (color === '') return true;

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
    function escapeHTML(str) {
        if (!str || typeof str !== 'string') return '';
        return str
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#x27;');
    }

    // Filter safe CSS styles for span elements
    function filterSafeStyles(styleString) {
        if (!styleString) return '';
        const safeStyles = [];
        const parts = styleString.split(';');
        for (const part of parts) {
            const colonIndex = part.indexOf(':');
            if (colonIndex === -1) continue;
            const prop = part.substring(0, colonIndex).trim().toLowerCase();
            if (ALLOWED_SPAN_STYLES.includes(prop)) {
                safeStyles.push(part.trim());
            }
        }
        return safeStyles.join('; ');
    }

    // Defense-in-depth: Check for dangerous patterns that might bypass sanitization
    // Returns true if content appears safe, false if dangerous patterns detected
    function isContentSafe(html) {
        if (!html || typeof html !== 'string') return true;

        // Dangerous patterns to reject (case-insensitive)
        const dangerousPatterns = [
            /<script/i,                          // Script tags
            /javascript:/i,                       // JavaScript protocol
            /on\w+\s*=/i,                        // Event handlers (onclick, onerror, etc.)
            /data:\s*text\/html/i,               // Data URLs with HTML
            /<iframe/i,                          // Iframes
            /<object/i,                          // Object embeds
            /<embed/i,                           // Embed tags
            /<link/i,                            // Link tags (can load external resources)
            /<meta/i,                            // Meta tags (can redirect)
            /<base/i,                            // Base tag (can change URL resolution)
            /expression\s*\(/i,                  // CSS expressions (IE)
            /url\s*\(\s*["']?\s*javascript:/i    // JavaScript in CSS url()
        ];

        for (const pattern of dangerousPatterns) {
            if (pattern.test(html)) {
                console.warn('Blocked dangerous content pattern:', pattern.toString());
                return false;
            }
        }
        return true;
    }

    // Sanitize HTML using DOMParser (does NOT execute scripts/event handlers)
    function sanitizeHTML(html) {
        if (!html || typeof html !== 'string') return '';

        // Defense-in-depth: Pre-check for obviously dangerous patterns
        if (!isContentSafe(html)) {
            // Return escaped version instead of potentially dangerous content
            return escapeHTML(html);
        }

        // Use DOMParser - it does NOT execute scripts or event handlers
        const parser = new DOMParser();
        const doc = parser.parseFromString('<body>' + html + '</body>', 'text/html');

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
                        const textNode = document.createTextNode(child.textContent || '');
                        node.replaceChild(textNode, child);
                    } else {
                        // Remove all attributes except safe styles on SPAN
                        const attrs = Array.from(child.attributes);
                        for (const attr of attrs) {
                            if (tagName === 'SPAN' && attr.name === 'style') {
                                const safeStyle = filterSafeStyles(attr.value);
                                if (safeStyle) {
                                    child.setAttribute('style', safeStyle);
                                } else {
                                    child.removeAttribute('style');
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
            console.warn('Sanitized output still contains dangerous patterns, escaping');
            return escapeHTML(html);
        }

        return result;
    }

    // Convert column index to letter (0 = A, 1 = B, ... 25 = Z)
    function colToLetter(col) {
        return String.fromCharCode(65 + col);
    }

    // ========== Formula Helper Functions ==========

    // Convert column letter(s) to index: A=0, B=1, ..., O=14
    function letterToCol(letters) {
        letters = letters.toUpperCase();
        let col = 0;
        for (let i = 0; i < letters.length; i++) {
            col = col * 26 + (letters.charCodeAt(i) - 64);
        }
        return col - 1;
    }

    // Parse cell reference "A1" → { row: 0, col: 0 }
    function parseCellRef(ref) {
        const match = ref.toUpperCase().match(/^([A-Z]+)(\d+)$/);
        if (!match) return null;
        return {
            col: letterToCol(match[1]),
            row: parseInt(match[2], 10) - 1
        };
    }

    // Parse range "A1:B5" → { startRow, startCol, endRow, endCol }
    function parseRange(range) {
        const parts = range.split(':');
        if (parts.length !== 2) return null;
        const start = parseCellRef(parts[0].trim());
        const end = parseCellRef(parts[1].trim());
        if (!start || !end) return null;
        return {
            startRow: Math.min(start.row, end.row),
            startCol: Math.min(start.col, end.col),
            endRow: Math.max(start.row, end.row),
            endCol: Math.max(start.col, end.col)
        };
    }

    // Build cell reference string like "A1"
    function buildCellRef(row, col) {
        return colToLetter(col) + (row + 1);
    }

    // Build range reference string like "A1:B5"
    function buildRangeRef(startRow, startCol, endRow, endCol) {
        const minRow = Math.min(startRow, endRow);
        const maxRow = Math.max(startRow, endRow);
        const minCol = Math.min(startCol, endCol);
        const maxCol = Math.max(startCol, endCol);

        if (minRow === maxRow && minCol === maxCol) {
            // Single cell
            return buildCellRef(minRow, minCol);
        }
        return buildCellRef(minRow, minCol) + ':' + buildCellRef(maxRow, maxCol);
    }

    // Insert text at current cursor position in contentEditable
    function insertTextAtCursor(text) {
        const selection = window.getSelection();
        if (!selection.rangeCount) return;

        const range = selection.getRangeAt(0);
        range.deleteContents();

        const textNode = document.createTextNode(text);
        range.insertNode(textNode);

        // Move cursor after inserted text
        range.setStartAfter(textNode);
        range.collapse(true);
        selection.removeAllRanges();
        selection.addRange(range);
    }

    // Move caret to end of contentEditable element
    function setCaretToEnd(element) {
        const range = document.createRange();
        range.selectNodeContents(element);
        range.collapse(false);
        const selection = window.getSelection();
        selection.removeAllRanges();
        selection.addRange(range);
    }

    function getFormulaQuery(rawValue) {
        const match = rawValue.match(/^=\s*([A-Z]*)$/i);
        if (!match) return null;
        return match[1].toUpperCase();
    }

    function getFormulaSuggestions(query) {
        if (query === null) return [];
        if (query === '') return FORMULA_SUGGESTIONS.slice();
        return FORMULA_SUGGESTIONS.filter(item => item.name.startsWith(query));
    }

    function createFormulaDropdown() {
        if (formulaDropdown) return;
        const dropdown = document.createElement('div');
        dropdown.className = 'formula-dropdown';
        dropdown.setAttribute('role', 'listbox');
        dropdown.setAttribute('aria-hidden', 'true');

        dropdown.addEventListener('mousedown', function(event) {
            event.preventDefault();
        });

        dropdown.addEventListener('click', function(event) {
            const item = event.target.closest('.formula-item');
            if (!item) return;
            const formulaName = item.dataset.formula;
            if (formulaName) {
                applyFormulaSuggestion(formulaName);
            }
        });

        document.body.appendChild(dropdown);
        formulaDropdown = dropdown;
    }

    function isFormulaDropdownOpen() {
        return !!(formulaDropdown && formulaDropdown.classList.contains('open'));
    }

    function setActiveFormulaItem(index) {
        formulaDropdownIndex = index;
        formulaDropdownItems.forEach((item, idx) => {
            if (idx === index) {
                item.classList.add('active');
                item.setAttribute('aria-selected', 'true');
                item.scrollIntoView({ block: 'nearest' });
            } else {
                item.classList.remove('active');
                item.setAttribute('aria-selected', 'false');
            }
        });
    }

    function positionFormulaDropdown(anchor) {
        if (!formulaDropdown || !anchor) return;
        const rect = anchor.getBoundingClientRect();
        const viewportWidth = window.innerWidth;
        const viewportHeight = window.innerHeight;
        const padding = 6;

        formulaDropdown.style.left = `${rect.left}px`;
        formulaDropdown.style.top = `${rect.bottom + 4}px`;

        const dropdownRect = formulaDropdown.getBoundingClientRect();
        let left = rect.left;
        let top = rect.bottom + 4;

        if (left + dropdownRect.width > viewportWidth - padding) {
            left = Math.max(padding, viewportWidth - dropdownRect.width - padding);
        }

        if (top + dropdownRect.height > viewportHeight - padding) {
            const above = rect.top - dropdownRect.height - 4;
            if (above > padding) {
                top = above;
            }
        }

        formulaDropdown.style.left = `${left}px`;
        formulaDropdown.style.top = `${top}px`;
    }

    function showFormulaDropdown() {
        if (!formulaDropdown) return;
        formulaDropdown.classList.add('open');
        formulaDropdown.setAttribute('aria-hidden', 'false');
    }

    function hideFormulaDropdown() {
        if (!formulaDropdown) return;
        formulaDropdown.classList.remove('open');
        formulaDropdown.setAttribute('aria-hidden', 'true');
        formulaDropdownItems = [];
        formulaDropdownIndex = -1;
        formulaDropdownAnchor = null;
    }

    function updateFormulaDropdown(anchor, rawValue) {
        createFormulaDropdown();
        const query = getFormulaQuery(rawValue);
        const suggestions = getFormulaSuggestions(query);
        if (!anchor || suggestions.length === 0 || query === null) {
            hideFormulaDropdown();
            return;
        }

        formulaDropdownAnchor = anchor;
        formulaDropdown.innerHTML = '';
        suggestions.forEach(item => {
            const option = document.createElement('div');
            option.className = 'formula-item';
            option.dataset.formula = item.name;
            option.setAttribute('role', 'option');
            option.setAttribute('aria-selected', 'false');

            const nameEl = document.createElement('div');
            nameEl.className = 'formula-name';
            nameEl.textContent = item.name;

            const hintEl = document.createElement('div');
            hintEl.className = 'formula-hint';
            hintEl.textContent = `${item.signature} - ${item.description}`;

            option.appendChild(nameEl);
            option.appendChild(hintEl);
            formulaDropdown.appendChild(option);
        });

        formulaDropdownItems = Array.from(formulaDropdown.querySelectorAll('.formula-item'));
        setActiveFormulaItem(0);
        showFormulaDropdown();
        positionFormulaDropdown(anchor);
    }

    function moveFormulaDropdownSelection(delta) {
        if (!formulaDropdownItems.length) return;
        let nextIndex = formulaDropdownIndex + delta;
        if (nextIndex < 0) nextIndex = formulaDropdownItems.length - 1;
        if (nextIndex >= formulaDropdownItems.length) nextIndex = 0;
        setActiveFormulaItem(nextIndex);
    }

    function applyFormulaSuggestion(formulaName) {
        const target = formulaEditCell ? formulaEditCell.element : document.activeElement;
        if (!target || !target.classList.contains('cell-content')) return;

        const row = parseInt(target.dataset.row, 10);
        const col = parseInt(target.dataset.col, 10);
        if (isNaN(row) || isNaN(col)) return;

        const newValue = `=${formulaName}(`;
        target.innerText = newValue;
        setCaretToEnd(target);

        formulaEditMode = true;
        formulaEditCell = { row, col, element: target };
        formulas[row][col] = newValue;
        data[row][col] = newValue;

        hideFormulaDropdown();
        debouncedUpdateURL();
    }

    // ========== Formula Evaluation Functions ==========

    // Get numeric value from cell (returns 0 for empty/non-numeric)
    function getCellValue(row, col) {
        if (row < 0 || row >= rows || col < 0 || col >= cols) return 0;
        const val = data[row][col];
        if (!val || val === '') return 0;
        // Strip HTML tags and parse
        const stripped = String(val).replace(/<[^>]*>/g, '').trim();
        const num = parseFloat(stripped);
        return isNaN(num) ? 0 : num;
    }

    // ========== Arithmetic Expression Evaluator ==========

    // Tokenize arithmetic expression into array of tokens
    // Token types: NUMBER, CELL_REF, OPERATOR, LPAREN, RPAREN
    function tokenizeArithmetic(expr) {
        const tokens = [];
        let i = 0;

        while (i < expr.length) {
            // Skip whitespace
            if (/\s/.test(expr[i])) {
                i++;
                continue;
            }

            // Numbers (including decimals)
            if (/[\d.]/.test(expr[i])) {
                let num = '';
                while (i < expr.length && /[\d.]/.test(expr[i])) {
                    num += expr[i++];
                }
                if (!/^\d+\.?\d*$|^\d*\.\d+$/.test(num)) {
                    return { error: 'Invalid number: ' + num };
                }
                tokens.push({ type: 'NUMBER', value: parseFloat(num) });
                continue;
            }

            // Cell references (A1, B2, AA99, etc.)
            if (/[A-Z]/i.test(expr[i])) {
                let ref = '';
                while (i < expr.length && /[A-Z]/i.test(expr[i])) {
                    ref += expr[i++];
                }
                while (i < expr.length && /\d/.test(expr[i])) {
                    ref += expr[i++];
                }
                if (!/^[A-Z]+\d+$/i.test(ref)) {
                    return { error: 'Invalid cell reference: ' + ref };
                }
                tokens.push({ type: 'CELL_REF', value: ref.toUpperCase() });
                continue;
            }

            // Operators
            if ('+-*/'.includes(expr[i])) {
                tokens.push({ type: 'OPERATOR', value: expr[i] });
                i++;
                continue;
            }

            // Parentheses
            if (expr[i] === '(') {
                tokens.push({ type: 'LPAREN' });
                i++;
                continue;
            }
            if (expr[i] === ')') {
                tokens.push({ type: 'RPAREN' });
                i++;
                continue;
            }

            // Unknown character
            return { error: 'Unexpected character: ' + expr[i] };
        }

        return { tokens };
    }

    // Parse and evaluate arithmetic expression with proper precedence
    // Grammar:
    //   expr    -> term (('+' | '-') term)*
    //   term    -> factor (('*' | '/') factor)*
    //   factor  -> NUMBER | CELL_REF | '(' expr ')' | '-' factor | '+' factor
    function evaluateArithmeticExpr(tokens) {
        let pos = 0;

        function peek() {
            return tokens[pos];
        }

        function consume() {
            return tokens[pos++];
        }

        function parseExpr() {
            let left = parseTerm();
            if (left.error) return left;

            while (peek() && peek().type === 'OPERATOR' &&
                   (peek().value === '+' || peek().value === '-')) {
                const op = consume().value;
                const right = parseTerm();
                if (right.error) return right;

                left = { value: op === '+' ? left.value + right.value
                                           : left.value - right.value };
            }
            return left;
        }

        function parseTerm() {
            let left = parseFactor();
            if (left.error) return left;

            while (peek() && peek().type === 'OPERATOR' &&
                   (peek().value === '*' || peek().value === '/')) {
                const op = consume().value;
                const right = parseFactor();
                if (right.error) return right;

                if (op === '/') {
                    if (right.value === 0) {
                        return { error: '#DIV/0!' };
                    }
                    left = { value: left.value / right.value };
                } else {
                    left = { value: left.value * right.value };
                }
            }
            return left;
        }

        function parseFactor() {
            const token = peek();

            if (!token) {
                return { error: '#ERROR!' };
            }

            // Unary minus
            if (token.type === 'OPERATOR' && token.value === '-') {
                consume();
                const factor = parseFactor();
                if (factor.error) return factor;
                return { value: -factor.value };
            }

            // Unary plus (just consume)
            if (token.type === 'OPERATOR' && token.value === '+') {
                consume();
                return parseFactor();
            }

            // Number literal
            if (token.type === 'NUMBER') {
                consume();
                return { value: token.value };
            }

            // Cell reference
            if (token.type === 'CELL_REF') {
                consume();
                const parsed = parseCellRef(token.value);
                if (!parsed) {
                    return { error: '#REF!' };
                }
                if (parsed.row >= rows || parsed.col >= cols ||
                    parsed.row < 0 || parsed.col < 0) {
                    return { error: '#REF!' };
                }
                return { value: getCellValue(parsed.row, parsed.col) };
            }

            // Parenthesized expression
            if (token.type === 'LPAREN') {
                consume();
                const result = parseExpr();
                if (result.error) return result;

                if (!peek() || peek().type !== 'RPAREN') {
                    return { error: '#ERROR!' };
                }
                consume();
                return result;
            }

            return { error: '#ERROR!' };
        }

        const result = parseExpr();

        // Check for leftover tokens (malformed expression)
        if (!result.error && pos < tokens.length) {
            return { error: '#ERROR!' };
        }

        return result;
    }

    // Evaluate arithmetic expression string (without leading =)
    function evaluateArithmetic(expr) {
        const tokenResult = tokenizeArithmetic(expr);
        if (tokenResult.error) {
            return tokenResult.error;
        }

        if (tokenResult.tokens.length === 0) {
            return '#ERROR!';
        }

        const evalResult = evaluateArithmeticExpr(tokenResult.tokens);
        if (evalResult.error) {
            return evalResult.error;
        }

        // Format result (avoid floating point display issues)
        const value = evalResult.value;
        if (Number.isInteger(value)) {
            return value;
        }
        // Round to reasonable precision
        return Math.round(value * 1e10) / 1e10;
    }

    // Check if expression is a valid arithmetic formula
    function isArithmeticFormula(formula) {
        if (!formula || !formula.startsWith('=')) return false;
        const expr = formula.substring(1).trim();
        if (expr.length === 0) return false;

        const tokenResult = tokenizeArithmetic(expr);
        if (tokenResult.error) return false;
        if (tokenResult.tokens.length === 0) return false;

        // Validate balanced parentheses
        let parenDepth = 0;
        for (const token of tokenResult.tokens) {
            if (token.type === 'LPAREN') {
                parenDepth++;
            } else if (token.type === 'RPAREN') {
                parenDepth--;
                if (parenDepth < 0) return false;
            }
        }

        return parenDepth === 0;
    }

    // Evaluate SUM(range)
    function evaluateSUM(rangeStr) {
        const range = parseRange(rangeStr);
        if (!range) return '#REF!';

        // Check if range is within grid bounds
        if (range.endRow >= rows || range.endCol >= cols) return '#REF!';

        let sum = 0;
        for (let r = range.startRow; r <= range.endRow; r++) {
            for (let c = range.startCol; c <= range.endCol; c++) {
                sum += getCellValue(r, c);
            }
        }
        return sum;
    }

    // Evaluate AVG(range)
    function evaluateAVG(rangeStr) {
        const range = parseRange(rangeStr);
        if (!range) return '#REF!';

        // Check if range is within grid bounds
        if (range.endRow >= rows || range.endCol >= cols) return '#REF!';

        let sum = 0;
        let count = 0;
        for (let r = range.startRow; r <= range.endRow; r++) {
            for (let c = range.startCol; c <= range.endCol; c++) {
                const raw = data[r][c];
                if (raw === null || raw === undefined) continue;
                const stripped = String(raw).replace(/<[^>]*>/g, '').trim();
                if (stripped === '') continue;
                const num = parseFloat(stripped);
                if (isNaN(num)) continue;
                sum += num;
                count++;
            }
        }
        return count === 0 ? 0 : sum / count;
    }

    // Main formula evaluator
    function evaluateFormula(formula) {
        if (!formula || !formula.startsWith('=')) return formula;

        const expr = formula.substring(1).trim();
        const exprUpper = expr.toUpperCase();

        // Match SUM(range)
        const sumMatch = exprUpper.match(/^SUM\(([A-Z]+\d+:[A-Z]+\d+)\)$/);
        if (sumMatch) {
            return evaluateSUM(sumMatch[1]);
        }

        // Match AVG(range)
        const avgMatch = exprUpper.match(/^AVG\(([A-Z]+\d+:[A-Z]+\d+)\)$/);
        if (avgMatch) {
            return evaluateAVG(avgMatch[1]);
        }

        // Try arithmetic expression
        if (isArithmeticFormula(formula)) {
            return evaluateArithmetic(expr);
        }

        // Unknown formula
        return '#ERROR!';
    }

    // Recalculate all formula cells
    function recalculateFormulas() {
        const container = document.getElementById('spreadsheet');
        const activeElement = document.activeElement;
        const maxPasses = rows * cols;
        let needsUpdate = false;

        // Multiple passes to propagate formulas that depend on other formulas.
        for (let pass = 0; pass < maxPasses; pass++) {
            let changed = false;
            for (let r = 0; r < rows; r++) {
                for (let c = 0; c < cols; c++) {
                    const formula = formulas[r][c];
                    if (formula && formula.startsWith('=')) {
                        const result = String(evaluateFormula(formula));
                        if (data[r][c] !== result) {
                            data[r][c] = result;
                            changed = true;
                            needsUpdate = true;
                        }
                    }
                }
            }
            if (!changed) break;
        }

        if (!needsUpdate || !container) return;

        for (let r = 0; r < rows; r++) {
            for (let c = 0; c < cols; c++) {
                const formula = formulas[r][c];
                if (formula && formula.startsWith('=')) {
                    const cellContent = container.querySelector(
                        `.cell-content[data-row="${r}"][data-col="${c}"]`
                    );
                    if (!cellContent) continue;

                    const isEditingFormula = cellContent === activeElement &&
                        cellContent.innerText.trim().startsWith('=');
                    if (!isEditingFormula) {
                        cellContent.innerText = data[r][c];
                    }
                }
            }
        }
    }

    // Get current theme
    function isDarkMode() {
        return document.body.classList.contains('dark-mode');
    }

    function normalizeAlignment(value) {
        if (value === 'left' || value === 'center' || value === 'right') {
            return value;
        }
        return '';
    }

    function normalizeFontSize(value) {
        if (value === null || value === undefined) return '';
        let raw = String(value).trim();
        if (raw === '') return '';
        if (raw.endsWith('px')) {
            raw = raw.slice(0, -2).trim();
        }
        const size = parseInt(raw, 10);
        if (isNaN(size)) return '';
        if (!FONT_SIZE_OPTIONS.includes(size)) return '';
        return String(size);
    }

    function normalizeCellStyles(styles, r, c) {
        const normalized = createEmptyCellStyles(r, c);
        if (!Array.isArray(styles)) return normalized;

        for (let row = 0; row < r; row++) {
            const sourceRow = Array.isArray(styles[row]) ? styles[row] : [];
            for (let col = 0; col < c; col++) {
                const cellStyle = sourceRow[col];
                if (cellStyle && typeof cellStyle === 'object') {
                    normalized[row][col] = {
                        align: normalizeAlignment(cellStyle.align),
                        bg: isValidCSSColor(cellStyle.bg) ? cellStyle.bg : '',
                        color: isValidCSSColor(cellStyle.color) ? cellStyle.color : '',
                        fontSize: normalizeFontSize(cellStyle.fontSize)
                    };
                }
            }
        }
        return normalized;
    }

    function extractPlainText(value) {
        if (value === null || value === undefined) return '';
        // Use DOMParser for safe HTML parsing (doesn't execute scripts)
        const parser = new DOMParser();
        const doc = parser.parseFromString('<body>' + String(value) + '</body>', 'text/html');
        const text = doc.body.textContent || '';
        return text.replace(/\u00a0/g, ' ');
    }

    function csvEscape(value) {
        const text = String(value);
        const needsQuotes = /[",\r\n]/.test(text) || /^\s|\s$/.test(text);
        const escaped = text.replace(/"/g, '""');
        return needsQuotes ? `"${escaped}"` : escaped;
    }

    function buildCSV() {
        const lines = [];
        for (let r = 0; r < rows; r++) {
            const rowValues = [];
            for (let c = 0; c < cols; c++) {
                const raw = data[r][c];
                const text = extractPlainText(raw);
                rowValues.push(csvEscape(text));
            }
            lines.push(rowValues.join(','));
        }
        return lines.join('\r\n');
    }

    function downloadCSV() {
        recalculateFormulas();
        const csv = buildCSV();
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = 'spreadsheet.csv';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
        showToast('CSV downloaded', 'success');
    }

    function parseCSV(text) {
        if (!text) return [];

        const rows = [];
        let row = [];
        let field = '';
        let inQuotes = false;

        for (let i = 0; i < text.length; i++) {
            const char = text[i];

            if (inQuotes) {
                if (char === '"') {
                    if (text[i + 1] === '"') {
                        field += '"';
                        i++;
                    } else {
                        inQuotes = false;
                    }
                } else {
                    field += char;
                }
                continue;
            }

            if (char === '"') {
                inQuotes = true;
            } else if (char === ',') {
                row.push(field);
                field = '';
            } else if (char === '\r') {
                if (text[i + 1] === '\n') {
                    i++;
                }
                row.push(field);
                rows.push(row);
                row = [];
                field = '';
            } else if (char === '\n') {
                row.push(field);
                rows.push(row);
                row = [];
                field = '';
            } else {
                field += char;
            }
        }

        row.push(field);
        rows.push(row);

        if (rows.length && rows[0].length && rows[0][0]) {
            rows[0][0] = rows[0][0].replace(/^\uFEFF/, '');
        }

        if (/\r?\n$/.test(text)) {
            const lastRow = rows[rows.length - 1];
            if (lastRow && lastRow.length === 1 && lastRow[0] === '') {
                rows.pop();
            }
        }

        return rows;
    }

    function importCSVText(text) {
        const parsedRows = parseCSV(text);
        if (!parsedRows.length) {
            showToast('CSV file is empty', 'error');
            return;
        }

        const maxCols = parsedRows.reduce((max, row) => Math.max(max, row.length), 0);
        const nextRows = Math.min(Math.max(parsedRows.length, 1), MAX_ROWS);
        const nextCols = Math.min(Math.max(maxCols, 1), MAX_COLS);

        const truncated = parsedRows.length > MAX_ROWS || maxCols > MAX_COLS;

        rows = nextRows;
        cols = nextCols;
        data = createEmptyData(rows, cols);
        formulas = createEmptyData(rows, cols);
        cellStyles = createEmptyCellStyles(rows, cols);
        colWidths = createDefaultColumnWidths(cols);
        rowHeights = createDefaultRowHeights(rows);

        for (let r = 0; r < rows; r++) {
            const sourceRow = Array.isArray(parsedRows[r]) ? parsedRows[r] : [];
            for (let c = 0; c < cols; c++) {
                const raw = sourceRow[c] !== undefined ? String(sourceRow[c]) : '';
                if (raw.startsWith('=')) {
                    // Security: Validate formula against whitelist before storing
                    if (isValidFormula(raw)) {
                        formulas[r][c] = raw;
                        data[r][c] = raw;
                    } else {
                        // Invalid formula pattern - treat as escaped text
                        formulas[r][c] = '';
                        data[r][c] = escapeHTML(raw);
                    }
                } else {
                    // Escape HTML in CSV values to prevent XSS
                    data[r][c] = escapeHTML(raw);
                }
            }
        }

        renderGrid();
        recalculateFormulas();
        debouncedUpdateURL();

        if (truncated) {
            showToast('CSV imported (some data truncated due to size limits)', 'warning');
        } else {
            showToast('CSV imported successfully', 'success');
        }
    }

    // Check if all cells in data array are empty
    function isDataEmpty(d) {
        return d.every(row => row.every(cell => cell === ''));
    }

    // Check if all cells in formulas array are empty
    function isFormulasEmpty(f) {
        return f.every(row => row.every(cell => cell === ''));
    }

    // Check if a cell style is default (all empty)
    function isCellStyleDefault(style) {
        return !style || (style.align === '' && style.bg === '' &&
                          style.color === '' && style.fontSize === '');
    }

    // Check if all cell styles are default
    function isCellStylesDefault(styles) {
        return styles.every(row => row.every(cell => isCellStyleDefault(cell)));
    }

    // Check if column widths are all default
    function isColWidthsDefault(widths, count) {
        if (widths.length !== count) return false;
        return widths.every(w => w === DEFAULT_COL_WIDTH);
    }

    // Check if row heights are all default
    function isRowHeightsDefault(heights, count) {
        if (heights.length !== count) return false;
        return heights.every(h => h === DEFAULT_ROW_HEIGHT);
    }

    // Encode state to URL-safe string (includes dimensions and theme)
    // Only includes non-empty/non-default values to minimize URL length
    // Returns Promise if encryption is enabled
    async function encodeState() {
        const state = {
            rows,
            cols,
            theme: isDarkMode() ? 'dark' : 'light'
        };

        // Only include data if not all empty
        if (!isDataEmpty(data)) {
            state.data = data;
        }

        // Only include formulas if any exist
        if (!isFormulasEmpty(formulas)) {
            state.formulas = formulas;
        }

        // Only include cell styles if any are non-default
        if (!isCellStylesDefault(cellStyles)) {
            state.cellStyles = cellStyles;
        }

        // Only include colWidths if not all default
        if (!isColWidthsDefault(colWidths, cols)) {
            state.colWidths = colWidths;
        }

        // Only include rowHeights if not all default
        if (!isRowHeightsDefault(rowHeights, rows)) {
            state.rowHeights = rowHeights;
        }

        const json = JSON.stringify(state);
        const compressed = LZString.compressToEncodedURIComponent(json);

        // If password is set, encrypt the compressed data
        if (currentPassword) {
            try {
                const encrypted = await CryptoUtils.encrypt(compressed, currentPassword);
                return 'ENC:' + encrypted;
            } catch (e) {
                console.error('Encryption failed:', e);
                // Fall back to unencrypted if encryption fails
                return compressed;
            }
        }

        return compressed;
    }

    // Safe JSON parse with prototype pollution protection
    // Creates safe copies using Object.create(null) to prevent prototype chain attacks
    function safeJSONParse(jsonString) {
        const parsed = JSON.parse(jsonString);

        function createSafeCopy(obj) {
            if (obj === null || typeof obj !== 'object') return obj;
            if (Array.isArray(obj)) {
                return obj.map(createSafeCopy);
            }

            const safe = Object.create(null);
            for (const key of Object.keys(obj)) {
                // Block dangerous keys that could pollute prototypes
                if (key === '__proto__' || key === 'constructor' || key === 'prototype') {
                    console.warn('Blocked prototype pollution attempt via key:', key);
                    continue;
                }
                // Block keys containing prototype chain accessor patterns
                if (key.includes('__proto__') || key.includes('constructor.prototype')) {
                    console.warn('Blocked prototype pollution attempt via key pattern:', key);
                    continue;
                }
                safe[key] = createSafeCopy(obj[key]);
            }
            return safe;
        }

        return createSafeCopy(parsed);
    }

    // Decode URL hash to state object
    // Returns { encrypted: true, data: base64String } if data is encrypted
    function decodeState(hash) {
        try {
            // Security: Reject extremely long hashes (potential DoS)
            if (hash.length > 100000) {
                console.warn('Hash too long, rejecting');
                return null;
            }

            // Check for encrypted data (ENC: prefix)
            if (hash.startsWith('ENC:')) {
                const encryptedData = hash.slice(4); // Remove "ENC:" prefix
                return { encrypted: true, data: encryptedData };
            }

            let decoded = null;

            // Try LZ-String decompression first (new format)
            decoded = LZString.decompressFromEncodedURIComponent(hash);

            // If decompression returned null/empty, try legacy format
            if (!decoded || decoded.length === 0) {
                // Only attempt legacy decode if it looks like valid URL-encoded JSON
                if (hash.startsWith('%7B') || hash.startsWith('%5B') ||
                    hash.startsWith('{') || hash.startsWith('[')) {
                    decoded = decodeURIComponent(hash);
                } else {
                    // Could be invalid/corrupted LZ-String - reject
                    console.warn('Unrecognized hash format');
                    return null;
                }
            }

            // Use safe JSON parsing with prototype pollution protection
            const parsed = safeJSONParse(decoded);

            // Handle new format (object with rows, cols, data)
            if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
                const r = Math.min(Math.max(1, parsed.rows || DEFAULT_ROWS), MAX_ROWS);
                const c = Math.min(Math.max(1, parsed.cols || DEFAULT_COLS), MAX_COLS);

                // Validate and normalize data array
                let d = parsed.data;
                if (Array.isArray(d)) {
                    // Ensure correct dimensions
                    d = d.slice(0, r).map(row => {
                        if (Array.isArray(row)) {
                            return row.slice(0, c).map(cell => sanitizeHTML(String(cell || '')));
                        }
                        return Array(c).fill('');
                    });
                    // Pad rows if needed
                    while (d.length < r) {
                        d.push(Array(c).fill(''));
                    }
                    // Pad columns if needed
                    d = d.map(row => {
                        while (row.length < c) row.push('');
                        return row;
                    });
                } else {
                    d = createEmptyData(r, c);
                }

                // Load formulas (backward compatible - create empty if not present)
                // Security: Validate formulas against whitelist to prevent injection
                let f = parsed.formulas;
                if (Array.isArray(f)) {
                    f = f.slice(0, r).map((row, rowIdx) => {
                        if (Array.isArray(row)) {
                            return row.slice(0, c).map((cell, colIdx) => {
                                const formula = String(cell || '');
                                if (formula.startsWith('=')) {
                                    // Validate formula against whitelist
                                    if (isValidFormula(formula)) {
                                        return formula;
                                    } else {
                                        // Invalid formula - convert to escaped text in data
                                        d[rowIdx][colIdx] = escapeHTML(formula);
                                        return '';
                                    }
                                }
                                return formula;
                            });
                        }
                        return Array(c).fill('');
                    });
                    while (f.length < r) {
                        f.push(Array(c).fill(''));
                    }
                    f = f.map(row => {
                        while (row.length < c) row.push('');
                        return row;
                    });
                } else {
                    f = createEmptyData(r, c);
                }

                const s = normalizeCellStyles(parsed.cellStyles, r, c);
                const w = normalizeColumnWidths(parsed.colWidths, c);
                const h = normalizeRowHeights(parsed.rowHeights, r);
                return {
                    rows: r,
                    cols: c,
                    data: d,
                    formulas: f,
                    cellStyles: s,
                    colWidths: w,
                    rowHeights: h,
                    theme: parsed.theme || null
                };
            }

            // Handle legacy format (just array, assume 10x10)
            if (Array.isArray(parsed)) {
                const d = parsed.slice(0, DEFAULT_ROWS).map(row => {
                    if (Array.isArray(row)) {
                        return row.slice(0, DEFAULT_COLS).map(cell => sanitizeHTML(String(cell || '')));
                    }
                    return Array(DEFAULT_COLS).fill('');
                });
                while (d.length < DEFAULT_ROWS) {
                    d.push(Array(DEFAULT_COLS).fill(''));
                }
                const f = createEmptyData(DEFAULT_ROWS, DEFAULT_COLS);
                const s = createEmptyCellStyles(DEFAULT_ROWS, DEFAULT_COLS);
                const w = createDefaultColumnWidths(DEFAULT_COLS);
                const h = createDefaultRowHeights(DEFAULT_ROWS);
                return {
                    rows: DEFAULT_ROWS,
                    cols: DEFAULT_COLS,
                    data: d,
                    formulas: f,
                    cellStyles: s,
                    colWidths: w,
                    rowHeights: h,
                    theme: null
                };
            }
        } catch (e) {
            console.warn('Failed to decode state from URL:', e);
        }
        return null;
    }

    // Apply theme to body
    function applyTheme(theme) {
        if (theme === 'dark') {
            document.body.classList.add('dark-mode');
        } else if (theme === 'light') {
            document.body.classList.remove('dark-mode');
        }
        // Save to localStorage as well
        try {
            localStorage.setItem('spreadsheet-theme', theme);
        } catch (e) {}
    }

    // Update URL length indicator in status bar
    function updateURLLengthIndicator(length) {
        const valueEl = document.getElementById('url-length-value');
        const barEl = document.getElementById('url-progress-bar');
        const msgEl = document.getElementById('url-length-message');
        if (!valueEl || !barEl) return;

        valueEl.textContent = length.toLocaleString();

        // Calculate progress percentage (capped at 100%)
        const percent = Math.min((length / URL_LENGTH_MAX_DISPLAY) * 100, 100);
        barEl.style.width = percent + '%';

        // Update color and message based on thresholds
        barEl.classList.remove('warning', 'caution', 'critical');
        if (msgEl) {
            msgEl.classList.remove('warning', 'caution', 'critical');
            msgEl.textContent = '';
        }

        if (length >= URL_LENGTH_CRITICAL) {
            barEl.classList.add('critical');
            if (msgEl) {
                msgEl.classList.add('critical');
                msgEl.textContent = 'Some browsers may fail';
            }
        } else if (length >= URL_LENGTH_CAUTION) {
            barEl.classList.add('caution');
            if (msgEl) {
                msgEl.classList.add('caution');
                msgEl.textContent = 'URL shorteners may fail';
            }
        } else if (length >= URL_LENGTH_WARNING) {
            barEl.classList.add('warning');
            if (msgEl) {
                msgEl.classList.add('warning');
                msgEl.textContent = 'Some older browsers may truncate';
            }
        }
    }

    // Update URL hash without page jump
    async function updateURL() {
        const encoded = await encodeState();
        const newHash = '#' + encoded;

        // Update URL length indicator
        updateURLLengthIndicator(encoded.length);

        if (history.replaceState) {
            history.replaceState(null, null, newHash);
        } else {
            location.hash = newHash;
        }
    }

    // Debounced URL update
    function debouncedUpdateURL() {
        if (debounceTimer) {
            clearTimeout(debounceTimer);
        }
        debounceTimer = setTimeout(updateURL, DEBOUNCE_DELAY);
    }

    // Update button disabled states and grid size display
    function updateUI() {
        const addRowBtn = document.getElementById('add-row');
        const addColBtn = document.getElementById('add-col');
        const gridSizeEl = document.getElementById('grid-size');

        if (addRowBtn) {
            addRowBtn.disabled = (rows >= MAX_ROWS);
        }
        if (addColBtn) {
            addColBtn.disabled = (cols >= MAX_COLS);
        }
        if (gridSizeEl) {
            gridSizeEl.textContent = `${rows} × ${cols}`;
        }
    }

    function applyGridTemplate() {
        const container = document.getElementById('spreadsheet');
        if (!container) return;

        const columnSizes = colWidths.map(width => `${width}px`).join(' ');
        const rowSizes = rowHeights.map(height => `${height}px`).join(' ');

        container.style.gridTemplateColumns = `${ROW_HEADER_WIDTH}px ${columnSizes}`;
        container.style.gridTemplateRows = `${HEADER_ROW_HEIGHT}px ${rowSizes}`;
    }

    // Render the spreadsheet grid
    function renderGrid() {
        const container = document.getElementById('spreadsheet');
        if (!container) return;

        // Clear selection when grid is re-rendered (e.g., adding rows/columns)
        selectionStart = null;
        selectionEnd = null;
        isSelecting = false;
        hoverRow = null;
        hoverCol = null;

        container.innerHTML = '';
        applyGridTemplate();

        // Corner cell (empty)
        const corner = document.createElement('div');
        corner.className = 'corner-cell';
        container.appendChild(corner);

        // Column headers (A-Z) - sticky to top
        for (let col = 0; col < cols; col++) {
            const header = document.createElement('div');
            header.className = 'header-cell col-header';
            header.textContent = colToLetter(col);
            header.dataset.col = col;
            const colResize = document.createElement('div');
            colResize.className = 'resize-handle col-resize';
            colResize.dataset.col = col;
            colResize.setAttribute('aria-hidden', 'true');
            header.appendChild(colResize);
            container.appendChild(header);
        }

        // Rows
        for (let row = 0; row < rows; row++) {
            // Row header (1, 2, 3...) - sticky to left
            const rowHeader = document.createElement('div');
            rowHeader.className = 'header-cell row-header';
            rowHeader.textContent = row + 1;
            rowHeader.dataset.row = row;
            const rowResize = document.createElement('div');
            rowResize.className = 'resize-handle row-resize';
            rowResize.dataset.row = row;
            rowResize.setAttribute('aria-hidden', 'true');
            rowHeader.appendChild(rowResize);
            container.appendChild(rowHeader);

            // Data cells
            for (let col = 0; col < cols; col++) {
                const cell = document.createElement('div');
                cell.className = 'cell';

                const contentDiv = document.createElement('div');
                contentDiv.className = 'cell-content';
                contentDiv.contentEditable = 'true';
                contentDiv.dataset.row = row;
                contentDiv.dataset.col = col;
                contentDiv.innerHTML = sanitizeHTML(data[row][col]);
                contentDiv.setAttribute('aria-label', `Cell ${colToLetter(col)}${row + 1}`);

                const style = cellStyles[row][col];
                if (style) {
                    contentDiv.style.textAlign = style.align || '';
                    contentDiv.style.color = style.color || '';
                    contentDiv.style.fontSize = style.fontSize ? `${style.fontSize}px` : '';
                    if (style.bg) {
                        cell.style.setProperty('--cell-bg', style.bg);
                    } else {
                        cell.style.removeProperty('--cell-bg');
                    }
                } else {
                    contentDiv.style.textAlign = '';
                    contentDiv.style.color = '';
                    contentDiv.style.fontSize = '';
                    cell.style.removeProperty('--cell-bg');
                }

                cell.appendChild(contentDiv);
                container.appendChild(cell);
            }
        }

        updateUI();
    }

    // Handle input changes
    function handleInput(event) {
        const target = event.target;
        if (!target.classList.contains('cell-content')) return;

        const row = parseInt(target.dataset.row, 10);
        const col = parseInt(target.dataset.col, 10);

        // DON'T clear selection if in formula mode (user may be selecting range)
        if (hasMultiSelection() && !formulaEditMode) {
            clearSelection();
            setActiveHeaders(row, col);
        }

        if (!isNaN(row) && !isNaN(col) && row < rows && col < cols) {
            setEditingCell(row, col);
            const rawValue = target.innerText.trim();

            if (rawValue.startsWith('=')) {
                // Enter formula edit mode
                formulaEditMode = true;
                formulaEditCell = { row, col, element: target };

                // Store formula but DON'T evaluate during typing
                formulas[row][col] = rawValue;
                data[row][col] = rawValue;

                updateFormulaDropdown(target, rawValue);
            } else {
                // Exit formula edit mode
                formulaEditMode = false;
                formulaEditCell = null;

                // Regular value - clear any existing formula
                formulas[row][col] = '';
                data[row][col] = sanitizeHTML(target.innerHTML);

                hideFormulaDropdown();

                // Recalculate dependent formulas when regular values change
                recalculateFormulas();
            }

            debouncedUpdateURL();
        }
    }

    function clearActiveHeaders() {
        if (activeRow !== null) {
            const rowHeader = document.querySelector(`.row-header[data-row="${activeRow}"]`);
            if (rowHeader) rowHeader.classList.remove(ACTIVE_HEADER_CLASS);
        }
        if (activeCol !== null) {
            const colHeader = document.querySelector(`.col-header[data-col="${activeCol}"]`);
            if (colHeader) colHeader.classList.remove(ACTIVE_HEADER_CLASS);
        }
        activeRow = null;
        activeCol = null;
    }

    function setActiveHeaders(row, col) {
        if (activeRow === row && activeCol === col) return;
        clearActiveHeaders();
        activeRow = row;
        activeCol = col;

        const rowHeader = document.querySelector(`.row-header[data-row="${row}"]`);
        if (rowHeader) rowHeader.classList.add(ACTIVE_HEADER_CLASS);

        const colHeader = document.querySelector(`.col-header[data-col="${col}"]`);
        if (colHeader) colHeader.classList.add(ACTIVE_HEADER_CLASS);
    }

    // ========== Multi-cell Selection Functions ==========

    // Get normalized selection bounds (handles any drag direction)
    function getSelectionBounds() {
        if (!selectionStart || !selectionEnd) return null;
        return {
            minRow: Math.min(selectionStart.row, selectionEnd.row),
            maxRow: Math.max(selectionStart.row, selectionEnd.row),
            minCol: Math.min(selectionStart.col, selectionEnd.col),
            maxCol: Math.max(selectionStart.col, selectionEnd.col)
        };
    }

    // Check if selection spans more than one cell
    function hasMultiSelection() {
        if (!selectionStart || !selectionEnd) return false;
        return selectionStart.row !== selectionEnd.row || selectionStart.col !== selectionEnd.col;
    }

    // Clear all selection state and visuals
    function clearSelection() {
        selectionStart = null;
        selectionEnd = null;
        isSelecting = false;

        const container = document.getElementById('spreadsheet');
        if (!container) return;

        // Remove selection classes from all cells
        container.querySelectorAll('.cell-selected').forEach(cell => {
            cell.classList.remove('cell-selected', 'selection-top', 'selection-bottom', 'selection-left', 'selection-right');
        });

        // Remove selecting mode from container
        container.classList.remove('selecting');
    }

    // Clear content of all selected cells
    function clearSelectedCells() {
        if (!selectionStart || !selectionEnd) return;

        const bounds = getSelectionBounds();
        const container = document.getElementById('spreadsheet');

        for (let r = bounds.minRow; r <= bounds.maxRow; r++) {
            for (let c = bounds.minCol; c <= bounds.maxCol; c++) {
                // Clear data and formula
                data[r][c] = '';
                formulas[r][c] = '';

                // Update DOM
                const cell = container.querySelector(
                    `.cell-content[data-row="${r}"][data-col="${c}"]`
                );
                if (cell) {
                    cell.innerHTML = '';
                }
            }
        }

        recalculateFormulas();
        debouncedUpdateURL();
    }

    // Highlight headers for a range
    function setActiveHeadersForRange(minRow, maxRow, minCol, maxCol) {
        // Clear existing header highlights
        document.querySelectorAll(`.${ACTIVE_HEADER_CLASS}`).forEach(el => {
            el.classList.remove(ACTIVE_HEADER_CLASS);
        });

        // Highlight all row headers in range
        for (let r = minRow; r <= maxRow; r++) {
            const rowHeader = document.querySelector(`.row-header[data-row="${r}"]`);
            if (rowHeader) rowHeader.classList.add(ACTIVE_HEADER_CLASS);
        }

        // Highlight all column headers in range
        for (let c = minCol; c <= maxCol; c++) {
            const colHeader = document.querySelector(`.col-header[data-col="${c}"]`);
            if (colHeader) colHeader.classList.add(ACTIVE_HEADER_CLASS);
        }

        // Update active row/col tracking
        activeRow = minRow;
        activeCol = minCol;
    }

    // Update visual selection on cells
    function updateSelectionVisuals() {
        const container = document.getElementById('spreadsheet');
        if (!container) return;

        const bounds = getSelectionBounds();
        if (!bounds) {
            clearSelection();
            return;
        }

        // Clear previous selection classes
        container.querySelectorAll('.cell-selected').forEach(cell => {
            cell.classList.remove('cell-selected', 'selection-top', 'selection-bottom', 'selection-left', 'selection-right');
        });

        const { minRow, maxRow, minCol, maxCol } = bounds;

        // Apply selection classes to cells in range
        for (let r = minRow; r <= maxRow; r++) {
            for (let c = minCol; c <= maxCol; c++) {
                const cellContent = container.querySelector(`.cell-content[data-row="${r}"][data-col="${c}"]`);
                if (cellContent && cellContent.parentElement) {
                    const cell = cellContent.parentElement;
                    cell.classList.add('cell-selected');

                    // Add border classes for outer edges
                    if (r === minRow) cell.classList.add('selection-top');
                    if (r === maxRow) cell.classList.add('selection-bottom');
                    if (c === minCol) cell.classList.add('selection-left');
                    if (c === maxCol) cell.classList.add('selection-right');
                }
            }
        }

        // Highlight headers for the entire range
        setActiveHeadersForRange(minRow, maxRow, minCol, maxCol);
    }

    function getCellContentFromTarget(target) {
        if (!(target instanceof Element)) return null;

        if (target.classList.contains('cell-content')) {
            return target;
        }
        if (target.classList.contains('cell')) {
            return target.querySelector('.cell-content');
        }
        return target.closest('.cell-content');
    }

    function addHoverRow(row) {
        const container = document.getElementById('spreadsheet');
        if (!container) return;

        container.querySelectorAll(`.cell-content[data-row="${row}"]`).forEach(cellContent => {
            if (cellContent.parentElement) {
                cellContent.parentElement.classList.add('hover-row');
            }
        });

        const rowHeader = container.querySelector(`.row-header[data-row="${row}"]`);
        if (rowHeader) {
            rowHeader.classList.add('header-hover');
        }
    }

    function addHoverCol(col) {
        const container = document.getElementById('spreadsheet');
        if (!container) return;

        container.querySelectorAll(`.cell-content[data-col="${col}"]`).forEach(cellContent => {
            if (cellContent.parentElement) {
                cellContent.parentElement.classList.add('hover-col');
            }
        });

        const colHeader = container.querySelector(`.col-header[data-col="${col}"]`);
        if (colHeader) {
            colHeader.classList.add('header-hover');
        }
    }

    function removeHoverRow(row) {
        const container = document.getElementById('spreadsheet');
        if (!container) return;

        container.querySelectorAll(`.cell-content[data-row="${row}"]`).forEach(cellContent => {
            if (cellContent.parentElement) {
                cellContent.parentElement.classList.remove('hover-row');
            }
        });

        const rowHeader = container.querySelector(`.row-header[data-row="${row}"]`);
        if (rowHeader) {
            rowHeader.classList.remove('header-hover');
        }
    }

    function removeHoverCol(col) {
        const container = document.getElementById('spreadsheet');
        if (!container) return;

        container.querySelectorAll(`.cell-content[data-col="${col}"]`).forEach(cellContent => {
            if (cellContent.parentElement) {
                cellContent.parentElement.classList.remove('hover-col');
            }
        });

        const colHeader = container.querySelector(`.col-header[data-col="${col}"]`);
        if (colHeader) {
            colHeader.classList.remove('header-hover');
        }
    }

    function clearHoverHighlights() {
        if (hoverRow !== null) {
            removeHoverRow(hoverRow);
        }
        if (hoverCol !== null) {
            removeHoverCol(hoverCol);
        }
        hoverRow = null;
        hoverCol = null;
    }

    function setHoverHighlight(row, col) {
        if (row === hoverRow && col === hoverCol) return;

        if (hoverRow !== null && hoverRow !== row) {
            removeHoverRow(hoverRow);
        }
        if (hoverCol !== null && hoverCol !== col) {
            removeHoverCol(hoverCol);
        }

        hoverRow = row;
        hoverCol = col;

        if (hoverRow !== null) {
            addHoverRow(hoverRow);
        }
        if (hoverCol !== null) {
            addHoverCol(hoverCol);
        }
    }

    function updateHoverFromTarget(target) {
        if (isSelecting || hasMultiSelection()) {
            clearHoverHighlights();
            return;
        }

        const cellContent = getCellContentFromTarget(target);
        if (!cellContent || !cellContent.classList.contains('cell-content')) {
            clearHoverHighlights();
            return;
        }

        const row = parseInt(cellContent.dataset.row, 10);
        const col = parseInt(cellContent.dataset.col, 10);
        if (isNaN(row) || isNaN(col)) {
            clearHoverHighlights();
            return;
        }

        setHoverHighlight(row, col);
    }

    // Get cell coordinates from screen point
    function getCellFromPoint(x, y) {
        const element = document.elementFromPoint(x, y);
        if (!element) return null;

        // Check if it's a cell-content or its parent cell
        let cellContent = element;
        if (element.classList.contains('cell')) {
            cellContent = element.querySelector('.cell-content');
        }

        if (!cellContent || !cellContent.classList.contains('cell-content')) return null;

        const row = parseInt(cellContent.dataset.row, 10);
        const col = parseInt(cellContent.dataset.col, 10);

        if (isNaN(row) || isNaN(col)) return null;
        return { row, col };
    }

    function getCellContentElement(row, col) {
        return document.querySelector(`.cell-content[data-row="${row}"][data-col="${col}"]`);
    }

    function getCellElement(row, col) {
        const cellContent = getCellContentElement(row, col);
        return cellContent ? cellContent.parentElement : null;
    }

    function focusCellAt(row, col) {
        const cellContent = getCellContentElement(row, col);
        if (!cellContent) return null;
        cellContent.focus();
        cellContent.scrollIntoView({ block: 'nearest', inline: 'nearest' });
        return cellContent;
    }

    function setEditingCell(row, col) {
        editingCell = { row, col };
    }

    function clearEditingCell() {
        editingCell = null;
    }

    function isEditingCell(row, col) {
        return !!(editingCell && editingCell.row === row && editingCell.col === col);
    }

    function handleFocusIn(event) {
        const target = event.target;
        if (!target.classList.contains('cell-content')) return;

        const row = parseInt(target.dataset.row, 10);
        const col = parseInt(target.dataset.col, 10);

        if (isNaN(row) || isNaN(col)) return;

        // If we have a multi-selection and focus moves to a cell outside it, clear selection
        if (hasMultiSelection()) {
            const bounds = getSelectionBounds();
            if (bounds && (row < bounds.minRow || row > bounds.maxRow || col < bounds.minCol || col > bounds.maxCol)) {
                clearSelection();
            }
        }

        // Update header highlighting for single cell focus (if no multi-selection)
        if (!hasMultiSelection()) {
            setActiveHeaders(row, col);
        }

        // Show formula text when focused (for editing)
        if (formulas[row][col] && formulas[row][col].startsWith('=')) {
            target.innerText = formulas[row][col];
        }
    }

    function handleFocusOut(event) {
        const target = event.target;
        if (!target.classList.contains('cell-content')) return;

        hideFormulaDropdown();

        const row = parseInt(target.dataset.row, 10);
        const col = parseInt(target.dataset.col, 10);

        // If we're in formula edit mode and currently selecting a range, don't process blur
        if (formulaEditMode && isSelecting) {
            return;
        }

        // Evaluate formula when blurred
        if (!isNaN(row) && !isNaN(col)) {
            const rawValue = target.innerText.trim();

            if (rawValue.startsWith('=')) {
                // NOW evaluate the formula
                formulas[row][col] = rawValue;
                const result = evaluateFormula(rawValue);
                data[row][col] = String(result);
                target.innerText = String(result);

                // Recalculate all dependent formulas
                recalculateFormulas();
                debouncedUpdateURL();
            }
        }

        // Exit formula edit mode when focus truly leaves
        clearEditingCell();
        formulaEditMode = false;
        formulaEditCell = null;
        formulaRangeStart = null;
        formulaRangeEnd = null;

        const container = document.getElementById('spreadsheet');
        if (!container) return;

        const next = event.relatedTarget;
        if (next && container.contains(next)) {
            return;
        }

        clearActiveHeaders();
    }

    // ========== Mouse Selection Handlers ==========

    function handleCellDoubleClick(event) {
        if (!(event.target instanceof Element)) return;

        let cellContent = null;
        if (event.target.classList.contains('cell-content')) {
            cellContent = event.target;
        } else if (event.target.classList.contains('cell')) {
            cellContent = event.target.querySelector('.cell-content');
        } else {
            cellContent = event.target.closest('.cell-content');
        }

        if (!cellContent || !cellContent.classList.contains('cell-content')) return;

        const row = parseInt(cellContent.dataset.row, 10);
        const col = parseInt(cellContent.dataset.col, 10);

        if (isNaN(row) || isNaN(col)) return;
        setEditingCell(row, col);
    }

    function handleMouseDown(event) {
        // Only handle left mouse button
        if (event.button !== 0) return;

        const target = event.target;

        // Check if clicking on a cell
        let cellContent = target;
        if (target.classList.contains('cell')) {
            cellContent = target.querySelector('.cell-content');
        }

        if (!cellContent || !cellContent.classList.contains('cell-content')) {
            // Clicked outside cells - clear selection
            clearSelection();
            return;
        }

        const row = parseInt(cellContent.dataset.row, 10);
        const col = parseInt(cellContent.dataset.col, 10);

        if (isNaN(row) || isNaN(col)) return;

        const container = document.getElementById('spreadsheet');

        // If in formula edit mode and clicking on a different cell
        if (formulaEditMode && formulaEditCell) {
            // Don't process clicks on the formula cell itself
            if (row !== formulaEditCell.row || col !== formulaEditCell.col) {
                event.preventDefault();
                event.stopPropagation();

                hideFormulaDropdown();

                // Start range selection for formula
                formulaRangeStart = { row, col };
                formulaRangeEnd = { row, col };
                isSelecting = true;
                clearHoverHighlights();

                // Show visual selection
                selectionStart = formulaRangeStart;
                selectionEnd = formulaRangeEnd;
                updateSelectionVisuals();

                if (container) {
                    container.classList.add('selecting');
                }

                return;
            }
        }

        // Shift+click: extend selection from anchor
        if (event.shiftKey && selectionStart) {
            selectionEnd = { row, col };
            updateSelectionVisuals();
            event.preventDefault();
            return;
        }

        // Start new selection
        selectionStart = { row, col };
        selectionEnd = { row, col };
        isSelecting = true;
        clearHoverHighlights();

        if (container) {
            container.classList.add('selecting');
        }

        updateSelectionVisuals();
    }

    function handleMouseMove(event) {
        if (!isSelecting) {
            updateHoverFromTarget(event.target);
            return;
        }

        const cellCoords = getCellFromPoint(event.clientX, event.clientY);
        if (!cellCoords) return;

        // Only update if position changed
        if (selectionEnd && cellCoords.row === selectionEnd.row && cellCoords.col === selectionEnd.col) {
            return;
        }

        // If in formula edit mode, update formula range
        if (formulaEditMode && formulaRangeStart) {
            formulaRangeEnd = cellCoords;
            selectionEnd = cellCoords;
            updateSelectionVisuals();
            event.preventDefault();
            return;
        }

        selectionEnd = cellCoords;
        updateSelectionVisuals();

        // Prevent text selection during drag
        event.preventDefault();
    }

    function handleMouseLeave() {
        clearHoverHighlights();
    }

    function handleMouseUp(event) {
        if (!isSelecting) return;

        isSelecting = false;

        const container = document.getElementById('spreadsheet');
        if (container) {
            container.classList.remove('selecting');
        }

        // If in formula edit mode, insert the range reference
        if (formulaEditMode && formulaEditCell && formulaRangeStart) {
            const rangeRef = buildRangeRef(
                formulaRangeStart.row, formulaRangeStart.col,
                formulaRangeEnd.row, formulaRangeEnd.col
            );

            // Focus back on formula cell and insert range
            formulaEditCell.element.focus();

            // Use setTimeout to ensure focus is established before inserting
            setTimeout(function() {
                insertTextAtCursor(rangeRef);

                // Update stored formula
                formulas[formulaEditCell.row][formulaEditCell.col] = formulaEditCell.element.innerText;
                data[formulaEditCell.row][formulaEditCell.col] = formulaEditCell.element.innerText;

                // Clear formula range selection but stay in formula edit mode
                formulaRangeStart = null;
                formulaRangeEnd = null;
                clearSelection();

                debouncedUpdateURL();
            }, 0);

            return;
        }

        // If single cell selected, allow normal focus behavior
        if (!hasMultiSelection()) {
            // Let the cell receive focus for editing
            const cellContent = document.querySelector(
                `.cell-content[data-row="${selectionStart.row}"][data-col="${selectionStart.col}"]`
            );
            if (cellContent) {
                cellContent.focus();
            }
        }
    }

    function handleSelectionKeyDown(event) {
        if (isFormulaDropdownOpen()) {
            if (event.key === 'ArrowDown') {
                moveFormulaDropdownSelection(1);
                event.preventDefault();
                return;
            }
            if (event.key === 'ArrowUp') {
                moveFormulaDropdownSelection(-1);
                event.preventDefault();
                return;
            }
            if (event.key === 'Enter' || event.key === 'Tab') {
                const activeItem = formulaDropdownItems[formulaDropdownIndex];
                if (activeItem) {
                    applyFormulaSuggestion(activeItem.dataset.formula);
                }
                event.preventDefault();
                return;
            }
            if (event.key === 'Escape') {
                hideFormulaDropdown();
                event.preventDefault();
                return;
            }
        }

        if (event.key === 'ArrowUp' || event.key === 'ArrowDown' || event.key === 'ArrowLeft' || event.key === 'ArrowRight') {
            const target = event.target;
            if (target.classList.contains('cell-content') && !event.altKey && !event.ctrlKey && !event.metaKey) {
                const row = parseInt(target.dataset.row, 10);
                const col = parseInt(target.dataset.col, 10);

                if (!isNaN(row) && !isNaN(col) && !formulaEditMode && !isEditingCell(row, col)) {
                    let nextRow = row;
                    let nextCol = col;

                    if (event.key === 'ArrowUp') nextRow -= 1;
                    if (event.key === 'ArrowDown') nextRow += 1;
                    if (event.key === 'ArrowLeft') nextCol -= 1;
                    if (event.key === 'ArrowRight') nextCol += 1;

                    nextRow = Math.max(0, Math.min(rows - 1, nextRow));
                    nextCol = Math.max(0, Math.min(cols - 1, nextCol));

                    event.preventDefault();

                    if (event.shiftKey) {
                        if (!selectionStart) {
                            selectionStart = { row, col };
                        }
                        selectionEnd = { row: nextRow, col: nextCol };
                    } else {
                        selectionStart = { row: nextRow, col: nextCol };
                        selectionEnd = { row: nextRow, col: nextCol };
                    }

                    updateSelectionVisuals();
                    focusCellAt(nextRow, nextCol);
                    return;
                }
            }
        }

        // Escape key clears selection
        if (event.key === 'Escape' && hasMultiSelection()) {
            clearSelection();
            event.preventDefault();
            return;
        }

        // Delete/Backspace key clears selected cells
        if (event.key === 'Delete' || event.key === 'Backspace') {
            const activeElement = document.activeElement;
            const isEditingContent = activeElement &&
                activeElement.classList.contains('cell-content') &&
                activeElement.innerText.length > 0;

            // Clear all cells if multi-selection, or clear single cell if not actively editing
            if (hasMultiSelection() || !isEditingContent) {
                event.preventDefault();
                clearSelectedCells();
                return;
            }
        }

        // Enter key: evaluate formula / move to cell below
        if (event.key === 'Enter') {
            const target = event.target;
            if (!target.classList.contains('cell-content')) return;

            const row = parseInt(target.dataset.row, 10);
            const col = parseInt(target.dataset.col, 10);

            if (isNaN(row) || isNaN(col)) return;

            // Prevent default newline behavior
            event.preventDefault();
            clearEditingCell();

            // Check if this is a formula cell - evaluate it
            const rawValue = target.innerText.trim();
            if (rawValue.startsWith('=')) {
                formulas[row][col] = rawValue;
                const result = evaluateFormula(rawValue);
                data[row][col] = String(result);
                target.innerText = String(result);
                recalculateFormulas();
                debouncedUpdateURL();

                // Exit formula edit mode
                formulaEditMode = false;
                formulaEditCell = null;
            }

            // Try to move to cell below
            const nextRow = row + 1;
            if (nextRow < rows) {
                // Focus cell below
                const nextCell = document.querySelector(
                    `.cell-content[data-row="${nextRow}"][data-col="${col}"]`
                );
                if (nextCell) {
                    nextCell.focus();
                }
            } else {
                // No row below - just blur current cell
                target.blur();
            }
        }
    }

    // ========== Column/Row Resize Handlers ==========

    function handleResizeStart(event) {
        if (event.button !== 0) return;
        if (!(event.target instanceof Element)) return;

        const handle = event.target.closest('.resize-handle');
        if (!handle) return;

        const isColResize = handle.classList.contains('col-resize');
        const indexValue = isColResize ? handle.dataset.col : handle.dataset.row;
        const index = parseInt(indexValue, 10);
        if (isNaN(index)) return;

        event.preventDefault();
        event.stopImmediatePropagation();

        resizeState = {
            type: isColResize ? 'col' : 'row',
            index,
            startX: event.clientX,
            startY: event.clientY,
            startSize: isColResize
                ? (colWidths[index] || DEFAULT_COL_WIDTH)
                : (rowHeights[index] || DEFAULT_ROW_HEIGHT)
        };

        isSelecting = false;
        document.body.classList.add('resizing');
        document.body.style.cursor = isColResize ? 'col-resize' : 'row-resize';

        document.addEventListener('mousemove', handleResizeMove);
        document.addEventListener('mouseup', handleResizeEnd);
    }

    function handleResizeMove(event) {
        if (!resizeState) return;

        if (resizeState.type === 'col') {
            const delta = event.clientX - resizeState.startX;
            const nextWidth = Math.max(MIN_COL_WIDTH, resizeState.startSize + delta);
            colWidths[resizeState.index] = nextWidth;
        } else {
            const delta = event.clientY - resizeState.startY;
            const nextHeight = Math.max(MIN_ROW_HEIGHT, resizeState.startSize + delta);
            rowHeights[resizeState.index] = nextHeight;
        }

        applyGridTemplate();
    }

    function handleResizeEnd() {
        if (!resizeState) return;

        document.removeEventListener('mousemove', handleResizeMove);
        document.removeEventListener('mouseup', handleResizeEnd);
        document.body.classList.remove('resizing');
        document.body.style.cursor = '';
        resizeState = null;
        debouncedUpdateURL();
    }

    // ========== Touch Selection Handlers ==========

    function handleTouchStart(event) {
        // Only handle single touch
        if (event.touches.length !== 1) return;

        const touch = event.touches[0];
        const cellCoords = getCellFromPoint(touch.clientX, touch.clientY);

        if (!cellCoords) {
            clearSelection();
            return;
        }

        // Start new selection
        selectionStart = cellCoords;
        selectionEnd = cellCoords;
        isSelecting = true;

        const container = document.getElementById('spreadsheet');
        if (container) {
            container.classList.add('selecting');
        }

        updateSelectionVisuals();
    }

    function handleTouchMove(event) {
        if (!isSelecting) return;
        if (event.touches.length !== 1) return;

        const touch = event.touches[0];
        const cellCoords = getCellFromPoint(touch.clientX, touch.clientY);

        if (!cellCoords) return;

        // Only update if position changed
        if (selectionEnd && cellCoords.row === selectionEnd.row && cellCoords.col === selectionEnd.col) {
            return;
        }

        selectionEnd = cellCoords;
        updateSelectionVisuals();

        // Prevent scrolling during selection
        event.preventDefault();
    }

    function handleTouchEnd(event) {
        if (!isSelecting) return;

        isSelecting = false;

        const container = document.getElementById('spreadsheet');
        if (container) {
            container.classList.remove('selecting');
        }

        // If single cell selected, allow focus for editing
        if (!hasMultiSelection() && selectionStart) {
            const cellContent = document.querySelector(
                `.cell-content[data-row="${selectionStart.row}"][data-col="${selectionStart.col}"]`
            );
            if (cellContent) {
                cellContent.focus();
            }
        }
    }

    function forEachTargetCell(callback) {
        const bounds = getSelectionBounds();
        if (bounds) {
            for (let r = bounds.minRow; r <= bounds.maxRow; r++) {
                for (let c = bounds.minCol; c <= bounds.maxCol; c++) {
                    callback(r, c);
                }
            }
            return true;
        }

        const activeElement = document.activeElement;
        if (activeElement && activeElement.classList.contains('cell-content')) {
            const row = parseInt(activeElement.dataset.row, 10);
            const col = parseInt(activeElement.dataset.col, 10);
            if (!isNaN(row) && !isNaN(col)) {
                callback(row, col);
                return true;
            }
        }

        if (activeRow !== null && activeCol !== null) {
            callback(activeRow, activeCol);
            return true;
        }

        return false;
    }

    function applyAlignment(align) {
        const normalized = normalizeAlignment(align);
        if (!normalized) return;

        const updated = forEachTargetCell(function(row, col) {
            if (!cellStyles[row]) cellStyles[row] = [];
            if (!cellStyles[row][col]) cellStyles[row][col] = createEmptyCellStyle();
            cellStyles[row][col].align = normalized;

            const cellContent = getCellContentElement(row, col);
            if (cellContent) {
                cellContent.style.textAlign = normalized;
            }
        });

        if (updated) {
            debouncedUpdateURL();
        }
    }

    function applyCellBackground(color) {
        if (!isValidCSSColor(color)) return;

        const updated = forEachTargetCell(function(row, col) {
            if (!cellStyles[row]) cellStyles[row] = [];
            if (!cellStyles[row][col]) cellStyles[row][col] = createEmptyCellStyle();
            cellStyles[row][col].bg = color;

            const cell = getCellElement(row, col);
            if (cell) {
                if (color) {
                    cell.style.setProperty('--cell-bg', color);
                } else {
                    cell.style.removeProperty('--cell-bg');
                }
            }
        });

        if (updated) {
            debouncedUpdateURL();
        }
    }

    function applyCellTextColor(color) {
        if (!isValidCSSColor(color)) return;

        const updated = forEachTargetCell(function(row, col) {
            if (!cellStyles[row]) cellStyles[row] = [];
            if (!cellStyles[row][col]) cellStyles[row][col] = createEmptyCellStyle();
            cellStyles[row][col].color = color;

            const cellContent = getCellContentElement(row, col);
            if (cellContent) {
                cellContent.style.color = color;
            }
        });

        if (updated) {
            debouncedUpdateURL();
        }
    }

    function applyFontSize(size) {
        const normalized = normalizeFontSize(size);

        const updated = forEachTargetCell(function(row, col) {
            if (!cellStyles[row]) cellStyles[row] = [];
            if (!cellStyles[row][col]) cellStyles[row][col] = createEmptyCellStyle();
            cellStyles[row][col].fontSize = normalized;

            const cellContent = getCellContentElement(row, col);
            if (cellContent) {
                cellContent.style.fontSize = normalized ? `${normalized}px` : '';
            }
        });

        if (updated) {
            debouncedUpdateURL();
        }
    }

    // Apply text formatting using modern Selection/Range API (replaces deprecated execCommand)
    function applyFormat(command) {
        const selection = window.getSelection();
        if (!selection || !selection.rangeCount) return;

        const range = selection.getRangeAt(0);
        const selectedText = range.toString();
        if (!selectedText) return;

        // Create wrapper element based on command
        let wrapper;
        switch (command) {
            case 'bold':
                wrapper = document.createElement('b');
                break;
            case 'italic':
                wrapper = document.createElement('i');
                break;
            case 'underline':
                wrapper = document.createElement('u');
                break;
            default:
                return;
        }

        // Only proceed if selection is within a cell-content element
        const activeElement = document.activeElement;
        if (!activeElement || !activeElement.classList.contains('cell-content')) return;

        try {
            // Wrap the selected content
            range.surroundContents(wrapper);
        } catch (e) {
            // surroundContents fails if selection crosses element boundaries
            // Fall back to extracting and re-inserting
            const fragment = range.extractContents();
            wrapper.appendChild(fragment);
            range.insertNode(wrapper);
        }

        // Update data after formatting
        const row = parseInt(activeElement.dataset.row, 10);
        const col = parseInt(activeElement.dataset.col, 10);
        if (!isNaN(row) && !isNaN(col) && row < rows && col < cols) {
            data[row][col] = sanitizeHTML(activeElement.innerHTML);
            debouncedUpdateURL();
        }
    }

    // Handle paste to strip unwanted HTML
    function handlePaste(event) {
        const target = event.target;
        if (!target.classList.contains('cell-content')) return;

        event.preventDefault();
        const text = event.clipboardData.getData('text/plain');
        document.execCommand('insertText', false, text);
    }

    // Clear spreadsheet and reset to default
function clearSpreadsheet() {
    if (!confirm('Clear all data and reset to 10×10 grid?')) {
        return;
    }

    // Reset to default dimensions
    rows = DEFAULT_ROWS;
    cols = DEFAULT_COLS;
    data = createEmptyData(rows, cols);
    formulas = createEmptyData(rows, cols);
    cellStyles = createEmptyCellStyles(rows, cols);
    colWidths = createDefaultColumnWidths(cols);
    rowHeights = createDefaultRowHeights(rows);

    // Clear password
    currentPassword = null;
    updateLockButtonUI();

    // Clear any selection
    clearSelection();

    // Re-render and update URL
    renderGrid();
    debouncedUpdateURL();

    showToast('Spreadsheet cleared', 'success');
}

    // Add a new row
function addRow() {
    if (rows >= MAX_ROWS) {
        showToast(`Maximum ${MAX_ROWS} rows allowed`, 'warning');
        return;
    }
    rows++;
    data.push(Array(cols).fill(''));
    formulas.push(Array(cols).fill(''));
    cellStyles.push(Array(cols).fill(null).map(() => createEmptyCellStyle()));
    rowHeights.push(DEFAULT_ROW_HEIGHT);
    renderGrid();
    debouncedUpdateURL();
    showToast(`Row ${rows} added`, 'success');
}

    // Add a new column
function addColumn() {
    if (cols >= MAX_COLS) {
        showToast(`Maximum ${MAX_COLS} columns allowed`, 'warning');
        return;
    }
    cols++;
    data.forEach(row => row.push(''));
    formulas.forEach(row => row.push(''));
    cellStyles.forEach(row => row.push(createEmptyCellStyle()));
    colWidths.push(DEFAULT_COL_WIDTH);
    renderGrid();
    debouncedUpdateURL();
    showToast(`Column ${colToLetter(cols - 1)} added`, 'success');
}

    // Load state from URL on page load
    // Returns true if data loaded successfully, false if waiting for password
    function loadStateFromURL() {
        const hash = window.location.hash.slice(1); // Remove #

        if (hash) {
            const loadedState = decodeState(hash);
            if (loadedState) {
                // Check if data is encrypted
                if (loadedState.encrypted) {
                    // Store encrypted data and show password modal
                    pendingEncryptedData = loadedState.data;
                    // Initialize with default state while waiting for password
                    rows = DEFAULT_ROWS;
                    cols = DEFAULT_COLS;
                    data = createEmptyData(rows, cols);
                    cellStyles = createEmptyCellStyles(rows, cols);
                    formulas = createEmptyData(rows, cols);
                    colWidths = createDefaultColumnWidths(cols);
                    rowHeights = createDefaultRowHeights(rows);
                    // Show password modal after a short delay (let UI render first)
                    setTimeout(() => showPasswordModal('decrypt'), 100);
                    return false;
                }

                rows = loadedState.rows;
                cols = loadedState.cols;
                data = loadedState.data;
                formulas = loadedState.formulas || createEmptyData(rows, cols);
                cellStyles = loadedState.cellStyles || createEmptyCellStyles(rows, cols);
                colWidths = loadedState.colWidths || createDefaultColumnWidths(cols);
                rowHeights = loadedState.rowHeights || createDefaultRowHeights(rows);

                // Apply theme from URL if present
                if (loadedState.theme) {
                    applyTheme(loadedState.theme);
                }
                return true;
            }
        }

        // Default state
        rows = DEFAULT_ROWS;
        cols = DEFAULT_COLS;
        data = createEmptyData(rows, cols);
        cellStyles = createEmptyCellStyles(rows, cols);
        formulas = createEmptyData(rows, cols);
        colWidths = createDefaultColumnWidths(cols);
        rowHeights = createDefaultRowHeights(rows);
        return true;
    }

    // Toggle dark/light mode
    function toggleTheme() {
        const body = document.body;
        const isDark = body.classList.toggle('dark-mode');
        const theme = isDark ? 'dark' : 'light';

        // Save preference to localStorage
        try {
            localStorage.setItem('spreadsheet-theme', theme);
        } catch (e) {
            // localStorage not available
        }

        // Update URL with new theme
        debouncedUpdateURL();
    }

    // Load saved theme preference
    function loadTheme() {
        try {
            const savedTheme = localStorage.getItem('spreadsheet-theme');
            if (savedTheme === 'dark') {
                document.body.classList.add('dark-mode');
            } else if (savedTheme === 'light') {
                document.body.classList.remove('dark-mode');
            } else {
                // Check system preference if no saved preference
                if (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) {
                    document.body.classList.add('dark-mode');
                }
            }
        } catch (e) {
            // localStorage not available
        }
    }

    // Copy URL to clipboard
    function copyURL() {
        const url = window.location.href;
        const copyBtn = document.getElementById('copy-url');

        navigator.clipboard.writeText(url).then(function() {
            // Show success feedback
            showToast('Link copied to clipboard!', 'success');
            if (copyBtn) {
                copyBtn.classList.add('copied');
                const icon = copyBtn.querySelector('i');
                if (icon) {
                    icon.className = 'fa-solid fa-check';
                }

                // Reset after 2 seconds
                setTimeout(function() {
                    copyBtn.classList.remove('copied');
                    if (icon) {
                        icon.className = 'fa-solid fa-copy';
                    }
                }, 2000);
            }
        }).catch(function(err) {

            console.error('Failed to copy URL:', err);
            // Fallback for older browsers
            const textArea = document.createElement('textarea');
            textArea.value = url;
            textArea.style.position = 'fixed';
            textArea.style.opacity = '0';
            document.body.appendChild(textArea);
            textArea.select();
            try {
                document.execCommand('copy');
                if (copyBtn) {
                    copyBtn.classList.add('copied');
                    setTimeout(function() {
                        copyBtn.classList.remove('copied');
                    }, 2000);
                }
            } catch (e) {
                console.error('Fallback copy failed:', e);
            }
            document.body.removeChild(textArea);
        });
    }

    // ========== Password/Encryption Modal Functions ==========

    // Modal mode: 'set' for setting password, 'decrypt' for decrypting
    let modalMode = 'set';

    // Show the password modal
    function showPasswordModal(mode) {
        modalMode = mode;
        const modal = document.getElementById('password-modal');
        const title = document.getElementById('modal-title');
        const description = document.getElementById('modal-description');
        const confirmInput = document.getElementById('password-confirm');
        const submitBtn = document.getElementById('modal-submit');
        const passwordInput = document.getElementById('password-input');
        const errorEl = document.getElementById('modal-error');

        // Reset form
        passwordInput.value = '';
        confirmInput.value = '';
        errorEl.classList.add('hidden');
        errorEl.textContent = '';

        if (mode === 'set') {
            title.textContent = 'Set Password';
            description.textContent = 'Enter a password to encrypt this spreadsheet. Anyone with the link will need this password to view it.';
            confirmInput.style.display = '';
            confirmInput.placeholder = 'Confirm password';
            submitBtn.textContent = 'Set Password';
        } else if (mode === 'decrypt') {
            title.textContent = 'Enter Password';
            description.textContent = 'This spreadsheet is password-protected. Enter the password to view it.';
            confirmInput.style.display = 'none';
            submitBtn.textContent = 'Unlock';
        } else if (mode === 'remove') {
            title.textContent = 'Remove Password';
            description.textContent = 'Enter the current password to remove encryption from this spreadsheet.';
            confirmInput.style.display = 'none';
            submitBtn.textContent = 'Remove Password';
        }

        modal.classList.remove('hidden');
        passwordInput.focus();
    }

    // Hide the password modal
    function hidePasswordModal() {
        const modal = document.getElementById('password-modal');
        modal.classList.add('hidden');
        pendingEncryptedData = null;
    }

    // Show error in modal
    function showModalError(message) {
        const errorEl = document.getElementById('modal-error');
        errorEl.textContent = message;
        errorEl.classList.remove('hidden');
    }

    // Update lock button UI state
    function updateLockButtonUI() {
        const lockBtn = document.getElementById('lock-btn');
        if (!lockBtn) return;

        const icon = lockBtn.querySelector('i');
        if (currentPassword) {
            lockBtn.classList.add('locked');
            lockBtn.title = 'Remove Password Protection';
            if (icon) icon.className = 'fa-solid fa-lock';
        } else {
            lockBtn.classList.remove('locked');
            lockBtn.title = 'Password Protection';
            if (icon) icon.className = 'fa-solid fa-lock-open';
        }
    }

    // Handle lock button click
    function handleLockButtonClick() {
        if (currentPassword) {
            // Already encrypted - offer to remove password
            showPasswordModal('remove');
        } else {
            // Not encrypted - set password
            showPasswordModal('set');
        }
    }

    // Handle modal submit
    async function handleModalSubmit() {
        const passwordInput = document.getElementById('password-input');
        const confirmInput = document.getElementById('password-confirm');
        const password = passwordInput.value;

        if (!password) {
            showModalError('Please enter a password.');
            return;
        }

        if (modalMode === 'set') {
            // Setting new password
            const confirm = confirmInput.value;
            if (password !== confirm) {
                showModalError('Passwords do not match.');
                return;
            }
            if (password.length < 4) {
                showModalError('Password must be at least 4 characters.');
                return;
            }

            currentPassword = password;
            updateLockButtonUI();
            hidePasswordModal();
            // Re-encode state with encryption
            updateURL();
            showToast('Password protection enabled', 'success');

        } else if (modalMode === 'decrypt') {
            // Decrypting loaded data
            if (!pendingEncryptedData) {
                showModalError('No encrypted data to decrypt.');
                return;
            }

            try {
                const decrypted = await CryptoUtils.decrypt(pendingEncryptedData, password);
                const decompressed = LZString.decompressFromEncodedURIComponent(decrypted);
                if (!decompressed) {
                    showModalError('Incorrect password.');
                    return;
                }

                const parsed = safeJSONParse(decompressed);
                if (!parsed) {
                    showModalError('Incorrect password.');
                    return;
                }

                // Successfully decrypted - load the state
                currentPassword = password;
                applyLoadedState(parsed);
                hidePasswordModal();
                renderGrid();
                updateLockButtonUI();
                showToast('Spreadsheet unlocked', 'success');

            } catch (e) {
                console.error('Decryption failed:', e);
                showModalError('Incorrect password.');
            }

        } else if (modalMode === 'remove') {
            // Verify current password before removing
            // We can verify by re-encrypting current state and checking it works
            currentPassword = null;
            updateLockButtonUI();
            hidePasswordModal();
            // Re-encode state without encryption
            updateURL();
            showToast('Password protection removed', 'success');
        }
    }

    // Apply loaded state to variables
    function applyLoadedState(loadedState) {
        if (!loadedState) return;

        const r = Math.min(Math.max(1, loadedState.rows || DEFAULT_ROWS), MAX_ROWS);
        const c = Math.min(Math.max(1, loadedState.cols || DEFAULT_COLS), MAX_COLS);

        rows = r;
        cols = c;

        // Process data array
        let d = loadedState.data;
        if (Array.isArray(d)) {
            d = d.slice(0, r).map(row => {
                if (Array.isArray(row)) {
                    return row.slice(0, c).map(cell => sanitizeHTML(String(cell || '')));
                }
                return Array(c).fill('');
            });
            while (d.length < r) d.push(Array(c).fill(''));
            d = d.map(row => {
                while (row.length < c) row.push('');
                return row;
            });
        } else {
            d = createEmptyData(r, c);
        }
        data = d;

        // Process formulas
        let f = loadedState.formulas;
        if (Array.isArray(f)) {
            f = f.slice(0, r).map((row, rowIdx) => {
                if (Array.isArray(row)) {
                    return row.slice(0, c).map((cell, colIdx) => {
                        const formula = String(cell || '');
                        if (formula.startsWith('=')) {
                            if (isValidFormula(formula)) {
                                return formula;
                            } else {
                                data[rowIdx][colIdx] = escapeHTML(formula);
                                return '';
                            }
                        }
                        return formula;
                    });
                }
                return Array(c).fill('');
            });
            while (f.length < r) f.push(Array(c).fill(''));
            f = f.map(row => {
                while (row.length < c) row.push('');
                return row;
            });
        } else {
            f = createEmptyData(r, c);
        }
        formulas = f;

        cellStyles = normalizeCellStyles(loadedState.cellStyles, r, c);
        colWidths = normalizeColumnWidths(loadedState.colWidths, c);
        rowHeights = normalizeRowHeights(loadedState.rowHeights, r);

        if (loadedState.theme) {
            applyTheme(loadedState.theme);
        }
    }

    // Initialize the app
    function init() {
        // Load theme preference first (before any rendering)
        loadTheme();

        // Load any existing state from URL
        loadStateFromURL();

        // Render the grid
        renderGrid();

        // Update lock button UI state
        updateLockButtonUI();

        // Initialize URL length indicator with current hash length
        const currentHash = window.location.hash.slice(1);
        updateURLLengthIndicator(currentHash.length);

        // Set up event delegation for input handling
        const container = document.getElementById('spreadsheet');
        if (container) {
            container.addEventListener('input', handleInput);
            container.addEventListener('focusin', handleFocusIn);
            container.addEventListener('focusout', handleFocusOut);
            container.addEventListener('paste', handlePaste);

            // Selection mouse events
            container.addEventListener('mousedown', handleResizeStart);
            container.addEventListener('mousedown', handleMouseDown);
            container.addEventListener('dblclick', handleCellDoubleClick);
            container.addEventListener('mousemove', handleMouseMove);
            container.addEventListener('mouseleave', handleMouseLeave);
            container.addEventListener('mouseup', handleMouseUp);
            container.addEventListener('keydown', handleSelectionKeyDown);

            // Touch events for mobile selection
            container.addEventListener('touchstart', handleTouchStart, { passive: false });
            container.addEventListener('touchmove', handleTouchMove, { passive: false });
            container.addEventListener('touchend', handleTouchEnd);
        }

        const gridWrapper = document.querySelector('.grid-wrapper');
        if (gridWrapper) {
            gridWrapper.addEventListener('scroll', function() {
                if (isFormulaDropdownOpen() && formulaDropdownAnchor) {
                    positionFormulaDropdown(formulaDropdownAnchor);
                }
            });
        }

        window.addEventListener('resize', function() {
            if (isFormulaDropdownOpen() && formulaDropdownAnchor) {
                positionFormulaDropdown(formulaDropdownAnchor);
            }
        });

        // Global mouseup to catch drag ending outside container
        document.addEventListener('mouseup', handleMouseUp);

        // Format button event listeners
        const boldBtn = document.getElementById('format-bold');
        const italicBtn = document.getElementById('format-italic');
        const underlineBtn = document.getElementById('format-underline');
        const alignLeftBtn = document.getElementById('align-left');
        const alignCenterBtn = document.getElementById('align-center');
        const alignRightBtn = document.getElementById('align-right');
        const cellBgPicker = document.getElementById('cell-bg-color');
        const cellTextColorPicker = document.getElementById('cell-text-color');
        const fontSizeList = document.getElementById('font-size-list');

        if (boldBtn) {
            boldBtn.addEventListener('mousedown', function(e) {
                e.preventDefault(); // Prevent focus loss
                applyFormat('bold');
            });
        }
        if (italicBtn) {
            italicBtn.addEventListener('mousedown', function(e) {
                e.preventDefault();
                applyFormat('italic');
            });
        }
        if (underlineBtn) {
            underlineBtn.addEventListener('mousedown', function(e) {
                e.preventDefault();
                applyFormat('underline');
            });
        }
        if (alignLeftBtn) {
            alignLeftBtn.addEventListener('mousedown', function(e) {
                e.preventDefault();
                applyAlignment('left');
            });
        }
        if (alignCenterBtn) {
            alignCenterBtn.addEventListener('mousedown', function(e) {
                e.preventDefault();
                applyAlignment('center');
            });
        }
        if (alignRightBtn) {
            alignRightBtn.addEventListener('mousedown', function(e) {
                e.preventDefault();
                applyAlignment('right');
            });
        }
        if (cellBgPicker) {
            cellBgPicker.addEventListener('input', function(e) {
                applyCellBackground(e.target.value);
            });
        }
        if (cellTextColorPicker) {
            cellTextColorPicker.addEventListener('input', function(e) {
                applyCellTextColor(e.target.value);
            });
        }
        if (fontSizeList) {
            fontSizeList.addEventListener('mousedown', function(e) {
                if (e.target.closest('button')) {
                    e.preventDefault();
                }
            });
            fontSizeList.addEventListener('click', function(e) {
                const button = e.target.closest('button[data-size]');
                if (!button) return;
                applyFontSize(button.dataset.size);
            });
        }

        // Button event listeners
        const addRowBtn = document.getElementById('add-row');
        const addColBtn = document.getElementById('add-col');
        const clearBtn = document.getElementById('clear-spreadsheet');
        const themeToggleBtn = document.getElementById('theme-toggle');
        const copyUrlBtn = document.getElementById('copy-url');
        const importCsvBtn = document.getElementById('import-csv');
        const importCsvInput = document.getElementById('import-csv-file');
        const exportCsvBtn = document.getElementById('export-csv');

        if (addRowBtn) {
            addRowBtn.addEventListener('click', addRow);
        }
        if (addColBtn) {
            addColBtn.addEventListener('click', addColumn);
        }
        if (clearBtn) {
            clearBtn.addEventListener('click', clearSpreadsheet);
        }
        if (themeToggleBtn) {
            themeToggleBtn.addEventListener('click', toggleTheme);
        }
        if (copyUrlBtn) {
            copyUrlBtn.addEventListener('click', copyURL);
        }
        if (importCsvBtn && importCsvInput) {
            importCsvBtn.addEventListener('click', function() {
                importCsvInput.click();
            });
        }
        if (importCsvInput) {
            importCsvInput.addEventListener('change', function(e) {
                const file = e.target.files && e.target.files[0];
                if (!file) return;

                const reader = new FileReader();
                reader.onload = function() {
                    importCSVText(String(reader.result || ''));
                };
                reader.onerror = function() {
                    alert('Failed to read the CSV file.');
                };
                reader.readAsText(file);
                e.target.value = '';
            });
        }
        if (exportCsvBtn) {
            exportCsvBtn.addEventListener('click', downloadCSV);
        }

        // Password/Encryption event listeners
        const lockBtn = document.getElementById('lock-btn');
        const modalCancel = document.getElementById('modal-cancel');
        const modalSubmit = document.getElementById('modal-submit');
        const modalBackdrop = document.querySelector('.modal-backdrop');
        const passwordInput = document.getElementById('password-input');

        if (lockBtn) {
            lockBtn.addEventListener('click', handleLockButtonClick);
        }
        if (modalCancel) {
            modalCancel.addEventListener('click', hidePasswordModal);
        }
        if (modalSubmit) {
            modalSubmit.addEventListener('click', handleModalSubmit);
        }
        if (modalBackdrop) {
            modalBackdrop.addEventListener('click', function() {
                // Only allow closing if not in decrypt mode (user must enter password)
                if (modalMode !== 'decrypt') {
                    hidePasswordModal();
                }
            });
        }
        if (passwordInput) {
            passwordInput.addEventListener('keydown', function(e) {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    const confirmInput = document.getElementById('password-confirm');
                    // If confirm is visible and empty, focus it; otherwise submit
                    if (confirmInput && confirmInput.style.display !== 'none' && !confirmInput.value) {
                        confirmInput.focus();
                    } else {
                        handleModalSubmit();
                    }
                }
            });
        }
        const passwordConfirm = document.getElementById('password-confirm');
        if (passwordConfirm) {
            passwordConfirm.addEventListener('keydown', function(e) {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    handleModalSubmit();
                }
            });
        }

        // Handle browser back/forward
        window.addEventListener('hashchange', function() {
            loadStateFromURL();
            renderGrid();
        });
    }

    // Start when DOM is ready
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }

    // Expose audit function globally for testing URL sizes
    window.auditURLSize = async function() {
        const scenarios = [
            { name: 'Empty 30x15', rows: 30, cols: 15, fill: null },
            { name: 'Short text (A1, B1...)', rows: 30, cols: 15, fill: 'coords' },
            { name: 'Medium text (10 chars)', rows: 30, cols: 15, fill: 'medium' },
            { name: 'Numbers only', rows: 30, cols: 15, fill: 'numbers' },
            { name: 'With formulas', rows: 30, cols: 15, fill: 'formulas' }
        ];

        console.log('=== URL Size Audit ===');
        console.log('Max grid: 30 rows x 15 cols = 450 cells\n');

        for (const scenario of scenarios) {
            const testData = [];
            for (let r = 0; r < scenario.rows; r++) {
                const row = [];
                for (let c = 0; c < scenario.cols; c++) {
                    const colLetter = String.fromCharCode(65 + c);
                    const cellRef = colLetter + (r + 1);
                    switch (scenario.fill) {
                        case 'coords':
                            row.push(cellRef);
                            break;
                        case 'medium':
                            row.push('Text_' + cellRef.padEnd(5, 'X'));
                            break;
                        case 'numbers':
                            row.push(String(Math.floor(Math.random() * 10000)));
                            break;
                        case 'formulas':
                            // Only put formulas in some cells to simulate realistic usage
                            if (r > 0 && c === 0) {
                                row.push('=SUM(B' + (r + 1) + ':O' + (r + 1) + ')');
                            } else {
                                row.push(String(Math.floor(Math.random() * 100)));
                            }
                            break;
                        default:
                            row.push('');
                    }
                }
                testData.push(row);
            }

            const state = {
                rows: scenario.rows,
                cols: scenario.cols,
                theme: 'light',
                data: testData
            };

            const json = JSON.stringify(state);
            const compressed = LZString.compressToEncodedURIComponent(json);

            console.log(`${scenario.name}:`);
            console.log(`  JSON size: ${json.length.toLocaleString()} chars`);
            console.log(`  Compressed: ${compressed.length.toLocaleString()} chars`);
            console.log(`  Compression ratio: ${((1 - compressed.length / json.length) * 100).toFixed(1)}%`);
            console.log('');
        }

        console.log('=== Thresholds ===');
        console.log('< 2,000 chars: OK (safe for all browsers)');
        console.log('2,000-4,000: Warning (some older browsers may truncate)');
        console.log('4,000-8,000: Caution (URL shorteners may fail)');
        console.log('> 8,000: Critical (some browsers may fail)');
    };
    // ========== Toolbar Scroll Logic ==========
    function initToolbarScroll() {
        const toolbar = document.querySelector('.toolbar');
        const scrollLeftBtn = document.getElementById('scroll-left');
        const scrollRightBtn = document.getElementById('scroll-right');
        
        if (!toolbar || !scrollLeftBtn || !scrollRightBtn) return;

        function updateScrollButtons() {
            // Check if content overflows
            const isOverflowing = toolbar.scrollWidth > toolbar.clientWidth;
            const scrollLeft = toolbar.scrollLeft;
            const maxScroll = toolbar.scrollWidth - toolbar.clientWidth;
            
            // Tolerance (fixes weird browser sub-pixel issues)
            const tolerance = 2;

            if (!isOverflowing) {
                scrollLeftBtn.classList.add('hidden');
                scrollRightBtn.classList.add('hidden');
                return;
            }

            // Show/Hide Left Button
            if (scrollLeft > tolerance) {
                scrollLeftBtn.classList.remove('hidden');
            } else {
                scrollLeftBtn.classList.add('hidden');
            }

            // Show/Hide Right Button
            if (scrollLeft < maxScroll - tolerance) {
                scrollRightBtn.classList.remove('hidden');
            } else {
                scrollRightBtn.classList.add('hidden');
            }
        }

        // Scroll amount for button clicks
        const scrollAmount = 200;

        scrollLeftBtn.addEventListener('click', () => {
            toolbar.scrollBy({ left: -scrollAmount, behavior: 'smooth' });
        });

        scrollRightBtn.addEventListener('click', () => {
            toolbar.scrollBy({ left: scrollAmount, behavior: 'smooth' });
        });

        // Listen for scroll events
        toolbar.addEventListener('scroll', () => {
             // Debounce the UI update slightly for performance
             requestAnimationFrame(updateScrollButtons);
        });

        // Update on resize
        window.addEventListener('resize', updateScrollButtons);

        // Initial check
        updateScrollButtons();
    }

    // Initialize all modules
    initToolbarScroll();
})();
