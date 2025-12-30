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
    const FORMULA_SUGGESTIONS = [
        { name: 'SUM', signature: 'SUM(range)', description: 'Adds numbers in a range' },
        { name: 'AVG', signature: 'AVG(range)', description: 'Average of numbers in a range' }
    ];
    const FONT_SIZE_OPTIONS = [10, 12, 14, 16, 18, 24];

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

        const expr = formula.substring(1).trim().toUpperCase();

        // Match SUM(range)
        const sumMatch = expr.match(/^SUM\(([A-Z]+\d+:[A-Z]+\d+)\)$/);
        if (sumMatch) {
            return evaluateSUM(sumMatch[1]);
        }

        // Match AVG(range)
        const avgMatch = expr.match(/^AVG\(([A-Z]+\d+:[A-Z]+\d+)\)$/);
        if (avgMatch) {
            return evaluateAVG(avgMatch[1]);
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
                        bg: typeof cellStyle.bg === 'string' ? cellStyle.bg : '',
                        color: typeof cellStyle.color === 'string' ? cellStyle.color : '',
                        fontSize: normalizeFontSize(cellStyle.fontSize)
                    };
                }
            }
        }
        return normalized;
    }

    function extractPlainText(value) {
        if (value === null || value === undefined) return '';
        const temp = document.createElement('div');
        temp.innerHTML = String(value);
        const text = temp.textContent || '';
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
            alert('CSV file is empty.');
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
                    formulas[r][c] = raw;
                    data[r][c] = raw;
                } else {
                    data[r][c] = raw;
                }
            }
        }

        renderGrid();
        recalculateFormulas();
        debouncedUpdateURL();

        if (truncated) {
            alert('CSV exceeded the max size (30 rows x 15 columns). Extra data was truncated.');
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
    function encodeState() {
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
        return encodeURIComponent(json);
    }

    // Decode URL hash to state object
    function decodeState(hash) {
        try {
            const decoded = decodeURIComponent(hash);
            const parsed = JSON.parse(decoded);

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
                            return row.slice(0, c).map(cell => String(cell || ''));
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
                let f = parsed.formulas;
                if (Array.isArray(f)) {
                    f = f.slice(0, r).map(row => {
                        if (Array.isArray(row)) {
                            return row.slice(0, c).map(cell => String(cell || ''));
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
                        return row.slice(0, DEFAULT_COLS).map(cell => String(cell || ''));
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

    // Update URL hash without page jump
    function updateURL() {
        const encoded = encodeState();
        const newHash = '#' + encoded;

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
                contentDiv.innerHTML = data[row][col];
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
                data[row][col] = target.innerHTML;

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
        if (typeof color !== 'string') return;

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
        if (typeof color !== 'string') return;

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

    // Apply text formatting using execCommand
    function applyFormat(command) {
        document.execCommand(command, false, null);
        // Update data after formatting
        const activeElement = document.activeElement;
        if (activeElement && activeElement.classList.contains('cell-content')) {
            const row = parseInt(activeElement.dataset.row, 10);
            const col = parseInt(activeElement.dataset.col, 10);
            if (!isNaN(row) && !isNaN(col) && row < rows && col < cols) {
                data[row][col] = activeElement.innerHTML;
                debouncedUpdateURL();
            }
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

        // Clear any selection
        clearSelection();

        // Re-render and update URL
        renderGrid();
        debouncedUpdateURL();
    }

    // Add a new row
    function addRow() {
        if (rows >= MAX_ROWS) return;
        rows++;
        data.push(Array(cols).fill(''));
        formulas.push(Array(cols).fill(''));
        cellStyles.push(Array(cols).fill(null).map(() => createEmptyCellStyle()));
        rowHeights.push(DEFAULT_ROW_HEIGHT);
        renderGrid();
        debouncedUpdateURL();
    }

    // Add a new column
    function addColumn() {
        if (cols >= MAX_COLS) return;
        cols++;
        data.forEach(row => row.push(''));
        formulas.forEach(row => row.push(''));
        cellStyles.forEach(row => row.push(createEmptyCellStyle()));
        colWidths.push(DEFAULT_COL_WIDTH);
        renderGrid();
        debouncedUpdateURL();
    }

    // Load state from URL on page load
    function loadStateFromURL() {
        const hash = window.location.hash.slice(1); // Remove #

        if (hash) {
            const loadedState = decodeState(hash);
            if (loadedState) {
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
                return;
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

    // Initialize the app
    function init() {
        // Load theme preference first (before any rendering)
        loadTheme();

        // Load any existing state from URL
        loadStateFromURL();

        // Render the grid
        renderGrid();

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
})();
