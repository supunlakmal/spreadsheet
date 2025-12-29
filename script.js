// Dynamic Spreadsheet Web App
// Data persists in URL hash for easy sharing

(function() {
    'use strict';

    // Limits
    const MAX_ROWS = 30;
    const MAX_COLS = 15;
    const DEBOUNCE_DELAY = 200;
    const ACTIVE_HEADER_CLASS = 'header-active';

    // Default starting size
    const DEFAULT_ROWS = 10;
    const DEFAULT_COLS = 10;

    // Dynamic dimensions
    let rows = DEFAULT_ROWS;
    let cols = DEFAULT_COLS;

    // Data model - dynamic 2D array
    let data = createEmptyData(rows, cols);

    // Debounce timer
    let debounceTimer = null;

    // Active header tracking for row/column highlight
    let activeRow = null;
    let activeCol = null;

    // Multi-cell selection state
    let selectionStart = null;  // { row, col } anchor point
    let selectionEnd = null;    // { row, col } current end
    let isSelecting = false;    // true during mouse drag

    // Create empty data array with specified dimensions
    function createEmptyData(r, c) {
        return Array(r).fill(null).map(() => Array(c).fill(''));
    }

    // Convert column index to letter (0 = A, 1 = B, ... 25 = Z)
    function colToLetter(col) {
        return String.fromCharCode(65 + col);
    }

    // Get current theme
    function isDarkMode() {
        return document.body.classList.contains('dark-mode');
    }

    // Encode state to URL-safe string (includes dimensions and theme)
    function encodeState() {
        const state = { rows, cols, data, theme: isDarkMode() ? 'dark' : 'light' };
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

                return { rows: r, cols: c, data: d, theme: parsed.theme || null };
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
                return { rows: DEFAULT_ROWS, cols: DEFAULT_COLS, data: d, theme: null };
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

    // Render the spreadsheet grid
    function renderGrid() {
        const container = document.getElementById('spreadsheet');
        if (!container) return;

        // Clear selection when grid is re-rendered (e.g., adding rows/columns)
        selectionStart = null;
        selectionEnd = null;
        isSelecting = false;

        // Set dynamic grid columns
        container.style.gridTemplateColumns = `40px repeat(${cols}, minmax(80px, 1fr))`;
        container.innerHTML = '';

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
            container.appendChild(header);
        }

        // Rows
        for (let row = 0; row < rows; row++) {
            // Row header (1, 2, 3...) - sticky to left
            const rowHeader = document.createElement('div');
            rowHeader.className = 'header-cell row-header';
            rowHeader.textContent = row + 1;
            rowHeader.dataset.row = row;
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

        // Clear multi-selection when user starts typing
        if (hasMultiSelection()) {
            clearSelection();
            setActiveHeaders(row, col);
        }

        if (!isNaN(row) && !isNaN(col) && row < rows && col < cols) {
            data[row][col] = target.innerHTML;
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
    }

    function handleFocusOut(event) {
        const target = event.target;
        if (!target.classList.contains('cell-content')) return;

        const container = document.getElementById('spreadsheet');
        if (!container) return;

        const next = event.relatedTarget;
        if (next && container.contains(next)) {
            return;
        }

        clearActiveHeaders();
    }

    // ========== Mouse Selection Handlers ==========

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

        if (container) {
            container.classList.add('selecting');
        }

        updateSelectionVisuals();
    }

    function handleMouseMove(event) {
        if (!isSelecting) return;

        const cellCoords = getCellFromPoint(event.clientX, event.clientY);
        if (!cellCoords) return;

        // Only update if position changed
        if (selectionEnd && cellCoords.row === selectionEnd.row && cellCoords.col === selectionEnd.col) {
            return;
        }

        selectionEnd = cellCoords;
        updateSelectionVisuals();

        // Prevent text selection during drag
        event.preventDefault();
    }

    function handleMouseUp(event) {
        if (!isSelecting) return;

        isSelecting = false;

        const container = document.getElementById('spreadsheet');
        if (container) {
            container.classList.remove('selecting');
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
        // Escape key clears selection
        if (event.key === 'Escape' && hasMultiSelection()) {
            clearSelection();
            event.preventDefault();
        }
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
        renderGrid();
        debouncedUpdateURL();
    }

    // Add a new column
    function addColumn() {
        if (cols >= MAX_COLS) return;
        cols++;
        data.forEach(row => row.push(''));
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
            container.addEventListener('mousedown', handleMouseDown);
            container.addEventListener('mousemove', handleMouseMove);
            container.addEventListener('mouseup', handleMouseUp);
            container.addEventListener('keydown', handleSelectionKeyDown);

            // Touch events for mobile selection
            container.addEventListener('touchstart', handleTouchStart, { passive: false });
            container.addEventListener('touchmove', handleTouchMove, { passive: false });
            container.addEventListener('touchend', handleTouchEnd);
        }

        // Global mouseup to catch drag ending outside container
        document.addEventListener('mouseup', handleMouseUp);

        // Format button event listeners
        const boldBtn = document.getElementById('format-bold');
        const italicBtn = document.getElementById('format-italic');
        const underlineBtn = document.getElementById('format-underline');

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

        // Button event listeners
        const addRowBtn = document.getElementById('add-row');
        const addColBtn = document.getElementById('add-col');
        const clearBtn = document.getElementById('clear-spreadsheet');
        const themeToggleBtn = document.getElementById('theme-toggle');
        const copyUrlBtn = document.getElementById('copy-url');

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
