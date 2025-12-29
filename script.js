// Dynamic Spreadsheet Web App
// Data persists in URL hash for easy sharing

(function() {
    'use strict';

    // Limits
    const MAX_ROWS = 100;
    const MAX_COLS = 26;
    const DEBOUNCE_DELAY = 200;

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
            container.appendChild(header);
        }

        // Rows
        for (let row = 0; row < rows; row++) {
            // Row header (1, 2, 3...) - sticky to left
            const rowHeader = document.createElement('div');
            rowHeader.className = 'header-cell row-header';
            rowHeader.textContent = row + 1;
            container.appendChild(rowHeader);

            // Data cells
            for (let col = 0; col < cols; col++) {
                const cell = document.createElement('div');
                cell.className = 'cell';

                const input = document.createElement('input');
                input.type = 'text';
                input.dataset.row = row;
                input.dataset.col = col;
                input.value = data[row][col];
                input.setAttribute('aria-label', `Cell ${colToLetter(col)}${row + 1}`);

                cell.appendChild(input);
                container.appendChild(cell);
            }
        }

        updateUI();
    }

    // Handle input changes
    function handleInput(event) {
        const input = event.target;
        if (input.tagName !== 'INPUT') return;

        const row = parseInt(input.dataset.row, 10);
        const col = parseInt(input.dataset.col, 10);

        if (!isNaN(row) && !isNaN(col) && row < rows && col < cols) {
            data[row][col] = input.value;
            debouncedUpdateURL();
        }
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
        }

        // Button event listeners
        const addRowBtn = document.getElementById('add-row');
        const addColBtn = document.getElementById('add-col');
        const themeToggleBtn = document.getElementById('theme-toggle');

        if (addRowBtn) {
            addRowBtn.addEventListener('click', addRow);
        }
        if (addColBtn) {
            addColBtn.addEventListener('click', addColumn);
        }
        if (themeToggleBtn) {
            themeToggleBtn.addEventListener('click', toggleTheme);
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
