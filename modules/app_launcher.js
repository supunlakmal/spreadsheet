/**
 * AppLauncher Module
 * Handles the "Waffle" menu for ecosystem apps.
 */
export class AppLauncher {
    constructor() {
        this.apps = [
            {
                name: 'Hash Calendar',
                url: 'https://hash-calendar.netlify.app/',
                icon: 'fa-solid fa-calendar-days',
                color: '#1a73e8'
            },
            // Current app (Spreadsheet Live) logic could be added to disable link if active
            {
                name: 'Spreadsheet Live',
                url: 'https://spreadsheetlive.netlify.app/',
                icon: 'fa-solid fa-table-cells',
                color: '#107c41'
            }
        ];

        this.init();
    }

    init() {
        this.renderMenu();
        this.setupEventListeners();
    }

    renderMenu() {
        const menuContainer = document.getElementById('app-launcher-menu');
        if (!menuContainer) return;

        // Clear existing content just in case
        menuContainer.innerHTML = '';

        // Create the grid
        const grid = document.createElement('div');
        grid.className = 'app-launcher-grid';

        this.apps.forEach(app => {
            const appItem = document.createElement('a');
            appItem.href = app.url;
            appItem.target = '_blank';
            appItem.className = 'app-launcher-item';
            appItem.rel = 'noopener noreferrer';
            
            appItem.innerHTML = `
                <div class="app-icon-wrapper" style="background-color: ${app.color}15; color: ${app.color}">
                    <i class="${app.icon}"></i>
                </div>
                <span class="app-name">${app.name}</span>
            `;

            grid.appendChild(appItem);
        });

        menuContainer.appendChild(grid);
    }

    setupEventListeners() {
        const btn = document.getElementById('app-launcher-btn');
        const menu = document.getElementById('app-launcher-menu');

        if (!btn || !menu) return;

        // Toggle menu
        btn.addEventListener('click', (e) => {
            e.stopPropagation();
            const isHidden = menu.classList.contains('hidden');
            
            if (isHidden) {
                this.openMenu();
            } else {
                this.closeMenu();
            }
        });

        // Close on click outside
        document.addEventListener('click', (e) => {
            if (!menu.classList.contains('hidden') && !menu.contains(e.target) && !btn.contains(e.target)) {
                this.closeMenu();
            }
        });

        // Close on Escape key
        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape' && !menu.classList.contains('hidden')) {
                this.closeMenu();
            }
        });
    }

    openMenu() {
        const menu = document.getElementById('app-launcher-menu');
        const btn = document.getElementById('app-launcher-btn');
        
        menu.classList.remove('hidden');
        btn.setAttribute('aria-expanded', 'true');
    }

    closeMenu() {
        const menu = document.getElementById('app-launcher-menu');
        const btn = document.getElementById('app-launcher-btn');
        
        menu.classList.add('hidden');
        btn.setAttribute('aria-expanded', 'false');
    }
}
