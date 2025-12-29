# Spreadsheet

A lightweight, client-only spreadsheet web application. All data persists in the URL hash for instant sharing—no backend required.

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Vanilla JS](https://img.shields.io/badge/vanilla-JavaScript-yellow.svg)
![No Dependencies](https://img.shields.io/badge/dependencies-none-green.svg)

## Features

### Core Functionality
- **Zero Backend** - All state saved in URL hash, works completely offline
- **Instant Sharing** - Copy URL to share your spreadsheet with anyone
- **Dynamic Grid** - Expandable up to 30 rows and 15 columns (A-O)
- **Persistent State** - Browser back/forward buttons restore previous states

### Text Formatting
- **Bold** (Ctrl+B) - Apply bold formatting to selected text
- **Italic** (Ctrl+I) - Apply italic formatting to selected text
- **Underline** (Ctrl+U) - Apply underline formatting to selected text
- HTML-based formatting preserved in cell content

### Multi-Cell Selection (Google Sheets Style)
- **Click & Drag** - Select rectangular ranges by dragging
- **Shift+Click** - Extend selection from anchor point
- **Visual Feedback** - Selected cells highlighted with blue background
- **Border Outline** - Blue border around selection edges
- **Header Highlighting** - Row/column headers highlight for selected range
- **Escape to Clear** - Press Escape to deselect

### Grid Management
- **Add Row** - Expand grid rows (max 30)
- **Add Column** - Expand grid columns (max 15)
- **Clear Spreadsheet** - Reset to empty 10×10 grid with confirmation
- **Live Grid Size** - Display shows current dimensions

### Theme Support
- **Dark/Light Mode** - Toggle with sun/moon button
- **System Detection** - Respects OS dark mode preference
- **Persistent Theme** - Saved to localStorage and URL
- **Smooth Transitions** - Elegant theme switching animation

### Mobile & Touch Support
- **Touch Selection** - Drag-to-select works on touch devices
- **Responsive Design** - Adapts to all screen sizes
- **iOS Optimized** - 16px font prevents auto-zoom
- **Touch-Friendly** - Larger tap targets on mobile

### Accessibility
- **ARIA Labels** - Screen reader support (e.g., "Cell A1")
- **Keyboard Navigation** - Full keyboard access
- **Tooltips** - Descriptive hints on all controls
- **Focus Management** - Proper focus handling

## Quick Start

Open `index.html` directly in your browser, or use a local server:

```bash
npx serve .
```

## How It Works

Your spreadsheet data is stored entirely in the URL hash:

```
https://yoursite.com/spreadsheet/#{"rows":10,"cols":10,"data":[["A1","B1"],["A2","B2"]],"theme":"light"}
```

When you edit cells, the URL updates automatically (debounced at 200ms). Share the URL to share your data—no account or database needed.

## Usage

| Action | How |
|--------|-----|
| Edit cell | Click any cell and type |
| Format text | Select text, click B/I/U buttons or use Ctrl+B/I/U |
| Select range | Click and drag across cells |
| Extend selection | Shift+Click on another cell |
| Clear selection | Press Escape |
| Add row | Click "+ Row" button (max 30) |
| Add column | Click "+ Column" button (max 15) |
| Clear all | Click "Clear" button (with confirmation) |
| Share | Click copy button to copy URL |
| Toggle theme | Click sun/moon icon |

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| Ctrl+B | Bold |
| Ctrl+I | Italic |
| Ctrl+U | Underline |
| Escape | Clear selection |
| Shift+Click | Extend selection |
| Tab | Move to next cell |

## Tech Stack

- Vanilla HTML/CSS/JavaScript (no frameworks)
- CSS Grid for spreadsheet layout
- CSS Custom Properties for theming
- Font Awesome 6.5.1 (icons via CDN)
- No build tools required

## Browser Support

| Browser | Support |
|---------|---------|
| Chrome | Full |
| Firefox | Full |
| Safari | Full |
| Edge | Full |
| Mobile Safari | Full |
| Chrome Android | Full |
| Older browsers | Graceful degradation |

## Limitations

- Maximum 30 rows
- Maximum 15 columns (A-O)
- Default grid: 10 × 10
- URL length limits may apply for very large spreadsheets

## File Structure

```
spreadsheet/
├── index.html      # Single-page app structure
├── styles.css      # All styling including dark mode
├── script.js       # Application logic (IIFE module)
├── CLAUDE.md       # Development documentation
└── README.md       # This file
```

## Architecture

- **State Management** - `data` (2D array), `rows`, `cols` variables
- **URL Sync** - JSON serialization with debounced updates
- **Event Delegation** - All cell events handled on container
- **CSS Grid** - Dynamic column template set via JavaScript
- **Sticky Headers** - Row/column headers with z-index layering

## Development

No build process required. Edit files and refresh browser.

```bash
# Start local server
npx serve .

# Open in browser
http://localhost:3000
```

## Recent Updates

### v1.2 - Multi-Cell Selection & Clear
- Added Google Sheets-style multi-cell selection
- Click and drag to select cell ranges
- Shift+Click to extend selections
- Visual selection with border outline
- Added Clear button to reset spreadsheet
- Touch/mobile selection support

### v1.1 - Text Formatting
- Added Bold, Italic, Underline buttons
- Keyboard shortcuts (Ctrl+B/I/U)
- Active header highlighting

### v1.0 - Initial Release
- Core spreadsheet functionality
- URL-based state persistence
- Dark/light theme toggle
- Dynamic grid sizing
- Mobile responsive design

## License

MIT

## Contributing

1. Fork the repository
2. Create your feature branch
3. Commit your changes
4. Push to the branch
5. Open a Pull Request

## Links

- [GitHub Repository](https://github.com/supunlakmal/spreadsheet)
- [Live Demo](https://supunlakmal.github.io/spreadsheet)
