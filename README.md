# Spreadsheet

![Spreadsheet logo](logo.png)

A lightweight, client-only spreadsheet web application. All data persists in the URL hash for instant sharing—no backend required. Optional AES-GCM password protection keeps shared links locked without a server.

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Vanilla JS](https://img.shields.io/badge/vanilla-JavaScript-yellow.svg)
![No Build Tools](https://img.shields.io/badge/build-none-brightgreen.svg)

## Features

### Core Functionality
- **Client-only** - All state saved in URL hash; no backend required
- **Compressed URL State** - LZ-String compression keeps share links short
- **Instant Sharing** - Copy URL to share your spreadsheet with anyone
- **Dynamic Grid** - Expandable up to 30 rows and 15 columns (A-O)
- **Arrow Key Navigation** - Move the active cell with arrow keys; Shift+Arrow expands selection
- **Persistent History** - Browser back/forward buttons restore previous states

### Text & Cell Styling
- **Bold** (Ctrl+B) - Apply bold formatting to selected text
- **Italic** (Ctrl+I) - Apply italic formatting to selected text
- **Underline** (Ctrl+U) - Apply underline formatting to selected text
- **Alignment** - Left/center/right alignment per cell or selection
- **Font Size** - Quick size buttons (Auto, 10-24px)
- **Cell Colors** - Background and text color pickers
- **Style Persistence** - Alignment, colors, and sizes saved in the URL
- Sanitized formatting preserved in cell content (B/I/U/STRONG/EM/SPAN with safe styles)

### Multi-Cell Selection (Google Sheets Style)
- **Click & Drag** - Select rectangular ranges by dragging
- **Shift+Click** - Extend selection from anchor point
- **Shift+Arrow** - Extend selection with the keyboard
- **Visual Feedback** - Selected cells highlighted with blue background
- **Border Outline** - Blue border around selection edges
- **Header Highlighting** - Row/column headers highlight for selected range
- **Hover Highlighting** - Row/column hover highlights for quick scanning
- **Escape to Clear** - Press Escape to deselect

### Grid Management
- **Add Row** - Expand grid rows (max 30)
- **Add Column** - Expand grid columns (max 15)
- **Resize Rows/Columns** - Drag header handles to adjust sizes
- **Clear Spreadsheet** - Reset to empty 10x10 grid with confirmation
- **Live Grid Size** - Display shows current dimensions

### Data Import/Export
- **CSV Import** - Load .csv files from the toolbar
- **CSV Export** - Download the current grid as `spreadsheet.csv`
- **Formula-aware Import** - SUM/AVG formulas are preserved; unsupported formulas are imported as text

### Password Protection & Sharing
- **One-click lock** - Set a password from the toolbar lock button; password never leaves the browser
- **AES-GCM 256 + PBKDF2** - 100k iterations with random salt/IV, stored as URL-safe Base64
- **ENC-prefixed URLs** - Encrypted hashes use `ENC:`; recipients must unlock via modal
- **Optional** - Unencrypted links continue to work exactly as before

### Formula Support
- **SUM / AVG Functions** - Calculate totals or averages with `=SUM(A1:B5)` / `=AVG(A1:B5)`
- **Formula Autocomplete** - Dropdown suggestions appear when typing `=`
- **Range Selection** - Click/drag cells while editing to insert range references
- **Live Evaluation** - Formulas evaluate on Enter or when leaving the cell
- **Recalculation** - Dependent formulas update when referenced cells change
- **Error Handling** - Shows `#REF!` for invalid ranges, `#ERROR!` for unknown formulas
- **Shareable Formulas** - Formulas preserved in URL for sharing

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
- **Keyboard Navigation** - Arrow keys move selection; formatting shortcuts supported
- **Tooltips** - Descriptive hints on all controls
- **Focus Management** - Proper focus handling

### Security & Privacy
- **Password-protected links** - AES-GCM encryption with PBKDF2 key derivation; tampering detection via GCM tag
- **HTML Sanitization** - DOMParser-based whitelist for tags and safe styles
- **Formula Validation** - Only SUM/AVG range syntax is accepted in formulas
- **Safe URL Parsing** - Prototype pollution guards and hash length checks
- **Style Guardrails** - CSS color validation for user-provided styles
- **Content Security Policy** - CSP restricts scripts/styles/imgs/fonts to trusted sources
- **Analytics Hygiene** - URL hash excluded from Google Analytics page tracking

## Quick Start

Open `index.html` directly in your browser, or use a local server:

```bash
npx serve .
```

## Demo Links

- [Open example](https://spreadsheetlive.netlify.app/#N4IgTg9g7gziBcBGAnAGhAYwgGzvAHOgC4AWApgLZkIgAmAhmANYjoNH0IDaXIACmTAwIAO3rYABABV6ADwkBhcRgCu2ekQCWoiQAoA8nwCiAOQCUrEJevobV2wF1UvO69vurT3gFlRpbACeEgCSIlhUlhISACwADLGo8bEAdPESbvaZ1l4gAJpkjIEhYRAR6FHRqPjxifGpsekeGdnOWU3tWTkAYhBgZJoA5iISAMoQKmAYZHAdzY6tJmRQ0nISAEoaZBIAtBKIAKwApHNtna0ATLHn+wD0l+cAbJF7VTVJ9Y0gUQDMtQnvaUsB2O5Ri+3+dUB6G+AHZwbFXDkTs0kbM0Z4Fr0KOIJAIhKIcTJZMjZqivi9qhCUlDyb8kn8PpZYiDybsoijWs9EH8GTSouceQCGpYHiyog8qYzSZzQfDecLQdy5ULPoh8GKJMhJTTETLycrIQrydz6SrLOdohrEJd5Z9dS5ZW9DZ8ogbqUbvsyueDbRyHZlXU73Z9vqKudVfdLeFIIBxJESTuKg1LTiAySTOk5MGRsNgRkQAthptwuKBxIMRDRLAAjAZV9BYbC9GgAYgAZm2kgj0G2-CNNAAvajwECIaIgAC+qDL2Ar9ZAtfnjebI8svZERH7Q6rU5nc9X6EXB8wOBXmXXm8Hw6su5A5aG86PmWXYHnF6318n07vs4fx6flgvm+fZXju373pW-51seQHHu+oGrreEGPtBz6nq+cEgduiHgb+kGZABDbocBG4fjuXh7n+BGoYBxGYaRCE3rh+7UUudHnlhn5IXhKFsU2GEcQx2FMZR+E1jRRH8SRl7CV+om8TB7FrpxYHyVBfFnspQlccxVHiRpAlaTJOlqaxilSfRxnkc4pn6eZmk9ipOG2YeEknhZglWc5P4sXZaEeUZZHech6n2YZjnaapPl6a5BnSUFInRWJsVhfFjFyUlCn+Q5IDwbJ3G+Sl2XhblTk3hRmWhcVaX5bpyULm5sGeQlGUhWZ1WWS1BUxQ1cWdel3X1YR7k5XlJmVe1tEBRFXmJW1flTaNZWtTxVWLSVY3WaWE0LZJS2RcFq2TXtG3LYNWXrTV43zUVl39bVLm9al93XUdu0jadB1zW9t0nVdUU3U9HXNQNFWA0+7ZtmQZDxHdIAtl2sMzV1dUXX9L0Az9QNw4FA2o2t6Mgw9O2-R9-2HYV2OE7jxPg41SnI3jj3DU1NNcWDWMswzpVfStlNc9NPOzXzPUC-twvnQTZMYxTov04Lm2y0N8viyjzMq59Ev48d0tE+zNkk1Tuts5j-Ma+T31m31eum3L1sm0raPG4ztOc+bMuW3bz0247Uusy7+vbXT9sB7bysh0LauG2LmtR8H3sO574cJ6Hvs6-7kdM9H7s+0nTsZ4r5UG-HwOJyLyel6ned+9zhfl-ntdndr70F036sR3Xkvp43vNdy3PcSxzVsp5nrvD5Xo+vePONV-XNcK232cd4vJcz5PYcNwvvfN6Tre90PXsT53O9G3vWvtyPx8X0fK9u8v2-X2vV9L5ft-T9Ts8TgfFdP2-h+-w-F+N9AGrw-uvNO-ct7nyAQA6BoDnbgOrt3KBXVv6b1VlneBZ84531fiA3BwC4EENgTg9+CDn5YIHqQ-+YDO5oPnhgseNDyF-x-rQ1h6DY6YOIew-BZDsHcP4VQwRzCBH5XocgxhU9RHCKYWwlhfCZEoJEfIsR0jVGyPUZwi2c9JFcPEcXHhCiiFCOUXI7RHtdGQKkRvBh+itF2J0X3XemjbF6KcRI6x9i3FeKcSfGOfjH68JMUomxECXFmIce4yxzjT6uMQp4iJYSkG+Jif4nOZdYkBLSUE4x1CNGRJ8Uk7x4S4mFISVmRsAB1TQtBSB4C4Ocb4aBEA1FaQkdpqBOndLafEBwE4gA)
- [Encrypted example](http://127.0.0.1:5500/#ENC:bkatoH8LlG5ourw0QT6Da4mdMea3KGRtSrl9CdMhz0U1_sFkZFwINWlj1qqt87y8ZJnA1JJFXLyvH8EkdVNBCwzrnhNwU62iIGhiF6uWTJg1D1gPmx-G4YrKn04xlGWN9JPpzVUkX9IqcZJ1nPuqxEEBegUeLCY5b5flEf-jsmdurr3YLJSxYKw7Kp3cu6HaEv9RSh8IclFSyehhxVsTmt4m1H3yx6gvxPJ8betmlvt1Zl67Y8yNbD8Z3MHBNl392PldIyJ8e9e8Pfl0jeLAYncaeZUCzdYnlmcFuMr1HaD-FdqdKNOz1e01B4euGoZJTBZpYBgUD51TBK77bwNoWUNopOvLOKnt6LvRRfx_1CYl_H62yab1F_8qzuiksMK2cuZfdDHoaOo7lExpQIzRArOFbRcPXk1kXsy2F6gGAMHAT_2wQycMKMg4SM3fKihww9wp_6Qh6poD9Zhe_u6UCy-d_fSMd0KuKYx7V-UvcHh8-9_h3vpgWfmmIssMq2TKdaCkoM0qMGcEEjTi36lYpOs2zZOX1Yg9NLEnkd_2_-NhGwNOMI_xFfWjG3H_3crDFbrahG-PeB6OuymbMxDrdbVK0fyJj8ALQYnjAbvnneYSjKOfZXFcD9YItpQ0t2_8Lnm9SSB8RDxcijpA4GO8OiESt-IVQAL6xMcBS9UxDi3poKimcdowb8IMTuJMIE-TU_GDQ0LeHB7wjbZeEz3gWwfn8l9BqWPRopJArlGJ6ZStHtuJ7IjSVahCR5s9UBCEjG19CGLYijQQqTAHEAd72a_87rW5HMYM_rTCpH_uLGE7gohx8dzf2B8r_Jas8E2XQWZ6Mr8sIohWS0jNhNZIrr9BVebfKOLO5Oirqv_1UJ7xm0fkcHcPmNUrWy7Rwsoz0UUx_1d6OWfXxhsiepx6qVeccgR6YPoF8PNemvCqLLXpIYuHcCkvYcuzisiLNW-zKEV9BWJwVnif8ytySlF24HqPWiemQJospnYeQuRgbhLwFqIRKIG56dXLKuCVwafSuk_LOVVYGr8Q1_o6RP7rH8BvR5GSjKnep3qbQxTE-z201Wpg9Ha2SQdWgGn2YwRWYd6HaMbhN-HfM4n_6SrE6miM1qkfFGrmQfVsk1q4lOIdbh4LRPBKSGRLV8Kij_jjBUXeCrIERlsQ4fUtGeNBPzWrdMWc7HAf2I7NTPNjTt_qWwyA2QaldQM21fd9JPHo0ZwRZ_bH9UF27GIa6hjOOc6rYC0_yzEvEEu0SqoxZQbJXu4bj8OnxRo1k_fRdAVFYRXsYhvZdPoqQj0ZnrqZbSWW9SiJCB5wvEwMveLtWUjYPlNbRN7xw63QynbYWmS50hLPTz6DTTUsP1GziPWQW4zpm_PaJ9V4C_otPie6HKvfChHTEZTxZqrzLIrjC-5zoYzGD1rZIFeCv7XvoFTkwDv8s3Z_XFcJrXWO5aQa8Uovr8cT-oETO0-vjHoIBHkGKjc6f6AfDL1STklZvquCPZaWQfiHc6lJC0OMHi8gZ-CNsEBbBykkBhutGkfj-N3z4Zzftf2QNNmFzHiLbBRJugi0EBTm3BKOxB-0ICM0WKDBGs-0sGlyv3Kc5s5M9Q7eb8JMhvK_umc1WlAHtBuid3PLBFCSRM-X63mAXF3LNE6TalTN7r0Bs1tE1zFoLfA4i7Y13puazpLL7VHzsvQ0uPDFjibtI3CgKN0gvm-WetqJoCHkp6VAaUBWH10RUdr7KeipyZ-0zT4lnTrno-Y2Y3Bc-x3098Tl7MFgmVxrcfS_yQe6O-JUaXflbBqsAiYzf-l2a2mIwVtAiLfoPHv9AXEfxHJtbc9Rqul4zMPQIwNgc-t5BzT-UKoYziraxoofMy3iFUGSQtiP_FYpQNCUL-hBKhcpL9Hfs5u4IwENaYgVT4Arvy-exZbvHy0BBRJWtBMkO4_viTjmOuaeubG-yVEBpkxKCHH0KeINqXLjhbqJQR1NTGPKIBlZteqHPEGw-wo6YJInM60ZcWdOWGAmgro6JrQZQQpn1zv3D1FeEp-pzJAgRJhFpYm9m9j2RrcA6eyrEKyf0M2Qdcs5ET3MzbFR8Xlw-S9eHE4h1V9HW8i4KPuHR0Z-KSNHdcn66tX3KOAo-9vW4-DCKo2BTfwrk1ebCCaaILCQOeFNcRGsDZwi2pjdn6U4bMC_QxZpLcqj3Q_xC4IvJMpKtnmJV3BT5f68ErKbnoGkL9bQ2L0c3WTBYsEfoi8Yul51D_kewdgWr8KsIonOSvntq6D6a5JogScA1tGJ6M_uQU-f92YyI1jUaRda--R24W-GLSx_YoEbYgjFD7Xm6ICFxRbZ15xw1CO0D4A5S6WLmTgy5PGiJZwHuHjc1eonZcRwnybEnyWnQtGYAK8C5Yo-XXEwjRvCoDw_bP-ysIgG0ALB4TQ3a_kx2d6E3FK4_7TN8MibO_ZtCQJmJjzmABNi57G0moP_HeRi5GX8RG6M5BfGS04F_DfEPCaggJA5BXZTd4dyPNYh-UCCTFyzsJT_bELvhoU4LOe-h_ToH5OkZcO3AS8774_OPLXA8kDm58KYU64FiqKGVSNFTOND-DqoByekRKWR_PldfVxoS86YzJw)
  - Password: `A2P7peq8aVixgB2`

## How It Works

Your spreadsheet state is stored entirely in the URL hash. The hash is LZ-String compressed JSON to keep links short, and only non-default data is included. When password protection is enabled, that compressed string is encrypted with AES-GCM (256-bit) using a PBKDF2-derived key (100k iterations, random salt/IV) and stored as URL-safe Base64 with an `ENC:` prefix. The password never leaves the browser; recipients must enter it to decrypt locally.

Example state (decompressed, before encryption):

```
{
  "rows": 2,
  "cols": 2,
  "data": [["A1", "B1"], ["A2", "B2"]],
  "formulas": [["", "=SUM(A1:A2)"]],
  "cellStyles": [[{"align": "center", "bg": "#f5f5f5", "color": "#111", "fontSize": "14"}]],
  "colWidths": [120, 100],
  "rowHeights": [32, 32],
  "theme": "light"
}
```

When you edit cells, the URL updates automatically (debounced at 200ms). Formulas are stored separately from displayed values, so both the results and the original formulas are preserved. Column widths, row heights, and cell styles are saved too. Incoming URL state is sanitized and validated (DOMParser whitelist, formula regex, safe JSON parsing), oversized hashes are rejected, and legacy uncompressed hashes are still supported.

## Usage

| Action | How |
|--------|-----|
| Edit cell | Double-click a cell or click and start typing |
| Format text | Select text, click B/I/U buttons or use Ctrl+B/I/U |
| Align text | Click left/center/right alignment buttons |
| Set font size | Use size buttons (Auto, 10-24) |
| Set cell colors | Use background/text color pickers |
| Resize column/row | Drag header resize handles |
| Navigate cells | Arrow keys (when not editing) |
| Select range | Click and drag across cells |
| Extend selection | Shift+Click or Shift+Arrow |
| Clear selection | Press Escape |
| Add row | Click "+ Row" button (max 30) |
| Add column | Click "+ Column" button (max 15) |
| Clear all | Click "Clear" button (with confirmation) |
| Import CSV | Click import button and choose a .csv file |
| Export CSV | Click download button |
| Enter formula | Type `=` followed by function (e.g., `=SUM(A1:B5)`) |
| Select formula range | Click/drag cells while editing a formula |
| Share | Click copy button to copy URL |
| Lock with password | Click the lock icon (open) and set a password in the modal |
| Unlock encrypted link | Open the link, enter password in the modal to decrypt |
| Remove password | Click the lock icon (closed) and confirm removal |
| Toggle theme | Click sun/moon icon |

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| Ctrl+B | Bold |
| Ctrl+I | Italic |
| Ctrl+U | Underline |
| Arrow keys | Move selection (when not editing) |
| Shift+Arrow | Extend selection |
| Enter | Evaluate formula and move down (or insert formula suggestion when dropdown is open) |
| Escape | Clear selection / Close formula dropdown |
| Arrow Up/Down (formula dropdown) | Navigate suggestions |
| Tab (formula dropdown) | Insert active suggestion |

## Tech Stack

- Vanilla HTML/CSS/JavaScript (no frameworks)
- CSS Grid for spreadsheet layout
- CSS Custom Properties for theming
- LZ-String (URL state compression via CDN)
- Web Crypto API (AES-GCM + PBKDF2) for optional password protection
- Font Awesome 6.5.1 (icons via CDN)
- Google Analytics (gtag.js) for usage tracking
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
- Default grid: 10 rows x 10 columns
- Formulas limited to SUM and AVG range syntax
- CSV imports larger than 30x15 are truncated
- URL length limits may apply for very large spreadsheets; encrypted links are longer
- Losing the password means the encrypted data cannot be recovered

## File Structure

```
spreadsheet/
|-- index.html      # Single-page app structure
|-- styles.css      # All styling including dark mode
|-- script.js       # Application logic (IIFE module)
|-- logo.png        # App logo
|-- favicon.png     # Browser favicon
|-- CLAUDE.md       # Development documentation
`-- README.md       # This file
```

## Architecture

- **State Management** - `data`, `formulas`, `cellStyles`, `rows`, `cols`, `colWidths`, `rowHeights`
- **URL Sync** - LZ-String compressed JSON with debounced updates and legacy fallback
- **Event Delegation** - All cell events handled on container
- **CSS Grid** - Dynamic column template and row heights set via JavaScript
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

### Latest - Optional Password Protection
- Added AES-GCM (256-bit) encryption with PBKDF2 (100k iterations) for URL hashes (`ENC:` prefix)
- Lock/unlock toolbar button with modal flows for setting, unlocking, and removing passwords
- URL-safe Base64 payloads; encryption failures fall back safely without breaking sharing

### v1.4 - Keyboard Navigation
- Arrow keys move the active selection without entering edit mode
- Shift+Arrow expands selections from the anchor cell
- Double-click or start typing to enter edit mode

### v1.3 - Formula Support
- Added formula evaluation with `=SUM(range)` function
- Formula autocomplete dropdown with suggestions
- Click-to-select range references while editing formulas
- Keyboard navigation in formula dropdown (Up/Down/Enter/Tab/Escape)
- Formulas preserved in shareable URLs
- Enter key evaluates formula and moves to next row

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
