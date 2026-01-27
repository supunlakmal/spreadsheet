# HashCal

HashCal is a lightweight, client-only calendar. Your calendar state lives entirely in the URL hash, so sharing is as simple as copying the link. No backend, no accounts.

## Features

- URL-based storage with LZ-String compression
- Day, week, month, and year views
- Create, edit, and delete events
- All-day events and duration-based events
- Recurring events (daily, weekly, monthly, yearly)
- Copy-link sharing
- Export calendar to JSON
- Import events from .ics files
- Optional password lock (AES-GCM + PBKDF2)
- Theme toggle (light/dark)
- Week start toggle (Sunday/Monday)

## How it works

- The calendar state is serialized to JSON and compressed into the URL hash.
- If a password is set, the compressed payload is encrypted using AES-GCM.
- No data leaves the browser unless you share the link.

Encrypted links start with `#ENC:`. Clearing the hash resets the calendar.

## Getting started

Open `index.html` in your browser.

For clipboard and file import features, run a local server:

```bash
npx serve .
```

## Project structure

- `index.html` - UI markup
- `styles.css` - Styles
- `script.js` - App logic
- `modules/`
  - `calendarRender.js` - Calendar rendering
  - `recurrenceEngine.js` - Recurring event expansion
  - `icsImporter.js` - .ics parsing
  - `hashcalUrlManager.js` - URL read/write + compression
  - `cryptoManager.js` - Encryption helpers
  - `lz-string.min.js` - Compression library

## License

MIT. See `LICENSE`.
