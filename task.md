This is an excellent application for this architecture. A calendar app requires relatively low data volume (mostly short text strings and timestamps) compared to a spreadsheet, making it highly suitable for the **"State-in-URL"** pattern.

Below is the **Architectural Guide** for your Project: **"HashCal" (Serverless Privacy Calendar)**.

---

### 1. Core Philosophy & Limitations

- **Zero Knowledge:** Since there is no database, you (the developer) cannot see user data.
- **The URL is the Save File:** Bookmarking the page "saves" the calendar. Sharing the link shares the calendar.
- **Optimization:** Every byte counts. The Data Schema must be aggressively minified to keep URLs manageable (target < 4000 characters).

### 2. Technology Stack

- **Core:** Vanilla JavaScript (ES6 Modules)
- **Rendering:** CSS Grid (for the calendar layout)
- **Compression:** `lz-string` (Essential for packing JSON into URL)
- **Encryption:** Web Crypto API (AES-GCM 256-bit)
- **Time Management:** `date-fns` (lightweight) or native `Intl.DateTimeFormat` (to zero dependency).
- **Import/Export:** `.ics` (iCalendar) parser/generator (essential for migrating from Google Calendar).

---

### 3. Minified Data Schema (The "Secret Sauce")

To make this work, you cannot store verbose JSON keys like `{"title": "Meeting"}`. You must map keys to single characters.

**The State Object:**

```json
{
  "t": "Calendar Title",
  "c": ["#ff0000", "#00ff00"], // Array of custom category colors
  "e": [
    // Array of Events (Sparse Array)
    // [StartTimestamp (minified), Duration (mins), Title, Type/ColorIndex, RecurrenceRule]
    [1706601600, 60, "Dentist", 0], // One-time event
    [1706688000, 30, "Daily Standup", 1, "d"], // Recur: "d"=daily
    [1706601600, 0, "B-day", 2, "y"] // All-day event (duration 0), recur yearly
  ],
  "s": { "d": 0, "m": 1 } // Settings: (d)ark mode off, (m)onday start
}
```

**Optimization Strategies:**

1.  **Base Timestamps:** Store a `base_year` in the object. Store event times as _minutes elapsed since start of that year_ (integers) rather than full Unix timestamps or ISO strings.
2.  **Indexing:** Do not store "Work" or "Personal" strings in every event. Store them in a legend `c` (categories) and use their index `0, 1` in the event array.

---

### 4. Component Architecture (Modules)

You should organize your code into these exact files:

#### A. `urlManager.js` (The "Hard Drive")

- **Listen:** `window.addEventListener("hashchange")`
- **Read:**
  1. Grab hash.
  2. Check if it starts with `ENC:` (Encrypted).
  3. If yes, prompt password -> Decrypt.
  4. Decompress via `lz-string`.
  5. `JSON.parse`.
- **Write (Debounced):**
  1. User adds event.
  2. Wait 500ms (debounce to prevent lagging).
  3. `JSON.stringify` state.
  4. Compress.
  5. If password exists -> Encrypt.
  6. `history.replaceState(..., '#' + newString)`.

#### B. `calendarRender.js` (The View)

- **Grid System:** dynamic HTML generation.
  - 35-42 divs (7 columns x 5-6 rows) for Month View.
  - Columns for Week View.
- **Virtualization:** Since we have the data, calculate which events fall in the _visible range_ and only render those DOM nodes.

#### C. `recurrenceEngine.js` (The Logic)

This is the hardest part. You don't save _every_ daily meeting in the URL. You save the _rule_.

- Input: `Start Time`, `Rule ("d" | "w" | "m" | "y")`.
- Function: `getEventsForRange(startDate, endDate)`.
- Logic: Just like Google Calendar, calculate "ghost" events dynamically when the user switches months.

#### D. `cryptoManager.js`

- Use `window.crypto.subtle`.
- **Derive Key:** PBKDF2 (Password + Salt).
- **Encryption:** AES-GCM (Includes an "integrity tag" so if the password is wrong, it fails instantly rather than outputting garbage).

#### E. `icsImporter.js`

- Allow users to drag & drop a `.ics` file.
- Parse standard RFC 5545 data into your "Minified Schema".
- _Crucial:_ This allows users to migrate from Google/Apple Calendar easily.

---

### 5. Flow of Interaction

#### Use Case 1: Creating a Calendar

1. User opens `site.com`.
2. App generates default empty state `{ t: "My Cal", e: [] }`.
3. URL Hash is updated immediately.
4. User sets a password "MySecret".
5. App encrypts state -> updates URL hash to `#ENC:xyz...`.
6. User clicks "Share", app copies the URL to clipboard.

#### Use Case 2: Adding a Recurring Event

1. User clicks Jan 28th -> "Add Event".
2. Inputs: "Gym", "Weekly".
3. App updates state: push `[timestamp, 60, "Gym", 0, "w"]` to array.
4. `recurrenceEngine` immediately draws this event on Jan 28, Feb 4, Feb 11... in the view.
5. `urlManager` compresses state -> Updates Browser URL.

---

### 6. Roadmap & Guide

**Phase 1: The Core (No Crypto)**

1.  Set up the HTML/CSS Grid for a static calendar month.
2.  Implement `state <-> lz-string <-> URL Hash`.
3.  Create the `eventManager` to add simple one-time events to the array.

**Phase 2: Recurrence & Complexity**

1.  Implement logic for repeating events (Daily/Weekly).
2.  Handle collision (visual stacking) when two events happen at the same time (CSS logic).

**Phase 3: Privacy & Security**

1.  Add the Password Modal.
2.  Implement AES-GCM encryption/decryption flow.
3.  Add "Destructive Burn": Since data is in the URL, "deleting" the calendar is just `window.location.hash = ""`.

**Phase 4: Integrations**

1.  Add `.ics` file drop support (Parsing libraries available on npm, can be bundled).
2.  **Notification System:** Since this is client-side, use `Notification API`. Note: This only works if the tab is open.

### 7. Known Limitation to Mitigate

**URL Length limit:**
Browsers support long URLs (Chrome handles 2MB+ in hash), but _sharing_ links on WhatsApp/Twitter/Slack often truncates them around 2048 chars.

- **Mitigation:** Add a "Snapshot" feature. If the calendar gets too big, allow the user to "Export to File" (save JSON) or provide a specialized short-link service (which breaks the serverless purity, so sticking to File Export is safer).

### 8. Directory Structure Example

```text
/
├── index.html       (Main UI)
├── css/
│   ├── grid.css     (Calendar Grid Logic)
│   └── style.css    (Theming)
├── js/
│   ├── app.js       (Controller/Main)
│   ├── state.js     (Compress/Decompress/Storage)
│   ├── crypto.js    (Encrypt/Decrypt)
│   ├── recurrence.js (Math for repeating dates)
│   └── ui.js        (DOM Manipulation)
├── vendor/
│   └── lz-string.js
└── Dockerfile       (For deployment via Caddy like the original project)
```

This approach replicates the brilliant simplicity of the Spreadsheet project but solves a problem (Personal Planning) that is arguably more universally useful for mobile users.
