This is a **brilliant idea**. By exposing the raw JSON, you effectively create an **"AI Interface"** for your spreadsheet.

Since the application is serverless, the "State" is just a JavaScript object. If you let the user view and edit this object as text, they can copy it to ChatGPT, ask for changes (e.g., _"Populate this with 50 fake names and emails"_), and paste the result back.

Here is the **Development Specification** for the **"Raw JSON Editor / AI Bridge"**.

---

# Feature Specification: JSON Data Editor (AI Bridge)

## 1. Overview

A feature that allows users to view the current spreadsheet state as a raw JSON string in a modal. The user can:

1.  **Copy** the JSON to use in LLMs (ChatGPT/Claude).
2.  **Paste** modified JSON back into the box.
3.  **Update** the grid instantly (which also regenerates the URL hash).

## 2. Technical Architecture

- **Data Flow:** `Grid State` <-> `JSON.stringify` <-> `Textarea` <-> `JSON.parse` <-> `Grid State`.
- **Constraint:** The JSON must be "Pretty Printed" (indented) so it is readable by humans and LLMs.
- **Safety:** The import function must have strict error handling to prevent the app from crashing if the user/LLM pastes invalid JSON.

## 3. Implementation Plan

### 3.1 UI Changes (`index.html`)

We need a modal with a large text area and action buttons.

```html
<!-- Add this to your HTML body -->
<div id="json-modal" style="display:none; position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.85); z-index:9999; justify-content:center; align-items:center;">
    <div style="background:white; width:80%; height:80%; padding:20px; border-radius:8px; display:flex; flex-direction:column; gap:10px;">

        <div style="display:flex; justify-content:space-between; align-items:center;">
            <h3 style="margin:0;">🔧 Raw Data (AI Interface)</h3>
            <button id="close-json-modal" style="background:none; border:none; font-size:20px; cursor:pointer;">&times;</button>
        </div>

        <p style="margin:0; font-size:12px; color:#666;">
            Copy this to ChatGPT/Claude to modify data, then paste the result back here.
        </button>

        <!-- The Editor Area -->
        <textarea id="json-input" style="flex:1; font-family:monospace; font-size:12px; padding:10px; border:1px solid #ccc;"></textarea>

        <div style="display:flex; gap:10px; justify-content:flex-end;">
            <button id="copy-json-btn" style="padding:8px 15px;">📋 Copy JSON</button>
            <button id="save-json-btn" style="padding:8px 15px; background:green; color:white; border:none; cursor:pointer;">💾 Update Grid</button>
        </div>

    </div>
</div>
```

### 3.2 Logic Implementation (`script.js`)

You need to integrate this with your existing `getData()` and `loadData()` functions.

```javascript
document.addEventListener("DOMContentLoaded", () => {
  // DOM Elements
  const modal = document.getElementById("json-modal");
  const textarea = document.getElementById("json-input");
  const openBtn = document.getElementById("open-json-btn"); // Add this ID to a button in your toolbar
  const closeBtn = document.getElementById("close-json-modal");
  const saveBtn = document.getElementById("save-json-btn");
  const copyBtn = document.getElementById("copy-json-btn");

  // 1. OPEN MODAL & EXPORT DATA
  function openModal() {
    // Assume 'getData()' is your existing function that returns the sheet object
    // If your app uses a global variable like 'data', use that.
    const currentState = window.getData ? window.getData() : {};

    // Pretty print with 2 spaces indentation
    textarea.value = JSON.stringify(currentState, null, 2);

    modal.style.display = "flex";
  }

  // 2. IMPORT DATA & UPDATE GRID
  function saveAndClose() {
    const rawString = textarea.value;

    try {
      // A. Validate JSON
      const newState = JSON.parse(rawString);

      // B. Basic Schema Validation (Prevent crashes)
      // Check if the object looks somewhat correct (e.g., has 'rows' or 'data')
      if (typeof newState !== "object" || newState === null) {
        throw new Error("Invalid Data Structure");
      }

      // C. Update the Grid
      // Assume 'loadSpreadsheet(data)' is your existing function
      if (window.loadSpreadsheet) {
        window.loadSpreadsheet(newState);
      }

      // D. Update URL Hash
      // We re-encode the data to base64 to update the shareable link
      // Assuming your app has a function to update hash, otherwise:
      const encoded = btoa(JSON.stringify(newState));
      window.location.hash = encoded;

      modal.style.display = "none";
      alert("Grid updated successfully!");
    } catch (e) {
      alert("❌ Error: Invalid JSON.\n\nPlease check your syntax. If you used ChatGPT, make sure it returned valid JSON.");
      console.error(e);
    }
  }

  // 3. COPY TO CLIPBOARD
  function copyToClipboard() {
    textarea.select();
    document.execCommand("copy"); // Fallback or use Navigator API
    // navigator.clipboard.writeText(textarea.value); // Modern way

    const originalText = copyBtn.innerText;
    copyBtn.innerText = "Copied! ✅";
    setTimeout(() => (copyBtn.innerText = originalText), 2000);
  }

  // Event Listeners
  if (openBtn) openBtn.addEventListener("click", openModal);
  if (closeBtn) closeBtn.addEventListener("click", () => (modal.style.display = "none"));
  if (saveBtn) saveBtn.addEventListener("click", saveAndClose);
  if (copyBtn) copyBtn.addEventListener("click", copyToClipboard);

  // Close on background click
  modal.addEventListener("click", (e) => {
    if (e.target === modal) modal.style.display = "none";
  });
});
```

---

## 4. Edge Cases & AI Specific Handling

Since this feature is designed for LLM interaction, we need to handle specific issues that arise when pasting from AI.

| Edge Case                   | Problem                                                                                                                  | Handling Strategy                                                                                                                                                                       |
| :-------------------------- | :----------------------------------------------------------------------------------------------------------------------- | :-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Markdown Wrapping**       | LLMs often wrap code in triple backticks (e.g., \`\`\`json ... \`\`\`). If the user pastes this, `JSON.parse` will fail. | **Sanitization:** Before parsing, strip the markdown. <br> ` let cleanStr = rawString.replace(/```json/g, "").replace(/```/g, ""); `                                                    |
| **Partial JSON**            | User might accidentally copy only half the JSON string.                                                                  | **Try/Catch:** The `try...catch` block in the code above handles this. Show a clear error: "Unexpected end of data".                                                                    |
| **Huge Data Sets**          | Stringifying a massive spreadsheet might freeze the browser for a second.                                                | **No Fix Needed:** Modern browsers handle ~5MB text in textareas fine. Just be aware of the lag.                                                                                        |
| **Structure Hallucination** | The LLM might change the key names (e.g., changing `cell_id` to `id`).                                                   | **Soft Validation:** In the `saveAndClose` function, check if critical keys exist. If they are missing, alert the user: "The JSON structure seems wrong. Did the AI change the format?" |

## 5. Suggested "System Prompt" for Users

To make this feature truly useful, add a small "Help" link in the modal that gives the user a prompt to copy.

**Add this text to the Modal UI:**

> **Tip:** When asking AI to edit this, use this prompt:
> _"Here is the JSON data for a spreadsheet. Please keep the exact same structure and keys, but [insert request here, e.g., add 10 rows of random user data]. Return ONLY the JSON."_
