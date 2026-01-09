Based on the architectural patterns observed in the **supunlakmal/spreadsheet** repository (Vanilla JavaScript, DOM-based interaction, state stored in a central object), here is the **Extended Development Specification** for the **Smart Status Bar**.

---

# Feature Specification: Smart Status Bar (Sum / Avg / Count)

## 1. Overview

The **Smart Status Bar** mimics the functionality of Excel/Google Sheets. When a user selects multiple cells in the grid, a fixed footer bar will automatically appear (or update) to display:

- **Sum:** Total of all numeric values.
- **Average:** The mean of the numeric values.
- **Count:** Total number of non-empty selected cells.
- **Min/Max:** (Optional) The lowest and highest numbers in the selection.

## 2. Technical Architecture

### 2.1 Component Structure

- **UI Layer:** A new HTML `<footer>` element fixed to the bottom of the viewport.
- **Logic Layer:** A JavaScript function triggered by the existing `mouseup` or `selection` events.
- **Data Source:** The feature will read from the **internal state object** (usually named `data` or derived from the DOM `innerText`), not just the raw HTML, to ensure calculation accuracy.

### 2.2 Performance Constraints

- **Debouncing:** Calculations must be performant. If a user selects 500 cells, the loop must run instantly.
- **Zero Dependencies:** No external math libraries (like math.js) allowed.

---

## 3. Implementation Plan

### 3.1 UI Changes (`index.html` & `style.css`)

We need a container to display the stats. It should overlap the bottom of the sheet without hiding data (add padding to the body if necessary).

**HTML:**

```html
<!-- Add before </body> -->
<div id="status-bar" class="status-bar-hidden">
  <span class="stat-item">Count: <b id="stat-count">0</b></span>
  <span class="stat-item">Sum: <b id="stat-sum">0</b></span>
  <span class="stat-item">Avg: <b id="stat-avg">0</b></span>
</div>
```

**CSS:**

```css
#status-bar {
  position: fixed;
  bottom: 0;
  right: 0;
  width: 100%;
  background: #f8f9fa;
  border-top: 1px solid #ccc;
  padding: 5px 20px;
  text-align: right;
  font-family: sans-serif;
  font-size: 14px;
  z-index: 100;
  display: none; /* Hidden by default */
}

#status-bar.active {
  display: block;
}

.stat-item {
  margin-left: 20px;
  color: #333;
}
```

### 3.2 Logic Integration (`script.js`)

You need to hook into the event that handles cell selection. In this codebase, look for the mouse event handlers (likely `mouseup` or a specific function like `handleSelection`).

**The Algorithm:**

1.  Identify all DOM elements that currently have the class `.selected` (or `.active`).
2.  Iterate through them and extract the raw value.
3.  **Sanitize:** Distinguish between Text ("Apple"), Numbers ("10.5"), and Empty strings.
4.  **Compute:** Calculate stats.
5.  **Render:** Update the `innerHTML` of the status bar.

---

## 4. Edge Cases & Handling (Critical Analysis)

Since spreadsheets contain messy user data, these edge cases **must** be handled to prevent `NaN` errors.

| Edge Case                 | Example Input        | Handling Rule                                                                                               |
| :------------------------ | :------------------- | :---------------------------------------------------------------------------------------------------------- |
| **Mixed Content**         | `10`, `20`, `Banana` | **Ignore Text.** The Sum is 30. The Count is 3 (or 2 depending on preference, usually count all non-empty). |
| **Floating Point Math**   | `0.1`, `0.2`         | Javascript sums this as `0.300000004`. **Fix:** Use `result.toFixed(2)` for display to avoid ugly decimals. |
| **Currency Symbols**      | `$100`, `€50`        | **Parse Logic:** Strip non-numeric characters before parsing. Use regex: `val.replace(/[^0-9.-]+/g,"")`.    |
| **Empty Cells**           | `""` (Empty String)  | **Exclude:** Do not count as 0. Do not increment the "Count" metric.                                        |
| **Date Strings**          | `2023-01-01`         | **Ignore:** Unless you implement complex date math, treat dates as Strings (ignore for Sum/Avg).            |
| **Single Cell Selection** | User clicks 1 cell   | **Hide Bar:** The status bar should usually be hidden unless >1 cell is selected to reduce visual noise.    |
| **Formula Cells**         | `=SUM(A1:A5)`        | **Read Value, Not Formula:** Ensure you read the _rendered_ text (result), not the input value (formula).   |

---

## 5. Development Specs (Code Snippets)

### Step 1: The Calculation Function

Add this utility function to `script.js`.

```javascript
function updateStatusBar() {
  // 1. Get all selected cells
  // Note: Adjust '.cell.selected' to match the actual CSS class used in the repo for highlighting
  const selectedCells = document.querySelectorAll(".cell.selected");

  // UI Elements
  const bar = document.getElementById("status-bar");
  const elCount = document.getElementById("stat-count");
  const elSum = document.getElementById("stat-sum");
  const elAvg = document.getElementById("stat-avg");

  // 2. Optimization: If 0 or 1 cell, hide bar
  if (selectedCells.length < 2) {
    bar.classList.remove("active");
    return;
  }

  let sum = 0;
  let countNumeric = 0;
  let countTotal = 0;

  // 3. Loop and Calculate
  selectedCells.forEach((cell) => {
    // Get raw value (innerText is usually safer than value for contenteditable divs)
    let rawValue = cell.innerText || cell.textContent;

    // Clean whitespace
    rawValue = rawValue.trim();

    if (rawValue !== "") {
      countTotal++; // It's not empty, so it counts

      // Attempt to parse number
      // Remove currency symbols like $, but keep negative signs and decimals
      let cleanNumber = rawValue.replace(/[^0-9.-]+/g, "");
      let val = parseFloat(cleanNumber);

      if (!isNaN(val)) {
        sum += val;
        countNumeric++;
      }
    }
  });

  // 4. Update UI
  // Logic: If no numbers were selected, don't show Sum/Avg, just Count
  if (countNumeric > 0) {
    elSum.innerText = sum % 1 !== 0 ? sum.toFixed(2) : sum; // Show decimals only if needed
    elAvg.innerText = (sum / countNumeric).toFixed(2);
    elSum.parentElement.style.display = "inline";
    elAvg.parentElement.style.display = "inline";
  } else {
    // Hide Sum/Avg if only text is selected
    elSum.parentElement.style.display = "none";
    elAvg.parentElement.style.display = "none";
  }

  elCount.innerText = countTotal;
  bar.classList.add("active");
}
```

### Step 2: Hooking the Event

You must find where the spreadsheet handles the "End of Selection".

- Look for `window.addEventListener('mouseup', ...)` or the function that runs when dragging stops.
- Add the call to `updateStatusBar()` inside that handler.

```javascript
// Example injection point in existing code:
document.addEventListener("mouseup", () => {
  // ... existing selection logic ...

  updateStatusBar(); // <--- Add this line
});

// Also add to keyup (in case user changes data inside a selected range)
document.addEventListener("keyup", () => {
  updateStatusBar();
});
```

### Step 3: CSS Class Hook

Ensure the repository actually adds a class (like `.selected`) to highlighted cells.

- _If the repo uses inline styles for selection (e.g., `background: blue`),_ you will need to modify the selection logic to add a CSS class instead, or update the `querySelectorAll` logic to find cells by background color (which is messy/not recommended).
- **Recommendation:** Modify the selection function to toggle a class named `selected`.

---

## 6. Testing Protocol

1.  **Basic Math:** Select cells containing `10`, `10`, `10`. Expect: Sum `30`, Avg `10`, Count `3`.
2.  **The "Text" Test:** Select `10`, `Apple`, `20`. Expect: Sum `30`, Count `3`.
3.  **The "Empty" Test:** Select `10`, `[Empty]`, `20`. Expect: Sum `30`, Count `2`.
4.  **Visual Test:** Click a single cell. Ensure the bar disappears.
5.  **Update Test:** Select 3 cells. Change a number in one of them while it is still selected. Does the sum update instantly? (Requires the `keyup` listener).
