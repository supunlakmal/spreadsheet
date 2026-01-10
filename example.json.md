## What `example.json` Represents

`example.json` is a minified spreadsheet state snapshot—exactly the shape we pack into the URL hash (via `minifyState` in `modules/urlManager.js`). It is **not** the expanded grid; it is the compact form with sparse arrays and short keys.

## Field Reference
- `r`: row count (int).
- `c`: column count (int).
- `t`: theme string (`"light"` or `"dark"`).
- `d`: data as sparse triplets `[row, col, value]`. Only non-empty cells appear.
- `f`: formulas as sparse triplets `[row, col, "=FORMULA()"]`.
- `s`: cell styles as sparse triplets `[row, col, styleObject]` where style keys are minified: `a` align, `b` background color, `c` text color, `z` font size (px).
- `w`: column widths array (length = `c`). Defaults are omitted, so presence here means custom widths.

## Supported Formulas

All formulas must start with `=`. The app supports:

### Range Functions
| Function | Syntax | Description |
|----------|--------|-------------|
| `SUM` | `=SUM(A1:B5)` | Adds all numbers in the range |
| `AVG` | `=AVG(A1:B5)` | Average of non-empty numeric cells in range |

### Arithmetic Expressions
Cell references and basic math with proper operator precedence:
- **Cell references**: `=A1`, `=B2`, `=AA99`
- **Operators**: `+` `-` `*` `/`
- **Parentheses**: `=(A1+B1)/2`
- **Unary**: `=-A1`, `=+B2`

Examples:
```
=A1+B1
=A1*B2-C3
=(A1+A2+A3)/3
=B5*1.15
```

### Error Codes
| Code | Meaning |
|------|---------|
| `#DIV/0!` | Division by zero |
| `#REF!` | Invalid or out-of-bounds cell reference |
| `#ERROR!` | Malformed expression or unknown formula |

## How to Read It
1) Treat `d`, `f`, and `s` as sparse arrays—build an empty `r x c` grid first, then set only the listed coordinates.
2) Expand style keys: `{ "b": "#ffee00", "c": "#000000" }` ⇒ `{ bg: "#ffee00", color: "#000000" }`.
3) If a field is missing, assume defaults: empty cells, no formulas, default styles, default row heights/col widths, light theme, read-only off.

## Typical Flow in Code
- Decode hash → `expandState`/`validateAndNormalizeState` → render.
- Encode for sharing → `minifyState` → compress/encrypt → hash.

This file is already decompressed; you can load it directly by parsing JSON, expanding sparse arrays to a dense grid, and applying styles/formulas.

## Example JSON

```json
{
  "r": 19,
  "c": 8,
  "t": "dark",
  "d": [
    [0, 0, "Personal Tax Calculation (OPEN)"],
    [2, 0, "Monthly Income"],
    [2, 1, "  400,000.00 "],
    [3, 0, "Yearly Income"],
    [3, 1, "  4,800,000.00 "],
    [3, 2, "404"],
    [5, 0, "Foreign Sources"],
    [6, 0, "New Tax Rate - 15%"],
    [7, 0, "2025/2026"],
    [7, 1, "  1,800,000.00 "],
    [7, 2, "  3,000,000.00 "],
    [7, 3, "15%"],
    [7, 4, "  450,000.00 "],
    [7, 5, "37500"],
    [10, 0, "Normal Personal Tax"],
    [11, 1, "  1,800,000.00 "],
    [11, 2, "  3,000,000.00 "],
    [11, 3, "0%"],
    [11, 4, "  -   "],
    [12, 1, "  1,000,000.00 "],
    [12, 2, "  2,000,000.00 "],
    [12, 3, "6%"],
    [12, 4, "  60,000.00 "],
    [13, 1, "  500,000.00 "],
    [13, 2, "  1,500,000.00 "],
    [13, 3, "18%"],
    [13, 4, "  90,000.00 "],
    [14, 1, "  500,000.00 "],
    [14, 2, "  1,000,000.00 "],
    [14, 3, "24%"],
    [14, 4, "  120,000.00 "],
    [15, 1, "  500,000.00 "],
    [15, 2, "  500,000.00 "],
    [15, 3, "30%"],
    [15, 4, "  150,000.00 "],
    [16, 2, "  500,000.00 "],
    [16, 3, "36%"],
    [16, 4, "  180,000.00 "],
    [17, 0, "Total Tax"],
    [17, 4, "  600,000.00 "]
  ],
  "f": [[3, 2, "=SUM(B3:B4)"]],
  "s": [
    [0, 0, { "c": "#ff0000", "z": "14" }],
    [5, 0, { "b": "#ffee00", "c": "#000000" }]
  ],
  "w": [239, 100, 100, 100, 100, 100, 100, 100]
}
```
