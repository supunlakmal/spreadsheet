import { FORMULA_SUGGESTIONS } from "./constants.js";

// ==========================================
// Formula Helper Functions
// ==========================================

// Convert column index to letters (0 = A, 25 = Z, 26 = AA)
export function colToLetter(col) {
  if (!Number.isInteger(col) || col < 0) return "";
  let n = col;
  let letters = "";
  while (n >= 0) {
    const remainder = n % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    n = Math.floor(n / 26) - 1;
  }
  return letters;
}

// Convert column letter(s) to index: A=0, B=1, ..., O=14
export function letterToCol(letters) {
  letters = letters.toUpperCase();
  let col = 0;
  for (let i = 0; i < letters.length; i++) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  return col - 1;
}

// Parse cell reference "A1" → { row: 0, col: 0 }
export function parseCellRef(ref) {
  const match = ref.toUpperCase().match(/^([A-Z]+)(\d+)$/);
  if (!match) return null;
  return {
    col: letterToCol(match[1]),
    row: parseInt(match[2], 10) - 1,
  };
}

// Parse range "A1:B5" → { startRow, startCol, endRow, endCol }
export function parseRange(range) {
  const parts = range.split(":");
  if (parts.length !== 2) return null;
  const start = parseCellRef(parts[0].trim());
  const end = parseCellRef(parts[1].trim());
  if (!start || !end) return null;
  return {
    startRow: Math.min(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endRow: Math.max(start.row, end.row),
    endCol: Math.max(start.col, end.col),
  };
}

// Build cell reference string like "A1"
export function buildCellRef(row, col) {
  return colToLetter(col) + (row + 1);
}

// Build range reference string like "A1:B5"
export function buildRangeRef(startRow, startCol, endRow, endCol) {
  const minRow = Math.min(startRow, endRow);
  const maxRow = Math.max(startRow, endRow);
  const minCol = Math.min(startCol, endCol);
  const maxCol = Math.max(startCol, endCol);

  if (minRow === maxRow && minCol === maxCol) {
    // Single cell
    return buildCellRef(minRow, minCol);
  }
  return buildCellRef(minRow, minCol) + ":" + buildCellRef(maxRow, maxCol);
}

const VISUAL_FORMULA_REGEX = /^=\s*(PROGRESS|TAG|RATING|CHART)\s*\((.*)\)\s*$/i;

export function isVisualFormula(formula) {
  if (!formula || typeof formula !== "string") return false;
  const match = formula.match(VISUAL_FORMULA_REGEX);
  if (!match) return false;
  return match[2].trim().length > 0;
}

// Check if a formula string is valid (arithmetic or supported functions)
export function isValidFormula(formula) {
  if (!formula || !formula.startsWith("=")) return false;

  if (isVisualFormula(formula)) return true;

  // Check supported functions
  if (/^=SUM\([A-Z]+\d+:[A-Z]+\d+\)$/i.test(formula)) return true;
  if (/^=AVG\([A-Z]+\d+:[A-Z]+\d+\)$/i.test(formula)) return true;

  // Check arithmetic
  return FormulaEvaluator.isArithmeticFormula(formula);
}

// ==========================================
// Formula Evaluation Logic
// ==========================================

export const FormulaEvaluator = {
  // Tokenize arithmetic expression into array of tokens
  tokenizeArithmetic(expr) {
    const tokens = [];
    let i = 0;

    while (i < expr.length) {
      // Skip whitespace
      if (/\s/.test(expr[i])) {
        i++;
        continue;
      }

      // Numbers (including decimals)
      if (/[\d.]/.test(expr[i])) {
        // Lookahead to see if it's actually part of a cell ref (like 123... wait, cell refs start with letters)
        // Actually, cell refs MUST start with letters. So if it starts with digit, it's a number.
        let num = "";
        while (i < expr.length && /[\d.]/.test(expr[i])) {
          num += expr[i++];
        }
        if (!/^\d+\.?\d*$|^\d*\.\d+$/.test(num)) {
          return { error: "Invalid number: " + num };
        }
        tokens.push({ type: "NUMBER", value: parseFloat(num) });
        continue;
      }

      // Identifiers: Cell References (A1) OR Function Names (AVG)
      // Both start with letters.
      if (/[A-Z]/i.test(expr[i])) {
        let ident = "";
        while (i < expr.length && /[A-Z]/i.test(expr[i])) {
          ident += expr[i++];
        }
        // If followed by digits, it might be a cell ref (e.g. A1) or a function with number in name (none currently)
        // Standard cell ref: Letters + Digits.
        // Function name: Letters only (usually).
        
        // Peek ahead for digits
        let digits = "";
        if (i < expr.length && /\d/.test(expr[i])) {
            while (i < expr.length && /\d/.test(expr[i])) {
                digits += expr[i++];
            }
        }

        if (digits.length > 0) {
            // It's a cell reference (e.g. A1)
            tokens.push({ type: "CELL_REF", value: (ident + digits).toUpperCase() });
        } else {
            // It's a function name or just letters (e.g. AVG, SUM, PI)
            tokens.push({ type: "IDENTIFIER", value: ident.toUpperCase() });
        }
        continue;
      }

      // Operators
      if ("+-*/".includes(expr[i])) {
        tokens.push({ type: "OPERATOR", value: expr[i] });
        i++;
        continue;
      }

      // Colon for ranges
      if (expr[i] === ":") {
        tokens.push({ type: "COLON" });
        i++;
        continue;
      }

      // Parentheses
      if (expr[i] === "(") {
        tokens.push({ type: "LPAREN" });
        i++;
        continue;
      }
      if (expr[i] === ")") {
        tokens.push({ type: "RPAREN" });
        i++;
        continue;
      }

      // Unknown character
      return { error: "Unexpected character: " + expr[i] };
    }

    return { tokens };
  },

  // Parse and evaluate arithmetic expression with proper precedence
  evaluateArithmeticExpr(tokens, context) {
    let pos = 0;
    // We bind 'this' to access evaluateSUM/AVG, or we can just access FormulaEvaluator directly if strict.
    // Better to use 'self' or access properties from outer scope if possible, 
    // but here we are inside an object method. Context 'this' might get lost in inner functions.
    const evaluator = this; 

    const { getCellValue, rows, cols } = context;

    function peek() {
      return tokens[pos];
    }

    function consume() {
      return tokens[pos++];
    }

    function parseExpr() {
      let left = parseTerm();
      if (left.error) return left;

      while (peek() && peek().type === "OPERATOR" && (peek().value === "+" || peek().value === "-")) {
        const op = consume().value;
        const right = parseTerm();
        if (right.error) return right;

        left = { value: op === "+" ? left.value + right.value : left.value - right.value };
      }
      return left;
    }

    function parseTerm() {
      let left = parseFactor();
      if (left.error) return left;

      while (peek() && peek().type === "OPERATOR" && (peek().value === "*" || peek().value === "/")) {
        const op = consume().value;
        const right = parseFactor();
        if (right.error) return right;

        if (op === "/") {
          if (right.value === 0) {
            return { error: "#DIV/0!" };
          }
          left = { value: left.value / right.value };
        } else {
          left = { value: left.value * right.value };
        }
      }
      return left;
    }

    function parseFactor() {
      const token = peek();

      if (!token) {
        return { error: "#ERROR!" };
      }

      // Unary minus
      if (token.type === "OPERATOR" && token.value === "-") {
        consume();
        const factor = parseFactor();
        if (factor.error) return factor;
        return { value: -factor.value };
      }

      // Unary plus (just consume)
      if (token.type === "OPERATOR" && token.value === "+") {
        consume();
        return parseFactor();
      }

      // Number literal
      if (token.type === "NUMBER") {
        consume();
        return { value: token.value };
      }

      // Cell reference
      if (token.type === "CELL_REF") {
        consume();
        const parsed = parseCellRef(token.value);
        if (!parsed) {
          return { error: "#REF!" };
        }
        if (parsed.row >= rows || parsed.col >= cols || parsed.row < 0 || parsed.col < 0) {
          return { error: "#REF!" };
        }
        // getCellValue returns a value. If string, try to parse? 
        // For arithmetic it usually expects numbers. 
        // Standard behavior: if cell string, 0 or error? 
        // Existing logic passed it through.
        let val = getCellValue(parsed.row, parsed.col);
        if (val === null || val === "") val = 0;
        const num = parseFloat(val);
        return { value: isNaN(num) ? 0 : num }; 
      }

      // Function Call (IDENTIFIER + LPAREN)
      if (token.type === "IDENTIFIER") {
        const funcName = consume().value; 
        
        if (peek() && peek().type === "LPAREN") {
            consume(); // eat (
            
            // For SUM/AVG we expect a Range (A1:B2)
            // Range parsing in tokens: CELL_REF COLON CELL_REF
            // or just CELL_REF (single cell range)
            
            const first = consume();
            if(!first || first.type !== "CELL_REF") return { error: "#NAME?" };
            
            let rangeStr = first.value;
            
            if (peek() && peek().type === "COLON") {
                consume(); // eat :
                const second = consume();
                if (!second || second.type !== "CELL_REF") return { error: "#NAME?" };
                rangeStr += ":" + second.value;
            }
            
            if (!peek() || peek().type !== "RPAREN") return { error: "#ERROR!" };
            consume(); // eat )
            
            if (funcName === "SUM") {
                const val = evaluator.evaluateSUM(rangeStr, context);
                if (typeof val === 'string' && val.startsWith("#")) return { error: val };
                return { value: val };
            }
            if (funcName === "AVG") {
                const val = evaluator.evaluateAVG(rangeStr, context);
                if (typeof val === 'string' && val.startsWith("#")) return { error: val };
                return { value: val };
            }
            
            return { error: "#NAME?" };
        } else {
             // Just an identifier without parens? Not supported yet (maybe named ranges later)
             return { error: "#NAME?" };
        }
      }

      // Parenthesized expression
      if (token.type === "LPAREN") {
        consume();
        const result = parseExpr();
        if (result.error) return result;

        if (!peek() || peek().type !== "RPAREN") {
          return { error: "#ERROR!" };
        }
        consume();
        return result;
      }

      return { error: "#ERROR!" };
    }

    const result = parseExpr();

    // Check for leftover tokens (malformed expression)
    if (!result.error && pos < tokens.length) {
      return { error: "#ERROR!" };
    }

    return result;
  },

  // Evaluate arithmetic expression string (without leading =)
  evaluateArithmetic(expr, context) {
    const tokenResult = this.tokenizeArithmetic(expr);
    if (tokenResult.error) {
      return tokenResult.error;
    }

    if (tokenResult.tokens.length === 0) {
      return "#ERROR!";
    }

    const evalResult = this.evaluateArithmeticExpr(tokenResult.tokens, context);
    if (evalResult.error) {
      return evalResult.error;
    }

    // Format result (avoid floating point display issues)
    const value = evalResult.value;
    if (Number.isInteger(value)) {
      return value;
    }
    // Round to reasonable precision
    return Math.round(value * 1e10) / 1e10;
  },

  // Check if expression is a valid arithmetic formula
  isArithmeticFormula(formula) {
    if (!formula || !formula.startsWith("=")) return false;
    const expr = formula.substring(1).trim();
    if (expr.length === 0) return false;

    // The new tokenizer handles everything, so if it tokenizes without error, we consider it a formula candidate.
    // However, to be safe against plain strings, we might want to check if it LOOKS like a formula.
    // But since the user explicitly typed '=', checking if it tokenizes is good enough for our logic.
    
    // Quick check: Previous logic had regexes for SUM/AVG. 
    // Now we rely on tokenizer.
    const tokenResult = this.tokenizeArithmetic(expr);
    if (tokenResult.error) return false;
    if (tokenResult.tokens.length === 0) return false;

    // Validate balanced parentheses
    let parenDepth = 0;
    for (const token of tokenResult.tokens) {
      if (token.type === "LPAREN") {
        parenDepth++;
      } else if (token.type === "RPAREN") {
        parenDepth--;
        if (parenDepth < 0) return false;
      }
    }

    return parenDepth === 0;
  },

  // Evaluate SUM(range)
  evaluateSUM(rangeStr, context) {
    const { getCellValue, rows, cols } = context;
    const range = parseRange(rangeStr);
    if (!range) return "#REF!";

    // Check if range is within grid bounds
    if (range.endRow >= rows || range.endCol >= cols) return "#REF!";

    let sum = 0;
    for (let r = range.startRow; r <= range.endRow; r++) {
      for (let c = range.startCol; c <= range.endCol; c++) {
        // getCellValue returns the visual value or calculated value.
        // If it's another formula, it should have been calculated already or will be handled by dependency order (ideally).
        // Since we are adding, we try to parse float.
        const val = getCellValue(r, c);
        const num = parseFloat(val);
        if (!isNaN(num)) {
            sum += num;
        }
      }
    }
    return sum;
  },

  // Evaluate AVG(range)
  evaluateAVG(rangeStr, context) {
    const { data, rows, cols } = context;
    const range = parseRange(rangeStr);
    if (!range) return "#REF!";

    // Check if range is within grid bounds
    if (range.endRow >= rows || range.endCol >= cols) return "#REF!";

    let sum = 0;
    let count = 0;
    for (let r = range.startRow; r <= range.endRow; r++) {
      for (let c = range.startCol; c <= range.endCol; c++) {
        // use data directly to skip empty cells distinctions
        const raw = data[r][c];
        if (raw === null || raw === undefined) continue;
        const stripped = String(raw).replace(/<[^>]*>/g, "").trim();
        if (stripped === "") continue;
        
        // If it's a formula in that cell, we should technically use its calculated value...
        // But the previous implementation used 'raw' and parsed it? 
        // If the cell contains "=SUM(A1:A2)", stripped is "=SUM(A1:A2)". parseFloat is NaN.
        // This suggests AVG only worked on literal numbers before? 
        // Let's improve it: use getCellValue if raw starts with =? 
        // Actually, 'data' usually stores the Formula string for formulas.
        // 'getCellValue' returns the computed value for that cell (if script.js is doing its job right).
        // Let's switch to checking the computed value via getCellValue for consistency, 
        // BUT we need to ignore empty cells.
        
        // Wait, context.getCellValue implementation in script.js might just be retrieving data[r][c] or the cell textContent.
        // Let's stick to the previous logic but maybe robustify it slightly if needed.
        // Original logic:
        /*
        const normalized = stripped.replace(/,/g, "");
        const num = parseFloat(normalized);
        if (isNaN(num)) continue;
        sum += num;
        count++;
        */
       
       // If the referenced cell is a formula, raw is "=..." which is NaN. 
       // We really should use computed values. 
       // However, to minimize regression risk on this specific refactor, I will adapt the referenced implementation 
       // but strictly speaking, AVG SHOULD use computed values. 
       // If I change it to use context.getCellValue(r,c), I might break "ignore empty strings" logic if getCellValue returns "" for empty.
       
       // Improving: Try computed value first.
       let val = context.getCellValue(r, c);
       if (val === null || val === "" || val === undefined) continue;
       
       // value might be "123" or 123.
       const num = parseFloat(val);
       if (!isNaN(num)) {
           sum += num;
           count++;
       }
      }
    }
    return count === 0 ? 0 : sum / count;
  },

  // Main formula evaluator
  evaluate(formula, context) {
    if (!formula || !formula.startsWith("=")) return formula;
    if (isVisualFormula(formula)) return formula;

    const expr = formula.substring(1).trim();
    
    // We strictly use the arithmetic evaluator now because it handles SUM/AVG too.
    // So we don't need the separate heavy regex checks for SUM/AVG unless we want 
    // to keep them for performance or edge cases (like legacy behavior).
    // The new parser should be superset.
    
    // Try arithmetic expression (which now includes functions)
    // We allow isArithmeticFormula to decide.
    if (this.isArithmeticFormula(formula)) {
      return this.evaluateArithmetic(expr, context);
    }

    // Unknown formula
    return "#ERROR!";
  },
};

// ==========================================
// Formula Dropdown UI Manager
// ==========================================

export const FormulaDropdownManager = {
  element: null,
  items: [],
  activeIndex: -1,
  anchor: null,
  onSelect: null, // Callback function(formulaName)

  init(onSelectCallback) {
    if (this.element) return;
    this.onSelect = onSelectCallback;

    const dropdown = document.createElement("div");
    dropdown.className = "formula-dropdown";
    dropdown.setAttribute("role", "listbox");
    dropdown.setAttribute("aria-hidden", "true");

    dropdown.addEventListener("mousedown", (event) => {
      event.preventDefault(); // Prevent blur
    });

    dropdown.addEventListener("click", (event) => {
      const item = event.target.closest(".formula-item");
      if (!item) return;
      const formulaName = item.dataset.formula;
      if (formulaName && this.onSelect) {
        this.onSelect(formulaName);
      }
    });

    document.body.appendChild(dropdown);
    this.element = dropdown;
  },

  getFormulaQuery(rawValue) {
    const match = rawValue.match(/^=\s*([A-Z]*)$/i);
    if (!match) return null;
    return match[1].toUpperCase();
  },

  getSuggestions(query) {
    if (query === null) return [];
    if (query === "") return FORMULA_SUGGESTIONS.slice();
    return FORMULA_SUGGESTIONS.filter((item) => item.name.startsWith(query));
  },

  isOpen() {
    return !!(this.element && this.element.classList.contains("open"));
  },

  setActiveItem(index) {
    this.activeIndex = index;
    this.items.forEach((item, idx) => {
      if (idx === index) {
        item.classList.add("active");
        item.setAttribute("aria-selected", "true");
        item.scrollIntoView({ block: "nearest" });
      } else {
        item.classList.remove("active");
        item.setAttribute("aria-selected", "false");
      }
    });
  },

  position(anchor) {
    if (!this.element || !anchor) return;
    const rect = anchor.getBoundingClientRect();
    const viewportWidth = window.innerWidth;
    const viewportHeight = window.innerHeight;
    const padding = 6;

    this.element.style.left = `${rect.left}px`;
    this.element.style.top = `${rect.bottom + 4}px`;

    const dropdownRect = this.element.getBoundingClientRect();
    let left = rect.left;
    let top = rect.bottom + 4;

    if (left + dropdownRect.width > viewportWidth - padding) {
      left = Math.max(padding, viewportWidth - dropdownRect.width - padding);
    }

    if (top + dropdownRect.height > viewportHeight - padding) {
      const above = rect.top - dropdownRect.height - 4;
      if (above > padding) {
        top = above;
      }
    }

    this.element.style.left = `${left}px`;
    this.element.style.top = `${top}px`;
  },

  show() {
    if (!this.element) return;
    this.element.classList.add("open");
    this.element.setAttribute("aria-hidden", "false");
  },

  hide() {
    if (!this.element) return;
    this.element.classList.remove("open");
    this.element.setAttribute("aria-hidden", "true");
    this.items = [];
    this.activeIndex = -1;
    this.anchor = null;
  },

  update(anchor, rawValue) {
    this.init(this.onSelect); // Ensure initialized

    const query = this.getFormulaQuery(rawValue);
    const suggestions = this.getSuggestions(query);

    if (!anchor || suggestions.length === 0 || query === null) {
      this.hide();
      return;
    }

    this.anchor = anchor;
    this.element.innerHTML = "";
    suggestions.forEach((item) => {
      const option = document.createElement("div");
      option.className = "formula-item";
      option.dataset.formula = item.name;
      option.setAttribute("role", "option");
      option.setAttribute("aria-selected", "false");

      const nameEl = document.createElement("div");
      nameEl.className = "formula-name";
      nameEl.textContent = item.name;

      const hintEl = document.createElement("div");
      hintEl.className = "formula-hint";
      hintEl.textContent = `${item.signature} - ${item.description}`;

      option.appendChild(nameEl);
      option.appendChild(hintEl);
      this.element.appendChild(option);
    });

    this.items = Array.from(this.element.querySelectorAll(".formula-item"));
    this.setActiveItem(0);
    this.show();
    this.position(anchor);
  },

  moveSelection(delta) {
    if (!this.items.length) return;
    let nextIndex = this.activeIndex + delta;
    if (nextIndex < 0) nextIndex = this.items.length - 1;
    if (nextIndex >= this.items.length) nextIndex = 0;
    this.setActiveItem(nextIndex);
  },

  // Get currently active item's formula name
  getActiveFormulaName() {
    if (this.activeIndex >= 0 && this.activeIndex < this.items.length) {
      return this.items[this.activeIndex].dataset.formula;
    }
    return null;
  },
};
