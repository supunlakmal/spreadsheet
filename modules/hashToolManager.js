const defaultCallbacks = {
  showToast: () => {},
};

const elements = {
  modal: null,
  input: null,
  output: null,
  errorEl: null,
  inputLen: null,
  outputLen: null,
  ratio: null,
  actionBtn: null,
  inputLabel: null,
  inputLabel: null,
  outputLabel: null,
  docs: null,
  docsToggle: null,
  modeButtons: [],
};

const MODES = {
  JSON_TO_HASH: "json-to-hash",
  HASH_TO_JSON: "hash-to-json",
};

let callbacks = { ...defaultCallbacks };
let activeMode = MODES.JSON_TO_HASH;

function clearError() {
  if (elements.errorEl) {
    elements.errorEl.textContent = "";
    elements.errorEl.classList.add("hidden");
  }
}

function showError(message) {
  if (elements.errorEl) {
    elements.errorEl.textContent = message;
    elements.errorEl.classList.remove("hidden");
  }
}

function updateStats(inputLen, outputLen) {
  if (elements.inputLen) {
    elements.inputLen.textContent = inputLen ? inputLen.toLocaleString() : "0";
  }
  if (elements.outputLen) {
    elements.outputLen.textContent = outputLen ? outputLen.toLocaleString() : "0";
  }
  if (elements.ratio) {
    if (!inputLen || !outputLen) {
      elements.ratio.textContent = "-";
    } else {
      const ratio = ((outputLen / inputLen) * 100).toFixed(1);
      elements.ratio.textContent = `${ratio}%`;
    }
  }
}

function clearOutputs() {
  if (elements.output) {
    elements.output.value = "";
  }
  updateStats(0, 0);
}

function extractHash(rawValue) {
  const trimmed = rawValue.trim();
  if (!trimmed) return "";
  const hashIndex = trimmed.indexOf("#");
  if (hashIndex >= 0) {
    return trimmed.slice(hashIndex + 1).trim();
  }
  return trimmed;
}

function getCopyLabel() {
  return activeMode === MODES.JSON_TO_HASH ? "Copy Hash" : "Copy JSON";
}

function getInputLabel() {
  return activeMode === MODES.JSON_TO_HASH ? "JSON Input" : "URL Hash or Full URL";
}

function getOutputLabel() {
  return activeMode === MODES.JSON_TO_HASH ? "URL Hash" : "JSON Output";
}

function getInputPlaceholder() {
  if (activeMode === MODES.JSON_TO_HASH) {
    return "Paste JSON here (use Copy JSON from the app for the shortest links).";
  }
  return "Paste a URL hash or full link here...";
}

function getOutputPlaceholder() {
  return activeMode === MODES.JSON_TO_HASH ? "Hash will appear here..." : "JSON will appear here...";
}

function buildFullUrl(hashValue) {
  const base = window.location.href.split("#")[0];
  const cleanHash = hashValue.replace(/^#/, "");
  return `${base}#${cleanHash}`;
}

async function copyToClipboard(text) {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch (err) {
    const temp = document.createElement("textarea");
    temp.value = text;
    temp.style.position = "fixed";
    temp.style.opacity = "0";
    document.body.appendChild(temp);
    temp.select();
    let success = false;
    try {
      success = document.execCommand("copy");
    } catch (e) {
      success = false;
    }
    document.body.removeChild(temp);
    return success;
  }
}

function setCopyFeedback(button, resetLabel) {
  if (!button) return;
  button.textContent = "Copied!";
  button.disabled = true;
  setTimeout(() => {
    button.disabled = false;
    button.textContent = resetLabel;
  }, 1800);
}

async function handleAction() {
  if (!elements.output) return;
  const rawOutput = elements.output.value.trim();
  if (!rawOutput) return;

  let textToCopy = rawOutput;
  let successLabel = "Copied!";
  let resetLabel = "Copy JSON";

  // If mode is JSON_TO_HASH, we want to copy the FULL URL, not just the hash
  if (activeMode === MODES.JSON_TO_HASH) {
    textToCopy = buildFullUrl(rawOutput); // rawOutput is the hash
    resetLabel = "Copy Link";
  }

  const success = await copyToClipboard(textToCopy);
  if (success) {
    setCopyFeedback(elements.actionBtn, resetLabel);
    if (callbacks.showToast) {
      const msg = activeMode === MODES.JSON_TO_HASH ? "Link copied to clipboard" : "JSON copied to clipboard";
      callbacks.showToast(msg, "success");
    }
  } else if (callbacks.showToast) {
    callbacks.showToast("Press Ctrl+C to copy", "warning");
  }
}

function convert() {
  if (!elements.input || !elements.output) return;

  const rawInput = elements.input.value.trim();
  clearError();

  if (!rawInput) {
    clearOutputs();
    return;
  }

  if (typeof LZString === "undefined") {
    showError("LZString is not available.");
    clearOutputs();
    return;
  }

  if (activeMode === MODES.JSON_TO_HASH) {
    try {
      const parsed = JSON.parse(rawInput);
      const normalized = JSON.stringify(parsed);
      const compressed = LZString.compressToEncodedURIComponent(normalized);
      elements.output.value = compressed;
      updateStats(normalized.length, compressed.length);
    } catch (err) {
      showError(`Invalid JSON: ${err.message}`);
      clearOutputs();
    }
    return;
  }

  const hash = extractHash(rawInput);
  if (!hash) {
    clearOutputs();
    return;
  }

  const hashUpper = hash.toUpperCase();
  if (hashUpper.startsWith("ENC:") || hashUpper.startsWith("ENC%3A")) {
    showError("Encrypted hash detected. Open the link and unlock with the password.");
    clearOutputs();
    updateStats(hash.length, 0);
    return;
  }

  const decompressed = LZString.decompressFromEncodedURIComponent(hash);
  if (!decompressed) {
    showError("Invalid or corrupted hash.");
    clearOutputs();
    updateStats(hash.length, 0);
    return;
  }

  let pretty = decompressed;
  try {
    pretty = JSON.stringify(JSON.parse(decompressed), null, 2);
  } catch (err) {
    // Keep raw string if parsing fails.
  }

  elements.output.value = pretty;
  updateStats(hash.length, pretty.length);
}

function setMode(mode) {
  if (mode !== MODES.JSON_TO_HASH && mode !== MODES.HASH_TO_JSON) return;
  activeMode = mode;

  if (elements.docs) {
    elements.docs.classList.remove("visible");
  }
  if (elements.docsToggle) {
    elements.docsToggle.textContent = "Show Schema Reference";
  }

  if (elements.input) {
    elements.input.value = "";
  }
  clearOutputs();
  clearError();

  if (elements.modeButtons.length) {
    elements.modeButtons.forEach((button) => {
      button.classList.toggle("active", button.dataset.mode === activeMode);
    });
  }

  if (elements.inputLabel) {
    elements.inputLabel.textContent = getInputLabel();
  }
  if (elements.outputLabel) {
    elements.outputLabel.textContent = getOutputLabel();
  }
  if (elements.input) {
    elements.input.placeholder = getInputPlaceholder();
  }
  if (elements.output) {
    elements.output.placeholder = getOutputPlaceholder();
  }
  if (elements.actionBtn) {
    elements.actionBtn.textContent = activeMode === MODES.JSON_TO_HASH ? "Copy Link" : "Copy JSON";
  }

  convert();
}

export const HashToolManager = {
  init(callbackOverrides = {}) {
    callbacks = { ...defaultCallbacks, ...callbackOverrides };

    elements.modal = document.getElementById("hash-tool-modal");
    elements.input = document.getElementById("hash-tool-input");
    elements.output = document.getElementById("hash-tool-output");
    elements.errorEl = document.getElementById("hash-tool-error");
    elements.inputLen = document.getElementById("hash-tool-input-len");
    elements.outputLen = document.getElementById("hash-tool-output-len");
    elements.ratio = document.getElementById("hash-tool-ratio");
    elements.actionBtn = document.getElementById("hash-tool-action-btn");
    elements.inputLabel = document.getElementById("hash-tool-input-label");
    elements.outputLabel = document.getElementById("hash-tool-output-label");
    elements.docs = document.getElementById("hash-tool-docs");
    elements.docsToggle = document.getElementById("hash-tool-docs-toggle");
    elements.modeButtons = Array.from(document.querySelectorAll(".hash-tool-mode-btn"));

    if (elements.input) {
      elements.input.addEventListener("input", convert);
    }
    if (elements.actionBtn) {
      elements.actionBtn.addEventListener("click", () => {
        handleAction();
      });
    }

    if (elements.docsToggle) {
      elements.docsToggle.addEventListener("click", () => {
        this.toggleDocs();
      });
    }

    if (elements.modeButtons.length) {
      elements.modeButtons.forEach((button) => {
        button.addEventListener("click", () => {
          setMode(button.dataset.mode);
        });
      });
    }

    setMode(activeMode);
  },

  openModal() {
    if (!elements.modal) return;
    elements.modal.classList.remove("hidden");
    clearError();
    convert();
    if (elements.input) {
      elements.input.focus();
    }
  },

  closeModal() {
    if (elements.modal) {
      elements.modal.classList.add("hidden");
    }
    if (elements.docs) {
      elements.docs.classList.remove("visible");
    }
    if (elements.docsToggle) {
      elements.docsToggle.textContent = "Show Schema Reference";
    }
    clearError();
  },

  toggleDocs() {
    if (!elements.docs || !elements.docsToggle) return;
    const isVisible = elements.docs.classList.contains("visible");
    if (isVisible) {
      elements.docs.classList.remove("visible");
      elements.docsToggle.textContent = "Show Schema Reference";
    } else {
      elements.docs.classList.add("visible");
      elements.docsToggle.textContent = "Hide Schema Reference";
    }
  },
};
