import { showToast } from "./toastManager.js";

// Default fallback ICE servers (STUN only)
const FALLBACK_ICE_SERVERS = {
  iceServers: [
    { urls: 'stun:stun.l.google.com:19302' },
    { urls: 'stun:stun1.l.google.com:19302' }
  ]
};

// Cached ICE configuration
let cachedIceConfig = null;
let cacheTimestamp = 0;
const CACHE_DURATION = 3600000; // 1 hour in milliseconds

// Fetch TURN credentials from Netlify Function (keeps API key secure)
async function fetchIceServers() {
  const now = Date.now();

  // Return cached config if still valid
  if (cachedIceConfig && (now - cacheTimestamp) < CACHE_DURATION) {
    return cachedIceConfig;
  }

  try {
    const response = await fetch('/api/turn-credentials');

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }

    const data = await response.json();

    if (data.fallback) {
      console.warn('Using fallback STUN-only configuration');
    }

    cachedIceConfig = { iceServers: data.iceServers };
    cacheTimestamp = now;

    return cachedIceConfig;
  } catch (error) {
    console.error('Failed to fetch ICE servers:', error);
    showToast('Using fallback connection (may have limited connectivity)', 'warning');
    return FALLBACK_ICE_SERVERS;
  }
}

const defaultCallbacks = {
  onHostReady: () => {},
  onPeerReady: () => {},
  onConnectionOpened: () => {},
  onInitialSync: () => {},
  onRemoteCellUpdate: () => {},
  onRemoteCursorMove: () => {},
  onSyncRequest: () => {},
  onPeerHello: () => {},
  onPeerActivity: () => {},
  onConnectionClosed: () => {},
  onConnectionError: () => {},
};

function getPeerConstructor() {
  if (typeof window === "undefined") return null;
  return window.Peer || null;
}

// UI state variables
let p2pUI = null;
let remoteCursorClass = "remote-active";

export const P2PManager = {
  peer: null,
  conn: null,
  isHost: false,
  myPeerId: null,
  callbacks: { ...defaultCallbacks },

  init(callbacks = {}) {
    this.callbacks = { ...defaultCallbacks, ...callbacks };
    // Store UI references if provided
    if (callbacks.p2pUI) {
      p2pUI = callbacks.p2pUI;
    }
    if (callbacks.remoteCursorClass) {
      remoteCursorClass = callbacks.remoteCursorClass;
    }
  },

  /**
   * Set P2P status message in UI
   * @param {string} message - Status message to display
   */
  setStatus(message) {
    if (p2pUI && p2pUI.statusEl) {
      p2pUI.statusEl.textContent = message;
    }
  },

  /**
   * Reset P2P controls to initial state
   */
  resetControls() {
    if (!p2pUI) return;

    if (p2pUI.startHostBtn) {
      p2pUI.startHostBtn.disabled = false;
      if (p2pUI.startHostLabel) {
        p2pUI.startHostBtn.innerHTML = p2pUI.startHostLabel;
      }
    }
    if (p2pUI.joinBtn) {
      p2pUI.joinBtn.disabled = false;
      if (p2pUI.joinLabel) {
        p2pUI.joinBtn.innerHTML = p2pUI.joinLabel;
      }
    }
    if (p2pUI.idDisplay) {
      p2pUI.idDisplay.classList.add("hidden");
    }
  },

  /**
   * Clear remote cursor highlighting from all cells
   */
  clearRemoteCursor() {
    document.querySelectorAll(`.${remoteCursorClass}`).forEach((el) => {
      el.classList.remove(remoteCursorClass);
      el.style.removeProperty("--peer-color");
      el.removeAttribute("data-peer");
    });
  },

  canSend() {
    return !!(this.conn && this.conn.open);
  },

  async startHosting() {
    const PeerCtor = getPeerConstructor();
    if (!PeerCtor) {
      showToast("PeerJS not available. Check network or CSP.", "error");
      return false;
    }

    this.disconnect({ silent: true });
    this.isHost = true;

    // Fetch ICE servers dynamically
    const iceConfig = await fetchIceServers();
    this.peer = new PeerCtor({ config: iceConfig });

    this.peer.on("open", (id) => {
      this.myPeerId = id;
      this.callbacks.onPeerReady(id, true);
      this.callbacks.onHostReady(id);
    });

    this.peer.on("connection", (conn) => {
      if (this.conn && this.conn.open) {
        showToast("Already connected to a peer", "warning");
        conn.close();
        return;
      }
      this.handleConnection(conn);
    });

    this.peer.on("error", (err) => {
      console.error(err);
      showToast(`P2P Error: ${err.type || "unknown"}`, "error");
      this.callbacks.onConnectionError(err);
    });

    this.peer.on("disconnected", () => {
      showToast("Disconnected from signaling server", "warning");
    });

    return true;
  },

  async joinSession(hostId) {
    const PeerCtor = getPeerConstructor();
    if (!PeerCtor) {
      showToast("PeerJS not available. Check network or CSP.", "error");
      return false;
    }

    if (!hostId) {
      showToast("Host ID required", "warning");
      return false;
    }

    this.disconnect({ silent: true });
    this.isHost = false;

    // Fetch ICE servers dynamically
    const iceConfig = await fetchIceServers();
    this.peer = new PeerCtor({ config: iceConfig });

    this.peer.on("open", (id) => {
      this.myPeerId = id;
      this.callbacks.onPeerReady(id, false);
      const conn = this.peer.connect(hostId);
      this.handleConnection(conn);
    });

    this.peer.on("error", (err) => {
      console.error(err);
      const msg = err && err.type === "peer-unavailable" ? "Host not found. Check the ID." : "Could not connect to host.";
      showToast(msg, "error");
      this.callbacks.onConnectionError(err);
    });

    return true;
  },

  handleConnection(conn) {
    this.conn = conn;

    conn.on("open", () => {
      this.callbacks.onConnectionOpened(this.isHost);
    });

    conn.on("data", (payload) => {
      this.handleIncomingData(payload);
    });

    conn.on("close", () => {
      this.conn = null;
      this.callbacks.onConnectionClosed();
    });

    conn.on("error", (err) => {
      console.error(err);
      showToast("Connection error", "error");
      this.callbacks.onConnectionError(err);
    });
  },

  handleIncomingData(payload) {
    if (!payload || typeof payload !== "object") return;
    const type = payload.type;
    if (!type) return;

    switch (type) {
      case "INITIAL_SYNC":
      case "FULL_SYNC":
        this.callbacks.onInitialSync(payload.data, type);
        break;
      case "UPDATE_CELL":
        this.callbacks.onRemoteCellUpdate(payload.data || {});
        break;
      case "UPDATE_CURSOR":
        this.callbacks.onRemoteCursorMove(payload.data || {});
        break;
      case "SYNC_REQUEST":
        this.callbacks.onSyncRequest();
        break;
      case "PEER_HELLO":
        this.callbacks.onPeerHello(payload.data || {});
        break;
      case "PEER_ACTIVITY":
        this.callbacks.onPeerActivity(payload.data || {});
        break;
      default:
        console.warn("Unknown P2P message:", type);
    }
  },

  sendPayload(payload) {
    if (!this.canSend()) return false;
    try {
      this.conn.send(payload);
      return true;
    } catch (err) {
      console.error(err);
      showToast("Failed to send update", "error");
      return false;
    }
  },

  sendInitialSync(fullState) {
    return this.sendPayload({ type: "INITIAL_SYNC", data: fullState });
  },

  sendFullSync(fullState) {
    return this.sendPayload({ type: "FULL_SYNC", data: fullState });
  },

  requestFullSync() {
    return this.sendPayload({ type: "SYNC_REQUEST" });
  },

  broadcastCellUpdate(row, col, value, formula) {
    return this.sendPayload({
      type: "UPDATE_CELL",
      data: { row, col, value, formula, senderId: this.myPeerId },
    });
  },

  broadcastCursor(row, col, color) {
    return this.sendPayload({
      type: "UPDATE_CURSOR",
      data: { row, col, color, senderId: this.myPeerId },
    });
  },

  sendPeerHello(peerMeta) {
    return this.sendPayload({ type: "PEER_HELLO", data: peerMeta });
  },

  sendPeerActivity(activity) {
    return this.sendPayload({ type: "PEER_ACTIVITY", data: activity });
  },

  disconnect({ silent = false } = {}) {
    if (this.conn) {
      try {
        this.conn.close();
      } catch (e) {}
    }
    if (this.peer) {
      try {
        this.peer.destroy();
      } catch (e) {}
    }

    this.peer = null;
    this.conn = null;
    this.myPeerId = null;
    this.isHost = false;

    if (!silent) {
      this.callbacks.onConnectionClosed();
    }
  },
};
