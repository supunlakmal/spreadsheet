import { showToast } from "./toastManager.js";

const defaultCallbacks = {
  onHostReady: () => {},
  onConnectionOpened: () => {},
  onInitialSync: () => {},
  onRemoteCellUpdate: () => {},
  onRemoteCursorMove: () => {},
  onSyncRequest: () => {},
  onConnectionClosed: () => {},
  onConnectionError: () => {},
};

function getPeerConstructor() {
  if (typeof window === "undefined") return null;
  return window.Peer || null;
}

export const P2PManager = {
  peer: null,
  conn: null,
  isHost: false,
  myPeerId: null,
  callbacks: { ...defaultCallbacks },

  init(callbacks = {}) {
    this.callbacks = { ...defaultCallbacks, ...callbacks };
  },

  canSend() {
    return !!(this.conn && this.conn.open);
  },

  startHosting() {
    const PeerCtor = getPeerConstructor();
    if (!PeerCtor) {
      showToast("PeerJS not available. Check network or CSP.", "error");
      return false;
    }

    this.disconnect({ silent: true });
    this.isHost = true;
    this.peer = new PeerCtor();

    this.peer.on("open", (id) => {
      this.myPeerId = id;
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

  joinSession(hostId) {
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
    this.peer = new PeerCtor();

    this.peer.on("open", () => {
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
      data: { row, col, value, formula },
    });
  },

  broadcastCursor(row, col, color) {
    return this.sendPayload({
      type: "UPDATE_CURSOR",
      data: { row, col, color },
    });
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
