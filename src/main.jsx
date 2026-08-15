import React from 'react'
import ReactDOM from 'react-dom/client'
import './responsive.css'
import App from '../app.jsx'

class ErrorBoundary extends React.Component {
  constructor(props) {
    super(props);
    this.state = { hasError: false, error: null };
  }

  static getDerivedStateFromError(error) {
    return { hasError: true, error };
  }

  componentDidCatch(error, errorInfo) {
    console.error("EDU TECH caught an error:", error, errorInfo);
  }

  render() {
    if (this.state.hasError) {
      return (
        <div
          style={{
            minHeight: "100vh",
            background: "#1a1a1a",
            color: "#f3eee1",
            display: "flex",
            flexDirection: "column",
            alignItems: "center",
            justifyContent: "center",
            padding: 24,
            fontFamily: "'Inter', sans-serif",
            textAlign: "center",
          }}
        >
          <h2 style={{ fontFamily: "'Spectral', serif", marginBottom: 8 }}>
            Something went wrong
          </h2>
          <p style={{ color: "#b8b0a0", fontSize: 13, maxWidth: 440, marginBottom: 20 }}>
            {this.state.error?.message || "An unexpected error occurred."}
          </p>
          <button
            onClick={() => {
              this.setState({ hasError: false, error: null });
              window.location.reload();
            }}
            style={{
              background: "#f3eee1",
              color: "#1a1a1a",
              border: "none",
              padding: "10px 18px",
              borderRadius: 6,
              fontSize: 13,
              fontWeight: 600,
              cursor: "pointer",
            }}
          >
            Reload Application
          </button>
        </div>
      );
    }
    return this.props.children;
  }
}

// ── window.storage polyfill (browser localStorage) ──────────────────────────
// The App uses window.storage.get/set/delete/list (Kiro IDE storage API).
// This shim maps those calls to localStorage so the app runs in any browser.

const PREFIX_SHARED = '__edutech_shared__';
const PREFIX_LOCAL  = '__edutech_local__';

window.storage = {
  get(key, shared = false) {
    try {
      const k = (shared ? PREFIX_SHARED : PREFIX_LOCAL) + key;
      const raw = localStorage.getItem(k);
      if (raw === null) return Promise.resolve(null);
      return Promise.resolve({ value: raw });
    } catch (e) {
      return Promise.resolve(null);
    }
  },

  set(key, value, shared = false) {
    try {
      const k = (shared ? PREFIX_SHARED : PREFIX_LOCAL) + key;
      localStorage.setItem(k, value);
    } catch (e) {}
    return Promise.resolve();
  },

  delete(key, shared = false) {
    try {
      const k = (shared ? PREFIX_SHARED : PREFIX_LOCAL) + key;
      localStorage.removeItem(k);
    } catch (e) {}
    return Promise.resolve();
  },

  // list all keys that start with the given prefix (shared store only)
  list(prefix, shared = false) {
    try {
      const storePrefix = (shared ? PREFIX_SHARED : PREFIX_LOCAL) + prefix;
      const keys = [];
      for (let i = 0; i < localStorage.length; i++) {
        const raw = localStorage.key(i);
        if (raw && raw.startsWith(storePrefix)) {
          // strip the store prefix so callers receive just the logical key
          keys.push(raw.slice(shared ? PREFIX_SHARED.length : PREFIX_LOCAL.length));
        }
      }
      return Promise.resolve({ keys });
    } catch (e) {
      return Promise.resolve({ keys: [] });
    }
  },
};
// ────────────────────────────────────────────────────────────────────────────

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <ErrorBoundary>
      <App />
    </ErrorBoundary>
  </React.StrictMode>,
)
