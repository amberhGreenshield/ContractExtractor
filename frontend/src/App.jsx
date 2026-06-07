import { useState, useEffect } from "react";

const API = "http://localhost:8000";

const STATUS = { IDLE: "idle", LOADING: "loading", EXTRACTING: "extracting", DONE: "done", ERROR: "error" };

function formatBytes(b) {
  if (!b) return "—";
  if (b < 1024) return `${b} B`;
  if (b < 1024 * 1024) return `${(b / 1024).toFixed(1)} KB`;
  return `${(b / (1024 * 1024)).toFixed(1)} MB`;
}

function cleanPath(path) {
  if (!path) return "";
  const parts = path.split("/drives/");
  if (parts.length > 1) return "/" + parts[1].split("/root:/")[1] || path;
  return path;
}

export default function App() {
  const [files, setFiles] = useState([]);
  const [selected, setSelected] = useState(new Set());
  const [results, setResults] = useState([]);
  const [status, setStatus] = useState(STATUS.IDLE);
  const [error, setError] = useState("");
  const [progress, setProgress] = useState(0);

  useEffect(() => {
    loadFiles();
  }, []);

  async function loadFiles() {
    setStatus(STATUS.LOADING);
    setError("");
    try {
      const res = await fetch(`${API}/contracts/list`);
      if (!res.ok) throw new Error(await res.text());
      const data = await res.json();
      setFiles(data.files || []);
      setStatus(STATUS.IDLE);
    } catch (e) {
      setError(`Failed to load files: ${e.message}`);
      setStatus(STATUS.ERROR);
    }
  }

  function toggleFile(id) {
    setSelected(prev => {
      const next = new Set(prev);
      next.has(id) ? next.delete(id) : next.add(id);
      return next;
    });
  }

  function toggleAll() {
    if (selected.size === files.length) setSelected(new Set());
    else setSelected(new Set(files.map(f => f.id)));
  }

  async function runExtraction() {
    if (selected.size === 0) return;
    setStatus(STATUS.EXTRACTING);
    setError("");
    setResults([]);
    setProgress(0);

    const ids = [...selected];
    try {
      const res = await fetch(`${API}/contracts/extract`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ file_ids: ids }),
      });
      if (!res.ok) throw new Error(await res.text());
      const data = await res.json();
      setResults(data.results || []);
      setStatus(STATUS.DONE);
    } catch (e) {
      setError(`Extraction failed: ${e.message}`);
      setStatus(STATUS.ERROR);
    }
  }

  async function downloadExcel() {
    setError("");
    try {
      const ids = results.length > 0
        ? results.map(r => r.file_id)
        : [...selected];

      const res = await fetch(`${API}/contracts/export-excel`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ file_ids: ids }),
      });
      if (!res.ok) throw new Error(await res.text());
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "contract_analysis.xlsx";
      a.click();
      URL.revokeObjectURL(url);
    } catch (e) {
      setError(`Export failed: ${e.message}`);
    }
  }

  const isExtracting = status === STATUS.EXTRACTING;
  const isLoading = status === STATUS.LOADING;
  const hasResults = results.length > 0;

  return (
    <div style={{
      minHeight: "100vh",
      background: "linear-gradient(135deg, #0f1e35 0%, #1a3a5c 50%, #0f2840 100%)",
      fontFamily: "'DM Sans', 'Segoe UI', sans-serif",
      color: "#e8edf3",
    }}>
      {/* Header */}
      <header style={{
        background: "rgba(255,255,255,0.04)",
        borderBottom: "1px solid rgba(255,255,255,0.08)",
        backdropFilter: "blur(12px)",
        padding: "0 2.5rem",
        height: 70,
        display: "flex",
        alignItems: "center",
        justifyContent: "space-between",
        position: "sticky",
        top: 0,
        zIndex: 100,
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
          <div style={{
            width: 36, height: 36,
            background: "linear-gradient(135deg, #2979ff, #00bcd4)",
            borderRadius: 10,
            display: "flex", alignItems: "center", justifyContent: "center",
            fontSize: 18,
            boxShadow: "0 4px 12px rgba(41,121,255,0.4)",
          }}>📑</div>
          <div>
            <div style={{ fontWeight: 700, fontSize: 17, letterSpacing: "-0.3px" }}>Contract Extractor</div>
            <div style={{ fontSize: 11, color: "#6b8cae", letterSpacing: "0.5px" }}>POWERED BY AZURE OPENAI</div>
          </div>
        </div>

        <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
          <button onClick={loadFiles} disabled={isLoading} style={btnStyle("ghost")}>
            {isLoading ? "⟳ Loading…" : "⟳ Refresh"}
          </button>
          {(hasResults || selected.size > 0) && (
            <button onClick={downloadExcel} style={btnStyle("accent")}>
              ⬇ Export Excel
            </button>
          )}
        </div>
      </header>

      <main style={{ maxWidth: 1200, margin: "0 auto", padding: "2rem 2.5rem" }}>

        {/* Error banner */}
        {error && (
          <div style={{
            background: "rgba(198,40,40,0.15)",
            border: "1px solid rgba(198,40,40,0.4)",
            borderRadius: 10,
            padding: "12px 18px",
            marginBottom: 20,
            fontSize: 13,
            color: "#ef9a9a",
          }}>
            ⚠ {error}
          </div>
        )}

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 24 }}>

          {/* Left: File browser */}
          <div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
              <h2 style={{ margin: 0, fontSize: 15, fontWeight: 700, color: "#c8daf0" }}>
                SharePoint Contracts
                <span style={{ marginLeft: 10, fontSize: 12, color: "#6b8cae", fontWeight: 400 }}>
                  {files.length} files found
                </span>
              </h2>
              {files.length > 0 && (
                <button onClick={toggleAll} style={btnStyle("ghost", true)}>
                  {selected.size === files.length ? "Deselect all" : "Select all"}
                </button>
              )}
            </div>

            <div style={{
              background: "rgba(255,255,255,0.04)",
              border: "1px solid rgba(255,255,255,0.08)",
              borderRadius: 14,
              overflow: "hidden",
              maxHeight: 520,
              overflowY: "auto",
            }}>
              {isLoading ? (
                <div style={{ padding: "3rem", textAlign: "center", color: "#6b8cae" }}>
                  <div style={{ fontSize: 28, marginBottom: 10 }}>⟳</div>
                  Loading from SharePoint…
                </div>
              ) : files.length === 0 ? (
                <div style={{ padding: "3rem", textAlign: "center", color: "#6b8cae" }}>
                  <div style={{ fontSize: 36, marginBottom: 10 }}>📁</div>
                  No contract files found.<br />
                  <span style={{ fontSize: 12 }}>Check your SharePoint folder path in .env</span>
                </div>
              ) : (
                files.map((file, i) => (
                  <FileRow
                    key={file.id}
                    file={file}
                    checked={selected.has(file.id)}
                    onToggle={() => toggleFile(file.id)}
                    striped={i % 2 === 0}
                  />
                ))
              )}
            </div>

            <div style={{ marginTop: 16, display: "flex", justifyContent: "flex-end" }}>
              <button
                onClick={runExtraction}
                disabled={selected.size === 0 || isExtracting}
                style={{
                  ...btnStyle("primary"),
                  opacity: selected.size === 0 ? 0.4 : 1,
                  cursor: selected.size === 0 ? "not-allowed" : "pointer",
                  padding: "12px 28px",
                  fontSize: 14,
                  fontWeight: 700,
                }}
              >
                {isExtracting
                  ? `⟳ Extracting ${selected.size} contracts…`
                  : `⚡ Extract ${selected.size > 0 ? selected.size : ""} Contract${selected.size !== 1 ? "s" : ""}`}
              </button>
            </div>
          </div>

          {/* Right: Results */}
          <div>
            <div style={{ marginBottom: 14 }}>
              <h2 style={{ margin: 0, fontSize: 15, fontWeight: 700, color: "#c8daf0" }}>
                Extracted Data
                {hasResults && (
                  <span style={{ marginLeft: 10, fontSize: 12, color: "#6b8cae", fontWeight: 400 }}>
                    {results.length} contracts processed
                  </span>
                )}
              </h2>
            </div>

            <div style={{
              background: "rgba(255,255,255,0.04)",
              border: "1px solid rgba(255,255,255,0.08)",
              borderRadius: 14,
              overflow: "hidden",
              maxHeight: 520,
              overflowY: "auto",
            }}>
              {isExtracting ? (
                <div style={{ padding: "3rem", textAlign: "center", color: "#6b8cae" }}>
                  <div style={{ fontSize: 36, marginBottom: 14 }}>🤖</div>
                  <div style={{ fontWeight: 600, color: "#aac4e4", marginBottom: 8 }}>
                    GPT-4 is reading your contracts…
                  </div>
                  <div style={{ fontSize: 12 }}>Extracting vendor names and costs</div>
                  <div style={{
                    width: "80%", height: 4,
                    background: "rgba(255,255,255,0.1)",
                    borderRadius: 99, margin: "20px auto 0",
                    overflow: "hidden",
                  }}>
                    <div style={{
                      height: "100%",
                      background: "linear-gradient(90deg, #2979ff, #00bcd4)",
                      borderRadius: 99,
                      animation: "pulse-bar 1.5s ease-in-out infinite",
                      width: "60%",
                    }} />
                  </div>
                </div>
              ) : !hasResults ? (
                <div style={{ padding: "3rem", textAlign: "center", color: "#6b8cae" }}>
                  <div style={{ fontSize: 36, marginBottom: 10 }}>📊</div>
                  Select contracts and click Extract<br />
                  <span style={{ fontSize: 12 }}>Results will appear here</span>
                </div>
              ) : (
                results.map((r, i) => (
                  <ResultCard key={r.file_id} result={r} striped={i % 2 === 0} />
                ))
              )}
            </div>

            {hasResults && (
              <div style={{
                marginTop: 14,
                background: "rgba(41,121,255,0.08)",
                border: "1px solid rgba(41,121,255,0.2)",
                borderRadius: 10,
                padding: "12px 16px",
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
              }}>
                <div style={{ fontSize: 13, color: "#90b4d8" }}>
                  ✓ Extraction complete — {results.filter(r => !r.error).length} successful,{" "}
                  {results.filter(r => r.error).length} errors
                </div>
                <button onClick={downloadExcel} style={btnStyle("accent", true)}>
                  ⬇ Download Excel
                </button>
              </div>
            )}
          </div>
        </div>
      </main>

      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=DM+Mono:wght@400;500&display=swap');
        * { box-sizing: border-box; }
        ::-webkit-scrollbar { width: 6px; }
        ::-webkit-scrollbar-track { background: transparent; }
        ::-webkit-scrollbar-thumb { background: rgba(255,255,255,0.12); border-radius: 3px; }
        @keyframes pulse-bar {
          0% { transform: translateX(-100%); }
          100% { transform: translateX(250%); }
        }
      `}</style>
    </div>
  );
}

function FileRow({ file, checked, onToggle, striped }) {
  const ext = file.name.split(".").pop().toUpperCase();
  const extColor = ext === "PDF" ? "#ef5350" : ext === "DOCX" ? "#1e88e5" : "#43a047";

  return (
    <div
      onClick={onToggle}
      style={{
        display: "flex",
        alignItems: "center",
        gap: 12,
        padding: "10px 16px",
        cursor: "pointer",
        background: checked
          ? "rgba(41,121,255,0.12)"
          : striped ? "rgba(255,255,255,0.02)" : "transparent",
        borderLeft: checked ? "3px solid #2979ff" : "3px solid transparent",
        transition: "all 0.15s",
        borderBottom: "1px solid rgba(255,255,255,0.05)",
      }}
    >
      <input
        type="checkbox"
        checked={checked}
        onChange={onToggle}
        onClick={e => e.stopPropagation()}
        style={{ accentColor: "#2979ff", width: 15, height: 15, flexShrink: 0 }}
      />
      <span style={{
        fontSize: 9, fontWeight: 700,
        background: extColor,
        color: "white",
        borderRadius: 4,
        padding: "2px 5px",
        fontFamily: "DM Mono, monospace",
        flexShrink: 0,
      }}>{ext}</span>
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{
          fontSize: 13, fontWeight: 500,
          color: checked ? "#a8c8f0" : "#c8daf0",
          overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
        }}>
          {file.name}
        </div>
        <div style={{
          fontSize: 11, color: "#4d6b85",
          overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
        }}>
          {cleanPath(file.path)} · {formatBytes(file.size)}
        </div>
      </div>
    </div>
  );
}

function ResultCard({ result, striped }) {
  const hasError = !!result.error;

  return (
    <div style={{
      padding: "14px 18px",
      borderBottom: "1px solid rgba(255,255,255,0.05)",
      background: hasError
        ? "rgba(198,40,40,0.07)"
        : striped ? "rgba(255,255,255,0.02)" : "transparent",
      borderLeft: hasError ? "3px solid #ef5350" : "3px solid #2979ff",
    }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 8 }}>
        <div style={{
          fontSize: 12, fontWeight: 600,
          color: hasError ? "#ef9a9a" : "#90b4d8",
          overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
          maxWidth: "60%",
        }}>
          {result.file_name}
        </div>
        {hasError ? (
          <span style={{ fontSize: 11, color: "#ef5350", flexShrink: 0 }}>⚠ Error</span>
        ) : (
          <span style={{
            fontSize: 13, fontWeight: 700,
            color: "#69f0ae",
            fontFamily: "DM Mono, monospace",
            flexShrink: 0,
          }}>
            {result.contract_value || "—"}
          </span>
        )}
      </div>

      {hasError ? (
        <div style={{ fontSize: 11, color: "#ef9a9a", marginTop: 4 }}>{result.error}</div>
      ) : (
        <div style={{ display: "flex", gap: 16, marginTop: 6 }}>
          <div>
            <span style={{ fontSize: 10, color: "#4d6b85", display: "block" }}>VENDOR</span>
            <span style={{ fontSize: 13, color: "#e8edf3", fontWeight: 500 }}>
              {result.vendor_name || "Not found"}
            </span>
          </div>
          {result.notes && result.notes !== "—" && (
            <div style={{ flex: 1 }}>
              <span style={{ fontSize: 10, color: "#4d6b85", display: "block" }}>NOTES</span>
              <span style={{ fontSize: 11, color: "#6b8cae" }}>{result.notes}</span>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function btnStyle(variant, small = false) {
  const base = {
    border: "none",
    borderRadius: 8,
    cursor: "pointer",
    fontFamily: "inherit",
    fontWeight: 600,
    transition: "all 0.15s",
    padding: small ? "7px 14px" : "9px 18px",
    fontSize: small ? 12 : 13,
  };
  if (variant === "primary") return { ...base, background: "linear-gradient(135deg, #2979ff, #1565c0)", color: "white", boxShadow: "0 4px 14px rgba(41,121,255,0.35)" };
  if (variant === "accent") return { ...base, background: "linear-gradient(135deg, #00bcd4, #0097a7)", color: "white", boxShadow: "0 4px 14px rgba(0,188,212,0.3)" };
  if (variant === "ghost") return { ...base, background: "rgba(255,255,255,0.06)", color: "#90b4d8", border: "1px solid rgba(255,255,255,0.1)" };
  return base;
}

function cleanPath(path) {
  if (!path) return "";
  const m = path.match(/root:(.*)/);
  return m ? m[1] : path;
}
