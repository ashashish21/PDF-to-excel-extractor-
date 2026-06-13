import { useState, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

// ── tiny design tokens ──────────────────────────────────────────────
const C = {
  bg: "#0f1117",
  surface: "#181c27",
  border: "#252a38",
  accent: "#5b8af5",
  accentDim: "#1e2d56",
  accentHover: "#7aa3ff",
  text: "#e8eaf0",
  muted: "#6b7280",
  success: "#34d399",
  error: "#f87171",
  warning: "#fbbf24",
};

const font = "'Inter', 'Segoe UI', sans-serif";

// ── helpers ──────────────────────────────────────────────────────────
function fileToBase64(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result.split(",")[1]);
    r.onerror = () => rej(new Error("File read failed"));
    r.readAsDataURL(file);
  });
}

async function extractDataWithClaude(pdfBase64, headers) {
  const prompt = `You are a precise data extraction assistant.

I will give you a PDF document. Extract structured data from it and map it to the following column headers:
${headers.map((h, i) => `${i + 1}. ${h}`).join("\n")}

Rules:
- Return ONLY a valid JSON array of objects. No explanation, no markdown, no code fences.
- Each object must have keys exactly matching the headers above.
- Extract as many rows as you can find in the document.
- If a field value is not found for a row, use null.
- Dates should be ISO strings. Numbers should be numbers, not strings.
- Example output shape: [{"${headers[0] || "Field1"}": "value", ...}]`;

  const response = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: 4000,
      messages: [
        {
          role: "user",
          content: [
            {
              type: "document",
              source: { type: "base64", media_type: "application/pdf", data: pdfBase64 },
            },
            { type: "text", text: prompt },
          ],
        },
      ],
    }),
  });

  if (!response.ok) {
    const err = await response.json().catch(() => ({}));
    throw new Error(err?.error?.message || `API error ${response.status}`);
  }

  const data = await response.json();
  const raw = data.content.map((b) => b.text || "").join("");
  const clean = raw.replace(/```json|```/g, "").trim();

  try {
    return JSON.parse(clean);
  } catch {
    const match = clean.match(/\[[\s\S]*\]/);
    if (match) return JSON.parse(match[0]);
    throw new Error("Could not parse AI response as JSON.");
  }
}

function buildXLSX(rows, headers) {
  const wb = XLSX.utils.book_new();
  const wsData = [
    headers,
    ...rows.map((row) => headers.map((h) => row[h] ?? "")),
  ];
  const ws = XLSX.utils.aoa_to_sheet(wsData);

  // Style header row widths
  ws["!cols"] = headers.map(() => ({ wch: 22 }));

  XLSX.utils.book_append_sheet(wb, ws, "Extracted Data");
  return wb;
}

// ── sub-components ───────────────────────────────────────────────────
function Tag({ label, onRemove }) {
  return (
    <span style={{
      display: "inline-flex", alignItems: "center", gap: 6,
      background: C.accentDim, color: C.accent,
      border: `1px solid ${C.accent}33`,
      borderRadius: 6, padding: "4px 10px", fontSize: 13, fontFamily: font,
    }}>
      {label}
      <button onClick={onRemove} style={{
        background: "none", border: "none", cursor: "pointer",
        color: C.muted, fontSize: 15, lineHeight: 1, padding: 0,
        display: "flex", alignItems: "center",
      }}>×</button>
    </span>
  );
}

function DropZone({ file, onFile, disabled }) {
  const [drag, setDrag] = useState(false);
  const inputRef = useRef();

  const handle = (f) => {
    if (f && f.type === "application/pdf") onFile(f);
  };

  return (
    <div
      onClick={() => !disabled && inputRef.current.click()}
      onDragOver={(e) => { e.preventDefault(); !disabled && setDrag(true); }}
      onDragLeave={() => setDrag(false)}
      onDrop={(e) => { e.preventDefault(); setDrag(false); handle(e.dataTransfer.files[0]); }}
      style={{
        border: `2px dashed ${drag ? C.accent : file ? C.success : C.border}`,
        borderRadius: 12, padding: "32px 24px", textAlign: "center",
        cursor: disabled ? "default" : "pointer",
        background: drag ? C.accentDim : file ? "#0d2a1f" : C.surface,
        transition: "all .2s", userSelect: "none",
      }}
    >
      <input
        ref={inputRef} type="file" accept="application/pdf"
        style={{ display: "none" }}
        onChange={(e) => handle(e.target.files[0])}
      />
      <div style={{ fontSize: 32, marginBottom: 8 }}>
        {file ? "📄" : "⬆️"}
      </div>
      <div style={{ color: file ? C.success : C.text, fontWeight: 600, fontSize: 15 }}>
        {file ? file.name : "Drop PDF here or click to browse"}
      </div>
      {file && (
        <div style={{ color: C.muted, fontSize: 12, marginTop: 4 }}>
          {(file.size / 1024).toFixed(1)} KB
        </div>
      )}
      {!file && (
        <div style={{ color: C.muted, fontSize: 13, marginTop: 6 }}>
          Only .pdf files accepted
        </div>
      )}
    </div>
  );
}

function DataTable({ headers, rows }) {
  return (
    <div style={{ overflowX: "auto", borderRadius: 10, border: `1px solid ${C.border}` }}>
      <table style={{
        width: "100%", borderCollapse: "collapse",
        fontFamily: font, fontSize: 13,
      }}>
        <thead>
          <tr style={{ background: C.accentDim }}>
            <th style={{ ...thStyle, color: C.muted, width: 40 }}>#</th>
            {headers.map((h) => (
              <th key={h} style={{ ...thStyle, color: C.accent }}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, i) => (
            <tr key={i} style={{ background: i % 2 === 0 ? C.surface : "#161a25" }}>
              <td style={{ ...tdStyle, color: C.muted }}>{i + 1}</td>
              {headers.map((h) => (
                <td key={h} style={tdStyle}>
                  {row[h] == null ? <span style={{ color: C.muted }}>—</span> : String(row[h])}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

const thStyle = {
  padding: "10px 14px", textAlign: "left", fontWeight: 600,
  borderBottom: `1px solid ${C.border}`, whiteSpace: "nowrap",
};
const tdStyle = {
  padding: "9px 14px", color: C.text, borderBottom: `1px solid ${C.border}22`,
  maxWidth: 240, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
};

// ── main app ─────────────────────────────────────────────────────────
export default function App() {
  const [headers, setHeaders] = useState([]);
  const [headerInput, setHeaderInput] = useState("");
  const [pdfFile, setPdfFile] = useState(null);
  const [status, setStatus] = useState("idle"); // idle | loading | done | error
  const [extractedRows, setExtractedRows] = useState([]);
  const [errorMsg, setErrorMsg] = useState("");
  const [sheetName, setSheetName] = useState("Extracted Data");

  const addHeader = useCallback(() => {
    const val = headerInput.trim();
    if (val && !headers.includes(val)) {
      setHeaders((h) => [...h, val]);
    }
    setHeaderInput("");
  }, [headerInput, headers]);

  const onKeyDown = (e) => {
    if (e.key === "Enter" || e.key === ",") { e.preventDefault(); addHeader(); }
  };

  const handleExtract = async () => {
    if (!pdfFile || headers.length === 0) return;
    setStatus("loading");
    setErrorMsg("");
    try {
      const b64 = await fileToBase64(pdfFile);
      const rows = await extractDataWithClaude(b64, headers);
      setExtractedRows(rows);
      setStatus("done");
    } catch (e) {
      setErrorMsg(e.message);
      setStatus("error");
    }
  };

  const handleDownload = () => {
    const wb = buildXLSX(extractedRows, headers);
    // Rename sheet
    if (wb.SheetNames[0]) wb.SheetNames[0] = sheetName;
    if (wb.Sheets["Extracted Data"] && sheetName !== "Extracted Data") {
      wb.Sheets[sheetName] = wb.Sheets["Extracted Data"];
      delete wb.Sheets["Extracted Data"];
    }
    XLSX.writeFile(wb, "extracted_data.xlsx");
  };

  const reset = () => {
    setStatus("idle");
    setExtractedRows([]);
    setErrorMsg("");
    setPdfFile(null);
  };

  const canExtract = pdfFile && headers.length > 0 && status !== "loading";

  return (
    <div style={{
      minHeight: "100vh", background: C.bg, color: C.text,
      fontFamily: font, padding: "0 0 60px",
    }}>
      {/* Header */}
      <div style={{
        borderBottom: `1px solid ${C.border}`,
        padding: "22px 32px", display: "flex", alignItems: "center", gap: 14,
        background: C.surface,
      }}>
        <span style={{ fontSize: 26 }}>🗂️</span>
        <div>
          <div style={{ fontWeight: 700, fontSize: 18, letterSpacing: "-0.3px" }}>
            PDF → Excel Extractor
          </div>
          <div style={{ color: C.muted, fontSize: 13 }}>
            AI-powered data extraction from any PDF into a structured spreadsheet
          </div>
        </div>
      </div>

      <div style={{ maxWidth: 860, margin: "0 auto", padding: "36px 24px", display: "flex", flexDirection: "column", gap: 28 }}>

        {/* Step 1: Define headers */}
        <Section step="1" title="Define your Excel columns">
          <p style={{ color: C.muted, fontSize: 14, margin: "0 0 14px" }}>
            Type each column header you want in your spreadsheet and press Enter or comma to add it.
          </p>
          <div style={{ display: "flex", gap: 8 }}>
            <input
              value={headerInput}
              onChange={(e) => setHeaderInput(e.target.value)}
              onKeyDown={onKeyDown}
              placeholder="e.g. Invoice Number, Date, Amount…"
              style={{
                flex: 1, background: C.bg, border: `1px solid ${C.border}`,
                borderRadius: 8, padding: "10px 14px", color: C.text,
                fontSize: 14, fontFamily: font, outline: "none",
              }}
              disabled={status === "loading"}
            />
            <button
              onClick={addHeader}
              disabled={!headerInput.trim() || status === "loading"}
              style={btnStyle(C.accent, !headerInput.trim())}
            >
              Add
            </button>
          </div>
          {headers.length > 0 && (
            <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginTop: 14 }}>
              {headers.map((h) => (
                <Tag key={h} label={h} onRemove={() => setHeaders(headers.filter((x) => x !== h))} />
              ))}
            </div>
          )}
        </Section>

        {/* Step 2: Upload PDF */}
        <Section step="2" title="Upload your PDF">
          <DropZone
            file={pdfFile}
            onFile={setPdfFile}
            disabled={status === "loading"}
          />
          {pdfFile && (
            <button
              onClick={() => setPdfFile(null)}
              style={{ ...btnStyle(C.muted, false), marginTop: 10, fontSize: 12, padding: "6px 14px" }}
            >
              Remove file
            </button>
          )}
        </Section>

        {/* Step 3: Sheet name + Extract */}
        <Section step="3" title="Configure & Extract">
          <div style={{ display: "flex", gap: 12, alignItems: "flex-end", flexWrap: "wrap" }}>
            <div style={{ flex: 1, minWidth: 180 }}>
              <label style={{ fontSize: 13, color: C.muted, display: "block", marginBottom: 6 }}>
                Sheet name
              </label>
              <input
                value={sheetName}
                onChange={(e) => setSheetName(e.target.value)}
                style={{
                  width: "100%", background: C.bg, border: `1px solid ${C.border}`,
                  borderRadius: 8, padding: "10px 14px", color: C.text,
                  fontSize: 14, fontFamily: font, outline: "none", boxSizing: "border-box",
                }}
              />
            </div>
            <button
              onClick={handleExtract}
              disabled={!canExtract}
              style={{ ...btnStyle(C.accent, !canExtract), padding: "11px 28px", fontSize: 15 }}
            >
              {status === "loading" ? "Extracting…" : "⚡ Extract Data"}
            </button>
          </div>

          {!canExtract && status === "idle" && (
            <div style={{ color: C.muted, fontSize: 13, marginTop: 10 }}>
              {!pdfFile && headers.length === 0 ? "Add column headers and upload a PDF to begin." :
               !pdfFile ? "Upload a PDF to continue." :
               "Add at least one column header to continue."}
            </div>
          )}
        </Section>

        {/* Loading */}
        {status === "loading" && (
          <div style={{
            background: C.surface, border: `1px solid ${C.border}`,
            borderRadius: 12, padding: "28px 24px", textAlign: "center",
          }}>
            <Spinner />
            <div style={{ color: C.text, fontWeight: 600, marginTop: 14 }}>
              Claude is reading your PDF…
            </div>
            <div style={{ color: C.muted, fontSize: 13, marginTop: 6 }}>
              Extracting and mapping {headers.length} column{headers.length !== 1 ? "s" : ""}
            </div>
          </div>
        )}

        {/* Error */}
        {status === "error" && (
          <div style={{
            background: "#2a1414", border: `1px solid ${C.error}44`,
            borderRadius: 12, padding: "20px 24px",
          }}>
            <div style={{ color: C.error, fontWeight: 600, marginBottom: 6 }}>⚠ Extraction failed</div>
            <div style={{ color: C.muted, fontSize: 13 }}>{errorMsg}</div>
            <button onClick={reset} style={{ ...btnStyle(C.error, false), marginTop: 14, fontSize: 13 }}>
              Try again
            </button>
          </div>
        )}

        {/* Results */}
        {status === "done" && extractedRows.length > 0 && (
          <div style={{
            background: C.surface, border: `1px solid ${C.border}`,
            borderRadius: 12, padding: "24px",
          }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 18, flexWrap: "wrap", gap: 12 }}>
              <div>
                <span style={{ color: C.success, fontWeight: 700, fontSize: 16 }}>
                  ✓ {extractedRows.length} row{extractedRows.length !== 1 ? "s" : ""} extracted
                </span>
                <span style={{ color: C.muted, fontSize: 13, marginLeft: 10 }}>
                  across {headers.length} columns
                </span>
              </div>
              <div style={{ display: "flex", gap: 10 }}>
                <button onClick={reset} style={{ ...btnStyle(C.muted, false), fontSize: 13, padding: "8px 16px" }}>
                  Start over
                </button>
                <button onClick={handleDownload} style={{ ...btnStyle(C.success, false), fontSize: 13, padding: "8px 18px" }}>
                  ⬇ Download .xlsx
                </button>
              </div>
            </div>
            <DataTable headers={headers} rows={extractedRows} />
          </div>
        )}

        {status === "done" && extractedRows.length === 0 && (
          <div style={{
            background: "#1c1a0d", border: `1px solid ${C.warning}44`,
            borderRadius: 12, padding: "20px 24px", color: C.warning,
          }}>
            ⚠ The AI couldn't find any matching data in this PDF for your headers. Try adjusting the column names to better match the document's content.
            <button onClick={reset} style={{ ...btnStyle(C.warning, false), marginTop: 12, fontSize: 13, display: "block" }}>
              Try again
            </button>
          </div>
        )}

      </div>
    </div>
  );
}

// ── micro components ─────────────────────────────────────────────────
function Section({ step, title, children }) {
  return (
    <div style={{
      background: C.surface, border: `1px solid ${C.border}`,
      borderRadius: 14, padding: "24px 24px 22px",
    }}>
      <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 16 }}>
        <span style={{
          background: C.accentDim, color: C.accent,
          fontWeight: 700, fontSize: 12, borderRadius: 6,
          padding: "3px 9px", letterSpacing: "0.5px",
        }}>
          STEP {step}
        </span>
        <span style={{ fontWeight: 600, fontSize: 16 }}>{title}</span>
      </div>
      {children}
    </div>
  );
}

function Spinner() {
  return (
    <div style={{
      width: 40, height: 40, borderRadius: "50%",
      border: `3px solid ${C.border}`,
      borderTopColor: C.accent,
      animation: "spin 0.8s linear infinite",
      margin: "0 auto",
    }}>
      <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
    </div>
  );
}

function btnStyle(color, disabled) {
  return {
    background: disabled ? C.border : color === C.accent ? C.accentDim : "transparent",
    color: disabled ? C.muted : color,
    border: `1px solid ${disabled ? C.border : color}`,
    borderRadius: 8, padding: "10px 20px",
    fontFamily: font, fontWeight: 600, fontSize: 14,
    cursor: disabled ? "not-allowed" : "pointer",
    transition: "all .15s",
    opacity: disabled ? 0.6 : 1,
  };
}
