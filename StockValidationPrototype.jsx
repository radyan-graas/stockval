import React, { useState, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  Package,
  Truck,
  Link2,
  Play,
  Download,
  CheckCircle2,
  AlertTriangle,
  XCircle,
  Info,
  ChevronRight,
} from "lucide-react";

// ---------------------------------------------------------------------------
// Design tokens
// ---------------------------------------------------------------------------
const C = {
  ink: "#12181F",
  slate: "#4A5568",
  faint: "#8A94A3",
  line: "#E4E8ED",
  panel: "#F7F9FB",
  bg: "#FFFFFF",
  accent: "#2F5DAA",
  accentSoft: "#EAF0FA",
  good: "#1E8E5A",
  goodSoft: "#E7F5EE",
  warn: "#B7791F",
  warnSoft: "#FCF3DF",
  bad: "#C13A3A",
  badSoft: "#FBEAEA",
};

const mono = { fontFamily: "'JetBrains Mono', 'IBM Plex Mono', monospace" };

// ---------------------------------------------------------------------------
// Parsing helpers (mirrors the Python core logic)
// ---------------------------------------------------------------------------
function readFileAsSheetRows(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const isCsv = file.name.toLowerCase().endsWith(".csv");
        const wb = isCsv
          ? XLSX.read(e.target.result, { type: "string" })
          : XLSX.read(e.target.result, { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });
        resolve(rows);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    const isCsv = file.name.toLowerCase().endsWith(".csv");
    if (isCsv) reader.readAsText(file);
    else reader.readAsArrayBuffer(file);
  });
}

function norm(v) {
  if (v === null || v === undefined) return "";
  return String(v).trim();
}
function normLower(v) {
  return norm(v).toLowerCase();
}
function toNum(v) {
  const n = parseFloat(v);
  return isNaN(n) ? 0 : n;
}

function parseMarketplace(rows) {
  let headerRowIdx = -1;
  for (let i = 0; i < Math.min(rows.length, 20); i++) {
    const rowLower = (rows[i] || []).map(normLower);
    if (rowLower.includes("parent sku") && rowLower.includes("sku")) {
      headerRowIdx = i;
      break;
    }
  }
  if (headerRowIdx === -1) {
    throw new Error("Could not find the header row (expected 'Parent SKU' and 'SKU' columns).");
  }
  const header = (rows[headerRowIdx] || []).map(normLower);
  const find = (label) => {
    const idx = header.indexOf(label);
    if (idx === -1) throw new Error(`Could not find required column '${label}' in Marketplace file.`);
    return idx;
  };
  const idxName = find("product name");
  const idxParent = find("parent sku");
  const idxSku = find("sku");
  const idxStock = find("stock");
  const idxPid = find("product id");

  const bySku = new Map();
  for (let i = headerRowIdx + 1; i < rows.length; i++) {
    const row = rows[i] || [];
    const pid = row[idxPid];
    if (pid === null || pid === undefined || norm(pid) === "" || isNaN(parseFloat(pid))) continue;
    const parentSku = norm(row[idxParent]);
    const skuF = norm(row[idxSku]);
    const finalSku = skuF !== "" ? skuF : parentSku;
    if (finalSku === "") continue;
    const skuSource = skuF !== "" ? "SKU (F)" : "Parent SKU (E)";
    const stock = toNum(row[idxStock]);
    const name = norm(row[idxName]);

    if (bySku.has(finalSku)) {
      const rec = bySku.get(finalSku);
      rec.marketplaceStock += stock;
    } else {
      bySku.set(finalSku, {
        sku: finalSku,
        productName: name,
        skuSource,
        marketplaceStock: stock,
      });
    }
  }
  return Array.from(bySku.values());
}

function parseWarehouse(rows) {
  const header = (rows[0] || []).map(normLower);
  const findAny = (...cands) => {
    for (const c of cands) {
      const idx = header.indexOf(c);
      if (idx !== -1) return idx;
    }
    throw new Error(`Could not find any of [${cands.join(", ")}] in Warehouse file columns.`);
  };
  const idxSku = findAny("internal reference", "sku");
  const idxQty = findAny("available quantity", "quantity", "stock");

  const bySku = new Map();
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i] || [];
    const sku = norm(row[idxSku]);
    if (sku === "") continue;
    const qty = toNum(row[idxQty]);
    bySku.set(sku, (bySku.get(sku) || 0) + qty);
  }
  return bySku;
}

function parsePlatform(rows) {
  const header = (rows[0] || []).map(normLower);
  const idxSku = header.indexOf("sellersku");
  if (idxSku === -1) throw new Error("Could not find 'sellerSKU' column in Platform file.");

  let idxQty = header.findIndex((h) => h.includes("siawms") && h.includes("quantity"));
  let qtyColName = idxQty !== -1 ? rows[0][idxQty] : null;
  if (idxQty === -1) {
    if (header.length > 10) {
      idxQty = 10; // fallback: column K by position
      qtyColName = rows[0][10];
    } else {
      throw new Error("Could not find 'SiAWMS-1 quantity' column in Platform file.");
    }
  }

  const bySku = new Map();
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i] || [];
    const sku = norm(row[idxSku]);
    if (sku === "") continue;
    const qty = toNum(row[idxQty]);
    bySku.set(sku, (bySku.get(sku) || 0) + qty);
  }
  return { map: bySku, colName: qtyColName };
}

function buildComparison(marketplaceRows, warehouseMap, platformMap) {
  const rows = marketplaceRows.map((mp) => {
    const foundWh = warehouseMap.has(mp.sku);
    const foundPl = platformMap.has(mp.sku);
    const whStock = foundWh ? warehouseMap.get(mp.sku) : 0;
    const plStock = foundPl ? platformMap.get(mp.sku) : 0;

    const mpDiff = mp.marketplaceStock - whStock;
    const plDiff = plStock - whStock;

    const mpStatus = !foundWh ? "SKU Not Found" : mpDiff === 0 ? "Match" : "Mismatch";
    const plStatus = !foundPl ? "SKU Not Found" : plDiff === 0 ? "Match" : "Mismatch";

    const discrepancyCount = (mpStatus !== "Match" ? 1 : 0) + (plStatus !== "Match" ? 1 : 0);
    const overallStatus = discrepancyCount === 0 ? "OK" : "Discrepancy";

    return {
      sku: mp.sku,
      productName: mp.productName,
      skuSource: mp.skuSource,
      marketplaceStock: mp.marketplaceStock,
      platformStock: plStock,
      warehouseStock: whStock,
      correctStock: whStock,
      mpDiff,
      plDiff,
      mpStatus,
      plStatus,
      discrepancyCount,
      overallStatus,
    };
  });
  rows.sort((a, b) => b.discrepancyCount - a.discrepancyCount || a.sku.localeCompare(b.sku));
  return rows;
}

// ---------------------------------------------------------------------------
// Export helpers
// ---------------------------------------------------------------------------
function exportToXlsx(rows, filename, sheetName) {
  const data = rows.map((r) => ({
    SKU: r.sku,
    "Product Name": r.productName,
    SKU_Source: r.skuSource,
    Marketplace_Stock: r.marketplaceStock,
    Platform_Stock: r.platformStock,
    Warehouse_Stock: r.warehouseStock,
    "Correct_Stock (Warehouse anchor)": r.correctStock,
    MP_vs_WH_Diff: r.mpDiff,
    PLT_vs_WH_Diff: r.plDiff,
    Marketplace_Status: r.mpStatus,
    Platform_Status: r.plStatus,
    Discrepancy_Count: r.discrepancyCount,
    Overall_Status: r.overallStatus,
  }));
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName.slice(0, 31));
  XLSX.writeFile(wb, filename);
}

// ---------------------------------------------------------------------------
// UI subcomponents
// ---------------------------------------------------------------------------
function StatusPill({ status }) {
  const map = {
    Match: { c: C.good, bg: C.goodSoft, Icon: CheckCircle2 },
    OK: { c: C.good, bg: C.goodSoft, Icon: CheckCircle2 },
    Mismatch: { c: C.bad, bg: C.badSoft, Icon: XCircle },
    Discrepancy: { c: C.bad, bg: C.badSoft, Icon: XCircle },
    "SKU Not Found": { c: C.warn, bg: C.warnSoft, Icon: AlertTriangle },
  };
  const cfg = map[status] || { c: C.slate, bg: C.panel, Icon: Info };
  const { c, bg, Icon } = cfg;
  return (
    <span
      style={{
        display: "inline-flex",
        alignItems: "center",
        gap: 5,
        padding: "3px 9px",
        borderRadius: 20,
        background: bg,
        color: c,
        fontSize: 12,
        fontWeight: 600,
        whiteSpace: "nowrap",
      }}
    >
      <Icon size={12} strokeWidth={2.5} />
      {status}
    </span>
  );
}

function UploadCard({ icon: Icon, label, sub, file, onFile, accept }) {
  const [dragOver, setDragOver] = useState(false);
  const inputId = `upload-${label.replace(/\s+/g, "-")}`;

  return (
    <div
      onDragOver={(e) => {
        e.preventDefault();
        setDragOver(true);
      }}
      onDragLeave={() => setDragOver(false)}
      onDrop={(e) => {
        e.preventDefault();
        setDragOver(false);
        if (e.dataTransfer.files?.[0]) onFile(e.dataTransfer.files[0]);
      }}
      style={{
        border: `1.5px dashed ${dragOver ? C.accent : file ? C.good : C.line}`,
        borderRadius: 12,
        background: file ? C.goodSoft : dragOver ? C.accentSoft : C.panel,
        padding: "18px 16px",
        transition: "all 0.15s ease",
        cursor: "pointer",
        flex: 1,
        minWidth: 0,
      }}
      onClick={() => document.getElementById(inputId).click()}
    >
      <input
        id={inputId}
        type="file"
        accept={accept}
        style={{ display: "none" }}
        onChange={(e) => e.target.files?.[0] && onFile(e.target.files[0])}
      />
      <div style={{ display: "flex", alignItems: "flex-start", gap: 12 }}>
        <div
          style={{
            width: 36,
            height: 36,
            borderRadius: 9,
            background: file ? C.good : C.accent,
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            flexShrink: 0,
          }}
        >
          <Icon size={18} color="#fff" strokeWidth={2} />
        </div>
        <div style={{ minWidth: 0 }}>
          <div style={{ fontWeight: 650, fontSize: 14, color: C.ink }}>{label}</div>
          <div style={{ fontSize: 12, color: C.faint, marginTop: 1 }}>{sub}</div>
          <div
            style={{
              marginTop: 8,
              fontSize: 12.5,
              color: file ? C.good : C.faint,
              fontWeight: file ? 600 : 400,
              overflow: "hidden",
              textOverflow: "ellipsis",
              whiteSpace: "nowrap",
              ...(file ? mono : {}),
            }}
          >
            {file ? `✓ ${file.name}` : "Click or drop file here"}
          </div>
        </div>
      </div>
    </div>
  );
}

function Metric({ label, value, tone }) {
  const toneColor = tone === "good" ? C.good : tone === "bad" ? C.bad : C.ink;
  return (
    <div
      style={{
        background: C.panel,
        border: `1px solid ${C.line}`,
        borderRadius: 10,
        padding: "14px 16px",
        flex: 1,
        minWidth: 120,
      }}
    >
      <div style={{ fontSize: 11.5, color: C.faint, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.4 }}>
        {label}
      </div>
      <div style={{ fontSize: 26, fontWeight: 700, color: toneColor, marginTop: 4, ...mono }}>{value}</div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Main App
// ---------------------------------------------------------------------------
export default function StockValidationApp() {
  const [mpFile, setMpFile] = useState(null);
  const [plFile, setPlFile] = useState(null);
  const [whFile, setWhFile] = useState(null);
  const [running, setRunning] = useState(false);
  const [error, setError] = useState(null);
  const [result, setResult] = useState(null); // { comparison, platformCol }
  const [tab, setTab] = useState("discrepancy");

  const canRun = mpFile && plFile && whFile && !running;

  const runValidation = useCallback(async () => {
    setRunning(true);
    setError(null);
    try {
      const [mpRows, plRows, whRows] = await Promise.all([
        readFileAsSheetRows(mpFile),
        readFileAsSheetRows(plFile),
        readFileAsSheetRows(whFile),
      ]);
      const marketplaceRows = parseMarketplace(mpRows);
      const warehouseMap = parseWarehouse(whRows);
      const { map: platformMap, colName } = parsePlatform(plRows);
      const comparison = buildComparison(marketplaceRows, warehouseMap, platformMap);
      setResult({ comparison, platformCol: colName });
      setTab("discrepancy");
    } catch (e) {
      setError(e.message || String(e));
      setResult(null);
    } finally {
      setRunning(false);
    }
  }, [mpFile, plFile, whFile]);

  const comparison = result?.comparison || [];
  const total = comparison.length;
  const okCount = comparison.filter((r) => r.overallStatus === "OK").length;
  const discCount = comparison.filter((r) => r.overallStatus === "Discrepancy").length;
  const mpMismatch = comparison.filter((r) => r.mpStatus !== "Match").length;
  const plMismatch = comparison.filter((r) => r.plStatus !== "Match").length;
  const discrepancyRows = comparison.filter((r) => r.overallStatus === "Discrepancy");
  const shownRows = tab === "discrepancy" ? discrepancyRows : comparison;

  return (
    <div
      style={{
        fontFamily: "'Inter', -apple-system, sans-serif",
        background: C.bg,
        minHeight: "100vh",
        color: C.ink,
      }}
    >
      <div style={{ maxWidth: 1120, margin: "0 auto", padding: "28px 20px 60px" }}>
        {/* Header */}
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 6 }}>
          <div
            style={{
              width: 34,
              height: 34,
              borderRadius: 9,
              background: C.ink,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
            }}
          >
            <Package size={18} color="#fff" />
          </div>
          <h1 style={{ fontSize: 21, fontWeight: 700, margin: 0 }}>Stock Validation</h1>
        </div>
        <p style={{ fontSize: 13.5, color: C.slate, margin: "0 0 22px", maxWidth: 640, lineHeight: 1.5 }}>
          Cross-checks stock across Marketplace, Platform, and Warehouse. Marketplace SKUs are the
          reference list; Warehouse is the source of truth when numbers disagree.
        </p>

        {/* Rule strip */}
        <div
          style={{
            display: "flex",
            gap: 18,
            flexWrap: "wrap",
            padding: "12px 16px",
            background: C.accentSoft,
            borderRadius: 10,
            marginBottom: 24,
            fontSize: 12.5,
            color: C.accent,
          }}
        >
          {[
            "SKU = Marketplace col F, falls back to Parent SKU (col E)",
            "Platform stock = SiAWMS-1 quantity (col K) only",
            "Warehouse = anchor for the correct stock value",
          ].map((t, i) => (
            <div key={i} style={{ display: "flex", alignItems: "center", gap: 6 }}>
              <ChevronRight size={13} />
              <span>{t}</span>
            </div>
          ))}
        </div>

        {/* Uploads */}
        <div style={{ display: "flex", gap: 14, flexWrap: "wrap", marginBottom: 18 }}>
          <UploadCard
            icon={Package}
            label="Marketplace"
            sub="Shopee export (.xlsx)"
            file={mpFile}
            onFile={setMpFile}
            accept=".xlsx,.xls"
          />
          <UploadCard
            icon={Link2}
            label="Platform"
            sub="Platform inventory (.csv / .xlsx)"
            file={plFile}
            onFile={setPlFile}
            accept=".csv,.xlsx,.xls"
          />
          <UploadCard
            icon={Truck}
            label="Warehouse"
            sub="Warehouse report (.xlsx) — anchor"
            file={whFile}
            onFile={setWhFile}
            accept=".xlsx,.xls"
          />
        </div>

        <button
          onClick={runValidation}
          disabled={!canRun}
          style={{
            display: "inline-flex",
            alignItems: "center",
            gap: 8,
            padding: "10px 20px",
            borderRadius: 9,
            border: "none",
            background: canRun ? C.ink : C.line,
            color: canRun ? "#fff" : C.faint,
            fontWeight: 650,
            fontSize: 14,
            cursor: canRun ? "pointer" : "not-allowed",
          }}
        >
          <Play size={15} />
          {running ? "Running..." : "Run Stock Validation"}
        </button>

        {error && (
          <div
            style={{
              marginTop: 16,
              padding: "12px 14px",
              background: C.badSoft,
              color: C.bad,
              borderRadius: 9,
              fontSize: 13,
              display: "flex",
              gap: 8,
              alignItems: "flex-start",
            }}
          >
            <XCircle size={16} style={{ flexShrink: 0, marginTop: 1 }} />
            <span>{error}</span>
          </div>
        )}

        {result && (
          <div style={{ marginTop: 30 }}>
            {/* Metrics */}
            <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 8 }}>
              <Metric label="Total SKU" value={total} />
              <Metric label="Matched" value={okCount} tone="good" />
              <Metric label="Discrepancies" value={discCount} tone="bad" />
              <Metric label="Marketplace mismatches" value={mpMismatch} />
              <Metric label="Platform mismatches" value={plMismatch} />
            </div>
            <div style={{ fontSize: 12, color: C.faint, marginBottom: 20 }}>
              Platform stock computed from column: <span style={mono}>{result.platformCol}</span>
            </div>

            {/* Tabs */}
            <div style={{ display: "flex", gap: 4, borderBottom: `1px solid ${C.line}`, marginBottom: 14 }}>
              {[
                { key: "discrepancy", label: `Discrepancy Result (${discrepancyRows.length})` },
                { key: "all", label: `Working Process — all SKU (${comparison.length})` },
              ].map((t) => (
                <button
                  key={t.key}
                  onClick={() => setTab(t.key)}
                  style={{
                    padding: "8px 14px",
                    fontSize: 13,
                    fontWeight: 600,
                    background: "none",
                    border: "none",
                    borderBottom: tab === t.key ? `2px solid ${C.ink}` : "2px solid transparent",
                    color: tab === t.key ? C.ink : C.faint,
                    cursor: "pointer",
                  }}
                >
                  {t.label}
                </button>
              ))}
              <div style={{ flex: 1 }} />
              <button
                onClick={() =>
                  exportToXlsx(
                    tab === "discrepancy" ? discrepancyRows : comparison,
                    tab === "discrepancy"
                      ? "Stock_Validation_Discrepancy_Result.xlsx"
                      : "Stock_Validation_Working_Process.xlsx",
                    tab === "discrepancy" ? "Discrepancies" : "Working Process"
                  )
                }
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 6,
                  padding: "7px 13px",
                  fontSize: 12.5,
                  fontWeight: 600,
                  borderRadius: 7,
                  border: `1px solid ${C.line}`,
                  background: C.bg,
                  color: C.ink,
                  cursor: "pointer",
                  alignSelf: "center",
                }}
              >
                <Download size={13} />
                Download .xlsx
              </button>
            </div>

            {/* Table */}
            <div style={{ overflowX: "auto", border: `1px solid ${C.line}`, borderRadius: 10 }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
                <thead>
                  <tr style={{ background: C.panel, textAlign: "left" }}>
                    {[
                      "SKU",
                      "Product",
                      "SKU Source",
                      "Marketplace",
                      "Platform",
                      "Warehouse (anchor)",
                      "MP−WH",
                      "PLT−WH",
                      "MP Status",
                      "PLT Status",
                    ].map((h) => (
                      <th
                        key={h}
                        style={{
                          padding: "9px 12px",
                          fontWeight: 650,
                          color: C.slate,
                          borderBottom: `1px solid ${C.line}`,
                          whiteSpace: "nowrap",
                        }}
                      >
                        {h}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {shownRows.length === 0 ? (
                    <tr>
                      <td colSpan={10} style={{ padding: "28px 12px", textAlign: "center", color: C.faint }}>
                        No rows to show — no discrepancies found. 🎉
                      </td>
                    </tr>
                  ) : (
                    shownRows.map((r, i) => (
                      <tr key={r.sku} style={{ borderBottom: `1px solid ${C.line}`, background: i % 2 ? C.bg : C.panel }}>
                        <td style={{ padding: "8px 12px", ...mono, fontWeight: 600 }}>{r.sku}</td>
                        <td style={{ padding: "8px 12px", color: C.slate, maxWidth: 180, overflow: "hidden", textOverflow: "ellipsis" }}>
                          {r.productName}
                        </td>
                        <td style={{ padding: "8px 12px", color: C.faint, whiteSpace: "nowrap" }}>{r.skuSource}</td>
                        <td style={{ padding: "8px 12px", ...mono }}>{r.marketplaceStock}</td>
                        <td style={{ padding: "8px 12px", ...mono }}>{r.platformStock}</td>
                        <td style={{ padding: "8px 12px", ...mono, fontWeight: 650 }}>{r.warehouseStock}</td>
                        <td style={{ padding: "8px 12px", ...mono, color: r.mpDiff === 0 ? C.faint : C.bad }}>
                          {r.mpDiff > 0 ? `+${r.mpDiff}` : r.mpDiff}
                        </td>
                        <td style={{ padding: "8px 12px", ...mono, color: r.plDiff === 0 ? C.faint : C.bad }}>
                          {r.plDiff > 0 ? `+${r.plDiff}` : r.plDiff}
                        </td>
                        <td style={{ padding: "8px 12px" }}>
                          <StatusPill status={r.mpStatus} />
                        </td>
                        <td style={{ padding: "8px 12px" }}>
                          <StatusPill status={r.plStatus} />
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {!result && !error && (
          <div style={{ marginTop: 30, fontSize: 13, color: C.faint, display: "flex", alignItems: "center", gap: 8 }}>
            <Info size={15} />
            Upload the three files above, then run validation. Everything happens locally in your browser.
          </div>
        )}
      </div>
    </div>
  );
}
