import React, { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import { getDocument, GlobalWorkerOptions } from "pdfjs-dist";
import pdfWorker from "pdfjs-dist/build/pdf.worker?url";
import { supabase } from "./lib/supabaseClient";

GlobalWorkerOptions.workerSrc = pdfWorker;

const HARD_CODED_ITEMS = [
  ["Crystalline Allulose", "IW741"],
  ["Monk Fruit Concentrate", "IW796"],
  ["Erythritol FN Powder", "IN954"],
  ["Erythritol STD Granular", "IW810"],
  ["Organic Coconut Sugar", "IW816"],
  ["Wonder Monday Crumbs - Erythritol", "IW815"],
  ["Wonder Monday Crumbs - Coconut Sugar", "IW797"],
  ["Wonder Monday Crumbs - Cane Sugar", "IW793"],
  ["Natural Graham Flavor", "IN838"],
  ["White Chocolate Raspberry Flavor", "IN843"],
  ["Maple Pecan Flavor", "IN840"],
  ["Salted Caramel Type Extract", "IN864"],
  ["Vanilla Bean", "IN926"],
  ["Lemon Meringue Nat Flavor", "IN928"],
  ["Enrobe", "IW777"],
  ["SB 2-Pack Bites Labels", "PFW1020-12721"],
  ["PB 2-Pack Bites Labels", "PFW1020-12730"],
  ['Wonder Monday 3" Lids', "PP5002"],
  ['Wonder Monday Blank 3" Bases', "PP5003"],
  ['Wonder Monday 3" Clear Base', "PP5004"],
  ["Target 6 pack carton", "PC2894"],
  ["SRP Mastercase 12pk", "PC2891"],
].map(([description, item_code], index) => ({
  id: `hardcoded-${index}`,
  description,
  item_code,
  row_order: index,
}));

const MONTHS = Array.from({ length: 12 }, (_, i) => {
  const d = new Date();
  d.setMonth(d.getMonth() - i);
  return (
    d.toLocaleString("en-US", { month: "long" }) +
    ` '${String(d.getFullYear()).slice(-2)}`
  );
});

const BATCHES = Array.from({ length: 10 }, (_, i) => `Batch ${i + 1}`);

function norm(v) {
  return String(v ?? "").trim().toLowerCase();
}

function num(v) {
  if (v === null || v === undefined || v === "") return 0;
  const n = Number(String(v).replace(/,/g, "").replace("%", ""));
  return Number.isFinite(n) ? n : 0;
}

function positiveNum(v) {
  return Math.abs(num(v));
}

function getProductionCodeFromFileName(fileName) {
  const match = String(fileName || "").match(/W(?:M)?\s*(\d+)(?:[.\-_]\d+)?/i);
  return match ? match[1] : "PDF";
}

function varianceValue(wm, cbf) {
  wm = Number(wm || 0);
  cbf = Number(cbf || 0);

  if (wm === 0 && cbf === 0) return 0;

  if (wm === 0 && cbf > 0) return 100;

  if (wm > 0 && cbf === 0) return -100;

  return ((cbf - wm) / wm) * 100;
}

function variancePct(wm, cbf) {
  return `${varianceValue(wm, cbf).toFixed(2)}%`;
}

function lbsDiff(wm, cbf) {
  return cbf - wm;
}

function allowedTolerance(itemCode) {
  const code = String(itemCode || "").trim().toUpperCase();
  if (code === "IW810" || code === "IN838") return 10.15;
  return 3;
}

function summaryClass(itemCode, variance) {
  if (Math.abs(variance) <= 0.0049) return "summary-neutral";
  return Math.abs(variance) > allowedTolerance(itemCode)
    ? "summary-red"
    : "summary-green";
}

function batchSummaryText(name, wm, cbf) {
  const diff = lbsDiff(wm, cbf);
  const pct = varianceValue(wm, cbf);

  if (Math.abs(diff) <= 0.005) return `${name}: No variance.`;

  if (diff > 0) {
    return `${name}: There is a ${pct.toFixed(
      2
    )}% surplus. CBF used ${diff.toFixed(
      2
    )} lbs more than WM expected.`;
  }

  return `${name}: There is a ${pct.toFixed(
    2
  )}% variance. CBF used ${diff.toFixed(2)} lbs less than expected.`;
}

function extractFirstEachQty(text) {
  const matches = [
    ...String(text || "").matchAll(/([0-9][0-9,]*\.?[0-9]*)\s*each\b/gi),
  ];

  if (!matches.length) return null;
  return num(matches[0][1]);
}

function extractLastLbQty(text) {
  const matches = [
    ...String(text || "").matchAll(/([0-9][0-9,]*\.?[0-9]*)\s*lb\.?\b/gi),
  ];

  if (!matches.length) return null;
  return num(matches[matches.length - 1][1]);
}

function isPackagingCode(code) {
  return /^(P|PL|PP|PC|PS)/i.test(String(code || ""));
}

async function extractPdfLines(file) {
  const buffer = await file.arrayBuffer();
  const pdf = await getDocument({ data: buffer }).promise;
  const lines = [];

  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const content = await page.getTextContent();

    const buckets = new Map();

    content.items.forEach((item) => {
      const x = item.transform[4];
      const y = Math.round(item.transform[5]);
      const key = `${p}-${y}`;

      if (!buckets.has(key)) buckets.set(key, []);
      buckets.get(key).push({ x, text: item.str });
    });

    [...buckets.entries()]
      .sort((a, b) => Number(b[0].split("-")[1]) - Number(a[0].split("-")[1]))
      .forEach(([, parts]) => {
        const text = parts
          .sort((a, b) => a.x - b.x)
          .map((p) => p.text)
          .join(" ")
          .replace(/\s+/g, " ")
          .trim();

        if (text) lines.push(text);
      });
  }

  return lines;
}

async function parseWmExcel(file, items) {
  const buffer = await file.arrayBuffer();
  const wb = XLSX.read(buffer, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];

  const rows = XLSX.utils.sheet_to_json(ws, {
    header: 1,
    raw: false,
    defval: "",
  });

  const headerRowIndex = rows.findIndex((row) =>
    row.some((cell) => /^Used\s*\(/i.test(String(cell)))
  );

  if (headerRowIndex < 0) throw new Error("No Used(...) columns found.");

  const headers = rows[headerRowIndex].map((h) => String(h || "").trim());
  const itemLookup = new Map(items.map((item) => [norm(item.description), item]));

  const groups = headers
    .map((header, index) => {
      const match = header.match(/^Used\s*\(([^)]+)\)/i);
      if (!match) return null;
      return { prod: match[1].trim(), usedCol: index };
    })
    .filter(Boolean);

  const parsed = [];

  for (let r = headerRowIndex + 1; r < rows.length; r++) {
    const description = String(rows[r][0] || "").trim();
    if (!description || norm(description) === "cbf") continue;

    const matched = itemLookup.get(norm(description));
    if (!matched) continue;

    groups.forEach((g) => {
      parsed.push({
        description: matched.description,
        item_code: matched.item_code,
        production_code: g.prod,
        wm_usage: positiveNum(rows[r][g.usedCol]),
        cbf_usage: null,
        row_order: matched.row_order,
      });
    });
  }

  return parsed;
}

function findPdfUsageForCode(lines, code) {
  let total = 0;
  let found = false;
  const codeUpper = String(code || "").toUpperCase();
  const packaging = isPackagingCode(codeUpper);

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i] || "";

    if (!line.toUpperCase().includes(codeUpper)) continue;

    let usage = null;

    if (packaging) {
      usage = extractFirstEachQty(line);

      if (usage === null) {
        const forwardBlock = lines.slice(i, i + 3).join(" ");
        usage = extractFirstEachQty(forwardBlock);
      }
    } else {
      for (let j = i - 1; j >= Math.max(0, i - 3); j--) {
        const previousLine = lines[j] || "";

        if (
          previousLine.toLowerCase().includes("historical") ||
          previousLine.toLowerCase().includes("cost")
        ) {
          continue;
        }

        const lbValue = extractLastLbQty(previousLine);

        if (lbValue !== null) {
          usage = lbValue;
          break;
        }
      }
    }

    if (usage !== null) {
      total += usage;
      found = true;
    }
  }

  return found ? total : null;
}

async function parseCbfPdf(file, items) {
  const productionCode = getProductionCodeFromFileName(file.name);
  const lines = await extractPdfLines(file);

  const parsed = [];

  items.forEach((item) => {
    const usage = findPdfUsageForCode(lines, item.item_code);

    if (usage !== null) {
      parsed.push({
        description: item.description,
        item_code: item.item_code,
        production_code: productionCode,
        wm_usage: null,
        cbf_usage: usage,
        row_order: item.row_order,
      });
    }
  });

  if (parsed.length === 0) {
    throw new Error("No matching CBF item numbers found in this PDF.");
  }

  return parsed;
}

export default function App() {
  const [page, setPage] = useState("checker");
  const [items] = useState(HARD_CODED_ITEMS);
  const [bomRows, setBomRows] = useState([]);
  const [modal, setModal] = useState(null);
  const [files, setFiles] = useState([]);
  const [message, setMessage] = useState("");
  const [batchVisible, setBatchVisible] = useState(true);

  const [form, setForm] = useState({
    month: MONTHS[0],
    batch_mode: "new",
    batch_name: "Batch 1",
  });

  useEffect(() => {
    loadData();
  }, [form.month, form.batch_name]);

  async function loadData() {
    const { data } = await supabase
      .from("bom_usage_rows")
      .select("*")
      .eq("month", form.month)
      .eq("batch", form.batch_name);

    if (data) setBomRows(data);
  }

  async function uploadBom() {
    if (!files.length) return alert("Choose a file first.");

    try {
      const bomType = modal;

      if (bomType === "WM") {
        await supabase
          .from("bom_usage_rows")
          .delete()
          .eq("month", form.month)
          .eq("batch", form.batch_name)
          .eq("bom_type", "WM");
      }

      let allRows = [];

      for (const file of files) {
        const parsed =
          bomType === "WM"
            ? await parseWmExcel(file, items)
            : await parseCbfPdf(file, items);

        const cleanMonth = form.month.replaceAll(" ", "-").replaceAll("'", "");
        const cleanBatch = form.batch_name.replaceAll(" ", "");
        const storagePath = `${bomType}/${cleanMonth}/${cleanBatch}/${Date.now()}-${file.name}`;

        const { error: storageError } = await supabase.storage
          .from("bom-files")
          .upload(storagePath, file, { upsert: true });

        if (storageError) throw storageError;

        const { data: uploadRow, error: uploadError } = await supabase
          .from("bom_uploads")
          .insert({
            bom_type: bomType,
            month: form.month,
            batch: form.batch_name,
            file_name: file.name,
            storage_path: storagePath,
          })
          .select()
          .single();

        if (uploadError) throw uploadError;

        const rows = parsed.map((r) => ({
          upload_id: uploadRow.id,
          bom_type: bomType,
          month: form.month,
          batch: form.batch_name,
          item_code: r.item_code,
          description: r.description,
          production_code: r.production_code,
          usage_lbs: bomType === "WM" ? r.wm_usage : r.cbf_usage,
          wm_usage: r.wm_usage,
          cbf_usage: r.cbf_usage,
          row_order: r.row_order,
        }));

        allRows = [...allRows, ...rows];
      }

      const { error } = await supabase.from("bom_usage_rows").insert(allRows);
      if (error) throw error;

      await loadData();

      setMessage(`${bomType} BOM uploaded: ${allRows.length} values parsed.`);
      setModal(null);
      setFiles([]);
      setBatchVisible(true);
    } catch (err) {
      console.error(err);
      alert(err.message || "Upload failed.");
    }
  }

  function downloadExcel() {
    const rows = [];

    visibleWideRows.forEach((row) => {
      const record = {
        "Item Description": row.description,
        "CBF Item #": row.item_code,
      };

      productionCodes.forEach((code) => {
        const v = row.values[code] || { wm: 0, cbf: 0 };
        record[`Used (${code})`] = Number(v.wm.toFixed(2));
        record[`CBF-${code}`] = Number(v.cbf.toFixed(2));
        record[`Variance ${code}`] = variancePct(v.wm, v.cbf);
      });

      const batchTotal = batchTotalsByRow[row.item_code] || {
        wm: 0,
        cbf: 0,
        diff: 0,
        variance: 0,
        summary: "",
      };

      record["Total WM"] = Number(batchTotal.wm.toFixed(2));
      record["Total CBF"] = Number(batchTotal.cbf.toFixed(2));
      record["Total Variance"] = `${batchTotal.variance.toFixed(2)}%`;
      record["Total LBS"] = `${batchTotal.diff.toFixed(2)} lbs`;
      record["Summary"] = batchTotal.summary;

      rows.push(record);
    });

    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(wb, ws, form.batch_name);

    XLSX.writeFile(
      wb,
      `BOM_Checker_${form.month
        .replaceAll(" ", "_")
        .replaceAll("'", "")}_${form.batch_name.replaceAll(" ", "_")}.xlsx`
    );
  }

  const productionCodes = useMemo(() => {
    return [...new Set(bomRows.map((r) => r.production_code))]
      .filter(Boolean)
      .sort((a, b) => Number(a) - Number(b));
  }, [bomRows]);

  const wideRows = useMemo(() => {
    const grouped = new Map();

    items.forEach((item) => {
      grouped.set(item.item_code, {
        description: item.description,
        item_code: item.item_code,
        row_order: item.row_order,
        values: {},
      });
    });

    bomRows.forEach((r) => {
      const row = grouped.get(r.item_code);
      if (!row) return;

      const current = row.values[r.production_code] || { wm: 0, cbf: 0 };

      if (r.bom_type === "WM") {
        current.wm += Number(r.wm_usage ?? r.usage_lbs ?? 0);
      }

      if (r.bom_type === "CBF") {
        current.cbf += Number(r.cbf_usage ?? r.usage_lbs ?? 0);
      }

      row.values[r.production_code] = current;
    });

    return [...grouped.values()].sort((a, b) => a.row_order - b.row_order);
  }, [items, bomRows]);

  const batchTotalsByRow = useMemo(() => {
    const output = {};

    wideRows.forEach((row) => {
      let wm = 0;
      let cbf = 0;

      productionCodes.forEach((code) => {
        wm += Number(row.values[code]?.wm || 0);
        cbf += Number(row.values[code]?.cbf || 0);
      });

      const variance = varianceValue(wm, cbf);
      const diff = cbf - wm;

      output[row.item_code] = {
        wm,
        cbf,
        diff,
        variance,
        summary: batchSummaryText(row.description, wm, cbf),
      };
    });

    return output;
  }, [wideRows, productionCodes]);

  const visibleWideRows = useMemo(() => {
    if (!bomRows.length) return wideRows;

    return wideRows.filter((row) => {
      const hasAnyValue = productionCodes.some((code) => {
        const v = row.values[code] || { wm: 0, cbf: 0 };
        return Number(v.wm || 0) !== 0 || Number(v.cbf || 0) !== 0;
      });

      return hasAnyValue;
    });
  }, [wideRows, bomRows, productionCodes]);

  const visibleProductionCodes = batchVisible ? productionCodes : [];

  return (
    <div className="app">
      <aside className="sidebar">
        <h2>BOM Checker</h2>

        <button
          onClick={() => setPage("checker")}
          className={page === "checker" ? "active" : ""}
        >
          BOM Checker
        </button>

        <button
          onClick={() => setPage("items")}
          className={page === "items" ? "active" : ""}
        >
          Item List
        </button>
      </aside>

      <main className="main">
        {page === "checker" && (
          <>
            <header className="topbar">
              <div>
                <h1>BOM Checker</h1>
                <p>Compare WM usage against CBF actual usage by batch.</p>
              </div>

              <div className="top-actions">
                <button className="ghost">Month: {form.month}</button>
                <button onClick={() => setModal("WM")}>WM BOM</button>
                <button onClick={() => setModal("CBF")}>CBF BOM</button>
                <button onClick={downloadExcel} className="download">
                  Download Excel
                </button>
              </div>
            </header>

            {message && <p className="notice">{message}</p>}

            <div className="tableCard">
              <div className="tableWrap">
                <table>
                  <thead>
                    <tr>
                      <th rowSpan="2" className="sticky-col sticky-desc">
                        Item Description
                      </th>

                      <th rowSpan="2" className="sticky-col sticky-code">
                        CBF Item #
                      </th>

                      {batchVisible && (
                        <th colSpan={productionCodes.length * 3 + 5}>
                          {form.batch_name}
                        </th>
                      )}
                    </tr>

                    <tr>
                      {visibleProductionCodes.map((code) => (
                        <React.Fragment key={code}>
                          <th>Used ({code})</th>
                          <th>CBF-{code}</th>
                          <th>Variance</th>
                        </React.Fragment>
                      ))}

                      {batchVisible && (
                        <>
                          <th>Total WM</th>
                          <th>Total CBF</th>
                          <th>Total Variance</th>
                          <th>Total LBS</th>
                          <th>Summary</th>
                        </>
                      )}
                    </tr>
                  </thead>

                  <tbody>
                    {visibleWideRows.map((row) => {
                      const batchTotal = batchTotalsByRow[row.item_code] || {
                        wm: 0,
                        cbf: 0,
                        diff: 0,
                        variance: 0,
                        summary: "",
                      };

                      return (
                        <tr key={row.item_code}>
                          <td className="sticky-col sticky-desc">
                            {row.description}
                          </td>

                          <td className="sticky-col sticky-code">
                            {row.item_code}
                          </td>

                          {visibleProductionCodes.map((code) => {
                            const v = row.values[code] || { wm: 0, cbf: 0 };
                            const diff = lbsDiff(v.wm, v.cbf);

                            return (
                              <React.Fragment key={code}>
                                <td>{v.wm.toFixed(2)}</td>
                                <td>{v.cbf.toFixed(2)}</td>
                                <td
                                  className={
                                    diff < 0 ? "bad" : diff > 0 ? "good" : ""
                                  }
                                >
                                  {variancePct(v.wm, v.cbf)}
                                </td>
                              </React.Fragment>
                            );
                          })}

                          {batchVisible && (
                            <>
                              <td>{batchTotal.wm.toFixed(2)}</td>
                              <td>{batchTotal.cbf.toFixed(2)}</td>
                              <td>{batchTotal.variance.toFixed(2)}%</td>
                              <td>{batchTotal.diff.toFixed(2)} lbs</td>
                              <td
                                className={`summary ${summaryClass(
                                  row.item_code,
                                  batchTotal.variance
                                )}`}
                              >
                                {batchTotal.summary}
                              </td>
                            </>
                          )}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          </>
        )}

        {page === "items" && (
          <>
            <header className="topbar">
              <div>
                <h1>Item List</h1>
                <p>This list is hardcoded in the app.</p>
              </div>
            </header>

            <div className="tableCard">
              <table>
                <thead>
                  <tr>
                    <th>Description</th>
                    <th>CBF Item #</th>
                  </tr>
                </thead>

                <tbody>
                  {items.map((item) => (
                    <tr key={item.item_code}>
                      <td>{item.description}</td>
                      <td>{item.item_code}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </>
        )}
      </main>

      {modal && (
        <div className="modal">
          <div className="box">
            <h2>Upload {modal} BOM</h2>

            <label>
              Month
              <select
                value={form.month}
                onChange={(e) => setForm({ ...form, month: e.target.value })}
              >
                {MONTHS.map((m) => (
                  <option key={m}>{m}</option>
                ))}
              </select>
            </label>

            <label>
              Batch
              <select
                value={form.batch_name}
                onChange={(e) =>
                  setForm({ ...form, batch_name: e.target.value })
                }
              >
                {BATCHES.map((b) => (
                  <option key={b}>{b}</option>
                ))}
              </select>
            </label>

            <input
              type="file"
              multiple={modal === "CBF"}
              accept={modal === "WM" ? ".xlsx,.xls" : ".pdf"}
              onChange={(e) => setFiles(Array.from(e.target.files || []))}
            />

            <div className="actions">
              <button onClick={() => setModal(null)} className="ghost">
                Cancel
              </button>
              <button onClick={uploadBom}>Upload</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}