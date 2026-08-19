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
  ['Wonder Monday 3" Lids', "PP2723"],
  ['Black 3" Bases', "PP2721"],
  ['Wonder Monday 3" Clear Base', "PP2722"],
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

function getNextBatchName(existingBatches = []) {
  const numbers = existingBatches
    .map((batch) => Number(String(batch).match(/\d+/)?.[0] || 0))
    .filter((value) => Number.isFinite(value));

  return `Batch ${Math.max(0, ...numbers) + 1}`;
}

function norm(v) {
  return String(v ?? "").trim().toLowerCase();
}

function compact(v) {
  return String(v ?? "").replace(/\s+/g, "").toUpperCase();
}

function num(v) {
  if (v === null || v === undefined || v === "") return 0;
  const n = Number(String(v).replace(/,/g, "").replace("%", ""));
  return Number.isFinite(n) ? n : 0;
}

function positiveNum(v) {
  return Math.abs(num(v));
}

function cleanProductionCode(value) {
  const raw = String(value || "").trim();
  const match = raw.match(/\d+/);
  return match ? match[0] : raw || "PDF";
}

function getProductionCodeFromFileName(fileName) {
  const namedMatch = String(fileName || "").match(/\b(?:WM|PO)[\s\-_]*(\d+)\b/i);
  if (namedMatch) return namedMatch[1];

  const fallbackMatch = String(fileName || "").match(/(\d+)/);
  return fallbackMatch ? fallbackMatch[1] : "PDF";
}

function wmColumnLabel(code) {
  return `PO-${cleanProductionCode(code)}`;
}

function cbfColumnLabel(code) {
  return `WM${cleanProductionCode(code)}`;
}

function uploadLabel(bomType, codes) {
  const cleanCodes = [...new Set(codes.map(cleanProductionCode).filter(Boolean))];
  const codeLabel = cleanCodes.length ? cleanCodes.join("-") : "PDF";
  return bomType === "WM" ? `WM-PO-${codeLabel}` : `CBF-WM${codeLabel}`;
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
    )}% surplus. CBF used ${diff.toFixed(2)} lbs more than WM expected.`;
  }

  return `${name}: There is a ${pct.toFixed(
    2
  )}% variance. CBF used ${diff.toFixed(2)} lbs less than expected.`;
}

function getWmHeaderProductionCode(header) {
  const value = String(header || "").trim();
  const usedMatch = value.match(/^Used\s*\(([^)]+)\)/i);
  if (usedMatch) return cleanProductionCode(usedMatch[1]);

  const poMatch = value.match(/^(?:WM[\s\-_]*)?PO[\s\-_]*(\d+)$/i);
  if (poMatch) return poMatch[1];

  return null;
}

function buildDescriptionLookup(items) {
  const lookup = new Map();

  items.forEach((item) => {
    const key = norm(item.description);
    if (!lookup.has(key)) lookup.set(key, item);
  });

  return lookup;
}

function findItemByRow(row, itemByCode, itemByDescription, firstUsageColumn) {
  const codeSearchLimit = firstUsageColumn > 0 ? firstUsageColumn : row.length;

  for (let index = 0; index < codeSearchLimit; index += 1) {
    const code = compact(row[index]);
    if (itemByCode.has(code)) return itemByCode.get(code);
  }

  return itemByDescription.get(norm(row[0]));
}

function extractFirstEachQty(text) {
  const clean = String(text || "").replace(/\s+/g, " ");
  const beforeCost = clean.split("$")[0];
  const matches = [
    ...beforeCost.matchAll(/([0-9][0-9,]*\.?[0-9]*)\s*each\b/gi),
  ];

  if (!matches.length) return null;
  return num(matches[0][1]);
}

function extractStandaloneQty(text) {
  const clean = String(text || "").replace(/\s+/g, " ").trim();
  if (!/^[0-9][0-9,]*\.?[0-9]*$/.test(clean)) return null;
  return num(clean);
}

function extractFirstLbQty(text) {
  const clean = String(text || "").replace(/\s+/g, " ");
  const beforeCost = clean.split("$")[0];

  const matches = [
    ...beforeCost.matchAll(/([0-9][0-9,]*\.?[0-9]*)\s*lb\.?\b/gi),
  ];

  if (!matches.length) return null;
  return num(matches[0][1]);
}

function isPackagingCode(code) {
  return /^(P|PL|PP|PC|PS)/i.test(String(code || ""));
}

function extractPackagingQty(lines, index) {
  const line = lines[index] || "";
  const direct = extractFirstEachQty(line);
  if (direct !== null) return direct;

  for (const offset of [1, -1, 2]) {
    const nearby = extractStandaloneQty(lines[index + offset]);
    if (nearby !== null) return nearby;
  }

  return extractFirstEachQty(lines.slice(index, index + 3).join(" "));
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
      .sort((a, b) => {
        const [pageA, yA] = a[0].split("-").map(Number);
        const [pageB, yB] = b[0].split("-").map(Number);

        if (pageA !== pageB) return pageA - pageB;
        return yB - yA;
      })
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
    row.some((cell) => getWmHeaderProductionCode(cell) !== null)
  );

  if (headerRowIndex < 0) {
    throw new Error("No WM PO columns found. Expected Used(...), PO-1000, or WM-PO-1000.");
  }

  const headers = rows[headerRowIndex].map((h) => String(h || "").trim());
  const itemByCode = new Map(items.map((item) => [compact(item.item_code), item]));
  const itemByDescription = buildDescriptionLookup(items);

  const groups = headers
    .map((header, index) => {
      const prod = getWmHeaderProductionCode(header);
      if (!prod) return null;
      return { prod, usedCol: index };
    })
    .filter(Boolean);

  const firstUsageColumn = Math.min(...groups.map((group) => group.usedCol));
  const parsed = [];

  for (let r = headerRowIndex + 1; r < rows.length; r++) {
    const description = String(rows[r][0] || "").trim();
    if (!description || norm(description) === "cbf") continue;

    const matched = findItemByRow(
      rows[r],
      itemByCode,
      itemByDescription,
      firstUsageColumn
    );

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
  const codeCompact = compact(code);
  const packaging = isPackagingCode(code);

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i] || "";

    if (!compact(line).includes(codeCompact)) continue;

    let usage = null;

    if (packaging) {
      usage = extractPackagingQty(lines, i);
    } else {
      for (let j = i - 1; j >= Math.max(0, i - 4); j--) {
        const previousLine = lines[j] || "";
        const lower = previousLine.toLowerCase();

        if (
          lower.includes("historical") ||
          lower.includes("cost") ||
          lower.includes("ingredient items") ||
          lower.includes("packaging")
        ) {
          continue;
        }

        const lbValue = extractFirstLbQty(previousLine);

        if (lbValue !== null) {
          usage = lbValue;
          break;
        }
      }

      if (usage === null) {
        usage = extractFirstLbQty(line);
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

  console.log("PDF:", file.name, "PO:", productionCode);
  console.log("PDF LINES:", lines);

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

  console.log("PARSED PDF ROWS:", parsed);

  if (parsed.length === 0) {
    throw new Error(
      `No matching CBF item numbers found in ${file.name}. Check console PDF LINES.`
    );
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
  const [visibleBatches, setVisibleBatches] = useState([]);
  const [showAllItems, setShowAllItems] = useState(false);

  const [form, setForm] = useState({
    month: MONTHS[0],
    batch_mode: "new",
    batch_name: "Batch 1",
  });

  useEffect(() => {
    loadData();
  }, [form.month]);

  async function loadData() {
    const { data, error } = await supabase
      .from("bom_usage_rows")
      .select("*")
      .eq("month", form.month);

    if (error) {
      console.error(error);
      setMessage(`Supabase load failed: ${error.message}`);
      return;
    }

    setBomRows(data || []);
  }

  async function saveUsageRows(rows) {
    for (const row of rows) {
      const match = {
        month: row.month,
        batch: row.batch,
        item_code: row.item_code,
        production_code: row.production_code,
        bom_type: row.bom_type,
      };

      const { data: existingRows, error: lookupError } = await supabase
        .from("bom_usage_rows")
        .select("id")
        .match(match)
        .limit(1);

      if (lookupError) throw lookupError;

      if (existingRows?.length) {
        const { error } = await supabase
          .from("bom_usage_rows")
          .update(row)
          .eq("id", existingRows[0].id);

        if (error) throw error;
      } else {
        const { error } = await supabase.from("bom_usage_rows").insert(row);
        if (error) throw error;
      }
    }
  }

  async function uploadBom() {
    if (!files.length) return alert("Choose a file first.");

    try {
      const bomType = modal;
      const targetBatch =
        String(form.batch_name || "").trim() || getNextBatchName(availableBatches);

      if (bomType === "WM") {
        await supabase
          .from("bom_usage_rows")
          .delete()
          .eq("month", form.month)
          .eq("batch", targetBatch)
          .eq("bom_type", "WM");
      }

      let allRows = [];
      const uploadNames = [];

      for (const file of files) {
        const parsed =
          bomType === "WM"
            ? await parseWmExcel(file, items)
            : await parseCbfPdf(file, items);

        const currentUploadLabel = uploadLabel(
          bomType,
          parsed.map((row) => row.production_code)
        );
        uploadNames.push(currentUploadLabel);
        const cleanMonth = form.month.replaceAll(" ", "-").replaceAll("'", "");
        const cleanBatch = targetBatch.replaceAll(" ", "");
        const cleanUpload = currentUploadLabel.replace(/[^a-z0-9-]+/gi, "-");
        const storagePath = `${bomType}/${cleanMonth}/${cleanBatch}/${cleanUpload}/${Date.now()}-${file.name}`;

        const { error: storageError } = await supabase.storage
          .from("bom-files")
          .upload(storagePath, file, { upsert: true });

        if (storageError) throw storageError;

        const { data: uploadRow, error: uploadError } = await supabase
          .from("bom_uploads")
          .insert({
            bom_type: bomType,
            month: form.month,
            batch: targetBatch,
            file_name: `${currentUploadLabel} - ${file.name}`,
            storage_path: storagePath,
          })
          .select()
          .single();

        if (uploadError) throw uploadError;

        const rows = parsed.map((r) => ({
          upload_id: uploadRow.id,
          bom_type: bomType,
          month: form.month,
          batch: targetBatch,
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

      await saveUsageRows(allRows);
      await loadData();

      setMessage(`${[...new Set(uploadNames)].join(", ")} uploaded: ${allRows.length} values parsed and saved.`);
      setModal(null);
      setFiles([]);
      setVisibleBatches((current) =>
        current.includes(targetBatch)
          ? current
          : [...current, targetBatch]
      );
    } catch (err) {
      console.error(err);
      alert(err.message || "Upload failed.");
    }
  }

  async function saveEditedUsage(row, group, bomType, rawValue) {
    const value = positiveNum(rawValue);
    const payload = {
      upload_id: null,
      bom_type: bomType,
      month: form.month,
      batch: group.batch,
      item_code: row.item_code,
      description: row.description,
      production_code: group.production_code,
      usage_lbs: value,
      wm_usage: bomType === "WM" ? value : null,
      cbf_usage: bomType === "CBF" ? value : null,
      row_order: row.row_order,
    };

    const match = {
      month: form.month,
      batch: group.batch,
      item_code: row.item_code,
      production_code: group.production_code,
      bom_type: bomType,
    };

    const { error: deleteError } = await supabase
      .from("bom_usage_rows")
      .delete()
      .match(match);

    if (deleteError) throw deleteError;

    const { data, error } = await supabase
      .from("bom_usage_rows")
      .insert(payload)
      .select()
      .single();

    if (error) throw error;

    setBomRows((current) => [
      ...current.filter(
        (candidate) =>
          !(
            candidate.month === form.month &&
            candidate.batch === group.batch &&
            candidate.item_code === row.item_code &&
            candidate.production_code === group.production_code &&
            candidate.bom_type === bomType
          )
      ),
      data || payload,
    ]);

    setMessage(
      `${bomType} ${row.item_code} ${group.batch} ${group.production_code} saved to Supabase.`
    );
  }

  async function editUsageValue(row, group, bomType, currentValue) {
    const label = bomType === "WM" ? wmColumnLabel(group.production_code) : cbfColumnLabel(group.production_code);
    const next = window.prompt(
      `Edit ${label} for ${row.description} in ${group.batch}`,
      Number(currentValue || 0).toFixed(2)
    );

    if (next === null) return;
    const value = positiveNum(next);

    try {
      await saveEditedUsage(row, group, bomType, value);
    } catch (error) {
      console.error(error);
      alert(error.message || "Could not save edited value.");
    }
  }

  function downloadExcel() {
    const rows = [];

    visibleWideRows.forEach((row) => {
      const record = {
        "Item Description": row.description,
        "CBF Item #": row.item_code,
      };

      visibleGroups.forEach((group) => {
        const v = row.values[group.key] || { wm: 0, cbf: 0 };
        record[`${group.batch} ${wmColumnLabel(group.production_code)}`] = Number(v.wm.toFixed(2));
        record[`${group.batch} ${cbfColumnLabel(group.production_code)}`] = Number(v.cbf.toFixed(2));
        record[`${group.batch} Variance ${group.production_code}`] = variancePct(v.wm, v.cbf);
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

    XLSX.utils.book_append_sheet(wb, ws, form.month);

    XLSX.writeFile(
      wb,
      `BOM_Checker_${form.month.replaceAll(" ", "_").replaceAll("'", "")}.xlsx`
    );
  }

  const availableBatches = useMemo(() => {
    return [...new Set(bomRows.map((r) => r.batch).filter(Boolean))].sort(
      (a, b) => Number(String(a).match(/\d+/)?.[0] || 0) - Number(String(b).match(/\d+/)?.[0] || 0)
    );
  }, [bomRows]);

  useEffect(() => {
    setVisibleBatches((current) => {
      if (!availableBatches.length) return [];
      const kept = current.filter((batch) => availableBatches.includes(batch));
      const added = availableBatches.filter((batch) => !current.includes(batch));
      return kept.length ? [...kept, ...added] : availableBatches;
    });
  }, [availableBatches.join("|")]);

  const filteredBomRows = useMemo(() => {
    if (!availableBatches.length) return bomRows;
    return bomRows.filter((r) => visibleBatches.includes(r.batch));
  }, [bomRows, visibleBatches, availableBatches]);

  const visibleGroups = useMemo(() => {
    const groups = new Map();

    filteredBomRows.forEach((row) => {
      if (!row.production_code) return;
      const key = `${row.batch}::${row.production_code}`;
      if (!groups.has(key)) {
        groups.set(key, {
          key,
          batch: row.batch,
          production_code: row.production_code,
        });
      }
    });

    return [...groups.values()].sort((a, b) => {
      const batchSort = Number(String(a.batch).match(/\d+/)?.[0] || 0) - Number(String(b.batch).match(/\d+/)?.[0] || 0);
      if (batchSort !== 0) return batchSort;
      return Number(a.production_code) - Number(b.production_code);
    });
  }, [filteredBomRows]);

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

    filteredBomRows.forEach((r) => {
      const row = grouped.get(r.item_code);
      if (!row) return;

      const key = `${r.batch}::${r.production_code}`;
      const current = row.values[key] || { wm: 0, cbf: 0 };

      if (r.bom_type === "WM") {
        current.wm += Number(r.wm_usage ?? r.usage_lbs ?? 0);
      }

      if (r.bom_type === "CBF") {
        current.cbf += Number(r.cbf_usage ?? r.usage_lbs ?? 0);
      }

      row.values[key] = current;
    });

    return [...grouped.values()].sort((a, b) => a.row_order - b.row_order);
  }, [items, filteredBomRows]);

  const batchTotalsByRow = useMemo(() => {
    const output = {};

    wideRows.forEach((row) => {
      let wm = 0;
      let cbf = 0;

      visibleGroups.forEach((group) => {
        wm += Number(row.values[group.key]?.wm || 0);
        cbf += Number(row.values[group.key]?.cbf || 0);
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
  }, [wideRows, visibleGroups]);

  const visibleWideRows = useMemo(() => {
    if (showAllItems) return wideRows;
    if (!bomRows.length) return wideRows;

    return wideRows.filter((row) => {
      const hasAnyValue = visibleGroups.some((group) => {
        const v = row.values[group.key] || { wm: 0, cbf: 0 };
        return Number(v.wm || 0) !== 0 || Number(v.cbf || 0) !== 0;
      });

      return hasAnyValue;
    });
  }, [wideRows, bomRows, visibleGroups, showAllItems]);

  const summaryRows = useMemo(() => {
    return visibleWideRows.map((row) => {
      const total = batchTotalsByRow[row.item_code] || {
        wm: 0,
        cbf: 0,
        diff: 0,
        variance: 0,
        summary: "",
      };

      return { ...row, total };
    });
  }, [visibleWideRows, batchTotalsByRow]);

  function toggleBatch(batch) {
    setVisibleBatches((current) =>
      current.includes(batch)
        ? current.filter((item) => item !== batch)
        : [...current, batch]
    );
  }

  function openUploadModal(type) {
    const mode = availableBatches.length ? "existing" : "new";
    const batchName = availableBatches.length
      ? availableBatches[0]
      : getNextBatchName(availableBatches);

    setFiles([]);
    setForm((current) => ({
      ...current,
      batch_mode: mode,
      batch_name: batchName,
    }));
    setModal(type);
  }

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
                <label className="monthFilter">
                  <span>Month</span>
                  <select
                    value={form.month}
                    onChange={(e) => setForm({ ...form, month: e.target.value })}
                  >
                    {MONTHS.map((m) => (
                      <option key={m}>{m}</option>
                    ))}
                  </select>
                </label>
                <button
                  className="ghost"
                  onClick={() => setShowAllItems((value) => !value)}
                >
                  {showAllItems ? "Hide Empty Items" : "Show All Items"}
                </button>
                <button onClick={() => openUploadModal("WM")}>WM BOM</button>
                <button onClick={() => openUploadModal("CBF")}>CBF BOM</button>
                <button onClick={downloadExcel} className="download">
                  Download Excel
                </button>
              </div>
            </header>

            {message && <p className="notice">{message}</p>}

            {!!availableBatches.length && (
              <div className="batchFilters">
                <span>Batch filter</span>
                {availableBatches.map((batch) => (
                  <button
                    key={batch}
                    className={visibleBatches.includes(batch) ? "active" : ""}
                    onClick={() => toggleBatch(batch)}
                  >
                    {batch}
                  </button>
                ))}
                <button
                  className={visibleBatches.length === availableBatches.length ? "active" : ""}
                  onClick={() => setVisibleBatches(availableBatches)}
                >
                  All
                </button>
              </div>
            )}

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

                      {visibleGroups.map((group) => (
                        <th key={group.key} colSpan="3">
                          {group.batch} | {group.production_code}
                        </th>
                      ))}

                      <th colSpan="5">Total</th>
                    </tr>

                    <tr>
                      {visibleGroups.map((group) => (
                        <React.Fragment key={group.key}>
                          <th>{wmColumnLabel(group.production_code)}</th>
                          <th>{cbfColumnLabel(group.production_code)}</th>
                          <th>Variance</th>
                        </React.Fragment>
                      ))}

                      <th>Total WM</th>
                      <th>Total CBF</th>
                      <th>Total Variance</th>
                      <th>Total LBS</th>
                      <th>Summary</th>
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

                          {visibleGroups.map((group) => {
                            const v = row.values[group.key] || { wm: 0, cbf: 0 };
                            const diff = lbsDiff(v.wm, v.cbf);

                            return (
                              <React.Fragment key={group.key}>
                                <td
                                  className="editableCell"
                                  title="Double-click to edit WM value"
                                  onDoubleClick={() =>
                                    editUsageValue(row, group, "WM", v.wm)
                                  }
                                >
                                  {v.wm.toFixed(2)}
                                </td>
                                <td
                                  className="editableCell"
                                  title="Double-click to edit CBF value"
                                  onDoubleClick={() =>
                                    editUsageValue(row, group, "CBF", v.cbf)
                                  }
                                >
                                  {v.cbf.toFixed(2)}
                                </td>
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
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>

            {!!summaryRows.length && (
              <section className="analysisPanel">
                <div className="analysisHeader">
                  <h2>Summary</h2>
                  <p>Visible batch totals for {form.month}</p>
                </div>

                <div className="analysisGrid">
                  {summaryRows.map((row) => (
                    <div
                      key={row.item_code}
                      className={`analysisLine ${summaryClass(
                        row.item_code,
                        row.total.variance
                      )}`}
                    >
                      <strong>{row.description}</strong>
                      <span>{row.total.summary}</span>
                    </div>
                  ))}
                </div>
              </section>
            )}
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

            <div className="modeGroup" role="group" aria-label="Batch mode">
              <button
                type="button"
                className={form.batch_mode === "existing" ? "active" : ""}
                disabled={!availableBatches.length}
                onClick={() =>
                  setForm({
                    ...form,
                    batch_mode: "existing",
                    batch_name: availableBatches[0] || getNextBatchName(availableBatches),
                  })
                }
              >
                Existing Batch
              </button>
              <button
                type="button"
                className={form.batch_mode === "new" ? "active" : ""}
                onClick={() =>
                  setForm({
                    ...form,
                    batch_mode: "new",
                    batch_name: getNextBatchName(availableBatches),
                  })
                }
              >
                New Batch
              </button>
            </div>

            <label>
              Batch
              {form.batch_mode === "existing" && availableBatches.length ? (
                <select
                  value={form.batch_name}
                  onChange={(e) =>
                    setForm({ ...form, batch_name: e.target.value })
                  }
                >
                  {availableBatches.map((b) => (
                    <option key={b}>{b}</option>
                  ))}
                </select>
              ) : (
                <input
                  value={form.batch_name}
                  placeholder="Batch 1"
                  onChange={(e) =>
                    setForm({ ...form, batch_name: e.target.value })
                  }
                />
              )}
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