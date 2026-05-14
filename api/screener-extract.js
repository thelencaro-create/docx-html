import mammoth from "mammoth";
import JSZip from "jszip";
import ExcelJS from "exceljs";
import pdfParse from "pdf-parse";

/**
 * screener-extract — DOCX/XLSX/PDF → HTML
 *
 * Vereinheitlichter Eintrittspunkt für die Preview-Generator-Pipeline.
 * Liefert für alle drei Formate einheitliches HTML mit [[SHADED]]-Markern
 * an hinterlegten Tabellen-Zellen.
 *
 * Konsistenzen mit docx-html.js (v2-shading):
 *  - selber Marker [[SHADED]]
 *  - selbe Occurrence-basierte Logik für DOCX
 *  - selbe Fail-Soft-Strategie bei Shading-Extraktion
 *
 * NEU gegenüber docx-html.js:
 *  - XLSX-Support via ExcelJS (Shading aus cell.fill.fgColor)
 *  - PDF-Support via pdf-parse (Plaintext → <p>-Tags, ohne Shading)
 *  - Format-Auto-Detection via mimeType oder fileName
 *
 * Endpoint: POST /api/screener-extract
 * Request:  { data: <base64>, fileName?: string, mimeType?: string }
 * Response: { html, _format, _shadedCount, _shadedSampleTexts, _version }
 *
 * Marker-Form [[SHADED]] passt zum bestehenden Hardening / Parser-Prompt.
 */

// =============================================================================
// HELPERS (kopiert/angepasst aus docx-html.js für Konsistenz)
// =============================================================================
function toBufferFromBase64(input) {
  const s = String(input ?? "").trim();
  const clean = s.includes("base64,") ? s.split("base64,").pop() : s;
  return Buffer.from(clean, "base64");
}
function isZip(buf) {
  return buf.length >= 4 && buf[0] === 0x50 && buf[1] === 0x4B;
}
function hasEOCD(buf) {
  const sig = Buffer.from([0x50, 0x4B, 0x05, 0x06]);
  const start = Math.max(0, buf.length - 70000);
  return buf.slice(start).indexOf(sig) !== -1;
}
function isPdf(buf) {
  return buf.length >= 4 && buf[0] === 0x25 && buf[1] === 0x50 && buf[2] === 0x44 && buf[3] === 0x46;
}
function normalizeText(s) {
  return String(s).replace(/\s+/g, " ").trim().toLowerCase();
}
function escapeHtml(s) {
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function detectFormat({ fileName, mimeType, buf }) {
  const fn = String(fileName || "").toLowerCase();
  const mt = String(mimeType || "").toLowerCase();
  if (fn.endsWith(".docx") || mt.includes("wordprocessingml")) return "docx";
  if (fn.endsWith(".xlsx") || fn.endsWith(".xls") || mt.includes("spreadsheetml") || mt.includes("excel")) return "xlsx";
  if (fn.endsWith(".pdf") || mt.includes("pdf")) return "pdf";
  // Magic-Bytes-Fallback
  if (buf && isPdf(buf)) return "pdf";
  return null;
}

// =============================================================================
// DOCX — identisch zu docx-html.js v2-shading
// =============================================================================
function isShadedFillDocx(fill) {
  if (!fill) return false;
  const f = String(fill).trim().toLowerCase();
  if (f === "auto" || f === "ffffff" || f === "fff" || f === "none") return false;
  if (!/^[0-9a-f]{3}$|^[0-9a-f]{6}$/.test(f)) return false;
  return true;
}

function extractDocxShadedOccurrences(docXml) {
  const tcRegex = /<w:tc(?:\s[^>]*)?>([\s\S]*?)<\/w:tc>/g;
  const tcBlocks = [];
  let m;
  while ((m = tcRegex.exec(docXml)) !== null) tcBlocks.push(m[1]);

  const cells = tcBlocks.map(inner => {
    const tcPrMatch = inner.match(/<w:tcPr>([\s\S]*?)<\/w:tcPr>/);
    let shaded = false;
    if (tcPrMatch) {
      const shdMatch = tcPrMatch[1].match(/<w:shd\b[^/>]*w:fill="([^"]+)"/);
      if (shdMatch) shaded = isShadedFillDocx(shdMatch[1]);
    }
    const textParts = [];
    const tRegex = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
    let tm;
    while ((tm = tRegex.exec(inner)) !== null) {
      const t = tm[1]
        .replace(/&amp;/g, "&")
        .replace(/&lt;/g, "<")
        .replace(/&gt;/g, ">")
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'");
      textParts.push(t);
    }
    return { shaded, text: textParts.join("").trim() };
  });

  const counter = new Map();
  const shadedOccurrences = new Map();
  for (const c of cells) {
    if (!c.text) continue;
    const key = normalizeText(c.text);
    const idx = counter.get(key) || 0;
    if (c.shaded) {
      if (!shadedOccurrences.has(key)) shadedOccurrences.set(key, []);
      shadedOccurrences.get(key).push(idx);
    }
    counter.set(key, idx + 1);
  }
  return shadedOccurrences;
}

function injectShadedMarkersInHtml(html, shadedOccurrences) {
  if (!shadedOccurrences || shadedOccurrences.size === 0) return html;
  const counter = new Map();
  return html.replace(/<td\b([^>]*)>([\s\S]*?)<\/td>/gi, (match, attrs, inner) => {
    const plain = inner.replace(/<[^>]+>/g, " ").replace(/&nbsp;/g, " ");
    const key = normalizeText(plain);
    if (!key) return match;
    const idx = counter.get(key) || 0;
    counter.set(key, idx + 1);
    const occs = shadedOccurrences.get(key);
    if (!occs || !occs.includes(idx)) return match;

    let newInner;
    const firstTextMatch = inner.match(/>([^<]*\S[^<]*)</);
    if (firstTextMatch) {
      newInner = inner.replace(firstTextMatch[0], `>[[SHADED]]${firstTextMatch[1]}<`);
    } else {
      newInner = `[[SHADED]]${inner}`;
    }
    return `<td${attrs}>${newInner}</td>`;
  });
}

async function extractDocx(buf) {
  if (!isZip(buf) || !hasEOCD(buf) || buf.length < 1024) {
    throw new Error("Invalid DOCX/ZIP payload");
  }
  const result = await mammoth.convertToHtml({ buffer: buf });
  let html = String(result.value || "").trim();

  let shadedCount = 0;
  let shadedSampleTexts = [];
  try {
    const zip = await JSZip.loadAsync(buf);
    const docXmlFile = zip.file("word/document.xml");
    if (docXmlFile) {
      const docXml = await docXmlFile.async("string");
      const shadedOccurrences = extractDocxShadedOccurrences(docXml);
      shadedCount = [...shadedOccurrences.values()].reduce((s, arr) => s + arr.length, 0);
      shadedSampleTexts = [...shadedOccurrences.keys()].slice(0, 10);
      if (shadedCount > 0) html = injectShadedMarkersInHtml(html, shadedOccurrences);
    }
  } catch (e) {
    console.warn("DOCX shading extraction failed (non-fatal):", e.message);
  }

  return { html, shadedCount, shadedSampleTexts };
}

// =============================================================================
// XLSX — via ExcelJS, [[SHADED]]-Marker an Zellen mit Hintergrund-Farbe
// =============================================================================
function isShadedFillXlsx(fgColor) {
  if (!fgColor) return false;
  const argb = String(fgColor.argb || "").toUpperCase();
  if (!argb) return false;
  if (argb === "FFFFFFFF" || argb === "00000000" || argb === "FFFFFF") return false;
  return true;
}

function parseRef(ref) {
  const m = String(ref || "").match(/^([A-Z]+)(\d+)$/);
  if (!m) return null;
  const letters = m[1];
  const row = parseInt(m[2], 10);
  let col = 0;
  for (let i = 0; i < letters.length; i++) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  return { row, col };
}

function getCellText(cell) {
  const v = cell.value;
  if (v == null) return "";
  if (typeof v === "object") {
    if (v.richText) return v.richText.map(t => t.text).join("");
    if (v.text != null) return String(v.text);
    if (v.result != null) return String(v.result);
    if (v.formula) return "";
    return String(v);
  }
  return String(v);
}

async function extractXlsx(buf) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);

  const parts = [];
  let totalRows = 0, totalCells = 0, shadedCount = 0;
  const shadedSampleSet = new Set();

  wb.worksheets.forEach((ws) => {
    parts.push(`<h2>${escapeHtml(ws.name)}</h2>`);
    parts.push('<table border="1" style="border-collapse:collapse">');

    const mergeSkip = new Set();
    const mergeSpans = new Map();
    if (ws.model && ws.model.merges) {
      for (const range of ws.model.merges) {
        try {
          const [topLeft, bottomRight] = range.split(":");
          const tl = parseRef(topLeft);
          const br = parseRef(bottomRight);
          if (!tl || !br) continue;
          mergeSpans.set(`${tl.row},${tl.col}`, {
            rowspan: br.row - tl.row + 1,
            colspan: br.col - tl.col + 1,
          });
          for (let r = tl.row; r <= br.row; r++) {
            for (let c = tl.col; c <= br.col; c++) {
              if (r === tl.row && c === tl.col) continue;
              mergeSkip.add(`${r},${c}`);
            }
          }
        } catch (e) { /* ignore */ }
      }
    }

    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      const cells = [];
      let hasContent = false;
      const maxCol = row.cellCount;
      for (let c = 1; c <= maxCol; c++) {
        const key = `${rowNumber},${c}`;
        if (mergeSkip.has(key)) continue;

        const cell = row.getCell(c);
        const text = getCellText(cell).trim();
        if (text) hasContent = true;
        totalCells++;

        let shaded = false;
        const fill = cell.fill;
        if (fill && fill.type === "pattern" && fill.pattern === "solid") {
          if (isShadedFillXlsx(fill.fgColor)) shaded = true;
        }
        if (shaded) {
          shadedCount++;
          if (text && shadedSampleSet.size < 10) shadedSampleSet.add(text.substring(0, 40));
        }

        const bold = !!(cell.font && cell.font.bold);

        const span = mergeSpans.get(key);
        const spanAttr = span
          ? `${span.rowspan > 1 ? ` rowspan="${span.rowspan}"` : ""}${span.colspan > 1 ? ` colspan="${span.colspan}"` : ""}`
          : "";

        let inner = escapeHtml(text);
        if (bold && text) inner = `<strong>${escapeHtml(text)}</strong>`;
        if (shaded) inner = `[[SHADED]]${inner}`;

        cells.push(`<td${spanAttr}>${inner}</td>`);
      }
      if (hasContent) {
        parts.push(`<tr>${cells.join("")}</tr>`);
        totalRows++;
      }
    });

    parts.push("</table>");
  });

  return {
    html: parts.join("\n"),
    shadedCount,
    shadedSampleTexts: [...shadedSampleSet],
    stats: { sheets: wb.worksheets.length, rows: totalRows, cells: totalCells },
  };
}

// =============================================================================
// PDF — via pdf-parse, Plaintext → <p>-Tags (kein Shading)
// =============================================================================
async function extractPdf(buf) {
  if (!isPdf(buf)) {
    throw new Error("Invalid PDF payload (missing %PDF header)");
  }
  const data = await pdfParse(buf);
  const text = data.text || "";
  const paragraphs = text
    .split(/\n{2,}/)
    .map(p => p.trim())
    .filter(Boolean)
    .map(p => `<p>${escapeHtml(p).replace(/\n/g, "<br/>")}</p>`);
  return {
    html: paragraphs.join("\n"),
    shadedCount: 0,
    shadedSampleTexts: [],
    stats: { pages: data.numpages, chars: text.length },
  };
}

// =============================================================================
// HTTP HANDLER
// =============================================================================
export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(204).end();

  try {
    if (req.method !== "POST") {
      return res.status(405).json({ error: "Method not allowed" });
    }
    let body = req.body;
    if (typeof body === "string") {
      try { body = JSON.parse(body); }
      catch (e) { return res.status(400).json({ error: "Body is not valid JSON" }); }
    }
    if (!body || typeof body !== "object") {
      return res.status(400).json({ error: "Missing JSON body" });
    }
    const { fileName, mimeType, data } = body;
    if (!data || typeof data !== "string") {
      return res.status(400).json({ error: "Missing 'data' (base64 string)" });
    }

    const buf = toBufferFromBase64(data);
    if (buf.length === 0) {
      return res.status(400).json({ error: "Empty buffer after base64 decode" });
    }

    const format = detectFormat({ fileName, mimeType, buf });
    if (!format) {
      return res.status(400).json({
        error: "Unknown format",
        details: { fileName, mimeType, fileSize: buf.length },
        hint: "Supported: .docx, .xlsx, .pdf",
      });
    }

    let result;
    if (format === "docx") result = await extractDocx(buf);
    else if (format === "xlsx") result = await extractXlsx(buf);
    else if (format === "pdf") result = await extractPdf(buf);

    return res.status(200).json({
      html: result.html,
      _format: format,
      _shadedCount: result.shadedCount,
      _shadedSampleTexts: result.shadedSampleTexts,
      _stats: result.stats || null,
      _version: "screener-extract-v1.0.0",
    });
  } catch (err) {
    console.error("screener-extract error:", err);
    return res.status(500).json({ error: String(err?.message || err) });
  }
}

export const config = {
  api: { bodyParser: { sizeLimit: "10mb" } },
};
