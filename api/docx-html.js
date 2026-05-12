import mammoth from "mammoth";
import JSZip from "jszip";

/**
 * docx-html v2 — DOCX→HTML mit Cell-Shading-Erkennung
 *
 * v2 (neu): Erkennt farbig hinterlegte Tabellen-Zellen aus dem DOCX-XML
 * (<w:tcPr><w:shd w:fill="HEXFARBE"/></w:tcPr>) und ergänzt deren
 * Textinhalt mit dem Marker [[SHADED]] vor dem ersten Zeichen jedes Runs.
 *
 * Hintergrund Mars Q19: Bestimmte Zellen (Matrix-Items × Skala-Codes) sind
 * im Original-Screener farbig markiert — das sind die erlaubten Antworten.
 * Mammoth verliert diese Information beim HTML-Export. Wir lesen sie deshalb
 * direkt aus der document.xml und re-injecten den Marker in den HTML-Output.
 *
 * Funktioniert mit ALLEN Shading-Farben (Hellblau, Grün, Gelb, Rosa, ...).
 * Ausgeschlossen werden nur "kein Fill"-Werte: "auto", "FFFFFF", "ffffff", null.
 *
 * Marker-Wahl: [[SHADED]] (statt [[BLAU]]) — farbneutral, weil die
 * Recruiter-Anweisung im Screener immer auf "die markierten Felder"
 * verweist, unabhängig von der konkreten Farbe.
 */

// -------------------------------------------------------------
// Helpers (unverändert aus v1)
// -------------------------------------------------------------
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

// -------------------------------------------------------------
// SHADING-EXTRAKTOR (neu)
// -------------------------------------------------------------

/**
 * Liest aus document.xml alle Tabellen-Zellen mit Cell-Shading und liefert
 * eine Liste der Texte, die in solchen Zellen stehen.
 *
 * Zwei Quellen pro Zelle:
 *   1) Direkt-Shading:   <w:tcPr><w:shd w:fill="HEX"/></w:tcPr>
 *   2) Style-Shading via <w:cnfStyle/> ist möglich, aber selten und wird
 *      hier nicht abgedeckt — die meisten Screener nutzen Direkt-Shading.
 *
 * Rückgabe: Set<string> mit normalisierten Texten (lowercase, getrimmt,
 * Whitespace kollabiert), die in shaded cells stehen. Wir nutzen das später
 * als Lookup beim HTML-Post-Processing.
 *
 * Mit Marker-Position: Ein Text kann mehrfach im Dokument vorkommen, aber
 * nur in einer der Vorkommen shaded sein. Wir speichern deshalb zusätzlich
 * die Sequenz (Reihenfolge im Doc), damit der Marker-Inject die N-te
 * Occurrence treffen kann.
 */
function isShadedFill(fill) {
  if (!fill) return false;
  const f = String(fill).trim().toLowerCase();
  // "auto" oder reines Weiß = kein Shading
  if (f === "auto" || f === "ffffff" || f === "fff" || f === "none") return false;
  // Hex-Pattern check — nur 3- oder 6-stellig
  if (!/^[0-9a-f]{3}$|^[0-9a-f]{6}$/.test(f)) return false;
  return true;
}

/**
 * Extrahiert alle shaded Text-Vorkommen aus document.xml.
 * Liefert ein Array von { text, occurrenceIdx } pro shaded Zelle.
 * Die occurrenceIdx ist die N-te Occurrence dieses Texts (0-basiert) bezogen
 * auf alle Tabellen-Zellen-Texte im Dokument (auch nicht-shaded).
 *
 * Diese Buchführung ist nötig, weil derselbe Text (z.B. "5") in vielen
 * Zellen einer Matrix auftaucht, aber nur in manchen shaded ist.
 */
function extractShadedCellTexts(docXml) {
  // Wir parsen iterativ mit Regex — voller XML-Parser wäre overkill und
  // bringt Dependencies. Die Struktur ist:
  //   <w:tbl>
  //     <w:tr>
  //       <w:tc>
  //         <w:tcPr>
  //           <w:shd w:fill="..."/>  (optional)
  //         </w:tcPr>
  //         <w:p>...Text...</w:p>
  //       </w:tc>
  //     </w:tr>
  //   </w:tbl>
  //
  // Strategie: alle <w:tc>...</w:tc> Blocks holen (in Reihenfolge), pro Block
  // (a) prüfen ob shaded, (b) gesamten Text extrahieren.

  const tcRegex = /<w:tc(?:\s[^>]*)?>([\s\S]*?)<\/w:tc>/g;
  const tcBlocks = [];
  let m;
  while ((m = tcRegex.exec(docXml)) !== null) {
    tcBlocks.push(m[1]);
  }

  // Pro Zelle: shaded? + Text
  const cells = tcBlocks.map(inner => {
    // Shading aus tcPr lesen
    const tcPrMatch = inner.match(/<w:tcPr>([\s\S]*?)<\/w:tcPr>/);
    let shaded = false;
    if (tcPrMatch) {
      const shdMatch = tcPrMatch[1].match(/<w:shd\b[^/>]*w:fill="([^"]+)"/);
      if (shdMatch) {
        shaded = isShadedFill(shdMatch[1]);
      }
    }
    // Text aus allen <w:t>...</w:t> innerhalb der Zelle zusammensetzen
    const textParts = [];
    const tRegex = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
    let tm;
    while ((tm = tRegex.exec(inner)) !== null) {
      // XML-Entities zurück
      const t = tm[1]
        .replace(/&amp;/g, "&")
        .replace(/&lt;/g, "<")
        .replace(/&gt;/g, ">")
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'");
      textParts.push(t);
    }
    const text = textParts.join("").trim();
    return { shaded, text };
  });

  // Pro Text die occurrence-Indices (0-basiert) der shaded-Treffer sammeln
  // Beispiel: Text "5" kommt 6x vor, davon shaded an Position 0 und 3.
  // → shadedOccurrences["5"] = [0, 3]
  const counter = new Map(); // text -> running count
  const shadedOccurrences = new Map(); // text -> [occIdx, ...]
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

function normalizeText(s) {
  return String(s).replace(/\s+/g, " ").trim().toLowerCase();
}

// -------------------------------------------------------------
// HTML-POST-PROCESSING: Marker einfügen
// -------------------------------------------------------------

/**
 * Geht das HTML nach <td>-Zellen durch. Pro Zelle:
 *  - Cell-Text normalisieren
 *  - Counter pro Text hochzählen
 *  - Wenn (text, counter) in shadedOccurrences enthalten → Marker einfügen
 *
 * Reihenfolge ist entscheidend: Mammoth gibt Tabellen-Zellen in derselben
 * Reihenfolge aus wie die DOCX-XML sie definiert. Das matcht 1:1 mit
 * unserer Zähl-Logik aus extractShadedCellTexts.
 *
 * Marker-Form: Wir setzen [[SHADED]] DIREKT vor den ersten sichtbaren
 * Zeichen-Block der Zelle. Damit bleibt der Marker an Codes wie "5"
 * sichtbar und der nachfolgende Plain-Text-Cleanup im Hardening kann ihn
 * leicht greifen.
 */
function injectShadedMarkers(html, shadedOccurrences) {
  if (!shadedOccurrences || shadedOccurrences.size === 0) return html;

  const counter = new Map();

  return html.replace(/<td\b([^>]*)>([\s\S]*?)<\/td>/gi, (match, attrs, inner) => {
    // Inneren Text ohne Tags ermitteln für Lookup
    const plain = inner.replace(/<[^>]+>/g, " ").replace(/&nbsp;/g, " ");
    const key = normalizeText(plain);
    if (!key) return match;

    const idx = counter.get(key) || 0;
    counter.set(key, idx + 1);

    const occs = shadedOccurrences.get(key);
    if (!occs || !occs.includes(idx)) return match;

    // Marker einfügen — vor dem ersten Zeichen des ersten Text-Blocks im inner
    // Trick: ersetze den ersten "Text"-Run. Wenn inner mit "<p>...</p>" anfängt,
    // landen wir nach dem öffnenden <p>. Wenn nicht, prefixen wir den inner direkt.
    let newInner;
    const firstTextMatch = inner.match(/>([^<]*\S[^<]*)</);
    if (firstTextMatch) {
      // Erstes nicht-leeres Text-Stück gefunden
      newInner = inner.replace(firstTextMatch[0], `>[[SHADED]]${firstTextMatch[1]}<`);
    } else {
      // Fallback: ganz vorn anflanschen
      newInner = `[[SHADED]]${inner}`;
    }
    return `<td${attrs}>${newInner}</td>`;
  });
}

// -------------------------------------------------------------
// MAIN HANDLER
// -------------------------------------------------------------

export default async function handler(req, res) {
  try {
    if (req.method !== "POST") {
      return res.status(405).json({ error: "Method not allowed" });
    }
    const { fileName, mimeType, data } = req.body ?? {};
    if (!data) {
      return res.status(400).json({ error: "Missing DOCX base64 in 'data'" });
    }
    const buf = toBufferFromBase64(data);

    const headHex = buf.slice(0, 4).toString("hex");
    const tailHex = buf.slice(-4).toString("hex");
    const okZip = isZip(buf);
    const okEOCD = hasEOCD(buf);

    if (!okZip || !okEOCD || buf.length < 1024) {
      return res.status(400).json({
        error: "Invalid DOCX/ZIP payload",
        details: { length: buf.length, headHex, tailHex, zipHeaderOK: okZip, eocdOK: okEOCD }
      });
    }

    // 1) Mammoth: DOCX -> HTML
    const result = await mammoth.convertToHtml({ buffer: buf });
    let html = String(result.value || "").trim();

    // 2) Shading-Extraktion direkt aus document.xml
    //    (Mammoth verliert die Cell-Shading-Info, wir holen sie selbst raus)
    let shadedCount = 0;
    let shadedSampleTexts = [];
    try {
      const zip = await JSZip.loadAsync(buf);
      const docXmlFile = zip.file("word/document.xml");
      if (docXmlFile) {
        const docXml = await docXmlFile.async("string");
        const shadedOccurrences = extractShadedCellTexts(docXml);
        shadedCount = [...shadedOccurrences.values()].reduce((s, arr) => s + arr.length, 0);
        shadedSampleTexts = [...shadedOccurrences.keys()].slice(0, 10);

        if (shadedCount > 0) {
          html = injectShadedMarkers(html, shadedOccurrences);
        }
      }
    } catch (e) {
      // Fail-soft: wenn Shading-Extraktion failt, geben wir das HTML ohne
      // Marker zurück und loggen den Fehler. Der bisherige Workflow läuft
      // damit weiter, nur Mars Q19 wird ohne Marker geparst (= bisheriges
      // Verhalten).
      console.warn("Shading extraction failed (non-fatal):", e.message);
    }

    return res.status(200).json({
      html,
      // Debug-Felder, damit du im n8n-Output sehen kannst, ob Shading erkannt wurde
      _shadedCount: shadedCount,
      _shadedSampleTexts: shadedSampleTexts,
      _version: "docx-html-v2-shading"
    });
  } catch (err) {
    console.error("docx-html error:", err);
    return res.status(500).json({ error: String(err?.message || err) });
  }
}
