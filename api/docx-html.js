import mammoth from "mammoth";

/**
 * Hilfen: robustes Base64 -> Buffer + ZIP/DOCX-Checks
 */
function toBufferFromBase64(input) {
  const s = String(input ?? "").trim();
  const clean = s.includes("base64,") ? s.split("base64,").pop() : s;
  return Buffer.from(clean, "base64");
}
function isZip(buf) {
  // ZIP beginnt mit "PK" 0x50 0x4B
  return buf.length >= 4 && buf[0] === 0x50 && buf[1] === 0x4B;
}
function hasEOCD(buf) {
  // End Of Central Directory-Signatur: 0x50 0x4B 0x05 0x06 irgendwo im letzten Fenster
  const sig = Buffer.from([0x50, 0x4B, 0x05, 0x06]);
  const start = Math.max(0, buf.length - 70000);
  return buf.slice(start).indexOf(sig) !== -1;
}

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

    // Diagnostik vor Mammoth: Länge + Signaturen prüfen
    const headHex = buf.slice(0, 4).toString("hex");       // sollte '504b0304' o.ä. sein
    const tailHex = buf.slice(-4).toString("hex");         // meist '504b0506' im Umfeld
    const okZip   = isZip(buf);
    const okEOCD  = hasEOCD(buf);

    if (!okZip || !okEOCD || buf.length < 1024) {
      // Detail-Fehler zurückgeben (400), damit du es direkt siehst
      return res.status(400).json({
        error: "Invalid DOCX/ZIP payload",
        details: {
          length: buf.length,
          headHex,
          tailHex,
          zipHeaderOK: okZip,
          eocdOK: okEOCD
        }
      });
    }

    // DOCX -> HTML
    const result = await mammoth.convertToHtml({ buffer: buf });
    const html = String(result.value || "").trim();

    return res.status(200).json({ html });

  } catch (err) {
    console.error("docx-html error:", err);
    return res.status(500).json({ error: String(err?.message || err) });
  }
}
