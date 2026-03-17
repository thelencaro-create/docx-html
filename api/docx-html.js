import mammoth from "mammoth";

export default async function handler(req, res) {
  try {
    if (req.method !== "POST") {
      return res.status(405).json({ error: "Method not allowed" });
    }

    const { fileName, mimeType, data } = req.body ?? {};
    if (!data) {
      return res.status(400).json({ error: "Missing DOCX base64 in 'data'" });
    }

    const base64 = String(data).includes("base64,")
      ? String(data).split("base64,").pop()
      : String(data);
    const buffer = Buffer.from(base64, "base64");

    const result = await mammoth.convertToHtml({ buffer });
    const html = String(result.value || "").trim();

    return res.status(200).json({ html });
  } catch (err) {
    console.error("docx-html error:", err);
    return res.status(500).json({ error: String(err?.message || err) });
  }
}
