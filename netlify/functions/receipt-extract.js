// netlify/functions/receipt-extract.js
// Reads a receipt / invoice (image or PDF) with Claude and returns
// { supplier, invoice_no, amount, date, confidence }.
// Env: ANTHROPIC_API_KEY (required), RECEIPT_MODEL (optional).

const MODEL = process.env.RECEIPT_MODEL || "claude-sonnet-4-6";

exports.handler = async (event) => {
  if (event.httpMethod !== "POST") return { statusCode: 405, body: "POST only" };
  if (!process.env.ANTHROPIC_API_KEY) return { statusCode: 500, body: "ANTHROPIC_API_KEY not set" };

  let body;
  try { body = JSON.parse(event.body || "{}"); } catch { return { statusCode: 400, body: "bad json" }; }
  const { src, name = "", type = "" } = body;
  if (!src) return { statusCode: 400, body: "src required" };

  // Get the bytes as base64 + a media type, from a data: URL or a public URL.
  let b64, mime;
  try {
    if (src.startsWith("data:")) {
      const m = src.match(/^data:([^;]+);base64,(.*)$/s);
      if (!m) return { statusCode: 400, body: "unsupported data url" };
      mime = m[1]; b64 = m[2];
    } else {
      const r = await fetch(src);
      if (!r.ok) return { statusCode: 400, body: `fetch ${r.status}` };
      mime = r.headers.get("content-type") || type || "application/octet-stream";
      b64 = Buffer.from(await r.arrayBuffer()).toString("base64");
    }
  } catch (e) { return { statusCode: 400, body: "could not read file: " + e.message }; }

  const isPdf = mime.includes("pdf") || /\.pdf$/i.test(name);
  const block = isPdf
    ? { type: "document", source: { type: "base64", media_type: "application/pdf", data: b64 } }
    : { type: "image", source: { type: "base64", media_type: mime.split(";")[0], data: b64 } };

  const prompt = `This is a receipt or vendor invoice attached to a construction daily report.
Extract these fields and reply with ONLY a JSON object, no prose, no code fences:
{"supplier": "vendor / store name as printed",
 "invoice_no": "invoice, ticket, or receipt number (string, or null)",
 "amount": total amount due as a number (grand total after tax; null if not visible),
 "date": "invoice or purchase date as YYYY-MM-DD (null if not visible)",
 "confidence": "high" | "medium" | "low"}
If the document is not a receipt or invoice, return {"supplier":null,"invoice_no":null,"amount":null,"date":null,"confidence":"low"}.`;

  try {
    const res = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "x-api-key": process.env.ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
      },
      body: JSON.stringify({
        model: MODEL, max_tokens: 300,
        messages: [{ role: "user", content: [block, { type: "text", text: prompt }] }],
      }),
    });
    const data = await res.json();
    if (!res.ok) return { statusCode: 502, body: JSON.stringify(data) };
    const text = (data.content || []).filter(c => c.type === "text").map(c => c.text).join("\n");
    const clean = text.replace(/```json|```/g, "").trim();
    const start = clean.indexOf("{"), end = clean.lastIndexOf("}");
    const parsed = JSON.parse(clean.slice(start, end + 1));
    // normalise
    const out = {
      supplier: parsed.supplier ? String(parsed.supplier).trim() : null,
      invoice_no: parsed.invoice_no != null ? String(parsed.invoice_no).trim() : null,
      amount: parsed.amount != null && !isNaN(parseFloat(parsed.amount)) ? Math.round(parseFloat(parsed.amount) * 100) / 100 : null,
      date: /^\d{4}-\d{2}-\d{2}$/.test(String(parsed.date || "")) ? parsed.date : null,
      confidence: parsed.confidence || "medium",
    };
    return { statusCode: 200, headers: { "content-type": "application/json" }, body: JSON.stringify(out) };
  } catch (e) {
    return { statusCode: 500, body: "extract failed: " + e.message };
  }
};
