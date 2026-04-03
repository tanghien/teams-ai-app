import mammoth from "mammoth";
import pdfParse from "pdf-parse";
import XLSX from "xlsx";
import officeParser from "officeparser";

// ─────────────────────────────────────────────────────────────
// CONFIG
// ─────────────────────────────────────────────────────────────
const MAX_CHARS = 15000;
const MAX_FILE_MB = 20;

const TEXT_EXTS = new Set([".txt", ".md", ".csv", ".json", ".xml", ".html", ".htm", ".log"]);
const PDF_EXTS = new Set([".pdf"]);
const DOCX_EXTS = new Set([".docx"]);
const XLSX_EXTS = new Set([".xlsx", ".xls"]);
const PPTX_EXTS = new Set([".pptx"]);

const SUPPORTED_EXTS = new Set([
  ...TEXT_EXTS,
  ...PDF_EXTS,
  ...DOCX_EXTS,
  ...XLSX_EXTS,
  ...PPTX_EXTS
]);

// ─────────────────────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────────────────────
function truncate(text) {
  return text.length > MAX_CHARS
    ? text.substring(0, MAX_CHARS) + "\n...[bị cắt bớt]"
    : text;
}

function isRateLimitError(e) {
  return (
    e.status === 429 ||
    /quota|rate|limit|exceeded|too many/i.test(e.message || "")
  );
}

// ─────────────────────────────────────────────────────────────
// LLM PROVIDERS
// ─────────────────────────────────────────────────────────────

// 🥇 GROQ
async function callGroq(prompt, systemPrompt, maxTokens, key) {
  const r = await fetch("https://api.groq.com/openai/v1/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${key}`
    },
    body: JSON.stringify({
      model: "llama-3.3-70b-versatile",
      max_tokens: maxTokens,
      temperature: 0.3,
      messages: [
        ...(systemPrompt ? [{ role: "system", content: systemPrompt }] : []),
        { role: "user", content: prompt }
      ]
    })
  });

  const data = await r.json();
  if (data.error) {
    const err = new Error(data.error.message);
    err.status = r.status;
    throw err;
  }

  return data.choices?.[0]?.message?.content?.trim() || "";
}

// 🥈 GEMINI
async function callGemini(prompt, systemPrompt, maxTokens, key) {
  const r = await fetch(
    `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${key}`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: {
          maxOutputTokens: maxTokens,
          temperature: 0.3
        }
      })
    }
  );

  const data = await r.json();
  if (data.error) {
    const err = new Error(data.error.message);
    err.status = r.status;
    throw err;
  }

  return data.candidates?.[0]?.content?.parts?.[0]?.text || "";
}

// 🥉 OPENROUTER (GIẢM MODEL)
async function callOpenRouter(prompt, systemPrompt, maxTokens, key) {
  const models = [
    "openrouter/free",
    "deepseek/deepseek-r1:free"
  ];

  for (const model of models) {
    try {
      const r = await fetch("https://openrouter.ai/api/v1/chat/completions", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${key}`
        },
        body: JSON.stringify({
          model,
          max_tokens: maxTokens,
          temperature: 0.3,
          messages: [
            ...(systemPrompt ? [{ role: "system", content: systemPrompt }] : []),
            { role: "user", content: prompt }
          ]
        })
      });

      const data = await r.json();
      if (data.error) continue;

      const content = data.choices?.[0]?.message?.content;
      if (content) return content;
    } catch {}
  }

  throw new Error("OpenRouter unavailable");
}

// HF
async function callHF(prompt, systemPrompt, maxTokens, key) {
  const r = await fetch("https://router.huggingface.co/v1/chat/completions", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${key}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      model: "meta-llama/Llama-3.1-8B-Instruct:fastest",
      max_tokens: maxTokens,
      messages: [{ role: "user", content: prompt }]
    })
  });

  const data = await r.json();
  if (data.error) throw new Error(data.error.message);

  return data.choices?.[0]?.message?.content || "";
}

// NVIDIA
async function callNvidia(prompt, systemPrompt, maxTokens, key) {
  const r = await fetch("https://integrate.api.nvidia.com/v1/chat/completions", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${key}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      model: "thudm/glm-4-7b",
      max_tokens: maxTokens,
      messages: [{ role: "user", content: prompt }]
    })
  });

  const data = await r.json();
  if (data.error) throw new Error(data.error.message);

  return data.choices?.[0]?.message?.content || "";
}

// 🔀 MASTER FALLBACK
async function callLLM(prompt, systemPrompt, maxTokens, env) {
  const {
    GROQ_API_KEY,
    GEMINI_API_KEY,
    OPENROUTER_API_KEY,
    HF_TOKEN,
    NVIDIA_API_KEY
  } = env;

  // 1. GROQ
  if (GROQ_API_KEY) {
    try { return await callGroq(prompt, systemPrompt, maxTokens, GROQ_API_KEY); }
    catch (e) { if (!isRateLimitError(e)) throw e; }
  }

  // 2. GEMINI
  if (GEMINI_API_KEY) {
    try { return await callGemini(prompt, systemPrompt, maxTokens, GEMINI_API_KEY); }
    catch (e) { if (!isRateLimitError(e)) throw e; }
  }

  // 3. OPENROUTER
  if (OPENROUTER_API_KEY) {
    try { return await callOpenRouter(prompt, systemPrompt, maxTokens, OPENROUTER_API_KEY); }
    catch {}
  }

  // 4. HF
  if (HF_TOKEN) {
    try { return await callHF(prompt, systemPrompt, maxTokens, HF_TOKEN); }
    catch {}
  }

  // 5. NVIDIA
  if (NVIDIA_API_KEY) {
    return await callNvidia(prompt, systemPrompt, maxTokens, NVIDIA_API_KEY);
  }

  throw new Error("No LLM available");
}

// ─────────────────────────────────────────────────────────────
// FILE PARSERS
// ─────────────────────────────────────────────────────────────
async function extractText(buffer, ext) {
  if (TEXT_EXTS.has(ext)) return truncate(buffer.toString());

  if (PDF_EXTS.has(ext)) {
    const data = await pdfParse(buffer);
    return truncate(data.text || "");
  }

  if (DOCX_EXTS.has(ext)) {
    const r = await mammoth.extractRawText({ buffer });
    return truncate(r.value || "");
  }

  if (XLSX_EXTS.has(ext)) {
    const wb = XLSX.read(buffer, { type: "buffer" });
    return truncate(
      wb.SheetNames.map(n => XLSX.utils.sheet_to_csv(wb.Sheets[n])).join("\n")
    );
  }

  if (PPTX_EXTS.has(ext)) {
    return truncate(
      await new Promise((res, rej) =>
        officeParser.parseOffice(buffer, (d, e) => e ? rej(e) : res(d))
      )
    );
  }

  throw new Error("Unsupported file");
}

// ─────────────────────────────────────────────────────────────
// MAIN HANDLER
// ─────────────────────────────────────────────────────────────
export default async function handler(req, res) {
  if (req.method !== "POST")
    return res.status(405).json({ error: "POST only" });

  try {
    let body = typeof req.body === "string"
      ? JSON.parse(req.body)
      : req.body;

    const question = body?.question?.trim();
    if (!question)
      return res.status(400).json({ error: "Missing question" });

    // ─── AUTH ─────────────────
    const tokenRes = await fetch(
      `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          client_id: process.env.AZURE_CLIENT_ID,
          client_secret: process.env.AZURE_CLIENT_SECRET,
          scope: "https://graph.microsoft.com/.default",
          grant_type: "client_credentials"
        })
      }
    );

    const tokenData = await tokenRes.json();
    const accessToken = tokenData.access_token;

    // ─── GET FILE ─────────────────
    const siteRes = await fetch(
      `https://graph.microsoft.com/v1.0/sites/tbcball.sharepoint.com:/sites/${process.env.SHAREPOINT_SITE}`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );

    const site = await siteRes.json();

    const drivesRes = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${site.id}/drives`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );

    const drives = (await drivesRes.json()).value;
    const drive = drives[0];

    const filesRes = await fetch(
      `https://graph.microsoft.com/v1.0/drives/${drive.id}/root/children`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );

    const files = (await filesRes.json()).value;

    const fileList = files
      .map((f, i) => `${i + 1}. ${f.name}`)
      .join("\n");

    const idx = parseInt(
      await callLLM(
        `Q: ${question}\n\n${fileList}`,
        "Chọn file, trả số",
        20,
        process.env
      )
    );

    if (!idx || !files[idx - 1]) {
      return res.json({ answer: "Không tìm thấy file phù hợp" });
    }

    const file = files[idx - 1];

    const dl = await fetch(
      `https://graph.microsoft.com/v1.0/drives/${drive.id}/items/${file.id}/content`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );

    const buffer = Buffer.from(await dl.arrayBuffer());
    const ext = file.name.slice(file.name.lastIndexOf(".")).toLowerCase();

    const text = await extractText(buffer, ext);

    const answer = await callLLM(
      `Tài liệu:\n${text}\n\nCâu hỏi:${question}`,
      "Trả lời ngắn gọn tiếng Việt",
      1000,
      process.env
    );

    res.json({ answer });

  } catch (e) {
    res.status(500).json({ error: e.message });
  }
}