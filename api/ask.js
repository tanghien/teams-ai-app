import mammoth from "mammoth";
import pdfParse from "pdf-parse";
import XLSX from "xlsx";
import officeParser from "officeparser";
import { promisify } from "util";

const parseOffice = promisify(officeParser.parseOfficeAsync?.bind(officeParser) ?? officeParser.parseOffice?.bind(officeParser));

export default async function handler(req, res) {
  if (!req || !res) return;
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed. Use POST." });
  }

  try {
    // ─── 1. Parse Body ───────────────────────────────────────────────────────
    let body = req.body;
    if (typeof body === "string") {
      try { body = JSON.parse(body); } catch { body = {}; }
    }
    if (!body || typeof body !== "object") body = {};

    const question = (body.question ?? "").trim();
    if (!question) return res.status(400).json({ error: "Thiếu tham số 'question'." });

    // ─── 2. Environment Variables ────────────────────────────────────────────
    const {
      AZURE_TENANT_ID,
      AZURE_CLIENT_ID,
      AZURE_CLIENT_SECRET,
      GROQ_API_KEY,
      DEEPSEEK_API_KEY,
      GEMINI_API_KEY,
      OPENROUTER_API_KEY,
      HF_TOKEN,
      NVIDIA_API_KEY
    } = process.env;

    if (!AZURE_TENANT_ID || !AZURE_CLIENT_ID || !AZURE_CLIENT_SECRET)
      return res.status(500).json({ error: "Thiếu biến môi trường Azure." });

    const hasAnyLLM = GROQ_API_KEY || DEEPSEEK_API_KEY || OPENROUTER_API_KEY || HF_TOKEN || NVIDIA_API_KEY || GEMINI_API_KEY;
    if (!hasAnyLLM)
      return res.status(500).json({ error: "Cần ít nhất một API key free: GROQ, OPENROUTER, HF, NVIDIA hoặc GEMINI." });

    // ─── 3. LLM PROVIDERS (FREE TIER — FALLBACK CHAIN 1→2→3→4→5) ────────────

    // 🔍 Chỉ fallback khi đúng lỗi rate limit / quota — lỗi thật thì throw luôn
    function isRateLimitError(e) {
      return (
        e.status === 429 ||
        e.message?.toLowerCase().includes("quota") ||
        e.message?.toLowerCase().includes("rate limit") ||
        e.message?.toLowerCase().includes("rate_limit") ||
        e.message?.toLowerCase().includes("too many") ||
        e.message?.toLowerCase().includes("exceeded") ||
        e.message?.toLowerCase().includes("request too large") ||  // token vượt TPM limit
        e.message?.toLowerCase().includes("reduce your message")   // Groq TPM message
      );
    }

    // 🔹 1. Groq — Free tier, nhanh nhất, ưu tiên số 1
    // model mặc định: llama-3.1-8b-instant (14,400 req/ngày)
    // model thứ 2:   llama-3.3-70b-versatile (mạnh hơn, RPD thấp hơn)
    async function callGroq(prompt, systemPrompt = "", maxTokens = 1024, model = "llama-3.1-8b-instant") {
      if (!GROQ_API_KEY) throw new Error("NO_GROQ_KEY");
      console.log("[LLM:1/5] → Calling Groq (Free)...");
      const r = await fetch("https://api.groq.com/openai/v1/chat/completions", {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${GROQ_API_KEY}` },
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
      const remaining = r.headers.get("x-ratelimit-remaining-requests");
      if (remaining) console.log(`[Groq] Remaining: ${remaining} requests`);
      const data = await r.json();
      if (data.error) {
        const err = new Error(data.error.message || "Groq error");
        err.status = r.status;
        if (r.status === 429) console.warn("[Groq] RATE LIMIT (429) — switching to next");
        throw err;
      }
      console.log("[Groq] ✓ Success");
      return data.choices?.[0]?.message?.content?.trim() ?? "";
    }

    // 🔹 2. DeepSeek — 5M token free khi đăng ký, không rate limit cứng
    async function callDeepSeek(prompt, systemPrompt = "", maxTokens = 1024) {
      if (!DEEPSEEK_API_KEY) throw new Error("NO_DEEPSEEK_KEY");
      console.log("[LLM:2/6] → Calling DeepSeek (Free 5M tokens)...");
      const r = await fetch("https://api.deepseek.com/chat/completions", {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${DEEPSEEK_API_KEY}` },
        body: JSON.stringify({
          model: "deepseek-chat", // DeepSeek-V3.2, 128K context
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
        const err = new Error(data.error.message || "DeepSeek error");
        err.status = r.status;
        console.warn(`[DeepSeek] Error: ${err.message}`);
        throw err;
      }
      console.log("[DeepSeek] ✓ Success");
      return data.choices?.[0]?.message?.content?.trim() ?? "";
    }

    // 🔹 3. OpenRouter — Free models
    // Thử lần lượt: openrouter/free (auto-router) → deepseek-r1 → llama-4-maverick → qwen3
    async function callOpenRouter(prompt, systemPrompt = "", maxTokens = 1024) {
      if (!OPENROUTER_API_KEY) throw new Error("NO_OPENROUTER_KEY");

      // openrouter/free: tự động chọn model free đang hoạt động — không bao giờ bị "No endpoints"
      const OR_MODELS = [
        "openrouter/free",                    // auto-router, ưu tiên số 1
        "deepseek/deepseek-r1:free",          // fallback 1
        "meta-llama/llama-4-maverick:free",   // fallback 2
        "qwen/qwen3-235b-a22b:free",          // fallback 3
        "deepseek/deepseek-chat-v3.1:free",   // fallback 4
      ];

      const headers = {
        "Content-Type": "application/json",
        Authorization: `Bearer ${OPENROUTER_API_KEY}`,
        "HTTP-Referer": "https://yourdomain.com",
        "X-Title": "AI Docs Agent"
      };

      for (const model of OR_MODELS) {
        console.log(`[LLM:2/5] → OpenRouter trying model: ${model}`);
        try {
          const r = await fetch("https://openrouter.ai/api/v1/chat/completions", {
            method: "POST",
            headers,
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
          if (data.error) {
            const msg = data.error.message || "OpenRouter error";
            // Lỗi "No endpoints" hoặc model không tồn tại → thử model tiếp theo
            if (msg.toLowerCase().includes("no endpoints") || msg.toLowerCase().includes("not found")) {
              console.warn(`[OpenRouter] Model ${model} unavailable → trying next`);
              continue;
            }
            const err = new Error(msg);
            err.status = r.status;
            throw err;
          }
          const content = data.choices?.[0]?.message?.content?.trim() ?? "";
          if (!content) { console.warn(`[OpenRouter] Empty response from ${model} → trying next`); continue; }
          console.log(`[OpenRouter] ✓ Success via ${model}`);
          return content;
        } catch (e) {
          // Nếu là lỗi rate limit → throw lên callLLM để fallback sang HuggingFace
          if (isRateLimitError(e)) throw e;
          // Lỗi khác → thử model tiếp theo trong list
          console.warn(`[OpenRouter] ${model} error: ${e.message} → trying next`);
        }
      }
      throw new Error("Tất cả OpenRouter free models đều không khả dụng.");
    }

    // 🔹 3. Hugging Face — Free inference API (thử lần lượt các model)
    async function callHuggingFace(prompt, systemPrompt = "", maxTokens = 1024) {
      if (!HF_TOKEN) throw new Error("NO_HF_TOKEN");

      const HF_MODELS = [
        "meta-llama/Llama-3.1-8B-Instruct",          // stable, không có :fastest
        "meta-llama/Llama-3.2-3B-Instruct",           // nhỏ hơn, nhanh hơn
        "Qwen/Qwen2.5-7B-Instruct",                   // backup tốt
        "mistralai/Mistral-7B-Instruct-v0.3",         // ổn định lâu dài
      ];

      const headers = {
        "Content-Type": "application/json",
        Authorization: `Bearer ${HF_TOKEN}`
      };

      for (const model of HF_MODELS) {
        console.log(`[LLM:3/5] → HuggingFace trying model: ${model}`);
        try {
          const r = await fetch("https://router.huggingface.co/v1/chat/completions", {
            method: "POST",
            headers,
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
          if (data.error) {
            const msg = data.error.message || "HuggingFace error";
            // Rate limit → throw lên callLLM để fallback sang NVIDIA
            if (isRateLimitError({ message: msg, status: r.status })) throw Object.assign(new Error(msg), { status: r.status });
            // Model không khả dụng / lỗi endpoint → thử model tiếp
            console.warn(`[HuggingFace] Model ${model} error: ${msg} → trying next`);
            continue;
          }
          const text = data.choices?.[0]?.message?.content?.trim() ?? "";
          if (!text) { console.warn(`[HuggingFace] Empty response from ${model} → trying next`); continue; }
          console.log(`[HuggingFace] ✓ Success via ${model}`);
          return text;
        } catch (e) {
          if (isRateLimitError(e)) throw e; // rate limit → callLLM xử lý fallback
          console.warn(`[HuggingFace] ${model} error: ${e.message} → trying next`);
        }
      }
      throw new Error("Tất cả HuggingFace models đều không khả dụng.");
    }

    // 🔹 4. NVIDIA NIM — Free tier
    async function callNvidia(prompt, systemPrompt = "", maxTokens = 1024) {
      if (!NVIDIA_API_KEY) throw new Error("NO_NVIDIA_KEY");
      console.log("[LLM:4/5] → Calling NVIDIA NIM (Free)...");
      const r = await fetch("https://integrate.api.nvidia.com/v1/chat/completions", {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${NVIDIA_API_KEY}` },
        body: JSON.stringify({
          model: "thudm/glm-4-7b",
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
        const err = new Error(data.error.message || "NVIDIA error");
        err.status = r.status;
        console.warn(`[NVIDIA] Error: ${err.message}`);
        throw err;
      }
      console.log("[NVIDIA] ✓ Success");
      return data.choices?.[0]?.message?.content?.trim() ?? "";
    }

    // 🔹 5. Gemini Free — Last resort (20 req/ngày, text only)
    async function callGemini(prompt, systemPrompt = "", maxTokens = 1024) {
      if (!GEMINI_API_KEY) throw new Error("NO_GEMINI_KEY");
      console.log("[LLM:5/5] → Calling Gemini Free (LAST RESORT)...");
      const r = await fetch(
        `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_API_KEY}`,
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            ...(systemPrompt && { system_instruction: { parts: [{ text: systemPrompt }] } }),
            contents: [{ parts: [{ text: prompt }] }],
            generationConfig: { maxOutputTokens: maxTokens, temperature: 0.3 }
          })
        }
      );
      const data = await r.json();
      if (data.error) {
        if (data.error.message?.includes("quota") || r.status === 429) {
          console.error("[Gemini] QUOTA EXCEEDED — No more free providers available!");
          throw new Error("HẾT LIMIT TẤT CẢ PROVIDER FREE. Vui lòng thử lại sau hoặc nâng cấp API key.");
        }
        console.error(`[Gemini] Error: ${data.error.message}`);
        throw new Error(`Gemini error: ${data.error.message}`);
      }
      console.log("[Gemini] ✓ Success");
      return data.candidates?.[0]?.content?.parts?.[0]?.text?.trim() ?? "";
    }

    // 🔀 Fallback chain theo thứ tự tối ưu:
    //   1. Groq llama-3.3-70b-versatile — mạnh, TPM cao hơn 8b, context lớn
    //   2. DeepSeek deepseek-chat       — 5M token free, không rate limit cứng, 128K context
    //   3. Groq llama-3.1-8b-instant    — nhanh, 14,400 req/ngày nhưng TPM chỉ 6000
    //   4. HuggingFace                  — ổn định, limit không cố định
    //   5. NVIDIA NIM                   — backup
    //   6. OpenRouter/free              — auto-router, 200 req/ngày (để dành, hay bị lỗi endpoint)
    //   7. Gemini Free                  — last resort, 20 req/ngày
    async function callLLM(prompt, systemPrompt = "", maxTokens = 1024) {
      // 1️⃣ Groq llama-3.3-70b-versatile: TPM cao hơn 8b, xử lý tốt file lớn
      if (GROQ_API_KEY) {
        try { return await callGroq(prompt, systemPrompt, maxTokens, "llama-3.3-70b-versatile"); }
        catch (e) {
          if (isRateLimitError(e)) console.warn("[Fallback 1→2] Groq 70b hết quota/TPM → thử DeepSeek");
          else throw e;
        }
      }
      // 2️⃣ DeepSeek: 5M token free, không rate limit cứng, context 128K — rất phù hợp file lớn
      if (DEEPSEEK_API_KEY) {
        try { return await callDeepSeek(prompt, systemPrompt, maxTokens); }
        catch (e) {
          if (isRateLimitError(e)) console.warn("[Fallback 2→3] DeepSeek hết token free → thử Groq 8b");
          else throw e;
        }
      }
      // 3️⃣ Groq llama-3.1-8b-instant: 14,400 req/ngày — tốt cho prompt ngắn/vừa
      if (GROQ_API_KEY) {
        try { return await callGroq(prompt, systemPrompt, maxTokens, "llama-3.1-8b-instant"); }
        catch (e) {
          if (isRateLimitError(e)) console.warn("[Fallback 3→4] Groq 8b hết quota/TPM → thử HuggingFace");
          else throw e;
        }
      }
      // 4️⃣ HuggingFace: ổn định, không có hard limit/ngày rõ ràng
      if (HF_TOKEN) {
        try { return await callHuggingFace(prompt, systemPrompt, maxTokens); }
        catch (e) {
          if (isRateLimitError(e)) console.warn("[Fallback 4→5] HuggingFace hết quota → thử NVIDIA");
          else throw e;
        }
      }
      // 5️⃣ NVIDIA NIM: backup tốt, limit ~1000 req/tháng
      if (NVIDIA_API_KEY) {
        try { return await callNvidia(prompt, systemPrompt, maxTokens); }
        catch (e) {
          if (isRateLimitError(e)) console.warn("[Fallback 5→6] NVIDIA hết quota → thử OpenRouter");
          else throw e;
        }
      }
      // 6️⃣ OpenRouter: để dành gần cuối vì hay gặp lỗi endpoint, 200 req/ngày
      if (OPENROUTER_API_KEY) {
        try { return await callOpenRouter(prompt, systemPrompt, maxTokens); }
        catch (e) {
          if (isRateLimitError(e)) console.warn("[Fallback 6→7] OpenRouter hết quota → thử Gemini");
          else throw e;
        }
      }
      // 7️⃣ Gemini Free: LAST RESORT — chỉ 20 req/ngày, tiết kiệm tối đa
      console.warn("[Fallback] Gemini Free LAST RESORT (20 req/ngày) — tất cả provider khác đã hết quota");
      return await callGemini(prompt, systemPrompt, maxTokens);
    }

    // ─── 4. Local File Parsers ────────────────────────────────────────────────

    const TEXT_EXTS    = new Set([".txt", ".md", ".csv", ".json", ".xml", ".html", ".htm", ".log"]);
    const PDF_EXTS     = new Set([".pdf"]);
    const DOCX_EXTS    = new Set([".docx"]);
    const XLSX_EXTS    = new Set([".xlsx", ".xls"]);
    const PPTX_EXTS    = new Set([".pptx"]);
    const MAX_CHARS    = 15000; // giới hạn ký tự gửi LLM
    const MAX_FILE_MB  = 20;    // giới hạn kích thước file tải về

    function truncate(text) {
      return text.length > MAX_CHARS ? text.substring(0, MAX_CHARS) + "\n...[bị cắt bớt]" : text;
    }

    // Đọc nội dung text thuần
    async function parseText(buffer) {
      return buffer.toString("utf-8");
    }

    // Đọc PDF bằng pdf-parse (không cần Gemini Vision)
    async function parsePdf(buffer) {
      try {
        const data = await pdfParse(buffer);
        const text = data.text?.trim() ?? "";
        if (!text) throw new Error("pdf-parse không trích xuất được text (có thể là PDF scan).");
        console.log(`[PDF] Extracted ${text.length} chars, ${data.numpages} pages`);
        return text;
      } catch (e) {
        console.warn(`[PDF] pdf-parse failed: ${e.message}`);
        throw e;
      }
    }

    // Đọc DOCX bằng mammoth
    async function parseDocx(buffer) {
      const result = await mammoth.extractRawText({ buffer });
      const text = result.value?.trim() ?? "";
      if (!text) throw new Error("mammoth không trích xuất được nội dung .docx.");
      console.log(`[DOCX] Extracted ${text.length} chars`);
      return text;
    }

    // Đọc XLSX/XLS bằng xlsx
    async function parseXlsx(buffer) {
      const wb = XLSX.read(buffer, { type: "buffer" });
      const sheets = wb.SheetNames.map(name => {
        const csv = XLSX.utils.sheet_to_csv(wb.Sheets[name]);
        return `=== Sheet: ${name} ===\n${csv}`;
      });
      const text = sheets.join("\n\n").trim();
      if (!text) throw new Error("xlsx không trích xuất được nội dung.");
      console.log(`[XLSX] Extracted ${text.length} chars, ${wb.SheetNames.length} sheets`);
      return text;
    }

    // Đọc PPTX bằng officeparser
    async function parsePptx(buffer) {
      try {
        // officeparser nhận buffer hoặc path
        const text = await new Promise((resolve, reject) => {
          officeParser.parseOffice(buffer, (data, err) => {
            if (err) reject(err);
            else resolve(data);
          });
        });
        const result = (text ?? "").trim();
        if (!result) throw new Error("officeparser không trích xuất được nội dung .pptx.");
        console.log(`[PPTX] Extracted ${result.length} chars`);
        return result;
      } catch (e) {
        console.warn(`[PPTX] officeparser failed: ${e.message}`);
        throw e;
      }
    }

    // Router: chọn parser theo ext, trả về text đã truncate
    async function extractText(buffer, ext) {
      if (TEXT_EXTS.has(ext))  return truncate(await parseText(buffer));
      if (PDF_EXTS.has(ext))   return truncate(await parsePdf(buffer));
      if (DOCX_EXTS.has(ext))  return truncate(await parseDocx(buffer));
      if (XLSX_EXTS.has(ext))  return truncate(await parseXlsx(buffer));
      if (PPTX_EXTS.has(ext))  return truncate(await parsePptx(buffer));
      throw new Error(`Định dạng "${ext}" chưa được hỗ trợ. Dùng: TXT, PDF, DOCX, XLSX, XLS, PPTX.`);
    }

    // ─── 5. SharePoint Auth & File Listing ───────────────────────────────────
    const tokenRes = await fetch(
      `https://login.microsoftonline.com/${AZURE_TENANT_ID}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          client_id: AZURE_CLIENT_ID,
          client_secret: AZURE_CLIENT_SECRET,
          scope: "https://graph.microsoft.com/.default",
          grant_type: "client_credentials"
        })
      }
    );
    const tokenData = await tokenRes.json();
    if (!tokenData.access_token)
      return res.status(502).json({ error: "Lấy token thất bại", detail: tokenData });
    const accessToken = tokenData.access_token;

    const siteRes = await fetch(
      `https://graph.microsoft.com/v1.0/sites/tbcball.sharepoint.com:/sites/${process.env.SHAREPOINT_SITE}`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );
    const siteData = await siteRes.json();
    if (!siteData.id)
      return res.status(502).json({ error: "Không lấy được site ID", detail: siteData });
    const siteId = siteData.id;

    const drivesRes = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drives?$select=id,name`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );
    const drivesData = await drivesRes.json();
    const drives = drivesData.value || [];

    if (process.env.DEBUG_FILES === "1")
      return res.status(200).json({ _debug: true, drives });

    const targetDrive =
      drives.find(d => d.name?.toLowerCase().includes("approved sop")) ||
      drives.find(d => d.name?.toLowerCase().includes("document")) ||
      drives[0];

    if (!targetDrive)
      return res.status(502).json({ error: "Không tìm thấy Document Library nào.", drives });

    const driveId = targetDrive.id;
    const allFiles = [];

    async function fetchChildren(itemId, depth = 0) {
      if (depth > 3 || allFiles.length >= 200) return;
      const url = itemId === "root"
        ? `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children?$select=id,name,size,file,folder,parentReference`
        : `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/children?$select=id,name,size,file,folder,parentReference`;
      const r = await fetch(url, { headers: { Authorization: `Bearer ${accessToken}` } });
      const d = await r.json();
      for (const item of (d.value || [])) {
        if (item.file) {
          allFiles.push({
            id: item.id,
            name: item.name,
            size: item.size || 0,
            path: item.parentReference?.path?.replace(`/drives/${driveId}/root:`, "") || ""
          });
        } else if (item.folder) {
          await fetchChildren(item.id, depth + 1);
        }
      }
    }
    await fetchChildren("root");

    if (allFiles.length === 0)
      return res.status(200).json({
        answer: "Không tìm thấy file nào trong thư viện tài liệu.",
        _debug: { driveId, driveName: targetDrive.name }
      });

    // ─── 6. AI Chọn File ─────────────────────────────────────────────────────
    const SUPPORTED_EXTS = new Set([...TEXT_EXTS, ...PDF_EXTS, ...DOCX_EXTS, ...XLSX_EXTS, ...PPTX_EXTS]);

    // Chỉ hiển thị file có định dạng hỗ trợ cho AI chọn
    const supportedFiles = allFiles.filter(f => {
      const ext = f.name.substring(f.name.lastIndexOf(".")).toLowerCase();
      return SUPPORTED_EXTS.has(ext);
    });

    const fileList = supportedFiles
      .map((f, i) => `${i + 1}. [${f.path || "/"}] ${f.name} (${Math.round(f.size / 1024)} KB)`)
      .join("\n");

    const selectedIndexStr = await callLLM(
      `Câu hỏi: "${question}"\n\nDanh sách file:\n${fileList}`,
      "Chọn file liên quan nhất đến câu hỏi. Trả lời CHỈ bằng số thứ tự (ví dụ: 5). Nếu không có file liên quan, trả lời: 0.",
      50
    );
    const selectedIndex = parseInt(selectedIndexStr.trim(), 10);

    let answer = "";
    let selectedFile = null;
    let usedProvider = "none";
    const systemPrompt = "Bạn là trợ lý AI tra cứu tài liệu nội bộ. Trả lời ngắn gọn và chính xác bằng tiếng Việt.";

    if (selectedIndex > 0 && selectedIndex <= supportedFiles.length) {
      selectedFile = supportedFiles[selectedIndex - 1];
      const ext = selectedFile.name.substring(selectedFile.name.lastIndexOf(".")).toLowerCase();

      // Kiểm tra kích thước trước khi tải
      const fileSizeMB = selectedFile.size / (1024 * 1024);
      if (fileSizeMB > MAX_FILE_MB) {
        answer = `File "${selectedFile.name}" quá lớn (${fileSizeMB.toFixed(1)} MB). Giới hạn hiện tại là ${MAX_FILE_MB} MB.`;
      } else {
        // Tải file
        const dlRes = await fetch(
          `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${selectedFile.id}/content`,
          { headers: { Authorization: `Bearer ${accessToken}` } }
        );

        if (!dlRes.ok) {
          answer = `Không tải được file "${selectedFile.name}" (HTTP ${dlRes.status}).`;
        } else {
          const arrayBuffer = await dlRes.arrayBuffer();
          const buffer = Buffer.from(arrayBuffer);

          try {
            console.log(`[Parser] Extracting text from "${selectedFile.name}" (${ext})...`);
            const extractedText = await extractText(buffer, ext);

            answer = await callLLM(
              `Nội dung tài liệu:\n\n${extractedText}\n\n---\n\nCâu hỏi: ${question}`,
              systemPrompt,
              1024
            );
            usedProvider = `local-parse(${ext})->llm-free-chain`;
          } catch (parseErr) {
            console.error(`[Parser] Failed for ${ext}: ${parseErr.message}`);
            answer = `Không thể đọc file "${selectedFile.name}": ${parseErr.message}`;
          }
        }
      }
    } else {
      // Không tìm thấy file phù hợp → gợi ý
      answer = await callLLM(
        `Câu hỏi: "${question}"\n\nDanh sách file hiện có:\n${fileList.substring(0, 3000)}`,
        "Không tìm thấy file phù hợp. Hãy gợi ý người dùng nên tìm trong file nào dựa trên danh sách.",
        512
      );
      usedProvider = "fallback-llm-free-chain";
    }

    return res.status(200).json({
      answer: answer || "Không nhận được câu trả lời.",
      meta: {
        fileSelected: selectedFile ? `${selectedFile.path}/${selectedFile.name}` : null,
        totalFiles: allFiles.length,
        supportedFiles: supportedFiles.length,
        library: targetDrive.name,
        provider: usedProvider,
        supportedFormats: [...SUPPORTED_EXTS].join(", "),
        note: "Đang dùng 100% free tier. PDF/DOCX/XLSX/PPTX được đọc local — không cần Gemini Vision."
      }
    });

  } catch (err) {
    console.error("[ask.js] Unhandled error:", err);
    return res.status(500).json({
      crash: true,
      message: err.message ?? "Unknown error",
      tip: "Nếu lỗi 'quota exceeded': đợi reset lúc 00:00 UTC hoặc thêm API key free khác vào .env"
    });
  }
}
