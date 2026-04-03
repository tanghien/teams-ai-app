import mammoth from "mammoth";
import * as XLSX from "xlsx";

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
      AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET,
      GROQ_API_KEY, GEMINI_API_KEY, OPENROUTER_API_KEY
    } = process.env;

    if (!AZURE_TENANT_ID || !AZURE_CLIENT_ID || !AZURE_CLIENT_SECRET)
      return res.status(500).json({ error: "Thiếu biến môi trường Azure." });
    if (!GROQ_API_KEY && !OPENROUTER_API_KEY && !GEMINI_API_KEY)
      return res.status(500).json({ error: "Cần ít nhất một API key: GROQ, OPENROUTER hoặc GEMINI." });

    // ─── 3. LLM HELPERS ─────────────────────────────────────────────────────
    async function callGroq(prompt, systemPrompt = "", maxTokens = 1024) {
      if (!GROQ_API_KEY) throw new Error("NO_GROQ_KEY");
      const r = await fetch("https://api.groq.com/openai/v1/chat/completions", {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${GROQ_API_KEY}` },
        body: JSON.stringify({
          model: "llama-3.3-70b-versatile", max_tokens: maxTokens, temperature: 0.3,
          messages: [
            ...(systemPrompt ? [{ role: "system", content: systemPrompt }] : []),
            { role: "user", content: prompt }
          ]
        })
      });
      const data = await r.json();
      if (data.error) { const err = new Error(data.error.message || "Groq error"); err.status = r.status; throw err; }
      return data.choices?.[0]?.message?.content?.trim() ?? "";
    }

    async function callOpenRouter(prompt, systemPrompt = "", maxTokens = 1024) {
      if (!OPENROUTER_API_KEY) throw new Error("NO_OPENROUTER_KEY");
      const r = await fetch("https://openrouter.ai/api/v1/chat/completions", {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${OPENROUTER_API_KEY}`, "HTTP-Referer": "https://yourdomain.com", "X-Title": "AI Docs Agent" },
        body: JSON.stringify({
          model: "meta-llama/llama-3-8b-instruct:free", max_tokens: maxTokens, temperature: 0.3,
          messages: [
            ...(systemPrompt ? [{ role: "system", content: systemPrompt }] : []),
            { role: "user", content: prompt }
          ]
        })
      });
      const data = await r.json();
      if (data.error) { const err = new Error(data.error.message || "OpenRouter error"); err.status = r.status; throw err; }
      return data.choices?.[0]?.message?.content?.trim() ?? "";
    }

    async function callGemini(prompt, systemPrompt = "", maxTokens = 1024) {
      if (!GEMINI_API_KEY) throw new Error("NO_GEMINI_KEY");
      const r = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_API_KEY}`, {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          ...(systemPrompt && { system_instruction: { parts: [{ text: systemPrompt }] } }),
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: { maxOutputTokens: maxTokens, temperature: 0.3 }
        })
      });
      const data = await r.json();
      if (data.error) throw new Error(`Gemini error: ${data.error.message}`);
      return data.candidates?.[0]?.content?.parts?.[0]?.text?.trim() ?? "";
    }

    async function callLLM(prompt, systemPrompt = "", maxTokens = 1024) {
      if (GROQ_API_KEY) try { return await callGroq(prompt, systemPrompt, maxTokens); } catch (e) { console.warn(`[Fallback] Groq → OpenRouter: ${e.message}`); }
      if (OPENROUTER_API_KEY) try { return await callOpenRouter(prompt, systemPrompt, maxTokens); } catch (e) { console.warn(`[Fallback] OpenRouter → Gemini: ${e.message}`); }
      return await callGemini(prompt, systemPrompt, maxTokens);
    }

    async function callGeminiWithFile(fileBase64, mimeType, prompt, systemPrompt = "") {
      if (!GEMINI_API_KEY) throw new Error("NO_GEMINI_KEY");
      const r = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_API_KEY}`, {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          ...(systemPrompt && { system_instruction: { parts: [{ text: systemPrompt }] } }),
          contents: [{ parts: [{ inline_data: { mime_type: mimeType, data: fileBase64 } }, { text: prompt }] }],
          generationConfig: { maxOutputTokens: 2048, temperature: 0.3 }
        })
      });
      const data = await r.json();
      if (data.error) throw new Error(`Gemini Vision error: ${data.error.message}`);
      return data.candidates?.[0]?.content?.parts?.[0]?.text?.trim() ?? "";
    }

    // ─── 4. Utilities & File Parsers ─────────────────────────────────────────
    function getMimeType(ext) {
      const map = {
        ".pdf": "application/pdf", ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".xls": "application/vnd.ms-excel",
        ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation", ".txt": "text/plain",
        ".csv": "text/csv", ".md": "text/markdown", ".html": "text/html", ".htm": "text/html",
        ".json": "application/json", ".xml": "application/xml", ".png": "image/png",
        ".jpg": "image/jpeg", ".jpeg": "image/jpeg"
      };
      return map[ext] || null;
    }

    const TEXT_EXTS = [".txt", ".md", ".csv", ".json", ".xml", ".html", ".htm", ".log"];

    async function extractOfficeText(buffer, ext) {
      if (ext === ".docx") {
        const result = await mammoth.extractRawText({ buffer });
        return result.value || "[Không trích xuất được nội dung]";
      }
      if (ext === ".xlsx" || ext === ".xls") {
        const wb = XLSX.read(buffer);
        return wb.SheetNames.slice(0, 5).map(name => XLSX.utils.sheet_to_csv(wb.Sheets[name])).join("\n\n");
      }
      return null; // Fallback cho các định dạng khác
    }

    // ─── 5. SharePoint Auth & File Listing ───────────────────────────────────
    const tokenRes = await fetch(`https://login.microsoftonline.com/${AZURE_TENANT_ID}/oauth2/v2.0/token`, {
      method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({ client_id: AZURE_CLIENT_ID, client_secret: AZURE_CLIENT_SECRET, scope: "https://graph.microsoft.com/.default", grant_type: "client_credentials" })
    });
    const tokenData = await tokenRes.json();
    if (!tokenData.access_token) return res.status(502).json({ error: "Lấy token thất bại", detail: tokenData });
    const accessToken = tokenData.access_token;

    const siteRes = await fetch(`https://graph.microsoft.com/v1.0/sites/tbcball.sharepoint.com:/sites/${process.env.SHAREPOINT_SITE}`, { headers: { Authorization: `Bearer ${accessToken}` } });
    const siteData = await siteRes.json();
    if (!siteData.id) return res.status(502).json({ error: "Không lấy được site ID", detail: siteData });
    const siteId = siteData.id;

    const drivesRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/drives?$select=id,name`, { headers: { Authorization: `Bearer ${accessToken}` } });
    const drivesData = await drivesRes.json();
    const drives = drivesData.value || [];

    if (process.env.DEBUG_FILES === "1") return res.status(200).json({ _debug: true, drives });

    const targetDrive =
      drives.find(d => d.name?.toLowerCase().includes("approved sop")) ||
      drives.find(d => d.name?.toLowerCase().includes("document")) ||
      drives[0];

    if (!targetDrive) return res.status(502).json({ error: "Không tìm thấy Document Library nào.", drives });
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
          allFiles.push({ id: item.id, name: item.name, size: item.size || 0, path: item.parentReference?.path?.replace(`/drives/${driveId}/root:`, "") || "" });
        } else if (item.folder) {
          await fetchChildren(item.id, depth + 1);
        }
      }
    }
    await fetchChildren("root");
    if (allFiles.length === 0) return res.status(200).json({ answer: "Không tìm thấy file nào trong thư viện tài liệu.", _debug: { driveId, driveName: targetDrive.name } });

    // ─── 6. AI Chọn File + Xử Lý + Trả Lời ───────────────────────────────────
    const fileList = allFiles.map((f, i) => `${i + 1}. [${f.path || "/"}] ${f.name} (${Math.round(f.size / 1024)} KB)`).join("\n");

    const selectedIndexStr = await callLLM(
      `Câu hỏi: "${question}"\n\nDanh sách file:\n${fileList}`,
      "Chọn file liên quan nhất đến câu hỏi. Trả lời CHỈ bằng số thứ tự (ví dụ: 5). Nếu không có file liên quan, trả lời: 0.", 50
    );
    const selectedIndex = parseInt(selectedIndexStr.trim(), 10);

    let answer = "";
    let selectedFile = null;
    let usedProvider = "none";

    if (selectedIndex > 0 && selectedIndex <= allFiles.length) {
      selectedFile = allFiles[selectedIndex - 1];
      const ext = selectedFile.name.substring(selectedFile.name.lastIndexOf(".")).toLowerCase();
      const mimeType = getMimeType(ext);

      const dlRes = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${selectedFile.id}/content`, { headers: { Authorization: `Bearer ${accessToken}` } });
      if (!dlRes.ok) {
        answer = `Không tải được file "${selectedFile.name}" (HTTP ${dlRes.status}).`;
      } else {
        const buffer = Buffer.from(await dlRes.arrayBuffer());

        if (TEXT_EXTS.includes(ext)) {
          // Text trực tiếp
          let text = buffer.toString("utf8");
          if (text.length > 15000) text = text.substring(0, 15000) + "\n...[bị cắt bớt]";
          answer = await callLLM(`Nội dung tài liệu:\n\n${text}\n\n---\n\nCâu hỏi: ${question}`, "Bạn là trợ lý AI tra cứu tài liệu nội bộ. Trả lời ngắn gọn, chính xác bằng tiếng Việt.", 1024);
          usedProvider = "text->llm";
        } 
        else if (ext === ".pdf" && GEMINI_API_KEY) {
          // PDF: Dùng Gemini Vision
          const contentLength = parseInt(dlRes.headers.get("content-length") || "0", 10);
          if (contentLength > 18 * 1024 * 1024) {
            answer = `File "${selectedFile.name}" quá lớn (${Math.round(contentLength / 1024 / 1024)} MB). Vui lòng hỏi về file nhỏ hơn 18MB.`;
          } else {
            const base64 = buffer.toString("base64");
            answer = await callGeminiWithFile(base64, mimeType, question, "Bạn là trợ lý AI tra cứu tài liệu. Đọc file PDF và trả lời câu hỏi bằng tiếng Việt.");
            usedProvider = "gemini-vision-pdf";
          }
        } 
        else if ((ext === ".docx" || ext === ".xlsx" || ext === ".xls")) {
          // Office: Parse local -> LLM
          try {
            const text = await extractOfficeText(buffer, ext);
            if (!text || text.length < 50) throw new Error("Nội dung rỗng");
            answer = await callLLM(`Nội dung trích xuất từ file ${ext}:\n\n${text.substring(0, 15000)}\n\n---\n\nCâu hỏi: ${question}`, "Bạn là trợ lý AI tra cứu tài liệu. Dựa vào nội dung trích xuất để trả lời ngắn gọn bằng tiếng Việt.", 1024);
            usedProvider = `local-${ext.slice(1)}->llm`;
          } catch (e) {
            answer = `Không trích xuất được nội dung từ file ${ext}. Vui lòng thử file PDF hoặc liên hệ quản trị.`;
          }
        } 
        else {
          answer = `Định dạng ${ext} chưa được hỗ trợ đọc tự động. Vui lòng chuyển sang PDF hoặc TXT.`;
        }
      }
    } else {
      answer = await callLLM(`Câu hỏi: "${question}"\n\nDanh sách file hiện có:\n${fileList.substring(0, 3000)}`, "Không tìm thấy file phù hợp. Hãy gợi ý người dùng nên tìm trong file nào dựa trên danh sách.", 512);
      usedProvider = "groq/openrouter/gemini";
    }

    return res.status(200).json({
      answer: answer || "Không nhận được câu trả lời.",
      meta: { fileSelected: selectedFile ? `${selectedFile.path}/${selectedFile.name}` : null, totalFiles: allFiles.length, library: targetDrive.name, provider: usedProvider }
    });

  } catch (err) {
    console.error("[ask.js] Unhandled error:", err);
    return res.status(500).json({ crash: true, message: err.message ?? "Unknown error" });
  }
}