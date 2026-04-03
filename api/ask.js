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
    if (!question) {
      return res.status(400).json({ error: "Thiếu tham số 'question'." });
    }

    // ─── 2. Environment Variables ────────────────────────────────────────────
    const {
      AZURE_TENANT_ID,
      AZURE_CLIENT_ID,
      AZURE_CLIENT_SECRET,
      GROQ_API_KEY,
      GEMINI_API_KEY,
      OPENROUTER_API_KEY
    } = process.env;

    if (!AZURE_TENANT_ID || !AZURE_CLIENT_ID || !AZURE_CLIENT_SECRET) {
      return res.status(500).json({ error: "Thiếu biến môi trường Azure." });
    }
    if (!GROQ_API_KEY && !OPENROUTER_API_KEY && !GEMINI_API_KEY) {
      return res.status(500).json({ error: "Cần ít nhất một API key: GROQ, OPENROUTER hoặc GEMINI." });
    }

    // ─── 3. LLM HELPERS ─────────────────────────────────────────────────────
    async function callGroq(prompt, systemPrompt = "", maxTokens = 1024) {
      if (!GROQ_API_KEY) throw new Error("NO_GROQ_KEY");
      const res = await fetch("https://api.groq.com/openai/v1/chat/completions", {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${GROQ_API_KEY}` },
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
      const data = await res.json();
      if (data.error) {
        const err = new Error(data.error.message || "Groq error");
        err.status = res.status;
        throw err;
      }
      return data.choices?.[0]?.message?.content?.trim() ?? "";
    }

    async function callOpenRouter(prompt, systemPrompt = "", maxTokens = 1024) {
      if (!OPENROUTER_API_KEY) throw new Error("NO_OPENROUTER_KEY");
      const res = await fetch("https://openrouter.ai/api/v1/chat/completions", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${OPENROUTER_API_KEY}`,
          "HTTP-Referer": "https://yourdomain.com",
          "X-Title": "AI Docs Agent"
        },
        body: JSON.stringify({
          model: "meta-llama/llama-3-8b-instruct:free", // ✅ MODEL MIỄN PHÍ
          max_tokens: maxTokens,
          temperature: 0.3,
          messages: [
            ...(systemPrompt ? [{ role: "system", content: systemPrompt }] : []),
            { role: "user", content: prompt }
          ]
        })
      });
      const data = await res.json();
      if (data.error) {
        const err = new Error(data.error.message || "OpenRouter error");
        err.status = res.status;
        throw err;
      }
      return data.choices?.[0]?.message?.content?.trim() ?? "";
    }

    async function callGemini(prompt, systemPrompt = "", maxTokens = 1024) {
      if (!GEMINI_API_KEY) throw new Error("NO_GEMINI_KEY");
      const res = await fetch(
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
      const data = await res.json();
      if (data.error) throw new Error(`Gemini error: ${data.error.message}`);
      return data.candidates?.[0]?.content?.parts?.[0]?.text?.trim() ?? "";
    }

    // 🔹 CHUỖI FALLBACK: Groq → OpenRouter (Free) → Gemini
    async function callLLM(prompt, systemPrompt = "", maxTokens = 1024) {
      if (GROQ_API_KEY) {
        try { return await callGroq(prompt, systemPrompt, maxTokens); }
        catch (e) { console.warn(`[Fallback] Groq → OpenRouter: ${e.message}`); }
      }
      if (OPENROUTER_API_KEY) {
        try { return await callOpenRouter(prompt, systemPrompt, maxTokens); }
        catch (e) { console.warn(`[Fallback] OpenRouter → Gemini: ${e.message}`); }
      }
      return await callGemini(prompt, systemPrompt, maxTokens);
    }

    async function callGeminiWithFile(fileBase64, mimeType, prompt, systemPrompt = "") {
      if (!GEMINI_API_KEY) throw new Error("NO_GEMINI_KEY");
      const res = await fetch(
        `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_API_KEY}`,
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            ...(systemPrompt && { system_instruction: { parts: [{ text: systemPrompt }] } }),
            contents: [{ parts: [{ inline_data: { mime_type: mimeType, data: fileBase64 } }, { text: prompt }] }],
            generationConfig: { maxOutputTokens: 2048, temperature: 0.3 }
          })
        }
      );
      const data = await res.json();
      if (data.error) throw new Error(`Gemini Vision error: ${data.error.message}`);
      return data.candidates?.[0]?.content?.parts?.[0]?.text?.trim() ?? "";
    }

    // ─── 4. Utilities ────────────────────────────────────────────────────────
    function getMimeType(ext) {
      const map = {
        ".pdf": "application/pdf",
        ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".xls": "application/vnd.ms-excel",
        ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        ".txt": "text/plain",
        ".csv": "text/csv",
        ".md": "text/markdown",
        ".html": "text/html",
        ".htm": "text/html",
        ".json": "application/json",
        ".xml": "application/xml",
        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg"
      };
      return map[ext] || null;
    }

    const TEXT_EXTS = [".txt", ".md", ".csv", ".json", ".xml", ".html", ".htm", ".log"];

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
    if (!tokenData.access_token) return res.status(502).json({ error: "Lấy token thất bại", detail: tokenData });
    const accessToken = tokenData.access_token;

    const siteRes = await fetch(
      `https://graph.microsoft.com/v1.0/sites/tbcball.sharepoint.com:/sites/${process.env.SHAREPOINT_SITE}`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );
    const siteData = await siteRes.json();
    if (!siteData.id) return res.status(502).json({ error: "Không lấy được site ID", detail: siteData });
    const siteId = siteData.id;

    const drivesRes = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drives?$select=id,name`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );
    const drivesData = await drivesRes.json();
    const drives = drivesData.value || [];

    if (process.env.DEBUG_FILES === "1") {
      return res.status(200).json({ _debug: true, drives });
    }

    const targetDrive =
      drives.find((d) => d.name?.toLowerCase().includes("approved sop")) ||
      drives.find((d) => d.name?.toLowerCase().includes("document")) ||
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
    if (allFiles.length === 0) {
      return res.status(200).json({ answer: "Không tìm thấy file nào trong thư viện tài liệu.", _debug: { driveId, driveName: targetDrive.name } });
    }

    // ─── 6. AI Chọn File + Xử Lý + Trả Lời ───────────────────────────────────
    const fileList = allFiles
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

    if (selectedIndex > 0 && selectedIndex <= allFiles.length) {
      selectedFile = allFiles[selectedIndex - 1];
      const ext = selectedFile.name.substring(selectedFile.name.lastIndexOf(".")).toLowerCase();
      const mimeType = getMimeType(ext);

      const dlRes = await fetch(
        `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${selectedFile.id}/content`,
        { headers: { Authorization: `Bearer ${accessToken}` } }
      );

      if (!dlRes.ok) {
        answer = `Không tải được file "${selectedFile.name}" (HTTP ${dlRes.status}).`;
      } else if (TEXT_EXTS.includes(ext)) {
        let text = await dlRes.text();
        if (text.length > 12000) text = text.substring(0, 12000) + "\n...[bị cắt bớt]";
        answer = await callLLM(
          `Nội dung tài liệu:\n\n${text}\n\n---\n\nCâu hỏi: ${question}`,
          "Bạn là trợ lý AI tra cứu tài liệu nội bộ. Trả lời ngắn gọn và chính xác bằng tiếng Việt.",
          1024
        );
        usedProvider = "groq/openrouter/gemini";
      } else if (mimeType && GEMINI_API_KEY) {
        const contentLength = parseInt(dlRes.headers.get("content-length") || "0", 10);
        if (contentLength > 18 * 1024 * 1024) {
          answer = `File "${selectedFile.name}" quá lớn (${Math.round(contentLength / 1024 / 1024)} MB). Vui lòng hỏi về file nhỏ hơn 18MB.`;
        } else {
          const base64 = Buffer.from(await dlRes.arrayBuffer()).toString("base64");
          answer = await callGeminiWithFile(
            base64,
            mimeType,
            question,
            "Bạn là trợ lý AI tra cứu tài liệu nội bộ công ty. Đọc nội dung tài liệu và trả lời câu hỏi bằng tiếng Việt, ngắn gọn và chính xác. Nếu không tìm thấy thông tin liên quan, hãy nói rõ."
          );
          usedProvider = "gemini-vision";
        }
      } else if (mimeType && !GEMINI_API_KEY) {
        answer = `File "${selectedFile.name}" là định dạng ${ext} cần Gemini để đọc, nhưng GEMINI_API_KEY chưa được cấu hình.`;
      } else {
        answer = `Định dạng ${ext} chưa được hỗ trợ.`;
      }
    } else {
      answer = await callLLM(
        `Câu hỏi: "${question}"\n\nDanh sách file hiện có:\n${fileList.substring(0, 3000)}`,
        "Bạn là trợ lý AI tra cứu tài liệu nội bộ. Không tìm thấy file phù hợp. Hãy gợi ý người dùng nên tìm trong file nào dựa trên danh sách.",
        512
      );
      usedProvider = "groq/openrouter/gemini";
    }

    return res.status(200).json({
      answer: answer || "Không nhận được câu trả lời.",
      meta: {
        fileSelected: selectedFile ? `${selectedFile.path}/${selectedFile.name}` : null,
        totalFiles: allFiles.length,
        library: targetDrive.name,
        provider: usedProvider
      }
    });

  } catch (err) {
    console.error("[ask.js] Unhandled error:", err);
    return res.status(500).json({ crash: true, message: err.message ?? "Unknown error" });
  }
}