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
      GEMINI_API_KEY,
      OPENROUTER_API_KEY,
      HF_TOKEN,
      NVIDIA_API_KEY
    } = process.env;

    if (!AZURE_TENANT_ID || !AZURE_CLIENT_ID || !AZURE_CLIENT_SECRET)
      return res.status(500).json({ error: "Thiếu biến môi trường Azure." });
    
    const hasAnyLLM = GROQ_API_KEY || OPENROUTER_API_KEY || HF_TOKEN || NVIDIA_API_KEY || GEMINI_API_KEY;
    if (!hasAnyLLM)
      return res.status(500).json({ error: "Cần ít nhất một API key: GROQ, OPENROUTER, HF, NVIDIA hoặc GEMINI." });

    // ─── 3. LLM PROVIDERS ────────────────────────────────────────────────────
    
    // 🔹 1. Groq — Nhanh nhất, ưu tiên số 1
    async function callGroq(prompt, systemPrompt = "", maxTokens = 1024) {
      if (!GROQ_API_KEY) throw new Error("NO_GROQ_KEY");
      console.log("[LLM:1/5] → Calling Groq...");
      
      const r = await fetch("https://api.groq.com/openai/v1/chat/completions", {
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

      const remaining = r.headers.get("x-ratelimit-remaining-requests");
      if (remaining) console.log(`[Groq] Remaining: ${remaining} requests`);

      const data = await r.json();
      if (data.error) {
        const err = new Error(data.error.message || "Groq error");
        err.status = r.status;
        if (r.status === 429) console.warn("[Groq] RATE LIMIT (429)");
        throw err;
      }
      console.log("[Groq] ✓ Success");
      return data.choices?.[0]?.message?.content?.trim() ?? "";
    }

    // 🔹 2. OpenRouter — Free models, đa dạng
    async function callOpenRouter(prompt, systemPrompt = "", maxTokens = 1024) {
      if (!OPENROUTER_API_KEY) throw new Error("NO_OPENROUTER_KEY");
      console.log("[LLM:2/5] → Calling OpenRouter (Free)...");
      
      const r = await fetch("https://openrouter.ai/api/v1/chat/completions", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${OPENROUTER_API_KEY}`,
          "HTTP-Referer": "https://yourdomain.com",
          "X-Title": "AI Docs Agent"
        },
        body: JSON.stringify({
          model: "meta-llama/llama-3-8b-instruct:free", // ✅ Model miễn phí
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
        const err = new Error(data.error.message || "OpenRouter error");
        err.status = r.status;
        console.warn(`[OpenRouter] Error: ${err.message}`);
        throw err;
      }
      console.log("[OpenRouter] ✓ Success");
      return data.choices?.[0]?.message?.content?.trim() ?? "";
    }

    // 🔹 3. Hugging Face — Model chuyên biệt
    async function callHuggingFace(prompt, systemPrompt = "", maxTokens = 1024) {
      if (!HF_TOKEN) throw new Error("NO_HF_TOKEN");
      console.log("[LLM:3/5] → Calling Hugging Face...");
      
      const r = await fetch("https://router.huggingface.co/v1/chat/completions", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${HF_TOKEN}`
        },
        body: JSON.stringify({
          model: "meta-llama/Llama-3.1-8B-Instruct:fastest",
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
        const err = new Error(data.error.message || "Hugging Face error");
        err.status = r.status;
        console.warn(`[HuggingFace] Error: ${err.message}`);
        throw err;
      }
      console.log("[HuggingFace] ✓ Success");
      return data.choices?.[0]?.message?.content?.trim() ?? "";
    }

    // 🔹 4. NVIDIA NIM — Backup model Trung Quốc
    async function callNvidia(prompt, systemPrompt = "", maxTokens = 1024) {
      if (!NVIDIA_API_KEY) throw new Error("NO_NVIDIA_KEY");
      console.log("[LLM:4/5] → Calling NVIDIA NIM...");
      
      const r = await fetch("https://integrate.api.nvidia.com/v1/chat/completions", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${NVIDIA_API_KEY}`
        },
        body: JSON.stringify({
          model: "thudm/glm-4-7b", // ✅ Model free của NVIDIA
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

    // 🔹 5. Gemini — Last resort, hỗ trợ Vision
    async function callGemini(prompt, systemPrompt = "", maxTokens = 1024) {
      if (!GEMINI_API_KEY) throw new Error("NO_GEMINI_KEY");
      console.log("[LLM:5/5] → Calling Gemini (last resort)...");
      
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
        console.error(`[Gemini] Error: ${data.error.message}`);
        throw new Error(`Gemini error: ${data.error.message}`);
      }
      console.log("[Gemini] ✓ Success");
      return data.candidates?.[0]?.content?.parts?.[0]?.text?.trim() ?? "";
    }

    // 🔹 Gemini Vision — Đọc file PDF/ảnh
    async function callGeminiWithFile(fileBase64, mimeType, prompt, systemPrompt = "") {
      if (!GEMINI_API_KEY) throw new Error("NO_GEMINI_KEY");
      console.log(`[Vision] → Calling Gemini with file (${mimeType})`);
      
      const r = await fetch(
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
      const data = await r.json();
      if (data.error) throw new Error(`Gemini Vision error: ${data.error.message}`);
      return data.candidates?.[0]?.content?.parts?.[0]?.text?.trim() ?? "";
    }

    // ─── FALLBACK CHAIN: 1→2→3→4→5 ─────────────────────────────────────────
    async function callLLM(prompt, systemPrompt = "", maxTokens = 1024) {
      // 1️⃣ Groq
      if (GROQ_API_KEY) {
        try { return await callGroq(prompt, systemPrompt, maxTokens); }
        catch (e) { console.warn(`[Fallback 1→2] Groq failed: ${e.message}`); }
      }
      // 2️⃣ OpenRouter
      if (OPENROUTER_API_KEY) {
        try { return await callOpenRouter(prompt, systemPrompt, maxTokens); }
        catch (e) { console.warn(`[Fallback 2→3] OpenRouter failed: ${e.message}`); }
      }
      // 3️⃣ Hugging Face
      if (HF_TOKEN) {
        try { return await callHuggingFace(prompt, systemPrompt, maxTokens); }
        catch (e) { console.warn(`[Fallback 3→4] HuggingFace failed: ${e.message}`); }
      }
      // 4️⃣ NVIDIA NIM
      if (NVIDIA_API_KEY) {
        try { return await callNvidia(prompt, systemPrompt, maxTokens); }
        catch (e) { console.warn(`[Fallback 4→5] NVIDIA failed: ${e.message}`); }
      }
      // 5️⃣ Gemini (cuối cùng)
      console.log("[Fallback] Using Gemini as last resort");
      return await callGemini(prompt, systemPrompt, maxTokens);
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
    if (allFiles.length === 0) return res.status(200).json({ answer: "Không tìm thấy file nào trong thư viện tài liệu.", _debug: { driveId, driveName: targetDrive.name } });

    // ─── 6. AI Chọn File + Xử Lý + Trả Lời ───────────────────────────────────
    const fileList = allFiles.map((f, i) => `${i + 1}. [${f.path || "/"}] ${f.name} (${Math.round(f.size / 1024)} KB)`).join("\n");

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
        // Text file: đọc thẳng → gửi LLM chain
        let text = await dlRes.text();
        if (text.length > 15000) text = text.substring(0, 15000) + "\n...[bị cắt bớt]";
        answer = await callLLM(
          `Nội dung tài liệu:\n\n${text}\n\n---\n\nCâu hỏi: ${question}`,
          "Bạn là trợ lý AI tra cứu tài liệu nội bộ. Trả lời ngắn gọn và chính xác bằng tiếng Việt.",
          1024
        );
        usedProvider = "text->llm-chain";
      } else if (ext === ".pdf" && GEMINI_API_KEY) {
        // PDF: Dùng Gemini Vision
        const contentLength = parseInt(dlRes.headers.get("content-length") || "0", 10);
        if (contentLength > 18 * 1024 * 1024) {
          answer = `File "${selectedFile.name}" quá lớn (${Math.round(contentLength / 1024 / 1024)} MB).`;
        } else {
          const base64 = Buffer.from(await dlRes.arrayBuffer()).toString("base64");
          answer = await callGeminiWithFile(base64, mimeType, question, "Đọc file PDF và trả lời câu hỏi bằng tiếng Việt.");
          usedProvider = "gemini-vision-pdf";
        }
      } else if ((ext === ".docx" || ext === ".xlsx" || ext === ".xls") && GEMINI_API_KEY) {
        // Office: Gemini Vision hỗ trợ một số định dạng
        const contentLength = parseInt(dlRes.headers.get("content-length") || "0", 10);
        if (contentLength > 18 * 1024 * 1024) {
          answer = `File "${selectedFile.name}" quá lớn. Vui lòng chuyển sang PDF.`;
        } else {
          const base64 = Buffer.from(await dlRes.arrayBuffer()).toString("base64");
          answer = await callGeminiWithFile(base64, mimeType, question, "Đọc nội dung file và trả lời bằng tiếng Việt.");
          usedProvider = "gemini-vision-office";
        }
      } else {
        answer = `Định dạng ${ext} chưa được hỗ trợ. Vui lòng chuyển sang PDF hoặc TXT.`;
      }
    } else {
      // Không tìm thấy file phù hợp → gợi ý bằng LLM chain
      answer = await callLLM(
        `Câu hỏi: "${question}"\n\nDanh sách file hiện có:\n${fileList.substring(0, 3000)}`,
        "Không tìm thấy file phù hợp. Hãy gợi ý người dùng nên tìm trong file nào dựa trên danh sách.",
        512
      );
      usedProvider = "fallback-llm-chain";
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