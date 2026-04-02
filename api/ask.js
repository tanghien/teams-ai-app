export default async function handler(req, res) {
  // ─── Guard ───────────────────────────────────────────────────────────────────
  if (!req || !res) return;
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed. Use POST." });
  }

  try {
    // ─── Parse body ──────────────────────────────────────────────────────────────
    let body = req.body;
    if (typeof body === "string") {
      try { body = JSON.parse(body); } catch { body = {}; }
    }
    if (!body || typeof body !== "object") body = {};

    const question = (body.question ?? "").trim();
    if (!question) {
      return res.status(400).json({ error: "Thiếu tham số 'question'." });
    }

    // ─── Kiểm tra env ────────────────────────────────────────────────────────────
    const {
      AZURE_TENANT_ID,
      AZURE_CLIENT_ID,
      AZURE_CLIENT_SECRET,
      GROQ_API_KEY,
    } = process.env;

    if (!AZURE_TENANT_ID || !AZURE_CLIENT_ID || !AZURE_CLIENT_SECRET) {
      return res.status(500).json({ error: "Thiếu biến môi trường Azure." });
    }
    if (!GROQ_API_KEY) {
      return res.status(500).json({ error: "Thiếu biến môi trường GROQ_API_KEY." });
    }

    // ─── 1. Lấy Access Token ─────────────────────────────────────────────────────
    const tokenRes = await fetch(
      `https://login.microsoftonline.com/${AZURE_TENANT_ID}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          client_id: AZURE_CLIENT_ID,
          client_secret: AZURE_CLIENT_SECRET,
          scope: "https://graph.microsoft.com/.default",
          grant_type: "client_credentials",
        }),
      }
    );
    const tokenData = await tokenRes.json();
    if (!tokenData.access_token) {
      return res.status(502).json({ error: "Lấy token thất bại", detail: tokenData });
    }
    const accessToken = tokenData.access_token;

    // ─── 2. Lấy Site ID ──────────────────────────────────────────────────────────
    const siteRes = await fetch(
      "https://graph.microsoft.com/v1.0/sites/tbcball.sharepoint.com:/sites/Document",
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );
    const siteData = await siteRes.json();
    if (!siteData.id) {
      return res.status(502).json({ error: "Không lấy được site ID", detail: siteData });
    }
    const siteId = siteData.id;

    // ─── 3. Lấy danh sách file ───────────────────────────────────────────────────
    const fileRes = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root/children?$select=id,name,size,file,lastModifiedDateTime`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );
    const fileData = await fileRes.json();
    const allFiles = (fileData.value || []).filter((f) => f.file);

    if (allFiles.length === 0) {
      return res.status(200).json({ answer: "Không tìm thấy file nào trong SharePoint." });
    }

    // ─── 4. Dùng Groq chọn file liên quan nhất ───────────────────────────────────
    const fileList = allFiles
      .map((f, i) => `${i + 1}. ${f.name} (${Math.round((f.size || 0) / 1024)} KB)`)
      .join("\n");

    const selectRes = await fetch("https://api.groq.com/openai/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${GROQ_API_KEY}`,
      },
      body: JSON.stringify({
        model: "llama-3.3-70b-versatile",
        max_tokens: 100,
        temperature: 0,
        messages: [
          {
            role: "system",
            content:
              "Bạn là assistant giúp chọn file liên quan. Trả lời CHỈ bằng số thứ tự của file liên quan nhất (ví dụ: 3). Nếu không có file nào liên quan, trả lời: 0.",
          },
          {
            role: "user",
            content: `Câu hỏi: "${question}"\n\nDanh sách file:\n${fileList}\n\nFile nào liên quan nhất?`,
          },
        ],
      }),
    });
    const selectData = await selectRes.json();
    const selectedIndexStr = (selectData.choices?.[0]?.message?.content ?? "0").trim();
    const selectedIndex = parseInt(selectedIndexStr, 10);

    // ─── 5. Download nội dung file được chọn ─────────────────────────────────────
    let fileContent = "";
    let selectedFileName = null;

    if (selectedIndex > 0 && selectedIndex <= allFiles.length) {
      const selectedFile = allFiles[selectedIndex - 1];
      selectedFileName = selectedFile.name;

      const ext = selectedFile.name
        .substring(selectedFile.name.lastIndexOf("."))
        .toLowerCase();

      const textExtensions = [".txt", ".md", ".csv", ".json", ".xml", ".html", ".htm", ".log"];

      if (textExtensions.includes(ext)) {
        // Download trực tiếp cho file text
        const dlRes = await fetch(
          `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${selectedFile.id}/content`,
          { headers: { Authorization: `Bearer ${accessToken}` } }
        );
        if (dlRes.ok) {
          fileContent = await dlRes.text();
          // Giới hạn 8000 ký tự tránh vượt context
          if (fileContent.length > 8000) {
            fileContent = fileContent.substring(0, 8000) + "\n...[nội dung bị cắt bớt]";
          }
        }
      } else {
        // File Office/PDF: không đọc binary, báo cho Groq biết
        fileContent = `[File: ${selectedFile.name}]\nĐịnh dạng ${ext} — không thể đọc nội dung trực tiếp qua API. Hãy trả lời dựa trên tên file và ngữ cảnh câu hỏi.`;
      }
    }

    // ─── 6. Groq trả lời câu hỏi ─────────────────────────────────────────────────
    const contextText =
      fileContent ||
      `Không tìm thấy file liên quan đến câu hỏi.\n\nDanh sách file hiện có trong hệ thống:\n${fileList}`;

    const answerRes = await fetch("https://api.groq.com/openai/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${GROQ_API_KEY}`,
      },
      body: JSON.stringify({
        model: "llama-3.3-70b-versatile",
        max_tokens: 1024,
        temperature: 0.3,
        messages: [
          {
            role: "system",
            content:
              "Bạn là trợ lý AI giúp tra cứu tài liệu nội bộ công ty. Trả lời bằng tiếng Việt, ngắn gọn và chính xác dựa trên nội dung tài liệu được cung cấp. Nếu không đủ thông tin, hãy nói rõ và gợi ý người dùng tìm file phù hợp hơn.",
          },
          {
            role: "user",
            content: `Nội dung tài liệu:\n\n${contextText}\n\n---\n\nCâu hỏi: ${question}`,
          },
        ],
      }),
    });

    const answerData = await answerRes.json();
    const answer =
      answerData.choices?.[0]?.message?.content ??
      "Không nhận được câu trả lời từ Groq.";

    // ─── 7. Trả kết quả ──────────────────────────────────────────────────────────
    return res.status(200).json({
      answer,
      meta: {
        fileSelected: selectedFileName,
        totalFiles: allFiles.length,
      },
    });

  } catch (err) {
    console.error("[ask.js] Unhandled error:", err);
    return res.status(500).json({ crash: true, message: err.message ?? "Unknown error" });
  }
}