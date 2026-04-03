export default async function handler(req, res) {
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

    // ─── Env ─────────────────────────────────────────────────────────────────────
    const { AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, GROQ_API_KEY } = process.env;
    if (!AZURE_TENANT_ID || !AZURE_CLIENT_ID || !AZURE_CLIENT_SECRET)
      return res.status(500).json({ error: "Thiếu biến môi trường Azure." });
    if (!GROQ_API_KEY)
      return res.status(500).json({ error: "Thiếu biến môi trường GROQ_API_KEY." });

    // ─── 1. Access Token ─────────────────────────────────────────────────────────
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
    if (!tokenData.access_token)
      return res.status(502).json({ error: "Lấy token thất bại", detail: tokenData });
    const accessToken = tokenData.access_token;

    // ─── 2. Site ID ──────────────────────────────────────────────────────────────
    const siteRes = await fetch(
      `https://graph.microsoft.com/v1.0/sites/tbcball.sharepoint.com:/sites/${process.env.SHAREPOINT_SITE}`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );
    const siteData = await siteRes.json();
    if (!siteData.id)
      return res.status(502).json({ error: "Không lấy được site ID", detail: siteData });
    const siteId = siteData.id;

    // ─── 3. Tìm Document Library ─────────────────────────────────────────────────
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

    if (!targetDrive)
      return res.status(502).json({ error: "Không tìm thấy Document Library nào.", drives });

    const driveId = targetDrive.id;

    // ─── 4. Đệ quy lấy tất cả file (tối đa 200) ─────────────────────────────────
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
            path: item.parentReference?.path?.replace("/drives/" + driveId + "/root:", "") || "",
          });
        } else if (item.folder) {
          await fetchChildren(item.id, depth + 1);
        }
      }
    }

    await fetchChildren("root");

    if (allFiles.length === 0) {
      return res.status(200).json({
        answer: "Không tìm thấy file nào trong thư viện tài liệu.",
        _debug: { driveId, driveName: targetDrive.name },
      });
    }

    // ─── 5. Groq chọn file liên quan nhất ────────────────────────────────────────
    const fileList = allFiles
      .map((f, i) => `${i + 1}. [${f.path || "/"}] ${f.name} (${Math.round(f.size / 1024)} KB)`)
      .join("\n");

    const selectRes = await fetch("https://api.groq.com/openai/v1/chat/completions", {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: `Bearer ${GROQ_API_KEY}` },
      body: JSON.stringify({
        model: "llama-3.3-70b-versatile",
        max_tokens: 50,
        temperature: 0,
        messages: [
          {
            role: "system",
            content: "Chọn file liên quan nhất đến câu hỏi. Trả lời CHỈ bằng số thứ tự (ví dụ: 5). Nếu không có file liên quan, trả lời: 0.",
          },
          {
            role: "user",
            content: `Câu hỏi: "${question}"\n\nDanh sách file:\n${fileList}`,
          },
        ],
      }),
    });
    const selectData = await selectRes.json();
    const selectedIndex = parseInt((selectData.choices?.[0]?.message?.content ?? "0").trim(), 10);

    // ─── 6. Download và extract nội dung file ────────────────────────────────────
    let fileContent = "";
    let selectedFile = null;

    if (selectedIndex > 0 && selectedIndex <= allFiles.length) {
      selectedFile = allFiles[selectedIndex - 1];
      const ext = selectedFile.name.substring(selectedFile.name.lastIndexOf(".")).toLowerCase();

      // Download binary từ SharePoint
      const dlRes = await fetch(
        `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${selectedFile.id}/content`,
        { headers: { Authorization: `Bearer ${accessToken}` } }
      );

      if (!dlRes.ok) {
        fileContent = `[Không tải được file: ${selectedFile.name} — HTTP ${dlRes.status}]`;

      } else if ([".txt", ".md", ".csv", ".json", ".xml", ".html", ".htm", ".log"].includes(ext)) {
        // ── Text thuần ───────────────────────────────────────────────────────────
        fileContent = await dlRes.text();

      } else if (ext === ".pdf") {
        // ── PDF: extract text bằng pdf-parse ────────────────────────────────────
        try {
          const pdfParse = (await import("pdf-parse/lib/pdf-parse.js")).default;
          const buffer = Buffer.from(await dlRes.arrayBuffer());
          const data = await pdfParse(buffer, { max: 15 });
          fileContent = data.text?.trim() || `[PDF: ${selectedFile.name} — không có text, có thể là PDF scan]`;
        } catch (e) {
          fileContent = `[PDF lỗi: ${e.message}]`;
        }

      } else if (ext === ".docx") {
        // ── DOCX: extract text bằng mammoth ─────────────────────────────────────
        try {
          const mammoth = (await import("mammoth")).default;
          const buffer = Buffer.from(await dlRes.arrayBuffer());
          const result = await mammoth.extractRawText({ buffer });
          fileContent = result.value?.trim() || `[DOCX: ${selectedFile.name} — không có text]`;
        } catch (e) {
          fileContent = `[DOCX lỗi: ${e.message}]`;
        }

      } else if (ext === ".xlsx" || ext === ".xls") {
        // ── XLSX: đọc bằng SheetJS ───────────────────────────────────────────────
        try {
          const XLSX = (await import("xlsx")).default;
          const buffer = Buffer.from(await dlRes.arrayBuffer());
          const workbook = XLSX.read(buffer, { type: "buffer" });
          const lines = [];
          for (const sheetName of workbook.SheetNames) {
            const sheet = workbook.Sheets[sheetName];
            const csv = XLSX.utils.sheet_to_csv(sheet);
            lines.push(`=== Sheet: ${sheetName} ===\n${csv}`);
          }
          fileContent = lines.join("\n\n").trim() || `[XLSX: ${selectedFile.name} — không có dữ liệu]`;
        } catch (e) {
          fileContent = `[XLSX lỗi: ${e.message}]`;
        }

      } else if (ext === ".pptx") {
        // ── PPTX: extract text bằng officeparser ────────────────────────────────
        try {
          const officeParser = (await import("officeparser")).default;
          const buffer = Buffer.from(await dlRes.arrayBuffer());
          fileContent = await officeParser.parseOfficeAsync(buffer) || `[PPTX: ${selectedFile.name} — không có text]`;
        } catch (e) {
          fileContent = `[PPTX lỗi: ${e.message}]`;
        }

      } else {
        fileContent = `[Định dạng ${ext} chưa hỗ trợ: ${selectedFile.name}]`;
      }

      // Giới hạn độ dài context gửi cho Groq
      if (fileContent.length > 10000)
        fileContent = fileContent.substring(0, 10000) + "\n...[nội dung bị cắt bớt]";
    }

    // ─── 7. Groq trả lời ─────────────────────────────────────────────────────────
    const contextText = fileContent ||
      `Không tìm thấy file phù hợp.\n\nCác file hiện có (${allFiles.length} file):\n${fileList.substring(0, 3000)}`;

    const answerRes = await fetch("https://api.groq.com/openai/v1/chat/completions", {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: `Bearer ${GROQ_API_KEY}` },
      body: JSON.stringify({
        model: "llama-3.3-70b-versatile",
        max_tokens: 1024,
        temperature: 0.3,
        messages: [
          {
            role: "system",
            content: "Bạn là trợ lý AI tra cứu tài liệu nội bộ l. Trả lời bằng ngắn gọn và chính xác dựa trên nội dung tài liệu. Nếu không đủ thông tin, hãy nói rõ.",
          },
          {
            role: "user",
            content: `Nội dung tài liệu:\n\n${contextText}\n\n---\n\nCâu hỏi: ${question}`,
          },
        ],
      }),
    });

    const answerData = await answerRes.json();
    const answer = answerData.choices?.[0]?.message?.content ?? "Không nhận được câu trả lời từ Groq.";

    // ─── 8. Response ─────────────────────────────────────────────────────────────
    return res.status(200).json({
      answer,
      meta: {
        fileSelected: selectedFile ? `${selectedFile.path}/${selectedFile.name}` : null,
        totalFiles: allFiles.length,
        library: targetDrive.name,
      },
    });

  } catch (err) {
    console.error("[ask.js] Unhandled error:", err);
    return res.status(500).json({ crash: true, message: err.message ?? "Unknown error" });
  }
}