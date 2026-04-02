export default async function handler(req, res) {
  // ─── Guard: kiểm tra req / res hợp lệ ───────────────────────────────────────
  if (!req || !res) {
    return res?.json({ error: "Handler called without req/res" });
  }

  // ─── Chỉ cho phép POST ────────────────────────────────────────────────────────
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed. Use POST." });
  }

  try {
    // ─── Parse body an toàn (hỗ trợ cả parsed object lẫn raw string) ────────────
    let body = req.body;

    if (typeof body === "string") {
      try {
        body = JSON.parse(body);
      } catch {
        body = {};
      }
    }

    if (!body || typeof body !== "object") {
      body = {};
    }

    const question = (body.question ?? "").trim();

    if (!question) {
      return res.status(400).json({
        error: "Thiếu tham số 'question' trong request body.",
      });
    }

    // ─── Kiểm tra biến môi trường ────────────────────────────────────────────────
    const { AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET } = process.env;

    if (!AZURE_TENANT_ID || !AZURE_CLIENT_ID || !AZURE_CLIENT_SECRET) {
      return res.status(500).json({
        error: "Thiếu biến môi trường Azure. Kiểm tra AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET.",
      });
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

    if (!tokenRes.ok) {
      const raw = await tokenRes.text();
      return res.status(502).json({
        error: "Lấy token thất bại (HTTP error)",
        status: tokenRes.status,
        detail: raw,
      });
    }

    const tokenData = await tokenRes.json();

    if (!tokenData.access_token) {
      return res.status(502).json({
        error: "Lấy token thất bại (không có access_token)",
        detail: tokenData,
      });
    }

    const accessToken = tokenData.access_token;

    // ─── 2. Lấy SharePoint Site ──────────────────────────────────────────────────
    const siteRes = await fetch(
      "https://graph.microsoft.com/v1.0/sites/tbcball.sharepoint.com:/sites/Document",
      {
        headers: { Authorization: `Bearer ${accessToken}` },
      }
    );

    if (!siteRes.ok) {
      const raw = await siteRes.text();
      return res.status(502).json({
        error: "Không lấy được SharePoint site (HTTP error)",
        status: siteRes.status,
        detail: raw,
      });
    }

    const siteData = await siteRes.json();

    if (!siteData.id) {
      return res.status(502).json({
        error: "Không lấy được site ID",
        detail: siteData,
      });
    }

    const siteId = siteData.id;

    // ─── 3. Lấy danh sách file ───────────────────────────────────────────────────
    const fileRes = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root/children`,
      {
        headers: { Authorization: `Bearer ${accessToken}` },
      }
    );

    if (!fileRes.ok) {
      const raw = await fileRes.text();
      return res.status(502).json({
        error: "Không lấy được danh sách file (HTTP error)",
        status: fileRes.status,
        detail: raw,
      });
    }

    const fileData = await fileRes.json();

    // ─── Trả về kết quả thành công ───────────────────────────────────────────────
    return res.status(200).json({
      success: true,
      question,
      siteId,
      files: fileData,
    });

  } catch (err) {
    console.error("[ask.js] Unhandled error:", err);
    return res.status(500).json({
      crash: true,
      message: err.message ?? "Unknown error",
    });
  }
}