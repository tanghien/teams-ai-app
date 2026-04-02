export default async function handler(req, res) {
  try {
    let body = req.body;
    if (typeof body === "string") {
      try { body = JSON.parse(body); } catch { body = {}; }
    }
    const question = (body?.question ?? "").trim();

    console.log("✅ Step 0 - question:", question);

    // ─── Test token ───────────────────────────────────────
    const tokenRes = await fetch(
      `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          client_id: process.env.AZURE_CLIENT_ID,
          client_secret: process.env.AZURE_CLIENT_SECRET,
          scope: "https://graph.microsoft.com/.default",
          grant_type: "client_credentials",
        }),
      }
    );
    const tokenData = await tokenRes.json();
    console.log("✅ Step 1 - token status:", tokenRes.status, "has_token:", !!tokenData.access_token);

    if (!tokenData.access_token) {
      return res.json({ error: "Token thất bại", detail: tokenData });
    }

    // ─── Test site ────────────────────────────────────────
    const siteRes = await fetch(
      "https://graph.microsoft.com/v1.0/sites/tbcball.sharepoint.com:/sites/Document",
      { headers: { Authorization: `Bearer ${tokenData.access_token}` } }
    );
    const siteData = await siteRes.json();
    console.log("✅ Step 2 - site status:", siteRes.status, "siteId:", siteData.id);

    // Trả thẳng kết quả để debug
    return res.json({
      tokenOk: !!tokenData.access_token,
      siteStatus: siteRes.status,
      siteData: siteData,
    });

  } catch (err) {
    console.error("❌ Crash:", err);
    return res.json({ crash: true, message: err.message });
  }
}