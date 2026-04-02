export default async function handler(req, res) {
  try {
    const question = req.body?.question || "";

    // 🔹 1. Lấy token
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

    if (!tokenData.access_token) {
      return res.json({
        error: "Lấy token thất bại",
        detail: tokenData
      });
    }

    const accessToken = tokenData.access_token;

    // 🔹 2. Lấy site
    const siteRes = await fetch(
      "https://graph.microsoft.com/v1.0/sites/tbcball.sharepoint.com:/sites/Document",
      {
        headers: { Authorization: `Bearer ${accessToken}` }
      }
    );

    const siteData = await siteRes.json();

    if (!siteData.id) {
      return res.json({
        error: "Không lấy được site",
        detail: siteData
      });
    }

    const siteId = siteData.id;

    // 🔹 3. Lấy file
    const fileRes = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root/children`,
      {
        headers: { Authorization: `Bearer ${accessToken}` }
      }
    );

    const fileData = await fileRes.json();

    return res.json({
      success: true,
      siteId,
      files: fileData
    });

  } catch (err) {
    return res.json({
      crash: true,
      message: err.message
    });
  }
}