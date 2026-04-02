export default async function handler(req, res) {
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
  const accessToken = tokenData.access_token;

  // 🔹 2. Lấy site cụ thể của bạn
  const siteRes = await fetch(
    "https://graph.microsoft.com/v1.0/sites/tbcball.sharepoint.com:/sites/Document",
    {
      headers: { Authorization: `Bearer ${accessToken}` }
    }
  );

  const siteData = await siteRes.json();
  const siteId = siteData.id;

  let content = "Không có dữ liệu";

  // 🔹 3. Lấy danh sách file
  if (siteId) {
    const fileRes = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root/children`,
      {
        headers: { Authorization: `Bearer ${accessToken}` }
      }
    );

    const fileData = await fileRes.json();

    content = fileData.value
      ?.map(f => `- ${f.name}`)
      .join("\n");
  }

  // 🔹 4. Gửi sang Groq
  const aiRes = await fetch("https://api.groq.com/openai/v1/chat/completions", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${process.env.GROQ_API_KEY}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      model: "llama3-70b-8192",
      messages: [
        {
          role: "system",
          content: "Bạn là AI đọc dữ liệu SharePoint"
        },
        {
          role: "user",
          content: `Danh sách file:\n${content}\n\nCâu hỏi:\n${question}`
        }
      ]
    })
  });

  const aiData = await aiRes.json();

  res.json({
    answer: aiData.choices?.[0]?.message?.content || "Không có câu trả lời"
  });
}