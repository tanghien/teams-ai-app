export default async function handler(req, res) {
  const question = req.body.question;

  const response = await fetch("https://api.groq.com/openai/v1/chat/completions", {
    method: "POST",
    headers: {
      "Authorization": `Bearer ${process.env.GROQ_API_KEY}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      model: "llama3-70b-8192",
      messages: [
        {
          role: "system",
          content: "Bạn là AI hỗ trợ tài liệu nội bộ SharePoint."
        },
        {
          role: "user",
          content: question
        }
      ]
    })
  });

  const data = await response.json();
  res.status(200).json({
    answer: data.choices?.[0]?.message?.content || "Không có câu trả lời"
  });
}