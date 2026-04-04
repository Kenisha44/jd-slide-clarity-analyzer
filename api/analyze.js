import OpenAI from "openai";

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  try {
    if (!process.env.OPENAI_API_KEY) {
      return res.status(500).json({ error: "Missing OPENAI_API_KEY in Vercel environment variables." });
    }

    const client = new OpenAI({
      apiKey: process.env.OPENAI_API_KEY,
    });

    const { title, bullets } = req.body || {};

    if (!title && !bullets) {
      return res.status(400).json({ error: "Please enter slide content." });
    }

    const prompt = `
You are an expert presentation strategist and slide copy consultant.

Analyze this single slide and help improve its clarity, structure, and executive-level communication.

Slide title:
${title || ""}

Slide bullets/content:
${bullets || ""}

Return exactly in this format:

Clarity Score: [1-10]

Top Issues:
- [issue 1]
- [issue 2]
- [issue 3]

Suggested Rewrite:
Title: [better title]
Bullets:
- [bullet 1]
- [bullet 2]
- [bullet 3]

Design Notes:
- [how to visually improve the slide]
- [what to highlight]
- [what to cut or simplify]

Keep the answer concise, practical, and client-facing.
`;

    const response = await client.chat.completions.create({
      model: "gpt-4o-mini",
      messages: [{ role: "user", content: prompt }],
    });

    return res.status(200).json({
      result: response.choices[0].message.content,
    });
  } catch (err) {
    console.error("API /api/analyze error:", err);
    return res.status(500).json({
      error: err?.message || "Server error",
    });
  }
}