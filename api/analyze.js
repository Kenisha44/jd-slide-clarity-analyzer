export default async function handler(req, res) {
    if (req.method !== "POST") {
      return res.status(405).json({ error: "Method not allowed" });
    }
  
    try {
      const { title = "", bullets = "" } = req.body || {};
  
      const cleanTitle = String(title).trim();
      const cleanBullets = String(bullets).trim();
  
      if (!cleanTitle && !cleanBullets) {
        return res.status(400).json({
          error: "Please provide a slide title or bullets."
        });
      }
  
      const slideText = [cleanTitle, cleanBullets].filter(Boolean).join("\n\n");
  
      // Later, replace this with your OpenAI/Vercel AI call.
      const result = buildLocalAnalysis(cleanTitle, cleanBullets, slideText);
  
      return res.status(200).json({ result });
    } catch (error) {
      console.error("analyze.js error:", error);
      return res.status(500).json({
        error: error.message || "Failed to analyze typed slide."
      });
    }
  }
  
  function buildLocalAnalysis(title, bullets, slideText) {
    const bulletLines = bullets
      .split("\n")
      .map((line) => line.trim())
      .filter(Boolean);
  
    let score = 8;
    const issues = [];
    const suggestions = [];
  
    if (!title) {
      score -= 2;
      issues.push("Missing a clear slide title.");
      suggestions.push("Add a short takeaway title that tells the audience the point of the slide.");
    }
  
    if (bulletLines.length > 5) {
      score -= 1;
      issues.push("There are many bullets, which may feel dense.");
      suggestions.push("Trim to the 3 to 5 strongest points.");
    }
  
    if (bulletLines.some((line) => line.length > 90)) {
      score -= 1;
      issues.push("Some bullets are too long and may be hard to scan.");
      suggestions.push("Shorten long bullets into crisp, one-line statements.");
    }
  
    if (/[A-Za-z]/.test(slideText) && !/\d/.test(slideText)) {
      issues.push("The slide may benefit from a metric, result, or concrete proof point.");
      suggestions.push("Add one or two numbers to make the message more credible.");
    }
  
    if (!issues.length) {
      issues.push("The slide is reasonably clear, but could still be sharpened.");
      suggestions.push("Make the title more insight-led and tighten wording for faster scanning.");
    }
  
    if (score < 1) score = 1;
    if (score > 10) score = 10;
  
    const rewrittenTitle = title ? rewriteTitle(title) : "Key Takeaway";
  
    const rewrittenBullets =
      bulletLines.length > 0
        ? bulletLines.map((line) => `- ${rewriteBullet(line)}`).join("\n")
        : "- Add concise supporting points here";
  
    return `Clarity Score: ${score}/10
  
  Issues Found:
  ${issues.map((item) => `- ${item}`).join("\n")}
  
  Suggested Improvements:
  ${suggestions.map((item) => `- ${item}`).join("\n")}
  
  Suggested Rewrite:
  Title: ${rewrittenTitle}
  
  ${rewrittenBullets}`;
  }
  
  function rewriteTitle(title) {
    return title.replace(/\s+/g, " ").trim();
  }
  
  function rewriteBullet(line) {
    return line
      .replace(/^[-•*]\s*/, "")
      .replace(/\s+/g, " ")
      .trim();
  }