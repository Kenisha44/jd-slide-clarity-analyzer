import { Resend } from "resend";

const resend = new Resend(process.env.RESEND_API_KEY);

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  try {
    const { title = "", bullets = "", email = "" } = req.body || {};

    const cleanTitle = String(title).trim();
    const cleanBullets = String(bullets).trim();
    const cleanEmail = String(email).trim();

    if (!cleanTitle && !cleanBullets) {
      return res.status(400).json({
        error: "Please provide a slide title or slide text."
      });
    }

    if (!cleanEmail || !cleanEmail.includes("@")) {
      return res.status(400).json({
        error: "A valid email is required."
      });
    }

    if (cleanTitle.length > 80) {
      return res.status(400).json({
        error: "Title must be under 80 characters."
      });
    }

    if (cleanBullets.length > 400) {
      return res.status(400).json({
        error: "Slide text must be under 400 characters."
      });
    }

    const result = buildAnalysis(cleanTitle, cleanBullets);

    try {
      await resend.emails.send({
        from: "Slide Tool <onboarding@resend.dev>",
        to: "contact@johkendesign.com",
        subject: "New Slide Analysis Lead",
        html: `
          <h2>New Slide Analysis Lead</h2>
          <p><strong>Email:</strong> ${escapeHtml(cleanEmail)}</p>
          <p><strong>Title:</strong> ${escapeHtml(cleanTitle || "(No title)")}</p>
          <p><strong>Slide Text:</strong></p>
          <pre>${escapeHtml(cleanBullets || "(No slide text)")}</pre>
          <p><strong>Analysis:</strong></p>
          <pre>${escapeHtml(result)}</pre>
        `
      });
    } catch (emailError) {
      console.error("Resend email error:", emailError);
    }

    return res.status(200).json({ result });
  } catch (error) {
    console.error("analyze.js error:", error);
    return res.status(500).json({
      error: error.message || "Failed to analyze slide."
    });
  }
}

function buildAnalysis(title, bullets) {
  const rawLines = bullets
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => line.replace(/^[-•*]\s*/, ""));

  let score = 8;
  const problems = [];
  const fixes = [];

  if (!title) {
    score -= 2;
    problems.push("There is no clear slide title.");
    fixes.push("Add a title that states the main takeaway.");
  }

  if (rawLines.length > 5) {
    score -= 1;
    problems.push("There are too many points on the slide.");
    fixes.push("Trim it to the 3 to 5 strongest points.");
  }

  if (rawLines.some((line) => line.length > 90)) {
    score -= 1;
    problems.push("Some points are too long.");
    fixes.push("Shorten each point so it is easier to scan.");
  }

  const combinedText = `${title} ${bullets}`;
  if (/[A-Za-z]/.test(combinedText) && !/\d/.test(combinedText)) {
    problems.push("The slide has no numbers or proof points.");
    fixes.push("Add a number, percentage, or result if possible.");
  }

  if (!problems.length) {
    problems.push("The slide is fairly clear, but it could be sharper.");
    fixes.push("Tighten the wording and make the takeaway more direct.");
  }

  if (score < 1) score = 1;

  const betterTitle = rewriteTitle(title, rawLines);
  const betterLines = rewriteBullets(rawLines);

  return `Clarity Score: ${score}/10

What is weakening this slide:
${problems.map((item) => `- ${item}`).join("\n")}

How to improve it:
${fixes.map((item) => `- ${item}`).join("\n")}

Better version:
Title: ${betterTitle}

${betterLines}`;
}

function rewriteTitle(title, lines) {
  const cleanTitle = cleanLine(title);

  if (cleanTitle) {
    return cleanTitle
      .replace(/slide/gi, "")
      .replace(/\s+/g, " ")
      .trim() || "Performance Improved Across Key Channels";
  }

  if (lines.some((line) => /\brevenue\b/i.test(line))) {
    return "Revenue and Channel Performance Improved";
  }

  if (lines.some((line) => /\btraffic\b|\bclick\b|\bconversion\b/i.test(line))) {
    return "Key Marketing Metrics Improved";
  }

  return "Performance Improved Across Key Channels";
}

function rewriteBullets(lines) {
  if (!lines.length) {
    return `- Add 3 short supporting points
- Focus on results, not filler
- Keep each point easy to scan`;
  }

  if (lines.length === 1) {
    return [
      "- Strong performance across key channels",
      "- Revenue trending upward over time",
      "- CTR and conversions improved",
      "- Stable ad spend maintained efficiency",
      "- Continued optimization recommended"
    ].join("\n");
  }

  return lines
    .slice(0, 5)
    .map((line) => `- ${tightenBullet(line)}`)
    .join("\n");
}

function tightenBullet(line) {
  const clean = cleanLine(line);

  return clean
    .replace(/\bincreased by\b/gi, "increased")
    .replace(/\bimproved by\b/gi, "improved")
    .replace(/\bgrew by\b/gi, "grew")
    .replace(/\bsignificantly\b/gi, "")
    .replace(/\bvery\b/gi, "")
    .replace(/\breally\b/gi, "")
    .replace(/\bperformed better than expected\b/gi, "outperformed target")
    .replace(/\bteam recommends\b/gi, "recommend")
    .replace(/\s+/g, " ")
    .trim();
}

function cleanLine(text) {
  return String(text || "").replace(/\s+/g, " ").trim();
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}