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
        error: "Please provide a slide title or slide text."
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
    if (/performance/i.test(cleanTitle)) return "Performance Improved Across Key Channels";
    if (/marketing/i.test(cleanTitle)) return "Marketing Results Improved Across Key Channels";
    if (/sales/i.test(cleanTitle)) return "Sales Performance Increased This Period";
    if (/growth/i.test(cleanTitle)) return "Growth Accelerated Across Core Metrics";
    return cleanTitle;
  }

  if (lines.some((line) => /\brevenue\b/i.test(line))) {
    return "Revenue and Channel Performance Improved";
  }

  if (lines.some((line) => /\btraffic\b|\bclick\b|\bconversion\b/i.test(line))) {
    return "Key Marketing Metrics Improved";
  }

  return "Main takeaway";
}

function rewriteBullets(lines) {
  if (!lines.length) {
    return `- Add 3 short supporting points
- Focus on results, not filler
- Keep each point easy to scan`;
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