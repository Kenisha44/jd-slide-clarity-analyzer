import fs from "fs/promises";
import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import formidable from "formidable";

export const config = {
  api: {
    bodyParser: false
  }
};

function parseForm(req) {
  const form = formidable({
    multiples: false,
    keepExtensions: true,
    maxFileSize: 15 * 1024 * 1024
  });

  return new Promise((resolve, reject) => {
    form.parse(req, (err, fields, files) => {
      if (err) return reject(err);
      resolve({ fields, files });
    });
  });
}

function cleanExtractedText(text) {
  return String(text || "")
    .replace(/\r/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .replace(/[ \t]+/g, " ")
    .trim();
}

function extractTextFromNode(node, acc = []) {
  if (!node) return acc;

  if (typeof node === "string") {
    const trimmed = node.trim();
    if (trimmed) acc.push(trimmed);
    return acc;
  }

  if (Array.isArray(node)) {
    for (const item of node) extractTextFromNode(item, acc);
    return acc;
  }

  if (typeof node === "object") {
    for (const [key, value] of Object.entries(node)) {
      if (key === "a:t" || key.endsWith(":t")) {
        if (Array.isArray(value)) {
          value.forEach((v) => {
            const txt = String(v || "").trim();
            if (txt) acc.push(txt);
          });
        } else {
          const txt = String(value || "").trim();
          if (txt) acc.push(txt);
        }
      } else {
        extractTextFromNode(value, acc);
      }
    }
  }

  return acc;
}

async function extractSlidesFromPptx(filePath) {
  const buffer = await fs.readFile(filePath);
  const zip = await JSZip.loadAsync(buffer);
  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: "@_"
  });

  const slideEntries = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => {
      const aNum = Number(a.match(/slide(\d+)\.xml/)?.[1] || 0);
      const bNum = Number(b.match(/slide(\d+)\.xml/)?.[1] || 0);
      return aNum - bNum;
    });

  const slides = [];

  for (let i = 0; i < slideEntries.length; i++) {
    const xml = await zip.files[slideEntries[i]].async("string");
    const parsed = parser.parse(xml);
    const textRuns = extractTextFromNode(parsed);
    const joined = cleanExtractedText(textRuns.join("\n"));

    slides.push({
      slideNumber: i + 1,
      text: joined,
      preview: joined || "[No text found on this slide]"
    });
  }

  return slides;
}

function buildLocalAnalysis(slideText) {
  const lines = String(slideText)
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);

  let score = 8;
  const issues = [];
  const suggestions = [];

  if (!lines.length) {
    return `Clarity Score: 2/10

Issues Found:
- No readable text was extracted from this slide.

Suggested Improvements:
- Use a text-based slide or confirm the slide contains selectable text.

Suggested Rewrite:
Title: Add a clear takeaway title

- Add 3 to 5 concise bullets`;
  }

  if (lines.length > 6) {
    score -= 1;
    issues.push("This slide may contain too much text.");
    suggestions.push("Reduce to the most important points.");
  }

  if (lines.some((line) => line.length > 100)) {
    score -= 1;
    issues.push("Some lines are too long for quick scanning.");
    suggestions.push("Shorten longer statements.");
  }

  if (!/\d/.test(slideText)) {
    issues.push("The slide may need a specific metric or proof point.");
    suggestions.push("Add at least one number, percentage, or measurable result.");
  }

  if (!issues.length) {
    issues.push("The slide is fairly clear, but can be sharpened.");
    suggestions.push("Use a more insight-led title and tighten wording.");
  }

  if (score < 1) score = 1;

  const title = lines[0] || "Key Takeaway";
  const body = lines.slice(1, 6).map((line) => `- ${line}`).join("\n") || "- Add concise supporting points";

  return `Clarity Score: ${score}/10

Issues Found:
${issues.map((item) => `- ${item}`).join("\n")}

Suggested Improvements:
${suggestions.map((item) => `- ${item}`).join("\n")}

Suggested Rewrite:
Title: ${title}

${body}`;
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  let uploadedFilePath;

  try {
    const { fields, files } = await parseForm(req);

    const deckFile = Array.isArray(files.deck) ? files.deck[0] : files.deck;

    if (!deckFile) {
      return res.status(400).json({
        error: "No PPTX file uploaded. Use field name 'deck'."
      });
    }

    const originalName = deckFile.originalFilename || "uploaded.pptx";
    const mimetype = deckFile.mimetype || "";
    const filepath = deckFile.filepath;
    uploadedFilePath = filepath;

    const looksLikePptx =
      mimetype === "application/vnd.openxmlformats-officedocument.presentationml.presentation" ||
      originalName.toLowerCase().endsWith(".pptx");

    if (!looksLikePptx) {
      return res.status(400).json({
        error: "Only .pptx files are allowed."
      });
    }

    const slideNumberRaw = Array.isArray(fields.slideNumber)
      ? fields.slideNumber[0]
      : fields.slideNumber;

    const slideNumber = Number(slideNumberRaw || 1);

    if (!Number.isInteger(slideNumber) || slideNumber < 1) {
      return res.status(400).json({
        error: "slideNumber must be a positive integer."
      });
    }

    const slides = await extractSlidesFromPptx(filepath);

    if (!slides.length) {
      return res.status(400).json({
        error: "No slides were found in this deck."
      });
    }

    const selected = slides.find((slide) => slide.slideNumber === slideNumber);

    if (!selected) {
      return res.status(400).json({
        error: `Slide ${slideNumber} was not found. Deck has ${slides.length} slide(s).`
      });
    }

    const result = buildLocalAnalysis(selected.text);

    return res.status(200).json({
      fileName: originalName,
      slideCount: slides.length,
      slideNumber: selected.slideNumber,
      slideText: selected.text,
      result
    });
  } catch (error) {
    console.error("analyze-pptx-slide.js error:", error);
    return res.status(500).json({
      error: error.message || "Failed to analyze PPTX slide."
    });
  } finally {
    if (uploadedFilePath) {
      try {
        await fs.unlink(uploadedFilePath);
      } catch {
        // ignore cleanup failure
      }
    }
  }
}