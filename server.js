import express from "express";
import cors from "cors";
import dotenv from "dotenv";
import OpenAI from "openai";
import path from "path";
import fs from "fs/promises";
import multer from "multer";
import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import { fileURLToPath } from "url";
import rateLimit from "express-rate-limit";

const app = express();

const limiter = rateLimit({
  windowMs: 15 * 60 * 1000, // 15 min
  max: 50, // limit requests
});


app.use("/upload-pptx", limiter);
app.use("/analyze-pptx-slide", limiter);
app.use("/analyze", limiter);
console.log("USING NEW JSZIP PARSER SERVER");

dotenv.config();


app.use(cors());
app.use(express.json());

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
});

const upload = multer({
  dest: path.join(__dirname, "uploads"),
  limits: {
    fileSize: 15 * 1024 * 1024,
  },
  fileFilter: (req, file, cb) => {
    const isPptx =
      file.mimetype ===
        "application/vnd.openxmlformats-officedocument.presentationml.presentation" ||
      file.originalname.toLowerCase().endsWith(".pptx");

    if (!isPptx) {
      return cb(new Error("Only .pptx files are allowed."));
    }

    cb(null, true);
  },
});

function cleanExtractedText(text) {
  return String(text || "")
    .replace(/\r/g, "\n")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function ensureArray(value) {
  if (!value) return [];
  return Array.isArray(value) ? value : [value];
}

function collectText(node, parts = []) {
  if (!node) return parts;

  // Only extract actual text nodes
  if (typeof node === "string") {
    const text = node.trim();

    // FILTER OUT JUNK
    if (
      text.length > 2 &&
      !text.startsWith("http") &&
      !text.includes("schemas.microsoft") &&
      !text.match(/^\{.*\}$/) &&
      !text.match(/^[0-9.]+$/) &&
      !["UTF-8", "en-US", "body", "tx1"].includes(text)
    ) {
      parts.push(text);
    }

    return parts;
  }

  if (Array.isArray(node)) {
    node.forEach(item => collectText(item, parts));
    return parts;
  }

  if (typeof node === "object") {
    for (const [key, value] of Object.entries(node)) {
      // ONLY extract text nodes from PPT structure
      if (key === "a:t") {
        collectText(value, parts);
      } else {
        collectText(value, parts);
      }
    }
  }

  return parts;
}

async function extractSlidesFromPptx(filePath) {
  const buffer = await fs.readFile(filePath);
  const zip = await JSZip.loadAsync(buffer);

  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: "@_",
    trimValues: true,
  });

  const slideFiles = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => {
      const aNum = Number(a.match(/slide(\d+)\.xml$/)?.[1] || 0);
      const bNum = Number(b.match(/slide(\d+)\.xml$/)?.[1] || 0);
      return aNum - bNum;
    });

  const slides = [];

  for (let i = 0; i < slideFiles.length; i++) {
    const slidePath = slideFiles[i];
    const xml = await zip.files[slidePath].async("string");
    const parsed = parser.parse(xml);

    const texts = collectText(parsed);
    const joined = cleanExtractedText([...new Set(texts)].join("\n"));

    slides.push({
      slideNumber: i + 1,
      text: joined || "[No extractable text found on this slide]",
    });
  }

  return slides;
}

app.get("/ping", (req, res) => {
  res.json({ ok: true, message: "pong" });
});

app.use((req, res, next) => {
  console.log(`${req.method} ${req.url}`);
  next();
});

app.post("/upload-pptx", upload.single("deck"), async (req, res) => {
  let filePath = null;

  try {
    if (!req.file) {
      return res.status(400).json({ error: "No file uploaded." });
    }

    filePath = req.file.path;

    const slides = await extractSlidesFromPptx(filePath);
    const allText = slides.map((s) => `Slide ${s.slideNumber}\n${s.text}`).join("\n\n");

    return res.json({
      fileName: req.file.originalname,
      slideCount: slides.length,
      extractedTextPreview: allText.slice(0, 1500),
      slides: slides.map((slide) => ({
        slideNumber: slide.slideNumber,
        preview: slide.text.slice(0, 400),
      })),
    });
  } catch (error) {
    console.error("UPLOAD ERROR:", error);
    return res.status(500).json({
      error: error.message || "Failed to parse PPTX.",
    });
  } finally {
    if (filePath) {
      try {
        await fs.unlink(filePath);
      } catch {}
    }
  }
});

app.post("/analyze-pptx-slide", upload.single("deck"), async (req, res) => {
  let filePath = null;

  try {
    if (!req.file) {
      return res.status(400).json({ error: "No file uploaded." });
    }

    filePath = req.file.path;
    const requestedSlide = Number(req.body.slideNumber || 1);

    const slides = await extractSlidesFromPptx(filePath);

    if (!slides.length) {
      return res.status(400).json({ error: "No slides found in PPTX." });
    }

    const selected =
      slides.find((slide) => slide.slideNumber === requestedSlide) || slides[0];

    const prompt = `
You are an expert presentation strategist and slide copy consultant.

Analyze this single slide and help improve its clarity, structure, and executive-level communication.

Slide number:
${selected.slideNumber}

Slide content:
${selected.text}

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

    const response = await openai.chat.completions.create({
      model: "gpt-4o-mini",
      messages: [{ role: "user", content: prompt }],
    });

    return res.json({
      slideNumber: selected.slideNumber,
      slideText: selected.text,
      result: response.choices[0].message.content,
    });
  } catch (error) {
    console.error("ANALYZE PPTX ERROR:", error);
    return res.status(500).json({
      error: error.message || "Failed to analyze uploaded slide.",
    });
  } finally {
    if (filePath) {
      try {
        await fs.unlink(filePath);
      } catch {}
    }
  }
});

app.post("/analyze", async (req, res) => {
  try {
    const { title, bullets } = req.body;

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

    const response = await openai.chat.completions.create({
      model: "gpt-4o-mini",
      messages: [{ role: "user", content: prompt }],
    });

    return res.json({
      result: response.choices[0].message.content,
    });
  } catch (error) {
    console.error("ANALYZE TEXT ERROR:", error);
    return res.status(500).json({
      error: error.message || "Something went wrong.",
    });
  }
});

app.use(express.static(__dirname));

app.use((err, req, res, next) => {
  console.error("SERVER ERROR:", err);

  if (res.headersSent) {
    return next(err);
  }

  return res.status(500).json({
    error: err.message || "Server error.",
  });
});

app.listen(3000, () => {
  console.log("Server running on http://localhost:3000");
});