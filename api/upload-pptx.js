import multer from "multer";
import fs from "fs/promises";
import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";

export const config = {
  api: {
    bodyParser: false
  }
};

const upload = multer({
  dest: "/tmp",
  limits: {
    fileSize: 15 * 1024 * 1024
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
  }
});

function runMiddleware(req, res, fn) {
  return new Promise((resolve, reject) => {
    fn(req, res, (result) => {
      if (result instanceof Error) return reject(result);
      return resolve(result);
    });
  });
}

function cleanExtractedText(text) {
  return String(text || "")
    .replace(/\r/g, "\n")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function collectText(node, parts = []) {
  if (!node) return parts;

  if (typeof node === "string") {
    const text = node.trim();
    if (
      text.length > 1 &&
      !text.startsWith("http") &&
      !text.includes("schemas.microsoft") &&
      !text.match(/^\{.*\}$/) &&
      !text.match(/^[0-9.]+$/) &&
      !["UTF-8", "en-US", "body", "tx1", "ppt"].includes(text)
    ) {
      parts.push(text);
    }
    return parts;
  }

  if (Array.isArray(node)) {
    node.forEach((item) => collectText(item, parts));
    return parts;
  }

  if (typeof node === "object") {
    for (const [key, value] of Object.entries(node)) {
      if (key === "a:t" || key.endsWith(":t")) {
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
    trimValues: true
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
      text: joined || "[No extractable text found on this slide]"
    });
  }

  return slides;
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  let filePath = null;

  try {
    await runMiddleware(req, res, upload.single("deck"));

    if (!req.file) {
      return res.status(400).json({ error: "No file uploaded." });
    }

    filePath = req.file.path;

    const slides = await extractSlidesFromPptx(filePath);

    return res.status(200).json({
      fileName: req.file.originalname,
      slideCount: slides.length,
      slides: slides.map((slide) => ({
        slideNumber: slide.slideNumber,
        preview: slide.text.slice(0, 400)
      }))
    });
  } catch (err) {
    console.error("UPLOAD PPTX API ERROR:", err);
    return res.status(500).json({
      error: err?.message || "Failed to parse PPTX."
    });
  } finally {
    if (filePath) {
      try {
        await fs.unlink(filePath);
      } catch {}
    }
  }
}