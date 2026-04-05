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
      preview: joined || "[No text found on this slide]"
    });
  }

  return slides;
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  let uploadedFilePath = null;

  try {
    const { files } = await parseForm(req);
    const deckFile = Array.isArray(files.deck) ? files.deck[0] : files.deck;

    if (!deckFile) {
      return res.status(400).json({
        error: "No PPTX file uploaded. Use field name 'deck'."
      });
    }

    const filepath = deckFile.filepath;
    const originalName = deckFile.originalFilename || "uploaded.pptx";
    const mimetype = deckFile.mimetype || "";
    uploadedFilePath = filepath;

    const looksLikePptx =
      mimetype === "application/vnd.openxmlformats-officedocument.presentationml.presentation" ||
      originalName.toLowerCase().endsWith(".pptx");

    if (!looksLikePptx) {
      return res.status(400).json({
        error: "Only .pptx files are allowed."
      });
    }

    const slides = await extractSlidesFromPptx(filepath);

    return res.status(200).json({
      fileName: originalName,
      slideCount: slides.length,
      slides
    });
  } catch (error) {
    console.error("upload-pptx.js error:", error);
    return res.status(500).json({
      error: error.message || "Failed to preview PPTX."
    });
  } finally {
    if (uploadedFilePath) {
      try {
        await fs.unlink(uploadedFilePath);
      } catch {}
    }
  }
}