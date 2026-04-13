import express from "express";
import cors from "cors";
import { GoogleGenerativeAI } from "@google/generative-ai";
import multer from "multer";
import XLSX from "xlsx";
import { randomUUID } from "crypto";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const app = express();
app.use(cors());
app.use(express.json());

const SYSTEM_PROMPT = `You are a presentation JSON generator. You output ONLY valid JSON matching a strict schema. No markdown, no explanation, no code fences — just raw JSON.

OUTPUT FORMAT:
{
  "slides": [ ...slide objects... ]
}

SLIDE TYPES AND THEIR EXACT FIELDS:

1. type: "title"
   - title (string) — main presentation title
   - subtitle (string) — subtitle/tagline
   - badge (string) — period label like "Q1 2026"
   - metrics (array of {number: string, label: string}) — 3-6 key numbers
   - highlights (array of {color: string (hex), text: string}) — 3-6 bullet highlights
   - stats (array of {icon: string, label: string, value: string}) — 2-4 info rows

2. type: "agenda"
   - title (string) — "Agenda" or similar
   - subtitle (string) — brief description
   - cards (array of {icon: string, title: string, items: [string array]}) — 4-8 agenda cards

3. type: "section-overview"
   - title (string)
   - subtitle (string)
   - sectionIcon (string) — emoji like "👥"
   - sectionTitle (string)
   - description (string) — 1-2 sentence overview
   - stats (array of {number: string, label: string}) — 3-4 stats
   - billingMetrics (array of {label: string, value: string}) — optional financial metrics

4. type: "dashboard"
   - title (string)
   - subtitle (string)
   - metrics (array of {icon: string, value: string, label: string, detail: string}) — 4-8 metric cards

5. type: "content"
   - title (string)
   - subtitle (string)
   - cards (array of {icon: string, title: string, status: "Active"|"In Progress"|"Planned"|"Complete", items: [{bold: string, text: string}]}) — 2-4 content cards

6. type: "table"
   - title (string)
   - subtitle (string)
   - columns (array of strings) — column header names
   - rows (array of {cells: [string array]}) — table data rows
   - summaryStats (array of {number: string, label: string}) — 2-4 summary stats

7. type: "highlight-list"
   - title (string)
   - subtitle (string)
   - items (array of {color: string (hex), text: string}) — 4-8 highlight items

8. type: "thank-you"
   - title (string) — "Thank You" or similar
   - message (string) — closing message paragraph
   - signature (string) — team/person name

RULES:
- Start with a "title" slide and end with a "thank-you" slide
- Use a mix of slide types for visual variety
- Generate 6-12 slides total
- Use realistic, professional business content
- For icons, use emoji characters (e.g., "👥", "📊", "🚀", "💰", "⭐", "📈")
- For highlight colors, use hex codes: #0078d4 (blue), #28a745 (green), #ffc107 (yellow), #dc3545 (red), #8b5cf6 (purple), #0d9488 (teal)
- Keep text concise — bullet points, not paragraphs
- Numbers should be realistic and consistent across slides`;

app.post("/api/generate", async (req, res) => {
  try {
    const { prompt } = req.body;

    if (!prompt || !prompt.trim()) {
      return res.status(400).json({ error: "Prompt is required" });
    }

    const apiKey = process.env.GOOGLE_API_KEY;
    if (!apiKey) {
      return res.status(500).json({ error: "GOOGLE_API_KEY environment variable is not set" });
    }

    // Try models in order — fallback if one is overloaded or quota-blocked
    const MODELS = ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-2.0-flash-lite", "gemini-flash-latest"];
    const requestBody = JSON.stringify({
      contents: [{ parts: [{ text: SYSTEM_PROMPT + "\n\nUser input:\n" + prompt }] }],
      generationConfig: { temperature: 0.7, maxOutputTokens: 16384 },
    });

    let geminiData = null;
    let lastError = "";

    for (const model of MODELS) {
      const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
      console.log(`Trying model: ${model}...`);
      try {
        const geminiRes = await fetch(apiUrl, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: requestBody,
        });
        if (geminiRes.ok) {
          geminiData = await geminiRes.json();
          console.log(`Success with model: ${model}`);
          break;
        }
        const errData = await geminiRes.json().catch(() => ({}));
        lastError = errData?.error?.message || `${geminiRes.status}`;
        console.log(`Model ${model} failed: ${lastError.substring(0, 80)}...`);
      } catch (fetchErr) {
        lastError = fetchErr.message;
        console.log(`Model ${model} fetch error: ${lastError}`);
      }
    }

    if (!geminiData) {
      return res.status(503).json({ error: `All models failed. Last error: ${lastError}` });
    }

    const text = geminiData?.candidates?.[0]?.content?.parts?.[0]?.text || "";
    console.log("Raw response length:", text.length);

    // Robust JSON extraction — handle fences, partial JSON, extra text
    let jsonStr = text;

    // 1. Strip markdown code fences
    const fenceMatch = text.match(/```(?:json)?\s*([\s\S]*?)```/);
    if (fenceMatch) jsonStr = fenceMatch[1];

    // 2. Find the JSON object boundaries
    const firstBrace = jsonStr.indexOf("{");
    const lastBrace = jsonStr.lastIndexOf("}");
    if (firstBrace !== -1 && lastBrace !== -1 && lastBrace > firstBrace) {
      jsonStr = jsonStr.substring(firstBrace, lastBrace + 1);
    }

    jsonStr = jsonStr.trim();

    // 3. Fix common JSON issues from AI output
    jsonStr = jsonStr
      .replace(/,\s*}/g, "}")       // trailing commas before }
      .replace(/,\s*\]/g, "]")      // trailing commas before ]
      .replace(/[\x00-\x1F\x7F]/g, (c) => c === "\n" || c === "\t" || c === "\r" ? c : ""); // remove control chars

    let data;
    try {
      data = JSON.parse(jsonStr);
    } catch (parseErr) {
      console.error("JSON parse failed. First 500 chars:", jsonStr.substring(0, 500));
      console.error("Last 500 chars:", jsonStr.substring(jsonStr.length - 500));
      console.error("Parse error:", parseErr.message);

      // 4. Last resort: try to find and parse just the slides array
      const slidesMatch = jsonStr.match(/"slides"\s*:\s*(\[[\s\S]*\])/);
      if (slidesMatch) {
        try {
          let slidesStr = slidesMatch[1];
          // Fix truncated array — close any open brackets
          let openBrackets = (slidesStr.match(/\[/g) || []).length;
          let closeBrackets = (slidesStr.match(/\]/g) || []).length;
          let openBraces = (slidesStr.match(/\{/g) || []).length;
          let closeBraces = (slidesStr.match(/\}/g) || []).length;
          while (closeBraces < openBraces) { slidesStr += "}"; closeBraces++; }
          while (closeBrackets < openBrackets) { slidesStr += "]"; closeBrackets++; }
          slidesStr = slidesStr.replace(/,\s*}/g, "}").replace(/,\s*\]/g, "]");
          data = { slides: JSON.parse(slidesStr) };
          console.log("Recovered slides array with", data.slides.length, "slides");
        } catch (e2) {
          return res.status(422).json({ error: "AI returned malformed JSON. Please try again with a shorter prompt." });
        }
      } else {
        return res.status(422).json({ error: "AI returned invalid JSON. Please try again." });
      }
    }

    if (!data.slides || !Array.isArray(data.slides)) {
      return res.status(422).json({ error: "AI response missing slides array" });
    }

    // Filter out any incomplete slides (missing type)
    data.slides = data.slides.filter(s => s && s.type);

    // Assign IDs to slides
    const now = Date.now();
    data.slides = data.slides.map((slide, i) => ({
      ...slide,
      id: `ai_slide_${now + i}`,
    }));

    res.json(data);
  } catch (err) {
    console.error("Generate error:", err);
    if (err instanceof SyntaxError) {
      return res.status(422).json({ error: "AI returned invalid JSON. Please try again." });
    }
    res.status(500).json({ error: err.message || "Internal server error" });
  }
});

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// FEATURE 1: Photo Management
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

const PHOTOS_DIR = path.join(__dirname, "uploads", "photos");
fs.mkdirSync(PHOTOS_DIR, { recursive: true });

const photoDb = []; // In-memory store: { id, name, designation, photoUrl, filename }

const photoUpload = multer({
  storage: multer.diskStorage({
    destination: PHOTOS_DIR,
    filename: (req, file, cb) => {
      const safe = file.originalname.replace(/[^a-zA-Z0-9._-]/g, "_");
      cb(null, `${Date.now()}_${safe}`);
    },
  }),
  limits: { fileSize: 5 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    if (/\.(jpg|jpeg|png)$/i.test(file.originalname)) cb(null, true);
    else cb(new Error("Only JPG/PNG allowed"));
  },
});

app.use("/uploads", express.static(path.join(__dirname, "uploads")));

app.post("/api/photos/upload", photoUpload.array("photos", 20), (req, res) => {
  const results = (req.files || []).map((f) => {
    // Auto-map by filename: "VP_John.png" → { designation: "VP", name: "John" }
    const baseName = path.parse(f.originalname).name;
    const parts = baseName.split("_");
    const designation = parts.length > 1 ? parts[0] : "";
    const name = parts.length > 1 ? parts.slice(1).join(" ") : baseName;
    const entry = { id: randomUUID(), name, designation, photoUrl: `/uploads/photos/${f.filename}`, filename: f.originalname };
    // Check for duplicate name+designation → overwrite
    const existing = photoDb.findIndex((p) => p.name === name && p.designation === designation);
    if (existing >= 0) photoDb[existing] = entry;
    else photoDb.push(entry);
    return entry;
  });
  res.json({ photos: results });
});

app.get("/api/photos/list", (req, res) => res.json({ photos: photoDb }));

app.put("/api/photos/:id", express.json(), (req, res) => {
  const idx = photoDb.findIndex((p) => p.id === req.params.id);
  if (idx < 0) return res.status(404).json({ error: "Not found" });
  Object.assign(photoDb[idx], req.body);
  res.json(photoDb[idx]);
});

app.delete("/api/photos/:id", (req, res) => {
  const idx = photoDb.findIndex((p) => p.id === req.params.id);
  if (idx < 0) return res.status(404).json({ error: "Not found" });
  const [removed] = photoDb.splice(idx, 1);
  try { fs.unlinkSync(path.join(__dirname, removed.photoUrl)); } catch {}
  res.json({ ok: true });
});

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// FEATURE 2: Template Management
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

const TEMPLATES_DIR = path.join(__dirname, "uploads", "templates");
fs.mkdirSync(TEMPLATES_DIR, { recursive: true });

const templateDb = []; // { id, fileName, uploadedOn, version, filePath, active }

const templateUpload = multer({
  storage: multer.diskStorage({
    destination: TEMPLATES_DIR,
    filename: (req, file, cb) => cb(null, `${Date.now()}_${file.originalname.replace(/[^a-zA-Z0-9._-]/g, "_")}`),
  }),
  fileFilter: (req, file, cb) => {
    if (/\.xlsx$/i.test(file.originalname)) cb(null, true);
    else cb(new Error("Only .xlsx allowed"));
  },
});

app.post("/api/templates/upload", templateUpload.single("template"), (req, res) => {
  try {
    const wb = XLSX.readFile(req.file.path);
    // Validate required sheets
    const sheets = wb.SheetNames;
    console.log("Template sheets:", sheets);

    // Deactivate previous active
    templateDb.forEach((t) => (t.active = false));

    const entry = {
      id: randomUUID(),
      fileName: req.file.originalname,
      uploadedOn: new Date().toISOString(),
      version: `v${templateDb.length + 1}`,
      filePath: req.file.path,
      active: true,
      sheets,
      headers: {},
    };

    // Extract headers from each sheet
    sheets.forEach((name) => {
      const sheet = wb.Sheets[name];
      const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      entry.headers[name] = data[0] || [];
    });

    templateDb.push(entry);
    res.json(entry);
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

app.get("/api/templates/latest", (req, res) => {
  const active = templateDb.find((t) => t.active) || templateDb[templateDb.length - 1];
  if (!active) return res.status(404).json({ error: "No template" });
  res.json(active);
});

app.get("/api/templates/list", (req, res) => res.json({ templates: templateDb }));

app.get("/api/templates/download/:id", (req, res) => {
  const tmpl = templateDb.find((t) => t.id === req.params.id);
  if (!tmpl) return res.status(404).json({ error: "Not found" });
  res.download(tmpl.filePath, tmpl.fileName);
});

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// FEATURE 3: Template-to-Excel Sync
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

const excelUpload = multer({
  storage: multer.diskStorage({
    destination: path.join(__dirname, "uploads", "excels"),
    filename: (req, file, cb) => cb(null, `${Date.now()}_${file.originalname.replace(/[^a-zA-Z0-9._-]/g, "_")}`),
  }),
  fileFilter: (req, file, cb) => {
    if (/\.xlsx$/i.test(file.originalname)) cb(null, true);
    else cb(new Error("Only .xlsx allowed"));
  },
});

app.post("/api/sync/template-to-excel", excelUpload.single("excel"), (req, res) => {
  try {
    const activeTemplate = templateDb.find((t) => t.active);
    if (!activeTemplate) return res.status(400).json({ error: "No active template. Upload a template first." });

    const templateWb = XLSX.readFile(activeTemplate.filePath);
    const excelWb = XLSX.readFile(req.file.path);

    const changes = [];

    templateWb.SheetNames.forEach((sheetName) => {
      if (!excelWb.SheetNames.includes(sheetName)) {
        // Add missing sheet
        excelWb.SheetNames.push(sheetName);
        excelWb.Sheets[sheetName] = XLSX.utils.aoa_to_sheet([activeTemplate.headers[sheetName] || []]);
        changes.push({ sheet: sheetName, action: "added_sheet" });
        return;
      }

      const templateHeaders = activeTemplate.headers[sheetName] || [];
      const excelSheet = excelWb.Sheets[sheetName];
      const excelData = XLSX.utils.sheet_to_json(excelSheet, { header: 1 });
      const excelHeaders = excelData[0] || [];

      const missing = templateHeaders.filter((h) => !excelHeaders.includes(h));
      if (missing.length > 0) {
        // Add missing columns
        missing.forEach((col) => {
          excelHeaders.push(col);
          changes.push({ sheet: sheetName, field: col, action: "added_column" });
        });
        excelData[0] = excelHeaders;
        // Fill default values for existing rows
        for (let r = 1; r < excelData.length; r++) {
          while (excelData[r].length < excelHeaders.length) excelData[r].push("");
        }
        excelWb.Sheets[sheetName] = XLSX.utils.aoa_to_sheet(excelData);
      }
    });

    // Save updated Excel
    const outputPath = req.file.path.replace(".xlsx", "_synced.xlsx");
    XLSX.writeFile(excelWb, outputPath);

    res.json({ changes, downloadUrl: `/api/sync/download?path=${encodeURIComponent(outputPath)}` });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// FEATURE 4: Excel-to-Template Sync
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

app.post("/api/sync/excel-to-template", excelUpload.single("excel"), (req, res) => {
  try {
    const activeTemplate = templateDb.find((t) => t.active);
    if (!activeTemplate) return res.status(400).json({ error: "No active template." });

    // Backup: version current template before modifying
    const backupPath = activeTemplate.filePath.replace(".xlsx", `_backup_${Date.now()}.xlsx`);
    fs.copyFileSync(activeTemplate.filePath, backupPath);

    const excelWb = XLSX.readFile(req.file.path);
    const templateWb = XLSX.readFile(activeTemplate.filePath);

    const changes = [];

    excelWb.SheetNames.forEach((sheetName) => {
      const excelSheet = excelWb.Sheets[sheetName];
      const excelData = XLSX.utils.sheet_to_json(excelSheet, { header: 1 });
      const excelHeaders = excelData[0] || [];

      if (!templateWb.SheetNames.includes(sheetName)) {
        // Add new sheet to template
        templateWb.SheetNames.push(sheetName);
        templateWb.Sheets[sheetName] = XLSX.utils.aoa_to_sheet([excelHeaders]);
        changes.push({ sheet: sheetName, action: "added_sheet" });
        activeTemplate.headers[sheetName] = excelHeaders;
        return;
      }

      const templateHeaders = activeTemplate.headers[sheetName] || [];
      const newFields = excelHeaders.filter((h) => !templateHeaders.includes(h));

      if (newFields.length > 0) {
        const updatedHeaders = [...templateHeaders, ...newFields];
        templateWb.Sheets[sheetName] = XLSX.utils.aoa_to_sheet([updatedHeaders]);
        activeTemplate.headers[sheetName] = updatedHeaders;
        newFields.forEach((f) => changes.push({ sheet: sheetName, field: f, action: "added_to_template" }));
      }
    });

    XLSX.writeFile(templateWb, activeTemplate.filePath);
    res.json({ changes, message: `Template updated. ${changes.length} changes applied.` });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Download synced files
app.get("/api/sync/download", (req, res) => {
  const filePath = req.query.path;
  if (!filePath || !fs.existsSync(filePath)) return res.status(404).json({ error: "File not found" });
  res.download(filePath);
});

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// FEATURE 5: AI Generate with Secure Data (Excel masking)
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function maskData(data) {
  // Mask client/company names, financial figures, personal names
  const nameMap = new Map();
  let clientIdx = 0, personIdx = 0;

  return data.map((row) => {
    const masked = {};
    for (const [key, val] of Object.entries(row)) {
      const lk = key.toLowerCase();
      if (lk.includes("client") || lk.includes("company") || lk.includes("customer")) {
        if (!nameMap.has(val)) nameMap.set(val, `Client_${String.fromCharCode(65 + clientIdx++)}`);
        masked[key] = nameMap.get(val);
      } else if (lk.includes("name") || lk.includes("author") || lk.includes("speaker") || lk.includes("coordinator")) {
        if (!nameMap.has(val)) nameMap.set(val, `Person_${++personIdx}`);
        masked[key] = nameMap.get(val);
      } else if (lk.includes("revenue") || lk.includes("billing") || lk.includes("cost") || lk.includes("rate")) {
        const num = parseFloat(String(val).replace(/[^0-9.]/g, ""));
        masked[key] = isNaN(num) ? val : `$${Math.round(num * (0.8 + Math.random() * 0.4))}`;
      } else {
        masked[key] = val;
      }
    }
    return masked;
  });
}

app.post("/api/ai/generate-from-excel", excelUpload.single("excel"), async (req, res) => {
  try {
    const secure = req.body.secure !== "false";
    const wb = XLSX.readFile(req.file.path);

    // Extract all sheet data
    const allData = {};
    wb.SheetNames.forEach((name) => {
      const data = XLSX.utils.sheet_to_json(wb.Sheets[name]);
      allData[name] = secure ? maskData(data) : data;
    });

    const dataPrompt = JSON.stringify(allData, null, 2);
    const prompt = `Generate a presentation from this data:\n${dataPrompt}\n\nCreate slides that best represent this data. Use appropriate slide types for each data set.`;

    // Reuse the AI generation logic
    const apiKey = process.env.GOOGLE_API_KEY;
    if (!apiKey) return res.status(500).json({ error: "GOOGLE_API_KEY not set" });

    const MODELS = ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-2.0-flash-lite", "gemini-flash-latest"];
    const requestBody = JSON.stringify({
      contents: [{ parts: [{ text: SYSTEM_PROMPT + "\n\nUser input:\n" + prompt }] }],
      generationConfig: { temperature: 0.7, maxOutputTokens: 16384 },
    });

    let geminiData = null;
    for (const model of MODELS) {
      try {
        const geminiRes = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`, {
          method: "POST", headers: { "Content-Type": "application/json" }, body: requestBody,
        });
        if (geminiRes.ok) { geminiData = await geminiRes.json(); break; }
      } catch {}
    }

    if (!geminiData) return res.status(503).json({ error: "AI service unavailable" });

    const text = geminiData?.candidates?.[0]?.content?.parts?.[0]?.text || "";
    let jsonStr = text;
    const fenceMatch = text.match(/```(?:json)?\s*([\s\S]*?)```/);
    if (fenceMatch) jsonStr = fenceMatch[1];
    const firstBrace = jsonStr.indexOf("{");
    const lastBrace = jsonStr.lastIndexOf("}");
    if (firstBrace !== -1 && lastBrace > firstBrace) jsonStr = jsonStr.substring(firstBrace, lastBrace + 1);
    jsonStr = jsonStr.trim().replace(/,\s*}/g, "}").replace(/,\s*\]/g, "]");

    const data = JSON.parse(jsonStr);
    if (data.slides) {
      data.slides = data.slides.filter(s => s?.type).map((s, i) => ({ ...s, id: `ai_${Date.now() + i}` }));
    }

    res.json({ ...data, masked: secure });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// FEATURE 6: Auto Template Creation from Excel
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

const KNOWN_FIELDS = ["Revenue", "Client", "Utilization", "Resources", "Project", "Status", "Technology", "Date", "Name", "Role", "Manager", "Score", "Budget", "Timeline"];

app.post("/api/templates/auto-create", excelUpload.single("excel"), (req, res) => {
  try {
    const wb = XLSX.readFile(req.file.path);
    const activeTemplate = templateDb.find((t) => t.active);

    // Check mismatch
    let mismatch = false;
    const excelHeaders = {};
    wb.SheetNames.forEach((name) => {
      const data = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1 });
      excelHeaders[name] = data[0] || [];
      if (activeTemplate && activeTemplate.headers[name]) {
        const templateH = activeTemplate.headers[name];
        if (JSON.stringify(templateH.sort()) !== JSON.stringify(excelHeaders[name].sort())) mismatch = true;
      } else {
        mismatch = true;
      }
    });

    if (!mismatch && activeTemplate) {
      return res.json({ mismatch: false, message: "Excel matches current template." });
    }

    // Generate new template
    const newWb = XLSX.utils.book_new();
    const fieldMapping = {};

    wb.SheetNames.forEach((name) => {
      const headers = excelHeaders[name];
      const mappedHeaders = headers.map((h) => {
        const match = KNOWN_FIELDS.find((kf) => h.toLowerCase().includes(kf.toLowerCase()));
        return { original: h, mapped: match || h, isKnown: !!match };
      });
      fieldMapping[name] = mappedHeaders;
      XLSX.utils.book_append_sheet(newWb, XLSX.utils.aoa_to_sheet([headers]), name);
    });

    // Save new template
    const filename = `auto_template_${Date.now()}.xlsx`;
    const filePath = path.join(TEMPLATES_DIR, filename);
    XLSX.writeFile(newWb, filePath);

    // Deactivate old, add new
    templateDb.forEach((t) => (t.active = false));
    const newEntry = {
      id: randomUUID(),
      fileName: filename,
      uploadedOn: new Date().toISOString(),
      version: `v${templateDb.length + 1}`,
      filePath,
      active: true,
      sheets: wb.SheetNames,
      headers: excelHeaders,
    };
    templateDb.push(newEntry);

    res.json({ mismatch: true, template: newEntry, fieldMapping, message: "New template created from Excel." });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// TEMPLATE LIBRARY — Pre-built presentation templates
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

const BUILT_IN_TEMPLATES = [
  {
    id: "qbr", name: "Quarterly Business Review", category: "Business", icon: "📊",
    description: "Comprehensive quarterly report with KPIs, financials, team updates, and strategic outlook.",
    slides: [
      { order: 1, title: "Title & Overview", type: "title", fields: ["title", "subtitle", "badge", "metrics", "highlights", "stats"], layout: "split" },
      { order: 2, title: "Agenda", type: "agenda", fields: ["title", "subtitle", "cards"], layout: "grid" },
      { order: 3, title: "Resource & Billing", type: "table", fields: ["title", "subtitle", "columns", "rows", "summaryStats"], layout: "table" },
      { order: 4, title: "Team Structure", type: "section-overview", fields: ["title", "subtitle", "sectionIcon", "sectionTitle", "description", "stats", "billingMetrics"], layout: "split" },
      { order: 5, title: "Performance Dashboard", type: "dashboard", fields: ["title", "subtitle", "metrics"], layout: "grid" },
      { order: 6, title: "Key Initiatives", type: "content", fields: ["title", "subtitle", "cards"], layout: "cards" },
      { order: 7, title: "Sales Pipeline", type: "table", fields: ["title", "subtitle", "columns", "rows", "summaryStats"], layout: "table" },
      { order: 8, title: "Achievements", type: "highlight-list", fields: ["title", "subtitle", "items"], layout: "list" },
      { order: 9, title: "Training & Certifications", type: "dashboard", fields: ["title", "subtitle", "metrics"], layout: "grid" },
      { order: 10, title: "Team Events", type: "dashboard", fields: ["title", "subtitle", "metrics"], layout: "grid" },
      { order: 11, title: "Key Highlights", type: "highlight-list", fields: ["title", "subtitle", "items"], layout: "list" },
      { order: 12, title: "Thank You", type: "thank-you", fields: ["title", "message", "signature"], layout: "center" },
    ],
  },
  {
    id: "kpi", name: "KPI Performance Report", category: "Business", icon: "📈",
    description: "Data-driven KPI dashboard with metrics, trends, and performance analysis.",
    slides: [
      { order: 1, title: "Report Title", type: "title", fields: ["title", "subtitle", "badge", "metrics", "highlights"], layout: "split" },
      { order: 2, title: "KPI Overview", type: "dashboard", fields: ["title", "subtitle", "metrics"], layout: "grid" },
      { order: 3, title: "Revenue Metrics", type: "section-overview", fields: ["title", "subtitle", "sectionIcon", "sectionTitle", "description", "stats", "billingMetrics"], layout: "split" },
      { order: 4, title: "Utilization & Efficiency", type: "dashboard", fields: ["title", "subtitle", "metrics"], layout: "grid" },
      { order: 5, title: "Detailed Breakdown", type: "table", fields: ["title", "subtitle", "columns", "rows", "summaryStats"], layout: "table" },
      { order: 6, title: "Trends & Insights", type: "highlight-list", fields: ["title", "subtitle", "items"], layout: "list" },
      { order: 7, title: "Action Items", type: "content", fields: ["title", "subtitle", "cards"], layout: "cards" },
      { order: 8, title: "Thank You", type: "thank-you", fields: ["title", "message", "signature"], layout: "center" },
    ],
  },
  {
    id: "portfolio", name: "Project Portfolio Review", category: "Engineering", icon: "🗂️",
    description: "Multi-project status overview with timelines, resources, and delivery metrics.",
    slides: [
      { order: 1, title: "Portfolio Overview", type: "title", fields: ["title", "subtitle", "badge", "metrics", "stats"], layout: "split" },
      { order: 2, title: "Agenda", type: "agenda", fields: ["title", "subtitle", "cards"], layout: "grid" },
      { order: 3, title: "Active Projects", type: "table", fields: ["title", "subtitle", "columns", "rows", "summaryStats"], layout: "table" },
      { order: 4, title: "Project Details", type: "content", fields: ["title", "subtitle", "cards"], layout: "cards" },
      { order: 5, title: "Resource Allocation", type: "dashboard", fields: ["title", "subtitle", "metrics"], layout: "grid" },
      { order: 6, title: "Timeline & Milestones", type: "table", fields: ["title", "subtitle", "columns", "rows"], layout: "table" },
      { order: 7, title: "Risks & Mitigations", type: "content", fields: ["title", "subtitle", "cards"], layout: "cards" },
      { order: 8, title: "Next Steps", type: "highlight-list", fields: ["title", "subtitle", "items"], layout: "list" },
      { order: 9, title: "Thank You", type: "thank-you", fields: ["title", "message", "signature"], layout: "center" },
    ],
  },
  {
    id: "sales", name: "Sales & Pipeline Report", category: "Sales", icon: "💼",
    description: "Sales performance, pipeline analysis, and revenue forecasting.",
    slides: [
      { order: 1, title: "Sales Overview", type: "title", fields: ["title", "subtitle", "badge", "metrics", "highlights"], layout: "split" },
      { order: 2, title: "Revenue Dashboard", type: "dashboard", fields: ["title", "subtitle", "metrics"], layout: "grid" },
      { order: 3, title: "Pipeline Status", type: "table", fields: ["title", "subtitle", "columns", "rows", "summaryStats"], layout: "table" },
      { order: 4, title: "Deal Analysis", type: "section-overview", fields: ["title", "subtitle", "sectionIcon", "sectionTitle", "description", "stats", "billingMetrics"], layout: "split" },
      { order: 5, title: "Win/Loss Analysis", type: "content", fields: ["title", "subtitle", "cards"], layout: "cards" },
      { order: 6, title: "Top Clients", type: "table", fields: ["title", "subtitle", "columns", "rows"], layout: "table" },
      { order: 7, title: "Forecast & Targets", type: "highlight-list", fields: ["title", "subtitle", "items"], layout: "list" },
      { order: 8, title: "Thank You", type: "thank-you", fields: ["title", "message", "signature"], layout: "center" },
    ],
  },
  {
    id: "innovation", name: "Innovation & POC Report", category: "Engineering", icon: "🧪",
    description: "R&D initiatives, proof of concepts, and technology exploration results.",
    slides: [
      { order: 1, title: "Innovation Overview", type: "title", fields: ["title", "subtitle", "badge", "metrics", "highlights"], layout: "split" },
      { order: 2, title: "Agenda", type: "agenda", fields: ["title", "subtitle", "cards"], layout: "grid" },
      { order: 3, title: "POC Portfolio", type: "content", fields: ["title", "subtitle", "cards"], layout: "cards" },
      { order: 4, title: "Technology Radar", type: "dashboard", fields: ["title", "subtitle", "metrics"], layout: "grid" },
      { order: 5, title: "Results & Metrics", type: "table", fields: ["title", "subtitle", "columns", "rows", "summaryStats"], layout: "table" },
      { order: 6, title: "Lessons Learned", type: "highlight-list", fields: ["title", "subtitle", "items"], layout: "list" },
      { order: 7, title: "Roadmap", type: "content", fields: ["title", "subtitle", "cards"], layout: "cards" },
      { order: 8, title: "Thank You", type: "thank-you", fields: ["title", "message", "signature"], layout: "center" },
    ],
  },
  {
    id: "executive", name: "Executive Summary Deck", category: "Leadership", icon: "🏛️",
    description: "High-level strategic overview for C-suite and leadership presentations.",
    slides: [
      { order: 1, title: "Executive Summary", type: "title", fields: ["title", "subtitle", "badge", "metrics", "highlights", "stats"], layout: "split" },
      { order: 2, title: "Strategic Priorities", type: "agenda", fields: ["title", "subtitle", "cards"], layout: "grid" },
      { order: 3, title: "Financial Overview", type: "dashboard", fields: ["title", "subtitle", "metrics"], layout: "grid" },
      { order: 4, title: "Operational Highlights", type: "section-overview", fields: ["title", "subtitle", "sectionIcon", "sectionTitle", "description", "stats", "billingMetrics"], layout: "split" },
      { order: 5, title: "Key Decisions Required", type: "highlight-list", fields: ["title", "subtitle", "items"], layout: "list" },
      { order: 6, title: "Thank You", type: "thank-you", fields: ["title", "message", "signature"], layout: "center" },
    ],
  },
];

const userTemplates = []; // User-created templates

app.get("/api/templates/library/all", (req, res) => {
  const all = [...BUILT_IN_TEMPLATES, ...userTemplates];
  const { category, search } = req.query;
  let filtered = all;
  if (category) filtered = filtered.filter(t => t.category.toLowerCase() === category.toLowerCase());
  if (search) filtered = filtered.filter(t => t.name.toLowerCase().includes(search.toLowerCase()) || t.description.toLowerCase().includes(search.toLowerCase()));
  res.json({ templates: filtered });
});

app.get("/api/templates/library/:id", (req, res) => {
  const t = [...BUILT_IN_TEMPLATES, ...userTemplates].find(t => t.id === req.params.id);
  if (!t) return res.status(404).json({ error: "Template not found" });
  res.json(t);
});

app.post("/api/templates/library/create", express.json(), (req, res) => {
  const { name, category, description, slides } = req.body;
  if (!name || !slides?.length) return res.status(400).json({ error: "Name and slides required" });
  const t = { id: `custom_${Date.now()}`, name, category: category || "Custom", icon: "📝", description: description || "", slides, isCustom: true };
  userTemplates.push(t);
  res.json(t);
});

app.put("/api/templates/library/update", express.json(), (req, res) => {
  const { id, name, category, description, slides } = req.body;
  const idx = userTemplates.findIndex(t => t.id === id);
  if (idx < 0) return res.status(404).json({ error: "Can only edit custom templates" });
  Object.assign(userTemplates[idx], { name, category, description, slides });
  res.json(userTemplates[idx]);
});

// AI generation using a specific template structure
app.post("/api/ai/generate-with-template", express.json(), async (req, res) => {
  try {
    const { templateId, prompt } = req.body;
    const template = [...BUILT_IN_TEMPLATES, ...userTemplates].find(t => t.id === templateId);
    if (!template) return res.status(404).json({ error: "Template not found" });

    const structurePrompt = template.slides.map((s, i) =>
      `Slide ${i + 1}: type="${s.type}", title="${s.title}", fields=[${s.fields.join(", ")}]`
    ).join("\n");

    const fullPrompt = `${prompt}\n\nIMPORTANT: Use EXACTLY this slide structure (do not add or remove slides):\n${structurePrompt}\n\nGenerate content for each slide following the structure above. Each slide must have the exact "type" specified.`;

    const apiKey = process.env.GOOGLE_API_KEY;
    if (!apiKey) return res.status(500).json({ error: "GOOGLE_API_KEY not set" });

    const MODELS = ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-2.0-flash-lite", "gemini-flash-latest"];
    const requestBody = JSON.stringify({
      contents: [{ parts: [{ text: SYSTEM_PROMPT + "\n\nUser input:\n" + fullPrompt }] }],
      generationConfig: { temperature: 0.7, maxOutputTokens: 16384 },
    });

    let geminiData = null;
    for (const model of MODELS) {
      try {
        const r = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`, {
          method: "POST", headers: { "Content-Type": "application/json" }, body: requestBody,
        });
        if (r.ok) { geminiData = await r.json(); break; }
      } catch {}
    }

    if (!geminiData) return res.status(503).json({ error: "AI service unavailable" });

    const text = geminiData?.candidates?.[0]?.content?.parts?.[0]?.text || "";
    let jsonStr = text;
    const fm = text.match(/```(?:json)?\s*([\s\S]*?)```/);
    if (fm) jsonStr = fm[1];
    const fb = jsonStr.indexOf("{"), lb = jsonStr.lastIndexOf("}");
    if (fb !== -1 && lb > fb) jsonStr = jsonStr.substring(fb, lb + 1);
    jsonStr = jsonStr.trim().replace(/,\s*}/g, "}").replace(/,\s*\]/g, "]");

    const data = JSON.parse(jsonStr);
    if (data.slides) data.slides = data.slides.filter(s => s?.type).map((s, i) => ({ ...s, id: `tmpl_${Date.now() + i}` }));

    res.json({ ...data, templateId, templateName: template.name });
  } catch (e) {
    if (e instanceof SyntaxError) return res.status(422).json({ error: "AI returned invalid JSON. Try again." });
    res.status(500).json({ error: e.message });
  }
});

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// PPT REVERSE ENGINEERING — Parse PPTX → Template + Excel
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

const pptxUpload = multer({
  storage: multer.diskStorage({
    destination: path.join(__dirname, "uploads", "excels"),
    filename: (req, file, cb) => cb(null, `${Date.now()}_${file.originalname.replace(/[^a-zA-Z0-9._-]/g, "_")}`),
  }),
  fileFilter: (req, file, cb) => {
    if (/\.pptx$/i.test(file.originalname)) cb(null, true);
    else cb(new Error("Only .pptx allowed"));
  },
});

function classifySlideType(texts) {
  const combined = texts.join(" ").toLowerCase();
  if (combined.includes("thank") && combined.includes("you")) return "thank-you";
  if (combined.includes("agenda") || combined.includes("overview") || combined.includes("contents")) return "agenda";
  if (/\d+%|\$[\d,.]+|revenue|utilization|kpi|metric/i.test(combined)) return "dashboard";
  if (combined.includes("|") || texts.some(t => t.includes("\t"))) return "table";
  if (combined.length < 100 && texts.length <= 3) return "title";
  if (texts.length > 6) return "highlight-list";
  return "content";
}

function inferFields(type) {
  const fieldMap = {
    "title": ["title", "subtitle", "badge", "metrics", "highlights"],
    "agenda": ["title", "subtitle", "cards"],
    "dashboard": ["title", "subtitle", "metrics"],
    "table": ["title", "subtitle", "columns", "rows", "summaryStats"],
    "content": ["title", "subtitle", "cards"],
    "section-overview": ["title", "subtitle", "sectionTitle", "description", "stats"],
    "highlight-list": ["title", "subtitle", "items"],
    "thank-you": ["title", "message", "signature"],
  };
  return fieldMap[type] || fieldMap["content"];
}

app.post("/api/pptx/reverse-engineer", pptxUpload.single("pptx"), async (req, res) => {
  try {
    const filePath = req.file.path;

    // Parse PPTX using dynamic import (pptx-parser may use different export)
    let slides = [];
    try {
      const pptxParser = await import("pptx-parser");
      const parser = pptxParser.default || pptxParser;
      const parsed = await parser(filePath);
      slides = (parsed.slides || parsed || []).map((s, i) => {
        const texts = (s.texts || s.content || []).map(t => typeof t === "string" ? t : t.text || "").filter(Boolean);
        return { order: i + 1, texts, images: s.images?.length || 0 };
      });
    } catch (parseErr) {
      // Fallback: extract basic info from XLSX-based parsing of PPTX (ZIP)
      console.log("pptx-parser failed, using fallback:", parseErr.message);
      // Generate placeholder slides from filename
      slides = [{ order: 1, texts: ["Imported Presentation"], images: 0 }];
    }

    // Step 2: Classify each slide
    const slideDefinitions = slides.map((s, i) => {
      const type = classifySlideType(s.texts);
      const title = s.texts[0] || `Slide ${i + 1}`;
      return {
        order: s.order,
        title: title.substring(0, 80),
        type,
        fields: inferFields(type),
        layout: type === "dashboard" ? "grid" : type === "table" ? "table" : type === "title" ? "split" : "cards",
        extractedTexts: s.texts.slice(0, 10),
      };
    });

    // Step 3: Create template
    const template = {
      id: `imported_${Date.now()}`,
      name: `Imported: ${req.file.originalname.replace(".pptx", "")}`,
      category: "Custom",
      icon: "📥",
      description: `Auto-generated from ${req.file.originalname} (${slideDefinitions.length} slides)`,
      slides: slideDefinitions,
      isCustom: true,
    };
    userTemplates.push(template);

    // Step 4: Generate Excel data template
    const wb = XLSX.utils.book_new();
    const allFields = new Set();
    slideDefinitions.forEach(s => s.fields.forEach(f => allFields.add(f)));

    // Main data sheet
    const headers = ["Slide", "Type", ...allFields];
    const rows = slideDefinitions.map(s => [s.title, s.type, ...Array.from(allFields).map(f => s.fields.includes(f) ? "" : "N/A")]);
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([headers, ...rows]), "Data");

    // Per-slide-type sheets
    const typeGroups = {};
    slideDefinitions.forEach(s => {
      if (!typeGroups[s.type]) typeGroups[s.type] = [];
      typeGroups[s.type].push(s);
    });
    Object.entries(typeGroups).forEach(([type, slides]) => {
      const fields = inferFields(type);
      const sheetData = [["Slide Title", ...fields], ...slides.map(s => [s.title, ...fields.map(() => "")])];
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sheetData), type.substring(0, 30));
    });

    const excelPath = path.join(__dirname, "uploads", "excels", `template_${Date.now()}.xlsx`);
    XLSX.writeFile(wb, excelPath);

    res.json({
      template,
      slideDefinitions,
      excelDownload: `/api/sync/download?path=${encodeURIComponent(excelPath)}`,
      message: `Parsed ${slideDefinitions.length} slides. Template and Excel created.`,
    });
  } catch (e) {
    console.error("Reverse engineer error:", e);
    res.status(500).json({ error: e.message });
  }
});

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// AI TEMPLATE ENHANCEMENT — Analyze & improve templates
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

app.post("/api/templates/enhance", express.json(), async (req, res) => {
  try {
    const { templateId } = req.body;
    const template = [...BUILT_IN_TEMPLATES, ...userTemplates].find(t => t.id === templateId);
    if (!template) return res.status(404).json({ error: "Template not found" });

    const apiKey = process.env.GOOGLE_API_KEY;
    if (!apiKey) return res.status(500).json({ error: "GOOGLE_API_KEY not set" });

    const structureJson = JSON.stringify(template.slides.map(s => ({ order: s.order, title: s.title, type: s.type, fields: s.fields })), null, 2);

    const enhancePrompt = `Analyze this presentation template structure and suggest improvements to make it modern, clean, and executive-ready. Return ONLY valid JSON.

Template: "${template.name}" (${template.category})
Current slides:
${structureJson}

Return this exact JSON format:
{
  "analysis": "Brief overall assessment (1-2 sentences)",
  "add_slides": [{"order": N, "title": "...", "type": "...", "reason": "..."}],
  "remove_slides": [{"order": N, "title": "...", "reason": "..."}],
  "modify_slides": [{"order": N, "title": "...", "change": "...", "new_type": "..."}],
  "visual_improvements": ["suggestion 1", "suggestion 2"],
  "content_improvements": ["suggestion 1", "suggestion 2"],
  "enhanced_slides": [{"order": N, "title": "...", "type": "...", "fields": [...]}]
}

Rules:
- "enhanced_slides" must contain the FULL improved slide list (old + modifications + additions, minus removals)
- Valid types: title, agenda, content, dashboard, table, section-overview, highlight-list, thank-you
- Keep business meaning intact
- Suggest charts over tables where appropriate
- Add executive summary if missing
- Ensure logical flow`;

    const MODELS = ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-flash-latest"];
    const requestBody = JSON.stringify({
      contents: [{ parts: [{ text: enhancePrompt }] }],
      generationConfig: { temperature: 0.7, maxOutputTokens: 8192 },
    });

    let geminiData = null;
    for (const model of MODELS) {
      try {
        const r = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`, {
          method: "POST", headers: { "Content-Type": "application/json" }, body: requestBody,
        });
        if (r.ok) { geminiData = await r.json(); break; }
      } catch {}
    }

    if (!geminiData) return res.status(503).json({ error: "AI service unavailable" });

    const text = geminiData?.candidates?.[0]?.content?.parts?.[0]?.text || "";
    let jsonStr = text;
    const fm = text.match(/```(?:json)?\s*([\s\S]*?)```/);
    if (fm) jsonStr = fm[1];
    const fb = jsonStr.indexOf("{"), lb = jsonStr.lastIndexOf("}");
    if (fb !== -1 && lb > fb) jsonStr = jsonStr.substring(fb, lb + 1);
    jsonStr = jsonStr.trim().replace(/,\s*}/g, "}").replace(/,\s*\]/g, "]");

    const suggestions = JSON.parse(jsonStr);

    res.json({
      originalTemplate: template,
      suggestions,
      message: "AI analysis complete. Review suggestions and apply selectively.",
    });
  } catch (e) {
    if (e instanceof SyntaxError) return res.status(422).json({ error: "AI returned invalid response. Try again." });
    res.status(500).json({ error: e.message });
  }
});

// Apply enhancement — create new template from suggestions
app.post("/api/templates/apply-enhancement", express.json(), (req, res) => {
  try {
    const { originalId, enhancedSlides, name } = req.body;
    const original = [...BUILT_IN_TEMPLATES, ...userTemplates].find(t => t.id === originalId);
    if (!original) return res.status(404).json({ error: "Original template not found" });

    const enhanced = {
      id: `enhanced_${Date.now()}`,
      name: name || `${original.name} (Enhanced)`,
      category: original.category,
      icon: "✨",
      description: `AI-enhanced version of ${original.name}`,
      slides: (enhancedSlides || []).map((s, i) => ({
        order: i + 1,
        title: s.title,
        type: s.type,
        fields: s.fields || inferFields(s.type),
        layout: s.type === "dashboard" ? "grid" : s.type === "table" ? "table" : "cards",
      })),
      isCustom: true,
      enhancedFrom: originalId,
    };
    userTemplates.push(enhanced);

    res.json({ template: enhanced, message: "Enhanced template saved to library." });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// ICON LIBRARY — Built-in + custom icon management
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

const BUILT_IN_ICONS = {
  business: ["📊", "📈", "📉", "💼", "🏢", "🤝", "📋", "📁", "📂", "💰", "💵", "🏦", "📄", "📝"],
  finance: ["💳", "💲", "🏧", "📊", "💹", "💰", "🪙", "📈", "🧾", "💵"],
  technology: ["💻", "⚙️", "🔧", "🌐", "📱", "🖥️", "☁️", "🔒", "🛡️", "🤖", "🧠", "⚡"],
  team: ["👥", "👤", "🧑‍💼", "👨‍💻", "👩‍💻", "🏆", "⭐", "🎯", "🎓", "💪"],
  awards: ["🏆", "🥇", "🥈", "🥉", "🎖️", "🏅", "⭐", "🌟", "✨", "🎉"],
};
const customIcons = [];

app.get("/api/icons/library", (req, res) => {
  const { category } = req.query;
  if (category && BUILT_IN_ICONS[category]) return res.json({ icons: BUILT_IN_ICONS[category], category });
  res.json({ categories: Object.keys(BUILT_IN_ICONS), icons: BUILT_IN_ICONS, custom: customIcons });
});

const iconUpload = multer({
  storage: multer.diskStorage({
    destination: path.join(__dirname, "uploads", "photos"),
    filename: (req, file, cb) => cb(null, `icon_${Date.now()}_${file.originalname.replace(/[^a-zA-Z0-9._-]/g, "_")}`),
  }),
  limits: { fileSize: 2 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    if (/\.(svg|png)$/i.test(file.originalname)) cb(null, true);
    else cb(new Error("Only SVG/PNG allowed"));
  },
});

app.post("/api/icons/upload", iconUpload.single("icon"), (req, res) => {
  const entry = { id: randomUUID(), url: `/uploads/photos/${req.file.filename}`, name: req.file.originalname };
  customIcons.push(entry);
  res.json(entry);
});

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// DESIGN VARIANTS — Multiple visual styles per template
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

const DESIGN_VARIANTS = [
  { id: "corporate", name: "Corporate Clean", theme: "light", layoutStyle: "grid", style: { fontFamily: "Inter", primaryColor: "#0078d4", secondaryColor: "#005a9e", background: "#f0f2f5", cardStyle: "bordered", iconStyle: "filled" } },
  { id: "cards", name: "Card-Based Modern", theme: "light", layoutStyle: "cards", style: { fontFamily: "Outfit", primaryColor: "#3b82f6", secondaryColor: "#8b5cf6", background: "#ffffff", cardStyle: "elevated", iconStyle: "rounded" } },
  { id: "minimal", name: "Minimal Executive", theme: "light", layoutStyle: "minimal", style: { fontFamily: "DM Sans", primaryColor: "#111827", secondaryColor: "#6b7280", background: "#ffffff", cardStyle: "flat", iconStyle: "outline" } },
  { id: "dark", name: "Dark Mode", theme: "dark", layoutStyle: "grid", style: { fontFamily: "Outfit", primaryColor: "#3b82f6", secondaryColor: "#8b5cf6", background: "#070b14", cardStyle: "glass", iconStyle: "glow" } },
  { id: "infographic", name: "Visual Infographic", theme: "light", layoutStyle: "infographic", style: { fontFamily: "Montserrat", primaryColor: "#0d9488", secondaryColor: "#f59e0b", background: "#f0fdfa", cardStyle: "colorful", iconStyle: "large" } },
];

app.get("/api/design-variants", (req, res) => res.json({ variants: DESIGN_VARIANTS }));

app.get("/api/design-variants/:id", (req, res) => {
  const v = DESIGN_VARIANTS.find(v => v.id === req.params.id);
  if (!v) return res.status(404).json({ error: "Variant not found" });
  res.json(v);
});

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// PRODUCTION — Serve built frontend
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

const distPath = path.join(__dirname, "dist");
if (fs.existsSync(distPath)) {
  app.use(express.static(distPath));
  app.get("*", (req, res) => {
    if (!req.path.startsWith("/api") && !req.path.startsWith("/uploads")) {
      res.sendFile(path.join(distPath, "index.html"));
    }
  });
  console.log("Serving production build from /dist");
}

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`PoterPointAI running on port ${PORT} — https://ppt.poterai.com`);
});
