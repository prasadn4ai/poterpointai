import XLSX from "xlsx";

const BASE = "http://localhost:3001";
let pass = 0, fail = 0, results = [];

async function test(name, fn) {
  try {
    await fn();
    pass++;
    results.push("  PASS: " + name);
  } catch (e) {
    fail++;
    results.push("  FAIL: " + name + " — " + e.message);
  }
}

function assert(cond, msg) { if (!cond) throw new Error(msg || "Assertion failed"); }

async function run() {
  console.log("================================================================");
  console.log("PoterPointAI — Production Test Suite");
  console.log("================================================================\n");

  // ━━━ SUITE 1: Backend API Tests ━━━
  console.log("--- SUITE 1: Backend API Endpoints ---");

  await test("1.01 Server health check", async () => {
    const r = await fetch(BASE + "/api/photos/list");
    assert(r.ok, "Status: " + r.status);
  });

  await test("1.02 Photo upload with auto-mapping", async () => {
    const form = new FormData();
    form.append("photos", new Blob(["fake"], { type: "image/png" }), "VP_TestPerson.png");
    const r = await fetch(BASE + "/api/photos/upload", { method: "POST", body: form });
    const d = await r.json();
    assert(d.photos?.length === 1, "Expected 1 photo");
    assert(d.photos[0].name === "TestPerson", "Auto-map name: " + d.photos[0].name);
    assert(d.photos[0].designation === "VP", "Auto-map role: " + d.photos[0].designation);
  });

  await test("1.03 Photo list", async () => {
    const r = await fetch(BASE + "/api/photos/list");
    const d = await r.json();
    assert(d.photos?.length >= 1, "Expected photos");
  });

  // Create test Excel
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["Name", "Revenue", "Status"], ["Test", "50000", "Active"]]), "KPI");
  const xlsxBuf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

  await test("1.04 Template upload (.xlsx)", async () => {
    const form = new FormData();
    form.append("template", new Blob([xlsxBuf]), "test.xlsx");
    const r = await fetch(BASE + "/api/templates/upload", { method: "POST", body: form });
    const d = await r.json();
    assert(d.version, "Expected version");
    assert(d.active === true, "Expected active");
  });

  await test("1.05 Template list", async () => {
    const r = await fetch(BASE + "/api/templates/list");
    const d = await r.json();
    assert(d.templates?.length >= 1, "Expected templates");
  });

  await test("1.06 Template latest returns active", async () => {
    const r = await fetch(BASE + "/api/templates/latest");
    const d = await r.json();
    assert(d.active === true, "Expected active");
  });

  await test("1.07 Template library: 6+ built-in templates", async () => {
    const r = await fetch(BASE + "/api/templates/library/all");
    const d = await r.json();
    assert(d.templates?.length >= 6, "Got " + d.templates?.length);
  });

  await test("1.08 Template library: category filter (Sales)", async () => {
    const r = await fetch(BASE + "/api/templates/library/all?category=Sales");
    const d = await r.json();
    assert(d.templates?.length >= 1, "Expected Sales templates");
    assert(d.templates.every(t => t.category === "Sales"), "Non-Sales found");
  });

  await test("1.09 Template library: search filter", async () => {
    const r = await fetch(BASE + "/api/templates/library/all?search=executive");
    const d = await r.json();
    assert(d.templates?.length >= 1, "Expected search results");
  });

  await test("1.10 Template library: get by ID (qbr)", async () => {
    const r = await fetch(BASE + "/api/templates/library/qbr");
    const d = await r.json();
    assert(d.name === "Quarterly Business Review", "Wrong name");
    assert(d.slides?.length === 12, "Expected 12 slides, got " + d.slides?.length);
  });

  await test("1.11 Template-to-Excel sync", async () => {
    const form = new FormData();
    form.append("excel", new Blob([xlsxBuf]), "data.xlsx");
    const r = await fetch(BASE + "/api/sync/template-to-excel", { method: "POST", body: form });
    const d = await r.json();
    assert(Array.isArray(d.changes), "Expected changes array");
  });

  await test("1.12 Excel-to-Template sync", async () => {
    const wb2 = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb2, XLSX.utils.aoa_to_sheet([["Name", "Revenue", "Status", "NewField"], ["A", "1", "X", "Y"]]), "KPI");
    const buf2 = XLSX.write(wb2, { type: "buffer", bookType: "xlsx" });
    const form = new FormData();
    form.append("excel", new Blob([buf2]), "extra.xlsx");
    const r = await fetch(BASE + "/api/sync/excel-to-template", { method: "POST", body: form });
    const d = await r.json();
    assert(d.changes?.length >= 1, "Expected changes");
  });

  await test("1.13 Icon library: 5 categories, 50+ icons", async () => {
    const r = await fetch(BASE + "/api/icons/library");
    const d = await r.json();
    assert(d.categories?.length === 5, "Expected 5 categories");
    const total = Object.values(d.icons).flat().length;
    assert(total >= 50, "Expected 50+ icons, got " + total);
  });

  await test("1.14 Design variants: 5 variants", async () => {
    const r = await fetch(BASE + "/api/design-variants");
    const d = await r.json();
    assert(d.variants?.length === 5, "Expected 5");
    assert(d.variants[0].style?.fontFamily, "Expected style.fontFamily");
  });

  await test("1.15 Auto-create template from mismatched Excel", async () => {
    const wb3 = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb3, XLSX.utils.aoa_to_sheet([["Unique1", "Unique2"], ["a", "b"]]), "NewSheet");
    const buf3 = XLSX.write(wb3, { type: "buffer", bookType: "xlsx" });
    const form = new FormData();
    form.append("excel", new Blob([buf3]), "mismatch.xlsx");
    const r = await fetch(BASE + "/api/templates/auto-create", { method: "POST", body: form });
    const d = await r.json();
    assert(d.mismatch === true, "Expected mismatch");
    assert(d.template?.id, "Expected template");
  });

  await test("1.16 Create custom template", async () => {
    const r = await fetch(BASE + "/api/templates/library/create", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ name: "Test Custom", category: "Custom", slides: [{ order: 1, title: "Slide1", type: "title", fields: ["title"] }] }),
    });
    const d = await r.json();
    assert(d.id?.startsWith("custom_"), "Expected custom_ prefix");
    assert(d.isCustom === true, "Expected isCustom");
  });

  await test("1.17 AI generate endpoint responds (not 500)", async () => {
    const r = await fetch(BASE + "/api/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ prompt: "Create 2 test slides" }),
    });
    assert(r.status !== 500, "Server error: " + r.status);
  });

  await test("1.18 AI generate rejects empty prompt", async () => {
    const r = await fetch(BASE + "/api/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ prompt: "" }),
    });
    assert(r.status === 400, "Expected 400, got " + r.status);
  });

  await test("1.19 Template enhance endpoint responds", async () => {
    const r = await fetch(BASE + "/api/templates/enhance", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ templateId: "executive" }),
    });
    assert(r.status !== 500, "Server error");
  });

  await test("1.20 Template enhance: 404 for bad ID", async () => {
    const r = await fetch(BASE + "/api/templates/enhance", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ templateId: "nonexistent" }),
    });
    assert(r.status === 404, "Expected 404");
  });

  await test("1.21 AI generate with template", async () => {
    const r = await fetch(BASE + "/api/ai/generate-with-template", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ templateId: "executive", prompt: "Create executive deck for Q1" }),
    });
    assert(r.status !== 500, "Server error");
  });

  await test("1.22 Static file serving (/uploads path)", async () => {
    const r = await fetch(BASE + "/uploads/");
    // Should not 500 — may 404 or 403 which is fine
    assert(r.status !== 500, "Server error");
  });

  await test("1.23 File validation: reject non-xlsx template", async () => {
    const form = new FormData();
    form.append("template", new Blob(["not excel"]), "bad.txt");
    const r = await fetch(BASE + "/api/templates/upload", { method: "POST", body: form });
    assert(!r.ok, "Should reject non-xlsx");
  });

  await test("1.24 Design variant get by ID", async () => {
    const r = await fetch(BASE + "/api/design-variants/corporate");
    const d = await r.json();
    assert(d.name === "Corporate Clean", "Wrong name: " + d.name);
    assert(d.theme === "light", "Expected light theme");
  });

  await test("1.25 Design variant 404 for bad ID", async () => {
    const r = await fetch(BASE + "/api/design-variants/nonexistent");
    assert(r.status === 404, "Expected 404");
  });

  // ━━━ Print Results ━━━
  console.log("");
  results.forEach(r => console.log(r));
  console.log("\n================================================================");
  console.log("TOTAL: " + pass + " passed, " + fail + " failed out of " + (pass + fail) + " tests");
  console.log("================================================================");

  if (fail > 0) process.exit(1);
}

run().catch(e => { console.error("Runner error:", e); process.exit(1); });
