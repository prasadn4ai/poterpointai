import { useState, useCallback, useRef, useEffect, createContext, useContext, useMemo } from "react";

/*
 ╔══════════════════════════════════════════════════════════════════════╗
 ║  DECK STUDIO v4 — Step 4: Export Engine (PPTX + PDF)               ║
 ║  ──────────────────────────────────────────────────────────────────  ║
 ║  ✅ PPTX Export → PptxGenJS with per-slide-type renderers           ║
 ║  ✅ PDF Export  → html2canvas + jsPDF pixel-perfect capture         ║
 ║  ✅ JSON Export → Raw data download                                 ║
 ║  ──────────────────────────────────────────────────────────────────  ║
 ║  Steps 1-3 preserved: Store, Schema, Editor, Slide Components      ║
 ╚══════════════════════════════════════════════════════════════════════╝
*/

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 🗄️  STORE — Zustand-pattern state management
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

const StoreContext = createContext(null);

function createStore(initialState) {
  let state = initialState;
  const listeners = new Set();
  return {
    getState: () => state,
    setState: (updater) => {
      state = typeof updater === "function" ? updater(state) : { ...state, ...updater };
      listeners.forEach((l) => l(state));
    },
    subscribe: (listener) => {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
  };
}

function useStore(selector) {
  const store = useContext(StoreContext);
  const [slice, setSlice] = useState(() => selector(store.getState()));
  useEffect(() => {
    return store.subscribe((s) => {
      const next = selector(s);
      setSlice(next);
    });
  }, [store, selector]);
  return slice;
}

function useStoreActions() {
  const store = useContext(StoreContext);
  return useMemo(() => ({
    addSlide: (type) => {
      const defaults = SLIDE_SCHEMA[type];
      if (!defaults) return;
      store.setState((s) => ({
        ...s,
        slides: [...s.slides, { ...JSON.parse(JSON.stringify(defaults.defaults)), id: `slide_${Date.now()}` }],
        activeSlide: s.slides.length,
      }));
    },
    updateSlide: (id, data) => {
      store.setState((s) => ({
        ...s,
        slides: s.slides.map((sl) => (sl.id === id ? { ...sl, ...data } : sl)),
      }));
    },
    removeSlide: (index) => {
      store.setState((s) => {
        if (s.slides.length <= 1) return s;
        const slides = s.slides.filter((_, i) => i !== index);
        return { ...s, slides, activeSlide: Math.min(s.activeSlide, slides.length - 1) };
      });
    },
    setActive: (index) => {
      store.setState((s) => ({ ...s, activeSlide: index }));
    },
    moveSlide: (from, dir) => {
      store.setState((s) => {
        const to = from + dir;
        if (to < 0 || to >= s.slides.length) return s;
        const slides = [...s.slides];
        [slides[from], slides[to]] = [slides[to], slides[from]];
        return { ...s, slides, activeSlide: to };
      });
    },
    duplicateSlide: (index) => {
      store.setState((s) => {
        const dup = { ...JSON.parse(JSON.stringify(s.slides[index])), id: `slide_${Date.now()}` };
        const slides = [...s.slides];
        slides.splice(index + 1, 0, dup);
        return { ...s, slides, activeSlide: index + 1 };
      });
    },
    updatePresentation: (data) => {
      store.setState((s) => ({ ...s, presentation: { ...s.presentation, ...data } }));
    },
    setCanvasMode: (enabled) => {
      store.setState((s) => ({ ...s, canvasMode: enabled, canvasSelectedId: null, canvasSelectedIds: [] }));
    },
    setCanvasSelectedId: (id) => {
      store.setState((s) => ({ ...s, canvasSelectedId: id, canvasSelectedIds: id ? [id] : [] }));
    },
    setCanvasSelectedIds: (ids) => {
      store.setState((s) => ({ ...s, canvasSelectedIds: ids, canvasSelectedId: ids[0] || null }));
    },
    toggleCanvasSelection: (id) => {
      store.setState((s) => {
        const cur = s.canvasSelectedIds || [];
        const next = cur.includes(id) ? cur.filter(i => i !== id) : [...cur, id];
        return { ...s, canvasSelectedIds: next, canvasSelectedId: next[0] || null };
      });
    },
    groupElements: (slideId, elementIds, groupId) => {
      store.setState((s) => ({
        ...s,
        slides: s.slides.map(sl => sl.id === slideId
          ? { ...sl, elementGroups: { ...(sl.elementGroups || {}), ...Object.fromEntries(elementIds.map(eid => [eid, groupId])) } }
          : sl),
      }));
    },
    ungroupElements: (slideId, groupId) => {
      store.setState((s) => ({
        ...s,
        slides: s.slides.map(sl => {
          if (sl.id !== slideId) return sl;
          const eg = { ...(sl.elementGroups || {}) };
          Object.keys(eg).forEach(k => { if (eg[k] === groupId) delete eg[k]; });
          return { ...sl, elementGroups: eg };
        }),
      }));
    },
    updateElementStyle: (slideId, elementId, styleUpdates) => {
      store.setState((s) => ({
        ...s,
        slides: s.slides.map((sl) =>
          sl.id === slideId
            ? { ...sl, elementStyles: { ...(sl.elementStyles || {}), [elementId]: { ...(sl.elementStyles?.[elementId] || {}), ...styleUpdates } } }
            : sl
        ),
      }));
    },
    updateLayout: (slideId, elementId, rect) => {
      store.setState((s) => ({
        ...s,
        slides: s.slides.map((sl) =>
          sl.id === slideId
            ? { ...sl, layout: { ...(sl.layout || {}), [elementId]: { ...(sl.layout?.[elementId] || {}), ...rect } } }
            : sl
        ),
      }));
    },
    resetLayout: (slideId) => {
      store.setState((s) => ({
        ...s,
        slides: s.slides.map((sl) =>
          sl.id === slideId ? { ...sl, layout: undefined } : sl
        ),
      }));
    },
    setThemeMode: (mode) => {
      store.setState((s) => ({ ...s, themeMode: mode }));
    },
    setColorCombo: (combo) => {
      store.setState((s) => ({ ...s, colorCombo: combo }));
    },
    setDesignVariant: (variant) => {
      store.setState((s) => ({ ...s, designVariant: variant }));
    },
    loadPresentation: (slides) => {
      const now = Date.now();
      store.setState((s) => ({
        ...s,
        slides: slides.map((sl, i) => ({ ...sl, id: sl.id || `ai_${now + i}` })),
        activeSlide: 0,
      }));
    },
  }), [store]);
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 📐  SCHEMA — Slide type definitions & field configs
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

const SLIDE_SCHEMA = {
  title: {
    label: "Title Slide",
    icon: "◆",
    description: "Opening slide with key metrics",
    fields: [
      { name: "title", label: "Title", type: "text", placeholder: "Presentation title" },
      { name: "subtitle", label: "Subtitle", type: "text", placeholder: "Subtitle or tagline" },
      { name: "badge", label: "Badge / Period", type: "text", placeholder: "e.g. Q1 2026" },
      {
        name: "metrics", label: "Key Metrics", type: "array",
        itemFields: [
          { name: "number", label: "Value", type: "text", width: "30%" },
          { name: "label", label: "Label", type: "text", width: "70%" },
        ],
        addLabel: "Add Metric",
      },
      {
        name: "highlights", label: "Highlights", type: "array",
        itemFields: [
          { name: "color", label: "Color", type: "color", width: "48px" },
          { name: "text", label: "Text", type: "text", width: "auto" },
        ],
        addLabel: "Add Highlight",
        defaultItem: { color: "#0078d4", text: "New highlight" },
      },
      {
        name: "stats", label: "Info Stats", type: "array",
        itemFields: [
          { name: "icon", label: "Icon", type: "text", width: "60px" },
          { name: "label", label: "Label", type: "text", width: "35%" },
          { name: "value", label: "Value", type: "text", width: "35%" },
        ],
        addLabel: "Add Stat",
        defaultItem: { icon: "📊", label: "Stat", value: "Value" },
      },
    ],
    defaults: {
      type: "title",
      title: "Quarterly Business Review",
      subtitle: "Microsoft Practice Division — Driving Cloud Excellence",
      badge: "Q1 2026",
      metrics: [
        { number: "24", label: "Resources" },
        { number: "$1.2M", label: "Revenue" },
        { number: "94%", label: "Utilization" },
      ],
      highlights: [
        { color: "#22c55e", text: "Revenue up 18% year-over-year" },
        { color: "#3b82f6", text: "3 new enterprise clients onboarded" },
      ],
      stats: [
        { icon: "👥", label: "Team Size", value: "24 Resources" },
        { icon: "🏢", label: "Active Clients", value: "12 Enterprises" },
        { icon: "📈", label: "Growth Rate", value: "+18% YoY" },
      ],
    },
  },

  agenda: {
    label: "Agenda",
    icon: "☰",
    description: "Meeting agenda with topic cards",
    fields: [
      { name: "title", label: "Title", type: "text" },
      { name: "subtitle", label: "Subtitle", type: "text" },
      {
        name: "cards", label: "Agenda Sections", type: "compound-array",
        itemFields: [
          { name: "icon", label: "Icon", type: "text", width: "60px", placeholder: "emoji" },
          { name: "title", label: "Section Title", type: "text" },
        ],
        subArray: { name: "items", label: "Points", type: "string-array", addLabel: "Add Point" },
        addLabel: "Add Section",
        defaultItem: { icon: "📌", title: "New Section", items: ["Point 1"] },
      },
    ],
    defaults: {
      type: "agenda",
      title: "Today's Agenda",
      subtitle: "Quarterly Review — Q1 2026",
      cards: [
        { icon: "👥", title: "Team & Resources", items: ["Headcount update", "New hires & departures", "Skill matrix review"] },
        { icon: "📊", title: "Performance Metrics", items: ["Revenue & billing", "Utilization rates", "Customer satisfaction"] },
        { icon: "🚀", title: "Initiatives & Roadmap", items: ["Active projects", "POC pipeline", "Q2 goals"] },
        { icon: "🏆", title: "Achievements", items: ["Awards & recognition", "Team events", "Knowledge sharing"] },
      ],
    },
  },

  content: {
    label: "Content",
    icon: "▤",
    description: "Two-column card layout",
    fields: [
      { name: "title", label: "Title", type: "text" },
      { name: "subtitle", label: "Subtitle", type: "text" },
      {
        name: "cards", label: "Content Cards", type: "compound-array",
        itemFields: [
          { name: "icon", label: "Icon", type: "text", width: "60px" },
          { name: "title", label: "Card Title", type: "text" },
          { name: "status", label: "Status", type: "select", options: ["Active", "In Progress", "Planned", "Complete"] },
        ],
        subArray: {
          name: "items", label: "Details", type: "kv-array",
          addLabel: "Add Detail",
          keyPlaceholder: "Label:",
          valuePlaceholder: "Description",
        },
        addLabel: "Add Card",
        defaultItem: { icon: "📋", title: "New Card", status: "Active", items: [{ bold: "Key:", text: "Value" }] },
      },
    ],
    defaults: {
      type: "content",
      title: "Key Initiatives",
      subtitle: "Current strategic focus areas",
      cards: [
        { icon: "⚡", title: "Cloud Migration Program", status: "Active", items: [{ bold: "Scope:", text: "12 workloads across 3 clients" }, { bold: "Progress:", text: "67% complete, on schedule" }] },
        { icon: "🔒", title: "Zero Trust Security", status: "In Progress", items: [{ bold: "Timeline:", text: "Complete by end of Q2" }, { bold: "Impact:", text: "All enterprise clients" }] },
        { icon: "🤖", title: "AI/ML Practice Launch", status: "Planned", items: [{ bold: "Goal:", text: "3 pilot projects by Q3" }, { bold: "Investment:", text: "$200K training budget" }] },
        { icon: "📐", title: "DevOps Acceleration", status: "Active", items: [{ bold: "Metric:", text: "Reduce deployment time 40%" }, { bold: "Tooling:", text: "GitHub Actions + Azure DevOps" }] },
      ],
    },
  },

  dashboard: {
    label: "Dashboard",
    icon: "◫",
    description: "Metrics grid with KPIs",
    fields: [
      { name: "title", label: "Title", type: "text" },
      { name: "subtitle", label: "Subtitle", type: "text" },
      {
        name: "metrics", label: "Metrics", type: "array",
        itemFields: [
          { name: "icon", label: "Icon", type: "text", width: "60px" },
          { name: "value", label: "Value", type: "text", width: "25%" },
          { name: "label", label: "Label", type: "text", width: "35%" },
          { name: "detail", label: "Change", type: "text", width: "25%" },
        ],
        addLabel: "Add Metric",
        defaultItem: { icon: "📊", value: "0", label: "New Metric", detail: "" },
      },
    ],
    defaults: {
      type: "dashboard",
      title: "Performance Dashboard",
      subtitle: "Key metrics at a glance — Q1 2026",
      metrics: [
        { icon: "👥", value: "24", label: "Team Size", detail: "+3 this quarter" },
        { icon: "💰", value: "$1.2M", label: "Revenue", detail: "+18% YoY" },
        { icon: "📈", value: "94%", label: "Utilization", detail: "Target: 90%" },
        { icon: "⭐", value: "4.8", label: "CSAT Score", detail: "Industry: 4.2" },
        { icon: "🏆", value: "12", label: "Projects Delivered", detail: "+5 vs last Q" },
        { icon: "🔄", value: "98%", label: "Retention Rate", detail: "All-time high" },
      ],
    },
  },

  table: {
    label: "Data Table",
    icon: "▦",
    description: "Structured data in table format",
    fields: [
      { name: "title", label: "Title", type: "text" },
      { name: "subtitle", label: "Subtitle", type: "text" },
      { name: "columns", label: "Columns (comma-separated)", type: "csv" },
      { name: "rows", label: "Rows", type: "table-rows", addLabel: "Add Row" },
      {
        name: "summaryStats", label: "Summary Stats", type: "array",
        itemFields: [
          { name: "number", label: "Value", type: "text", width: "30%" },
          { name: "label", label: "Label", type: "text", width: "70%" },
        ],
        addLabel: "Add Stat",
      },
    ],
    defaults: {
      type: "table",
      title: "Active Engagements",
      subtitle: "Current project portfolio",
      columns: ["Project", "Client", "Resources", "Status", "Timeline"],
      rows: [
        { cells: ["Azure Migration", "Contoso Ltd", "6", "Active", "Q1–Q2 2026"] },
        { cells: ["Power Platform", "Fabrikam", "4", "Active", "Q1 2026"] },
        { cells: ["M365 Rollout", "Woodgrove", "3", "Planning", "Q2 2026"] },
        { cells: ["Security Audit", "Northwind", "2", "Complete", "Q4 2025"] },
      ],
      summaryStats: [
        { number: "4", label: "Projects" },
        { number: "15", label: "Resources Deployed" },
      ],
    },
  },

  "section-overview": {
    label: "Section Overview",
    icon: "◈",
    description: "Section intro with stats & billing",
    fields: [
      { name: "title", label: "Title", type: "text" },
      { name: "subtitle", label: "Subtitle", type: "text" },
      { name: "sectionIcon", label: "Section Icon", type: "text" },
      { name: "sectionTitle", label: "Section Name", type: "text" },
      { name: "description", label: "Description", type: "textarea" },
      {
        name: "stats", label: "Stats", type: "array",
        itemFields: [
          { name: "number", label: "Value", type: "text", width: "30%" },
          { name: "label", label: "Label", type: "text", width: "70%" },
        ],
        addLabel: "Add Stat",
      },
      {
        name: "billingMetrics", label: "Billing Metrics", type: "array",
        itemFields: [
          { name: "label", label: "Label", type: "text", width: "50%" },
          { name: "value", label: "Value", type: "text", width: "50%" },
        ],
        addLabel: "Add Billing Metric",
      },
    ],
    defaults: {
      type: "section-overview",
      title: "Resource vs Billing",
      subtitle: "Team allocation and billing overview",
      sectionIcon: "👥",
      sectionTitle: "Resource Management",
      description: "Overview of our team allocation, utilization rates, and billing performance across all active engagements for the current quarter.",
      stats: [
        { number: "24", label: "Total Resources" },
        { number: "22", label: "Billable" },
        { number: "94%", label: "Utilization" },
      ],
      billingMetrics: [
        { label: "Total Hours", value: "4,320" },
        { label: "Billed Hours", value: "4,060" },
        { label: "Revenue", value: "$1.2M" },
        { label: "Avg Rate", value: "$185/hr" },
      ],
    },
  },

  "highlight-list": {
    label: "Highlights",
    icon: "★",
    description: "Key highlight points",
    fields: [
      { name: "title", label: "Title", type: "text" },
      { name: "subtitle", label: "Subtitle", type: "text" },
      {
        name: "items", label: "Highlight Items", type: "array",
        itemFields: [
          { name: "color", label: "Color", type: "color", width: "48px" },
          { name: "text", label: "Text", type: "text", width: "auto" },
        ],
        addLabel: "Add Highlight",
        defaultItem: { color: "#3b82f6", text: "New highlight" },
      },
    ],
    defaults: {
      type: "highlight-list",
      title: "Key Highlights",
      subtitle: "Quarter achievements and milestones",
      items: [
        { color: "#22c55e", text: "Revenue exceeded target by 12%, reaching $1.2M for Q1" },
        { color: "#3b82f6", text: "Successfully onboarded 3 new enterprise clients" },
        { color: "#8b5cf6", text: "Launched AI/ML practice with 2 pilot projects" },
        { color: "#f59e0b", text: "Team utilization hit 94% — highest in 6 quarters" },
        { color: "#06b6d4", text: "Zero critical incidents across all managed environments" },
        { color: "#ec4899", text: "Won Microsoft Partner of the Year nomination" },
      ],
    },
  },

  "thank-you": {
    label: "Thank You",
    icon: "♥",
    description: "Closing slide",
    fields: [
      { name: "title", label: "Title", type: "text" },
      { name: "message", label: "Closing Message", type: "textarea" },
      { name: "signature", label: "Team / Signature", type: "text" },
    ],
    defaults: {
      type: "thank-you",
      title: "Thank You",
      message: "We appreciate your time and partnership. Looking forward to an exceptional Q2.",
      signature: "Microsoft Practice Team",
    },
  },
};

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 🖱️  CANVAS ELEMENT MAP — Generates draggable blocks per slide type
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/*
  Each entry returns an array of { id, label, defaultRect: {x,y,w,h} }
  Coordinates are in 960×540 slide-space.
  `render` functions are added later (in the CanvasBlock component)
  to keep this map lightweight — blocks render inline from slide data.
*/

function generateCanvasElements(slide) {
  const fn = CANVAS_ELEMENT_MAP[slide.type];
  const slideEls = fn ? fn(slide) : [];
  // Add header bar, footer bar, and their child elements
  return [
    { id: "header-bar", label: "Header Bar", defaultRect: { x: 0, y: 0, w: 960, h: 48 }, subType: "header-bar" },
    { id: "header-logo", label: "Logo", defaultRect: { x: 14, y: 8, w: 32, h: 32 }, subType: "header" },
    { id: "header-title", label: "Header Title", defaultRect: { x: 52, y: 10, w: 160, h: 28 }, subType: "header" },
    { id: "header-period", label: "Period", defaultRect: { x: 810, y: 12, w: 80, h: 24 }, subType: "header" },
    ...slideEls,
    { id: "footer-bar", label: "Footer Bar", defaultRect: { x: 0, y: 504, w: 960, h: 36 }, subType: "footer-bar" },
    { id: "footer-logo", label: "Footer Logo", defaultRect: { x: 14, y: 508, w: 100, h: 24 }, subType: "footer" },
    { id: "footer-page", label: "Page", defaultRect: { x: 870, y: 510, w: 70, h: 20 }, subType: "footer" },
  ];
}

const CANVAS_ELEMENT_MAP = {
  "title": (s) => {
    const els = [];
    if (s.badge) els.push({ id: "badge", label: "Badge", defaultRect: { x: 36, y: 90, w: 130, h: 32 } });
    els.push({ id: "title", label: "Title", defaultRect: { x: 36, y: 130, w: 460, h: 65 } });
    els.push({ id: "subtitle", label: "Subtitle", defaultRect: { x: 36, y: 200, w: 460, h: 35 } });
    if (s.highlights?.length) {
      (s.highlights).forEach((_, i) => {
        els.push({ id: `highlights[${i}]`, label: `Highlight ${i + 1}`, defaultRect: { x: 36, y: 245 + i * 30, w: 460, h: 26 } });
      });
    }
    if (s.metrics?.length) {
      (s.metrics).forEach((m, i) => {
        const col = i % 3, row = Math.floor(i / 3);
        const mx = 520 + col * 135, my = 90 + row * 85;
        els.push({ id: `metrics[${i}].bg`, label: `Metric ${i+1} BG`, defaultRect: { x: mx, y: my, w: 125, h: 75 }, subType: "card-bg" });
        els.push({ id: `metrics[${i}].value`, label: m.number || m.value || "0", defaultRect: { x: mx + 5, y: my + 8, w: 115, h: 35 }, subType: "metric-value" });
        els.push({ id: `metrics[${i}].label`, label: m.label, defaultRect: { x: mx + 5, y: my + 45, w: 115, h: 20 }, subType: "metric-label" });
      });
    }
    if (s.stats?.length) {
      const metricsRows = s.metrics?.length ? Math.ceil(s.metrics.length / 3) : 0;
      const statsStartY = 90 + metricsRows * 85 + 10;
      (s.stats).forEach((_, i) => {
        els.push({ id: `stats[${i}]`, label: `Stat ${i + 1}`, defaultRect: { x: 520, y: statsStartY + i * 50, w: 400, h: 44 } });
      });
    }
    return els;
  },

  "agenda": (s) => {
    const els = [
      { id: "title", label: "Title", defaultRect: { x: 36, y: 65, w: 700, h: 35 } },
      { id: "subtitle", label: "Subtitle", defaultRect: { x: 36, y: 100, w: 700, h: 22 } },
    ];
    const cards = s.cards || [];
    const cardW = cards.length <= 3 ? 270 : cards.length <= 6 ? 140 : 105;
    const gap = 10;
    cards.forEach((card, i) => {
      const col = i % 4, row = Math.floor(i / 4);
      const cx = 36 + col * (cardW + gap), cy = 135 + row * 185;
      // Card background container
      els.push({ id: `cards[${i}].bg`, label: `Card ${i+1} BG`, defaultRect: { x: cx, y: cy, w: cardW, h: 175 }, subType: "card-bg" });
      // Icon
      if (card.icon) els.push({ id: `cards[${i}].icon`, label: `${card.icon}`, defaultRect: { x: cx + 8, y: cy + 10, w: 28, h: 28 }, subType: "icon" });
      // Card title
      els.push({ id: `cards[${i}].title`, label: card.title || `Card ${i+1}`, defaultRect: { x: cx + 40, y: cy + 10, w: cardW - 50, h: 24 }, subType: "card-title" });
      // Each bullet item
      (card.items || []).forEach((item, j) => {
        els.push({ id: `cards[${i}].items[${j}]`, label: (typeof item === "string" ? item : item.text || "").substring(0, 25), defaultRect: { x: cx + 12, y: cy + 45 + j * 24, w: cardW - 24, h: 20 }, subType: "bullet" });
      });
    });
    return els;
  },

  "content": (s) => {
    const els = [
      { id: "title", label: "Title", defaultRect: { x: 36, y: 65, w: 700, h: 35 } },
      { id: "subtitle", label: "Subtitle", defaultRect: { x: 36, y: 100, w: 700, h: 22 } },
    ];
    (s.cards || []).forEach((card, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const cx = 36 + col * 450, cy = 135 + row * 200;
      const cw = 430;
      // Card background
      els.push({ id: `cards[${i}].bg`, label: `Card ${i+1} BG`, defaultRect: { x: cx, y: cy, w: cw, h: 185 }, subType: "card-bg" });
      // Icon
      if (card.icon) els.push({ id: `cards[${i}].icon`, label: `${card.icon}`, defaultRect: { x: cx + 10, y: cy + 10, w: 30, h: 30 }, subType: "icon" });
      // Card title
      els.push({ id: `cards[${i}].title`, label: card.title || `Card ${i+1}`, defaultRect: { x: cx + 46, y: cy + 10, w: cw - 160, h: 26 }, subType: "card-title" });
      // Status badge
      if (card.status) els.push({ id: `cards[${i}].status`, label: card.status, defaultRect: { x: cx + cw - 100, y: cy + 12, w: 90, h: 22 }, subType: "badge" });
      // Each item (bold+text pair or string)
      (card.items || []).forEach((item, j) => {
        const text = typeof item === "string" ? item : `${item.bold || ""} ${item.text || ""}`;
        els.push({ id: `cards[${i}].items[${j}]`, label: text.substring(0, 30), defaultRect: { x: cx + 12, y: cy + 48 + j * 28, w: cw - 24, h: 24 }, subType: "bullet" });
      });
    });
    return els;
  },

  "dashboard": (s) => {
    const els = [
      { id: "title", label: "Title", defaultRect: { x: 36, y: 65, w: 700, h: 35 } },
      { id: "subtitle", label: "Subtitle", defaultRect: { x: 36, y: 100, w: 700, h: 22 } },
    ];
    (s.metrics || []).forEach((m, i) => {
      const col = i % 3, row = Math.floor(i / 3);
      const mx = 36 + col * 300, my = 135 + row * 150;
      // Metric card background
      els.push({ id: `metrics[${i}].bg`, label: `Metric ${i+1} BG`, defaultRect: { x: mx, y: my, w: 285, h: 135 }, subType: "card-bg" });
      // Icon
      if (m.icon) els.push({ id: `metrics[${i}].icon`, label: `${m.icon}`, defaultRect: { x: mx + 120, y: my + 10, w: 45, h: 30 }, subType: "icon" });
      // Value number
      els.push({ id: `metrics[${i}].value`, label: m.value || m.number || "0", defaultRect: { x: mx + 20, y: my + 42, w: 245, h: 38 }, subType: "metric-value" });
      // Label
      els.push({ id: `metrics[${i}].label`, label: m.label, defaultRect: { x: mx + 20, y: my + 82, w: 245, h: 20 }, subType: "metric-label" });
      // Detail
      if (m.detail) els.push({ id: `metrics[${i}].detail`, label: m.detail, defaultRect: { x: mx + 20, y: my + 104, w: 245, h: 18 }, subType: "metric-detail" });
    });
    return els;
  },

  "table": (s) => {
    const els = [
      { id: "title", label: "Title", defaultRect: { x: 36, y: 65, w: 700, h: 35 } },
      { id: "subtitle", label: "Subtitle", defaultRect: { x: 36, y: 100, w: 700, h: 22 } },
      { id: "table", label: "Table", defaultRect: { x: 36, y: 135, w: 888, h: 40 + (s.rows?.length || 0) * 36 } },
    ];
    if (s.summaryStats?.length) {
      (s.summaryStats).forEach((st, i) => {
        els.push({ id: `summaryStats[${i}]`, label: st.label || `Summary ${i + 1}`, defaultRect: { x: 36 + i * 160, y: 410, w: 145, h: 50 } });
      });
    }
    return els;
  },

  "section-overview": (s) => {
    const els = [
      { id: "section-icon", label: "Section Icon", defaultRect: { x: 36, y: 80, w: 56, h: 56 } },
      { id: "section-title", label: "Section Title", defaultRect: { x: 100, y: 80, w: 380, h: 50 } },
      { id: "description", label: "Description", defaultRect: { x: 36, y: 145, w: 440, h: 80 } },
    ];
    if (s.stats?.length) {
      (s.stats).forEach((st, i) => {
        els.push({ id: `stats[${i}]`, label: st.label || `Stat ${i + 1}`, defaultRect: { x: 36 + i * 110, y: 240, w: 100, h: 75 } });
      });
    }
    if (s.billingMetrics?.length) {
      (s.billingMetrics).forEach((bm, i) => {
        els.push({ id: `billingMetrics[${i}]`, label: bm.label || `Billing ${i + 1}`, defaultRect: { x: 520, y: 80 + i * 45, w: 400, h: 40 } });
      });
    }
    return els;
  },

  "highlight-list": (s) => {
    const els = [
      { id: "title", label: "Title", defaultRect: { x: 36, y: 65, w: 700, h: 35 } },
      { id: "subtitle", label: "Subtitle", defaultRect: { x: 36, y: 100, w: 700, h: 22 } },
    ];
    (s.items || []).forEach((item, i) => {
      els.push({ id: `items[${i}]`, label: item.text?.substring(0, 30) || `Item ${i + 1}`, defaultRect: { x: 36, y: 135 + i * 52, w: 888, h: 44 } });
    });
    return els;
  },

  "thank-you": (s) => [
    { id: "icon", label: "Heart Icon", defaultRect: { x: 430, y: 80, w: 100, h: 100 } },
    { id: "title", label: "Title", defaultRect: { x: 130, y: 200, w: 700, h: 70 } },
    { id: "message", label: "Message", defaultRect: { x: 200, y: 280, w: 560, h: 80 } },
    { id: "signature", label: "Signature", defaultRect: { x: 200, y: 375, w: 560, h: 30 } },
  ],
};

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 🎨  DESIGN TOKENS — Reactive Theme System
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━���━━━

const COLOR_COMBOS = {
  blue:   { accent: "#0078d4", accentDark: "#005a9e", accentLight: "#e8f4fd", accentGlow: "rgba(0,120,212,0.15)", label: "Blue" },
  purple: { accent: "#8b5cf6", accentDark: "#7c3aed", accentLight: "#ede9fe", accentGlow: "rgba(139,92,246,0.15)", label: "Purple" },
  teal:   { accent: "#0d9488", accentDark: "#0f766e", accentLight: "#ccfbf1", accentGlow: "rgba(13,148,136,0.15)", label: "Teal" },
  rose:   { accent: "#e11d48", accentDark: "#be123c", accentLight: "#ffe4e6", accentGlow: "rgba(225,29,72,0.15)", label: "Rose" },
  amber:  { accent: "#d97706", accentDark: "#b45309", accentLight: "#fef3c7", accentGlow: "rgba(217,119,6,0.15)", label: "Amber" },
};

// Miracle Software Systems logos (served from /public)
const MIRACLE_LOGO_SMALL = "/miracle-logo-small.png";  // Logo 1: compact version for title slide
const MIRACLE_LOGO = "/miracle-logo-large.png";         // Logo 2: full version for footer

function getTheme(mode, combo) {
  const c = COLOR_COMBOS[combo] || COLOR_COMBOS.blue;
  const shared = {
    accent: c.accent,
    accentDark: c.accentDark,
    accentLight: c.accentLight,
    accentGlow: c.accentGlow,
    success: "#22c55e",
    danger: "#ef4444",
    dangerBg: "rgba(239,68,68,0.08)",
    purple: "#8b5cf6",
    radius: "10px",
    radiusSm: "7px",
    radiusLg: "14px",
    font: "'Outfit', 'DM Sans', sans-serif",
    fontMono: "'JetBrains Mono', monospace",
    mode,
  };

  if (mode === "light") {
    return {
      ...shared,
      bg: "#f0f2f5",
      surface: "#ffffff",
      surfaceHover: "#f1f5f9",
      border: "#e5e7eb",
      borderHover: "#d1d5db",
      text: "#111827",
      textMuted: "#6b7280",
      textDim: "#9ca3af",
      warning: "#f59e0b",
      slideBg: `linear-gradient(145deg, #f8fafc 0%, #ffffff 50%, #f0f2f5 100%)`,
      glow: `radial-gradient(ellipse at 70% 30%, ${c.accentGlow.replace("0.15", "0.06")} 0%, transparent 60%)`,
      slideHeaderBg: c.accent,
      slideHeaderText: "#ffffff",
      cardBg: "#ffffff",
      cardBorder: "#e5e7eb",
      slideText: "#111827",
      slideTextSec: "#6b7280",
      slideTextDim: "#9ca3af",
      footerBg: "#ffffff",
      footerBorder: "#e5e7eb",
      slideAccentBg: c.accentLight,
    };
  }

  // dark mode — improved contrast and visual hierarchy
  return {
    ...shared,
    bg: "#06090f",
    surface: "#0c1220",
    surfaceHover: "#111827",
    border: "rgba(148,163,194,0.12)",
    borderHover: "rgba(148,163,194,0.22)",
    text: "#e2e8f0",
    textMuted: "#64748b",
    textDim: "#475569",
    warning: "#eab308",
    slideBg: "linear-gradient(145deg, #070b14 0%, #0f1729 50%, #0a1020 100%)",
    glow: `radial-gradient(ellipse at 70% 30%, ${c.accentGlow.replace("0.15", "0.08")} 0%, transparent 60%)`,
    slideHeaderBg: "rgba(255,255,255,0.06)",
    slideHeaderText: "#f1f5f9",
    cardBg: "#111d32",
    cardBorder: "rgba(148,163,194,0.15)",
    slideText: "#f1f5f9",
    slideTextSec: "#cbd5e1",
    slideTextDim: "#94a3b8",
    footerBg: "rgba(255,255,255,0.04)",
    footerBorder: "rgba(148,163,194,0.12)",
    slideAccentBg: `rgba(${parseInt(c.accent.slice(1,3),16)},${parseInt(c.accent.slice(3,5),16)},${parseInt(c.accent.slice(5,7),16)},0.15)`,
  };
}

function getCss(T) {
  return {
    input: {
      width: "100%", padding: "9px 13px", borderRadius: T.radiusSm,
      border: `1px solid ${T.border}`, background: T.mode === "light" ? "#f8fafc" : "rgba(255,255,255,0.02)",
      color: T.text, fontSize: "13px", outline: "none", boxSizing: "border-box",
      fontFamily: T.font, transition: "border-color 0.2s, box-shadow 0.2s",
    },
    label: {
      color: T.textMuted, fontSize: "10px", fontWeight: 700,
      textTransform: "uppercase", letterSpacing: "1.2px",
      marginBottom: "6px", display: "block", fontFamily: T.font,
    },
    btnIcon: {
      background: "transparent", border: `1px solid ${T.border}`,
      color: T.textMuted, width: "30px", height: "30px", borderRadius: T.radiusSm,
      cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center",
      fontSize: "13px", transition: "all 0.15s",
    },
    btnPrimary: {
      background: `linear-gradient(135deg, ${T.accent}, ${T.accentDark})`,
      border: "none", color: "#fff", padding: "8px 18px", borderRadius: T.radiusSm,
      cursor: "pointer", fontSize: "12px", fontWeight: 700, fontFamily: T.font,
      letterSpacing: "0.4px", transition: "transform 0.15s, box-shadow 0.15s",
    },
    btnGhost: {
      background: "transparent", border: `1px dashed ${T.border}`,
      color: T.accent, padding: "7px 14px", borderRadius: T.radiusSm,
      cursor: "pointer", fontSize: "11px", fontWeight: 600, width: "100%",
      fontFamily: T.font, transition: "all 0.15s",
    },
    btnDanger: {
      background: T.dangerBg, border: "none", color: T.danger,
      borderRadius: T.radiusSm, padding: "5px 9px", cursor: "pointer",
      fontSize: "11px", transition: "all 0.15s",
    },
    card: {
      background: T.surface, border: `1px solid ${T.border}`,
      borderRadius: T.radius, padding: "12px", marginBottom: "8px",
    },
  };
}

function useTheme() {
  const mode = useStore((s) => s.themeMode);
  const combo = useStore((s) => s.colorCombo);
  return useMemo(() => {
    const T = getTheme(mode, combo);
    const css = getCss(T);
    return { T, css };
  }, [mode, combo]);
}

// Legacy global references — used by non-component code (export helpers, etc.)
// These default to dark+blue for backward compatibility in non-React contexts
let T = getTheme("dark", "blue");
let css = getCss(T);

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 🧱  COMMON UI PRIMITIVES
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function Input({ value, onChange, placeholder, style: sx, ...rest }) {
  const { T, css } = useTheme();
  const [focused, setFocused] = useState(false);
  return (
    <input {...rest} style={{ ...css.input, ...sx, borderColor: focused ? T.accent : T.border, boxShadow: focused ? `0 0 0 3px ${T.accentGlow}` : "none" }}
      value={value || ""} placeholder={placeholder} onChange={(e) => onChange(e.target.value)}
      onFocus={() => setFocused(true)} onBlur={() => setFocused(false)} />
  );
}

function TextArea({ value, onChange, placeholder, rows = 3, style: sx }) {
  const { T, css } = useTheme();
  const [focused, setFocused] = useState(false);
  return (
    <textarea style={{ ...css.input, ...sx, resize: "vertical", minHeight: `${rows * 22}px`, borderColor: focused ? T.accent : T.border, boxShadow: focused ? `0 0 0 3px ${T.accentGlow}` : "none" }}
      value={value || ""} placeholder={placeholder} onChange={(e) => onChange(e.target.value)}
      onFocus={() => setFocused(true)} onBlur={() => setFocused(false)} />
  );
}

function Select({ value, onChange, options }) {
  const { css } = useTheme();
  return (
    <select style={{ ...css.input, cursor: "pointer", appearance: "none", paddingRight: "28px",
      backgroundImage: `url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' fill='%2364748b' viewBox='0 0 16 16'%3E%3Cpath d='M8 11L3 6h10z'/%3E%3C/svg%3E")`,
      backgroundRepeat: "no-repeat", backgroundPosition: "right 10px center" }}
      value={value || options[0]} onChange={(e) => onChange(e.target.value)}>
      {options.map((o) => <option key={o} value={o}>{o}</option>)}
    </select>
  );
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 🧩  EDITOR — Dynamic form engine
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

// Field description tooltips for help text
const FIELD_HELP = {
  title: "Main heading displayed prominently on the slide",
  subtitle: "Supporting text below the title",
  badge: "Period label (e.g., Q1 2026) shown as a colored badge",
  metrics: "Key numbers displayed in card format",
  highlights: "Colored bullet points for key takeaways",
  stats: "Info rows with icon, label, and value",
  cards: "Content sections with icon, title, and bullet points",
  columns: "Comma-separated column headers for the data table",
  rows: "Table data rows — one row per entry",
  summaryStats: "Summary numbers displayed below the table",
  items: "List items with optional color coding",
  message: "Closing message or paragraph text",
  signature: "Name/team shown at the bottom",
  sectionIcon: "Emoji icon for the section (e.g., 👥, 📊)",
  sectionTitle: "Section heading displayed prominently",
  description: "Detailed description text for the section",
  billingMetrics: "Key-value pairs for billing/financial data",
};

function FieldRenderer({ field, value, onChange }) {
  const { T, css } = useTheme();
  const helpText = FIELD_HELP[field.name];
  const fieldLabel = (
    <label style={css.label} title={helpText}>
      {field.label}
      {helpText && <span style={{ color: T.textDim, fontWeight: 400, fontSize: "9px", textTransform: "none", letterSpacing: 0, marginLeft: "6px" }}>{helpText}</span>}
    </label>
  );
  switch (field.type) {
    case "text":
      return <div style={{ marginBottom: "14px" }}>{fieldLabel}<Input value={value} onChange={onChange} placeholder={field.placeholder} /></div>;
    case "textarea":
      return <div style={{ marginBottom: "14px" }}>{fieldLabel}<TextArea value={value} onChange={onChange} placeholder={field.placeholder} /></div>;
    case "csv":
      return <div style={{ marginBottom: "14px" }}>{fieldLabel}<Input value={(value || []).join(", ")} onChange={(v) => onChange(v.split(",").map((s) => s.trim()).filter(Boolean))} /></div>;
    case "color":
      return <div style={{ marginBottom: "14px" }}><label style={css.label}>{field.label}</label><input type="color" value={value || "#3b82f6"} onChange={(e) => onChange(e.target.value)} style={{ width: "40px", height: "34px", border: "none", borderRadius: T.radiusSm, cursor: "pointer", padding: 0, background: "transparent" }} /></div>;
    case "select":
      return <div style={{ marginBottom: "14px" }}><label style={css.label}>{field.label}</label><Select value={value} onChange={onChange} options={field.options} /></div>;
    case "array":
      return <ArrayFieldRenderer field={field} value={value || []} onChange={onChange} />;
    case "compound-array":
      return <CompoundArrayRenderer field={field} value={value || []} onChange={onChange} />;
    case "table-rows":
      return <TableRowsRenderer field={field} value={value || []} onChange={onChange} />;
    default:
      return null;
  }
}

function ArrayFieldRenderer({ field, value, onChange }) {
  const { T, css } = useTheme();
  const items = value || [];
  const update = (i, key, val) => { const next = [...items]; next[i] = { ...next[i], [key]: val }; onChange(next); };
  const add = () => onChange([...items, field.defaultItem || Object.fromEntries(field.itemFields.map((f) => [f.name, ""]))]);
  const remove = (i) => onChange(items.filter((_, j) => j !== i));
  return (
    <div style={{ marginBottom: "14px" }}>
      <label style={css.label}>{field.label}</label>
      {items.map((item, i) => (
        <div key={i} style={{ ...css.card, display: "flex", gap: "6px", alignItems: "center", padding: "8px 10px" }}>
          {field.itemFields.map((f) => (
            <div key={f.name} style={{ flex: f.width === "auto" ? 1 : "none", width: f.width !== "auto" ? f.width : undefined }}>
              {f.type === "color" ? (
                <input type="color" value={item[f.name] || "#3b82f6"} onChange={(e) => update(i, f.name, e.target.value)}
                  style={{ width: "34px", height: "34px", border: "none", borderRadius: "6px", cursor: "pointer", padding: 0, background: "transparent" }} />
              ) : (
                <Input value={item[f.name]} onChange={(v) => update(i, f.name, v)} placeholder={f.label} style={{ fontSize: "12px", padding: "7px 10px" }} />
              )}
            </div>
          ))}
          <button onClick={() => remove(i)} style={css.btnDanger}>✕</button>
        </div>
      ))}
      <button onClick={add} style={css.btnGhost}>+ {field.addLabel || "Add Item"}</button>
    </div>
  );
}

function CompoundArrayRenderer({ field, value, onChange }) {
  const { T, css } = useTheme();
  const items = value || [];
  const update = (i, key, val) => { const next = [...items]; next[i] = { ...next[i], [key]: val }; onChange(next); };
  const add = () => onChange([...items, field.defaultItem || { title: "New" }]);
  const remove = (i) => onChange(items.filter((_, j) => j !== i));
  return (
    <div style={{ marginBottom: "14px" }}>
      <label style={css.label}>{field.label}</label>
      {items.map((item, i) => (
        <div key={i} style={{ ...css.card, padding: "10px 12px" }}>
          <div style={{ display: "flex", gap: "6px", alignItems: "center", marginBottom: "8px" }}>
            {field.itemFields.map((f) => (
              <div key={f.name} style={{ flex: f.width ? "none" : 1, width: f.width || undefined }}>
                {f.type === "select" ? <Select value={item[f.name]} onChange={(v) => update(i, f.name, v)} options={f.options} /> :
                  <Input value={item[f.name]} onChange={(v) => update(i, f.name, v)} placeholder={f.placeholder || f.label} style={{ fontSize: "12px", padding: "7px 10px" }} />}
              </div>
            ))}
            <button onClick={() => remove(i)} style={css.btnDanger}>✕</button>
          </div>
          {field.subArray && <SubArrayRenderer config={field.subArray} value={item[field.subArray.name] || []} onChange={(v) => update(i, field.subArray.name, v)} />}
        </div>
      ))}
      <button onClick={add} style={css.btnGhost}>+ {field.addLabel}</button>
    </div>
  );
}

function SubArrayRenderer({ config, value, onChange }) {
  const { T, css } = useTheme();
  if (config.type === "string-array") {
    return (
      <div style={{ paddingLeft: "8px", borderLeft: `2px solid ${T.border}` }}>
        {(value || []).map((item, i) => (
          <div key={i} style={{ display: "flex", gap: "4px", marginBottom: "4px" }}>
            <Input value={item} onChange={(v) => { const n = [...value]; n[i] = v; onChange(n); }} style={{ fontSize: "11px", padding: "5px 8px" }} placeholder={`Point ${i + 1}`} />
            <button onClick={() => onChange(value.filter((_, j) => j !== i))} style={{ ...css.btnDanger, padding: "3px 7px", fontSize: "10px" }}>✕</button>
          </div>
        ))}
        <button onClick={() => onChange([...(value || []), ""])} style={{ ...css.btnGhost, fontSize: "10px", padding: "4px 10px" }}>+ {config.addLabel}</button>
      </div>
    );
  }
  if (config.type === "kv-array") {
    return (
      <div style={{ paddingLeft: "8px", borderLeft: `2px solid ${T.border}` }}>
        {(value || []).map((item, i) => (
          <div key={i} style={{ display: "flex", gap: "4px", marginBottom: "4px" }}>
            <Input value={item.bold} onChange={(v) => { const n = [...value]; n[i] = { ...n[i], bold: v }; onChange(n); }} style={{ fontSize: "11px", padding: "5px 8px", width: "100px", flex: "none" }} placeholder={config.keyPlaceholder} />
            <Input value={item.text} onChange={(v) => { const n = [...value]; n[i] = { ...n[i], text: v }; onChange(n); }} style={{ fontSize: "11px", padding: "5px 8px" }} placeholder={config.valuePlaceholder} />
            <button onClick={() => onChange(value.filter((_, j) => j !== i))} style={{ ...css.btnDanger, padding: "3px 7px", fontSize: "10px" }}>✕</button>
          </div>
        ))}
        <button onClick={() => onChange([...(value || []), { bold: "", text: "" }])} style={{ ...css.btnGhost, fontSize: "10px", padding: "4px 10px" }}>+ {config.addLabel}</button>
      </div>
    );
  }
  return null;
}

function TableRowsRenderer({ field, value, onChange }) {
  const { T, css } = useTheme();
  const columns = useStore((s) => s.slides[s.activeSlide]?.columns || []);
  const add = () => onChange([...(value || []), { cells: Array(columns.length).fill("") }]);
  return (
    <div style={{ marginBottom: "14px" }}>
      <label style={css.label}>{field.label || "Rows"}</label>
      {(value || []).map((row, i) => (
        <div key={i} style={{ display: "flex", gap: "4px", marginBottom: "4px", alignItems: "center" }}>
          {(row.cells || []).map((cell, j) => (
            <Input key={j} value={cell} placeholder={columns[j] || `Col ${j + 1}`}
              onChange={(v) => { const rows = [...value]; const cells = [...rows[i].cells]; cells[j] = v; rows[i] = { ...rows[i], cells }; onChange(rows); }}
              style={{ fontSize: "11px", padding: "6px 8px", flex: 1 }} />
          ))}
          <button onClick={() => onChange(value.filter((_, j) => j !== i))} style={css.btnDanger}>✕</button>
        </div>
      ))}
      <button onClick={add} style={css.btnGhost}>+ {field.addLabel || "Add Row"}</button>
    </div>
  );
}

// Map canvas element ID → form field names it corresponds to
function canvasIdToFieldNames(canvasId) {
  if (!canvasId) return [];
  // Direct matches
  if (["title", "subtitle", "badge", "metrics", "highlights", "stats", "message", "signature", "description", "sectionTitle", "sectionIcon"].includes(canvasId)) return [canvasId];
  // page-title maps to title+subtitle
  if (canvasId === "page-title") return ["title", "subtitle"];
  // Array items map to the parent array field
  const arrMatch = canvasId.match(/^(\w+)\[\d+\]$/);
  if (arrMatch) return [arrMatch[1]];
  // Section overview specifics
  if (canvasId === "section-title") return ["sectionTitle"];
  if (canvasId === "section-icon") return ["sectionIcon"];
  if (canvasId === "billing") return ["billingMetrics"];
  if (canvasId === "table") return ["columns", "rows"];
  if (canvasId === "summary") return ["summaryStats"];
  if (canvasId === "icon") return [];
  return [canvasId];
}

function SlideForm({ slide }) {
  const { T, css } = useTheme();
  const schema = SLIDE_SCHEMA[slide.type] || SLIDE_SCHEMA.content;
  const { updateSlide } = useStoreActions();
  const canvasSelectedId = useStore((s) => s.canvasSelectedId);
  const highlightedFields = useMemo(() => canvasIdToFieldNames(canvasSelectedId), [canvasSelectedId]);
  const highlightRef = useRef(null);

  // Auto-scroll highlighted field into view
  useEffect(() => {
    if (highlightRef.current) {
      highlightRef.current.scrollIntoView({ behavior: "smooth", block: "nearest" });
    }
  }, [canvasSelectedId]);

  return (
    <div>
      {schema.fields.map((field) => {
        const isHighlighted = highlightedFields.includes(field.name);
        return (
          <div key={field.name} ref={isHighlighted ? highlightRef : undefined}
            style={{
              borderRadius: T.radiusSm,
              border: isHighlighted ? `2px solid ${T.accent}` : "2px solid transparent",
              background: isHighlighted ? T.accentGlow : "transparent",
              padding: isHighlighted ? "6px" : "0",
              marginBottom: isHighlighted ? "4px" : "0",
              transition: "all 0.2s ease",
            }}>
            <FieldRenderer field={field} value={slide[field.name]}
              onChange={(val) => updateSlide(slide.id, { [field.name]: val })} />
          </div>
        );
      })}
    </div>
  );
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 🎥  PREVIEW — Step 3: Modular SlideFrame / SlideHeader / SlideFooter
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/*
  Architecture (from Step 3 spec):
  ┌──────────────────────────────────────┐
  │  SlideFrame                          │
  │  ┌──────────────────────────────┐    │
  │  │  SlideHeader (logo + meta)   │    │
  │  ├──────────────────────────────┤    │
  │  │  SlideBody (children)        │    │
  │  ├──────────────────────────────┤    │
  │  │  SlideFooter (report + page) │    │
  │  └──────────────────────────────┘    │
  └──────────────────────────────────────┘
*/

// SLIDE_BG and GLOW are now part of the theme (T.slideBg, T.glow)

// ─── SlideFrame.jsx ────────────────────────────────────────────
function SlideFrame({ children, compact, noChrome }) {
  const { T } = useTheme();
  const isLight = T.mode === "light";
  return (
    <div style={{
      height: "100%", width: "100%", background: T.slideBg,
      position: "relative", overflow: "hidden", fontFamily: T.font,
    }}>
      <div style={{ position: "absolute", inset: 0, background: T.glow, pointerEvents: "none" }} />
      {/* Decorative corner accents */}
      <div style={{ position: "absolute", top: 0, left: 0, width: compact ? "30px" : "80px", height: "3px", background: `linear-gradient(90deg, ${T.accent}, transparent)`, opacity: isLight ? 0.6 : 0.4 }} />
      <div style={{ position: "absolute", top: 0, left: 0, width: "3px", height: compact ? "30px" : "80px", background: `linear-gradient(180deg, ${T.accent}, transparent)`, opacity: isLight ? 0.6 : 0.4 }} />
      <div style={{ position: "absolute", bottom: 0, right: 0, width: compact ? "30px" : "80px", height: "3px", background: `linear-gradient(270deg, ${T.purple}, transparent)`, opacity: 0.3 }} />
      <div style={{ position: "absolute", bottom: 0, right: 0, width: "3px", height: compact ? "30px" : "80px", background: `linear-gradient(0deg, ${T.purple}, transparent)`, opacity: 0.3 }} />
      <div style={{
        position: "relative", height: "100%",
        padding: compact ? "6px 10px" : "0",
        display: "flex", flexDirection: "column",
      }}>
        {children}
      </div>
    </div>
  );
}

// ─── SlideHeader.jsx ───────────────────────────────────────────
function SlideHeaderBar({ compact, period, presenter }) {
  const { T } = useTheme();
  const isLight = T.mode === "light";

  if (compact) {
    return (
      <div style={{
        display: "flex", justifyContent: "space-between", alignItems: "center",
        padding: isLight ? "3px 4px" : "0 2px", marginBottom: "4px", flexShrink: 0,
        background: isLight ? T.accent : "transparent",
        borderRadius: isLight ? "2px 2px 0 0" : "0",
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: "3px" }}>
          <div style={{
            width: "10px", height: "10px", borderRadius: "3px",
            background: isLight ? "white" : `linear-gradient(135deg, ${T.accent}, ${T.purple})`,
            display: "flex", alignItems: "center", justifyContent: "center",
            fontSize: "5px", color: isLight ? T.accent : "#fff", fontWeight: 900,
          }}>{isLight ? <span style={{ fontSize: "5px" }}>M</span> : "M"}</div>
          <span style={{ fontSize: "4px", fontWeight: 700, color: isLight ? "rgba(255,255,255,0.9)" : "rgba(255,255,255,0.5)" }}>Microsoft Practice</span>
        </div>
        <div style={{ display: "flex", gap: "4px" }}>
          {period && <span style={{ fontSize: "3.5px", color: isLight ? "rgba(255,255,255,0.7)" : T.textDim, background: isLight ? "rgba(255,255,255,0.15)" : "rgba(255,255,255,0.04)", padding: "1px 4px", borderRadius: "3px" }}>{period}</span>}
        </div>
      </div>
    );
  }

  if (isLight) {
    // Light mode: solid accent color header bar (matching HTML reference)
    return (
      <div style={{
        display: "flex", justifyContent: "space-between", alignItems: "center",
        padding: "0 32px", flexShrink: 0, height: "54px",
        background: T.accent,
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
          <div style={{
            width: "28px", height: "28px", borderRadius: "5px", background: "white",
            display: "flex", alignItems: "center", justifyContent: "center",
            fontSize: "14px", color: T.accent, fontWeight: 900,
          }}>M</div>
          <span style={{ color: "white", fontSize: "15px", fontWeight: 700, fontFamily: T.font }}>Microsoft Practice</span>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: "20px", color: "rgba(255,255,255,0.85)", fontSize: "13px" }}>
          {period && <span style={{ display: "flex", alignItems: "center", gap: "6px" }}><span style={{ fontSize: "11px" }}>📅</span> {period}</span>}
          <span style={{ display: "flex", alignItems: "center", gap: "6px" }}><span style={{ fontSize: "11px" }}>📄</span> Quarterly Report</span>
        </div>
      </div>
    );
  }

  // Dark mode: existing design
  return (
    <div style={{
      display: "flex", justifyContent: "space-between", alignItems: "center",
      padding: "16px 36px", flexShrink: 0,
      borderBottom: `1px solid ${T.footerBorder}`,
    }}>
      <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
        <div style={{
          width: "36px", height: "36px", borderRadius: "10px",
          background: `linear-gradient(135deg, ${T.accent}, ${T.purple})`,
          display: "flex", alignItems: "center", justifyContent: "center",
          fontSize: "14px", color: "#fff", fontWeight: 900,
          boxShadow: `0 4px 16px ${T.accentGlow}`,
        }}>M</div>
        <div>
          <div style={{ fontSize: "14px", fontWeight: 800, color: T.slideHeaderText, letterSpacing: "-0.01em" }}>Microsoft Practice</div>
          <div style={{ fontSize: "10px", color: T.textDim, marginTop: "1px" }}>Quarterly Report</div>
        </div>
      </div>
      <div style={{ display: "flex", gap: "10px", alignItems: "center" }}>
        {period && (
          <span style={{
            fontSize: "11px", color: T.textMuted, fontWeight: 600,
            background: "rgba(255,255,255,0.08)", padding: "5px 14px",
            borderRadius: "8px", border: `1px solid ${T.cardBorder}`,
            display: "flex", alignItems: "center", gap: "6px",
          }}>
            <span style={{ fontSize: "10px" }}>📅</span> {period}
          </span>
        )}
        {presenter && (
          <span style={{
            fontSize: "11px", color: T.textMuted, fontWeight: 600,
            background: "rgba(255,255,255,0.08)", padding: "5px 14px",
            borderRadius: "8px", border: `1px solid ${T.cardBorder}`,
            display: "flex", alignItems: "center", gap: "6px",
          }}>
            <span style={{ fontSize: "10px" }}>👤</span> {presenter}
          </span>
        )}
      </div>
    </div>
  );
}

// ─── SlideFooter.jsx ───────────────────────────────────────────
function SlideFooterBar({ compact, index, total }) {
  const { T } = useTheme();
  const isLight = T.mode === "light";

  if (compact) {
    return (
      <div style={{
        display: "flex", justifyContent: "space-between", alignItems: "center",
        padding: "2px 2px 0", marginTop: "auto", flexShrink: 0,
      }}>
        <span style={{ fontSize: "3.5px", color: isLight ? T.slideTextDim : "rgba(255,255,255,0.2)", display: "flex", alignItems: "center", gap: "2px" }}>
          <img src={MIRACLE_LOGO} style={{ height: "5px", objectFit: "contain" }} alt="" />
        </span>
        <span style={{ fontSize: "3.5px", color: isLight ? T.slideTextDim : "rgba(255,255,255,0.25)", fontFamily: T.fontMono }}>
          {index + 1} / {total}
        </span>
      </div>
    );
  }

  return (
    <div style={{
      display: "flex", justifyContent: "space-between", alignItems: "center",
      padding: isLight ? "0 32px" : "14px 36px", flexShrink: 0,
      height: isLight ? "38px" : "auto",
      borderTop: `1px solid ${T.footerBorder}`,
      marginTop: "auto",
      background: T.footerBg,
    }}>
      <div style={{ display: "flex", alignItems: "center", gap: "8px", fontSize: "12px" }}>
        <img src={MIRACLE_LOGO} style={{ height: "24px", objectFit: "contain" }} alt="Miracle Software Systems" />
      </div>
      <div style={{ display: "flex", alignItems: "center", gap: "16px", fontSize: "12px", color: T.slideTextDim }}>
        <span style={{ display: "flex", alignItems: "center", gap: "5px" }}>
          <span style={{ color: T.accent, fontSize: "11px" }}>👤</span> Presented by: Leadership
        </span>
        <div style={{
          background: isLight ? T.bg : "rgba(255,255,255,0.04)", borderRadius: "6px",
          padding: "3px 12px", fontSize: "11px", fontFamily: T.fontMono,
          color: T.slideTextSec, fontWeight: 600,
        }}>
          {index + 1} <span style={{ opacity: 0.4 }}>/</span> {total}
        </div>
      </div>
    </div>
  );
}

// ─── PageTitle (reusable for non-title slides) ─────────────────
function PageTitle({ title, subtitle, compact, badge }) {
  const { T } = useTheme();
  return (
    <div style={{
      marginBottom: compact ? "6px" : "22px", flexShrink: 0,
      padding: compact ? "0" : "0 36px",
    }}>
      <div style={{ display: "flex", alignItems: "center", gap: compact ? "4px" : "12px" }}>
        <h2 style={{
          color: T.slideText, fontSize: compact ? "11px" : "26px",
          fontWeight: 800, margin: 0, letterSpacing: "-0.025em", lineHeight: 1.15,
        }}>{title}</h2>
        {badge && (
          <span style={{
            background: T.accentGlow, color: T.accentLight,
            fontSize: compact ? "4px" : "10px", padding: compact ? "1px 4px" : "3px 10px",
            borderRadius: "8px", fontWeight: 700,
          }}>{badge}</span>
        )}
      </div>
      {subtitle && (
        <p style={{
          color: T.textDim, fontSize: compact ? "5px" : "13px",
          margin: compact ? "2px 0 0" : "4px 0 0", fontWeight: 400,
        }}>{subtitle}</p>
      )}
    </div>
  );
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 🎯  SLIDE TYPE COMPONENTS (Step 3 pattern: Frame→Header→Body→Footer)
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

// ─── TitleSlide.jsx ────────────────────────────────────────────
function TitleSlide({ data: s, compact, index, total }) {
  const { T } = useTheme();
  return (
    <SlideFrame compact={compact}>
      <SlideHeaderBar compact={compact} period={s.badge} presenter="Leadership" />

      {/* slide-body: Title layout — split left/right */}
      <div style={{
        flex: 1, display: "flex",
        padding: compact ? "0" : "0 36px",
        gap: compact ? "8px" : "32px",
        alignItems: "center",
      }}>
        {/* LEFT — Title content */}
        <div style={{ flex: 1.2, display: "flex", flexDirection: "column", justifyContent: "center" }}>
          {s.badge && (
            <div style={{
              background: `linear-gradient(135deg, ${T.accent}, #2563eb)`, color: "#fff",
              padding: compact ? "2px 8px" : "6px 18px", borderRadius: "16px",
              fontSize: compact ? "5px" : "11px", fontWeight: 700,
              letterSpacing: "0.8px", textTransform: "uppercase",
              alignSelf: "flex-start", marginBottom: compact ? "6px" : "16px",
            }}>{s.badge}</div>
          )}
          <h1 style={{
            color: T.slideText, fontSize: compact ? "13px" : "34px", fontWeight: 800,
            margin: 0, lineHeight: 1.15, letterSpacing: "-0.03em",
          }}>{s.title}</h1>
          <p style={{
            color: T.textDim, fontSize: compact ? "5px" : "14px",
            marginTop: compact ? "3px" : "10px", lineHeight: 1.5,
          }}>{s.subtitle}</p>

          {/* Highlights */}
          {s.highlights?.length > 0 && (
            <div style={{
              marginTop: compact ? "6px" : "20px",
              display: "flex", flexDirection: "column", gap: compact ? "2px" : "7px",
            }}>
              {s.highlights.map((h, i) => (
                <div key={i} style={{
                  display: "flex", alignItems: "center", gap: compact ? "4px" : "10px",
                  fontSize: compact ? "4.5px" : "12px", color: h.color || T.success, fontWeight: 600,
                }}>
                  <span style={{
                    width: compact ? "4px" : "7px", height: compact ? "4px" : "7px",
                    borderRadius: "50%", background: h.color || T.success, flexShrink: 0,
                  }} />
                  {h.text}
                </div>
              ))}
            </div>
          )}
        </div>

        {/* RIGHT — Metrics + Stats */}
        <div style={{ flex: 0.8, display: "flex", flexDirection: "column", gap: compact ? "4px" : "12px" }}>
          {/* Metric strip */}
          {s.metrics?.length > 0 && (
            <div style={{ display: "flex", gap: compact ? "4px" : "10px", flexWrap: "wrap" }}>
              {s.metrics.map((m, i) => (
                <div key={i} style={{
                  background: T.cardBg, border: `1px solid ${T.cardBorder}`,
                  borderRadius: compact ? "6px" : "12px",
                  padding: compact ? "5px 8px" : "14px 22px",
                  textAlign: "center", flex: 1, minWidth: compact ? "30px" : "70px",
                }}>
                  <div style={{ color: T.accent, fontSize: compact ? "10px" : "24px", fontWeight: 800, lineHeight: 1 }}>{m.number}</div>
                  <div style={{ color: T.textDim, fontSize: compact ? "3.5px" : "10px", marginTop: compact ? "2px" : "5px", textTransform: "uppercase", letterSpacing: "0.5px" }}>{m.label}</div>
                </div>
              ))}
            </div>
          )}

          {/* Info stat rows */}
          {s.stats?.length > 0 && (
            <div style={{ display: "flex", flexDirection: "column", gap: compact ? "3px" : "8px" }}>
              {s.stats.map((st, i) => (
                <div key={i} style={{
                  background: T.cardBg, border: `1px solid ${T.cardBorder}`,
                  borderRadius: compact ? "6px" : "10px",
                  padding: compact ? "4px 8px" : "10px 16px",
                  display: "flex", alignItems: "center", gap: compact ? "5px" : "12px",
                }}>
                  <span style={{
                    fontSize: compact ? "8px" : "18px",
                    width: compact ? "14px" : "32px", height: compact ? "14px" : "32px",
                    borderRadius: "8px", background: T.accentGlow,
                    display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0,
                  }}>{st.icon}</span>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontSize: compact ? "3.5px" : "10px", color: T.textDim, fontWeight: 500 }}>{st.label}</div>
                    <div style={{ fontSize: compact ? "5px" : "13px", color: T.slideTextSec, fontWeight: 700 }}>{st.value}</div>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>

      {/* Miracle Logo — bottom left */}
      {!compact && (
        <div style={{ padding: "0 36px 8px", flexShrink: 0 }}>
          <img src={MIRACLE_LOGO_SMALL} alt="Miracle Software Systems" style={{ height: "36px", objectFit: "contain" }} />
        </div>
      )}

      <SlideFooterBar compact={compact} index={index} total={total} />
    </SlideFrame>
  );
}

// ─── AgendaSlide.jsx ───────────────────────────────────────────
function AgendaSlide({ data: s, compact, index, total }) {
  const { T } = useTheme();
  const cards = s.cards || [];
  return (
    <SlideFrame compact={compact}>
      <SlideHeaderBar compact={compact} period="Q1 2026" />

      <PageTitle title={s.title} subtitle={s.subtitle} compact={compact} />

      {/* slide-body: Agenda grid */}
      <div style={{
        display: "grid",
        gridTemplateColumns: `repeat(${Math.min(cards.length, 4)}, 1fr)`,
        gap: compact ? "4px" : "14px", flex: 1, alignContent: "start",
        padding: compact ? "0" : "0 36px",
      }}>
        {cards.map((card, i) => (
          <div key={i} style={{
            background: T.cardBg, border: `1px solid ${T.cardBorder}`,
            borderRadius: compact ? "5px" : T.radiusLg,
            padding: compact ? "6px" : "20px",
            display: "flex", flexDirection: "column", position: "relative", overflow: "hidden",
          }}>
            <div style={{
              position: "absolute", top: 0, left: 0, right: 0, height: compact ? "2px" : "3px",
              background: `linear-gradient(90deg, ${T.accent}, ${T.purple})`, opacity: 0.5,
            }} />
            <div style={{ display: "flex", alignItems: "center", gap: compact ? "3px" : "10px", marginBottom: compact ? "4px" : "12px", marginTop: compact ? "2px" : "4px" }}>
              <span style={{ fontSize: compact ? "8px" : "20px" }}>{card.icon}</span>
              <span style={{ color: T.slideText, fontSize: compact ? "5.5px" : "13px", fontWeight: 700 }}>{card.title}</span>
            </div>
            {(card.items || []).map((item, j) => (
              <div key={j} style={{
                display: "flex", alignItems: "flex-start", gap: compact ? "3px" : "8px",
                marginBottom: compact ? "2px" : "5px",
              }}>
                <span style={{
                  width: compact ? "3px" : "6px", height: compact ? "3px" : "6px",
                  borderRadius: "50%", background: "rgba(59,130,246,0.4)",
                  flexShrink: 0, marginTop: compact ? "1px" : "4px",
                }} />
                <span style={{
                  color: T.textDim, fontSize: compact ? "4px" : "11px", lineHeight: 1.5,
                }}>{item}</span>
              </div>
            ))}
          </div>
        ))}
      </div>

      <SlideFooterBar compact={compact} index={index} total={total} />
    </SlideFrame>
  );
}

// ─── ContentSlide.jsx ──────────────────────────────────────────
function ContentSlide({ data: s, compact, index, total }) {
  const { T } = useTheme();
  const cards = s.cards || [];
  return (
    <SlideFrame compact={compact}>
      <SlideHeaderBar compact={compact} period="Q1 2026" />
      <PageTitle title={s.title} subtitle={s.subtitle} compact={compact} />

      <div style={{
        display: "grid", gridTemplateColumns: "1fr 1fr",
        gap: compact ? "4px" : "12px", flex: 1, alignContent: "start",
        padding: compact ? "0" : "0 36px",
      }}>
        {cards.map((card, i) => (
          <div key={i} style={{
            background: T.cardBg, border: `1px solid ${T.cardBorder}`,
            borderRadius: compact ? "5px" : T.radiusLg, padding: compact ? "6px" : "20px",
          }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: compact ? "4px" : "12px" }}>
              <div style={{ display: "flex", alignItems: "center", gap: compact ? "3px" : "10px" }}>
                <span style={{ fontSize: compact ? "7px" : "18px" }}>{card.icon}</span>
                <span style={{ color: T.slideText, fontSize: compact ? "5.5px" : "13px", fontWeight: 700 }}>{card.title}</span>
              </div>
              {card.status && (
                <span style={{
                  background: card.status === "Active" ? "rgba(34,197,94,0.12)" : card.status === "Complete" ? "rgba(59,130,246,0.12)" : card.status === "In Progress" ? "rgba(234,179,8,0.12)" : "rgba(148,163,194,0.08)",
                  color: card.status === "Active" ? T.success : card.status === "Complete" ? T.accent : card.status === "In Progress" ? "#eab308" : T.textMuted,
                  fontSize: compact ? "3.5px" : "10px", padding: compact ? "1px 3px" : "3px 10px",
                  borderRadius: "10px", fontWeight: 700,
                }}>{card.status}</span>
              )}
            </div>
            {(card.items || []).map((item, j) => (
              <div key={j} style={{ color: T.textDim, fontSize: compact ? "4px" : "11.5px", marginBottom: compact ? "2px" : "5px", lineHeight: 1.6 }}>
                {item.bold && <span style={{ color: T.slideTextSec, fontWeight: 700 }}>{item.bold} </span>}
                {item.text || (typeof item === "string" ? item : "")}
              </div>
            ))}
          </div>
        ))}
      </div>

      <SlideFooterBar compact={compact} index={index} total={total} />
    </SlideFrame>
  );
}

// ─── DashboardSlide.jsx ────────────────────────────────────────
function DashboardSlide({ data: s, compact, index, total }) {
  const { T } = useTheme();
  return (
    <SlideFrame compact={compact}>
      <SlideHeaderBar compact={compact} period="Q1 2026" />
      <PageTitle title={s.title} subtitle={s.subtitle} compact={compact} />

      <div style={{
        display: "grid", gridTemplateColumns: "repeat(3, 1fr)",
        gap: compact ? "4px" : "12px", flex: 1, alignContent: "start",
        padding: compact ? "0" : "0 36px",
      }}>
        {(s.metrics || []).map((m, i) => (
          <div key={i} style={{
            background: T.cardBg, border: `1px solid ${T.cardBorder}`,
            borderRadius: compact ? "5px" : T.radiusLg,
            padding: compact ? "5px" : "18px", textAlign: "center",
            position: "relative", overflow: "hidden",
          }}>
            <div style={{ position: "absolute", bottom: 0, left: 0, right: 0, height: "2px", background: `linear-gradient(90deg, ${T.accent}, transparent)`, opacity: 0.3 }} />
            <div style={{ fontSize: compact ? "8px" : "24px", marginBottom: compact ? "2px" : "6px" }}>{m.icon}</div>
            <div style={{ color: T.accent, fontSize: compact ? "10px" : "28px", fontWeight: 800, lineHeight: 1 }}>{m.value}</div>
            <div style={{ color: T.slideTextDim, fontSize: compact ? "4px" : "11px", marginTop: compact ? "2px" : "5px", fontWeight: 500 }}>{m.label}</div>
            {m.detail && <div style={{ color: T.success, fontSize: compact ? "3.5px" : "10px", marginTop: compact ? "1px" : "4px", fontWeight: 600, opacity: 0.8 }}>{m.detail}</div>}
          </div>
        ))}
      </div>

      <SlideFooterBar compact={compact} index={index} total={total} />
    </SlideFrame>
  );
}

// ─── TableSlide.jsx ────────────────────────────────────────────
function TableSlide({ data: s, compact, index, total }) {
  const { T } = useTheme();
  return (
    <SlideFrame compact={compact}>
      <SlideHeaderBar compact={compact} period="Q1 2026" />
      <PageTitle title={s.title} subtitle={s.subtitle} compact={compact} />

      <div style={{
        background: T.cardBg, border: `1px solid ${T.cardBorder}`,
        borderRadius: compact ? "5px" : T.radiusLg, overflow: "hidden", flex: 1,
        margin: compact ? "0" : "0 36px",
      }}>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead>
            <tr>
              {(s.columns || []).map((col, i) => (
                <th key={i} style={{
                  color: T.slideTextDim, fontSize: compact ? "4px" : "10px", fontWeight: 700,
                  padding: compact ? "3px 5px" : "12px 20px", textAlign: "left",
                  borderBottom: `1px solid ${T.border}`, textTransform: "uppercase",
                  letterSpacing: "0.8px", background: "rgba(59,130,246,0.04)",
                }}>{col}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {(s.rows || []).map((row, i) => (
              <tr key={i} style={{ background: i % 2 === 0 ? "transparent" : "rgba(255,255,255,0.04)" }}>
                {(row.cells || []).map((cell, j) => (
                  <td key={j} style={{
                    color: j === 0 ? "#cbd5e1" : T.textDim, fontWeight: j === 0 ? 600 : 400,
                    fontSize: compact ? "4px" : "11.5px",
                    padding: compact ? "2.5px 5px" : "10px 20px",
                    borderBottom: `1px solid rgba(148,163,194,0.04)`,
                  }}>{cell}</td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {s.summaryStats?.length > 0 && (
        <div style={{ display: "flex", gap: compact ? "4px" : "12px", marginTop: compact ? "4px" : "12px", padding: compact ? "0" : "0 36px" }}>
          {s.summaryStats.map((st, i) => (
            <div key={i} style={{
              background: T.accentGlow, borderRadius: T.radius,
              padding: compact ? "3px 6px" : "8px 18px", display: "flex", alignItems: "center", gap: compact ? "3px" : "8px",
            }}>
              <span style={{ color: T.accent, fontSize: compact ? "8px" : "20px", fontWeight: 800 }}>{st.number}</span>
              <span style={{ color: T.textDim, fontSize: compact ? "4px" : "11px", fontWeight: 500 }}>{st.label}</span>
            </div>
          ))}
        </div>
      )}

      <SlideFooterBar compact={compact} index={index} total={total} />
    </SlideFrame>
  );
}

// ─── SectionOverviewSlide.jsx ──────────────────────────────────
function SectionOverviewSlide({ data: s, compact, index, total }) {
  const { T } = useTheme();
  return (
    <SlideFrame compact={compact}>
      <SlideHeaderBar compact={compact} period="Q1 2026" />

      <div style={{
        flex: 1, display: "flex", gap: compact ? "6px" : "28px",
        padding: compact ? "0" : "0 36px", alignItems: "center",
      }}>
        {/* Left — Section info */}
        <div style={{ flex: 1 }}>
          <div style={{
            width: compact ? "20px" : "56px", height: compact ? "20px" : "56px",
            borderRadius: compact ? "8px" : "16px",
            background: T.accentGlow, border: `1px solid rgba(59,130,246,0.15)`,
            display: "flex", alignItems: "center", justifyContent: "center",
            fontSize: compact ? "10px" : "28px", marginBottom: compact ? "6px" : "18px",
          }}>{s.sectionIcon}</div>
          <h2 style={{ color: T.slideText, fontSize: compact ? "11px" : "26px", fontWeight: 800, margin: 0, letterSpacing: "-0.02em" }}>{s.sectionTitle || s.title}</h2>
          <p style={{ color: T.textDim, fontSize: compact ? "4px" : "12px", marginTop: compact ? "3px" : "10px", lineHeight: 1.6, maxWidth: "400px" }}>{s.description}</p>

          {s.stats?.length > 0 && (
            <div style={{ display: "flex", gap: compact ? "4px" : "12px", marginTop: compact ? "6px" : "20px" }}>
              {s.stats.map((st, i) => (
                <div key={i} style={{
                  background: T.cardBg, border: `1px solid ${T.cardBorder}`,
                  borderRadius: compact ? "6px" : "12px", padding: compact ? "4px 8px" : "12px 20px",
                  textAlign: "center",
                }}>
                  <div style={{ color: T.accent, fontSize: compact ? "9px" : "22px", fontWeight: 800 }}>{st.number}</div>
                  <div style={{ color: T.textDim, fontSize: compact ? "3.5px" : "10px", marginTop: "2px" }}>{st.label}</div>
                </div>
              ))}
            </div>
          )}
        </div>

        {/* Right — Billing metrics */}
        {s.billingMetrics?.length > 0 && (
          <div style={{
            background: T.cardBg, border: `1px solid ${T.cardBorder}`,
            borderRadius: compact ? "6px" : T.radiusLg,
            padding: compact ? "6px" : "24px", minWidth: compact ? "60px" : "240px",
          }}>
            <div style={{ fontSize: compact ? "5px" : "12px", color: T.textMuted, fontWeight: 700, marginBottom: compact ? "4px" : "14px", textTransform: "uppercase", letterSpacing: "0.8px" }}>Billing Overview</div>
            {s.billingMetrics.map((bm, i) => (
              <div key={i} style={{
                display: "flex", justifyContent: "space-between", alignItems: "center",
                padding: compact ? "2px 0" : "8px 0",
                borderBottom: i < s.billingMetrics.length - 1 ? `1px solid ${T.border}` : "none",
              }}>
                <span style={{ color: T.textDim, fontSize: compact ? "4px" : "11px" }}>{bm.label}</span>
                <span style={{ color: T.slideTextSec, fontSize: compact ? "5px" : "13px", fontWeight: 700 }}>{bm.value}</span>
              </div>
            ))}
          </div>
        )}
      </div>

      <SlideFooterBar compact={compact} index={index} total={total} />
    </SlideFrame>
  );
}

// ─── HighlightListSlide.jsx ────────────────────────────────────
function HighlightListSlide({ data: s, compact, index, total }) {
  const { T } = useTheme();
  return (
    <SlideFrame compact={compact}>
      <SlideHeaderBar compact={compact} period="Q1 2026" />
      <PageTitle title={s.title} subtitle={s.subtitle} compact={compact} />

      <div style={{
        flex: 1, display: "flex", flexDirection: "column", gap: compact ? "3px" : "10px",
        padding: compact ? "0" : "0 36px",
      }}>
        {(s.items || []).map((item, i) => (
          <div key={i} style={{
            display: "flex", alignItems: "center", gap: compact ? "6px" : "16px",
            background: T.cardBg, border: `1px solid ${T.cardBorder}`,
            borderRadius: compact ? "5px" : "12px",
            padding: compact ? "5px 8px" : "14px 22px",
            borderLeft: `3px solid ${item.color || T.accent}`,
          }}>
            <span style={{
              width: compact ? "5px" : "10px", height: compact ? "5px" : "10px",
              borderRadius: "50%", background: item.color || T.accent, flexShrink: 0,
              boxShadow: `0 0 8px ${item.color || T.accent}40`,
            }} />
            <span style={{ color: T.slideTextSec, fontSize: compact ? "4.5px" : "13px", fontWeight: 500, lineHeight: 1.5 }}>{item.text}</span>
          </div>
        ))}
      </div>

      <SlideFooterBar compact={compact} index={index} total={total} />
    </SlideFrame>
  );
}

// ─── ThankYouSlide.jsx ─────────────────────────────────────────
function ThankYouSlide({ data: s, compact, index, total }) {
  const { T } = useTheme();
  return (
    <SlideFrame compact={compact}>
      <SlideHeaderBar compact={compact} period="Q1 2026" />

      <div style={{ flex: 1, display: "flex", flexDirection: "column", justifyContent: "center", alignItems: "center", textAlign: "center" }}>
        <div style={{
          width: compact ? "20px" : "68px", height: compact ? "20px" : "68px", borderRadius: "50%",
          background: `linear-gradient(135deg, ${T.accent}, ${T.purple})`,
          display: "flex", alignItems: "center", justifyContent: "center",
          fontSize: compact ? "10px" : "28px", marginBottom: compact ? "6px" : "24px",
          boxShadow: "0 8px 40px rgba(59,130,246,0.25)",
        }}>♥</div>
        <h1 style={{
          color: T.slideText, fontSize: compact ? "14px" : "42px",
          fontWeight: 800, margin: 0, letterSpacing: "-0.03em",
        }}>{s.title}</h1>
        <p style={{
          color: T.textDim, fontSize: compact ? "5px" : "15px",
          marginTop: compact ? "4px" : "12px", maxWidth: "480px", lineHeight: 1.7,
        }}>{s.message}</p>
        <div style={{
          marginTop: compact ? "8px" : "32px",
          color: T.textMuted, fontSize: compact ? "4px" : "11px",
          fontWeight: 700, letterSpacing: "3px", textTransform: "uppercase",
        }}>{s.signature}</div>
      </div>

      <SlideFooterBar compact={compact} index={index} total={total} />
    </SlideFrame>
  );
}

// ─── SlideRenderer (Router) ───────────────────────────────────
function SlideRenderer({ slide, compact = false, index = 0, total = 1 }) {
  const props = { data: slide, compact, index, total };

  // If layout overrides exist and not compact (thumbnails), use layout-aware rendering
  if (slide.layout && Object.keys(slide.layout).length > 0 && !compact) {
    return <LayoutSlide slide={slide} index={index} total={total} />;
  }

  switch (slide.type) {
    case "title": return <TitleSlide {...props} />;
    case "agenda": return <AgendaSlide {...props} />;
    case "content": return <ContentSlide {...props} />;
    case "dashboard": return <DashboardSlide {...props} />;
    case "table": return <TableSlide {...props} />;
    case "section-overview": return <SectionOverviewSlide {...props} />;
    case "highlight-list": return <HighlightListSlide {...props} />;
    case "thank-you": return <ThankYouSlide {...props} />;
    default: return <ContentSlide {...props} />;
  }
}

// ─── Layout-Aware Slide Renderer ──────────────────────────────
// Renders elements at their canvas-dragged positions using absolute positioning
function LayoutSlide({ slide, index, total }) {
  const { T } = useTheme();
  const elements = useMemo(() => generateCanvasElements(slide), [slide]);

  const renderElementContent = (el, s) => renderCanvasContent(el, s, slide, T);

  return (
    <SlideFrame>
      <SlideHeaderBar period="Q1 2026" />
      <div style={{ flex: 1, position: "relative", overflow: "hidden" }}>
        {elements.map((el) => {
          const layout = slide.layout?.[el.id];
          const rect = layout ? { ...el.defaultRect, ...layout } : el.defaultRect;
          const left = (rect.x / 960) * 100 + "%";
          const top = (rect.y / 540) * 100 + "%";
          const width = (rect.w / 960) * 100 + "%";
          const height = (rect.h / 540) * 100 + "%";
          return (
            <div key={el.id} style={{ position: "absolute", left, top, width, height, overflow: "hidden" }}>
              {renderElementContent(el, slide)}
            </div>
          );
        })}
      </div>
      <SlideFooterBar index={index} total={total} />
    </SlideFrame>
  );
}

// ─── Shared element content renderer for Canvas & Layout ──────
function renderCanvasContent(el, s, slide, T) {
  // Apply per-element font/style overrides from elementStyles
  const es = slide.elementStyles?.[el.id] || {};
  const baseText = {
    fontFamily: es.fontFamily || T.font,
    fontSize: es.fontSize ? `${es.fontSize}px` : undefined,
    fontWeight: es.fontWeight || undefined,
    fontStyle: es.fontStyle || undefined,
    margin: 0,
  };
  switch (el.id) {
    case "badge":
        return s.badge ? <div style={{ background: `linear-gradient(135deg, ${T.accent}, ${T.accentDark})`, color: "#fff", padding: "6px 18px", borderRadius: "16px", fontSize: "11px", fontWeight: 700, letterSpacing: "0.8px", textTransform: "uppercase", display: "inline-block" }}>{s.badge}</div> : null;
      case "title":
        return <h1 style={{ color: T.slideText, fontSize: es.fontSize ? `${es.fontSize}px` : "34px", fontWeight: es.fontWeight || 800, letterSpacing: "-0.03em", lineHeight: 1.15, ...baseText }}>{s.title}</h1>;
      case "subtitle":
        return <p style={{ color: T.slideTextDim, fontSize: es.fontSize ? `${es.fontSize}px` : "14px", lineHeight: 1.5, ...baseText }}>{s.subtitle}</p>;
      case "page-title":
        return (<div><h2 style={{ ...baseText, color: T.slideText, fontSize: "26px", fontWeight: 800, letterSpacing: "-0.025em" }}>{s.title}</h2>
          {s.subtitle && <p style={{ ...baseText, color: T.slideTextDim, fontSize: "13px", marginTop: "4px" }}>{s.subtitle}</p>}</div>);
      case "highlights":
        return <div style={{ display: "flex", flexDirection: "column", gap: "7px" }}>{(s.highlights || []).map((h, i) => (
          <div key={i} style={{ display: "flex", alignItems: "center", gap: "10px", fontSize: "12px", color: h.color || T.success, fontWeight: 600 }}>
            <div style={{ width: "8px", height: "8px", borderRadius: "50%", background: h.color || T.success, flexShrink: 0 }} />{h.text}
          </div>))}</div>;
      case "metrics":
        return <div style={{ display: "flex", flexWrap: "wrap", gap: "10px" }}>{(s.metrics || []).map((m, i) => (
          <div key={i} style={{ background: T.cardBg, border: `1px solid ${T.cardBorder}`, borderRadius: "10px", padding: "12px 16px", textAlign: "center", minWidth: "90px", flex: "1" }}>
            <div style={{ fontSize: "22px", fontWeight: 800, color: T.accent }}>{m.number || m.value}</div>
            <div style={{ fontSize: "10px", color: T.slideTextDim, marginTop: "3px" }}>{m.label}</div>
          </div>))}</div>;
      case "stats":
        return <div style={{ display: "flex", flexDirection: "column", gap: "8px" }}>{(s.stats || []).map((st, i) => (
          <div key={i} style={{ background: T.cardBg, border: `1px solid ${T.cardBorder}`, borderRadius: "8px", padding: "10px 14px", display: "flex", alignItems: "center", gap: "10px" }}>
            {st.icon && <span style={{ fontSize: "16px" }}>{st.icon}</span>}
            <div><div style={{ fontSize: "10px", color: T.slideTextDim }}>{st.label}</div><div style={{ fontSize: "13px", fontWeight: 700, color: T.slideTextSec }}>{st.value}</div></div>
          </div>))}</div>;
      case "icon":
        return <div style={{ width: "100%", height: "100%", borderRadius: "50%", background: `linear-gradient(135deg, ${T.accent}, ${T.purple})`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: "28px" }}>♥</div>;
      case "message":
        return <p style={{ ...baseText, color: T.slideTextDim, fontSize: "15px", lineHeight: 1.7, textAlign: "center" }}>{s.message}</p>;
      case "signature":
        return <div style={{ ...baseText, color: T.slideTextDim, fontSize: "11px", fontWeight: 700, letterSpacing: "3px", textTransform: "uppercase", textAlign: "center" }}>{s.signature}</div>;
      case "table":
        return (<div style={{ background: T.cardBg, border: `1px solid ${T.cardBorder}`, borderRadius: "10px", overflow: "hidden", width: "100%", height: "100%" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "11px" }}>
            <thead><tr>{(s.columns || []).map((c, ci) => <th key={ci} style={{ padding: "8px 12px", textAlign: "left", color: T.slideTextDim, fontWeight: 700, fontSize: "9px", textTransform: "uppercase", borderBottom: `1px solid ${T.cardBorder}` }}>{c}</th>)}</tr></thead>
            <tbody>{(s.rows || []).map((r, ri) => <tr key={ri}>{(r.cells || []).map((cell, ci) => <td key={ci} style={{ padding: "6px 12px", color: ci === 0 ? T.slideTextSec : T.slideTextDim, fontWeight: ci === 0 ? 700 : 400, borderBottom: `1px solid ${T.cardBorder}` }}>{cell}</td>)}</tr>)}</tbody>
          </table></div>);
      case "summary":
        return <div style={{ display: "flex", gap: "12px" }}>{(s.summaryStats || []).map((st, i) => (
          <div key={i} style={{ background: T.cardBg, border: `1px solid ${T.cardBorder}`, borderRadius: "8px", padding: "8px 16px", textAlign: "center" }}>
            <div style={{ fontSize: "18px", fontWeight: 800, color: T.accent }}>{st.number}</div>
            <div style={{ fontSize: "9px", color: T.slideTextDim }}>{st.label}</div>
          </div>))}</div>;
      default:
        // Handle sub-elements: cards[N].icon, cards[N].title, cards[N].status, cards[N].items[M], metrics[N].icon, etc.
        const subMatch = el.id.match(/^(\w+)\[(\d+)\]\.(\w+)(?:\[(\d+)\])?$/);
        if (subMatch) {
          const [, field, idx, prop, subIdx] = subMatch;
          const parent = (slide[field] || [])[parseInt(idx)];
          if (!parent) return null;
          const bulletStyle = slide.bulletStyle || "•";

          if (prop === "bg") {
            return <div style={{ background: T.cardBg, border: `1px solid ${T.cardBorder}`, borderRadius: "10px", width: "100%", height: "100%", borderLeft: field === "cards" ? `3px solid ${T.accent}` : "none" }} />;
          }
          if (prop === "icon") {
            return <div style={{ fontSize: "20px", display: "flex", alignItems: "center", justifyContent: "center", width: "100%", height: "100%" }}>{parent.icon}</div>;
          }
          if (prop === "title") {
            return <div style={{ fontSize: es.fontSize ? `${es.fontSize}px` : "13px", fontWeight: es.fontWeight || 700, color: T.slideText, ...baseText }}>{parent.title}</div>;
          }
          if (prop === "status") {
            return <span style={{ fontSize: "9px", fontWeight: 600, color: T.accent, background: T.slideAccentBg, padding: "2px 8px", borderRadius: "8px", display: "inline-block" }}>{parent.status}</span>;
          }
          if (prop === "items" && subIdx !== undefined) {
            const item = (parent.items || [])[parseInt(subIdx)];
            if (!item) return null;
            return (<div style={{ display: "flex", gap: "6px", fontSize: "11px", color: T.slideTextDim, alignItems: "flex-start" }}>
              <span style={{ color: T.accent, flexShrink: 0, fontSize: "10px", marginTop: "1px" }}>{bulletStyle}</span>
              {typeof item === "string" ? <span>{item}</span> : <span><strong style={{ color: T.slideTextSec }}>{item.bold}</strong> {item.text}</span>}
            </div>);
          }
          if (prop === "value") {
            return <div style={{ fontSize: es.fontSize ? `${es.fontSize}px` : "22px", fontWeight: es.fontWeight || 800, color: T.accent, textAlign: "center", ...baseText }}>{parent.value || parent.number}</div>;
          }
          if (prop === "label") {
            return <div style={{ fontSize: es.fontSize ? `${es.fontSize}px` : "10px", color: T.slideTextDim, textAlign: "center", ...baseText }}>{parent.label}</div>;
          }
          if (prop === "detail") {
            return <div style={{ fontSize: "9px", color: T.success, fontWeight: 600, textAlign: "center" }}>{parent.detail}</div>;
          }
          return null;
        }

        // Handle simple array elements: cards[N], metrics[N], items[N], etc.
        const arrMatch = el.id.match(/^(\w+)\[(\d+)\]$/);
        if (arrMatch) {
          const [, field, idx] = arrMatch;
          const item = (slide[field] || [])[parseInt(idx)];
          if (!item) return null;
          if (field === "cards") {
            return (<div style={{ background: T.cardBg, border: `1px solid ${T.cardBorder}`, borderRadius: "10px", padding: "14px", height: "100%", borderLeft: `3px solid ${T.accent}` }}>
              <div style={{ display: "flex", alignItems: "center", gap: "8px", marginBottom: "10px" }}>
                {item.icon && <span style={{ fontSize: "16px" }}>{item.icon}</span>}
                <span style={{ fontSize: "13px", fontWeight: 700, color: T.slideText }}>{item.title}</span>
                {item.status && <span style={{ fontSize: "9px", fontWeight: 600, color: T.accent, background: T.slideAccentBg, padding: "2px 8px", borderRadius: "8px", marginLeft: "auto" }}>{item.status}</span>}
              </div>
              {(item.items || []).map((it, j) => (
                <div key={j} style={{ display: "flex", gap: "6px", fontSize: "11px", color: T.slideTextDim, padding: "3px 0" }}>
                  <span style={{ color: T.accent }}>•</span>
                  {typeof it === "string" ? it : <><span style={{ color: T.slideTextSec, fontWeight: 700 }}>{it.bold} </span>{it.text}</>}
                </div>
              ))}
            </div>);
          }
          if (field === "metrics") {
            return (<div style={{ background: T.cardBg, border: `1px solid ${T.cardBorder}`, borderRadius: "10px", padding: "14px", textAlign: "center", height: "100%" }}>
              {item.icon && <div style={{ fontSize: "20px", marginBottom: "6px" }}>{item.icon}</div>}
              <div style={{ fontSize: "22px", fontWeight: 800, color: T.accent }}>{item.value || item.number}</div>
              <div style={{ fontSize: "10px", color: T.slideTextDim, marginTop: "4px" }}>{item.label}</div>
              {item.detail && <div style={{ fontSize: "9px", color: T.success, fontWeight: 600, marginTop: "2px" }}>{item.detail}</div>}
            </div>);
          }
          if (field === "items") {
            return (<div style={{ background: T.cardBg, border: `1px solid ${T.cardBorder}`, borderRadius: "8px", padding: "10px 14px", display: "flex", alignItems: "center", gap: "12px", borderLeft: `3px solid ${item.color || T.accent}` }}>
              <div style={{ width: "8px", height: "8px", borderRadius: "50%", background: item.color || T.accent, flexShrink: 0 }} />
              <span style={{ fontSize: "12px", color: T.slideTextSec }}>{item.text}</span>
            </div>);
          }
          if (field === "highlights") {
            return (<div style={{ display: "flex", alignItems: "center", gap: "10px", fontSize: "11px" }}>
              <div style={{ width: "8px", height: "8px", borderRadius: "50%", background: item.color || T.success, flexShrink: 0 }} />
              <span style={{ color: item.color || T.success, fontWeight: 600 }}>{item.text}</span>
            </div>);
          }
          if (field === "stats") {
            return (<div style={{ background: T.cardBg, border: `1px solid ${T.cardBorder}`, borderRadius: "8px", padding: "8px 12px", display: "flex", alignItems: "center", gap: "10px", height: "100%" }}>
              {item.icon && <span style={{ fontSize: "16px" }}>{item.icon}</span>}
              <div><div style={{ fontSize: "9px", color: T.slideTextDim }}>{item.label}</div><div style={{ fontSize: "12px", fontWeight: 700, color: T.slideTextSec }}>{item.value || item.number}</div></div>
            </div>);
          }
          if (field === "billingMetrics") {
            return (<div style={{ background: T.cardBg, border: `1px solid ${T.cardBorder}`, borderRadius: "6px", padding: "8px 14px", display: "flex", justifyContent: "space-between", alignItems: "center", height: "100%" }}>
              <span style={{ fontSize: "11px", color: T.slideTextDim }}>{item.label}</span>
              <span style={{ fontSize: "12px", fontWeight: 700, color: T.accent }}>{item.value}</span>
            </div>);
          }
          if (field === "summaryStats") {
            return (<div style={{ background: T.cardBg, border: `1px solid ${T.cardBorder}`, borderRadius: "8px", padding: "8px 12px", textAlign: "center", height: "100%" }}>
              <div style={{ fontSize: "18px", fontWeight: 800, color: T.accent }}>{item.number}</div>
              <div style={{ fontSize: "9px", color: T.slideTextDim }}>{item.label}</div>
            </div>);
          }
        }
        // Section overview specific
        if (el.id === "section-icon") return <div style={{ width: "100%", height: "100%", background: T.slideAccentBg, borderRadius: "12px", display: "flex", alignItems: "center", justifyContent: "center", fontSize: "28px" }}>{slide.sectionIcon || "📊"}</div>;
        if (el.id === "section-title") return <h2 style={{ ...baseText, color: T.slideText, fontSize: "22px", fontWeight: 800, letterSpacing: "-0.02em" }}>{slide.sectionTitle || slide.title}</h2>;
        if (el.id === "description") return <p style={{ ...baseText, color: T.slideTextDim, fontSize: "12px", lineHeight: 1.6 }}>{slide.description}</p>;
        // Header/Footer elements
        if (el.id === "header-bar") {
          const isLight = T.mode === "light";
          return <div style={{ width: "100%", height: "100%", background: isLight ? T.accent : T.slideHeaderBg, borderBottom: isLight ? "none" : `1px solid ${T.footerBorder}` }} />;
        }
        if (el.id === "header-logo") return <div style={{ width: "30px", height: "30px", borderRadius: "7px", background: T.mode === "light" ? "#fff" : `linear-gradient(135deg, ${T.accent}, ${T.purple})`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: "13px", color: T.mode === "light" ? T.accent : "#fff", fontWeight: 900 }}>M</div>;
        if (el.id === "header-title") return <div style={{ fontSize: "12px", fontWeight: 700, color: T.mode === "light" ? "#fff" : T.slideText, ...baseText }}>Microsoft Practice</div>;
        if (el.id === "header-period") return <div style={{ fontSize: "9px", color: T.mode === "light" ? "rgba(255,255,255,0.8)" : T.slideTextDim, background: T.mode === "light" ? "rgba(255,255,255,0.15)" : T.cardBg, border: T.mode === "light" ? "none" : `1px solid ${T.cardBorder}`, borderRadius: "5px", padding: "3px 8px", textAlign: "center" }}>Q1 2026</div>;
        if (el.id === "footer-bar") return <div style={{ width: "100%", height: "100%", background: T.footerBg, borderTop: `1px solid ${T.footerBorder}` }} />;
        if (el.id === "footer-logo") return <img src="/miracle-logo-large.png" style={{ height: "18px", objectFit: "contain" }} alt="Miracle" />;
        if (el.id === "footer-page") return <div style={{ fontSize: "9px", color: T.slideTextDim, fontFamily: T.fontMono, textAlign: "center", background: T.cardBg, borderRadius: "4px", padding: "2px 8px" }}>{(slide._index || 0) + 1} / {slide._total || 1}</div>;

        // Legacy single "billing" block — kept for backward compat
        if (el.id === "billing") {
          return (<div style={{ background: T.cardBg, border: `1px solid ${T.cardBorder}`, borderRadius: "10px", padding: "16px", height: "100%" }}>
            <div style={{ fontSize: "9px", fontWeight: 700, color: T.slideTextDim, textTransform: "uppercase", marginBottom: "10px" }}>Billing Overview</div>
            {(slide.billingMetrics || []).map((bm, i) => (
              <div key={i} style={{ display: "flex", justifyContent: "space-between", padding: "6px 0", borderBottom: i < (slide.billingMetrics.length - 1) ? `1px solid ${T.cardBorder}` : "none", fontSize: "12px" }}>
                <span style={{ color: T.slideTextDim }}>{bm.label}</span>
                <span style={{ color: T.accent, fontWeight: 700 }}>{bm.value}</span>
              </div>))}
          </div>);
        }
        return null;
    }
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 📦  EXPORT ENGINE — PPTX (PptxGenJS) + PDF (canvas capture)
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/*
  Color palette for PPTX (no # prefix — PptxGenJS requirement)
  NEVER use "#" — causes file corruption
  NEVER reuse option objects — PptxGenJS mutates them
*/
function getPPT(themeT) {
  const isLight = themeT?.mode === "light";
  const accent = h(themeT?.accent || "#3B82F6");
  const accentDark = h(themeT?.accentDark || "#2563EB");
  if (isLight) {
    return {
      bg: "F0F2F5", bgCard: "FFFFFF", bgCardLight: "F8FAFC",
      accent, accentDark, purple: "8B5CF6",
      success: "28A745", warning: "F59E0B",
      white: "FFFFFF",
      textLight: "111827", textMid: "374151", textMuted: "6B7280", textDim: "9CA3AF",
      border: "E5E7EB",
      headerFont: "Trebuchet MS", bodyFont: "Calibri",
      headerBg: accent, headerText: "FFFFFF",
    };
  }
  return {
    bg: "070B14", bgCard: "0F1729", bgCardLight: "141E33",
    accent, accentDark, purple: "8B5CF6",
    success: "22C55E", warning: "EAB308",
    white: "F8FAFC",
    textLight: "F1F5F9", textMid: "CBD5E1", textMuted: "94A3B8", textDim: "64748B",
    border: "1E293B",
    headerFont: "Trebuchet MS", bodyFont: "Calibri",
    headerBg: "070B14", headerText: "F1F5F9",
  };
}
// Default for backward compat
let PPT = getPPT();

function h(hex) { return (hex || "").replace("#", ""); }

function makeShadow() {
  return { type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.2 };
}

// ─── Load PptxGenJS (npm module) ──────────────────────────────
async function loadPptxGenJS() {
  const mod = await import("pptxgenjs");
  return mod.default || mod;
}

// ─── Shared PPTX helpers ───────────────────────────────────────
function addSlideChrome(pptSlide, pres, index, total) {
  const isLight = PPT.headerBg === PPT.accent;

  // Background
  pptSlide.background = { color: PPT.bg };

  // Corner accent lines
  pptSlide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.8, h: 0.03, fill: { color: PPT.accent, transparency: isLight ? 40 : 60 } });
  pptSlide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.03, h: 0.8, fill: { color: PPT.accent, transparency: isLight ? 40 : 60 } });
  pptSlide.addShape(pres.shapes.RECTANGLE, { x: 9.2, y: 5.595, w: 0.8, h: 0.03, fill: { color: PPT.purple, transparency: 70 } });
  pptSlide.addShape(pres.shapes.RECTANGLE, { x: 9.97, y: 4.825, w: 0.03, h: 0.8, fill: { color: PPT.purple, transparency: 70 } });

  // Header bar — solid accent in light mode, dark bg in dark mode
  pptSlide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.55, fill: { color: PPT.headerBg } });
  if (!isLight) pptSlide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.55, w: 10, h: 0.01, fill: { color: PPT.border, transparency: 50 } });

  // Logo
  pptSlide.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.12, w: 0.32, h: 0.32, fill: { color: isLight ? "FFFFFF" : PPT.accent }, rectRadius: 0.06 });
  pptSlide.addText("M", { x: 0.4, y: 0.12, w: 0.32, h: 0.32, fontSize: 12, fontFace: PPT.headerFont, color: isLight ? PPT.accent : "FFFFFF", bold: true, align: "center", valign: "middle", margin: 0 });
  pptSlide.addText([
    { text: "Microsoft Practice", options: { fontSize: 10, bold: true, color: PPT.headerText, breakLine: true } },
    { text: "Quarterly Report", options: { fontSize: 7, color: isLight ? "E0E7FF" : PPT.textDim } },
  ], { x: 0.82, y: 0.1, w: 2, h: 0.38, fontFace: PPT.bodyFont, margin: 0 });

  // Period badge (top right)
  pptSlide.addShape(pres.shapes.RECTANGLE, { x: 8.2, y: 0.13, w: 0.85, h: 0.3, fill: { color: isLight ? "FFFFFF" : PPT.bgCard, transparency: isLight ? 70 : 0 }, line: isLight ? undefined : { color: PPT.border, width: 0.5 }, rectRadius: 0.05 });
  pptSlide.addText("Q1 2026", { x: 8.2, y: 0.13, w: 0.85, h: 0.3, fontSize: 7, fontFace: PPT.bodyFont, color: isLight ? "E0E7FF" : PPT.textMuted, align: "center", valign: "middle", margin: 0 });

  // Footer bar
  pptSlide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.2, w: 10, h: 0.01, fill: { color: PPT.border, transparency: 50 } });
  pptSlide.addText("Quarterly Report  •  Confidential", { x: 0.4, y: 5.25, w: 4, h: 0.3, fontSize: 7, fontFace: PPT.bodyFont, color: PPT.textDim, margin: 0 });
  pptSlide.addShape(pres.shapes.RECTANGLE, { x: 8.8, y: 5.28, w: 0.75, h: 0.24, fill: { color: isLight ? PPT.bgCardLight : PPT.bgCard }, rectRadius: 0.04 });
  pptSlide.addText(`${index + 1} / ${total}`, { x: 8.8, y: 5.28, w: 0.75, h: 0.24, fontSize: 8, fontFace: "Consolas", color: PPT.textMuted, align: "center", valign: "middle", margin: 0 });
}

function addPageTitle(pptSlide, title, subtitle) {
  pptSlide.addText(title, { x: 0.5, y: 0.7, w: 9, h: 0.45, fontSize: 22, fontFace: PPT.headerFont, color: PPT.textLight, bold: true, margin: 0 });
  if (subtitle) {
    pptSlide.addText(subtitle, { x: 0.5, y: 1.1, w: 9, h: 0.3, fontSize: 11, fontFace: PPT.bodyFont, color: PPT.textDim, margin: 0 });
  }
}

// ─── Per-slide-type PPTX renderers ─────────────────────────────

function renderTitlePPTX(pptSlide, pres, s, idx, total) {
  addSlideChrome(pptSlide, pres, idx, total);

  // Badge
  if (s.badge) {
    pptSlide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.0, w: 1.4, h: 0.3, fill: { color: PPT.accent }, rectRadius: 0.15 });
    pptSlide.addText(s.badge, { x: 0.6, y: 1.0, w: 1.4, h: 0.3, fontSize: 9, fontFace: PPT.bodyFont, color: PPT.headerText, bold: true, align: "center", valign: "middle", margin: 0 });
  }

  // Title + Subtitle (left side)
  pptSlide.addText(s.title || "", { x: 0.6, y: 1.45, w: 5.2, h: 0.8, fontSize: 28, fontFace: PPT.headerFont, color: PPT.textLight, bold: true, margin: 0 });
  pptSlide.addText(s.subtitle || "", { x: 0.6, y: 2.25, w: 5.2, h: 0.5, fontSize: 12, fontFace: PPT.bodyFont, color: PPT.textMuted, margin: 0 });

  // Highlights
  (s.highlights || []).forEach((hl, i) => {
    const hlColor = h(hl.color) || PPT.success;
    pptSlide.addShape(pres.shapes.OVAL, { x: 0.65, y: 2.9 + i * 0.35, w: 0.08, h: 0.08, fill: { color: hlColor } });
    pptSlide.addText(hl.text, { x: 0.85, y: 2.82 + i * 0.35, w: 4.5, h: 0.3, fontSize: 10, fontFace: PPT.bodyFont, color: hlColor, bold: true, margin: 0 });
  });

  // Right side — Metrics
  const metricsStartX = 6.2;
  (s.metrics || []).forEach((m, i) => {
    const mx = metricsStartX + (i % 3) * 1.2;
    const my = 1.0 + Math.floor(i / 3) * 0.9;
    pptSlide.addShape(pres.shapes.RECTANGLE, { x: mx, y: my, w: 1.1, h: 0.75, fill: { color: PPT.bgCard }, line: { color: PPT.border, width: 0.5 }, rectRadius: 0.08 });
    pptSlide.addText(m.number, { x: mx, y: my + 0.08, w: 1.1, h: 0.35, fontSize: 18, fontFace: PPT.headerFont, color: PPT.accent, bold: true, align: "center", valign: "middle", margin: 0 });
    pptSlide.addText(m.label, { x: mx, y: my + 0.45, w: 1.1, h: 0.22, fontSize: 7, fontFace: PPT.bodyFont, color: PPT.textDim, align: "center", valign: "middle", margin: 0 });
  });

  // Right side — Stat rows with icons
  (s.stats || []).forEach((st, i) => {
    const sy = 2.2 + (s.metrics ? Math.ceil((s.metrics.length) / 3) * 0.9 : 0) + i * 0.55;
    pptSlide.addShape(pres.shapes.RECTANGLE, { x: metricsStartX, y: sy, w: 3.5, h: 0.45, fill: { color: PPT.bgCard }, line: { color: PPT.border, width: 0.5 }, rectRadius: 0.06 });
    // Icon
    if (st.icon) {
      pptSlide.addText(st.icon, { x: metricsStartX + 0.1, y: sy + 0.05, w: 0.35, h: 0.35, fontSize: 12, align: "center", valign: "middle", margin: 0 });
    }
    pptSlide.addText(st.label, { x: metricsStartX + 0.5, y: sy, w: 1.5, h: 0.22, fontSize: 7, fontFace: PPT.bodyFont, color: PPT.textDim, margin: 0, valign: "bottom" });
    pptSlide.addText(st.value, { x: metricsStartX + 0.5, y: sy + 0.2, w: 2.5, h: 0.22, fontSize: 10, fontFace: PPT.bodyFont, color: PPT.textMid, bold: true, margin: 0, valign: "top" });
  });
}

function renderAgendaPPTX(pptSlide, pres, s, idx, total) {
  addSlideChrome(pptSlide, pres, idx, total);
  addPageTitle(pptSlide, s.title, s.subtitle);

  const cards = s.cards || [];
  const cardW = cards.length <= 3 ? 2.7 : 2.05;
  const gap = 0.2;
  const startX = 0.5;

  cards.forEach((card, i) => {
    const cx = startX + i * (cardW + gap);
    const cy = 1.55;
    const cardH = 3.3;

    pptSlide.addShape(pres.shapes.RECTANGLE, { x: cx, y: cy, w: cardW, h: cardH, fill: { color: PPT.bgCard }, line: { color: PPT.border, width: 0.5 }, rectRadius: 0.08 });
    // Top accent bar
    pptSlide.addShape(pres.shapes.RECTANGLE, { x: cx, y: cy, w: cardW, h: 0.035, fill: { color: PPT.accent, transparency: 40 } });

    // Card icon
    if (card.icon) {
      pptSlide.addText(card.icon, { x: cx + 0.12, y: cy + 0.12, w: 0.3, h: 0.3, fontSize: 13, align: "center", valign: "middle", margin: 0 });
      pptSlide.addText(card.title, { x: cx + 0.45, y: cy + 0.15, w: cardW - 0.6, h: 0.35, fontSize: 11, fontFace: PPT.headerFont, color: PPT.textLight, bold: true, margin: 0 });
    } else {
      pptSlide.addText(card.title, { x: cx + 0.15, y: cy + 0.15, w: cardW - 0.3, h: 0.35, fontSize: 11, fontFace: PPT.headerFont, color: PPT.textLight, bold: true, margin: 0 });
    }

    (card.items || []).forEach((item, j) => {
      pptSlide.addShape(pres.shapes.OVAL, { x: cx + 0.2, y: cy + 0.7 + j * 0.35, w: 0.06, h: 0.06, fill: { color: PPT.accent, transparency: 50 } });
      pptSlide.addText(item, { x: cx + 0.35, y: cy + 0.6 + j * 0.35, w: cardW - 0.55, h: 0.3, fontSize: 9, fontFace: PPT.bodyFont, color: PPT.textDim, margin: 0 });
    });
  });
}

function renderContentPPTX(pptSlide, pres, s, idx, total) {
  addSlideChrome(pptSlide, pres, idx, total);
  addPageTitle(pptSlide, s.title, s.subtitle);

  const cards = s.cards || [];
  cards.forEach((card, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const cx = 0.5 + col * 4.6;
    const cy = 1.55 + row * 1.75;
    const cw = 4.4;
    const ch = 1.6;

    pptSlide.addShape(pres.shapes.RECTANGLE, { x: cx, y: cy, w: cw, h: ch, fill: { color: PPT.bgCard }, line: { color: PPT.border, width: 0.5 }, rectRadius: 0.08 });
    // Left accent border
    pptSlide.addShape(pres.shapes.RECTANGLE, { x: cx, y: cy, w: 0.04, h: ch, fill: { color: PPT.accent } });

    // Icon + Title row
    if (card.icon) {
      pptSlide.addText(card.icon, { x: cx + 0.12, y: cy + 0.08, w: 0.3, h: 0.3, fontSize: 13, align: "center", valign: "middle", margin: 0 });
      pptSlide.addText(card.title, { x: cx + 0.45, y: cy + 0.1, w: cw - 1.6, h: 0.3, fontSize: 11, fontFace: PPT.headerFont, color: PPT.textLight, bold: true, margin: 0 });
    } else {
      pptSlide.addText(card.title, { x: cx + 0.15, y: cy + 0.1, w: cw - 1.3, h: 0.3, fontSize: 11, fontFace: PPT.headerFont, color: PPT.textLight, bold: true, margin: 0 });
    }

    // Status badge
    if (card.status) {
      const statusColor = card.status === "Active" ? PPT.success : card.status === "Complete" ? PPT.accent : card.status === "In Progress" ? PPT.warning : PPT.textMuted;
      pptSlide.addShape(pres.shapes.RECTANGLE, { x: cx + cw - 1.05, y: cy + 0.12, w: 0.9, h: 0.22, fill: { color: statusColor, transparency: 85 }, rectRadius: 0.1 });
      pptSlide.addText(card.status, { x: cx + cw - 1.05, y: cy + 0.12, w: 0.9, h: 0.22, fontSize: 7, fontFace: PPT.bodyFont, color: statusColor, bold: true, align: "center", valign: "middle", margin: 0 });
    }

    // Items
    (card.items || []).forEach((item, j) => {
      const textParts = [];
      if (item.bold) textParts.push({ text: item.bold + " ", options: { bold: true, color: PPT.textMid, fontSize: 9, fontFace: PPT.bodyFont } });
      textParts.push({ text: item.text || (typeof item === "string" ? item : ""), options: { color: PPT.textDim, fontSize: 9, fontFace: PPT.bodyFont } });
      pptSlide.addText(textParts, { x: cx + 0.15, y: cy + 0.45 + j * 0.3, w: cw - 0.3, h: 0.25, margin: 0 });
    });
  });
}

function renderDashboardPPTX(pptSlide, pres, s, idx, total) {
  addSlideChrome(pptSlide, pres, idx, total);
  addPageTitle(pptSlide, s.title, s.subtitle);

  const metrics = s.metrics || [];
  const cols = 3;
  const mw = 2.85;
  const mh = 1.35;
  const gap = 0.2;

  metrics.forEach((m, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const mx = 0.5 + col * (mw + gap);
    const my = 1.55 + row * (mh + gap);

    pptSlide.addShape(pres.shapes.RECTANGLE, { x: mx, y: my, w: mw, h: mh, fill: { color: PPT.bgCard }, line: { color: PPT.border, width: 0.5 }, rectRadius: 0.08 });
    // Bottom accent line
    pptSlide.addShape(pres.shapes.RECTANGLE, { x: mx, y: my + mh - 0.025, w: mw * 0.6, h: 0.025, fill: { color: PPT.accent, transparency: 60 } });

    // Icon
    if (m.icon) {
      pptSlide.addText(m.icon, { x: mx, y: my + 0.08, w: mw, h: 0.3, fontSize: 16, align: "center", valign: "middle", margin: 0 });
      pptSlide.addText(m.value, { x: mx, y: my + 0.35, w: mw, h: 0.35, fontSize: 22, fontFace: PPT.headerFont, color: PPT.accent, bold: true, align: "center", valign: "middle", margin: 0 });
    } else {
      pptSlide.addText(m.value, { x: mx, y: my + 0.15, w: mw, h: 0.5, fontSize: 24, fontFace: PPT.headerFont, color: PPT.accent, bold: true, align: "center", valign: "middle", margin: 0 });
    }
    pptSlide.addText(m.label, { x: mx, y: my + 0.7, w: mw, h: 0.25, fontSize: 9, fontFace: PPT.bodyFont, color: PPT.textMuted, align: "center", valign: "middle", margin: 0 });
    if (m.detail) {
      pptSlide.addText(m.detail, { x: mx, y: my + 0.9, w: mw, h: 0.22, fontSize: 8, fontFace: PPT.bodyFont, color: PPT.success, bold: true, align: "center", valign: "middle", margin: 0 });
    }
  });
}

function renderTablePPTX(pptSlide, pres, s, idx, total) {
  addSlideChrome(pptSlide, pres, idx, total);
  addPageTitle(pptSlide, s.title, s.subtitle);

  const cols = s.columns || [];
  const rows = s.rows || [];
  const colW = cols.map(() => (9 / cols.length));

  const headerRow = cols.map((c) => ({
    text: c, options: { fontSize: 8, fontFace: PPT.bodyFont, color: PPT.textMuted, bold: true, fill: { color: PPT.bgCardLight }, align: "left", margin: [4, 6, 4, 6] }
  }));

  const dataRows = rows.map((row, ri) => (row.cells || []).map((cell, ci) => ({
    text: cell, options: {
      fontSize: 9, fontFace: PPT.bodyFont,
      color: ci === 0 ? PPT.textMid : PPT.textDim,
      bold: ci === 0, fill: { color: ri % 2 === 0 ? PPT.bgCard : PPT.bg },
      align: "left", margin: [4, 6, 4, 6],
    }
  })));

  pptSlide.addTable([headerRow, ...dataRows], {
    x: 0.5, y: 1.55, w: 9, colW,
    border: { type: "solid", pt: 0.5, color: PPT.border },
    autoPage: false,
  });

  // Summary stats
  (s.summaryStats || []).forEach((st, i) => {
    const sx = 0.5 + i * 1.8;
    const sy = 4.6;
    pptSlide.addShape(pres.shapes.RECTANGLE, { x: sx, y: sy, w: 1.6, h: 0.4, fill: { color: PPT.accent, transparency: 85 }, rectRadius: 0.06 });
    pptSlide.addText([
      { text: st.number + "  ", options: { fontSize: 16, bold: true, color: PPT.accent, fontFace: PPT.headerFont } },
      { text: st.label, options: { fontSize: 9, color: PPT.textDim, fontFace: PPT.bodyFont } },
    ], { x: sx + 0.1, y: sy, w: 1.4, h: 0.4, valign: "middle", margin: 0 });
  });
}

function renderSectionOverviewPPTX(pptSlide, pres, s, idx, total) {
  addSlideChrome(pptSlide, pres, idx, total);

  // Section icon box with emoji
  pptSlide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 0.9, w: 0.55, h: 0.55, fill: { color: PPT.accent, transparency: 85 }, line: { color: PPT.accent, width: 0.5, transparency: 70 }, rectRadius: 0.1 });
  if (s.sectionIcon) {
    pptSlide.addText(s.sectionIcon, { x: 0.6, y: 0.9, w: 0.55, h: 0.55, fontSize: 20, align: "center", valign: "middle", margin: 0 });
  }

  pptSlide.addText(s.sectionTitle || s.title, { x: 0.6, y: 1.6, w: 4.5, h: 0.5, fontSize: 22, fontFace: PPT.headerFont, color: PPT.textLight, bold: true, margin: 0 });
  pptSlide.addText(s.description || "", { x: 0.6, y: 2.15, w: 4.5, h: 0.8, fontSize: 10, fontFace: PPT.bodyFont, color: PPT.textDim, margin: 0 });

  // Stats
  (s.stats || []).forEach((st, i) => {
    const sx = 0.6 + i * 1.4;
    pptSlide.addShape(pres.shapes.RECTANGLE, { x: sx, y: 3.1, w: 1.25, h: 0.75, fill: { color: PPT.bgCard }, line: { color: PPT.border, width: 0.5 }, rectRadius: 0.08 });
    pptSlide.addText(st.number, { x: sx, y: 3.15, w: 1.25, h: 0.4, fontSize: 18, fontFace: PPT.headerFont, color: PPT.accent, bold: true, align: "center", margin: 0 });
    pptSlide.addText(st.label, { x: sx, y: 3.55, w: 1.25, h: 0.22, fontSize: 7, fontFace: PPT.bodyFont, color: PPT.textDim, align: "center", margin: 0 });
  });

  // Billing panel (right)
  if (s.billingMetrics?.length > 0) {
    pptSlide.addShape(pres.shapes.RECTANGLE, { x: 5.8, y: 0.9, w: 3.8, h: 3.6, fill: { color: PPT.bgCard }, line: { color: PPT.border, width: 0.5 }, rectRadius: 0.1 });
    pptSlide.addText("BILLING OVERVIEW", { x: 6.0, y: 1.05, w: 3.4, h: 0.3, fontSize: 8, fontFace: PPT.bodyFont, color: PPT.textMuted, bold: true, margin: 0 });

    s.billingMetrics.forEach((bm, i) => {
      const by = 1.5 + i * 0.6;
      pptSlide.addText(bm.label, { x: 6.0, y: by, w: 1.8, h: 0.25, fontSize: 9, fontFace: PPT.bodyFont, color: PPT.textDim, margin: 0 });
      pptSlide.addText(bm.value, { x: 7.8, y: by, w: 1.6, h: 0.25, fontSize: 11, fontFace: PPT.bodyFont, color: PPT.textMid, bold: true, align: "right", margin: 0 });
      if (i < s.billingMetrics.length - 1) {
        pptSlide.addShape(pres.shapes.RECTANGLE, { x: 6.0, y: by + 0.35, w: 3.4, h: 0.005, fill: { color: PPT.border } });
      }
    });
  }
}

function renderHighlightListPPTX(pptSlide, pres, s, idx, total) {
  addSlideChrome(pptSlide, pres, idx, total);
  addPageTitle(pptSlide, s.title, s.subtitle);

  (s.items || []).forEach((item, i) => {
    const iy = 1.55 + i * 0.55;
    const color = h(item.color) || PPT.accent;

    pptSlide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: iy, w: 9, h: 0.45, fill: { color: PPT.bgCard }, line: { color: PPT.border, width: 0.5 }, rectRadius: 0.06 });
    // Left color accent
    pptSlide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: iy, w: 0.04, h: 0.45, fill: { color } });
    // Dot
    pptSlide.addShape(pres.shapes.OVAL, { x: 0.72, y: iy + 0.16, w: 0.1, h: 0.1, fill: { color } });
    // Text
    pptSlide.addText(item.text, { x: 0.95, y: iy + 0.03, w: 8.3, h: 0.38, fontSize: 10, fontFace: PPT.bodyFont, color: PPT.textMid, margin: 0, valign: "middle" });
  });
}

function renderThankYouPPTX(pptSlide, pres, s, idx, total) {
  addSlideChrome(pptSlide, pres, idx, total);

  // Center heart circle (gradient-like effect with two shapes)
  pptSlide.addShape(pres.shapes.OVAL, { x: 4.55, y: 1.2, w: 0.9, h: 0.9, fill: { color: PPT.accent } });
  pptSlide.addText("♥", { x: 4.55, y: 1.2, w: 0.9, h: 0.9, fontSize: 28, color: PPT.headerText, align: "center", valign: "middle", margin: 0 });

  pptSlide.addText(s.title || "Thank You", { x: 1, y: 2.3, w: 8, h: 0.8, fontSize: 42, fontFace: PPT.headerFont, color: PPT.textLight, bold: true, align: "center", margin: 0 });
  pptSlide.addText(s.message || "", { x: 2, y: 3.1, w: 6, h: 0.7, fontSize: 14, fontFace: PPT.bodyFont, color: PPT.textMuted, align: "center", margin: 0 });
  pptSlide.addText((s.signature || "").toUpperCase(), { x: 2, y: 4.0, w: 6, h: 0.3, fontSize: 10, fontFace: PPT.bodyFont, color: PPT.textDim, bold: true, align: "center", charSpacing: 3, margin: 0 });
}

// ─── Main PPTX export function ─────────────────────────────────
async function exportToPPTX(slides, themeT) {
  // Update PPT palette to match current theme
  PPT = getPPT(themeT);
  const PptxGenJS = await loadPptxGenJS();
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";
  pres.author = "PoterPointAI";
  pres.title = "Quarterly Business Review";

  slides.forEach((slide, idx) => {
    const pptSlide = pres.addSlide();
    switch (slide.type) {
      case "title": renderTitlePPTX(pptSlide, pres, slide, idx, slides.length); break;
      case "agenda": renderAgendaPPTX(pptSlide, pres, slide, idx, slides.length); break;
      case "content": renderContentPPTX(pptSlide, pres, slide, idx, slides.length); break;
      case "dashboard": renderDashboardPPTX(pptSlide, pres, slide, idx, slides.length); break;
      case "table": renderTablePPTX(pptSlide, pres, slide, idx, slides.length); break;
      case "section-overview": renderSectionOverviewPPTX(pptSlide, pres, slide, idx, slides.length); break;
      case "highlight-list": renderHighlightListPPTX(pptSlide, pres, slide, idx, slides.length); break;
      case "thank-you": renderThankYouPPTX(pptSlide, pres, slide, idx, slides.length); break;
      default:
        addSlideChrome(pptSlide, pres, idx, slides.length);
        pptSlide.addText(slide.title || "Slide", { x: 1, y: 2, w: 8, h: 1, fontSize: 24, fontFace: PPT.headerFont, color: PPT.textLight, bold: true, align: "center" });
    }
  });

  await pres.writeFile({ fileName: "Deck_Studio_Presentation.pptx" });
}

// ─── Load html2canvas + jsPDF (npm modules) ──────────────────
async function exportToPDF(slides, totalSlides, store) {
  const html2canvasMod = await import("html2canvas");
  const html2canvas = html2canvasMod.default || html2canvasMod;
  const jspdfMod = await import("jspdf");
  const { jsPDF } = jspdfMod;

  // Get current theme bg color for canvas capture
  const storeState = store.getState();
  const themeT = getTheme(storeState.themeMode, storeState.colorCombo);
  const bgColor = themeT.mode === "light" ? "#f0f2f5" : "#070B14";

  // Create an offscreen container at a fixed high-res size for consistent capture
  const container = document.createElement("div");
  container.style.cssText = "position:fixed;left:-9999px;top:0;width:960px;height:540px;z-index:-1;overflow:hidden;";
  document.body.appendChild(container);

  const pdf = new jsPDF({ orientation: "landscape", unit: "px", format: [960, 540] });

  try {
    for (let i = 0; i < slides.length; i++) {
      // Clear container and render each slide via React
      const { createRoot } = await import("react-dom/client");
      const slideWrapper = document.createElement("div");
      slideWrapper.style.cssText = "width:960px;height:540px;position:relative;overflow:hidden;";
      container.innerHTML = "";
      container.appendChild(slideWrapper);

      // Render slide inside StoreContext so useTheme() works
      const root = createRoot(slideWrapper);
      await new Promise((resolve) => {
        root.render(
          <StoreContext.Provider value={store}>
            <SlideRenderer slide={slides[i]} index={i} total={totalSlides} />
          </StoreContext.Provider>
        );
        // Give React a frame to paint
        requestAnimationFrame(() => requestAnimationFrame(resolve));
      });

      // Capture with html2canvas
      const canvas = await html2canvas(slideWrapper, {
        width: 960,
        height: 540,
        scale: 2,
        useCORS: true,
        backgroundColor: bgColor,
        logging: false,
      });

      const imgData = canvas.toDataURL("image/jpeg", 0.92);

      if (i > 0) pdf.addPage([960, 540], "landscape");
      pdf.addImage(imgData, "JPEG", 0, 0, 960, 540);

      // Cleanup React root
      root.unmount();
    }

    pdf.save("Deck_Studio_Presentation.pdf");
  } finally {
    document.body.removeChild(container);
  }
}

// ─── Export Dropdown Component ─────────────────────────────────
function ExportDropdown({ slides, onClose }) {
  const { T, css } = useTheme();
  const store = useContext(StoreContext);
  const [exporting, setExporting] = useState(null);

  const handlePPTX = async () => {
    setExporting("pptx");
    try { await exportToPPTX(slides, T); } catch (e) { console.error(e); alert("PPTX export failed: " + e.message); }
    setExporting(null);
    onClose();
  };

  const handlePDF = async () => {
    setExporting("pdf");
    try { await exportToPDF(slides, slides.length, store); } catch (e) { console.error(e); alert("PDF export failed: " + e.message); }
    setExporting(null);
    onClose();
  };

  const handleJSON = () => {
    const data = JSON.stringify({ presentation: { title: "My Presentation" }, slides }, null, 2);
    const b = new Blob([data], { type: "application/json" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(b);
    a.download = "presentation.json";
    a.click();
    onClose();
  };

  const exportOptions = [
    { id: "pptx", label: "PowerPoint (.pptx)", desc: "Editable presentation file", icon: "📊", action: handlePPTX },
    { id: "pdf", label: "PDF Document (.pdf)", desc: "Pixel-perfect, non-editable", icon: "📄", action: handlePDF },
    { id: "json", label: "JSON Data", desc: "Raw slide data for re-import", icon: "{ }", action: handleJSON },
  ];

  return (
    <div style={{
      position: "absolute", top: "44px", right: "0", zIndex: 300,
      background: T.mode === "light" ? "#ffffff" : "#111827", border: `1px solid ${T.borderHover}`,
      borderRadius: T.radiusLg, padding: "8px", minWidth: "260px",
      boxShadow: "0 20px 60px rgba(0,0,0,0.6)",
    }}>
      <div style={{ padding: "6px 10px 10px", color: T.textMuted, fontSize: "9px", fontWeight: 700, textTransform: "uppercase", letterSpacing: "1.5px" }}>Export As</div>
      {exportOptions.map((opt) => (
        <button key={opt.id} onClick={opt.action} disabled={exporting === opt.id}
          style={{
            display: "flex", alignItems: "center", gap: "12px", width: "100%",
            padding: "10px 12px", border: "none", background: "transparent",
            color: T.text, cursor: exporting ? "wait" : "pointer", borderRadius: T.radiusSm,
            fontSize: "13px", fontWeight: 500, textAlign: "left", fontFamily: T.font,
            opacity: exporting && exporting !== opt.id ? 0.4 : 1,
            transition: "background 0.12s",
          }}
          onMouseEnter={(e) => !exporting && (e.currentTarget.style.background = T.surfaceHover)}
          onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}>
          <span style={{
            width: "32px", height: "32px", borderRadius: "8px",
            background: T.accentGlow, display: "flex", alignItems: "center",
            justifyContent: "center", fontSize: "14px", flexShrink: 0,
          }}>{exporting === opt.id ? "⏳" : opt.icon}</span>
          <div>
            <div style={{ fontWeight: 700, fontSize: "12px" }}>{opt.label}</div>
            <div style={{ color: T.textDim, fontSize: "10px", marginTop: "1px" }}>
              {exporting === opt.id ? "Generating..." : opt.desc}
            </div>
          </div>
        </button>
      ))}
    </div>
  );
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 🖱️  CANVAS EDITOR — Drag & Drop Visual Editor
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

const CANVAS_W = 960;
const CANVAS_H = 540;
const GRID_SIZE = 10;
const snap = (v) => Math.round(v / GRID_SIZE) * GRID_SIZE;

// ─── Canvas Block (single draggable element) ──────────────────

function CanvasBlock({ el, rect, selected, scale, onSelect, onDragEnd, slide, renderContent, onInlineEdit, onMultiDragMove, onMultiDragEnd, isMultiSelected, dragOffset }) {
  const { T } = useTheme();
  const dragRef = useRef(null);
  const [dragPos, setDragPos] = useState(null);

  // Apply external drag offset (from multi-drag) or local drag
  const pos = dragOffset ? { ...rect, x: rect.x + dragOffset.dx, y: rect.y + dragOffset.dy } : dragPos || rect;

  const handlePointerDown = useCallback((e) => {
    if (e.button !== 0) return;
    e.stopPropagation();
    try { e.target.setPointerCapture(e.pointerId); } catch {}
    dragRef.current = { startX: e.clientX, startY: e.clientY, origX: rect.x, origY: rect.y };
    // Shift+click = toggle multi-select; normal click = single select
    if (e.shiftKey) {
      onSelect(el.id, "toggle");
    } else if (!isMultiSelected) {
      onSelect(el.id, "single");
    }
    // If multi-selected, notify parent to start multi-drag
    if (isMultiSelected || e.shiftKey) {
      onMultiDragMove?.("start", e.clientX, e.clientY);
    }
  }, [rect, el.id, onSelect, isMultiSelected, onMultiDragMove]);

  const [editing, setEditing] = useState(false);
  const editRef = useRef(null);

  const handleClick = useCallback((e) => {
    e.stopPropagation();
    if (!e.shiftKey && !isMultiSelected) onSelect(el.id, "single");
  }, [el.id, onSelect, isMultiSelected]);

  const handleDoubleClick = useCallback((e) => {
    e.stopPropagation();
    const editableFields = ["title", "subtitle", "badge", "message", "signature", "description", "page-title", "section-title"];
    // Also allow editing sub-element titles and labels
    const isSubTitle = el.id.match(/\.\w*(title|label|value)\b/i);
    if (editableFields.includes(el.id) || isSubTitle) {
      setEditing(true);
      setTimeout(() => editRef.current?.focus(), 50);
    }
  }, [el.id]);

  const handleInlineBlur = useCallback(() => {
    if (editing && editRef.current && onInlineEdit) {
      const fieldMap = { "page-title": "title", "section-title": "sectionTitle" };
      const fieldName = fieldMap[el.id] || el.id;
      onInlineEdit(fieldName, editRef.current.innerText);
    }
    setEditing(false);
  }, [editing, el.id, onInlineEdit]);

  const handlePointerMove = useCallback((e) => {
    if (!dragRef.current) return;
    const dx = (e.clientX - dragRef.current.startX) / scale;
    const dy = (e.clientY - dragRef.current.startY) / scale;
    if (isMultiSelected) {
      // Multi-drag: notify parent of delta
      onMultiDragMove?.("move", dx, dy);
    } else {
      setDragPos({ ...rect, x: snap(dragRef.current.origX + dx), y: snap(dragRef.current.origY + dy) });
    }
  }, [rect, scale, isMultiSelected, onMultiDragMove]);

  const handlePointerUp = useCallback((e) => {
    if (!dragRef.current) return;
    const dx = (e.clientX - dragRef.current.startX) / scale;
    const dy = (e.clientY - dragRef.current.startY) / scale;
    dragRef.current = null;
    if (isMultiSelected) {
      onMultiDragEnd?.(snap(dx), snap(dy));
    } else {
      setDragPos(null);
      onDragEnd(el.id, { x: snap(rect.x + dx), y: snap(rect.y + dy) });
    }
  }, [scale, el.id, onDragEnd, rect, isMultiSelected, onMultiDragEnd]);

  const cornerStyle = (cursor) => ({
    position: "absolute", width: "8px", height: "8px",
    background: T.accent, borderRadius: "2px", cursor,
    border: "1px solid rgba(255,255,255,0.5)",
  });

  return (
    <div
      onPointerDown={handlePointerDown}
      onPointerMove={handlePointerMove}
      onPointerUp={handlePointerUp}
      onClick={handleClick}
      onDoubleClick={handleDoubleClick}
      style={{
        position: "absolute",
        left: pos.x, top: pos.y, width: pos.w, height: pos.h,
        cursor: dragRef.current ? "grabbing" : "grab",
        border: selected ? `2px solid ${T.accent}` : isMultiSelected ? `2px dashed ${T.accent}80` : `1px solid transparent`,
        borderRadius: el.subType === "card-bg" ? "10px" : "4px",
        background: selected ? `${T.accent}0a` : isMultiSelected ? `${T.accent}05` : "transparent",
        transition: dragRef.current ? "none" : "all 0.12s ease",
        zIndex: selected ? 10 : isMultiSelected ? 5 : (el.subType === "card-bg" ? 0 : 1),
        outline: selected ? `1px solid ${T.accent}30` : "none",
        outlineOffset: "2px",
        boxSizing: "border-box",
      }}
    >
      {/* Label tooltip */}
      {(selected || isMultiSelected) && (
        <div style={{
          position: "absolute", top: "-18px", left: "0",
          background: selected ? T.accent : `${T.accent}90`, color: "#fff", fontSize: "8px",
          fontWeight: 600, padding: "1px 6px", borderRadius: "3px",
          whiteSpace: "nowrap", fontFamily: T.fontMono, pointerEvents: "none",
          boxShadow: "0 2px 6px rgba(0,0,0,0.15)",
          maxWidth: "120px", overflow: "hidden", textOverflow: "ellipsis",
        }}>{el.label}</div>
      )}
      {/* Corner handles */}
      {selected && (
        <>
          <div style={{ ...cornerStyle("nw-resize"), top: "-4px", left: "-4px" }} />
          <div style={{ ...cornerStyle("ne-resize"), top: "-4px", right: "-4px" }} />
          <div style={{ ...cornerStyle("sw-resize"), bottom: "-4px", left: "-4px" }} />
          <div style={{ ...cornerStyle("se-resize"), bottom: "-4px", right: "-4px" }} />
        </>
      )}
      {/* Render actual content — editable on double-click */}
      {editing ? (
        <div ref={editRef} contentEditable suppressContentEditableWarning
          onBlur={handleInlineBlur}
          onKeyDown={(e) => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); handleInlineBlur(); } if (e.key === "Escape") setEditing(false); }}
          style={{ width: "100%", height: "100%", overflow: "auto", outline: "none", padding: "4px",
            fontSize: el.id === "title" || el.id === "page-title" || el.id === "section-title" ? "28px" : el.id === "subtitle" ? "14px" : "13px",
            fontWeight: el.id === "title" || el.id === "page-title" || el.id === "section-title" ? 800 : 400,
            color: T.slideText, fontFamily: T.font, cursor: "text", background: "rgba(59,130,246,0.08)", borderRadius: "4px",
          }}>
          {slide[el.id === "page-title" ? "title" : el.id === "section-title" ? "sectionTitle" : el.id]}
        </div>
      ) : renderContent ? (
        <div style={{ width: "100%", height: "100%", overflow: "hidden", pointerEvents: "none", fontSize: "inherit" }}>
          {renderContent(el, slide)}
        </div>
      ) : null}
    </div>
  );
}

// ─── Canvas Toolbar ───────────────────────────────────────────

const FONT_FAMILIES = ["Outfit", "Inter", "Roboto", "Open Sans", "Lato", "Montserrat", "DM Sans", "Arial", "Calibri"];
const FONT_SIZES = [10, 11, 12, 13, 14, 16, 18, 20, 22, 24, 28, 32, 36, 42, 48];

const ICON_CATEGORIES = {
  business: ["📊", "📈", "📉", "💼", "🏢", "🤝", "📋", "📁", "💰", "💵", "🏦", "📄"],
  finance: ["💳", "💲", "🏧", "💹", "🪙", "🧾"],
  technology: ["💻", "⚙️", "🔧", "🌐", "📱", "🖥️", "☁️", "🔒", "🛡️", "🤖", "🧠", "⚡"],
  team: ["👥", "👤", "🧑‍💼", "👨‍💻", "👩‍💻", "🏆", "⭐", "🎯", "🎓", "💪"],
  awards: ["🏆", "🥇", "🥈", "🥉", "🎖️", "🏅", "⭐", "🌟", "✨", "🎉"],
};

function TextToolbar({ selectedId, slide, onUpdateStyle }) {
  const { T, css } = useTheme();
  const style = slide?.elementStyles?.[selectedId] || {};
  if (!selectedId) return null;

  const update = (key, val) => onUpdateStyle(selectedId, { ...style, [key]: val });

  return (
    <div style={{ display: "flex", alignItems: "center", gap: "6px", padding: "4px 0" }}>
      <select value={style.fontFamily || "Outfit"} onChange={(e) => update("fontFamily", e.target.value)}
        style={{ ...css.input, width: "110px", fontSize: "10px", padding: "3px 6px" }}>
        {FONT_FAMILIES.map(f => <option key={f} value={f}>{f}</option>)}
      </select>
      <select value={style.fontSize || 14} onChange={(e) => update("fontSize", Number(e.target.value))}
        style={{ ...css.input, width: "55px", fontSize: "10px", padding: "3px 6px" }}>
        {FONT_SIZES.map(s => <option key={s} value={s}>{s}px</option>)}
      </select>
      <button onClick={() => update("fontWeight", style.fontWeight === "bold" ? "normal" : "bold")}
        style={{ ...css.btnIcon, width: "26px", height: "26px", fontSize: "12px", fontWeight: 900, background: style.fontWeight === "bold" ? T.accentGlow : "transparent" }}>B</button>
      <button onClick={() => update("fontStyle", style.fontStyle === "italic" ? "normal" : "italic")}
        style={{ ...css.btnIcon, width: "26px", height: "26px", fontSize: "12px", fontStyle: "italic", background: style.fontStyle === "italic" ? T.accentGlow : "transparent" }}>I</button>
    </div>
  );
}

function IconPicker({ onSelect, onClose }) {
  const { T, css } = useTheme();
  const [activeCategory, setActiveCategory] = useState("business");
  const [search, setSearch] = useState("");

  const icons = ICON_CATEGORIES[activeCategory] || [];

  return (
    <div style={{ position: "absolute", top: "32px", left: 0, zIndex: 300, background: T.mode === "light" ? "#fff" : "#111827", border: `1px solid ${T.borderHover}`, borderRadius: T.radiusSm, padding: "10px", width: "260px", boxShadow: "0 12px 40px rgba(0,0,0,0.3)" }}>
      <div style={{ display: "flex", justifyContent: "space-between", marginBottom: "8px" }}>
        <span style={{ fontSize: "11px", fontWeight: 700, color: T.text }}>Icon Picker</span>
        <button onClick={onClose} style={{ background: "none", border: "none", color: T.textMuted, cursor: "pointer", fontSize: "12px" }}>✕</button>
      </div>
      <div style={{ display: "flex", gap: "3px", marginBottom: "8px", flexWrap: "wrap" }}>
        {Object.keys(ICON_CATEGORIES).map(c => (
          <button key={c} onClick={() => setActiveCategory(c)} style={{
            padding: "3px 8px", border: "none", cursor: "pointer", borderRadius: "4px", fontSize: "9px", fontWeight: 600,
            background: activeCategory === c ? T.accentGlow : "transparent", color: activeCategory === c ? T.accent : T.textMuted,
          }}>{c}</button>
        ))}
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(8, 1fr)", gap: "4px" }}>
        {icons.map((icon, i) => (
          <button key={i} onClick={() => { onSelect(icon); onClose(); }} style={{
            width: "28px", height: "28px", border: `1px solid ${T.border}`, borderRadius: "4px", cursor: "pointer",
            background: "transparent", fontSize: "16px", display: "flex", alignItems: "center", justifyContent: "center",
          }}>{icon}</button>
        ))}
      </div>
    </div>
  );
}

function CanvasToolbar({ elementCount, snapEnabled, onToggleSnap, onReset, selectedId, slide, onUpdateStyle, selectedCount, onGroup, onUngroup }) {
  const { T, css } = useTheme();
  const [showIcons, setShowIcons] = useState(false);
  return (
    <div style={{
      display: "flex", alignItems: "center", gap: "6px",
      padding: "5px 12px", background: T.mode === "light" ? T.surface : "rgba(255,255,255,0.03)",
      borderBottom: `1px solid ${T.border}`, fontSize: "11px", flexWrap: "wrap",
      minHeight: "36px",
    }}>
      <span style={{ color: T.textMuted, fontFamily: T.fontMono, fontSize: "9px", background: T.mode === "light" ? "#f1f5f9" : "rgba(255,255,255,0.06)", padding: "2px 8px", borderRadius: "4px" }}>
        {elementCount}
      </span>

      {/* Font controls when element selected */}
      {selectedId && <TextToolbar selectedId={selectedId} slide={slide} onUpdateStyle={onUpdateStyle} />}

      {/* Icon picker — shown for icon elements or any element to add/change icons */}
      {selectedId && (
        <div style={{ position: "relative" }}>
          <button onClick={() => setShowIcons(!showIcons)} style={{ ...css.btnIcon, width: "auto", padding: "0 8px", height: "26px", fontSize: "10px", fontWeight: 600, fontFamily: T.font }}>
            😊 Icon
          </button>
          {showIcons && <IconPicker onSelect={(newIcon) => {
            // Determine which slide data to update based on selected element ID
            // Pattern: field[N].icon or field[N].anything → update field[N].icon
            const subMatch = selectedId.match(/^(\w+)\[(\d+)\]/);
            if (subMatch) {
              const [, field, idx] = subMatch;
              const arr = [...(slide[field] || [])];
              if (arr[parseInt(idx)]) {
                arr[parseInt(idx)] = { ...arr[parseInt(idx)], icon: newIcon };
                onUpdateStyle("__slideData__", { [field]: arr });
              }
            } else if (selectedId === "section-icon") {
              onUpdateStyle("__slideData__", { sectionIcon: newIcon });
            } else {
              // For standalone elements, store in elementStyles
              onUpdateStyle(selectedId, { icon: newIcon });
            }
            setShowIcons(false);
          }} onClose={() => setShowIcons(false)} />}
        </div>
      )}

      {/* Bullet style picker */}
      {selectedId && selectedId.includes(".items[") && (
        <div style={{ display: "flex", gap: "2px", alignItems: "center" }}>
          <span style={{ fontSize: "9px", color: T.textMuted, marginRight: "4px" }}>Bullet:</span>
          {["•", "→", "▸", "◆", "✓", "★", "›", "–"].map(b => (
            <button key={b} onClick={() => onUpdateStyle("__bulletStyle__", { bulletStyle: b })}
              style={{
                width: "22px", height: "22px", border: `1px solid ${slide?.bulletStyle === b ? T.accent : T.border}`,
                borderRadius: "3px", cursor: "pointer", fontSize: "11px",
                background: slide?.bulletStyle === b ? T.accentGlow : "transparent",
                color: T.text, display: "flex", alignItems: "center", justifyContent: "center",
              }}>{b}</button>
          ))}
        </div>
      )}

      <div style={{ flex: 1 }} />
      <button onClick={onToggleSnap}
        style={{
          ...css.btnIcon, width: "auto", padding: "0 10px", height: "26px",
          fontSize: "10px", fontWeight: 600, fontFamily: T.font,
          background: snapEnabled ? T.accentGlow : "transparent",
          borderColor: snapEnabled ? "rgba(59,130,246,0.3)" : T.border,
        }}>
        ⊞ Snap {snapEnabled ? "ON" : "OFF"}
      </button>
      <button onClick={onReset}
        style={{ ...css.btnIcon, width: "auto", padding: "0 10px", height: "26px", fontSize: "10px", fontWeight: 600, fontFamily: T.font }}>
        ↺ Reset Layout
      </button>
      {/* Selection count + Group/Ungroup */}
      {selectedCount > 1 && (
        <>
          <span style={{ color: T.accent, fontSize: "10px", fontWeight: 700, fontFamily: T.fontMono }}>{selectedCount} selected</span>
          <button onClick={onGroup}
            style={{ ...css.btnIcon, width: "auto", padding: "0 10px", height: "26px", fontSize: "10px", fontWeight: 600, fontFamily: T.font, color: T.accent }}>
            ⊞ Group
          </button>
          <button onClick={onUngroup}
            style={{ ...css.btnIcon, width: "auto", padding: "0 10px", height: "26px", fontSize: "10px", fontWeight: 600, fontFamily: T.font }}>
            ⊟ Ungroup
          </button>
        </>
      )}
    </div>
  );
}

// ─── Canvas Editor (main component) ───────────────────────────

function CanvasEditor({ slide, slideIndex, totalSlides }) {
  const { T } = useTheme();
  const { updateLayout, resetLayout, setCanvasSelectedId, setCanvasSelectedIds, toggleCanvasSelection, updateElementStyle, updateSlide: updateSlideAction, groupElements, ungroupElements } = useStoreActions();
  const containerRef = useRef(null);
  const [scale, setScale] = useState(1);
  const selectedId = useStore((s) => s.canvasSelectedId);
  const selectedIds = useStore((s) => s.canvasSelectedIds) || [];
  const [snapEnabled, setSnapEnabled] = useState(true);
  const [multiDragOffset, setMultiDragOffset] = useState(null); // {dx, dy}
  const [marquee, setMarquee] = useState(null); // {x1,y1,x2,y2}
  const marqueeRef = useRef(null);

  // Measure container and compute scale
  useEffect(() => {
    const measure = () => {
      if (!containerRef.current) return;
      setScale(containerRef.current.getBoundingClientRect().width / CANVAS_W);
    };
    measure();
    const ro = new ResizeObserver(measure);
    if (containerRef.current) ro.observe(containerRef.current);
    return () => ro.disconnect();
  }, []);

  const elements = useMemo(() => generateCanvasElements(slide), [slide]);

  const getRect = useCallback((el) => {
    const override = slide.layout?.[el.id];
    return override ? { ...el.defaultRect, ...override } : el.defaultRect;
  }, [slide.layout]);

  // Selection handler — supports single, toggle (shift), and clear
  const handleSelect = useCallback((id, mode) => {
    if (mode === "toggle") {
      toggleCanvasSelection(id);
    } else {
      setCanvasSelectedIds([id]);
    }
  }, [setCanvasSelectedIds, toggleCanvasSelection]);

  // Multi-drag handler — called by any selected block during drag
  const handleMultiDragMove = useCallback((action, x, y) => {
    if (action === "start") {
      marqueeRef.current = { startX: x, startY: y };
    } else if (action === "move") {
      setMultiDragOffset({ dx: snap(x), dy: snap(y) });
    }
  }, []);

  const handleMultiDragEnd = useCallback((dx, dy) => {
    // Apply delta to all selected elements
    selectedIds.forEach((eid) => {
      const el = elements.find(e => e.id === eid);
      if (!el) return;
      const rect = getRect(el);
      updateLayout(slide.id, eid, { x: rect.x + dx, y: rect.y + dy });
    });
    setMultiDragOffset(null);
    marqueeRef.current = null;
  }, [selectedIds, elements, getRect, slide.id, updateLayout]);

  const handleDragEnd = useCallback((elementId, pos) => {
    updateLayout(slide.id, elementId, pos);
  }, [slide.id, updateLayout]);

  // Marquee selection — drag on background
  const handleBgPointerDown = useCallback((e) => {
    if (e.target !== e.currentTarget) return;
    const bounds = e.currentTarget.getBoundingClientRect();
    const x = (e.clientX - bounds.left) / scale;
    const y = (e.clientY - bounds.top) / scale;
    marqueeRef.current = { startX: e.clientX, startY: e.clientY, cx: x, cy: y };
    setMarquee({ x1: x, y1: y, x2: x, y2: y });
  }, [scale]);

  const handleBgPointerMove = useCallback((e) => {
    if (!marqueeRef.current?.cx) return;
    const bounds = e.currentTarget.getBoundingClientRect();
    const x = (e.clientX - bounds.left) / scale;
    const y = (e.clientY - bounds.top) / scale;
    setMarquee({ x1: marqueeRef.current.cx, y1: marqueeRef.current.cy, x2: x, y2: y });
  }, [scale]);

  const handleBgPointerUp = useCallback((e) => {
    if (marquee) {
      const left = Math.min(marquee.x1, marquee.x2);
      const top = Math.min(marquee.y1, marquee.y2);
      const right = Math.max(marquee.x1, marquee.x2);
      const bottom = Math.max(marquee.y1, marquee.y2);
      const w = right - left, h = bottom - top;
      if (w > 5 && h > 5) {
        // Select elements inside marquee
        const inside = elements.filter(el => {
          const r = getRect(el);
          return r.x >= left && r.y >= top && r.x + r.w <= right && r.y + r.h <= bottom;
        }).map(el => el.id);
        if (inside.length > 0) setCanvasSelectedIds(inside);
        else setCanvasSelectedIds([]);
      } else {
        setCanvasSelectedIds([]);
      }
    }
    setMarquee(null);
    marqueeRef.current = null;
  }, [marquee, elements, getRect, setCanvasSelectedIds]);

  const handleBackgroundClick = useCallback((e) => {
    if (e.target === e.currentTarget) setCanvasSelectedIds([]);
  }, [setCanvasSelectedIds]);

  // Keyboard: Escape, Delete, Ctrl+A, Arrow nudge (multi-aware)
  useEffect(() => {
    const handler = (e) => {
      if (e.key === "Escape") { setCanvasSelectedIds([]); return; }
      // Ctrl/Cmd + A = select all
      if ((e.ctrlKey || e.metaKey) && e.key === "a") {
        e.preventDefault();
        setCanvasSelectedIds(elements.map(el => el.id));
        return;
      }
      // Delete = reset selected to default positions
      if (e.key === "Delete" && selectedIds.length > 0) {
        selectedIds.forEach(eid => {
          const el = elements.find(e => e.id === eid);
          if (el) updateLayout(slide.id, eid, el.defaultRect);
        });
        setCanvasSelectedIds([]);
        return;
      }
      // Arrow nudge — all selected elements
      const nudges = { ArrowUp: [0, -GRID_SIZE], ArrowDown: [0, GRID_SIZE], ArrowLeft: [-GRID_SIZE, 0], ArrowRight: [GRID_SIZE, 0] };
      if (selectedIds.length > 0 && nudges[e.key]) {
        e.preventDefault();
        const [dx, dy] = nudges[e.key];
        selectedIds.forEach(eid => {
          const el = elements.find(el => el.id === eid);
          if (!el) return;
          const rect = getRect(el);
          updateLayout(slide.id, eid, { x: rect.x + dx, y: rect.y + dy });
        });
      }
    };
    window.addEventListener("keydown", handler);
    return () => window.removeEventListener("keydown", handler);
  }, [selectedIds, elements, slide, updateLayout, getRect, setCanvasSelectedIds]);

  // Group/Ungroup
  const handleGroup = useCallback(() => {
    if (selectedIds.length < 2) return;
    const gid = `group_${Date.now()}`;
    groupElements(slide.id, selectedIds, gid);
  }, [selectedIds, slide.id, groupElements]);

  const handleUngroup = useCallback(() => {
    const groups = slide.elementGroups || {};
    const gids = new Set(selectedIds.map(id => groups[id]).filter(Boolean));
    gids.forEach(gid => ungroupElements(slide.id, gid));
  }, [selectedIds, slide, ungroupElements]);

  // Multi-selection bounding box
  const multiBBox = useMemo(() => {
    if (selectedIds.length < 2) return null;
    let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
    selectedIds.forEach(eid => {
      const el = elements.find(e => e.id === eid);
      if (!el) return;
      const r = getRect(el);
      minX = Math.min(minX, r.x); minY = Math.min(minY, r.y);
      maxX = Math.max(maxX, r.x + r.w); maxY = Math.max(maxY, r.y + r.h);
    });
    const dx = multiDragOffset?.dx || 0, dy = multiDragOffset?.dy || 0;
    return { x: minX + dx, y: minY + dy, w: maxX - minX, h: maxY - minY };
  }, [selectedIds, elements, getRect, multiDragOffset]);

  return (
    <div style={{ display: "flex", flexDirection: "column", flex: 1, overflow: "hidden" }}>
      <CanvasToolbar
        elementCount={elements.length}
        snapEnabled={snapEnabled}
        onToggleSnap={() => setSnapEnabled(!snapEnabled)}
        onReset={() => { resetLayout(slide.id); setCanvasSelectedIds([]); }}
        selectedId={selectedId}
        slide={slide}
        onUpdateStyle={(elemId, styles) => {
          if (elemId === "__bulletStyle__" || elemId === "__slideData__") updateSlideAction(slide.id, styles);
          else updateElementStyle(slide.id, elemId, styles);
        }}
        selectedCount={selectedIds.length}
        onGroup={handleGroup}
        onUngroup={handleUngroup}
      />
      <div style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center", padding: "28px", overflow: "auto" }}>
        <div ref={containerRef}
          style={{ width: "100%", maxWidth: "860px", aspectRatio: "16/9", position: "relative", borderRadius: "14px", overflow: "hidden",
            boxShadow: `0 24px 80px rgba(0,0,0,0.45), 0 0 0 1px ${T.border}`,
          }}>
          {/* Background only — gradient + corner accents, no header/footer chrome (those are canvas elements now) */}
          <div style={{ position: "absolute", inset: 0, pointerEvents: "none" }}>
            <div style={{ width: "100%", height: "100%", background: T.slideBg, position: "relative" }}>
              <div style={{ position: "absolute", inset: 0, background: T.glow }} />
              <div style={{ position: "absolute", top: 0, left: 0, width: "80px", height: "3px", background: `linear-gradient(90deg, ${T.accent}, transparent)`, opacity: 0.4 }} />
              <div style={{ position: "absolute", top: 0, left: 0, width: "3px", height: "80px", background: `linear-gradient(180deg, ${T.accent}, transparent)`, opacity: 0.4 }} />
              <div style={{ position: "absolute", bottom: 0, right: 0, width: "80px", height: "3px", background: `linear-gradient(270deg, ${T.purple}, transparent)`, opacity: 0.3 }} />
              <div style={{ position: "absolute", bottom: 0, right: 0, width: "3px", height: "80px", background: `linear-gradient(0deg, ${T.purple}, transparent)`, opacity: 0.3 }} />
            </div>
          </div>

          {/* Scaled canvas layer */}
          <div
            onClick={handleBackgroundClick}
            onPointerDown={handleBgPointerDown}
            onPointerMove={handleBgPointerMove}
            onPointerUp={handleBgPointerUp}
            style={{
              position: "absolute", inset: 0,
              width: CANVAS_W, height: CANVAS_H,
              transform: `scale(${scale})`, transformOrigin: "top left",
              backgroundImage: `linear-gradient(rgba(59,130,246,0.08) 1px, transparent 1px), linear-gradient(90deg, rgba(59,130,246,0.08) 1px, transparent 1px)`,
              backgroundSize: `${GRID_SIZE}px ${GRID_SIZE}px`,
            }}
          >
            {elements.map((el) => (
              <CanvasBlock
                key={el.id}
                el={el}
                rect={getRect(el)}
                selected={selectedId === el.id && selectedIds.length <= 1}
                isMultiSelected={selectedIds.includes(el.id) && selectedIds.length > 1}
                scale={scale}
                onSelect={handleSelect}
                onDragEnd={handleDragEnd}
                slide={slide}
                renderContent={(el, s) => renderCanvasContent(el, s, slide, T)}
                onInlineEdit={(field, value) => updateSlideAction(slide.id, { [field]: value })}
                onMultiDragMove={selectedIds.includes(el.id) ? handleMultiDragMove : undefined}
                onMultiDragEnd={selectedIds.includes(el.id) ? handleMultiDragEnd : undefined}
                dragOffset={selectedIds.includes(el.id) && multiDragOffset ? multiDragOffset : null}
              />
            ))}

            {/* Multi-selection bounding box */}
            {multiBBox && (
              <div style={{
                position: "absolute", left: multiBBox.x, top: multiBBox.y, width: multiBBox.w, height: multiBBox.h,
                border: `2px dashed ${T.accent}`, borderRadius: "4px", pointerEvents: "none", zIndex: 20,
              }} />
            )}

            {/* Marquee selection rectangle */}
            {marquee && (
              <div style={{
                position: "absolute",
                left: Math.min(marquee.x1, marquee.x2), top: Math.min(marquee.y1, marquee.y2),
                width: Math.abs(marquee.x2 - marquee.x1), height: Math.abs(marquee.y2 - marquee.y1),
                background: "rgba(59,130,246,0.1)", border: `1px solid ${T.accent}`,
                borderRadius: "2px", pointerEvents: "none", zIndex: 50,
              }} />
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 🎬  PRESENTATION MODE
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function PresentationMode({ slides, onClose, startSlide = 0 }) {
  const { T, css } = useTheme();
  const [cur, setCur] = useState(startSlide);
  useEffect(() => {
    const h = (e) => {
      if (e.key === "ArrowRight" || e.key === " ") { e.preventDefault(); setCur((p) => Math.min(p + 1, slides.length - 1)); }
      else if (e.key === "ArrowLeft") setCur((p) => Math.max(p - 1, 0));
      else if (e.key === "Escape") onClose();
    };
    window.addEventListener("keydown", h);
    return () => window.removeEventListener("keydown", h);
  }, [slides.length, onClose]);

  return (
    <div style={{ position: "fixed", inset: 0, zIndex: 9999, background: "#000", display: "flex", flexDirection: "column" }}>
      <div style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center", padding: "24px" }}>
        <div style={{ width: "100%", maxWidth: "1200px", aspectRatio: "16/9", borderRadius: "12px", overflow: "hidden", boxShadow: "0 24px 80px rgba(0,0,0,0.6)" }}>
          <SlideRenderer slide={slides[cur]} index={cur} total={slides.length} />
        </div>
      </div>
      <div style={{ display: "flex", justifyContent: "center", alignItems: "center", gap: "14px", padding: "14px", background: "rgba(0,0,0,0.85)", borderTop: `1px solid ${T.border}` }}>
        <button onClick={() => setCur((p) => Math.max(0, p - 1))} disabled={cur === 0}
          style={{ ...css.btnIcon, opacity: cur === 0 ? 0.3 : 1, color: "#fff", borderColor: "rgba(255,255,255,0.15)" }}>←</button>
        <span style={{ color: T.textMuted, fontSize: "13px", fontFamily: T.fontMono, minWidth: "60px", textAlign: "center" }}>{cur + 1} / {slides.length}</span>
        <button onClick={() => setCur((p) => Math.min(slides.length - 1, p + 1))} disabled={cur === slides.length - 1}
          style={{ ...css.btnIcon, opacity: cur === slides.length - 1 ? 0.3 : 1, color: "#fff", borderColor: "rgba(255,255,255,0.15)" }}>→</button>
        <div style={{ width: "1px", height: "20px", background: "rgba(255,255,255,0.1)", margin: "0 6px" }} />
        <button onClick={onClose} style={{ ...css.btnIcon, borderColor: "rgba(239,68,68,0.3)", color: T.danger }}>✕</button>
      </div>
    </div>
  );
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 🖥️  SLIDE TYPE SELECTOR
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function SlideTypeSelector({ onSelect, onClose }) {
  const { T, css } = useTheme();
  return (
    <div style={{
      position: "absolute", top: "36px", left: "0", zIndex: 200,
      background: T.mode === "light" ? "#ffffff" : "#111827", border: `1px solid ${T.borderHover}`,
      borderRadius: T.radiusLg, padding: "8px", minWidth: "260px",
      boxShadow: "0 20px 60px rgba(0,0,0,0.6)",
    }}>
      <div style={{ padding: "6px 10px 10px", color: T.textMuted, fontSize: "9px", fontWeight: 700, textTransform: "uppercase", letterSpacing: "1.5px" }}>Add Slide</div>
      {Object.entries(SLIDE_SCHEMA).map(([key, schema]) => (
        <button key={key} onClick={() => { onSelect(key); onClose(); }}
          style={{
            display: "flex", alignItems: "center", gap: "12px", width: "100%",
            padding: "10px 12px", border: "none", background: "transparent",
            color: T.text, cursor: "pointer", borderRadius: T.radiusSm,
            fontSize: "13px", fontWeight: 500, textAlign: "left", fontFamily: T.font,
            transition: "background 0.12s",
          }}
          onMouseEnter={(e) => (e.currentTarget.style.background = T.surfaceHover)}
          onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}>
          <span style={{
            width: "32px", height: "32px", borderRadius: "8px",
            background: T.accentGlow, display: "flex", alignItems: "center",
            justifyContent: "center", fontSize: "15px", flexShrink: 0,
          }}>{schema.icon}</span>
          <div>
            <div style={{ fontWeight: 700, fontSize: "12px" }}>{schema.label}</div>
            <div style={{ color: T.textDim, fontSize: "10px", marginTop: "1px" }}>{schema.description}</div>
          </div>
        </button>
      ))}
    </div>
  );
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 🤖  AI SLIDE GENERATOR
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

const PROMPT_TEMPLATES = [
  { label: "Enterprise QBR", icon: "🏢", prompt: "Create a Microsoft Practice Quarterly Report for Q1 2026. Include: Title slide (7 resources, 4 projects, 3 new customers, $105K revenue), Agenda (8 sections), Resource vs Billing table (Jan-Mar breakdown), New Engagements table (4 projects), Performance Dashboard (6 KPIs: Team 7, Revenue $105K +18%, Utilization 94%, CSAT 4.8, 12 delivered, 98% retention), Process Improvements (2 cards), Sales Pipeline table (6 enquiries), Blogs & Webinars table, Certifications dashboard, Key Highlights list, and Thank You slide." },
  { label: "Quarterly Report", icon: "📊", prompt: "Create a quarterly business review presentation for Q1 2026 with 7 resources, 4 active projects, 3 new customers onboarded. Include sections for resource billing, new engagements, process improvements, major achievements, sales enquiries, training, and team events." },
  { label: "Startup Pitch", icon: "🚀", prompt: "Create a startup pitch deck for an AI-powered SaaS analytics platform called DataPulse. Include problem, solution, market size ($4.2B TAM), traction (500+ customers, $2M ARR), team highlights, and funding ask ($5M Series A)." },
  { label: "Sales Deck", icon: "💼", prompt: "Create a sales presentation for a cloud migration consulting service. Include company overview, service offerings (Azure, AWS, GCP), case studies with metrics, pricing tiers, client testimonials, and a call to action." },
  { label: "Project Update", icon: "📋", prompt: "Create a project status update presentation for a digital transformation initiative. Include project timeline, milestones completed (60%), team allocation (12 resources), budget utilization (75%), risks & mitigations, and next steps." },
];

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 📚 TEMPLATE LIBRARY — Gamma-style template gallery
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

const TEMPLATE_CATEGORIES = ["All", "Business", "Engineering", "Sales", "Leadership", "Custom"];

function TemplateGalleryModal({ onClose, onSelectTemplate }) {
  const { T, css } = useTheme();
  const [templates, setTemplates] = useState([]);
  const [category, setCategory] = useState("All");
  const [search, setSearch] = useState("");
  const [preview, setPreview] = useState(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    const params = new URLSearchParams();
    if (category !== "All") params.set("category", category);
    if (search) params.set("search", search);
    fetch(`/api/templates/library/all?${params}`).then(r => r.json()).then(d => {
      setTemplates(d.templates || []);
      setLoading(false);
    }).catch(() => setLoading(false));
  }, [category, search]);

  const handleSelect = (tmpl) => {
    onSelectTemplate(tmpl);
    onClose();
  };

  return (
    <div style={{ position: "fixed", inset: 0, zIndex: 9999, background: "rgba(0,0,0,0.6)", backdropFilter: "blur(8px)", display: "flex", alignItems: "center", justifyContent: "center" }}
      onClick={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <div style={{ background: T.mode === "light" ? "#fff" : "#111827", borderRadius: T.radiusLg, width: "820px", maxWidth: "95vw", maxHeight: "85vh", display: "flex", flexDirection: "column", border: `1px solid ${T.borderHover}`, boxShadow: "0 24px 80px rgba(0,0,0,0.5)" }}>

        {/* Header */}
        <div style={{ padding: "20px 24px 0", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
            <span style={{ fontSize: "22px" }}>📚</span>
            <div><div style={{ fontSize: "16px", fontWeight: 800, color: T.text }}>Template Library</div>
              <div style={{ fontSize: "11px", color: T.textMuted }}>Select a template to structure your presentation</div></div>
          </div>
          <button onClick={onClose} style={{ ...css.btnIcon, fontSize: "16px" }}>✕</button>
        </div>

        {/* Search + Filters */}
        <div style={{ padding: "16px 24px 0", display: "flex", gap: "12px", alignItems: "center", flexWrap: "wrap" }}>
          <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search templates..."
            style={{ ...css.input, width: "200px", flex: "none", fontSize: "12px", padding: "7px 12px" }} />
          <div style={{ display: "flex", gap: "4px", flexWrap: "wrap" }}>
            {TEMPLATE_CATEGORIES.map(c => (
              <button key={c} onClick={() => setCategory(c)} style={{
                padding: "5px 12px", border: "none", cursor: "pointer", borderRadius: "6px",
                fontSize: "11px", fontWeight: 600, fontFamily: T.font,
                background: category === c ? T.accentGlow : "transparent",
                color: category === c ? T.accent : T.textMuted,
                border: `1px solid ${category === c ? T.accent + "40" : T.border}`,
              }}>{c}</button>
            ))}
          </div>
        </div>

        {/* Template Grid */}
        <div style={{ flex: 1, overflow: "auto", padding: "16px 24px 24px" }}>
          {loading ? (
            <div style={{ textAlign: "center", padding: "40px", color: T.textMuted }}>Loading templates...</div>
          ) : templates.length === 0 ? (
            <div style={{ textAlign: "center", padding: "40px", color: T.textMuted }}>No templates found.</div>
          ) : (
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(230px, 1fr))", gap: "14px" }}>
              {templates.map(tmpl => (
                <div key={tmpl.id} onClick={() => setPreview(tmpl)}
                  style={{
                    background: T.surface, border: `1px solid ${preview?.id === tmpl.id ? T.accent : T.border}`,
                    borderRadius: T.radius, padding: "16px", cursor: "pointer",
                    transition: "all 0.15s", boxShadow: preview?.id === tmpl.id ? `0 0 0 2px ${T.accentGlow}` : "none",
                  }}>
                  {/* Icon + Category badge */}
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "10px" }}>
                    <span style={{ fontSize: "28px" }}>{tmpl.icon}</span>
                    <span style={{ fontSize: "9px", fontWeight: 700, color: T.accent, background: T.accentGlow, padding: "3px 8px", borderRadius: "6px", textTransform: "uppercase", letterSpacing: "0.5px" }}>{tmpl.category}</span>
                  </div>
                  <div style={{ fontSize: "14px", fontWeight: 700, color: T.text, marginBottom: "4px" }}>{tmpl.name}</div>
                  <div style={{ fontSize: "11px", color: T.textMuted, lineHeight: 1.5, marginBottom: "10px" }}>{tmpl.description}</div>
                  {/* Slide count + types preview */}
                  <div style={{ display: "flex", gap: "4px", flexWrap: "wrap" }}>
                    <span style={{ fontSize: "10px", fontWeight: 600, color: T.textDim, background: T.mode === "light" ? "#f1f5f9" : "rgba(255,255,255,0.05)", padding: "2px 8px", borderRadius: "4px" }}>
                      {tmpl.slides.length} slides
                    </span>
                    {[...new Set(tmpl.slides.map(s => s.type))].slice(0, 4).map(t => (
                      <span key={t} style={{ fontSize: "9px", color: T.textDim, background: T.mode === "light" ? "#f1f5f9" : "rgba(255,255,255,0.05)", padding: "2px 6px", borderRadius: "4px" }}>{t}</span>
                    ))}
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>

        {/* Preview Panel + Design Variant Selector */}
        {preview && (
          <div style={{ padding: "16px 24px", borderTop: `1px solid ${T.border}` }}>
            <div style={{ display: "flex", gap: "16px", alignItems: "center", marginBottom: "10px" }}>
              <div style={{ flex: 1 }}>
                <div style={{ fontSize: "13px", fontWeight: 700, color: T.text }}>{preview.icon} {preview.name}</div>
                <div style={{ fontSize: "11px", color: T.textMuted, marginTop: "2px" }}>
                  {preview.slides.map((s, i) => `${i + 1}. ${s.title}`).join("  →  ")}
                </div>
              </div>
              <button onClick={() => handleSelect(preview)} style={{ ...css.btnPrimary, display: "flex", alignItems: "center", gap: "6px" }}>
                ✨ Use Template
              </button>
            </div>
            {/* Design Variants */}
            <div style={{ display: "flex", gap: "6px", flexWrap: "wrap" }}>
              <span style={{ fontSize: "9px", fontWeight: 700, color: T.textMuted, textTransform: "uppercase", alignSelf: "center", marginRight: "4px" }}>Style:</span>
              {[
                { id: "corporate", label: "Corporate", color: "#0078d4" },
                { id: "cards", label: "Cards", color: "#3b82f6" },
                { id: "minimal", label: "Minimal", color: "#6b7280" },
                { id: "dark", label: "Dark", color: "#1e293b" },
                { id: "infographic", label: "Infographic", color: "#0d9488" },
              ].map(v => (
                <button key={v.id} onClick={() => handleSelect({ ...preview, designVariant: v.id })}
                  style={{
                    padding: "4px 10px", border: `1px solid ${T.border}`, borderRadius: "6px", cursor: "pointer",
                    fontSize: "10px", fontWeight: 600, fontFamily: T.font, background: "transparent", color: T.text,
                    display: "flex", alignItems: "center", gap: "4px",
                  }}>
                  <span style={{ width: "8px", height: "8px", borderRadius: "50%", background: v.color, display: "inline-block" }} />
                  {v.label}
                </button>
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

// 📂 DATA MANAGER — Photos, Templates, Sync, Excel AI
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function DataManagerModal({ onClose }) {
  const { T, css } = useTheme();
  const { loadPresentation, setCanvasMode } = useStoreActions();
  const [tab, setTab] = useState("photos");
  const [photos, setPhotos] = useState([]);
  const [templates, setTemplates] = useState([]);
  const [status, setStatus] = useState(null);
  const [loading, setLoading] = useState(false);

  // Fetch data on mount
  useEffect(() => {
    fetch("/api/photos/list").then(r => r.json()).then(d => setPhotos(d.photos || [])).catch(() => {});
    fetch("/api/templates/list").then(r => r.json()).then(d => setTemplates(d.templates || [])).catch(() => {});
  }, []);

  const showStatus = (msg, isError) => { setStatus({ msg, isError }); setTimeout(() => setStatus(null), 4000); };

  // ─── Photos Tab ───
  const handlePhotoUpload = async (e) => {
    const files = e.target.files;
    if (!files?.length) return;
    const fd = new FormData();
    Array.from(files).forEach(f => fd.append("photos", f));
    setLoading(true);
    try {
      const res = await fetch("/api/photos/upload", { method: "POST", body: fd });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error);
      setPhotos(prev => {
        const ids = new Set(data.photos.map(p => p.id));
        return [...prev.filter(p => !ids.has(p.id)), ...data.photos];
      });
      showStatus(`${data.photos.length} photo(s) uploaded`);
    } catch (e) { showStatus(e.message, true); }
    setLoading(false);
  };

  const deletePhoto = async (id) => {
    await fetch(`/api/photos/${id}`, { method: "DELETE" });
    setPhotos(prev => prev.filter(p => p.id !== id));
  };

  const updatePhoto = async (id, updates) => {
    const res = await fetch(`/api/photos/${id}`, { method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify(updates) });
    const data = await res.json();
    setPhotos(prev => prev.map(p => p.id === id ? data : p));
  };

  // ─── Templates Tab ───
  const handleTemplateUpload = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const fd = new FormData();
    fd.append("template", file);
    setLoading(true);
    try {
      const res = await fetch("/api/templates/upload", { method: "POST", body: fd });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error);
      setTemplates(prev => [...prev.map(t => ({ ...t, active: false })), data]);
      showStatus(`Template ${data.version} uploaded`);
    } catch (e) { showStatus(e.message, true); }
    setLoading(false);
  };

  // ─── Sync Tab ───
  const handleSync = async (direction) => {
    const input = document.createElement("input");
    input.type = "file"; input.accept = ".xlsx";
    input.onchange = async (e) => {
      const file = e.target.files?.[0];
      if (!file) return;
      const fd = new FormData();
      fd.append("excel", file);
      setLoading(true);
      try {
        const endpoint = direction === "template-to-excel" ? "/api/sync/template-to-excel" : "/api/sync/excel-to-template";
        const res = await fetch(endpoint, { method: "POST", body: fd });
        const data = await res.json();
        if (!res.ok) throw new Error(data.error);
        showStatus(`${data.changes?.length || 0} changes applied. ${data.message || ""}`);
        if (data.downloadUrl) {
          const a = document.createElement("a"); a.href = data.downloadUrl; a.download = "synced.xlsx"; a.click();
        }
      } catch (e) { showStatus(e.message, true); }
      setLoading(false);
    };
    input.click();
  };

  // ─── Excel AI Tab (Feature 5) ───
  const handleExcelAI = async (secure) => {
    const input = document.createElement("input");
    input.type = "file"; input.accept = ".xlsx";
    input.onchange = async (e) => {
      const file = e.target.files?.[0];
      if (!file) return;
      const fd = new FormData();
      fd.append("excel", file);
      fd.append("secure", String(secure));
      setLoading(true);
      try {
        const res = await fetch("/api/ai/generate-from-excel", { method: "POST", body: fd });
        const data = await res.json();
        if (!res.ok) throw new Error(data.error);
        if (data.slides?.length) {
          loadPresentation(data.slides);
          setCanvasMode(true);
          showStatus(`Generated ${data.slides.length} slides${data.masked ? " (masked data)" : ""}`);
          setTimeout(onClose, 500);
        }
      } catch (e) { showStatus(e.message, true); }
      setLoading(false);
    };
    input.click();
  };

  // ─── Auto Template (Feature 6) ───
  const handleAutoTemplate = async () => {
    const input = document.createElement("input");
    input.type = "file"; input.accept = ".xlsx";
    input.onchange = async (e) => {
      const file = e.target.files?.[0];
      if (!file) return;
      const fd = new FormData();
      fd.append("excel", file);
      setLoading(true);
      try {
        const res = await fetch("/api/templates/auto-create", { method: "POST", body: fd });
        const data = await res.json();
        if (!res.ok) throw new Error(data.error);
        if (data.mismatch) {
          setTemplates(prev => [...prev.map(t => ({ ...t, active: false })), data.template]);
          showStatus(`New template ${data.template.version} created from Excel`);
        } else {
          showStatus(data.message);
        }
      } catch (e) { showStatus(e.message, true); }
      setLoading(false);
    };
    input.click();
  };

  // ─── PPT Reverse Engineering ───
  const handleReverseEngineer = async () => {
    const input = document.createElement("input");
    input.type = "file"; input.accept = ".pptx";
    input.onchange = async (e) => {
      const file = e.target.files?.[0];
      if (!file) return;
      const fd = new FormData();
      fd.append("pptx", file);
      setLoading(true);
      try {
        const res = await fetch("/api/pptx/reverse-engineer", { method: "POST", body: fd });
        const data = await res.json();
        if (!res.ok) throw new Error(data.error);
        setTemplates(prev => [...prev, data.template]);
        showStatus(`Parsed ${data.slideDefinitions.length} slides. Template "${data.template.name}" created.`);
        if (data.excelDownload) {
          const a = document.createElement("a"); a.href = data.excelDownload; a.download = "data_template.xlsx"; a.click();
        }
      } catch (e) { showStatus(e.message, true); }
      setLoading(false);
    };
    input.click();
  };

  // ─── AI Template Enhancement ───
  const [enhancing, setEnhancing] = useState(null); // template id being enhanced
  const [suggestions, setSuggestions] = useState(null);

  const handleEnhance = async (tmplId) => {
    setEnhancing(tmplId);
    setSuggestions(null);
    setLoading(true);
    try {
      const res = await fetch("/api/templates/enhance", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ templateId: tmplId }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error);
      setSuggestions(data.suggestions);
      showStatus("AI analysis complete. Review suggestions below.");
    } catch (e) { showStatus(e.message, true); }
    setLoading(false);
  };

  const handleApplyEnhancement = async (tmplId) => {
    if (!suggestions?.enhanced_slides?.length) return;
    setLoading(true);
    try {
      const res = await fetch("/api/templates/apply-enhancement", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ originalId: tmplId, enhancedSlides: suggestions.enhanced_slides }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error);
      setTemplates(prev => [...prev, data.template]);
      setSuggestions(null);
      setEnhancing(null);
      showStatus(`Enhanced template "${data.template.name}" saved to library.`);
    } catch (e) { showStatus(e.message, true); }
    setLoading(false);
  };

  const TABS = [
    { id: "photos", label: "📸 Photos", icon: "📸" },
    { id: "templates", label: "📋 Templates", icon: "📋" },
    { id: "sync", label: "🔄 Sync", icon: "🔄" },
    { id: "excel-ai", label: "🤖 Excel AI", icon: "🤖" },
    { id: "reverse", label: "📥 Import PPT", icon: "📥" },
    { id: "enhance", label: "✨ Enhance", icon: "✨" },
  ];

  return (
    <div style={{ position: "fixed", inset: 0, zIndex: 9999, background: "rgba(0,0,0,0.6)", backdropFilter: "blur(8px)", display: "flex", alignItems: "center", justifyContent: "center" }}
      onClick={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <div style={{ background: T.mode === "light" ? "#fff" : "#111827", borderRadius: T.radiusLg, width: "680px", maxWidth: "95vw", maxHeight: "85vh", display: "flex", flexDirection: "column", border: `1px solid ${T.borderHover}`, boxShadow: "0 24px 80px rgba(0,0,0,0.5)" }}>
        {/* Header */}
        <div style={{ padding: "20px 24px 0", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
            <span style={{ fontSize: "22px" }}>📂</span>
            <div><div style={{ fontSize: "16px", fontWeight: 800, color: T.text }}>Data Manager</div>
              <div style={{ fontSize: "11px", color: T.textMuted }}>Photos, Templates, Sync & Excel AI</div></div>
          </div>
          <button onClick={onClose} style={{ ...css.btnIcon, fontSize: "16px" }}>✕</button>
        </div>

        {/* Tabs */}
        <div style={{ display: "flex", gap: "4px", padding: "16px 24px 0", borderBottom: `1px solid ${T.border}` }}>
          {TABS.map(t => (
            <button key={t.id} onClick={() => setTab(t.id)} style={{
              padding: "8px 16px", border: "none", cursor: "pointer", borderRadius: "8px 8px 0 0",
              fontSize: "12px", fontWeight: 700, fontFamily: T.font,
              background: tab === t.id ? T.accentGlow : "transparent",
              color: tab === t.id ? T.accent : T.textMuted,
              borderBottom: tab === t.id ? `2px solid ${T.accent}` : "2px solid transparent",
            }}>{t.label}</button>
          ))}
        </div>

        {/* Status */}
        {status && (
          <div style={{ margin: "12px 24px 0", padding: "8px 14px", borderRadius: T.radiusSm, fontSize: "12px", fontWeight: 600,
            background: status.isError ? T.dangerBg : T.accentGlow, color: status.isError ? T.danger : T.accent }}>
            {status.msg}
          </div>
        )}

        {/* Tab Content */}
        <div style={{ flex: 1, overflow: "auto", padding: "16px 24px 24px" }}>
          {loading && <div style={{ textAlign: "center", padding: "20px", color: T.textMuted, fontSize: "13px" }}>⏳ Processing...</div>}

          {/* ─── Photos Tab ─── */}
          {tab === "photos" && !loading && (
            <div>
              <div style={{ display: "flex", gap: "8px", marginBottom: "16px" }}>
                <label style={{ ...css.btnPrimary, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: "6px" }}>
                  📤 Upload Photos <input type="file" accept=".jpg,.jpeg,.png" multiple onChange={handlePhotoUpload} style={{ display: "none" }} />
                </label>
                <span style={{ fontSize: "11px", color: T.textMuted, alignSelf: "center" }}>JPG/PNG, max 5MB. Name files like "VP_John.png" for auto-mapping.</span>
              </div>
              {photos.length === 0 && <p style={{ color: T.textDim, fontSize: "13px" }}>No photos uploaded yet.</p>}
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(150px, 1fr))", gap: "12px" }}>
                {photos.map(p => (
                  <div key={p.id} style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: T.radius, padding: "12px", textAlign: "center" }}>
                    <img src={p.photoUrl} alt={p.name} style={{ width: "64px", height: "64px", borderRadius: "50%", objectFit: "cover", border: `2px solid ${T.accent}` }} />
                    <div style={{ marginTop: "8px" }}>
                      <input value={p.name} onChange={(e) => updatePhoto(p.id, { name: e.target.value })} placeholder="Name"
                        style={{ ...css.input, fontSize: "11px", padding: "4px 8px", textAlign: "center", marginBottom: "4px" }} />
                      <select value={p.designation} onChange={(e) => updatePhoto(p.id, { designation: e.target.value })}
                        style={{ ...css.input, fontSize: "10px", padding: "3px 6px" }}>
                        {["", "VP", "Director", "Manager", "Team Member"].map(d => <option key={d} value={d}>{d || "Select Role"}</option>)}
                      </select>
                    </div>
                    <button onClick={() => deletePhoto(p.id)} style={{ ...css.btnDanger, marginTop: "6px", fontSize: "10px" }}>Remove</button>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* ─── Templates Tab ─── */}
          {tab === "templates" && !loading && (
            <div>
              <div style={{ display: "flex", gap: "8px", marginBottom: "16px", flexWrap: "wrap" }}>
                <label style={{ ...css.btnPrimary, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: "6px" }}>
                  📤 Upload Template <input type="file" accept=".xlsx" onChange={handleTemplateUpload} style={{ display: "none" }} />
                </label>
                <button onClick={handleAutoTemplate} style={{ ...css.btnIcon, width: "auto", padding: "0 14px", fontSize: "11px", fontWeight: 600, fontFamily: T.font }}>
                  🔧 Auto-Create from Excel
                </button>
              </div>
              {templates.length === 0 && <p style={{ color: T.textDim, fontSize: "13px" }}>No templates uploaded yet.</p>}
              {templates.map(t => (
                <div key={t.id} style={{ background: T.surface, border: `1px solid ${t.active ? T.accent : T.border}`, borderRadius: T.radius, padding: "12px", marginBottom: "8px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <div>
                    <div style={{ fontSize: "13px", fontWeight: 700, color: T.text }}>{t.fileName} <span style={{ fontSize: "10px", color: T.accent, fontWeight: 600 }}>{t.version}</span></div>
                    <div style={{ fontSize: "10px", color: T.textMuted }}>
                      {t.sheets?.join(", ")} • {new Date(t.uploadedOn).toLocaleDateString()}
                      {t.active && <span style={{ color: T.success, fontWeight: 700 }}> • Active</span>}
                    </div>
                  </div>
                  <a href={`/api/templates/download/${t.id}`} style={{ ...css.btnIcon, width: "auto", padding: "0 10px", fontSize: "10px", fontWeight: 600, fontFamily: T.font, textDecoration: "none", color: T.textMuted }}>⬇ Download</a>
                </div>
              ))}
            </div>
          )}

          {/* ─── Sync Tab ─── */}
          {tab === "sync" && !loading && (
            <div style={{ display: "flex", flexDirection: "column", gap: "16px" }}>
              <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: T.radius, padding: "20px" }}>
                <div style={{ fontSize: "14px", fontWeight: 700, color: T.text, marginBottom: "6px" }}>📋 → 📊 Template to Excel</div>
                <p style={{ fontSize: "12px", color: T.textMuted, marginBottom: "12px" }}>Add new template columns to your Excel file. Missing fields will be appended.</p>
                <button onClick={() => handleSync("template-to-excel")} style={css.btnPrimary}>Select Excel & Sync from Template</button>
              </div>
              <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: T.radius, padding: "20px" }}>
                <div style={{ fontSize: "14px", fontWeight: 700, color: T.text, marginBottom: "6px" }}>📊 → 📋 Excel to Template</div>
                <p style={{ fontSize: "12px", color: T.textMuted, marginBottom: "12px" }}>Update template with new fields found in your Excel. Existing fields preserved.</p>
                <button onClick={() => handleSync("excel-to-template")} style={css.btnPrimary}>Select Excel & Sync to Template</button>
              </div>
            </div>
          )}

          {/* ─── Excel AI Tab ─── */}
          {tab === "excel-ai" && !loading && (
            <div style={{ display: "flex", flexDirection: "column", gap: "16px" }}>
              <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: T.radius, padding: "20px" }}>
                <div style={{ fontSize: "14px", fontWeight: 700, color: T.text, marginBottom: "6px" }}>🔒 Secure Mode (Recommended)</div>
                <p style={{ fontSize: "12px", color: T.textMuted, marginBottom: "12px" }}>Client names, financial data, and personal names are masked before sending to AI. No raw data leaves your server.</p>
                <button onClick={() => handleExcelAI(true)} style={css.btnPrimary}>📤 Select Excel & Generate (Masked)</button>
              </div>
              <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: T.radius, padding: "20px" }}>
                <div style={{ fontSize: "14px", fontWeight: 700, color: T.text, marginBottom: "6px" }}>⚠️ Direct Mode</div>
                <p style={{ fontSize: "12px", color: T.textMuted, marginBottom: "12px" }}>Send actual data to AI for more accurate slides. Use only with non-sensitive data.</p>
                <button onClick={() => handleExcelAI(false)} style={{ ...css.btnIcon, width: "auto", padding: "0 14px", fontSize: "11px", fontWeight: 600, fontFamily: T.font }}>Select Excel & Generate (Direct)</button>
              </div>
            </div>
          )}

          {/* ─── Reverse Engineer PPT Tab ─── */}
          {tab === "reverse" && !loading && (
            <div style={{ display: "flex", flexDirection: "column", gap: "16px" }}>
              <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: T.radius, padding: "20px" }}>
                <div style={{ fontSize: "14px", fontWeight: 700, color: T.text, marginBottom: "6px" }}>📥 Import PowerPoint File</div>
                <p style={{ fontSize: "12px", color: T.textMuted, marginBottom: "12px" }}>
                  Upload a .pptx file to reverse-engineer its structure. The system will:
                </p>
                <ul style={{ fontSize: "12px", color: T.textMuted, marginBottom: "16px", paddingLeft: "20px", lineHeight: 1.8 }}>
                  <li>Parse all slides and classify their types</li>
                  <li>Create a reusable template in your library</li>
                  <li>Generate an Excel data template for re-population</li>
                </ul>
                <button onClick={handleReverseEngineer} style={css.btnPrimary}>📂 Select .pptx File & Analyze</button>
              </div>
            </div>
          )}

          {/* ─── AI Enhance Tab ─── */}
          {tab === "enhance" && !loading && (
            <div style={{ display: "flex", flexDirection: "column", gap: "12px" }}>
              <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: T.radius, padding: "16px" }}>
                <div style={{ fontSize: "14px", fontWeight: 700, color: T.text, marginBottom: "6px" }}>✨ AI Template Enhancement</div>
                <p style={{ fontSize: "12px", color: T.textMuted, marginBottom: "12px" }}>Select a template to analyze and receive AI-powered improvement suggestions.</p>
              </div>
              {templates.map(t => (
                <div key={t.id} style={{ background: T.surface, border: `1px solid ${enhancing === t.id ? T.accent : T.border}`, borderRadius: T.radius, padding: "14px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <div>
                    <div style={{ fontSize: "13px", fontWeight: 700, color: T.text }}>{t.icon || "📋"} {t.name || t.fileName}</div>
                    <div style={{ fontSize: "10px", color: T.textMuted }}>{t.slides?.length || "?"} slides • {t.category || t.version || ""}</div>
                  </div>
                  <button onClick={() => handleEnhance(t.id)} disabled={enhancing === t.id}
                    style={{ ...css.btnIcon, width: "auto", padding: "0 12px", fontSize: "10px", fontWeight: 600, fontFamily: T.font, color: T.accent }}>
                    {enhancing === t.id ? "⏳ Analyzing..." : "✨ Enhance"}
                  </button>
                </div>
              ))}
              {templates.length === 0 && <p style={{ color: T.textDim, fontSize: "12px" }}>No templates available. Upload a template or import a PPTX first.</p>}

              {/* Suggestions Panel */}
              {suggestions && (
                <div style={{ background: T.surface, border: `1px solid ${T.accent}`, borderRadius: T.radius, padding: "16px" }}>
                  <div style={{ fontSize: "13px", fontWeight: 700, color: T.accent, marginBottom: "8px" }}>AI Suggestions</div>
                  {suggestions.analysis && <p style={{ fontSize: "12px", color: T.text, marginBottom: "10px", lineHeight: 1.5 }}>{suggestions.analysis}</p>}

                  {suggestions.add_slides?.length > 0 && (
                    <div style={{ marginBottom: "8px" }}>
                      <div style={{ fontSize: "10px", fontWeight: 700, color: T.success, textTransform: "uppercase", marginBottom: "4px" }}>+ Add</div>
                      {suggestions.add_slides.map((s, i) => <div key={i} style={{ fontSize: "11px", color: T.textMuted, padding: "2px 0" }}>Slide {s.order}: {s.title} ({s.type}) — {s.reason}</div>)}
                    </div>
                  )}
                  {suggestions.remove_slides?.length > 0 && (
                    <div style={{ marginBottom: "8px" }}>
                      <div style={{ fontSize: "10px", fontWeight: 700, color: T.danger, textTransform: "uppercase", marginBottom: "4px" }}>- Remove</div>
                      {suggestions.remove_slides.map((s, i) => <div key={i} style={{ fontSize: "11px", color: T.textMuted, padding: "2px 0" }}>Slide {s.order}: {s.title} — {s.reason}</div>)}
                    </div>
                  )}
                  {suggestions.modify_slides?.length > 0 && (
                    <div style={{ marginBottom: "8px" }}>
                      <div style={{ fontSize: "10px", fontWeight: 700, color: T.accent, textTransform: "uppercase", marginBottom: "4px" }}>~ Modify</div>
                      {suggestions.modify_slides.map((s, i) => <div key={i} style={{ fontSize: "11px", color: T.textMuted, padding: "2px 0" }}>Slide {s.order}: {s.change}</div>)}
                    </div>
                  )}
                  {suggestions.visual_improvements?.length > 0 && (
                    <div style={{ marginBottom: "8px" }}>
                      <div style={{ fontSize: "10px", fontWeight: 700, color: T.purple, textTransform: "uppercase", marginBottom: "4px" }}>Visual</div>
                      {suggestions.visual_improvements.map((s, i) => <div key={i} style={{ fontSize: "11px", color: T.textMuted, padding: "2px 0" }}>• {s}</div>)}
                    </div>
                  )}

                  <div style={{ display: "flex", gap: "8px", marginTop: "12px" }}>
                    <button onClick={() => handleApplyEnhancement(enhancing)} style={css.btnPrimary}>✅ Apply & Save Enhanced Template</button>
                    <button onClick={() => { setSuggestions(null); setEnhancing(null); }} style={{ ...css.btnIcon, width: "auto", padding: "0 12px", fontSize: "11px", fontWeight: 600, fontFamily: T.font }}>Dismiss</button>
                  </div>
                </div>
              )}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function AIGeneratorModal({ onClose, initialTemplate }) {
  const { T, css } = useTheme();
  const { loadPresentation, setCanvasMode } = useStoreActions();
  const [prompt, setPrompt] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [selectedTemplate, setSelectedTemplate] = useState(initialTemplate || null);
  const [showTemplateGallery, setShowTemplateGallery] = useState(false);

  const handleGenerate = async () => {
    if (!prompt.trim()) { setError("Please enter a prompt"); return; }
    setLoading(true);
    setError(null);
    try {
      let res;
      if (selectedTemplate) {
        // Use template-guided generation
        res = await fetch("/api/ai/generate-with-template", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ templateId: selectedTemplate.id, prompt: prompt.trim() }),
        });
      } else {
        // Free-form generation
        res = await fetch("/api/generate", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ prompt: prompt.trim() }),
        });
      }
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Generation failed");
      if (!data.slides || !Array.isArray(data.slides) || data.slides.length === 0) {
        throw new Error("No slides returned from AI");
      }
      loadPresentation(data.slides);
      setCanvasMode(true);
      onClose();
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{
      position: "fixed", inset: 0, zIndex: 9999,
      background: "rgba(0,0,0,0.6)", backdropFilter: "blur(8px)",
      display: "flex", alignItems: "center", justifyContent: "center",
    }} onClick={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <div style={{
        background: T.mode === "light" ? "#ffffff" : "#111827",
        borderRadius: T.radiusLg, padding: "28px", width: "560px", maxWidth: "90vw",
        boxShadow: "0 24px 80px rgba(0,0,0,0.5)",
        border: `1px solid ${T.borderHover}`,
      }}>
        {/* Header */}
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "20px" }}>
          <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
            <div style={{
              width: "36px", height: "36px", borderRadius: T.radiusSm,
              background: `linear-gradient(135deg, ${T.accent}, ${T.purple})`,
              display: "flex", alignItems: "center", justifyContent: "center",
              fontSize: "18px",
            }}>✨</div>
            <div>
              <div style={{ fontSize: "16px", fontWeight: 800, color: T.text }}>AI Slide Generator</div>
              <div style={{ fontSize: "11px", color: T.textMuted }}>Powered by Gemini AI</div>
            </div>
          </div>
          <button onClick={onClose} style={{ ...css.btnIcon, fontSize: "16px" }}>✕</button>
        </div>

        {/* Template Selection */}
        <div style={{ marginBottom: "14px" }}>
          <div style={{ ...css.label, marginBottom: "8px" }}>Slide Structure Template</div>
          {selectedTemplate ? (
            <div style={{ display: "flex", alignItems: "center", gap: "10px", padding: "8px 12px", borderRadius: T.radiusSm, background: T.accentGlow, border: `1px solid ${T.accent}40` }}>
              <span style={{ fontSize: "18px" }}>{selectedTemplate.icon}</span>
              <div style={{ flex: 1 }}>
                <div style={{ fontSize: "12px", fontWeight: 700, color: T.accent }}>{selectedTemplate.name}</div>
                <div style={{ fontSize: "10px", color: T.textMuted }}>{selectedTemplate.slides.length} slides • {selectedTemplate.category}</div>
              </div>
              <button onClick={() => setSelectedTemplate(null)} style={{ ...css.btnIcon, width: "24px", height: "24px", fontSize: "10px" }}>✕</button>
              <button onClick={() => setShowTemplateGallery(true)} style={{ ...css.btnIcon, width: "auto", padding: "0 10px", height: "24px", fontSize: "10px", fontWeight: 600, fontFamily: T.font }}>Change</button>
            </div>
          ) : (
            <button onClick={() => setShowTemplateGallery(true)} style={{
              ...css.btnGhost, display: "flex", alignItems: "center", justifyContent: "center", gap: "8px", padding: "10px",
            }}>📚 Browse Template Library</button>
          )}
        </div>

        {/* Prompt Templates */}
        <div style={{ marginBottom: "14px" }}>
          <div style={{ ...css.label, marginBottom: "8px" }}>Quick Prompts</div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: "6px" }}>
            {PROMPT_TEMPLATES.map((t) => (
              <button key={t.label} onClick={() => setPrompt(t.prompt)}
                style={{
                  background: prompt === t.prompt ? T.accentGlow : (T.mode === "light" ? "#f1f5f9" : "rgba(255,255,255,0.04)"),
                  border: `1px solid ${prompt === t.prompt ? T.accent : T.border}`,
                  color: prompt === t.prompt ? T.accent : T.text,
                  padding: "6px 12px", borderRadius: "8px", cursor: "pointer",
                  fontSize: "12px", fontWeight: 600, fontFamily: T.font,
                  display: "flex", alignItems: "center", gap: "6px",
                  transition: "all 0.15s",
                }}>
                <span>{t.icon}</span> {t.label}
              </button>
            ))}
          </div>
        </div>

        {/* Prompt Input */}
        <div style={{ marginBottom: "16px" }}>
          <label style={css.label}>Describe Your Presentation</label>
          <textarea
            value={prompt}
            onChange={(e) => setPrompt(e.target.value)}
            placeholder="E.g., Create a quarterly report for Microsoft Practice Q1 2026 with 7 resources, 4 projects, 3 new customers..."
            style={{
              ...css.input,
              resize: "vertical", minHeight: "100px", lineHeight: 1.5,
            }}
          />
        </div>

        {/* Error */}
        {error && (
          <div style={{
            background: T.dangerBg, color: T.danger, padding: "10px 14px",
            borderRadius: T.radiusSm, fontSize: "12px", marginBottom: "14px",
            fontWeight: 600,
          }}>
            {error}
          </div>
        )}

        {/* Actions */}
        <div style={{ display: "flex", justifyContent: "flex-end", gap: "8px" }}>
          <button onClick={onClose}
            style={{ ...css.btnIcon, width: "auto", padding: "0 16px", fontSize: "12px", fontWeight: 600, fontFamily: T.font }}>
            Cancel
          </button>
          <button onClick={handleGenerate} disabled={loading || !prompt.trim()}
            style={{
              ...css.btnPrimary,
              opacity: loading || !prompt.trim() ? 0.5 : 1,
              cursor: loading ? "wait" : "pointer",
              display: "flex", alignItems: "center", gap: "8px",
            }}>
            {loading ? (
              <>
                <span style={{ display: "inline-block", animation: "spin 1s linear infinite", fontSize: "14px" }}>⟳</span>
                Generating...
              </>
            ) : (
              <>✨ Generate Slides</>
            )}
          </button>
        </div>

        {/* Loading hint */}
        {loading && (
          <div style={{ textAlign: "center", marginTop: "14px", fontSize: "11px", color: T.textMuted }}>
            Claude is crafting your presentation — this may take 10-20 seconds...
          </div>
        )}
      </div>

      <style>{`@keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }`}</style>

      {/* Template Gallery Sub-Modal */}
      {showTemplateGallery && (
        <TemplateGalleryModal
          onClose={() => setShowTemplateGallery(false)}
          onSelectTemplate={(tmpl) => { setSelectedTemplate(tmpl); setShowTemplateGallery(false); }}
        />
      )}
    </div>
  );
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 🏠  MAIN APP
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 🤖 AI QUICK ACTIONS — Improve/Concise/Visual/Insights per slide
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

const AI_ACTIONS = [
  { id: "improve", label: "✨ Improve", desc: "Enhance clarity & professionalism" },
  { id: "concise", label: "📝 Concise", desc: "Reduce text to bullet points" },
  { id: "visual", label: "📊 Visual", desc: "Add icons, metrics, visual structure" },
  { id: "insights", label: "💡 Insights", desc: "Highlight key business insights" },
];

function AIQuickActions({ slide, updateSlide, T }) {
  const [loading, setLoading] = useState(null); // which action is loading
  const [error, setError] = useState(null);
  const [undoData, setUndoData] = useState(null); // previous slide data for undo

  const handleAction = useCallback(async (actionId) => {
    if (loading) return;
    setLoading(actionId);
    setError(null);

    // Save current state for undo
    setUndoData({ ...slide });

    try {
      const res = await fetch("/api/ai/edit-slide", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: actionId, slideData: slide }),
      });
      const data = await res.json();

      if (!res.ok) throw new Error(data.error || "AI edit failed");
      if (!data.slide) throw new Error("No slide data returned");

      // Merge AI result into current slide — preserve id, layout, elementStyles
      const merged = { ...data.slide, id: slide.id, type: slide.type };
      if (slide.layout) merged.layout = slide.layout;
      if (slide.elementStyles) merged.elementStyles = slide.elementStyles;
      if (slide.elementGroups) merged.elementGroups = slide.elementGroups;

      updateSlide(slide.id, merged);
    } catch (e) {
      setError(e.message);
      setTimeout(() => setError(null), 5000);
    } finally {
      setLoading(null);
    }
  }, [slide, loading, updateSlide]);

  const handleUndo = useCallback(() => {
    if (!undoData) return;
    updateSlide(undoData.id, undoData);
    setUndoData(null);
  }, [undoData, updateSlide]);

  return (
    <div style={{ marginBottom: "12px" }}>
      <div style={{ display: "flex", gap: "4px", flexWrap: "wrap", alignItems: "center" }}>
        {AI_ACTIONS.map(qa => (
          <button key={qa.id}
            onClick={() => handleAction(qa.id)}
            disabled={!!loading}
            title={qa.desc}
            style={{
              padding: "4px 8px", border: `1px solid ${T.border}`, borderRadius: "6px",
              background: loading === qa.id ? T.accentGlow : "transparent",
              color: loading === qa.id ? T.accent : T.textMuted,
              cursor: loading ? "wait" : "pointer",
              fontSize: "10px", fontWeight: 600, fontFamily: T.font,
              transition: "all 0.15s",
              opacity: loading && loading !== qa.id ? 0.4 : 1,
            }}
            onMouseEnter={(e) => { if (!loading) { e.currentTarget.style.background = T.accentGlow; e.currentTarget.style.color = T.accent; } }}
            onMouseLeave={(e) => { if (loading !== qa.id) { e.currentTarget.style.background = "transparent"; e.currentTarget.style.color = T.textMuted; } }}
          >{loading === qa.id ? "⏳" : ""}{qa.label}</button>
        ))}
        {undoData && (
          <button onClick={handleUndo} style={{ padding: "4px 8px", border: `1px solid ${T.border}`, borderRadius: "6px", background: "transparent", color: T.danger, cursor: "pointer", fontSize: "10px", fontWeight: 600, fontFamily: T.font }}>
            ↩ Undo
          </button>
        )}
      </div>
      {loading && <div style={{ fontSize: "9px", color: T.accent, marginTop: "4px", fontFamily: T.fontMono }}>AI is editing this slide...</div>}
      {error && <div style={{ fontSize: "9px", color: T.danger, marginTop: "4px", padding: "4px 8px", background: T.dangerBg, borderRadius: "4px" }}>{error}</div>}
    </div>
  );
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 🏠 LANDING PAGE
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function LandingPage({ onEnter }) {
  const { T } = useTheme();
  const [activeVariant, setActiveVariant] = useState(0);
  const isLight = T.mode === "light";
  const bg = isLight ? "#ffffff" : "#06090f";
  const cardBg = isLight ? "#f8fafc" : "#111827";
  const border = isLight ? "#e5e7eb" : "rgba(148,163,194,0.12)";
  const textPrimary = isLight ? "#111827" : "#f1f5f9";
  const textSec = isLight ? "#6b7280" : "#94a3b8";

  const variants = ["Corporate", "Cards", "Minimal", "Dark", "Infographic"];

  return (
    <div style={{ background: bg, minHeight: "100vh", fontFamily: T.font, color: textPrimary, overflowX: "hidden" }}>
      <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800;900&display=swap" rel="stylesheet" />
      <style>{`@keyframes fadeUp { from { opacity: 0; transform: translateY(30px); } to { opacity: 1; transform: translateY(0); } }
        @keyframes float { 0%,100% { transform: translateY(0); } 50% { transform: translateY(-10px); } }
        .fade-up { animation: fadeUp 0.7s ease forwards; }
        .fade-up-2 { animation: fadeUp 0.7s 0.15s ease forwards; opacity: 0; }
        .fade-up-3 { animation: fadeUp 0.7s 0.3s ease forwards; opacity: 0; }
        .float { animation: float 4s ease-in-out infinite; }
        .land-card:hover { transform: translateY(-6px); box-shadow: 0 20px 60px rgba(0,0,0,0.15); }
      `}</style>

      {/* ─── NAV ─── */}
      <nav style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "20px 48px", position: "sticky", top: 0, zIndex: 100, background: bg + "ee", backdropFilter: "blur(20px)" }}>
        <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
          <div style={{ width: "36px", height: "36px", borderRadius: "10px", background: `linear-gradient(135deg, ${T.accent}, ${T.purple})`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: "16px", color: "#fff", fontWeight: 900 }}>P</div>
          <span style={{ fontSize: "18px", fontWeight: 800, letterSpacing: "-0.02em" }}>PoterPointAI</span>
        </div>
        <button onClick={() => onEnter("editor")} style={{ padding: "10px 24px", background: `linear-gradient(135deg, ${T.accent}, ${T.accentDark})`, color: "#fff", border: "none", borderRadius: "10px", fontSize: "14px", fontWeight: 700, cursor: "pointer", fontFamily: T.font }}>
          Open Editor
        </button>
      </nav>

      {/* ─── SECTION 1: HERO ─── */}
      <section style={{ textAlign: "center", padding: "80px 48px 60px", maxWidth: "900px", margin: "0 auto" }}>
        <div className="fade-up" style={{ display: "inline-block", padding: "6px 16px", borderRadius: "20px", background: T.accentGlow, color: T.accent, fontSize: "13px", fontWeight: 700, marginBottom: "24px", border: `1px solid ${T.accent}30` }}>
          AI-Powered Presentations
        </div>
        <h1 className="fade-up" style={{ fontSize: "56px", fontWeight: 900, lineHeight: 1.1, letterSpacing: "-0.04em", margin: "0 0 20px" }}>
          Turn your data into<br /><span style={{ color: T.accent }}>stunning presentations</span><br />in minutes
        </h1>
        <p className="fade-up-2" style={{ fontSize: "18px", color: textSec, lineHeight: 1.7, maxWidth: "600px", margin: "0 auto 40px" }}>
          From Excel, templates, or AI prompts — create, customize, and present effortlessly with drag-and-drop precision.
        </p>
        <div className="fade-up-3" style={{ display: "flex", gap: "12px", justifyContent: "center", flexWrap: "wrap" }}>
          {[
            { label: "📊 Start from Excel", action: "excel", desc: "Upload data → AI generates slides" },
            { label: "📚 Start from Template", action: "template", desc: "Choose QBR, KPI, Sales..." },
            { label: "✨ Start with AI", action: "ai", desc: "Describe → Generate instantly" },
          ].map(cta => (
            <button key={cta.action} onClick={() => onEnter(cta.action)}
              style={{ padding: "14px 28px", border: `1px solid ${border}`, borderRadius: "12px", background: cardBg, cursor: "pointer", fontFamily: T.font, textAlign: "left", transition: "all 0.2s", minWidth: "200px" }}
              onMouseEnter={(e) => { e.currentTarget.style.borderColor = T.accent; e.currentTarget.style.transform = "translateY(-2px)"; }}
              onMouseLeave={(e) => { e.currentTarget.style.borderColor = border; e.currentTarget.style.transform = "translateY(0)"; }}>
              <div style={{ fontSize: "15px", fontWeight: 700, color: textPrimary, marginBottom: "4px" }}>{cta.label}</div>
              <div style={{ fontSize: "12px", color: textSec }}>{cta.desc}</div>
            </button>
          ))}
        </div>
      </section>

      {/* ─── SECTION 2: LIVE PREVIEW WITH VARIANTS ─── */}
      <section style={{ padding: "40px 48px 80px", maxWidth: "1000px", margin: "0 auto" }}>
        <div style={{ display: "flex", justifyContent: "center", gap: "8px", marginBottom: "24px" }}>
          {variants.map((v, i) => (
            <button key={v} onClick={() => setActiveVariant(i)}
              style={{ padding: "6px 16px", border: `1px solid ${activeVariant === i ? T.accent : border}`, borderRadius: "8px", background: activeVariant === i ? T.accentGlow : "transparent", color: activeVariant === i ? T.accent : textSec, fontSize: "12px", fontWeight: 600, cursor: "pointer", fontFamily: T.font, transition: "all 0.15s" }}>
              {v}
            </button>
          ))}
        </div>
        <div style={{ borderRadius: "16px", overflow: "hidden", boxShadow: "0 24px 80px rgba(0,0,0,0.2)", border: `1px solid ${border}`, maxWidth: "700px", margin: "0 auto", aspectRatio: "16/9" }}>
          <div style={{ width: "100%", height: "100%", position: "relative" }}>
            <div style={{ position: "absolute", inset: 0, background: activeVariant === 3 ? "linear-gradient(145deg, #070b14, #0f1729)" : activeVariant === 4 ? "linear-gradient(135deg, #f0fdfa, #ccfbf1)" : "linear-gradient(145deg, #f8fafc, #ffffff)", display: "flex", flexDirection: "column", padding: "28px" }}>
              <div style={{ display: "flex", alignItems: "center", gap: "12px", marginBottom: "30px" }}>
                <div style={{ width: "40px", height: "40px", borderRadius: "10px", background: T.accent, display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontWeight: 900, fontSize: "18px" }}>M</div>
                <div>
                  <div style={{ fontSize: "16px", fontWeight: 800, color: activeVariant === 3 ? "#f1f5f9" : "#111827" }}>Microsoft Practice</div>
                  <div style={{ fontSize: "11px", color: activeVariant === 3 ? "#64748b" : "#9ca3af" }}>Quarterly Report • Q1 2026</div>
                </div>
              </div>
              <div style={{ fontSize: "32px", fontWeight: 900, color: activeVariant === 3 ? "#f1f5f9" : "#111827", letterSpacing: "-0.03em", marginBottom: "16px" }}>
                Performance Dashboard
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: "16px", flex: 1 }}>
                {[{ icon: "👥", val: "24", lbl: "Team Size" }, { icon: "💰", val: "$1.2M", lbl: "Revenue" }, { icon: "📈", val: "94%", lbl: "Utilization" }].map(m => (
                  <div key={m.lbl} style={{ background: activeVariant === 3 ? "#111d32" : "#ffffff", border: `1px solid ${activeVariant === 3 ? "rgba(148,163,194,0.15)" : "#e5e7eb"}`, borderRadius: activeVariant === 4 ? "16px" : "12px", padding: "20px", textAlign: "center" }}>
                    <div style={{ fontSize: "24px", marginBottom: "8px" }}>{m.icon}</div>
                    <div style={{ fontSize: "28px", fontWeight: 800, color: T.accent }}>{m.val}</div>
                    <div style={{ fontSize: "12px", color: textSec, marginTop: "4px" }}>{m.lbl}</div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </div>
      </section>

      {/* ─── SECTION 3: FEATURE HIGHLIGHTS ─── */}
      <section style={{ padding: "80px 48px", background: isLight ? "#f8fafc" : "#0c1220" }}>
        <div style={{ maxWidth: "1000px", margin: "0 auto" }}>
          <h2 style={{ fontSize: "36px", fontWeight: 900, textAlign: "center", marginBottom: "12px", letterSpacing: "-0.03em" }}>Everything you need</h2>
          <p style={{ textAlign: "center", color: textSec, fontSize: "16px", marginBottom: "48px" }}>Powerful features that make presentation creation effortless</p>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: "16px" }}>
            {[
              { icon: "🤖", title: "Smart AI", items: ["Improve slides instantly", "Make text concise", "Add visual elements", "Extract insights"] },
              { icon: "🎨", title: "Design Control", items: ["9 font families", "5 design variants", "56+ icons library", "Bullet style picker"] },
              { icon: "📊", title: "Data Power", items: ["Excel import & sync", "Template library", "Data masking", "Auto-create templates"] },
              { icon: "👥", title: "Team & KPI", items: ["Photo management", "KPI dashboards", "Manager scorecards", "Team hierarchy"] },
            ].map(f => (
              <div key={f.title} className="land-card" style={{ background: cardBg, border: `1px solid ${border}`, borderRadius: "16px", padding: "24px", transition: "all 0.25s" }}>
                <div style={{ fontSize: "32px", marginBottom: "12px" }}>{f.icon}</div>
                <div style={{ fontSize: "16px", fontWeight: 800, marginBottom: "12px" }}>{f.title}</div>
                {f.items.map(item => (
                  <div key={item} style={{ fontSize: "13px", color: textSec, padding: "4px 0", display: "flex", gap: "8px", alignItems: "center" }}>
                    <span style={{ color: T.accent, fontSize: "10px" }}>✓</span> {item}
                  </div>
                ))}
              </div>
            ))}
          </div>
        </div>
      </section>

      {/* ─── SECTION 4: HOW IT WORKS ─── */}
      <section style={{ padding: "80px 48px" }}>
        <div style={{ maxWidth: "900px", margin: "0 auto" }}>
          <h2 style={{ fontSize: "36px", fontWeight: 900, textAlign: "center", marginBottom: "48px", letterSpacing: "-0.03em" }}>How it works</h2>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: "24px" }}>
            {[
              { step: "1", icon: "📤", title: "Upload or Describe", desc: "Upload Excel, choose template, or type an AI prompt" },
              { step: "2", icon: "✨", title: "AI Generates", desc: "AI creates structured slides matching your data" },
              { step: "3", icon: "🖱️", title: "Customize", desc: "Drag, edit, style — every element is selectable" },
              { step: "4", icon: "📥", title: "Export & Present", desc: "Download PPTX, PDF, or present directly" },
            ].map(s => (
              <div key={s.step} style={{ textAlign: "center" }}>
                <div style={{ width: "56px", height: "56px", borderRadius: "16px", background: T.accentGlow, display: "flex", alignItems: "center", justifyContent: "center", fontSize: "24px", margin: "0 auto 16px", border: `1px solid ${T.accent}30` }}>{s.icon}</div>
                <div style={{ fontSize: "11px", fontWeight: 800, color: T.accent, marginBottom: "6px", textTransform: "uppercase", letterSpacing: "1px" }}>Step {s.step}</div>
                <div style={{ fontSize: "16px", fontWeight: 700, marginBottom: "8px" }}>{s.title}</div>
                <div style={{ fontSize: "13px", color: textSec, lineHeight: 1.6 }}>{s.desc}</div>
              </div>
            ))}
          </div>
        </div>
      </section>

      {/* ─── SECTION 5: USE CASES ─── */}
      <section style={{ padding: "80px 48px", background: isLight ? "#f8fafc" : "#0c1220" }}>
        <div style={{ maxWidth: "1000px", margin: "0 auto" }}>
          <h2 style={{ fontSize: "36px", fontWeight: 900, textAlign: "center", marginBottom: "48px", letterSpacing: "-0.03em" }}>Built for every team</h2>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: "16px" }}>
            {[
              { icon: "📊", title: "Quarterly Reviews", desc: "Comprehensive QBR with KPIs, billing, team updates", color: "#0078d4" },
              { icon: "📈", title: "KPI Dashboards", desc: "Auto-generated metrics from Excel data", color: "#28a745" },
              { icon: "💼", title: "Sales Decks", desc: "Pipeline analysis, revenue forecasts, client wins", color: "#8b5cf6" },
              { icon: "📋", title: "Project Updates", desc: "Timeline, milestones, resource allocation reports", color: "#0d9488" },
            ].map(uc => (
              <div key={uc.title} className="land-card" style={{ background: cardBg, border: `1px solid ${border}`, borderRadius: "16px", padding: "24px", transition: "all 0.25s", borderTop: `3px solid ${uc.color}` }}>
                <div style={{ fontSize: "28px", marginBottom: "12px" }}>{uc.icon}</div>
                <div style={{ fontSize: "16px", fontWeight: 800, marginBottom: "8px" }}>{uc.title}</div>
                <div style={{ fontSize: "13px", color: textSec, lineHeight: 1.6 }}>{uc.desc}</div>
              </div>
            ))}
          </div>
        </div>
      </section>

      {/* ─── SECTION 6: FINAL CTA ─── */}
      <section style={{ padding: "100px 48px", textAlign: "center" }}>
        <h2 style={{ fontSize: "42px", fontWeight: 900, marginBottom: "16px", letterSpacing: "-0.03em" }}>
          Ready to create?
        </h2>
        <p style={{ fontSize: "18px", color: textSec, marginBottom: "32px" }}>
          Start building your next presentation in under a minute.
        </p>
        <button onClick={() => onEnter("editor")}
          style={{ padding: "16px 40px", background: `linear-gradient(135deg, ${T.accent}, ${T.accentDark})`, color: "#fff", border: "none", borderRadius: "14px", fontSize: "18px", fontWeight: 800, cursor: "pointer", fontFamily: T.font, boxShadow: `0 8px 32px ${T.accentGlow}`, transition: "transform 0.2s" }}
          onMouseEnter={(e) => e.currentTarget.style.transform = "translateY(-2px)"}
          onMouseLeave={(e) => e.currentTarget.style.transform = "translateY(0)"}>
          Create your first presentation →
        </button>
      </section>

      {/* ─── FOOTER ─── */}
      <footer style={{ padding: "24px 48px", borderTop: `1px solid ${border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
          <div style={{ width: "24px", height: "24px", borderRadius: "6px", background: `linear-gradient(135deg, ${T.accent}, ${T.purple})`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: "11px", color: "#fff", fontWeight: 900 }}>P</div>
          <span style={{ fontSize: "13px", fontWeight: 700 }}>PoterPointAI</span>
          <span style={{ fontSize: "11px", color: textSec }}>• ppt.poterai.com</span>
        </div>
        <div style={{ fontSize: "12px", color: textSec }}>Powered by Miracle Software Systems</div>
      </footer>
    </div>
  );
}

function AppInner({ pendingAction, onClearAction, onShowLanding }) {
  const { T, css } = useTheme();
  const slides = useStore((s) => s.slides);
  const activeSlide = useStore((s) => s.activeSlide);
  const canvasMode = useStore((s) => s.canvasMode);
  const themeMode = useStore((s) => s.themeMode);
  const colorCombo = useStore((s) => s.colorCombo);
  const { addSlide, setActive, moveSlide, duplicateSlide, removeSlide, setCanvasMode, setThemeMode, setColorCombo, setDesignVariant, updateSlide } = useStoreActions();
  const [showMenu, setShowMenu] = useState(false);
  const [presenting, setPresenting] = useState(false);
  const [showExport, setShowExport] = useState(false);
  const [showAI, setShowAI] = useState(false);
  const [showDataMgr, setShowDataMgr] = useState(false);
  const [showTemplates, setShowTemplates] = useState(false);

  // Handle pending action from landing page
  useEffect(() => {
    if (!pendingAction) return;
    if (pendingAction === "ai") setShowAI(true);
    else if (pendingAction === "template") setShowTemplates(true);
    else if (pendingAction === "excel") setShowDataMgr(true);
    onClearAction?.();
  }, [pendingAction, onClearAction]);

  const current = slides[activeSlide];
  const currentSchema = current ? SLIDE_SCHEMA[current.type] : null;

  if (presenting !== false) return <PresentationMode slides={slides} startSlide={presenting} onClose={() => setPresenting(false)} />;

  return (
    <div style={{ fontFamily: T.font, background: T.bg, height: "100vh", display: "flex", flexDirection: "column", color: T.text }}>
      <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800;900&family=JetBrains+Mono:wght@400;600&family=DM+Sans:wght@300;400;500;600;700&display=swap" rel="stylesheet" />
      <style>{`
        ::-webkit-scrollbar { width: 6px; height: 6px; }
        ::-webkit-scrollbar-track { background: transparent; }
        ::-webkit-scrollbar-thumb { background: ${T.mode === "light" ? "#cbd5e1" : "#334155"}; border-radius: 3px; }
        ::-webkit-scrollbar-thumb:hover { background: ${T.mode === "light" ? "#94a3b8" : "#475569"}; }
        * { scrollbar-width: thin; scrollbar-color: ${T.mode === "light" ? "#cbd5e1 transparent" : "#334155 transparent"}; }
      `}</style>

      {/* ─── Toolbar ─── */}
      <div style={{
        display: "flex", alignItems: "center", justifyContent: "space-between",
        padding: "10px 20px", borderBottom: `1px solid ${T.border}`,
        background: T.mode === "light" ? T.surface : "rgba(255,255,255,0.01)", backdropFilter: "blur(20px)",
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: "14px" }}>
          <div onClick={onShowLanding} style={{
            width: "34px", height: "34px", borderRadius: T.radiusSm,
            background: `linear-gradient(135deg, ${T.accent}, ${T.purple})`,
            display: "flex", alignItems: "center", justifyContent: "center",
            fontWeight: 900, fontSize: "15px", color: "#fff", cursor: "pointer",
          }}>P</div>
          <div>
            <div onClick={onShowLanding} style={{ fontWeight: 800, fontSize: "15px", letterSpacing: "-0.02em", cursor: "pointer" }}>PoterPointAI</div>
            <div style={{ color: T.textDim, fontSize: "10px", fontFamily: T.fontMono }}>{slides.length} slide{slides.length !== 1 ? "s" : ""} • ppt.poterai.com</div>
          </div>
        </div>

        {/* ─── Theme Controls ─── */}
        <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
          {/* Color Combo Picker */}
          <div style={{ display: "flex", alignItems: "center", gap: "5px" }}>
            {Object.entries(COLOR_COMBOS).map(([key, c]) => (
              <button key={key} onClick={() => setColorCombo(key)} title={c.label}
                style={{
                  width: "18px", height: "18px", borderRadius: "50%", border: "2px solid",
                  borderColor: colorCombo === key ? T.text : "transparent",
                  background: c.accent, cursor: "pointer", padding: 0,
                  boxShadow: colorCombo === key ? `0 0 0 2px ${T.bg}, 0 0 0 4px ${c.accent}` : "none",
                  transition: "all 0.15s",
                }} />
            ))}
          </div>

          {/* Theme Mode Toggle */}
          <button onClick={() => setThemeMode(themeMode === "light" ? "dark" : "light")}
            style={{
              ...css.btnIcon, width: "auto", padding: "0 12px", gap: "6px",
              display: "flex", alignItems: "center", fontSize: "12px", fontWeight: 600, fontFamily: T.font,
            }}>
            {themeMode === "light" ? "☀" : "☾"} {themeMode === "light" ? "Light" : "Dark"}
          </button>
        </div>

        <div style={{ display: "flex", gap: "8px", position: "relative" }}>
          <button onClick={() => setShowDataMgr(true)}
            style={{ ...css.btnIcon, width: "auto", padding: "0 14px", fontSize: "11px", fontWeight: 600, fontFamily: T.font, gap: "6px", display: "flex", alignItems: "center" }}>
            <span style={{ fontSize: "13px" }}>📂</span> Data
          </button>
          <button onClick={() => setShowAI(true)}
            style={{ ...css.btnIcon, width: "auto", padding: "0 14px", fontSize: "11px", fontWeight: 600, fontFamily: T.font, gap: "6px", display: "flex", alignItems: "center",
              background: `linear-gradient(135deg, ${T.accentGlow}, rgba(139,92,246,0.12))`,
              borderColor: T.purple,
            }}>
            <span style={{ fontSize: "13px" }}>✨</span> AI Generate
          </button>
          <button onClick={() => setShowExport(!showExport)}
            style={{ ...css.btnIcon, width: "auto", padding: "0 14px", fontSize: "11px", fontWeight: 600, fontFamily: T.font, gap: "6px", display: "flex", alignItems: "center",
              background: showExport ? T.accentGlow : "transparent",
              borderColor: showExport ? "rgba(59,130,246,0.3)" : T.border,
            }}>
            <span style={{ fontSize: "13px" }}>↓</span> Export
          </button>
          {showExport && <ExportDropdown slides={slides} onClose={() => setShowExport(false)} />}
          <button onClick={() => setPresenting(0)} style={css.btnPrimary}>▶ Present</button>
        </div>
      </div>

      {/* AI Generator Modal */}
      {showAI && <AIGeneratorModal onClose={() => setShowAI(false)} />}
      {showDataMgr && <DataManagerModal onClose={() => setShowDataMgr(false)} />}
      {showTemplates && <TemplateGalleryModal onClose={() => setShowTemplates(false)} onSelectTemplate={(tmpl) => { setShowTemplates(false); setShowAI(true); }} />}

      {/* ─── Main Layout ─── */}
      <div style={{ display: "flex", flex: 1, overflow: "hidden" }}>

        {/* LEFT PANEL — Slide Navigator */}
        <div style={{
          width: "190px", borderRight: `1px solid ${T.border}`,
          background: T.mode === "light" ? T.surface : "rgba(255,255,255,0.008)", display: "flex", flexDirection: "column",
        }}>
          <div style={{
            padding: "10px 12px", borderBottom: `1px solid ${T.border}`,
            display: "flex", justifyContent: "space-between", alignItems: "center",
          }}>
            <span style={{ fontSize: "9px", fontWeight: 700, textTransform: "uppercase", letterSpacing: "1.5px", color: T.textMuted }}>Navigator</span>
            <div style={{ position: "relative" }}>
              <button onClick={() => setShowMenu(!showMenu)}
                style={{ ...css.btnIcon, width: "26px", height: "26px", fontSize: "15px", color: T.accent, borderColor: "rgba(59,130,246,0.2)" }}>+</button>
              {showMenu && <SlideTypeSelector onSelect={addSlide} onClose={() => setShowMenu(false)} />}
            </div>
          </div>
          <div style={{ flex: 1, overflow: "auto", padding: "8px" }}>
            {slides.map((slide, i) => (
              <div key={slide.id} onClick={() => setActive(i)}
                style={{
                  borderRadius: T.radius, overflow: "hidden", marginBottom: "6px",
                  border: i === activeSlide ? `2px solid ${T.accent}` : `2px solid transparent`,
                  cursor: "pointer", transition: "border-color 0.15s",
                  boxShadow: i === activeSlide ? `0 0 0 3px ${T.accentGlow}` : "none",
                }}>
                <div style={{ aspectRatio: "16/9", position: "relative" }}>
                  <SlideRenderer slide={slide} compact={true} index={i} total={slides.length} />
                  <div style={{
                    position: "absolute", bottom: "3px", left: "3px",
                    background: "rgba(0,0,0,0.7)", borderRadius: "4px",
                    padding: "1px 5px", fontSize: "7px", color: "rgba(255,255,255,0.6)",
                    fontWeight: 700, fontFamily: T.fontMono,
                  }}>{i + 1}</div>
                  <div style={{
                    position: "absolute", bottom: "3px", right: "3px",
                    background: "rgba(0,0,0,0.7)", borderRadius: "4px",
                    padding: "1px 5px", fontSize: "6px", color: T.textDim, fontWeight: 600,
                  }}>{SLIDE_SCHEMA[slide.type]?.label}</div>
                </div>
              </div>
            ))}
          </div>
        </div>

        {/* CENTER — Live Preview / Canvas */}
        <div style={{
          flex: 1, display: "flex", flexDirection: "column",
          background: T.mode === "light" ? "#e5e7eb" : "rgba(0,0,0,0.25)", overflow: "hidden",
        }}>
          {/* Mode Toggle Bar */}
          <div style={{
            display: "flex", alignItems: "center", justifyContent: "center",
            padding: "8px 20px", borderBottom: `1px solid ${T.border}`,
            background: "rgba(255,255,255,0.01)",
          }}>
            <div style={{
              display: "flex", background: "rgba(255,255,255,0.04)",
              borderRadius: "8px", padding: "2px", border: `1px solid ${T.border}`,
            }}>
              {[{ key: false, label: "Form", icon: "✎" }, { key: true, label: "Canvas", icon: "⊞" }, { key: "flow", label: "Flow", icon: "↕" }].map((m) => (
                <button key={String(m.key)} onClick={() => setCanvasMode(m.key)}
                  style={{
                    padding: "5px 16px", border: "none", cursor: "pointer",
                    borderRadius: "6px", fontSize: "11px", fontWeight: 700,
                    fontFamily: T.font, display: "flex", alignItems: "center", gap: "6px",
                    background: canvasMode === m.key ? T.accentGlow : "transparent",
                    color: canvasMode === m.key ? T.accent : T.textMuted,
                    transition: "all 0.15s",
                  }}>
                  <span style={{ fontSize: "13px" }}>{m.icon}</span> {m.label}
                </button>
              ))}
            </div>
          </div>

          {/* Content area — switches between preview and canvas */}
          {canvasMode === true && current ? (
            <CanvasEditor slide={current} slideIndex={activeSlide} totalSlides={slides.length} />
          ) : canvasMode === "flow" ? (
            /* ─── Flow View: Vertical scroll through all slides (Gamma-style) ─── */
            <div style={{ flex: 1, overflow: "auto", padding: "24px", display: "flex", flexDirection: "column", alignItems: "center", gap: "16px" }}>
              {slides.map((sl, i) => (
                <div key={sl.id} onClick={() => setActive(i)}
                  style={{
                    width: "100%", maxWidth: "860px", aspectRatio: "16/9",
                    borderRadius: "12px", overflow: "hidden", cursor: "pointer",
                    boxShadow: i === activeSlide ? `0 0 0 3px ${T.accent}, 0 12px 40px rgba(0,0,0,0.25)` : `0 4px 20px rgba(0,0,0,0.15)`,
                    transition: "box-shadow 0.2s ease, transform 0.2s ease",
                    transform: i === activeSlide ? "scale(1)" : "scale(0.98)",
                    flexShrink: 0,
                  }}>
                  <SlideRenderer slide={sl} index={i} total={slides.length} />
                </div>
              ))}
            </div>
          ) : (
            <div style={{
              flex: 1, display: "flex", flexDirection: "column",
              alignItems: "center", justifyContent: "center",
              padding: "28px", overflow: "auto",
            }}>
              <div style={{
                width: "100%", maxWidth: "860px", aspectRatio: "16/9",
                borderRadius: "14px", overflow: "hidden",
                boxShadow: `0 24px 80px rgba(0,0,0,0.45), 0 0 0 1px ${T.border}`,
              }}>
                {current && <SlideRenderer slide={current} index={activeSlide} total={slides.length} />}
              </div>
              <div style={{ display: "flex", gap: "5px", marginTop: "18px" }}>
                {slides.map((_, i) => (
                  <button key={i} onClick={() => setActive(i)}
                    style={{
                      width: i === activeSlide ? "22px" : "7px", height: "7px",
                      borderRadius: "4px", cursor: "pointer", border: "none",
                      background: i === activeSlide ? T.accent : "rgba(255,255,255,0.12)",
                      transition: "all 0.25s ease",
                    }} />
                ))}
              </div>
            </div>
          )}
        </div>

        {/* RIGHT PANEL — Editor */}
        <div style={{
          width: "340px", borderLeft: `1px solid ${T.border}`,
          background: T.mode === "light" ? T.surface : "rgba(255,255,255,0.008)", display: "flex", flexDirection: "column",
        }}>
          <div style={{
            padding: "12px 16px", borderBottom: `1px solid ${T.border}`,
            display: "flex", justifyContent: "space-between", alignItems: "center",
          }}>
            <div>
              <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                {currentSchema && <span style={{ fontSize: "14px" }}>{currentSchema.icon}</span>}
                <span style={{ fontSize: "13px", fontWeight: 700 }}>{currentSchema?.label || "Slide"}</span>
              </div>
              <div style={{ fontSize: "10px", color: T.textDim, fontFamily: T.fontMono, marginTop: "2px" }}>
                Slide {activeSlide + 1} of {slides.length}
              </div>
            </div>
            <div style={{ display: "flex", gap: "3px" }}>
              <button onClick={() => setPresenting(activeSlide)} title="Present from this slide"
                style={{ ...css.btnIcon, width: "28px", height: "28px", fontSize: "11px", color: T.accent, borderColor: T.accent + "40" }}>▶</button>
              <button onClick={() => moveSlide(activeSlide, -1)} disabled={activeSlide === 0} title="Move up"
                style={{ ...css.btnIcon, width: "28px", height: "28px", fontSize: "11px", opacity: activeSlide === 0 ? 0.3 : 1 }}>↑</button>
              <button onClick={() => moveSlide(activeSlide, 1)} disabled={activeSlide === slides.length - 1} title="Move down"
                style={{ ...css.btnIcon, width: "28px", height: "28px", fontSize: "11px", opacity: activeSlide === slides.length - 1 ? 0.3 : 1 }}>↓</button>
              <button onClick={() => duplicateSlide(activeSlide)} title="Duplicate"
                style={{ ...css.btnIcon, width: "28px", height: "28px", fontSize: "11px" }}>⧉</button>
              <button onClick={() => removeSlide(activeSlide)} disabled={slides.length <= 1} title="Delete"
                style={{ ...css.btnIcon, width: "28px", height: "28px", fontSize: "11px", borderColor: "rgba(239,68,68,0.2)", color: T.danger, opacity: slides.length <= 1 ? 0.3 : 1 }}>✕</button>
            </div>
          </div>
          <div style={{ flex: 1, overflow: "auto", padding: "16px" }}>
            {/* AI Quick Actions */}
            {current && <AIQuickActions slide={current} updateSlide={updateSlide} T={T} />}
            {current && (
              <div style={{ marginBottom: "0" }}>
              </div>
            )}
            {current && <SlideForm slide={current} />}
          </div>
        </div>
      </div>
    </div>
  );
}

export default function App() {
  const [store] = useState(() =>
    createStore({
      presentation: { title: "Microsoft Practice Quarterly Status Report", subtitle: "Q1 2026" },
      slides: [
        // SLIDE 1: Title + Leadership
        { id: "s1", type: "title", title: "Microsoft Practice Quarterly Status Report", subtitle: "2026-Q1 — January to March 2026", badge: "Q1 2026",
          metrics: [{ number: "7", label: "Resources" }, { number: "4", label: "Projects" }, { number: "3", label: "New Customers" }],
          highlights: [
            { color: "#0078d4", text: "7 new engagements across 4 projects" },
            { color: "#28a745", text: "3 new customers onboarded" },
            { color: "#ffc107", text: "5 webinars hosted, 3 blogs published" },
            { color: "#8b5cf6", text: "3 certifications achieved" },
          ],
          stats: [{ icon: "👥", label: "Team Size", value: "7 Resources" }, { icon: "📈", label: "Growth Rate", value: "+15% QoQ" }, { icon: "💰", label: "Revenue", value: "$105,400" }],
        },
        // SLIDE 2: Agenda
        { id: "s2", type: "agenda", title: "Agenda", subtitle: "Q1 2026 Quarterly Review — Key Areas of Focus",
          cards: [
            { icon: "👥", title: "Resource & Billing", items: ["Resource allocation", "Billing status & metrics", "Utilization analysis"] },
            { icon: "🤝", title: "New Engagements", items: ["7 new engagements", "3 new customers", "Resource assignments"] },
            { icon: "⚙️", title: "Process Improvements", items: ["Learn Together Initiative", "Skill development", "Automation POCs"] },
            { icon: "🏆", title: "Achievements", items: ["5 webinars hosted", "3 blogs published", "3 certifications"] },
            { icon: "🔍", title: "Sales Enquiries", items: ["8 new enquiries", "POC demonstrations", "Proposal submissions"] },
            { icon: "🧠", title: "Training & POCs", items: ["Azure AI & RAG", "4 POCs developed", "Sales enablement"] },
            { icon: "⭐", title: "Appreciations", items: ["Team recognitions", "Client feedback", "Performance awards"] },
            { icon: "📅", title: "Team Events", items: ["Team outings", "Team building", "Celebrations"] },
          ],
        },
        // SLIDE 3: Resource vs Billing
        { id: "s3", type: "table", title: "Resource vs Billing", subtitle: "Monthly resource allocation and billing breakdown — Q1 2026",
          columns: ["Category", "January", "February", "March", "Q1 Total"],
          rows: [
            { cells: ["Total Resources", "7", "7", "7", "7"] },
            { cells: ["Billable", "5", "6", "6", "6"] },
            { cells: ["Non-Billable", "1", "1", "1", "1"] },
            { cells: ["Shadow", "1", "0", "0", "0"] },
            { cells: ["Trainees", "0", "0", "0", "0"] },
          ],
          summaryStats: [{ number: "7", label: "Resources" }, { number: "86%", label: "Avg Utilization" }, { number: "$105K", label: "Revenue" }],
        },
        // SLIDE 4: Team Structure
        { id: "s4", type: "section-overview", title: "Team Structure", subtitle: "Organizational hierarchy — Q1 2026",
          sectionIcon: "👥", sectionTitle: "Team Hierarchy",
          description: "Our Microsoft Practice team structure with clear reporting lines from VP through Directors, Managers, and individual contributors.",
          stats: [{ number: "1", label: "VP" }, { number: "1", label: "Director" }, { number: "3", label: "Managers" }, { number: "7", label: "Resources" }],
          billingMetrics: [{ label: "VP", value: "Leadership" }, { label: "Director", value: "Delivery Head" }, { label: "Manager — Azure", value: "3 Resources" }, { label: "Manager — D365", value: "2 Resources" }, { label: "Manager — QA", value: "2 Resources" }],
        },
        // SLIDE 5: New Engagements
        { id: "s5", type: "table", title: "New Engagements", subtitle: "Q1 2026 — Resource allocation and project details",
          columns: ["Project", "No. of Resources", "Details"],
          rows: [
            { cells: ["Client A — Cloud Portal", "1", "Cloud and Customer Portal App Support"] },
            { cells: ["Client B — SharePoint", "4", "Classic to Modern SharePoint Migration"] },
            { cells: ["Client C — D365", "1", "Dynamics 365 Manufacturing F&O Consultant"] },
            { cells: ["Client D — QA", "1", "QA Automation & Testing"] },
          ],
          summaryStats: [{ number: "4", label: "Projects" }, { number: "7", label: "Resources" }, { number: "100%", label: "Utilization" }],
        },
        // SLIDE 6: Active Project Portfolio
        { id: "s6", type: "content", title: "Active Project Portfolio", subtitle: "Current strategic focus areas — Q1 2026",
          cards: [
            { icon: "☁️", title: "Cloud Migration Program", status: "Active", items: [{ bold: "Client:", text: "Enterprise Corp — 12 workloads" }, { bold: "Resources:", text: "3 engineers deployed" }, { bold: "Timeline:", text: "Q1-Q2 2026" }] },
            { icon: "📱", title: "SharePoint Modernization", status: "Active", items: [{ bold: "Client:", text: "Global Retail — 50+ sites" }, { bold: "Resources:", text: "4 developers" }, { bold: "Status:", text: "Phase 2 in progress" }] },
            { icon: "🔧", title: "D365 Implementation", status: "In Progress", items: [{ bold: "Client:", text: "Manufacturing Inc" }, { bold: "Resources:", text: "1 functional consultant" }, { bold: "Module:", text: "Finance & Operations" }] },
            { icon: "🧪", title: "QA Automation", status: "Active", items: [{ bold: "Client:", text: "FinTech Solutions" }, { bold: "Resources:", text: "1 QA engineer" }, { bold: "Framework:", text: "Playwright + Selenium" }] },
          ],
        },
        // SLIDE 7: Process Improvements
        { id: "s7", type: "content", title: "Process Improvements", subtitle: "Q1 2026 — Key initiatives and enhancements",
          cards: [
            { icon: "🎓", title: "Learn Together Initiative", status: "Active", items: [{ bold: "Collaborative Learning:", text: "Daily AI topics assigned to team members" }, { bold: "Sales Enablement:", text: "Strategic approach for sales team insights" }, { bold: "Monitoring:", text: "Team leaders tracking contributions" }] },
            { icon: "💻", title: "Technical Development", status: "In Progress", items: [{ bold: "Node.js & React Native:", text: "APIs built, mobile app developed" }, { bold: "TCPOS Automation:", text: "POC for Paradies Lagardère" }, { bold: "Agentic AI Testing:", text: "Autonomous AI agents for QA" }] },
          ],
        },
        // SLIDE 8: Performance Dashboard
        { id: "s8", type: "dashboard", title: "Performance Dashboard", subtitle: "Key metrics at a glance — Q1 2026",
          metrics: [
            { icon: "👥", value: "7", label: "Team Size", detail: "Fully staffed" },
            { icon: "💰", value: "$105K", label: "Revenue", detail: "+18% QoQ" },
            { icon: "📈", value: "94%", label: "Utilization", detail: "Target: 90%" },
            { icon: "⭐", value: "4.8", label: "CSAT Score", detail: "Industry: 4.2" },
            { icon: "🏆", value: "12", label: "Projects Delivered", detail: "+5 vs last Q" },
            { icon: "🔄", value: "98%", label: "Retention Rate", detail: "All-time high" },
          ],
        },
        // SLIDE 9: Sales Enquiries
        { id: "s9", type: "table", title: "Sales Enquiries & Pipeline", subtitle: "Q1 2026 — Active sales opportunities",
          columns: ["Company", "Requirement", "Technology", "Status"],
          rows: [
            { cells: ["Client E", "WordPress to React Migration", "Umbraco, React", "On Hold"] },
            { cells: ["Client F", "Application Development", "SharePoint, Power Platform", "Started Mar 9"] },
            { cells: ["Client G", "SharePoint Migration", "SharePoint Online", "On Hold"] },
            { cells: ["Client H", "Unified Portal", "ReactJS", "In Progress"] },
            { cells: ["Client I", "D365 & E-commerce", "Dynamics 365, OROCommerce", "Onboarded"] },
            { cells: ["Client J", "UI Path to Power Automate", "Power Automate Desktop", "Awaiting Setup"] },
          ],
          summaryStats: [{ number: "8", label: "Total Enquiries" }, { number: "3", label: "In Progress" }, { number: "2", label: "Onboarded" }],
        },
        // SLIDE 10: Blogs & Webinars
        { id: "s10", type: "table", title: "Blogs & Webinars", subtitle: "Q1 2026 — Knowledge sharing sessions and publications",
          columns: ["Type", "Title / Topic", "Author / Speaker", "Date"],
          rows: [
            { cells: ["📝 Blog", "Using Generative AI Responsibly", "Ramji Bodda", "Mar 2026"] },
            { cells: ["📝 Blog", "Implementing RAG in .NET", "Sowjanya Kolli", "Feb 2026"] },
            { cells: ["📝 Blog", "Applitools Eyes: AI Visual Testing", "Praveen Pallapati", "Jan 2026"] },
            { cells: ["🎤 Webinar", "Azure AI Search", "Ashok Kotu", "Mar 2026"] },
            { cells: ["🎤 Webinar", "Azure Data Factory", "Sahithi Korupolu", "Feb 2026"] },
            { cells: ["🎤 Webinar", "API Testing with Playwright", "Rajesh Gullipalli", "Jan 2026"] },
          ],
          summaryStats: [{ number: "3", label: "Blogs" }, { number: "5", label: "Webinars" }, { number: "100%", label: "Published" }],
        },
        // SLIDE 11: Certifications & Training
        { id: "s11", type: "dashboard", title: "Certifications & Training", subtitle: "Q1 2026 — Team skill development",
          metrics: [
            { icon: "🎓", value: "2", label: "Certifications", detail: "PL-600 & AZ-900" },
            { icon: "📚", value: "2", label: "Trainings", detail: "Azure AI, Playwright" },
            { icon: "👨‍💻", value: "15", label: "Resources Trained", detail: "100% completion" },
            { icon: "🧪", value: "4", label: "POCs Created", detail: "Business cases" },
          ],
        },
        // SLIDE 12: POCs & Enablement
        { id: "s12", type: "content", title: "POCs & Sales Enablement", subtitle: "Q1 2026 — Proof of concepts and sales sessions",
          cards: [
            { icon: "🧩", title: "SharePoint Components", status: "Complete", items: [{ bold: "Custom web parts:", text: "Modern SharePoint site development" }, { bold: "Team:", text: "Bhavya, Gopi, Surya" }] },
            { icon: "🔄", title: "JIRA Cloud Migration", status: "Complete", items: [{ bold: "Scope:", text: "Cloud-to-cloud JIRA migration" }, { bold: "Team:", text: "Bhavya Geddamuri" }] },
            { icon: "🖥️", title: "Desktop Automation", status: "Complete", items: [{ bold: "Python automation:", text: "POC on Aronium sample app" }, { bold: "Team:", text: "Praveen Pallapati" }] },
            { icon: "🌐", title: "Cross-Browser Testing", status: "Complete", items: [{ bold: "Playwright + Jenkins:", text: "Parallel cross-browser automation" }, { bold: "Team:", text: "Pavan Gedela" }] },
          ],
        },
        // SLIDE 13: Monthly Awards
        { id: "s13", type: "highlight-list", title: "Monthly Awards & Star Performers", subtitle: "Q1 2026 — Celebrating excellence",
          items: [
            { color: "#ffc107", text: "⭐ January Star Performer — Outstanding delivery on Cloud Migration project" },
            { color: "#0078d4", text: "🏆 January Best Innovator — Agentic AI Testing POC development" },
            { color: "#28a745", text: "⭐ February Star Performer — SharePoint Modernization phase completion" },
            { color: "#8b5cf6", text: "🏆 February Best Collaborator — Cross-team knowledge sharing initiative" },
            { color: "#0d9488", text: "⭐ March Star Performer — D365 F&O successful go-live support" },
            { color: "#e11d48", text: "🏆 March Innovation Award — Desktop Automation Python POC" },
          ],
        },
        // SLIDE 14: Manager KPI
        { id: "s14", type: "section-overview", title: "Manager KPI & Achievements", subtitle: "Q1 2026 — Leadership performance",
          sectionIcon: "📊", sectionTitle: "Manager Scorecard",
          description: "Quarterly KPI scores and key achievements for each practice manager, aligned with organizational goals.",
          stats: [{ number: "4.5", label: "Avg KPI Score" }, { number: "3", label: "Managers" }, { number: "100%", label: "Goals Met" }],
          billingMetrics: [
            { label: "Manager — Azure Practice", value: "KPI: 4.6 / 5.0" },
            { label: "Manager — D365 Practice", value: "KPI: 4.4 / 5.0" },
            { label: "Manager — QA Practice", value: "KPI: 4.5 / 5.0" },
            { label: "Key Achievement", value: "Zero critical incidents" },
          ],
        },
        // SLIDE 15: Team Activities
        { id: "s15", type: "dashboard", title: "Team Activities & Culture", subtitle: "Q1 2026 — Events, celebrations, and team bonding",
          metrics: [
            { icon: "🍽️", value: "Team Lunch", label: "January", detail: "All hands gathering" },
            { icon: "🥾", value: "Team Outing", label: "February", detail: "Adventure park visit" },
            { icon: "🎂", value: "Birthdays", label: "Quarterly", detail: "3 celebrations" },
            { icon: "🎉", value: "Office Party", label: "March", detail: "Quarter close celebration" },
            { icon: "🧩", value: "Team Building", label: "Monthly", detail: "Problem-solving workshops" },
            { icon: "🌟", value: "Recognition", label: "Quarterly", detail: "Awards ceremony" },
          ],
        },
        // SLIDE 16: Key Highlights Summary
        { id: "s16", type: "highlight-list", title: "Key Highlights — Q1 2026", subtitle: "Revenue achievements, client wins, certifications, and strategic milestones",
          items: [
            { color: "#28a745", text: "💰 Revenue exceeded target by 18%, reaching $105,400 for Q1" },
            { color: "#0078d4", text: "🤝 Successfully onboarded 3 new enterprise clients" },
            { color: "#8b5cf6", text: "🚀 Launched AI/ML practice with 2 pilot POC projects" },
            { color: "#0d9488", text: "📈 Team utilization hit 94% — highest in 6 quarters" },
            { color: "#ffc107", text: "🛡️ Zero critical incidents across all managed environments" },
            { color: "#e11d48", text: "🏆 Won Microsoft Partner of the Year nomination" },
          ],
        },
        // SLIDE 17: Thank You
        { id: "s17", type: "thank-you", title: "Thank You", message: "We are grateful for the continued support and guidance from our leadership team. Your insights and direction have been invaluable in helping us achieve our goals. We look forward to building on our accomplishments and delivering even greater value in Q2 2026.", signature: "Microsoft Practice Team" },
      ],
      activeSlide: 0,
      canvasMode: false,
      canvasSelectedId: null,
      canvasSelectedIds: [],
      themeMode: "light",
      colorCombo: "blue",
      designVariant: "corporate",
    })
  );

  const [view, setView] = useState("landing"); // "landing" | "editor"
  const [pendingAction, setPendingAction] = useState(null);

  const handleEnter = useCallback((action) => {
    setView("editor");
    if (action && action !== "editor") setPendingAction(action);
  }, []);

  return (
    <StoreContext.Provider value={store}>
      {view === "landing" ? (
        <LandingPage onEnter={handleEnter} />
      ) : (
        <AppInner pendingAction={pendingAction} onClearAction={() => setPendingAction(null)} onShowLanding={() => setView("landing")} />
      )}
    </StoreContext.Provider>
  );
}
