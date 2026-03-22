/* ============================================================
   PPT ToolBox — Office Web Add-in
   Targets: PowerPoint on Mac, Windows, Web (Microsoft 365)
   Requires: PowerPoint JS API 1.5 (Office 2021 / M365)
   ============================================================ */

import "./taskpane.css";

// ── Swatch storage (localStorage) ───────────────────────────────
const SWATCH_KEY = "ppttoolbox_swatches";
const DEFAULT_SWATCHES = [
  "#1A1A2E","#E94F37","#FFFFFF","#000000",
  "#2E86AB","#F6C90E","#3DC47E","#A83F9E",
  "#FF6B35","#04A777","#D62246","#C0C0C0",
];

function loadSwatches(): string[] {
  try {
    const stored = localStorage.getItem(SWATCH_KEY);
    if (stored) {
      const parsed: string[] = JSON.parse(stored);
      if (Array.isArray(parsed) && parsed.length === 12) return parsed;
    }
  } catch { /* ignore */ }
  return [...DEFAULT_SWATCHES];
}

function saveSwatches(swatches: string[]): void {
  localStorage.setItem(SWATCH_KEY, JSON.stringify(swatches));
}

// ── Globals ──────────────────────────────────────────────────────
let swatches = loadSwatches();
let swatchEditMode = false;

// ── Entry point ──────────────────────────────────────────────────
Office.onReady(() => {
  initTabs();
  buildAllSwatches();
  attachHandlers();
});

// ── Status helpers ───────────────────────────────────────────────
function showStatus(msg: string, isError = false): void {
  const el = document.getElementById("status-bar")!;
  el.textContent = msg;
  el.className = isError ? "status-error" : "status-ok";
  setTimeout(() => { el.textContent = ""; el.className = ""; }, 3000);
}

// ── Safe wrapper ─────────────────────────────────────────────────
async function safe(fn: () => Promise<void>): Promise<void> {
  try {
    await fn();
  } catch (err: any) {
    showStatus(err.message ?? "An error occurred", true);
  }
}

// ── Tab switching ─────────────────────────────────────────────────
function initTabs(): void {
  document.querySelectorAll<HTMLButtonElement>(".tab-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      const tab = btn.dataset.tab!;
      document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
      document.querySelectorAll(".tab-panel").forEach(p => p.classList.remove("active"));
      btn.classList.add("active");
      document.getElementById("tab-" + tab)!.classList.add("active");
    });
  });
}

// ── Swatch rendering ─────────────────────────────────────────────
function buildSwatchPanel(
  panelId: string,
  onClick: (color: string, index: number, btn: HTMLButtonElement) => void
): void {
  const panel = document.getElementById(panelId)!;
  panel.innerHTML = "";
  swatches.forEach((color, idx) => {
    const btn = document.createElement("button");
    btn.className = "swatch-btn";
    btn.style.background = color;
    btn.title = color;
    btn.addEventListener("click", () => onClick(color, idx, btn));
    panel.appendChild(btn);
  });
}

function buildAllSwatches(): void {
  buildSwatchPanel("swatches-fill", (color, idx) => {
    if (swatchEditMode) openSwatchEditor(idx);
    else safe(() => applyFillColor(color));
  });
  buildSwatchPanel("swatches-font", (color, idx) => {
    if (swatchEditMode) openSwatchEditor(idx);
    else safe(() => applyFontColor(color));
  });
  buildSwatchPanel("swatches-outline", (color, idx) => {
    if (swatchEditMode) openSwatchEditor(idx);
    else safe(() => applyOutlineColor(color));
  });
}

function openSwatchEditor(index: number): void {
  const picker = document.createElement("input");
  picker.type = "color";
  picker.value = swatches[index];
  picker.style.display = "none";
  document.body.appendChild(picker);
  picker.addEventListener("input", () => {
    swatches[index] = picker.value;
    saveSwatches(swatches);
    buildAllSwatches();
    if (swatchEditMode) toggleEditMode(true); // re-apply dashed border
  });
  picker.click();
  picker.addEventListener("blur", () => document.body.removeChild(picker), { once: true });
}

function toggleEditMode(force?: boolean): void {
  swatchEditMode = force !== undefined ? force : !swatchEditMode;
  const btn = document.getElementById("btn-edit-swatches")!;
  btn.textContent  = swatchEditMode ? "Done Editing" : "Edit Swatches";
  (btn as HTMLButtonElement).style.background = swatchEditMode ? "#E94F37" : "";
  (btn as HTMLButtonElement).style.color      = swatchEditMode ? "#fff"    : "";

  ["swatches-fill", "swatches-font", "swatches-outline"].forEach(id => {
    const panel = document.getElementById(id)!;
    panel.classList.toggle("swatch-edit-mode", swatchEditMode);
  });
}

// ── Slide size ────────────────────────────────────────────────────
async function getSlideDimensions(): Promise<{ w: number; h: number }> {
  return await PowerPoint.run(async context => {
    const pres = context.presentation as any;
    pres.load("slideSizeType");
    await context.sync();
    const t = pres.slideSizeType as string;
    const map: Record<string, { w: number; h: number }> = {
      Widescreen:        { w: 960,   h: 540   },
      OnScreen:          { w: 720,   h: 540   },
      OnScreenShow16x9:  { w: 960,   h: 540   },
      OnScreenShow4x3:   { w: 720,   h: 540   },
      LetterPaper:       { w: 792,   h: 612   },
      A4Paper:           { w: 780.5, h: 540   },
      Custom:            { w: 960,   h: 540   },
    };
    return map[t] ?? { w: 960, h: 540 };
  });
}

// ── cm ↔ pt ───────────────────────────────────────────────────────
const cmToPt  = (s: string) => parseFloat(s.replace(",", ".")) / 2.54 * 72;
const ptToCm  = (pt: number) => (pt * 2.54 / 72).toFixed(2);

// ── Get selected shapes helper ────────────────────────────────────
async function withShapes<T>(
  props: string[],
  fn: (shapes: PowerPoint.Shape[], context: PowerPoint.RequestContext) => Promise<T>
): Promise<T> {
  return PowerPoint.run(async context => {
    const sel = context.presentation.getSelectedShapes();
    sel.load(props as any);
    await context.sync();
    if (sel.items.length === 0) throw new Error("No shapes selected");
    return fn(sel.items, context);
  });
}

// ════════════════════════════════════════════════════════════════
//  ARRANGE
// ════════════════════════════════════════════════════════════════

async function alignShapes(type: string): Promise<void> {
  const { w: slideW, h: slideH } = await getSlideDimensions();
  await withShapes(["left","top","width","height"], async shapes => {
    for (const s of shapes) {
      switch (type) {
        case "left":    s.left = 0;                           break;
        case "centerh": s.left = (slideW - s.width)  / 2;   break;
        case "right":   s.left =  slideW - s.width;          break;
        case "top":     s.top  = 0;                           break;
        case "middlev": s.top  = (slideH - s.height) / 2;   break;
        case "bottom":  s.top  =  slideH - s.height;         break;
      }
    }
    showStatus("Aligned");
  });
}

async function distributeShapes(horizontal: boolean): Promise<void> {
  await withShapes(["left","top","width","height"], async shapes => {
    if (shapes.length < 3) throw new Error("Select 3 or more shapes to distribute");
    const arr = [...shapes];
    if (horizontal) {
      arr.sort((a, b) => a.left - b.left);
      const span  = (arr[arr.length - 1].left + arr[arr.length - 1].width) - arr[0].left;
      const total = arr.reduce((s, sh) => s + sh.width, 0);
      const gap   = (span - total) / (arr.length - 1);
      let x = arr[0].left;
      arr.forEach(sh => { sh.left = x; x += sh.width + gap; });
    } else {
      arr.sort((a, b) => a.top - b.top);
      const span  = (arr[arr.length - 1].top + arr[arr.length - 1].height) - arr[0].top;
      const total = arr.reduce((s, sh) => s + sh.height, 0);
      const gap   = (span - total) / (arr.length - 1);
      let y = arr[0].top;
      arr.forEach(sh => { sh.top = y; y += sh.height + gap; });
    }
    showStatus("Distributed");
  });
}

async function zOrder(dir: string): Promise<void> {
  await withShapes(["id"], async shapes => {
    const map: Record<string, any> = {
      fwd:    PowerPoint.ShapeZOrder.bringForward,
      back:   PowerPoint.ShapeZOrder.sendBackward,
      front:  PowerPoint.ShapeZOrder.bringToFront,
      toback: PowerPoint.ShapeZOrder.sendToBack,
    };
    shapes.forEach(s => (s as any).incrementZOrder(map[dir]));
    showStatus("Z-order updated");
  });
}

async function matchSize(matchW: boolean, matchH: boolean): Promise<void> {
  await withShapes(["width","height"], async shapes => {
    if (shapes.length < 2) throw new Error("Select 2 or more shapes");
    const refW = shapes[0].width, refH = shapes[0].height;
    for (let i = 1; i < shapes.length; i++) {
      if (matchW) shapes[i].width  = refW;
      if (matchH) shapes[i].height = refH;
    }
    showStatus("Size matched");
  });
}

async function readGeometry(): Promise<void> {
  await withShapes(["left","top","width","height"], async shapes => {
    const s = shapes[0];
    (document.getElementById("txt-width")  as HTMLInputElement).value = ptToCm(s.width);
    (document.getElementById("txt-height") as HTMLInputElement).value = ptToCm(s.height);
    (document.getElementById("txt-x")      as HTMLInputElement).value = ptToCm(s.left);
    (document.getElementById("txt-y")      as HTMLInputElement).value = ptToCm(s.top);
  });
}

async function applyExactSize(): Promise<void> {
  const w = cmToPt((document.getElementById("txt-width")  as HTMLInputElement).value);
  const h = cmToPt((document.getElementById("txt-height") as HTMLInputElement).value);
  await withShapes(["id"], async shapes => {
    shapes.forEach(s => {
      if (!isNaN(w)) s.width  = w;
      if (!isNaN(h)) s.height = h;
    });
    showStatus("Size applied");
  });
}

async function applyPosition(): Promise<void> {
  const x = cmToPt((document.getElementById("txt-x") as HTMLInputElement).value);
  const y = cmToPt((document.getElementById("txt-y") as HTMLInputElement).value);
  await withShapes(["id"], async shapes => {
    shapes.forEach(s => {
      if (!isNaN(x)) s.left = x;
      if (!isNaN(y)) s.top  = y;
    });
    showStatus("Position applied");
  });
}

// ════════════════════════════════════════════════════════════════
//  FONT
// ════════════════════════════════════════════════════════════════

async function applyFont(): Promise<void> {
  const name = (document.getElementById("cmb-font") as HTMLSelectElement).value;
  const size = parseFloat((document.getElementById("txt-font-size") as HTMLInputElement).value);
  await withShapes(["id"], async shapes => {
    shapes.forEach(s => {
      const f = s.textFrame.textRange.font;
      if (name)     f.name = name;
      if (!isNaN(size)) f.size = size;
    });
    showStatus("Font applied");
  });
}

async function toggleBold(): Promise<void> {
  await PowerPoint.run(async context => {
    const sel = context.presentation.getSelectedShapes();
    sel.load(["id"]);
    await context.sync();
    if (!sel.items.length) { showStatus("No shapes selected", true); return; }
    const firstFont = sel.items[0].textFrame.textRange.font;
    firstFont.load("bold");
    await context.sync();
    const newBold = firstFont.bold !== true;
    sel.items.forEach(s => { s.textFrame.textRange.font.bold = newBold; });
    await context.sync();
  });
}

async function toggleItalic(): Promise<void> {
  await PowerPoint.run(async context => {
    const sel = context.presentation.getSelectedShapes();
    sel.load(["id"]);
    await context.sync();
    if (!sel.items.length) { showStatus("No shapes selected", true); return; }
    const firstFont = sel.items[0].textFrame.textRange.font;
    firstFont.load("italic");
    await context.sync();
    const newItalic = firstFont.italic !== true;
    sel.items.forEach(s => { s.textFrame.textRange.font.italic = newItalic; });
    await context.sync();
  });
}

async function toggleUnderline(): Promise<void> {
  await PowerPoint.run(async context => {
    const sel = context.presentation.getSelectedShapes();
    sel.load(["id"]);
    await context.sync();
    if (!sel.items.length) { showStatus("No shapes selected", true); return; }
    const firstFont = sel.items[0].textFrame.textRange.font;
    firstFont.load("underline");
    await context.sync();
    const isUnderlined = firstFont.underline !== PowerPoint.ShapeFontUnderlineStyle.none;
    sel.items.forEach(s => {
      s.textFrame.textRange.font.underline = isUnderlined
        ? PowerPoint.ShapeFontUnderlineStyle.none
        : PowerPoint.ShapeFontUnderlineStyle.single;
    });
    await context.sync();
  });
}

async function applyFontColor(color: string): Promise<void> {
  await withShapes(["id"], async shapes => {
    shapes.forEach(s => { s.textFrame.textRange.font.color = color; });
    showStatus("Font colour applied");
  });
}

// ════════════════════════════════════════════════════════════════
//  PARAGRAPH
// ════════════════════════════════════════════════════════════════

async function applyParaAlign(align: string): Promise<void> {
  const map: Record<string, PowerPoint.ParagraphHorizontalAlignment> = {
    left:    PowerPoint.ParagraphHorizontalAlignment.left,
    center:  PowerPoint.ParagraphHorizontalAlignment.center,
    right:   PowerPoint.ParagraphHorizontalAlignment.right,
    justify: PowerPoint.ParagraphHorizontalAlignment.justify,
  };
  await withShapes(["id"], async shapes => {
    shapes.forEach(s => {
      s.textFrame.textRange.paragraphFormat.horizontalAlignment = map[align];
    });
    showStatus("Alignment applied");
  });
}

async function applySpacing(): Promise<void> {
  const ls  = parseFloat((document.getElementById("txt-line-spacing")  as HTMLInputElement).value);
  const sbf = parseFloat((document.getElementById("txt-space-before")  as HTMLInputElement).value);
  const saf = parseFloat((document.getElementById("txt-space-after")   as HTMLInputElement).value);
  await withShapes(["id"], async shapes => {
    shapes.forEach(s => {
      const pf = s.textFrame.textRange.paragraphFormat as any;
      if (!isNaN(ls))  pf.lineSpacing  = ls;    // 0 = auto, positive = exact points
      if (!isNaN(sbf)) pf.spaceBefore  = sbf;
      if (!isNaN(saf)) pf.spaceAfter   = saf;
    });
    showStatus("Spacing applied");
  });
}

// ════════════════════════════════════════════════════════════════
//  FILL
// ════════════════════════════════════════════════════════════════

async function applyFillColor(color: string): Promise<void> {
  await withShapes(["id"], async shapes => {
    shapes.forEach(s => s.fill.setSolidColor(color));
    showStatus("Fill applied");
  });
}

async function applyNoFill(): Promise<void> {
  await withShapes(["id"], async shapes => {
    shapes.forEach(s => s.fill.clear());
    showStatus("Fill removed");
  });
}

async function applyTransparency(value: number): Promise<void> {
  await withShapes(["id"], async shapes => {
    shapes.forEach(s => { s.fill.transparency = value / 100; });
  });
}

// ════════════════════════════════════════════════════════════════
//  OUTLINE
// ════════════════════════════════════════════════════════════════

async function applyOutline(): Promise<void> {
  const color  = (document.getElementById("pick-outline-color") as HTMLInputElement).value;
  const weight = parseFloat((document.getElementById("txt-outline-width") as HTMLInputElement).value);
  await withShapes(["id"], async shapes => {
    shapes.forEach(s => {
      s.lineFormat.color   = color;
      if (!isNaN(weight)) s.lineFormat.weight = weight;
      s.lineFormat.visible = true;
    });
    showStatus("Outline applied");
  });
}

async function applyOutlineColor(color: string): Promise<void> {
  await withShapes(["id"], async shapes => {
    shapes.forEach(s => {
      s.lineFormat.color   = color;
      s.lineFormat.visible = true;
    });
    showStatus("Outline colour applied");
  });
}

async function applyNoOutline(): Promise<void> {
  await withShapes(["id"], async shapes => {
    shapes.forEach(s => { s.lineFormat.visible = false; });
    showStatus("Outline removed");
  });
}

// ════════════════════════════════════════════════════════════════
//  SHADOW  (via shape OOXML — requires API 1.5)
// ════════════════════════════════════════════════════════════════

// Outer shadow XML snippets (inserted into <p:spPr>)
const SHADOW_XML: Record<string, string> = {
  soft: `<a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst>`,
  hard: `<a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:outerShdw blurRad="0" dist="12700" dir="2700000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="75000"/></a:srgbClr></a:outerShdw></a:effectLst>`,
  bottom: `<a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:outerShdw blurRad="63500" dist="25400" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr></a:outerShdw></a:effectLst>`,
  perspective: `<a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:outerShdw blurRad="114300" dist="101600" dir="5400000" sy="-30000" algn="b" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr></a:outerShdw></a:effectLst>`,
  none: `<a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>`,
};

async function applyShadow(preset: string): Promise<void> {
  await PowerPoint.run(async context => {
    const sel = context.presentation.getSelectedShapes();
    sel.load(["id"]);
    await context.sync();
    if (!sel.items.length) { showStatus("No shapes selected", true); return; }

    const shadowXml = SHADOW_XML[preset];

    for (const shape of sel.items) {
      const s = shape as any;
      let xml: string = s.getOoxml().value;
      // Replace existing effectLst / effectDag, or inject before </p:spPr>
      if (/<a:effectLst[\s>]/.test(xml)) {
        xml = xml.replace(/<a:effectLst[\s\S]*?<\/a:effectLst>/g, shadowXml);
      } else if (/<a:effectDag[\s>]/.test(xml)) {
        xml = xml.replace(/<a:effectDag[\s\S]*?<\/a:effectDag>/g, shadowXml);
      } else {
        xml = xml.replace("</p:spPr>", shadowXml + "</p:spPr>");
      }
      s.setOoxml(xml);
    }

    await context.sync();
    showStatus(preset === "none" ? "Shadow removed" : "Shadow applied");
  });
}

// ════════════════════════════════════════════════════════════════
//  QUICK ACTIONS
// ════════════════════════════════════════════════════════════════

async function duplicateShapes(): Promise<void> {
  await withShapes(["id"], async shapes => {
    shapes.forEach(s => (s as any).duplicate());
    showStatus("Duplicated");
  });
}

async function savePresentation(): Promise<void> {
  // Office.js has no direct save API for presentations — trigger Ctrl+S via keyboard simulation
  showStatus("Press Ctrl+S (Cmd+S on Mac) to save");
}

// ════════════════════════════════════════════════════════════════
//  EVENT HANDLER WIRING
// ════════════════════════════════════════════════════════════════

function attachHandlers(): void {
  const on = (id: string, fn: () => Promise<void>) => {
    document.getElementById(id)?.addEventListener("click", () => safe(fn));
  };

  // ── Arrange ──
  on("btn-align-left",     () => alignShapes("left"));
  on("btn-align-centerh",  () => alignShapes("centerh"));
  on("btn-align-right",    () => alignShapes("right"));
  on("btn-align-top",      () => alignShapes("top"));
  on("btn-align-middlev",  () => alignShapes("middlev"));
  on("btn-align-bottom",   () => alignShapes("bottom"));
  on("btn-dist-h",         () => distributeShapes(true));
  on("btn-dist-v",         () => distributeShapes(false));
  on("btn-z-fwd",          () => zOrder("fwd"));
  on("btn-z-back",         () => zOrder("back"));
  on("btn-z-front",        () => zOrder("front"));
  on("btn-z-toback",       () => zOrder("toback"));
  on("btn-match-w",        () => matchSize(true,  false));
  on("btn-match-h",        () => matchSize(false, true));
  on("btn-match-both",     () => matchSize(true,  true));
  on("btn-apply-size",     () => applyExactSize());
  on("btn-apply-pos",      () => applyPosition());
  on("btn-read-geometry",  () => readGeometry());

  // ── Font ──
  on("btn-apply-font",       () => applyFont());
  on("btn-bold",             () => toggleBold());
  on("btn-italic",           () => toggleItalic());
  on("btn-underline",        () => toggleUnderline());
  on("btn-apply-font-color", async () => {
    const color = (document.getElementById("pick-font-color") as HTMLInputElement).value;
    await applyFontColor(color);
  });

  // ── Paragraph ──
  on("btn-para-left",     () => applyParaAlign("left"));
  on("btn-para-center",   () => applyParaAlign("center"));
  on("btn-para-right",    () => applyParaAlign("right"));
  on("btn-para-justify",  () => applyParaAlign("justify"));
  on("btn-apply-spacing", () => applySpacing());

  // ── Fill ──
  on("btn-apply-fill", async () => {
    const color = (document.getElementById("pick-fill-color") as HTMLInputElement).value;
    await applyFillColor(color);
  });
  on("btn-no-fill",       () => applyNoFill());
  on("btn-apply-outline", () => applyOutline());
  on("btn-no-outline",    () => applyNoOutline());

  document.getElementById("btn-edit-swatches")?.addEventListener("click", () => toggleEditMode());

  document.getElementById("trk-transparency")?.addEventListener("input", e => {
    const v = parseInt((e.target as HTMLInputElement).value);
    document.getElementById("lbl-transparency")!.textContent = v + "%";
    safe(() => applyTransparency(v));
  });

  // ── Shadow ──
  on("btn-shadow-soft",        () => applyShadow("soft"));
  on("btn-shadow-hard",        () => applyShadow("hard"));
  on("btn-shadow-bottom",      () => applyShadow("bottom"));
  on("btn-shadow-perspective", () => applyShadow("perspective"));
  on("btn-shadow-remove",      () => applyShadow("none"));

  // ── Quick ──
  on("btn-duplicate", () => duplicateShapes());
  on("btn-save",      () => savePresentation());
}
