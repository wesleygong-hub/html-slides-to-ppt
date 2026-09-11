// Browser rectangles include ancestor transforms; computed CSS lengths do not.
// Keep every exported length in the same (rendered CSS pixel) coordinate space.
export async function installCaptureHelpers(page) {
  await page.evaluate(() => {
    globalThis.__html2pptGeometry = (element) => {
      let matrix = new DOMMatrixReadOnly();
      for (let current = element; current; current = current.parentElement) {
        const style = getComputedStyle(current);
        const transform = new DOMMatrixReadOnly(style.transform === "none" ? undefined : style.transform);
        matrix = transform.multiply(matrix);
        const zoom = Number.parseFloat(style.zoom) || 1;
        matrix = new DOMMatrixReadOnly().scale(zoom).multiply(matrix);
      }
      return {
        scaleX: Math.hypot(matrix.a, matrix.b),
        scaleY: Math.hypot(matrix.c, matrix.d),
        rotation: Math.atan2(matrix.b, matrix.a) * 180 / Math.PI,
      };
    };
  });
}

export async function prepareCapturePage(page, selector) {
  return page.evaluate((selector) => {
    const slides = [...document.querySelectorAll(selector)];
    const insideSlide = (node) => slides.some((slide) => slide.contains(node));
    const containsSlide = (node) => slides.some((slide) => node.contains(slide));
    // Isolate presentation content from sibling controls without disturbing
    // flex/grid layout or fixed, pointer-transparent decorative overlays.
    let hidden = 0;
    for (const element of document.querySelectorAll("body *")) {
      if (element.hasAttribute("data-html2ppt-ignore")) {
        element.setAttribute("data-html2ppt-excluded", "1");
        hidden++;
        continue;
      }
      if (insideSlide(element) || containsSlide(element)) continue;
      const style = getComputedStyle(element);
      const isControl = element.matches("button,nav,input,select,textarea,[role=button],[role=navigation]");
      const isOverlay = ["fixed", "absolute"].includes(style.position) && style.pointerEvents !== "none";
      if (isControl || isOverlay) {
        element.setAttribute("data-html2ppt-excluded", "1");
        hidden++;
      }
    }
    const style = document.createElement("style");
    style.textContent = '[data-html2ppt-excluded="1"], [data-html2ppt-excluded="1"] * { visibility: hidden !important; }';
    document.head.append(style);
    return hidden;
  }, selector);
}

export function resolveFontFace(value, installedFonts) {
  const names = String(value || "Microsoft YaHei").match(/"[^"]*"|'[^']*'|[^,]+/g)
    .map((name) => name.trim().replace(/^['"]|['"]$/g, ""));
  const available = new Map((installedFonts || []).map((name) => [name.toLowerCase(), name]));
  for (const name of names) {
    if (available.has(name.toLowerCase())) return available.get(name.toLowerCase());
    if (/^(serif|sans-serif|system-ui|monospace)$/i.test(name)) break;
  }
  const first = names[0];
  // If platform font discovery is unavailable, preserve explicitly named
  // families rather than matching substrings such as "Serif" in Noto Serif SC.
  if (!installedFonts && !/^(serif|sans-serif|system-ui|monospace)$/i.test(first)) return first;
  const serif = names.some((name) => /serif|song|宋|明朝/i.test(name) && !/sans-serif/i.test(name));
  const mono = names.some((name) => /mono|consolas/i.test(name));
  const fallback = mono ? ["Consolas", "Courier New"] : serif
    ? ["Noto Serif SC", "SimSun", "Times New Roman"]
    : ["Noto Sans SC", "Microsoft YaHei", "Arial"];
  return fallback.find((name) => available.has(name.toLowerCase())) || fallback[0];
}

// Merge only adjacent inline fragments belonging to the same formatting
// context and animation. Flex items and separately positioned labels stay put.
export function mergeInlineTextBlocks(blocks) {
  const output = [];
  for (const block of blocks) {
    const prior = output.at(-1);
    const gap = prior ? block.x - (prior.x + prior.w) : Infinity;
    const sameLine = prior && Math.abs((prior.y + prior.h) - (block.y + block.h)) < Math.max(2, block.fontSize * 0.2);
    const compatible = prior && prior.flowId === block.flowId && prior.flowId
      && JSON.stringify(prior.animationRefs) === JSON.stringify(block.animationRefs)
      && Math.abs(prior.fontSize - block.fontSize) < 0.1
      && Math.abs(prior.rotation || 0) < 0.1 && Math.abs(block.rotation || 0) < 0.1;
    if (compatible && sameLine && gap >= -2 && gap < block.fontSize * 0.8) {
      prior.runs ||= [{ ...prior, runs: undefined }];
      const separator = gap > block.fontSize * 0.15 && !/\s$/.test(prior.runs.at(-1).text)
        && !/^\s/.test(block.text) ? " " : "";
      prior.runs.push({ ...block, text: separator + block.text });
      prior.text += separator + block.text;
      const right = block.x + block.w;
      const bottom = Math.max(prior.y + prior.h, block.y + block.h);
      prior.y = Math.min(prior.y, block.y);
      prior.h = bottom - prior.y;
      prior.w = right - prior.x;
    } else output.push({ ...block });
  }
  return output;
}
