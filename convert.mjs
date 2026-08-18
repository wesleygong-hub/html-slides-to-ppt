#!/usr/bin/env node

import fs from "node:fs/promises";
import { createHash } from "node:crypto";
import { execFile } from "node:child_process";
import path from "node:path";
import process from "node:process";
import { promisify } from "node:util";
import { fileURLToPath, pathToFileURL } from "node:url";
import { chromium } from "playwright";
import pptxgen from "pptxgenjs";
import {
  applyAnimationSemantics,
  mapHtmlEffectToPowerPoint,
  summarizeAnimationDirections,
} from "./animation-semantics.mjs";

const execFileAsync = promisify(execFile);

const HELP = `
HTML Slides to PPT

用法：
  node convert.mjs --input deck.html [--output deck.pptx]

选项：
  --input, -i <file>       输入 HTML 文件（必填）
  --output, -o <file>      输出 PPTX；默认与 HTML 同目录并带模式后缀
  --mode <mode>            objects | editable | image，默认 objects
  --selector <css>         自定义每页幻灯片的 CSS 选择器
  --browser <mode>         auto | msedge | chrome | chromium，默认 auto
  --width <px>             浏览器视口宽度，默认 1920
  --height <px>            浏览器视口高度，默认 1080
  --scale <n>              截图像素倍率，默认 2（4K）；小文件可设为 1
  --wait <ms>              切页后额外等待，默认 400
  --fit <mode>             contain | cover | stretch，默认 contain
  --page-base <mode>       layout | slide；objects 模式默认 layout
  --animations <mode>      auto | off | required，默认 auto
  --keep-images [dir]      同时保存逐页 PNG；不传目录时用 <输出名>-images
  --help, -h               显示帮助

支持自动识别：section.slide、Reveal.js、[data-slide]、.slide 和普通 section。
`;

function parseArgs(argv) {
  const opts = {
    input: "",
    output: "",
    selector: "",
    browser: "auto",
    mode: "objects",
    width: 1920,
    height: 1080,
    scale: 2,
    wait: 400,
    fit: "contain",
    pageBase: "layout",
    animations: "auto",
    keepImages: false,
    keepImagesDir: "",
  };

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    const next = () => {
      if (i + 1 >= argv.length) throw new Error(`${arg} 缺少参数`);
      i += 1;
      return argv[i];
    };

    if (arg === "--help" || arg === "-h") opts.help = true;
    else if (arg === "--input" || arg === "-i") opts.input = next();
    else if (arg === "--output" || arg === "-o") opts.output = next();
    else if (arg === "--selector") opts.selector = next();
    else if (arg === "--browser") opts.browser = next().toLowerCase();
    else if (arg === "--mode") opts.mode = next().toLowerCase();
    else if (arg === "--width") opts.width = Number(next());
    else if (arg === "--height") opts.height = Number(next());
    else if (arg === "--scale") opts.scale = Number(next());
    else if (arg === "--wait") opts.wait = Number(next());
    else if (arg === "--fit") opts.fit = next().toLowerCase();
    else if (arg === "--page-base") opts.pageBase = next().toLowerCase();
    else if (arg === "--animations") opts.animations = next().toLowerCase();
    else if (arg === "--keep-images") {
      opts.keepImages = true;
      if (i + 1 < argv.length && !argv[i + 1].startsWith("-")) {
        i += 1;
        opts.keepImagesDir = argv[i];
      }
    } else if (!arg.startsWith("-") && !opts.input) opts.input = arg;
    else throw new Error(`未知参数：${arg}`);
  }
  return opts;
}

function validateOptions(opts) {
  if (!opts.input) throw new Error("请通过 --input 指定 HTML 文件");
  if (!Number.isFinite(opts.width) || opts.width < 320) throw new Error("--width 必须大于等于 320");
  if (!Number.isFinite(opts.height) || opts.height < 240) throw new Error("--height 必须大于等于 240");
  if (!Number.isFinite(opts.scale) || opts.scale < 0.5 || opts.scale > 4) {
    throw new Error("--scale 必须在 0.5 到 4 之间");
  }
  if (!Number.isFinite(opts.wait) || opts.wait < 0 || opts.wait > 30000) {
    throw new Error("--wait 必须在 0 到 30000 之间");
  }
  if (!["auto", "msedge", "chrome", "chromium"].includes(opts.browser)) {
    throw new Error("--browser 只支持 auto、msedge、chrome、chromium");
  }
  if (!["objects", "editable", "image"].includes(opts.mode)) {
    throw new Error("--mode 只支持 objects、editable 或 image");
  }
  if (!["contain", "cover", "stretch"].includes(opts.fit)) {
    throw new Error("--fit 只支持 contain、cover、stretch");
  }
  if (!["layout", "slide"].includes(opts.pageBase)) {
    throw new Error("--page-base 只支持 layout 或 slide");
  }
  if (!["auto", "off", "required"].includes(opts.animations)) {
    throw new Error("--animations 只支持 auto、off 或 required");
  }
  if (opts.mode === "image" && opts.animations === "required") {
    throw new Error("image 模式只有整页静态图片，不能使用 --animations required");
  }
}

async function launchBrowser(mode) {
  const attempts = [];
  const launch = async (label, options) => {
    try {
      const browser = await chromium.launch({
        headless: true,
        ...options,
        args: [
          "--disable-background-networking",
          "--disable-component-update",
          "--disable-default-apps",
          "--disable-dev-shm-usage",
          "--disable-features=Translate,OptimizationHints,MediaRouter",
          "--disable-gpu",
          "--no-default-browser-check",
          "--no-first-run",
        ],
      });
      return { browser, label };
    } catch (error) {
      attempts.push(`${label}: ${error.message.split("\n")[0]}`);
      return null;
    }
  };

  if (mode === "msedge") {
    const result = await launch("Microsoft Edge", { channel: "msedge" });
    if (result) return result;
  } else if (mode === "chrome") {
    const result = await launch("Google Chrome", { channel: "chrome" });
    if (result) return result;
  } else if (mode === "chromium") {
    const result = await launch("Playwright Chromium", {});
    if (result) return result;
  } else {
    for (const [label, options] of [
      ["Microsoft Edge", { channel: "msedge" }],
      ["Google Chrome", { channel: "chrome" }],
      ["Playwright Chromium", {}],
    ]) {
      const result = await launch(label, options);
      if (result) return result;
    }
  }

  throw new Error(
    `无法启动浏览器。请确认 Edge/Chrome 已安装，或运行 npm run install-browser。\n${attempts.join("\n")}`,
  );
}

async function waitForAssets(page) {
  await page.evaluate(async () => {
    const timeout = new Promise((resolve) => setTimeout(resolve, 12000));
    const fonts = document.fonts?.ready ?? Promise.resolve();
    const images = Promise.all(
      Array.from(document.images).map((image) => {
        if (image.complete) return Promise.resolve();
        return new Promise((resolve) => {
          image.addEventListener("load", resolve, { once: true });
          image.addEventListener("error", resolve, { once: true });
        });
      }),
    );
    await Promise.race([Promise.all([fonts, images]), timeout]);
  });
}

async function detectSlides(page, requestedSelector) {
  if (requestedSelector) {
    const count = await page.locator(requestedSelector).count();
    if (!count) throw new Error(`选择器没有匹配到幻灯片：${requestedSelector}`);
    return { selector: requestedSelector, count, framework: "custom" };
  }

  const reveal = await page.evaluate(() => {
    if (!globalThis.Reveal || typeof globalThis.Reveal.getSlides !== "function") return 0;
    const slides = globalThis.Reveal.getSlides();
    slides.forEach((slide, index) => slide.setAttribute("data-html2ppt-reveal-index", String(index)));
    return slides.length;
  });
  if (reveal) {
    return {
      selector: "[data-html2ppt-reveal-index]",
      count: reveal,
      framework: "reveal",
    };
  }

  const candidates = [
    "section.slide",
    "[data-slide]",
    ".slides > .slide",
    ".slide",
    "section",
  ];
  for (const selector of candidates) {
    const count = await page.locator(selector).count();
    if (count) return { selector, count, framework: "generic" };
  }
  throw new Error("没有找到幻灯片。可通过 --selector 指定每一页的 CSS 选择器。");
}

async function activateSlide(page, info, index) {
  await page.evaluate(
    ({ selector, framework, index: targetIndex }) => {
      const slides = Array.from(document.querySelectorAll(selector));
      const target = slides[targetIndex];
      if (!target) throw new Error(`幻灯片索引不存在：${targetIndex}`);

      if (framework === "reveal" && globalThis.Reveal) {
        const indices = globalThis.Reveal.getIndices(target);
        globalThis.Reveal.slide(indices.h, indices.v, indices.f);
        return;
      }

      if (globalThis.deck && typeof globalThis.deck.showSlide === "function") {
        globalThis.deck.showSlide(targetIndex);
        return;
      }
      if (globalThis.deck && typeof globalThis.deck.show === "function") {
        globalThis.deck.show(targetIndex);
        return;
      }
      if (typeof globalThis.showSlide === "function") {
        globalThis.showSlide(targetIndex);
        return;
      }

      slides.forEach((slide, slideIndex) => {
        slide.style.removeProperty("display");
        slide.style.removeProperty("visibility");
        slide.style.removeProperty("opacity");
        slide.classList.toggle("active", slideIndex === targetIndex);
        slide.classList.toggle("present", slideIndex === targetIndex);
        slide.classList.toggle("visible", slideIndex === targetIndex);
        slide.setAttribute("aria-hidden", slideIndex === targetIndex ? "false" : "true");
        if (slideIndex !== targetIndex) {
          slide.style.setProperty("display", "none", "important");
        }
      });

      const style = getComputedStyle(target);
      if (style.display === "none") target.style.setProperty("display", "block", "important");
      target.style.setProperty("visibility", "visible", "important");
      target.style.setProperty("opacity", "1", "important");
    },
    { selector: info.selector, framework: info.framework, index },
  );
}

async function extractSlideAnimations(page, info, index, animationMode) {
  const captured = await page.evaluate(
    async ({ selector, index: targetIndex }) => {
      const slide = document.querySelectorAll(selector)[targetIndex];
      if (!slide) return [];

      slide.querySelectorAll("[data-html2ppt-animation-id]").forEach((element) => {
        element.removeAttribute("data-html2ppt-animation-id");
      });

      // Give CSS transitions and Web Animations one rendering opportunity after
      // the slide controller adds its active/visible class.
      await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));

      const elements = Array.from(slide.querySelectorAll("*"));
      const slideRect = slide.getBoundingClientRect();
      const nodes = [slide, ...elements];
      const nodeIds = new WeakMap(nodes.map((element, nodeIndex) => (
        [element, `html2ppt-node-${targetIndex + 1}-${nodeIndex + 1}`]
      )));
      const finalSnapshots = new Map();
      const milliseconds = (value) => {
        const text = String(value || "0s").trim().toLowerCase();
        const parsed = Number.parseFloat(text) || 0;
        return text.endsWith("ms") ? parsed : parsed * 1000;
      };
      const list = (value) => String(value || "").split(",").map((item) => item.trim());
      const at = (values, itemIndex, fallback = "") => (
        values.length ? values[itemIndex % values.length] : fallback
      );
      const finiteNumber = (value, fallback = 0) => {
        const number = Number(value);
        return Number.isFinite(number) ? number : fallback;
      };
      const styleSnapshot = (style) => ({
        opacity: style.opacity,
        transform: style.transform,
        clipPath: style.clipPath,
        strokeDashoffset: style.strokeDashoffset,
        width: style.width,
        height: style.height,
      });
      const elementLabel = (element) => {
        const tag = element.tagName?.toLowerCase?.() || "element";
        const id = element.id ? `#${element.id}` : "";
        const classes = Array.from(element.classList || []).slice(0, 3).map((name) => `.${name}`).join("");
        return `${tag}${id}${classes}`;
      };
      const revealHint = (element) => {
        let current = element;
        while (current && current !== slide.parentElement) {
          if (current.hasAttribute?.("data-html2ppt-reveal-from")) {
            const raw = current.getAttribute("data-html2ppt-reveal-from") ?? "";
            return { raw, value: raw.trim().toLowerCase(), owner: elementLabel(current) };
          }
          current = current.parentElement;
        }
        return null;
      };

      for (const element of elements) {
        const style = getComputedStyle(element);
        const animations = [];
        for (const animation of element.getAnimations({ subtree: false })) {
          if (!animation.animationName) continue; // CSS transitions are reconstructed below.
          const timing = animation.effect?.getTiming?.() || {};
          const duration = finiteNumber(timing.duration, 0);
          const iterations = finiteNumber(timing.iterations, 1);
          if (duration <= 0 || iterations <= 0 || !Number.isFinite(iterations)) continue;
          animations.push({
            type: "animation",
            name: animation.animationName || "",
            duration,
            delay: finiteNumber(timing.delay, 0),
            iterations,
            easing: timing.easing || "linear",
            keyframes: (animation.effect?.getKeyframes?.() || []).map((frame) => ({
              offset: finiteNumber(frame.offset, 0),
              opacity: frame.opacity,
              transform: frame.transform,
              clipPath: frame.clipPath,
              strokeDashoffset: frame.strokeDashoffset,
            })),
          });
        }

        // A completed CSS animation with `fill:none` may no longer have a
        // Web Animations object. Keep its declared timing so name-based mapping
        // can still produce a sensible native PowerPoint effect.
        const names = list(style.animationName).filter((name) => name && name !== "none");
        const durations = list(style.animationDuration).map(milliseconds);
        const delays = list(style.animationDelay).map(milliseconds);
        const iterationCounts = list(style.animationIterationCount).map((value) => (
          value === "infinite" ? Number.POSITIVE_INFINITY : Number.parseFloat(value) || 1
        ));
        names.forEach((name, nameIndex) => {
          if (animations.some((item) => item.name === name)) return;
          const duration = at(durations, nameIndex, 0);
          const iterations = at(iterationCounts, nameIndex, 1);
          if (duration <= 0 || !Number.isFinite(iterations)) return;
          animations.push({
            type: "animation",
            name,
            duration,
            delay: at(delays, nameIndex, 0),
            iterations,
            easing: at(list(style.animationTimingFunction), nameIndex, "linear"),
            keyframes: [],
          });
        });

        const transitionDurations = list(style.transitionDuration).map(milliseconds);
        const hasTransition = transitionDurations.some((duration) => duration > 0);
        if (hasTransition || animations.length) {
          finalSnapshots.set(element, {
            style: styleSnapshot(style),
            transitionProperties: list(style.transitionProperty),
            transitionDurations,
            transitionDelays: list(style.transitionDelay).map(milliseconds),
            transitionEasings: list(style.transitionTimingFunction),
            animations,
          });
        }
      }

      if (!finalSnapshots.size) return [];

      // Reconstruct transition start values by briefly removing the standard
      // slide-state classes with all motion disabled. This handles decks whose
      // first slide finished animating before conversion began.
      const originalClass = slide.getAttribute("class");
      const inspectStyle = document.createElement("style");
      inspectStyle.setAttribute("data-html2ppt-animation-inspection", "1");
      inspectStyle.textContent = `
        html[data-html2ppt-animation-inspection="1"] * {
          animation: none !important;
          transition: none !important;
        }
      `;
      document.head.appendChild(inspectStyle);
      document.documentElement.setAttribute("data-html2ppt-animation-inspection", "1");
      slide.classList.remove("active", "present", "visible", "current", "shown");
      void slide.offsetWidth;

      const initialSnapshots = new Map();
      for (const element of finalSnapshots.keys()) {
        initialSnapshots.set(element, styleSnapshot(getComputedStyle(element)));
      }

      if (originalClass === null) slide.removeAttribute("class");
      else slide.setAttribute("class", originalClass);
      void slide.offsetWidth;
      document.documentElement.removeAttribute("data-html2ppt-animation-inspection");
      inspectStyle.remove();

      const changed = (before, after, property) => {
        const a = String(before?.[property] ?? "");
        const b = String(after?.[property] ?? "");
        if (property === "opacity") return Math.abs((Number.parseFloat(a) || 0) - (Number.parseFloat(b) || 0)) > 0.01;
        return a !== b;
      };
      const results = [];
      let sequence = 0;
      for (const [element, final] of finalSnapshots) {
        const initial = initialSnapshots.get(element) || {};
        const effects = [...final.animations];
        const changedProperties = final.transitionProperties.filter((property) => {
          if (property === "all") {
            return ["opacity", "transform", "clipPath", "width", "height"]
              .some((candidate) => changed(initial, final.style, candidate));
          }
          const normalized = property === "clip-path" ? "clipPath" : property;
          return Object.hasOwn(final.style, normalized) && changed(initial, final.style, normalized);
        });
        // When a transition is already at its end state, Chromium can report
        // opacity as changed while a paired transform has been optimized away.
        // Preserve the declared transform so fade+move entrances map to Float.
        if (changedProperties.length
          && final.transitionProperties.includes("transform")
          && !changedProperties.includes("transform")) {
          changedProperties.push("transform");
        }
        if (changedProperties.length) {
          const duration = Math.max(...final.transitionDurations);
          if (duration > 0) {
            effects.push({
              type: "transition",
              name: changedProperties.join(","),
              properties: changedProperties,
              duration,
              delay: final.transitionDelays.length ? Math.max(...final.transitionDelays) : 0,
              easing: final.transitionEasings[0] || "ease",
              keyframes: [initial, final.style],
            });
          }
        }
        if (!effects.length) continue;
        const id = `html2ppt-animation-${targetIndex + 1}-${++sequence}`;
        element.setAttribute("data-html2ppt-animation-id", id);
        const finalAnimatedTransform = effects
          .flatMap((effect) => effect.keyframes || [])
          .map((frame) => frame?.transform)
          .filter((value) => value !== undefined && value !== null && value !== "")
          .at(-1);
        const previousTransform = element.style.getPropertyValue("transform");
        const previousTransformPriority = element.style.getPropertyPriority("transform");
        if (finalAnimatedTransform !== undefined) {
          element.style.setProperty("transform", finalAnimatedTransform || "none", "important");
        }
        const rect = element.getBoundingClientRect();
        if (finalAnimatedTransform !== undefined) {
          if (previousTransform) {
            element.style.setProperty("transform", previousTransform, previousTransformPriority);
          } else {
            element.style.removeProperty("transform");
          }
        }
        const ancestorChain = [];
        let ancestor = element.parentElement;
        while (ancestor && slide.contains(ancestor)) {
          ancestorChain.push(nodeIds.get(ancestor));
          if (ancestor === slide) break;
          ancestor = ancestor.parentElement;
        }
        results.push({
          id,
          effects,
          metadata: {
            rect: {
              x: rect.left - slideRect.left,
              y: rect.top - slideRect.top,
              w: rect.width,
              h: rect.height,
            },
            slideSize: { width: slideRect.width, height: slideRect.height },
            transformOrigin: getComputedStyle(element).transformOrigin,
            ancestorChain,
            revealHint: revealHint(element),
            elementLabel: elementLabel(element),
          },
        });
      }
      return results;
    },
    { selector: info.selector, index },
  );
  const semanticResult = applyAnimationSemantics(captured, {
    slideIndex: index + 1,
    mode: animationMode,
  });
  for (const warning of semanticResult.warnings) console.warn(`动画提示：${warning}`);
  if (semanticResult.errors.length) throw new Error(semanticResult.errors.join("；"));
  return semanticResult.animations;
}

async function finishSlideMotion(page, info, index, waitMs) {
  await page.waitForTimeout(waitMs);
  await page.evaluate(
    ({ selector, index: targetIndex }) => {
      const target = document.querySelectorAll(selector)[targetIndex];
      if (!target) return;
      for (const animation of target.getAnimations({ subtree: true })) {
        try {
          if (Number.isFinite(animation.effect?.getTiming?.().duration)) animation.finish();
        } catch {
          // Infinite decorative animations are intentionally left at their current frame.
        }
      }
    },
    { selector: info.selector, index },
  );
  await page.waitForTimeout(60);
}

async function installCleanCaptureStyle(page) {
  await page.addStyleTag({
    content: `
      html[data-html2ppt-clean="1"] [data-html2ppt-hide-text="1"] {
        color: transparent !important;
        text-decoration-color: transparent !important;
        text-shadow: none !important;
        -webkit-text-stroke-color: transparent !important;
      }
      html[data-html2ppt-clean="1"] [data-html2ppt-hide-text="1"]::before,
      html[data-html2ppt-clean="1"] [data-html2ppt-hide-text="1"]::after {
        color: transparent !important;
        text-decoration-color: transparent !important;
        text-shadow: none !important;
        -webkit-text-stroke-color: transparent !important;
      }
      [data-html2ppt-pseudo-before="1"]::before,
      [data-html2ppt-pseudo-after="1"]::after {
        content: none !important;
        display: none !important;
      }
      html[data-html2ppt-objects-clean="1"] [data-html2ppt-hide-box="1"] {
        background: transparent !important;
        background-color: transparent !important;
        background-image: none !important;
        border-color: transparent !important;
        box-shadow: none !important;
        outline-color: transparent !important;
      }
      html[data-html2ppt-objects-clean="1"] [data-html2ppt-hide-object="1"] {
        visibility: hidden !important;
      }
      html[data-html2ppt-object-capture] {
        background: transparent !important;
      }
      html[data-html2ppt-object-capture] body {
        background: transparent !important;
      }
      html[data-html2ppt-object-capture] body * {
        visibility: hidden !important;
      }
      html[data-html2ppt-object-capture="asset"] [data-html2ppt-capture-target="1"],
      html[data-html2ppt-object-capture="asset"] [data-html2ppt-capture-target="1"] * {
        visibility: visible !important;
      }
      html[data-html2ppt-object-capture="background"] [data-html2ppt-capture-target="1"] {
        visibility: visible !important;
      }
      html[data-html2ppt-object-capture="background"] [data-html2ppt-capture-target="1"] > * {
        visibility: hidden !important;
      }
    `,
  });
}

async function extractEditableText(page, info, index) {
  return page.evaluate(
    ({ selector, index: targetIndex }) => {
      const slide = document.querySelectorAll(selector)[targetIndex];
      if (!slide) return { slideRect: null, textBlocks: [] };
      const slideRect = slide.getBoundingClientRect();
      const walker = document.createTreeWalker(slide, NodeFilter.SHOW_TEXT);
      const blocks = [];
      const marked = new Set();
      let sequence = 0;

      const animationRefs = (element) => {
        const refs = [];
        let current = element;
        while (current && current !== slide.parentElement) {
          const id = current.getAttribute?.("data-html2ppt-animation-id");
          if (id) refs.unshift(id);
          current = current.parentElement;
        }
        return refs;
      };

      const rgb = (value) => {
        const match = String(value || "").match(/rgba?\(\s*([\d.]+)[, ]+\s*([\d.]+)[, ]+\s*([\d.]+)/i);
        if (!match) return "000000";
        return [match[1], match[2], match[3]]
          .map((part) => Math.max(0, Math.min(255, Math.round(Number(part)))).toString(16).padStart(2, "0"))
          .join("")
          .toUpperCase();
      };

      const effectiveOpacity = (element) => {
        let opacity = 1;
        let current = element;
        while (current && current !== slide.parentElement) {
          const style = getComputedStyle(current);
          if (style.display === "none" || style.visibility === "hidden") return 0;
          opacity *= Number.parseFloat(style.opacity || "1") || 0;
          current = current.parentElement;
        }
        return opacity;
      };

      const transformedText = (text, transform) => {
        if (transform === "uppercase") return text.toUpperCase();
        if (transform === "lowercase") return text.toLowerCase();
        if (transform === "capitalize") return text.replace(/(^|\s)\S/g, (part) => part.toUpperCase());
        return text;
      };

      let node;
      while ((node = walker.nextNode())) {
        const original = node.textContent || "";
        if (!original.trim()) continue;
        const element = node.parentElement;
        if (!element || element.closest("script,style,noscript,svg,canvas,textarea")) continue;
        const style = getComputedStyle(element);
        if (effectiveOpacity(element) < 0.01) continue;
        if (style.backgroundClip === "text" || style.webkitBackgroundClip === "text") continue;

        const fontSize = Number.parseFloat(style.fontSize) || 16;
        const range = document.createRange();
        const chars = [];
        for (let charIndex = 0; charIndex < original.length; charIndex += 1) {
          range.setStart(node, charIndex);
          range.setEnd(node, charIndex + 1);
          const rects = Array.from(range.getClientRects());
          if (!rects.length) continue;
          const rect = rects[0];
          if (rect.width <= 0 || rect.height <= 0) continue;
          chars.push({ char: original[charIndex].replace(/\u00a0/g, " "), rect });
        }
        if (!chars.length) continue;

        const lines = [];
        for (const item of chars) {
          const previous = lines[lines.length - 1];
          const tolerance = Math.max(2, fontSize * 0.18);
          if (!previous || Math.abs(previous.top - item.rect.top) > tolerance) {
            lines.push({ items: [item], top: item.rect.top });
          } else {
            previous.items.push(item);
          }
        }

        for (const line of lines) {
          const rawText = line.items.map((item) => item.char).join("").replace(/\s+/g, " ").trim();
          if (!rawText) continue;
          const visibleItems = line.items.filter((item) => item.char.trim());
          if (!visibleItems.length) continue;
          const left = Math.min(...visibleItems.map((item) => item.rect.left));
          const top = Math.min(...visibleItems.map((item) => item.rect.top));
          const right = Math.max(...visibleItems.map((item) => item.rect.right));
          const bottom = Math.max(...visibleItems.map((item) => item.rect.bottom));
          const matrix = new DOMMatrixReadOnly(style.transform === "none" ? undefined : style.transform);
          const rotation = Math.abs(matrix.b) > 0.0001 || Math.abs(matrix.a - 1) > 0.0001
            ? Math.atan2(matrix.b, matrix.a) * 180 / Math.PI
            : 0;
          blocks.push({
            id: `html2ppt-text-${targetIndex + 1}-${++sequence}`,
            text: transformedText(rawText, style.textTransform),
            x: left - slideRect.left,
            y: top - slideRect.top,
            w: right - left,
            h: bottom - top,
            fontSize,
            fontFamily: style.fontFamily || "Microsoft YaHei",
            fontWeight: style.fontWeight || "400",
            fontStyle: style.fontStyle || "normal",
            color: rgb(style.color),
            align: style.textAlign || "left",
            letterSpacing: Number.parseFloat(style.letterSpacing) || 0,
            rotation,
            animationRefs: animationRefs(element),
          });
        }

        marked.add(element);
      }

      for (const element of marked) element.setAttribute("data-html2ppt-hide-text", "1");
      return {
        slideRect: { width: slideRect.width, height: slideRect.height },
        textBlocks: blocks,
      };
    },
    { selector: info.selector, index },
  );
}

async function extractVisualObjects(page, info, index) {
  return page.evaluate(
    ({ selector, index: targetIndex }) => {
      const slide = document.querySelectorAll(selector)[targetIndex];
      if (!slide) return [];

      // Pseudo-elements inside clipped media containers (for example a photo
      // overlay inside a masked image arc) must stay grouped with that asset.
      // Chromium applies masks to true pseudo-elements more reliably than to a
      // materialized child during element screenshots, so discover these roots
      // before materializing the remaining pseudo-elements.
      const assetRoots = new Set();
      const mediaElements = Array.from(slide.querySelectorAll("img,svg,canvas,video"));

      const isRendered = (element) => {
        const style = getComputedStyle(element);
        if (style.display === "none" || style.visibility === "hidden"
          || Number.parseFloat(style.opacity || "1") < 0.01) return false;
        const rect = element.getBoundingClientRect();
        return rect.width >= 1.5 && rect.height >= 1.5;
      };

      const hasVisibleTextOutsideMedia = (candidate, asset) => {
        const walker = document.createTreeWalker(candidate, NodeFilter.SHOW_TEXT);
        let node;
        while ((node = walker.nextNode())) {
          if (!(node.textContent || "").trim()) continue;
          const parent = node.parentElement;
          if (!parent || asset.contains(parent)) continue;
          if (isRendered(parent)) return true;
        }
        return false;
      };

      const hasIndependentVisualOutsideMedia = (candidate, asset) => {
        for (const element of candidate.querySelectorAll("*")) {
          // Ignore the media itself, its SVG descendants, and layout wrappers
          // on the direct ancestry path from the media to the clip candidate.
          if (element === asset || asset.contains(element) || element.contains(asset)) continue;
          if (!isRendered(element)) continue;
          const style = getComputedStyle(element);
          const background = String(style.backgroundColor || "");
          const backgroundAlpha = Number(background.match(/rgba?\([^)]*[,/]\s*([\d.]+)\s*\)$/i)?.[1]
            ?? (background.startsWith("rgb(") ? 1 : 0));
          const hasFill = backgroundAlpha > 0.01;
          const hasBackgroundImage = style.backgroundImage && style.backgroundImage !== "none";
          const hasBorder = ["Top", "Right", "Bottom", "Left"].some((side) => {
            const width = Number.parseFloat(style[`border${side}Width`]) || 0;
            return width > 0.1 && style[`border${side}Style`] !== "none";
          });
          const hasEffect = (style.boxShadow && style.boxShadow !== "none")
            || (style.filter && style.filter !== "none");
          if (hasFill || hasBackgroundImage || hasBorder || hasEffect) return true;
        }
        return false;
      };

      const isDedicatedMediaWrapper = (candidate, asset) => {
        // A clip wrapper may be promoted only when it represents one coherent
        // visual asset. A content panel that happens to use overflow:hidden
        // must not swallow its cards, labels, or diagrams into one screenshot.
        const containedMedia = mediaElements.filter((item) => candidate.contains(item));
        if (containedMedia.length !== 1 || containedMedia[0] !== asset) return false;
        if (hasVisibleTextOutsideMedia(candidate, asset)) return false;
        if (hasIndependentVisualOutsideMedia(candidate, asset)) return false;
        return true;
      };

      for (const asset of mediaElements) {
        let root = asset;
        let current = asset.parentElement;
        while (current && current !== slide) {
          const currentStyle = getComputedStyle(current);
          const clips = [currentStyle.overflow, currentStyle.overflowX, currentStyle.overflowY].includes("hidden")
            || (currentStyle.clipPath && currentStyle.clipPath !== "none")
            || (currentStyle.maskImage && currentStyle.maskImage !== "none")
            || (currentStyle.webkitMaskImage && currentStyle.webkitMaskImage !== "none");
          if (clips) {
            if (isDedicatedMediaWrapper(current, asset)) root = current;
            break;
          }
          current = current.parentElement;
        }
        assetRoots.add(root);
      }

      // CSS pseudo-elements do not exist in the DOM, so they cannot normally
      // become separate PowerPoint objects. Replace every visible ::before and
      // ::after with an equivalent real element while preserving layout. The
      // replacement then flows through the same shape/image extraction path as
      // ordinary HTML elements.
      slide.querySelectorAll("[data-html2ppt-pseudo-clone]").forEach((element) => element.remove());
      slide.querySelectorAll("[data-html2ppt-pseudo-before]").forEach((element) => element.removeAttribute("data-html2ppt-pseudo-before"));
      slide.querySelectorAll("[data-html2ppt-pseudo-after]").forEach((element) => element.removeAttribute("data-html2ppt-pseudo-after"));
      const pseudoSources = Array.from(slide.querySelectorAll("*"));
      for (const source of pseudoSources) {
        if (Array.from(assetRoots).some((root) => root === source || root.contains(source))) continue;
        for (const pseudo of ["before", "after"]) {
          const pseudoStyle = getComputedStyle(source, `::${pseudo}`);
          const content = String(pseudoStyle.content || "").trim().toLowerCase();
          if (content === "none" || content === "normal" || pseudoStyle.display === "none"
            || pseudoStyle.visibility === "hidden" || Number.parseFloat(pseudoStyle.opacity || "1") < 0.01) continue;
          const clone = document.createElement("span");
          clone.setAttribute("aria-hidden", "true");
          clone.setAttribute("data-html2ppt-pseudo-clone", pseudo);
          for (const property of pseudoStyle) {
            clone.style.setProperty(property, pseudoStyle.getPropertyValue(property), pseudoStyle.getPropertyPriority(property));
          }
          clone.style.setProperty("animation", "none", "important");
          clone.style.setProperty("transition", "none", "important");
          clone.style.setProperty("pointer-events", "none", "important");
          source.setAttribute(`data-html2ppt-pseudo-${pseudo}`, "1");
          if (pseudo === "before") source.insertBefore(clone, source.firstChild);
          else source.appendChild(clone);
        }
      }

      const slideRect = slide.getBoundingClientRect();
      const objects = [];
      let sequence = 0;

      const animationRefs = (element) => {
        const refs = [];
        let current = element;
        while (current && current !== slide.parentElement) {
          const id = current.getAttribute?.("data-html2ppt-animation-id");
          if (id) refs.unshift(id);
          current = current.parentElement;
        }
        return refs;
      };

      const objectAnimationRefs = (element, includeDescendants) => {
        const refs = animationRefs(element);
        if (includeDescendants) {
          for (const animated of element.querySelectorAll("[data-html2ppt-animation-id]")) {
            const id = animated.getAttribute("data-html2ppt-animation-id");
            if (id && !refs.includes(id)) refs.push(id);
          }
        }
        return refs;
      };

      slide.querySelectorAll("[data-html2ppt-visual-id]").forEach((element) => {
        element.removeAttribute("data-html2ppt-visual-id");
        element.removeAttribute("data-html2ppt-hide-box");
        element.removeAttribute("data-html2ppt-hide-object");
      });

      const color = (value) => {
        const text = String(value || "");
        const rgba = text.match(/rgba?\(\s*([\d.]+)[, ]+\s*([\d.]+)[, ]+\s*([\d.]+)(?:\s*[,/]\s*([\d.]+))?/i);
        if (!rgba) return { color: "000000", alpha: 0 };
        return {
          color: [rgba[1], rgba[2], rgba[3]]
            .map((part) => Math.max(0, Math.min(255, Math.round(Number(part)))).toString(16).padStart(2, "0"))
            .join("")
            .toUpperCase(),
          alpha: rgba[4] === undefined ? 1 : Math.max(0, Math.min(1, Number(rgba[4]))),
        };
      };

      const effectiveOpacity = (element) => {
        let opacity = 1;
        let current = element;
        while (current && current !== slide.parentElement) {
          const style = getComputedStyle(current);
          if (style.display === "none" || style.visibility === "hidden") return 0;
          opacity *= Number.parseFloat(style.opacity || "1") || 0;
          current = current.parentElement;
        }
        return opacity;
      };

      const border = (style, side) => {
        const width = Number.parseFloat(style[`border${side}Width`]) || 0;
        const parsed = color(style[`border${side}Color`]);
        const borderStyle = style[`border${side}Style`] || "none";
        return {
          width: borderStyle === "none" ? 0 : width,
          color: parsed.color,
          alpha: parsed.alpha,
          style: borderStyle,
        };
      };

      for (const element of slide.querySelectorAll("*")) {
        const pseudo = element.getAttribute("data-html2ppt-pseudo-clone");
        const tag = pseudo ? `pseudo-${pseudo}` : element.tagName.toLowerCase();
        if (["script", "style", "noscript", "textarea", "source"].includes(tag)) continue;
        if (element.closest("svg") && tag !== "svg") continue;
        if (Array.from(assetRoots).some((root) => root !== element && root.contains(element))) continue;
        const style = getComputedStyle(element);
        const opacity = effectiveOpacity(element);
        if (opacity < 0.01 || style.display === "contents") continue;
        const rect = element.getBoundingClientRect();
        if (rect.width < 1.5 || rect.height < 1.5) continue;
        if (rect.right <= slideRect.left || rect.bottom <= slideRect.top || rect.left >= slideRect.right || rect.top >= slideRect.bottom) continue;

        const clippedX = Math.max(rect.left, slideRect.left) - slideRect.left;
        const clippedY = Math.max(rect.top, slideRect.top) - slideRect.top;
        const clippedRight = Math.min(rect.right, slideRect.right) - slideRect.left;
        const clippedBottom = Math.min(rect.bottom, slideRect.bottom) - slideRect.top;
        const clippedW = clippedRight - clippedX;
        const clippedH = clippedBottom - clippedY;
        if (clippedW < 1.5 || clippedH < 1.5) continue;

        const asset = assetRoots.has(element);
        const fill = color(style.backgroundColor);
        fill.alpha *= opacity;
        const backgroundImage = style.backgroundImage || "none";
        const borders = {
          top: border(style, "Top"),
          right: border(style, "Right"),
          bottom: border(style, "Bottom"),
          left: border(style, "Left"),
        };
        for (const item of Object.values(borders)) item.alpha *= opacity;
        // CSS commonly draws arrowheads and carets with a zero-content box:
        // one opaque border forms the triangular face while the two adjacent
        // borders are transparent. Treat that pattern as geometry instead of
        // rebuilding its opaque border as a straight PowerPoint line.
        const triangle = (() => {
          if (asset || fill.alpha > 0.01 || backgroundImage !== "none") return null;
          const visibleSides = Object.entries(borders)
            .filter(([, item]) => item.width > 0.1 && item.alpha > 0.01);
          if (visibleSides.length !== 1) return null;
          const [coloredSide, coloredBorder] = visibleSides[0];
          const candidates = {
            left: { direction: "right", adjacent: ["top", "bottom"], opposite: "right" },
            right: { direction: "left", adjacent: ["top", "bottom"], opposite: "left" },
            top: { direction: "down", adjacent: ["left", "right"], opposite: "bottom" },
            bottom: { direction: "up", adjacent: ["left", "right"], opposite: "top" },
          };
          const candidate = candidates[coloredSide];
          if (!candidate) return null;
          const transparentAdjacent = candidate.adjacent.every((side) => (
            borders[side].width > 0.1 && borders[side].alpha <= 0.01
          ));
          const emptyOpposite = borders[candidate.opposite].width <= 0.1
            || borders[candidate.opposite].alpha <= 0.01;
          if (!transparentAdjacent || !emptyOpposite) return null;
          return {
            direction: candidate.direction,
            color: coloredBorder.color,
            alpha: coloredBorder.alpha,
          };
        })();
        const hasBorder = Object.values(borders).some((item) => item.width > 0.1 && item.alpha > 0.01);
        const hasFill = fill.alpha > 0.01;
        const hasBackgroundImage = backgroundImage !== "none";
        const hasShadow = style.boxShadow && style.boxShadow !== "none";
        if (!asset && !hasFill && !hasBorder && !hasBackgroundImage && !hasShadow) continue;
        if (style.backgroundClip === "text" || style.webkitBackgroundClip === "text") continue;

        const matrix = new DOMMatrixReadOnly(style.transform === "none" ? undefined : style.transform);
        const rotation = Math.abs(matrix.b) > 0.0001 || Math.abs(matrix.a - 1) > 0.0001
          ? Math.atan2(matrix.b, matrix.a) * 180 / Math.PI
          : 0;
        const radiusValue = (value) => {
          const parsed = Number.parseFloat(value) || 0;
          return String(value).includes("%") ? minSidePx * parsed / 100 : parsed;
        };
        const minSidePx = Math.min(rect.width, rect.height);
        const radii = [
          style.borderTopLeftRadius,
          style.borderTopRightRadius,
          style.borderBottomRightRadius,
          style.borderBottomLeftRadius,
        ].map(radiusValue);
        const id = `html2ppt-${targetIndex + 1}-${++sequence}`;
        const kind = asset ? "asset" : hasBackgroundImage ? "background" : "shape";
        // Native PowerPoint shapes may extend beyond the slide and should keep
        // their original geometry (large circles are often intentionally clipped).
        // Raster captures are clipped to the visible slide/viewport region.
        const x = kind === "shape" ? rect.left - slideRect.left : clippedX;
        const y = kind === "shape" ? rect.top - slideRect.top : clippedY;
        const w = kind === "shape" ? rect.width : clippedW;
        const h = kind === "shape" ? rect.height : clippedH;
        element.setAttribute("data-html2ppt-visual-id", id);
        if (kind === "asset") element.setAttribute("data-html2ppt-hide-object", "1");
        else element.setAttribute("data-html2ppt-hide-box", "1");
        objects.push({
          id,
          kind,
          tag,
          x,
          y,
          w,
          h,
          fill,
          borders,
          radii,
          rotation,
          opacity,
          backgroundImage,
          objectFit: style.objectFit || "fill",
          geometry: triangle ? "triangle" : undefined,
          triangle,
          // SVG/canvas/media descendants are flattened into this one picture
          // object. Carry their individual animation sources onto the picture
          // so path drawing, arrow fades and bar growth still have a native
          // PowerPoint approximation.
          animationRefs: objectAnimationRefs(element, asset),
        });
      }
      return objects;
    },
    { selector: info.selector, index },
  );
}

async function captureRasterObject(page, visual) {
  const locator = page.locator(`[data-html2ppt-visual-id="${visual.id}"]`);
  await page.evaluate(
    ({ id, kind }) => {
      const target = document.querySelector(`[data-html2ppt-visual-id="${id}"]`);
      if (!target) return;
      const parseColor = (value) => {
        const match = String(value || "").match(/rgba?\(\s*([\d.]+)[, ]+\s*([\d.]+)[, ]+\s*([\d.]+)(?:\s*[,/]\s*([\d.]+))?/i);
        if (!match) return null;
        return {
          r: Math.round(Number(match[1])),
          g: Math.round(Number(match[2])),
          b: Math.round(Number(match[3])),
          a: match[4] === undefined ? 1 : Number(match[4]),
        };
      };
      const solidAncestorBackdrop = (element) => {
        let current = element.parentElement;
        while (current) {
          const style = getComputedStyle(current);
          // A gradient or image cannot be represented by one solid fallback.
          if (style.backgroundImage && style.backgroundImage !== "none") return null;
          const parsed = parseColor(style.backgroundColor);
          if (parsed?.a >= 0.99) return `rgb(${parsed.r}, ${parsed.g}, ${parsed.b})`;
          current = current.parentElement;
        }
        return null;
      };

      // Element screenshots are flattened by Chromium. Transparent SVGs,
      // canvases and PNGs can therefore inherit the slide/page color instead
      // of the local card color. Paint the nearest opaque ancestor color into
      // the capture target so the flattened image matches the HTML composite.
      if (kind === "asset") {
        const ownBackground = parseColor(getComputedStyle(target).backgroundColor);
        const backdrop = ownBackground?.a >= 0.99 ? null : solidAncestorBackdrop(target);
        if (backdrop) {
          target.setAttribute("data-html2ppt-capture-original-bg", target.style.getPropertyValue("background-color"));
          target.setAttribute("data-html2ppt-capture-original-bg-priority", target.style.getPropertyPriority("background-color"));
          target.style.setProperty("background-color", backdrop, "important");
          if (target.tagName.toLowerCase() === "svg") {
            const rect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
            const viewBox = target.viewBox?.baseVal;
            rect.setAttribute("x", String(viewBox?.x || 0));
            rect.setAttribute("y", String(viewBox?.y || 0));
            rect.setAttribute("width", String(viewBox?.width || "100%"));
            rect.setAttribute("height", String(viewBox?.height || "100%"));
            rect.setAttribute("fill", backdrop);
            rect.setAttribute("data-html2ppt-capture-backdrop", "1");
            target.insertBefore(rect, target.firstChild);
          }
        }
      }
      document.documentElement.setAttribute("data-html2ppt-clean", "1");
      document.documentElement.setAttribute("data-html2ppt-object-capture", kind);
      target.setAttribute("data-html2ppt-capture-target", "1");
    },
    { id: visual.id, kind: visual.kind },
  );
  try {
    return await locator.screenshot({
      animations: "disabled",
      caret: "hide",
      omitBackground: true,
      timeout: 60000,
    });
  } finally {
    await page.evaluate((id) => {
      const target = document.querySelector(`[data-html2ppt-visual-id="${id}"]`);
      if (target) {
        target.removeAttribute("data-html2ppt-capture-target");
        target.querySelectorAll('[data-html2ppt-capture-backdrop="1"]').forEach((element) => element.remove());
        if (target.hasAttribute("data-html2ppt-capture-original-bg")) {
          const value = target.getAttribute("data-html2ppt-capture-original-bg") || "";
          const priority = target.getAttribute("data-html2ppt-capture-original-bg-priority") || "";
          if (value) target.style.setProperty("background-color", value, priority);
          else target.style.removeProperty("background-color");
          target.removeAttribute("data-html2ppt-capture-original-bg");
          target.removeAttribute("data-html2ppt-capture-original-bg-priority");
        }
      }
      document.documentElement.removeAttribute("data-html2ppt-object-capture");
      document.documentElement.removeAttribute("data-html2ppt-clean");
    }, visual.id);
  }
}

async function captureCleanBackground(page, slideElement) {
  await page.evaluate(() => document.documentElement.setAttribute("data-html2ppt-clean", "1"));
  try {
    return await slideElement.screenshot({
      animations: "disabled",
      caret: "hide",
      timeout: 60000,
    });
  } finally {
    await page.evaluate(() => {
      document.documentElement.removeAttribute("data-html2ppt-clean");
      document.querySelectorAll("[data-html2ppt-hide-text]").forEach((element) => {
        element.removeAttribute("data-html2ppt-hide-text");
      });
    });
  }
}

async function captureObjectsBackground(page, slideElement) {
  await page.evaluate(() => {
    document.documentElement.setAttribute("data-html2ppt-clean", "1");
    document.documentElement.setAttribute("data-html2ppt-objects-clean", "1");
  });
  try {
    return await slideElement.screenshot({
      animations: "disabled",
      caret: "hide",
      timeout: 60000,
    });
  } finally {
    await page.evaluate(() => {
      document.documentElement.removeAttribute("data-html2ppt-clean");
      document.documentElement.removeAttribute("data-html2ppt-objects-clean");
      document.querySelectorAll("[data-html2ppt-hide-text]").forEach((element) => {
        element.removeAttribute("data-html2ppt-hide-text");
      });
      document.querySelectorAll("[data-html2ppt-visual-id]").forEach((element) => {
        element.removeAttribute("data-html2ppt-visual-id");
        element.removeAttribute("data-html2ppt-hide-box");
        element.removeAttribute("data-html2ppt-hide-object");
        element.removeAttribute("data-html2ppt-capture-target");
      });
      document.querySelectorAll("[data-html2ppt-pseudo-clone]").forEach((element) => element.remove());
      document.querySelectorAll("[data-html2ppt-pseudo-before]").forEach((element) => element.removeAttribute("data-html2ppt-pseudo-before"));
      document.querySelectorAll("[data-html2ppt-pseudo-after]").forEach((element) => element.removeAttribute("data-html2ppt-pseudo-after"));
    });
  }
}

function normalizeFontFace(value) {
  const first = String(value || "Microsoft YaHei").split(",")[0].trim().replace(/^['"]|['"]$/g, "");
  const lower = first.toLowerCase();
  // Web fonts are not embedded in PPTX. Use a Windows serif with similarly
  // compact numeral metrics so adjacent unit labels stay aligned.
  if (/crimson pro|adobe arabic/.test(lower)) return "Times New Roman";
  if (/pingfang|hiragino|noto sans cjk|source han sans|heiti/.test(lower)) return "Microsoft YaHei";
  if (/noto serif cjk|source han serif|songti/.test(lower)) return "SimSun";
  if (/monospace/.test(lower)) return "Consolas";
  if (/sans-serif|system-ui/.test(lower)) return "Microsoft YaHei";
  if (/serif/.test(lower)) return "SimSun";
  return first || "Microsoft YaHei";
}

function addEditableText(slide, block, sourceRect, backgroundBox) {
  const sx = backgroundBox.w / sourceRect.width;
  const sy = backgroundBox.h / sourceRect.height;
  const fontSize = Math.max(1, block.fontSize * sx * 72);
  const x = backgroundBox.x + block.x * sx;
  const y = backgroundBox.y + block.y * sy - fontSize / 72 * 0.05;
  // Keep each DOM text run close to its browser-measured width. Inflating boxes
  // makes adjacent inline runs (for example "13,526" + "人") collide in PPT.
  // With wrapping disabled, PowerPoint can shrink a wider glyph run into place.
  const w = Math.min(backgroundBox.x + backgroundBox.w - x, Math.max(block.w * sx * 1.01 + 0.01, 0.06));
  const h = Math.max(block.h * sy * 1.18, fontSize / 72 * 1.22);
  if (w <= 0 || h <= 0) return [];

  const objectName = `HTML text ${block.id}`;

  slide.addText(block.text, {
    x,
    y,
    w,
    h,
    margin: 0,
    fontFace: normalizeFontFace(block.fontFamily),
    fontSize,
    bold: Number.parseInt(block.fontWeight, 10) >= 600 || block.fontWeight === "bold",
    italic: block.fontStyle === "italic" || block.fontStyle === "oblique",
    color: block.color,
    align: ["center", "right", "justify"].includes(block.align) ? block.align : "left",
    valign: "mid",
    breakLine: false,
    wrap: false,
    fit: "shrink",
    charSpacing: block.letterSpacing ? block.letterSpacing * sx * 72 : undefined,
    rotate: Math.abs(block.rotation) > 0.1 ? block.rotation : undefined,
    isTextBox: true,
    lineSpacingMultiple: 1,
    paraSpaceAfterPt: 0,
    objectName,
  });
  return [objectName];
}

function scaledObjectBox(visual, sourceRect, backgroundBox) {
  const sx = backgroundBox.w / sourceRect.width;
  const sy = backgroundBox.h / sourceRect.height;
  return {
    x: backgroundBox.x + visual.x * sx,
    y: backgroundBox.y + visual.y * sy,
    w: Math.max(0.01, visual.w * sx),
    h: Math.max(0.01, visual.h * sy),
    sx,
    sy,
  };
}

function borderDash(style) {
  if (style === "dashed") return "dash";
  if (style === "dotted") return "sysDot";
  return "solid";
}

function addEditableVisual(slide, visual, sourceRect, backgroundBox, shapeTypes) {
  const box = scaledObjectBox(visual, sourceRect, backgroundBox);
  if (visual.kind === "asset" || visual.kind === "background") {
    if (!visual.png?.length) return [];
    const objectName = `HTML ${visual.tag || visual.kind} ${visual.id}`;
    slide.addImage({
      data: `data:image/png;base64,${visual.png.toString("base64")}`,
      x: box.x,
      y: box.y,
      w: box.w,
      h: box.h,
      rotate: Math.abs(visual.rotation) > 0.1 ? visual.rotation : undefined,
      objectName,
    });
    return [objectName];
  }

  if (visual.geometry === "triangle" && visual.triangle) {
    const rotation = { up: 0, right: 90, down: 180, left: 270 }[visual.triangle.direction] ?? 0;
    // `triangle` points upward by default. For a horizontal triangle, swap the
    // unrotated extents and keep the centre fixed so the rotated visual bounds
    // still match the browser's narrow arrowhead box.
    const horizontal = visual.triangle.direction === "left" || visual.triangle.direction === "right";
    const triangleBox = horizontal
      ? {
          x: box.x + (box.w - box.h) / 2,
          y: box.y + (box.h - box.w) / 2,
          w: box.h,
          h: box.w,
        }
      : box;
    const objectName = `HTML CSS triangle ${visual.triangle.direction} ${visual.id}`;
    slide.addShape(shapeTypes.triangle, {
      x: triangleBox.x,
      y: triangleBox.y,
      w: triangleBox.w,
      h: triangleBox.h,
      fill: {
        color: visual.triangle.color,
        transparency: Math.round((1 - visual.triangle.alpha) * 100),
      },
      line: { color: visual.triangle.color, transparency: 100, width: 0 },
      rotate: rotation || undefined,
      objectName,
    });
    return [objectName];
  }

  const minSidePx = Math.min(visual.w, visual.h);
  const maxRadius = Math.max(...visual.radii);
  const nearCircle = Math.abs(visual.w - visual.h) <= Math.max(2, Math.max(visual.w, visual.h) * 0.08)
    && maxRadius >= minSidePx * 0.42;
  const shapeType = nearCircle ? shapeTypes.ellipse : maxRadius >= 2 ? shapeTypes.roundRect : shapeTypes.rect;
  // PowerPoint's stock rounded rectangle defaults to a very large corner
  // adjustment (roughly one sixth of the short side). Pass the browser's
  // measured CSS radius in slide inches so PptxGenJS writes an exact OOXML
  // `adj` value instead of leaving the stock 1/4-circle-looking corners.
  const rectRadius = shapeType === shapeTypes.roundRect
    ? Math.min(maxRadius, minSidePx / 2) * ((box.sx + box.sy) / 2)
    : undefined;
  const borders = Object.values(visual.borders);
  const visibleBorders = borders.filter((item) => item.width > 0.1 && item.alpha > 0.01);
  const uniformBorder = visibleBorders.length === 4
    && borders.every((item) => Math.abs(item.width - borders[0].width) < 0.2
      && item.color === borders[0].color
      && Math.abs(item.alpha - borders[0].alpha) < 0.02
      && item.style === borders[0].style);
  const line = uniformBorder
    ? {
        color: borders[0].color,
        transparency: Math.round((1 - borders[0].alpha) * 100),
        width: Math.max(0.1, borders[0].width * ((box.sx + box.sy) / 2) * 72),
        dashType: borderDash(borders[0].style),
      }
    : { color: "FFFFFF", transparency: 100, width: 0 };
  const objectNames = [`HTML ${visual.tag || "shape"} ${visual.id}`];
  slide.addShape(shapeType, {
    x: box.x,
    y: box.y,
    w: box.w,
    h: box.h,
    fill: {
      color: visual.fill.color,
      transparency: Math.round((1 - visual.fill.alpha) * 100),
    },
    line,
    rectRadius,
    rotate: Math.abs(visual.rotation) > 0.1 ? visual.rotation : undefined,
    objectName: objectNames[0],
  });

  if (uniformBorder) return objectNames;
  const sides = [
    ["top", box.x, box.y, box.w, 0],
    ["right", box.x + box.w, box.y, 0, box.h],
    ["bottom", box.x, box.y + box.h, box.w, 0],
    ["left", box.x, box.y, 0, box.h],
  ];
  for (const [side, x, y, w, h] of sides) {
    const border = visual.borders[side];
    if (border.width <= 0.1 || border.alpha <= 0.01) continue;
    const objectName = `HTML border ${side} ${visual.id}`;
    slide.addShape(shapeTypes.line, {
      x,
      y,
      w,
      h,
      line: {
        color: border.color,
        transparency: Math.round((1 - border.alpha) * 100),
        width: Math.max(0.1, border.width * ((box.sx + box.sy) / 2) * 72),
        dashType: borderDash(border.style),
      },
      objectName,
    });
    objectNames.push(objectName);
  }
  return objectNames;
}

function pngDimensions(buffer) {
  if (buffer.length < 24 || buffer.toString("ascii", 1, 4) !== "PNG") {
    throw new Error("截图不是有效的 PNG");
  }
  return { width: buffer.readUInt32BE(16), height: buffer.readUInt32BE(20) };
}

function imageBox(imageWidth, imageHeight, slideWidth, slideHeight, fit) {
  if (fit === "stretch") return { x: 0, y: 0, w: slideWidth, h: slideHeight };
  const imageRatio = imageWidth / imageHeight;
  const slideRatio = slideWidth / slideHeight;
  const cover = fit === "cover";
  const useWidth = cover ? imageRatio < slideRatio : imageRatio > slideRatio;
  if (useWidth) {
    const w = slideWidth;
    const h = w / imageRatio;
    return { x: 0, y: (slideHeight - h) / 2, w, h };
  }
  const h = slideHeight;
  const w = h * imageRatio;
  return { x: (slideWidth - w) / 2, y: 0, w, h };
}

function safeFileStem(filePath) {
  return path.basename(filePath, path.extname(filePath)).replace(/[<>:"/\\|?*]+/g, "-");
}

function representativeAnimationScore(effect) {
  const source = String(effect.sourceName || "").toLowerCase();
  // A flattened SVG/canvas can contain many independently animated children,
  // but PowerPoint sees only one picture shape. Prefer the animation that best
  // describes the asset's main visual construction and avoid replaying that
  // whole picture once for every child path, marker or label.
  if (effect.effectType === 22 && /draw|stroke/.test(source)) return 700;
  if (effect.effectType === 22 && /grow|scale.?y/.test(source)) return 650;
  if (effect.effectType === 61) return 600;
  if (effect.effectType === 48) return 550;
  if (effect.effectType === 2) return 500;
  if (effect.effectType === 22) return 450;
  if (effect.effectType === 10) return 100;
  return 0;
}

function buildAnimationManifest(frames) {
  let totalEffects = 0;
  const slides = frames.map((frame, slideIndex) => {
    const byId = new Map((frame.animations || []).map((animation) => [animation.id, animation]));
    const effects = [];
    for (const target of frame.animationTargets || []) {
      const candidates = [];
      for (const ref of target.animationRefs || []) {
        const source = byId.get(ref);
        if (!source) continue;
        for (const rawEffect of source.effects || []) {
          candidates.push({
            animationId: ref,
            ...mapHtmlEffectToPowerPoint(rawEffect),
          });
        }
      }
      const selected = target.animationStrategy === "single" && candidates.length
        ? [candidates.sort((a, b) => (
            representativeAnimationScore(b) - representativeAnimationScore(a)
            || a.delay - b.delay
            || b.duration - a.duration
          ))[0]]
        : candidates;
      for (const candidate of selected) {
        for (const objectName of target.objectNames || []) {
          effects.push({
            objectName,
            ...candidate,
          });
        }
      }
    }
    effects.sort((a, b) => a.delay - b.delay || a.objectName.localeCompare(b.objectName));
    totalEffects += effects.length;
    return { slide: slideIndex + 1, effects };
  });
  return { version: 2, totalEffects, slides };
}

async function applyPowerPointAnimations(outputPath, manifest, mode) {
  if (!manifest.totalEffects) return false;
  if (process.platform !== "win32") {
    const error = new Error("原生 PPT 动画需要 Windows 桌面版 Microsoft PowerPoint");
    if (mode === "required") throw error;
    console.warn(`动画提示：${error.message}；已保留静态 PPT。`);
    return false;
  }

  const scriptPath = path.join(path.dirname(fileURLToPath(import.meta.url)), "apply-animations.ps1");
  const manifestPath = `${outputPath}.html2ppt-animations-${process.pid}.json`;
  await fs.writeFile(manifestPath, JSON.stringify(manifest, null, 2), "utf8");
  try {
    const { stdout } = await execFileAsync(
      "powershell.exe",
      [
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", scriptPath,
        "-PptxPath", outputPath,
        "-ManifestPath", manifestPath,
      ],
      { windowsHide: true, maxBuffer: 4 * 1024 * 1024, timeout: 180000 },
    );
    const summary = String(stdout || "").trim();
    console.log(`PowerPoint 原生动画：${manifest.totalEffects} 个效果${summary ? `（${summary}）` : ""}`);
    return true;
  } catch (error) {
    const detail = String(error.stderr || error.stdout || error.message).trim().split("\n").at(-1);
    if (mode === "required") throw new Error(`写入 PowerPoint 动画失败：${detail}`);
    console.warn(`动画提示：写入原生动画失败（${detail}）；已保留静态 PPT。`);
    return false;
  } finally {
    await fs.unlink(manifestPath).catch(() => {});
  }
}

async function main() {
  const opts = parseArgs(process.argv.slice(2));
  if (opts.help) {
    console.log(HELP.trim());
    return;
  }
  validateOptions(opts);

  const inputPath = path.resolve(opts.input);
  const inputStat = await fs.stat(inputPath).catch(() => null);
  if (!inputStat?.isFile()) throw new Error(`输入文件不存在：${inputPath}`);
  if (path.extname(inputPath).toLowerCase() !== ".html") throw new Error("输入文件必须是 .html");

  const outputPath = path.resolve(
    opts.output || path.join(path.dirname(inputPath), `${safeFileStem(inputPath)}-${opts.mode}.pptx`),
  );
  await fs.mkdir(path.dirname(outputPath), { recursive: true });

  let imageDir = "";
  if (opts.keepImages) {
    imageDir = path.resolve(
      opts.keepImagesDir || path.join(path.dirname(outputPath), `${safeFileStem(outputPath)}-images`),
    );
    await fs.mkdir(imageDir, { recursive: true });
  }

  const { browser, label: browserLabel } = await launchBrowser(opts.browser);
  console.log(`浏览器：${browserLabel}`);

  const context = await browser.newContext({
    viewport: { width: Math.round(opts.width), height: Math.round(opts.height) },
    deviceScaleFactor: opts.scale,
    bypassCSP: true,
    locale: "zh-CN",
  });
  const page = await context.newPage();
  const captureAnimations = opts.animations !== "off" && opts.mode !== "image";
  await page.emulateMedia({ reducedMotion: captureAnimations ? "no-preference" : "reduce" });
  page.setDefaultTimeout(30000);
  page.on("dialog", (dialog) => dialog.dismiss().catch(() => {}));

  const pageErrors = [];
  page.on("pageerror", (error) => pageErrors.push(error.message));

  try {
    await page.goto(pathToFileURL(inputPath).href, {
      waitUntil: "domcontentloaded",
      timeout: 60000,
    });
    await waitForAssets(page);
    await installCleanCaptureStyle(page);

    const info = await detectSlides(page, opts.selector);
    const title = (await page.title()) || safeFileStem(inputPath);
    console.log(`识别到 ${info.count} 页，选择器：${info.selector}`);

    const frames = [];
    for (let index = 0; index < info.count; index += 1) {
      await activateSlide(page, info, index);
      const animations = captureAnimations
        ? await extractSlideAnimations(page, info, index, opts.animations)
        : [];
      await finishSlideMotion(page, info, index, opts.wait);
      const slideElement = page.locator(info.selector).nth(index);
      await slideElement.waitFor({ state: "attached" });
      const box = await slideElement.boundingBox();
      if (!box || box.width < 10 || box.height < 10) {
        throw new Error(`第 ${index + 1} 页没有可截图的尺寸`);
      }
      let editable = { slideRect: { width: box.width, height: box.height }, textBlocks: [] };
      let visualObjects = [];
      let png;
      if (opts.mode === "objects") {
        editable = await extractEditableText(page, info, index);
        visualObjects = await extractVisualObjects(page, info, index);
        for (const visual of visualObjects) {
          if (visual.kind === "asset" || visual.kind === "background") {
            visual.png = await captureRasterObject(page, visual);
          }
        }
        png = await captureObjectsBackground(page, slideElement);
      } else if (opts.mode === "editable") {
        editable = await extractEditableText(page, info, index);
        png = await captureCleanBackground(page, slideElement);
      } else {
        png = await slideElement.screenshot({
          animations: "disabled",
          caret: "hide",
          timeout: 60000,
        });
      }
      const dimensions = pngDimensions(png);
      frames.push({ png, dimensions, visualObjects, animations, ...editable });
      if (imageDir) {
        const imagePath = path.join(imageDir, `slide-${String(index + 1).padStart(3, "0")}.png`);
        await fs.writeFile(imagePath, png);
      }
      const editableLabel = opts.mode === "objects"
        ? `，${editable.textBlocks.length} 个文本框，${visualObjects.length} 个视觉对象`
        : opts.mode === "editable" ? `，${editable.textBlocks.length} 个文本框` : "";
      const animationLabel = animations.length ? `，${animations.length} 个 HTML 动画源` : "";
      console.log(`已渲染 ${index + 1}/${info.count}：${dimensions.width}×${dimensions.height}${editableLabel}${animationLabel}`);
    }

    const pptx = new pptxgen();
    pptx.layout = "LAYOUT_WIDE";
    pptx.author = "HTML Slides to PPT";
    pptx.company = "Local converter";
    pptx.subject = `Converted from ${path.basename(inputPath)}`;
    pptx.title = title;
    pptx.lang = "zh-CN";
    pptx.theme = {
      headFontFace: "Microsoft YaHei",
      bodyFontFace: "Microsoft YaHei",
      lang: "zh-CN",
    };

    const slideWidth = 13.333333;
    const slideHeight = 7.5;
    const useObjectBackgroundLayouts = opts.mode === "objects" && opts.pageBase === "layout";
    if (useObjectBackgroundLayouts) {
      // The clean page base sits behind every editable object. Keep it on a
      // reusable PowerPoint layout instead of adding a selectable image to
      // each slide. Hashing the rendered base and its fitted bounds lets pages
      // with the same template background share one layout while covers and
      // other genuinely different page types retain their own background.
      const backgroundMasters = new Map();
      for (const frame of frames) {
        const box = imageBox(
          frame.dimensions.width,
          frame.dimensions.height,
          slideWidth,
          slideHeight,
          opts.fit,
        );
        const boxKey = [box.x, box.y, box.w, box.h].map((value) => value.toFixed(6)).join(":");
        const backgroundKey = createHash("sha256")
          .update(frame.png)
          .update(boxKey)
          .digest("hex");
        let masterName = backgroundMasters.get(backgroundKey);
        if (!masterName) {
          masterName = `HTML Objects Background ${backgroundMasters.size + 1}`;
          pptx.defineSlideMaster({
            title: masterName,
            background: { color: "FFFFFF" },
            objects: [{
              image: {
                data: `data:image/png;base64,${frame.png.toString("base64")}`,
                ...box,
                objectName: masterName,
              },
            }],
          });
          backgroundMasters.set(backgroundKey, masterName);
        }
        frame.backgroundBox = box;
        frame.backgroundMasterName = masterName;
      }
      console.log(`对象模式背景版式：${backgroundMasters.size} 个（相同底图已合并）`);
    }

    for (const [index, frame] of frames.entries()) {
      const slide = useObjectBackgroundLayouts
        ? pptx.addSlide({ masterName: frame.backgroundMasterName })
        : pptx.addSlide();
      const box = frame.backgroundBox || imageBox(
        frame.dimensions.width,
        frame.dimensions.height,
        slideWidth,
        slideHeight,
        opts.fit,
      );
      if (!useObjectBackgroundLayouts) {
        slide.background = { color: "FFFFFF" };
        slide.addImage({
          data: `data:image/png;base64,${frame.png.toString("base64")}`,
          ...box,
          objectName: opts.mode === "objects" ? `HTML page base ${index + 1}` : `HTML slide ${index + 1}`,
        });
      }
      if (opts.mode === "objects" && frame.slideRect) {
        frame.animationTargets = [];
        for (const visual of frame.visualObjects) {
          const objectNames = addEditableVisual(slide, visual, frame.slideRect, box, pptx.ShapeType);
          if (objectNames?.length && visual.animationRefs?.length) {
            frame.animationTargets.push({
              objectNames,
              animationRefs: visual.animationRefs,
              // SVG/canvas/media descendants are rasterized into one picture.
              // One representative entrance prevents the complete image from
              // appearing repeatedly for each internal HTML animation.
              animationStrategy: visual.kind === "asset" ? "single" : "all",
            });
          }
        }
      }
      if ((opts.mode === "objects" || opts.mode === "editable") && frame.slideRect) {
        frame.animationTargets ||= [];
        for (const textBlock of frame.textBlocks) {
          const objectNames = addEditableText(slide, textBlock, frame.slideRect, box);
          if (objectNames?.length && textBlock.animationRefs?.length) {
            frame.animationTargets.push({ objectNames, animationRefs: textBlock.animationRefs });
          }
        }
      }
      slide.addNotes(`Source HTML: ${path.basename(inputPath)}\nHTML slide: ${index + 1}`);
    }

    await pptx.writeFile({ fileName: outputPath, compression: true });
    if (captureAnimations) {
      const animationManifest = buildAnimationManifest(frames);
      const directionSummary = summarizeAnimationDirections(animationManifest);
      if (directionSummary) console.log(`动画方向：${directionSummary}`);
      await applyPowerPointAnimations(outputPath, animationManifest, opts.animations);
    }
    console.log(`完成：${outputPath}`);
    if (imageDir) console.log(`逐页图片：${imageDir}`);
    if (pageErrors.length) {
      console.warn(`提示：页面运行时出现 ${pageErrors.length} 个错误；PPT 已生成，请检查复杂交互内容。`);
    }
  } finally {
    await context.close();
    await browser.close();
  }
}

main().catch((error) => {
  console.error(`转换失败：${error.message}`);
  process.exitCode = 1;
});
