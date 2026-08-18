export const REVEAL_DIRECTIONS = Object.freeze({
  top: 1,
  right: 2,
  bottom: 3,
  left: 4,
});

const REVEAL_EDGES = new Set(["top", "right", "bottom", "left"]);

function finite(value, fallback = 0) {
  const number = Number(value);
  return Number.isFinite(number) ? number : fallback;
}

function propertyChanged(effect, property) {
  const values = (effect.keyframes || [])
    .map((frame) => String(frame?.[property] ?? "").trim().toLowerCase())
    .filter(Boolean);
  return values.length >= 2 && new Set(values).size > 1;
}

function matrixComponents(transform) {
  const match = String(transform || "").match(/matrix(?:3d)?\(([^)]+)\)/i);
  if (!match) return null;
  const values = match[1].split(",").map((value) => Number.parseFloat(value.trim()));
  if (!values.every(Number.isFinite)) return null;
  if (values.length === 6) {
    return {
      scaleX: Math.hypot(values[0], values[1]),
      scaleY: Math.hypot(values[2], values[3]),
      x: values[4],
      y: values[5],
    };
  }
  if (values.length === 16) {
    return {
      scaleX: Math.hypot(values[0], values[1], values[2]),
      scaleY: Math.hypot(values[4], values[5], values[6]),
      x: values[12],
      y: values[13],
    };
  }
  return null;
}

function transformScale(transform) {
  const text = String(transform || "").trim().toLowerCase();
  if (!text || text === "none") return { x: 1, y: 1 };
  const matrix = matrixComponents(text);
  if (matrix) return { x: matrix.scaleX, y: matrix.scaleY };

  let x = 1;
  let y = 1;
  const scale = text.match(/(?:^|\s)scale\(\s*([-+\d.e]+)(?:\s*,\s*([-+\d.e]+))?\s*\)/i);
  if (scale) {
    x = finite(scale[1], 1);
    y = scale[2] === undefined ? x : finite(scale[2], x);
  }
  const scaleX = text.match(/scalex\(\s*([-+\d.e]+)\s*\)/i);
  const scaleY = text.match(/scaley\(\s*([-+\d.e]+)\s*\)/i);
  if (scaleX) x *= finite(scaleX[1], 1);
  if (scaleY) y *= finite(scaleY[1], 1);
  return { x, y };
}

function matrixTranslation(transform) {
  const matrix = matrixComponents(transform);
  return matrix ? { x: matrix.x, y: matrix.y } : { x: 0, y: 0 };
}

function originRatio(token, size, lowKeyword, highKeyword) {
  const text = String(token || "").trim().toLowerCase();
  if (text === lowKeyword) return 0;
  if (text === highKeyword) return 1;
  if (text === "center") return 0.5;
  if (text.endsWith("%")) return finite(Number.parseFloat(text), 50) / 100;
  if (text.endsWith("px")) return size > 0 ? finite(Number.parseFloat(text), size / 2) / size : 0.5;
  const value = Number.parseFloat(text);
  return Number.isFinite(value) && size > 0 ? value / size : 0.5;
}

function transformOriginRatios(origin, rect = {}) {
  const tokens = String(origin || "50% 50%").trim().toLowerCase().split(/\s+/);
  let xToken = tokens[0] || "50%";
  let yToken = tokens[1] || "50%";
  if (["top", "bottom"].includes(xToken)) {
    yToken = xToken;
    xToken = tokens[1] || "50%";
  }
  return {
    x: originRatio(xToken, finite(rect.w), "left", "right"),
    y: originRatio(yToken, finite(rect.h), "top", "bottom"),
  };
}

function inferScaleReveal(effect, metadata) {
  const frames = (effect.keyframes || []).filter((frame) => frame?.transform);
  if (frames.length < 2) return null;
  const start = transformScale(frames[0].transform);
  const end = transformScale(frames.at(-1).transform);
  const xGrowth = start.x <= 0.25 && end.x >= 0.75 && end.x - start.x >= 0.5;
  const yGrowth = start.y <= 0.25 && end.y >= 0.75 && end.y - start.y >= 0.5;
  if (xGrowth === yGrowth) return null; // Uniform scale remains Zoom.

  const origin = transformOriginRatios(metadata.transformOrigin, metadata.rect);
  if (yGrowth) {
    if (origin.y <= 0.35) return "top";
    if (origin.y >= 0.65) return "bottom";
  } else {
    if (origin.x <= 0.35) return "left";
    if (origin.x >= 0.65) return "right";
  }
  return null;
}

function expandInset(values) {
  if (values.length === 1) return [values[0], values[0], values[0], values[0]];
  if (values.length === 2) return [values[0], values[1], values[0], values[1]];
  if (values.length === 3) return [values[0], values[1], values[2], values[1]];
  return values.slice(0, 4);
}

function parseInset(value, rect = {}) {
  const match = String(value || "").trim().toLowerCase().match(/^inset\(\s*([^)]*)\)$/i);
  if (!match) return null;
  const tokens = match[1]
    .split(/\s+round\s+/i)[0]
    .trim()
    .split(/\s+/);
  if (!tokens.length || tokens.some((token) => !Number.isFinite(Number.parseFloat(token)))) return null;
  return expandInset(tokens).map((token, edge) => {
    const number = Number.parseFloat(token);
    if (!token.endsWith("%")) return number;
    const size = edge === 0 || edge === 2 ? finite(rect.h) : finite(rect.w);
    return size * number / 100;
  });
}

function inferClipReveal(effect, metadata) {
  const frames = (effect.keyframes || []).filter((frame) => frame?.clipPath);
  if (frames.length < 2) return null;
  const start = parseInset(frames[0].clipPath, metadata.rect);
  const end = parseInset(frames.at(-1).clipPath, metadata.rect);
  if (!start || !end) return null;
  const decreases = start.map((value, index) => value - end[index]);
  const maximum = Math.max(...decreases);
  if (maximum <= 0.5) return null;
  const edge = decreases.indexOf(maximum);
  // A large inset means that edge is clipped away; the initially visible
  // content therefore grows outward from the opposite edge.
  return ["bottom", "left", "top", "right"][edge];
}

function isVerticalLayoutCandidate(effect) {
  const name = String(effect.name || "").toLowerCase();
  const properties = (effect.properties || []).join(" ").toLowerCase();
  const transforms = (effect.keyframes || []).map((frame) => String(frame?.transform || "")).join(" ").toLowerCase();
  return /grow|scale.?y/.test(`${name} ${transforms}`)
    || /clip|wipe/.test(`${name} ${properties}`)
    || propertyChanged(effect, "clipPath");
}

function geometryDirection(sources) {
  if (sources.length < 3) return null;
  const slideHeight = Math.max(...sources.map((source) => finite(source.metadata?.slideSize?.height)));
  const tolerance = Math.max(6, slideHeight * 0.01);
  const rects = sources.map((source) => source.metadata.rect);
  if (rects.some((rect) => !rect || finite(rect.w) <= 0 || finite(rect.h) <= 0)) return null;
  const tops = rects.map((rect) => finite(rect.y));
  const bottoms = rects.map((rect) => finite(rect.y) + finite(rect.h));
  const spread = (values) => Math.max(...values) - Math.min(...values);
  const topSpread = spread(tops);
  const bottomSpread = spread(bottoms);
  const clearDifference = Math.max(tolerance * 2, slideHeight * 0.02);
  if (bottomSpread <= tolerance && topSpread >= clearDifference) return "bottom";
  if (topSpread <= tolerance && bottomSpread >= clearDifference) return "top";
  return null;
}

function inferGeometry(sources) {
  const candidates = sources.filter((source) => (
    source.effects.some((effect) => !effect.revealFrom && isVerticalLayoutCandidate(effect))
    && source.metadata?.rect
  ));
  const groups = new Map();
  for (const source of candidates) {
    for (const [depth, ancestor] of (source.metadata.ancestorChain || []).entries()) {
      if (!groups.has(ancestor)) groups.set(ancestor, { depthTotal: 0, sources: new Map() });
      const group = groups.get(ancestor);
      group.depthTotal += depth;
      group.sources.set(source.id, source);
    }
  }
  return [...groups.values()]
    .map((group) => ({
      sources: [...group.sources.values()],
      averageDepth: group.depthTotal / Math.max(1, group.sources.size),
    }))
    .filter((group) => group.sources.length >= 3)
    .sort((a, b) => a.averageDepth - b.averageDepth || a.sources.length - b.sources.length);
}

function normalizeHint(metadata) {
  const hint = metadata?.revealHint;
  if (!hint) return { kind: "none" };
  const value = String(hint.value ?? hint.raw ?? "").trim().toLowerCase();
  if (value === "auto") return { kind: "auto" };
  if (REVEAL_EDGES.has(value)) return { kind: "explicit", value };
  return { kind: "invalid", value: String(hint.raw ?? hint.value ?? "") };
}

export function applyAnimationSemantics(inputSources, { slideIndex = 1, mode = "auto" } = {}) {
  const warnings = [];
  const errors = [];
  const sources = (inputSources || []).map((source) => ({
    ...source,
    metadata: { ...(source.metadata || {}) },
    effects: (source.effects || []).map((effect) => ({ ...effect })),
  }));

  for (const source of sources) {
    const hint = normalizeHint(source.metadata);
    if (hint.kind === "invalid") {
      const owner = source.metadata.revealHint?.owner || source.metadata.elementLabel || source.id;
      const message = `第 ${slideIndex} 页元素 ${owner} 的 data-html2ppt-reveal-from 值“${hint.value}”无效`;
      if (mode === "required") errors.push(message);
      else warnings.push(`${message}；已回退自动方向推断。`);
    }
    for (const effect of source.effects) {
      if (hint.kind === "explicit") {
        effect.revealFrom = hint.value;
        effect.directionReason = "explicit";
        effect.directionRequired = true;
        continue;
      }
      const scaleReveal = inferScaleReveal(effect, source.metadata);
      if (scaleReveal) {
        effect.revealFrom = scaleReveal;
        effect.directionReason = "transform-origin";
        effect.directionRequired = true;
        continue;
      }
      const clipReveal = inferClipReveal(effect, source.metadata);
      if (clipReveal) {
        effect.revealFrom = clipReveal;
        effect.directionReason = "clip-path";
        effect.directionRequired = true;
      }
    }
  }

  for (const group of inferGeometry(sources)) {
    const direction = geometryDirection(group.sources);
    if (!direction) continue;
    for (const source of group.sources) {
      for (const effect of source.effects) {
        if (!effect.revealFrom && isVerticalLayoutCandidate(effect)) {
          effect.revealFrom = direction;
          effect.directionReason = "staircase";
          effect.directionRequired = true;
        }
      }
    }
  }

  for (const source of sources) {
    for (const effect of source.effects) {
      if (effect.revealFrom) continue;
      const name = String(effect.name || "").toLowerCase();
      const transforms = (effect.keyframes || []).map((frame) => String(frame?.transform || "")).join(" ").toLowerCase();
      if (/grow|scale.?y/.test(`${name} ${transforms}`)) {
        effect.revealFrom = "bottom";
        effect.directionReason = "grow-fallback";
        effect.directionRequired = true;
      }
    }
  }

  return { animations: sources, warnings, errors };
}

export function mapHtmlEffectToPowerPoint(effect) {
  const name = String(effect.name || "").toLowerCase();
  const keyframes = Array.isArray(effect.keyframes) ? effect.keyframes : [];
  const transforms = keyframes.map((frame) => String(frame.transform || ""));
  const transformText = transforms.join(" ").toLowerCase();
  const properties = (effect.properties || []).join(" ").toLowerCase();
  const firstTranslation = matrixTranslation(transforms[0]);
  const lastTranslation = matrixTranslation(transforms.at(-1));
  const dx = firstTranslation.x - lastTranslation.x;
  const dy = firstTranslation.y - lastTranslation.y;

  let effectType = 10; // msoAnimEffectFade
  let direction = 0;
  let revealFrom = effect.revealFrom || "";
  let directionReason = effect.directionReason || "";
  let directionRequired = Boolean(effect.directionRequired);
  if (revealFrom && REVEAL_DIRECTIONS[revealFrom]) {
    effectType = 22;
    direction = REVEAL_DIRECTIONS[revealFrom];
  } else if (/draw|stroke/.test(name) || propertyChanged(effect, "strokeDashoffset")) {
    effectType = 22; // msoAnimEffectWipe
    direction = 4;
    revealFrom ||= "left";
    directionReason ||= "legacy-draw";
  } else if (/grow|scale.?y/.test(`${name} ${transformText}`)) {
    effectType = 22;
    direction = REVEAL_DIRECTIONS.bottom;
    revealFrom = "bottom";
    directionReason = "grow-fallback";
    directionRequired = true;
  } else if (/spin|rotate/.test(`${name} ${transformText}`)) {
    effectType = 61; // msoAnimEffectSpin
  } else if (/zoom|scale/.test(`${name} ${transformText}`)) {
    effectType = 48; // msoAnimEffectFadedZoom
  } else if (Math.abs(dy) > 0.5 || /translatey/.test(transformText)) {
    effectType = 10;
  } else if (Math.abs(dx) > 0.5 || /translatex/.test(transformText)) {
    effectType = 2; // msoAnimEffectFly
    direction = dx >= 0 ? 2 : 4;
  } else if (properties.includes("transform")) {
    effectType = 10;
  } else if (/clip|wipe/.test(`${name} ${properties}`)) {
    effectType = 22;
    direction = 4;
    revealFrom = "left";
    directionReason = "legacy-wipe";
  }

  return {
    effectType,
    direction,
    duration: Math.max(0.05, Math.min(30, (Number(effect.duration) || 500) / 1000)),
    delay: Math.max(0, Math.min(30, (Number(effect.delay) || 0) / 1000)),
    sourceType: effect.type || "animation",
    sourceName: effect.name || "unnamed",
    revealFrom: revealFrom || null,
    directionReason: directionReason || null,
    directionRequired,
  };
}

export function summarizeAnimationDirections(manifest) {
  const counts = new Map();
  for (const slide of manifest?.slides || []) {
    for (const effect of slide.effects || []) {
      if (!effect.revealFrom) continue;
      const edge = effect.revealFrom;
      if (!counts.has(edge)) counts.set(edge, { total: 0, reasons: new Map() });
      const item = counts.get(edge);
      item.total += 1;
      const reason = effect.directionReason || "legacy";
      item.reasons.set(reason, (item.reasons.get(reason) || 0) + 1);
    }
  }
  return ["top", "right", "bottom", "left"]
    .filter((edge) => counts.has(edge))
    .map((edge) => {
      const item = counts.get(edge);
      const details = [...item.reasons.entries()].map(([reason, count]) => `${reason} ${count}`).join("，");
      return `${edge} ${item.total}${details ? `（${details}）` : ""}`;
    })
    .join("；");
}
