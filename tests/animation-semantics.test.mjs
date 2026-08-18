import assert from "node:assert/strict";
import test from "node:test";

import {
  applyAnimationSemantics,
  mapHtmlEffectToPowerPoint,
  summarizeAnimationDirections,
} from "../animation-semantics.mjs";

function source({
  id = "source-1",
  name = "entrance",
  keyframes = [],
  properties = [],
  rect = { x: 0, y: 0, w: 100, h: 100 },
  transformOrigin = "50px 50px",
  ancestorChain = ["group", "slide"],
  revealHint = null,
} = {}) {
  return {
    id,
    effects: [{ type: "animation", name, duration: 800, delay: 120, keyframes, properties }],
    metadata: {
      rect,
      slideSize: { width: 1920, height: 1080 },
      transformOrigin,
      ancestorChain,
      revealHint,
      elementLabel: `div#${id}`,
    },
  };
}

function resolveOne(options) {
  const result = applyAnimationSemantics([source(options)]);
  assert.deepEqual(result.errors, []);
  return result.animations[0].effects[0];
}

test("scaleY uses bottom and top transform origins", () => {
  const keyframes = [{ transform: "scaleY(0)" }, { transform: "scaleY(1)" }];
  assert.equal(resolveOne({ keyframes, transformOrigin: "50px 100px" }).revealFrom, "bottom");
  assert.equal(resolveOne({ keyframes, transformOrigin: "50px 0px" }).revealFrom, "top");
});

test("scaleX uses left and right transform origins", () => {
  const keyframes = [{ transform: "scaleX(0)" }, { transform: "scaleX(1)" }];
  assert.equal(resolveOne({ keyframes, transformOrigin: "0px 50px" }).revealFrom, "left");
  assert.equal(resolveOne({ keyframes, transformOrigin: "100px 50px" }).revealFrom, "right");
});

test("uniform scale remains a Zoom candidate", () => {
  const effect = resolveOne({
    name: "zoom-in",
    keyframes: [{ transform: "scale(0)" }, { transform: "scale(1)" }],
    transformOrigin: "0px 100px",
  });
  assert.equal(effect.revealFrom, undefined);
  assert.equal(mapHtmlEffectToPowerPoint(effect).effectType, 48);
});

test("clip-path inset reveals from the opposite visible edge", () => {
  const cases = [
    ["inset(100% 0 0 0)", "bottom"],
    ["inset(0 100% 0 0)", "left"],
    ["inset(0 0 100% 0)", "top"],
    ["inset(0 0 0 100%)", "right"],
  ];
  for (const [initial, expected] of cases) {
    const effect = resolveOne({
      name: "clip-reveal",
      keyframes: [{ clipPath: initial }, { clipPath: "inset(0)" }],
    });
    assert.equal(effect.revealFrom, expected, initial);
    assert.equal(effect.directionReason, "clip-path");
  }
});

test("explicit hint overrides inference and auto cancels ancestor override", () => {
  const keyframes = [{ transform: "scaleY(0)" }, { transform: "scaleY(1)" }];
  const explicit = resolveOne({
    keyframes,
    transformOrigin: "50px 100px",
    revealHint: { value: "top", raw: "top", owner: "div.group" },
  });
  assert.equal(explicit.revealFrom, "top");
  assert.equal(explicit.directionReason, "explicit");

  const automatic = resolveOne({
    keyframes,
    transformOrigin: "50px 100px",
    revealHint: { value: "auto", raw: "auto", owner: "div.card" },
  });
  assert.equal(automatic.revealFrom, "bottom");
  assert.equal(automatic.directionReason, "transform-origin");
});

test("explicit edge converts non-directional effects to Wipe and overrides legacy direction", () => {
  const fade = resolveOne({
    name: "fade-in",
    keyframes: [{ opacity: 0 }, { opacity: 1 }],
    revealHint: { value: "right", raw: "right", owner: "div.fade" },
  });
  assert.deepEqual(
    { effectType: mapHtmlEffectToPowerPoint(fade).effectType, direction: mapHtmlEffectToPowerPoint(fade).direction },
    { effectType: 22, direction: 2 },
  );

  const draw = resolveOne({
    name: "draw-line",
    revealHint: { value: "bottom", raw: "bottom", owner: "svg.line" },
  });
  assert.equal(mapHtmlEffectToPowerPoint(draw).direction, 3);
});

test("invalid hint warns in auto and fails in required mode", () => {
  const invalid = source({ revealHint: { value: "up", raw: "up", owner: "div.bad" } });
  const automatic = applyAnimationSemantics([invalid], { slideIndex: 3, mode: "auto" });
  assert.equal(automatic.warnings.length, 1);
  assert.match(automatic.warnings[0], /第 3 页元素 div\.bad/);
  const required = applyAnimationSemantics([invalid], { slideIndex: 3, mode: "required" });
  assert.equal(required.errors.length, 1);
  assert.match(required.errors[0], /data-html2ppt-reveal-from/);
});

test("explicit hint does not synthesize animation for static elements", () => {
  const item = source({ revealHint: { value: "bottom", raw: "bottom", owner: "div.static" } });
  item.effects = [];
  const result = applyAnimationSemantics([item]);
  assert.deepEqual(result.animations[0].effects, []);
});

test("bottom-aligned staircase is inferred from geometry", () => {
  const items = [180, 260, 350, 440].map((height, index) => source({
    id: `step-${index + 1}`,
    name: "grow-step",
    keyframes: [],
    rect: { x: 100 + index * 210, y: 900 - height, w: 180, h: height },
    ancestorChain: ["steps", "slide"],
  }));
  const result = applyAnimationSemantics(items);
  assert.deepEqual(result.animations.map((item) => item.effects[0].revealFrom), Array(4).fill("bottom"));
  assert.deepEqual(result.animations.map((item) => item.effects[0].directionReason), Array(4).fill("staircase"));
});

test("top-aligned hanging structure is inferred from geometry", () => {
  const items = [180, 260, 350].map((height, index) => source({
    id: `hang-${index + 1}`,
    name: "wipe-panel",
    rect: { x: 100 + index * 210, y: 120, w: 180, h: height },
    ancestorChain: ["hanging", "slide"],
  }));
  const result = applyAnimationSemantics(items);
  assert.deepEqual(result.animations.map((item) => item.effects[0].revealFrom), Array(3).fill("top"));
  assert.deepEqual(result.animations.map((item) => item.effects[0].directionReason), Array(3).fill("staircase"));
});

test("geometry does not infer with two items, unstable baselines, or insufficient height difference", () => {
  const two = [0, 1].map((index) => source({
    id: `two-${index}`,
    name: "wipe-panel",
    rect: { x: index * 200, y: 100, w: 180, h: 200 + index * 100 },
  }));
  assert.ok(applyAnimationSemantics(two).animations.every((item) => !item.effects[0].revealFrom));

  const unstable = [
    { x: 0, y: 100, w: 100, h: 200 },
    { x: 120, y: 170, w: 100, h: 300 },
    { x: 240, y: 60, w: 100, h: 500 },
  ].map((rect, index) => source({ id: `unstable-${index}`, name: "wipe-panel", rect }));
  assert.ok(applyAnimationSemantics(unstable).animations.every((item) => !item.effects[0].revealFrom));

  const flat = [0, 1, 2].map((index) => source({
    id: `flat-${index}`,
    name: "wipe-panel",
    rect: { x: index * 120, y: 100 + index, w: 100, h: 200 },
  }));
  assert.ok(applyAnimationSemantics(flat).animations.every((item) => !item.effects[0].revealFrom));
});

test("grow fallback maps to PowerPoint bottom direction 3", () => {
  const mapped = mapHtmlEffectToPowerPoint(resolveOne({ name: "grow" }));
  assert.equal(mapped.effectType, 22);
  assert.equal(mapped.revealFrom, "bottom");
  assert.equal(mapped.direction, 3);
  assert.equal(mapped.directionRequired, true);
});

test("Fade, line drawing, and legacy horizontal Wipe do not regress", () => {
  const fade = mapHtmlEffectToPowerPoint({ name: "fade-in", duration: 500, keyframes: [{ opacity: 0 }, { opacity: 1 }] });
  assert.equal(fade.effectType, 10);
  assert.equal(fade.direction, 0);

  const draw = mapHtmlEffectToPowerPoint({ name: "draw-line", duration: 500, keyframes: [] });
  assert.equal(draw.effectType, 22);
  assert.equal(draw.direction, 4);

  const wipe = mapHtmlEffectToPowerPoint({ name: "wipe", duration: 500, keyframes: [] });
  assert.equal(wipe.effectType, 22);
  assert.equal(wipe.direction, 4);
});

test("direction summaries include edge and reason counts", () => {
  const summary = summarizeAnimationDirections({
    slides: [{ effects: [
      { revealFrom: "bottom", directionReason: "staircase" },
      { revealFrom: "bottom", directionReason: "staircase" },
      { revealFrom: "left", directionReason: "legacy-wipe" },
    ] }],
  });
  assert.equal(summary, "bottom 2（staircase 2）；left 1（legacy-wipe 1）");
});
