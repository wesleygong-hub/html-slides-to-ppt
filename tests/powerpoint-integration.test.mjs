import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);
const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const fixture = path.join(repoRoot, "tests", "fixtures", "staircase.html");
const inspectScript = path.join(repoRoot, "tests", "read-ppt-animations.ps1");
const renderScript = path.join(repoRoot, "tests", "render-ppt-slide.ps1");

test("PowerPoint keeps bottom-up staircase direction and HTML timing", {
  skip: process.platform !== "win32" ? "requires Windows desktop PowerPoint" : false,
  timeout: 240_000,
}, async () => {
  const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), "html2ppt-staircase-"));
  const animatedPptx = path.join(tempRoot, "animated.pptx");
  const staticPptx = path.join(tempRoot, "static.pptx");
  const animatedPng = path.join(tempRoot, "animated.png");
  const staticPng = path.join(tempRoot, "static.png");
  const convert = async (output, animations) => execFileAsync(
    process.execPath,
    [
      path.join(repoRoot, "convert.mjs"),
      "--input", fixture,
      "--output", output,
      "--mode", "objects",
      "--scale", "1",
      "--wait", "20",
      "--animations", animations,
    ],
    { cwd: repoRoot, windowsHide: true, timeout: 180_000, maxBuffer: 8 * 1024 * 1024 },
  );

  try {
    const animated = await convert(animatedPptx, "required");
    assert.match(animated.stdout, /动画方向：bottom 4（staircase 4）/);

    const inspected = await execFileAsync(
      "powershell.exe",
      ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", inspectScript, "-PptxPath", animatedPptx],
      { windowsHide: true, timeout: 60_000, maxBuffer: 4 * 1024 * 1024 },
    );
    const effects = JSON.parse(inspected.stdout.trim());
    assert.equal(effects.length, 4);
    for (const effect of effects) {
      assert.equal(effect.effectType, 22);
      assert.equal(effect.direction, 3);
      assert.ok(Math.abs(effect.duration - 0.8) <= 0.02);
      assert.ok(Math.abs(effect.delay) <= 0.02);
      assert.equal(effect.trigger, 2);
    }

    await convert(staticPptx, "off");
    await execFileAsync(
      "powershell.exe",
      ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", renderScript, "-PptxPath", animatedPptx, "-PngPath", animatedPng],
      { windowsHide: true, timeout: 60_000, maxBuffer: 4 * 1024 * 1024 },
    );
    await execFileAsync(
      "powershell.exe",
      ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", renderScript, "-PptxPath", staticPptx, "-PngPath", staticPng],
      { windowsHide: true, timeout: 60_000, maxBuffer: 4 * 1024 * 1024 },
    );
    const [animatedRender, staticRender] = await Promise.all([
      fs.readFile(animatedPng),
      fs.readFile(staticPng),
    ]);
    assert.deepEqual(animatedRender, staticRender, "animation semantics must not change the final PowerPoint rendering");
  } finally {
    await fs.unlink(animatedPng).catch(() => {});
    await fs.unlink(staticPng).catch(() => {});
    await fs.unlink(animatedPptx).catch(() => {});
    await fs.unlink(staticPptx).catch(() => {});
    await fs.rmdir(tempRoot).catch(() => {});
  }
});
