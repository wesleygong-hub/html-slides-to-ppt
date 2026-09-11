import test from 'node:test';
import assert from 'node:assert/strict';
import { chromium } from 'playwright';
import { installCaptureHelpers, prepareCapturePage } from '../capture-semantics.mjs';
import { activateSlide, installCleanCaptureStyle, extractVisualObjects, extractEditableText, captureRasterObject } from '../convert.mjs';

test('scaled HTML keeps font proportions, inline content, clipping and transparent backgrounds', async () => {
  const browser = await chromium.launch(process.platform === 'win32' ? {channel:'msedge'} : {});
  try {
    for (const scale of [1,1.5]) {
      const page = await browser.newPage({viewport:{width:1280*scale,height:720*scale}});
      await page.goto(new URL('./fixtures/scaled-layout.html',import.meta.url).href);
      await page.evaluate(()=>document.fonts.ready);
      await installCaptureHelpers(page);
      await installCleanCaptureStyle(page);
      await prepareCapturePage(page,'section.slide');
      const info = {selector:'section.slide',framework:'generic'};
      await activateSlide(page,info,0);
      assert.equal(await page.locator('.slide').first().evaluate(e=>getComputedStyle(e).display),'flex');
      assert.equal(await page.locator('nav').evaluate(e=>getComputedStyle(e).visibility),'hidden');
      assert.equal(await page.locator('.noise').evaluate(e=>getComputedStyle(e).visibility),'visible');
      const visuals = await extractVisualObjects(page,info,0);
      const editable = await extractEditableText(page,info,0);
      const title = editable.textBlocks.find(b=>b.text.startsWith('专业服务'));
      assert.equal(title.text,'专业服务保真转换案例');
      assert.equal(title.runs.length,3);
      assert.ok(Math.abs(title.fontSize / editable.slideRect.width * 960 - 23.25)<0.01);
      assert.ok(editable.textBlocks.some(b=>b.text==='·'));
      assert.ok(!editable.textBlocks.some(b=>b.text.includes('不应导出')));
      const card = visuals.find(v=>v.backgroundImage==='none' && v.kind==='background' && v.h>50*scale);
      assert.ok(card, 'decorative pseudo-element must stay in the card background');
      assert.ok(!visuals.some(v=>v.tag==='pseudo-before'));
      const label = visuals.find(v=>v.kind==='background' && v.h<40*scale);
      const png = await captureRasterObject(page,label);
      const alpha = await page.evaluate(async (data) => {
        const image = new Image(); image.src = 'data:image/png;base64,'+data; await image.decode();
        const canvas = document.createElement('canvas'); canvas.width=image.width; canvas.height=image.height;
        const ctx=canvas.getContext('2d');ctx.drawImage(image,0,0);
        return ctx.getImageData(0,0,1,1).data[3];
      },png.toString('base64'));
      assert.equal(alpha,0,'transparent label must not include the opaque slide background');
      assert.equal(await page.locator('.slide').first().evaluate(e=>getComputedStyle(e).visibility),'visible');
      await activateSlide(page,info,1);
      const second=await page.locator('.slide').nth(1).boundingBox();
      assert.ok(Math.abs(second.x)<1,'second slide stays on the canvas');
      await page.close();
    }
  } finally { await browser.close(); }
});
