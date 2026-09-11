import test from 'node:test';
import assert from 'node:assert/strict';
import { resolveFontFace, mergeInlineTextBlocks } from '../capture-semantics.mjs';
import { addEditableText } from '../convert.mjs';

test('installed Noto Serif SC stays Noto Serif SC instead of matching generic serif', () => {
  const fonts = ['Noto Serif SC', 'Noto Sans SC', 'SimSun', 'Microsoft YaHei'];
  assert.equal(resolveFontFace('"Noto Serif SC", serif', fonts), 'Noto Serif SC');
  assert.equal(resolveFontFace('"Noto Sans SC", sans-serif', fonts), 'Noto Sans SC');
  assert.equal(resolveFontFace('"Missing Serif", "Noto Serif SC", serif', fonts), 'Noto Serif SC');
  assert.equal(resolveFontFace('"Noto Serif SC", serif', null), 'Noto Serif SC');
});

test('merge adjacent inline emphasis without combining independently positioned or animated labels', () => {
  const base = {x:0,y:0,w:100,h:30,fontSize:20,flowId:1,animationRefs:['a']};
  const result = mergeInlineTextBlocks([
    {...base,text:'普通'}, {...base,x:100,w:60,text:'强调',fontWeight:'700',color:'FF0000'},
    {...base,x:160,text:'独立',flowId:2}, {...base,x:260,text:'后出现',flowId:2,animationRefs:['b']},
  ]);
  assert.equal(result.length,3);
  assert.equal(result[0].text,'普通强调');
  assert.equal(result[0].runs[1].color,'FF0000');
  assert.equal(result[0].w,160);
});

test('mixed emphasis keeps per-run font and weight without leaking bold to normal text', () => {
  const base = {id:'mixed',text:'强调正文',x:0,y:0,w:160,h:30,fontSize:31,
    fontFamily:'"Noto Serif SC", serif',fontWeight:'700',color:'FF0000'};
  let result;
  addEditableText({addText:(runs,options)=>{result={runs,options};}}, {
    ...base,runs:[{...base,text:'强调'},{...base,text:'正文',fontWeight:'400',color:'123456'}],
  }, {width:1280,height:720}, {x:0,y:0,w:13.3333333333,h:7.5});
  assert.equal(result.options.bold,false);
  assert.equal(result.runs[0].options.bold,true);
  assert.equal(result.runs[1].options.bold,false);
  assert.equal(result.runs[0].options.fontFace,'Noto Serif SC');
  assert.equal(result.runs[1].options.color,'123456');
  assert.ok(Math.abs(result.runs[0].options.fontSize-23.25)<0.01);
});
