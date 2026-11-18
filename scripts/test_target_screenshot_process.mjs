import test from 'node:test';
import assert from 'node:assert/strict';
import sharp from 'sharp';
import {processTargetScreenshotWithNode} from './capture_targeted_sheets.mjs';
test('host pixel processor preserves the measured native crop at retina scale',async()=>{
  const pixels=Buffer.from(Array.from({length:8*8},(_,i)=>[i,255-i,17]).flat());
  const source=await sharp(pixels,{raw:{width:8,height:8,channels:3}}).png().toBuffer();
  const result=await processTargetScreenshotWithNode(source,{viewport:{width:4,height:4},clip:{x:1,y:1,width:2,height:1}});
  assert.deepEqual(result.pixelCrop,{left:2,top:2,width:4,height:2});
  assert.equal(result.sourceExtension,'png');
  const {data,info}=await sharp(result.cropped).raw().toBuffer({resolveWithObject:true});
  assert.equal(info.width,4);assert.equal(info.height,2);
  assert.deepEqual([...data.subarray(0,3)],[18,237,17]);
  assert.deepEqual([...data.subarray(-3)],[29,226,17]);
  await assert.rejects(processTargetScreenshotWithNode(source,{viewport:{width:4,height:8},clip:{x:0,y:0,width:1,height:1}}),/scales disagree/);
});
